"""
API для учёта ЭЗПУ и трекеров - замена ручного ввода в Excel.

Написано на чистом Python: http.server (встроен в Python) + pg8000
(чистый Python-драйвер Postgres, без C-расширений). Ничего не требует
компиляции - работает на любой версии Python, включая самые новые.

Эндпоинты:
    GET  /health
    GET  /devices/{serial}
    GET  /devices/{serial}/history
    POST /operations

Запуск локально:
    pip install -r requirements.txt
    set DATABASE_URL=postgresql://...
    python app.py
"""
import os
import re
import json
import datetime
import zipfile
import time
import threading
import urllib.request
import http.cookiejar
import io
from xml.sax.saxutils import escape as xml_escape
from urllib.parse import urlparse, parse_qs
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pg8000

DATABASE_URL = os.environ.get("DATABASE_URL")
ALLOWED_OPERATION_TYPES = {"навешивание", "снятие", "передача", "возврат", "ремонт", "списание"}

DEVICE_RE = re.compile(r"^/devices/([^/]+)$")
DEVICE_HISTORY_RE = re.compile(r"^/devices/([^/]+)/history$")
TRIP_RE = re.compile(r"^/trips/(\d+)$")
TRIP_CLOSE_RE = re.compile(r"^/trips/(\d+)/close$")
TRIP_ASSIGN_RE = re.compile(r"^/trips/(\d+)/assign-device$")
TRIP_STOP_COMPLETE_RE = re.compile(r"^/trips/(\d+)/stops/(\d+)/complete$")
TRIP_STOP_ARRIVE_RE = re.compile(r"^/trips/(\d+)/stops/(\d+)/arrive$")
TRIP_STOP_ZPU_RE = re.compile(r"^/trips/(\d+)/stops/(\d+)/zpu$")
ACT_RE = re.compile(r"^/acts/(\d+)$")
ACT_DOWNLOAD_RE = re.compile(r"^/acts/(\d+)/download$")


def get_connection():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL не задан")
    u = urlparse(DATABASE_URL)
    return pg8000.connect(
        user=u.username,
        password=u.password,
        host=u.hostname,
        port=u.port or 5432,
        database=u.path.lstrip("/"),
    )


def json_default(o):
    if isinstance(o, (datetime.datetime, datetime.date)):
        return o.isoformat()
    return str(o)


# ---------- Логика работы с БД ----------

def db_list_devices(status=None, device_type=None, limit=200):
    conn = get_connection()
    try:
        cur = conn.cursor()
        query = """
            SELECT d.id, d.device_type, d.serial_number, d.status, d.current_location,
                   d.last_operation_at, p.name
            FROM devices d
            LEFT JOIN parties p ON p.id = d.current_holder_id
            WHERE 1=1
        """
        params = []
        if status:
            query += " AND d.status = %s"
            params.append(status)
        if device_type:
            query += " AND d.device_type = %s"
            params.append(device_type)
        query += " ORDER BY d.serial_number LIMIT %s"
        params.append(limit)
        cur.execute(query, params)
        return [
            {
                "id": r[0], "device_type": r[1], "serial_number": r[2],
                "status": r[3], "current_location": r[4],
                "last_operation_at": r[5], "current_holder": r[6],
            }
            for r in cur.fetchall()
        ]
    finally:
        conn.close()


def db_get_device(serial):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT d.id, d.device_type, d.serial_number, d.status, d.current_location,
                   d.last_operation_at, p.name
            FROM devices d
            LEFT JOIN parties p ON p.id = d.current_holder_id
            WHERE d.serial_number = %s
            """,
            (serial,),
        )
        row = cur.fetchone()
        if not row:
            return None
        return {
            "id": row[0], "device_type": row[1], "serial_number": row[2],
            "status": row[3], "current_location": row[4],
            "last_operation_at": row[5], "current_holder": row[6],
        }
    finally:
        conn.close()


def db_get_device_history(serial):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT id FROM devices WHERE serial_number = %s", (serial,))
        device = cur.fetchone()
        if not device:
            return None
        cur.execute(
            """
            SELECT o.operation_type, fp.name, tp.name, o.location, o.operation_dt, o.document_ref
            FROM operations o
            LEFT JOIN parties fp ON fp.id = o.from_party_id
            LEFT JOIN parties tp ON tp.id = o.to_party_id
            WHERE o.device_id = %s
            ORDER BY o.operation_dt DESC
            LIMIT 100
            """,
            (device[0],),
        )
        return [
            {
                "operation_type": r[0], "from": r[1], "to": r[2],
                "location": r[3], "operation_dt": r[4], "document_ref": r[5],
            }
            for r in cur.fetchall()
        ]
    finally:
        conn.close()


def _get_or_create_party(cur, name):
    if not name:
        return None
    cur.execute("SELECT id FROM parties WHERE name = %s", (name,))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute("INSERT INTO parties (name, type) VALUES (%s, 'internal') RETURNING id", (name,))
    return cur.fetchone()[0]


def _get_or_create_device(cur, serial, device_type_hint):
    cur.execute("SELECT id FROM devices WHERE serial_number = %s", (serial,))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute(
        "INSERT INTO devices (device_type, serial_number, status) VALUES (%s, %s, 'рабочая') RETURNING id",
        (device_type_hint, serial),
    )
    return cur.fetchone()[0]


def db_create_operation(payload):
    conn = get_connection()
    try:
        cur = conn.cursor()
        device_id = _get_or_create_device(cur, payload["device_serial"], payload["device_type_hint"])
        from_id = _get_or_create_party(cur, payload["from_party"])
        to_id = _get_or_create_party(cur, payload["to_party"])

        cur.execute(
            """
            INSERT INTO operations
                (device_id, operation_type, from_party_id, to_party_id,
                 location, operation_dt, document_ref, recorded_by)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id
            """,
            (device_id, payload["operation_type"], from_id, to_id,
             payload["location"], payload["operation_dt"], payload["document_ref"], payload["recorded_by"]),
        )
        new_id = cur.fetchone()[0]

        cur.execute(
            """
            UPDATE devices SET
                current_holder_id = COALESCE(%s, current_holder_id),
                current_location = COALESCE(%s, current_location),
                last_operation_at = %s,
                updated_at = now()
            WHERE id = %s
            """,
            (to_id, payload["location"], payload["operation_dt"], device_id),
        )
        conn.commit()
        return new_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ТМА — учёт ЭЗПУ и трекеров</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: -apple-system, Segoe UI, Arial, sans-serif; margin: 0; background: #f4f4f2; color: #222; }
  header { background: #1f2937; color: #fff; padding: 16px 24px; }
  header h1 { margin: 0; font-size: 18px; font-weight: 600; }
  .wrap { max-width: 1100px; margin: 0 auto; padding: 20px; }
  .panel { background: #fff; border-radius: 8px; padding: 16px; margin-bottom: 16px; box-shadow: 0 1px 2px rgba(0,0,0,0.08); }
  .table-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; margin: 0 -16px; padding: 0 16px; }
  .table-scroll table { min-width: 900px; }
  .filters { display: flex; gap: 8px; flex-wrap: wrap; align-items: end; }
  .field { display: flex; flex-direction: column; gap: 4px; }
  .field label { font-size: 12px; color: #666; }
  input, select, button { padding: 8px 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px; }
  td.actions-cell { white-space: normal; min-width: 160px; }
  td.actions-cell button { font-size: 12px; padding: 5px 8px; margin: 2px 3px 2px 0; }
  button { background: #1f2937; color: #fff; border: none; cursor: pointer; }
  button:hover { background: #374151; }
  button.secondary { background: #e5e7eb; color: #222; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { text-align: left; padding: 9px 10px; border-bottom: 1px solid #eee; }
  th { background: #f4f5f7; font-weight: 600; color: #444; font-size: 12px; text-transform: uppercase; letter-spacing: 0.02em; border-bottom: 2px solid #e2e4e8; }
  tr.device-row { cursor: pointer; }
  tr.device-row:hover { background: #f3f4f6; }
  .status-pill { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; background: #e5e7eb; }
  .status-рабочая, .status-рабочий { background: #dcfce7; color: #166534; }
  .status-неисправная, .status-неисправный { background: #fee2e2; color: #991b1b; }
  .count { color: #666; font-size: 13px; margin-bottom: 8px; }
  #history-panel, #form-panel { display: none; }
  #history-panel.open, #form-panel.open { display: block; }
  .close-btn { float: right; background: none; color: #666; font-size: 18px; padding: 0 4px; }
  .hist-row { padding: 6px 0; border-bottom: 1px dashed #eee; font-size: 13px; }
  .hist-row b { color: #1f2937; }
  .op-form { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
  .op-form .full { grid-column: 1 / -1; }
  #msg { margin-top: 10px; font-size: 13px; }
  #msg.ok { color: #166534; }
  #msg.err { color: #991b1b; }
  .tabs { display: flex; gap: 4px; margin-bottom: 16px; }
  .tab-btn { background: #e5e7eb; color: #444; border: none; padding: 10px 18px; border-radius: 6px 6px 0 0; cursor: pointer; font-size: 14px; }
  .tab-btn.active { background: #fff; color: #1f2937; font-weight: 600; }
  .tab-view { display: none; }
  .tab-view.active { display: block; }
  .trip-status { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; background: #e5e7eb; color: #444; }
  .trip-status-ожидание { background: #fef3c7; color: #92400e; }
  .trip-status-запланирован { background: #fef3c7; color: #92400e; }
  .trip-status-исполнено { background: #dcfce7; color: #166534; }
  .trip-status-снят { background: #dcfce7; color: #166534; }
  .trip-status-в-пути { background: #dbeafe; color: #1e40af; }
  .trip-main-row { font-weight: 600; }
  .trip-main-row td { background: linear-gradient(to bottom, #f3f4f6, #f8fafc); border-top: 2px solid #cbd5e1; padding-top: 12px; padding-bottom: 12px; }
  .trip-main-row td:first-child { border-left: 3px solid #1f2937; }
  .trip-leg-row { font-weight: 400; color: #8a8f98; font-size: 12.5px; }
  .trip-leg-row td { padding-top: 6px; padding-bottom: 6px; border-bottom: 1px dotted #eee; }
  .trip-leg-row td:first-child { padding-left: 20px; color: #b0b4bb; }
  .trip-leg-row .trip-status { font-size: 11px; }
  .trip-leg-row:hover td { background: #fafafa; }
  .btn-success { background: #16a34a; color: #fff; border: none; font-weight: 600; }
  .btn-success:hover { background: #15803d; }
  .waypoint-row { display: flex; gap: 8px; margin-bottom: 6px; align-items: center; }
  .waypoint-row input { flex: 1; }
  .waypoint-row button { padding: 6px 10px; }
  .stop-line { display: flex; align-items: center; gap: 6px; white-space: nowrap; }
  .stop-line:not(:last-child) { margin-bottom: 3px; }
  .stop-dot { width: 7px; height: 7px; border-radius: 50%; background: #ef4444; flex-shrink: 0; }
  .stop-dot.done { background: #22c55e; }
</style>
</head>
<body>
<header><h1>ТМА — учёт ЭЗПУ и трекеров</h1></header>
<div class="wrap">

  <div class="tabs">
    <button class="tab-btn active" id="tab-btn-devices" onclick="switchTab('devices')">Устройства</button>
    <button class="tab-btn" id="tab-btn-trips" onclick="switchTab('trips')">Рейсы</button>
    <button class="tab-btn" id="tab-btn-acts" onclick="switchTab('acts')">Акты</button>
    <button class="tab-btn" id="tab-btn-reports" onclick="switchTab('reports')">Отчёты</button>
  </div>

  <div class="tab-view active" id="tab-devices">

  <div class="panel">
    <div class="filters">
      <div class="field">
        <label>Серийный номер</label>
        <input id="f-search" placeholder="GNS20776">
      </div>
      <div class="field">
        <label>Статус</label>
        <input id="f-status" placeholder="рабочая, неисправная...">
      </div>
      <div class="field">
        <label>Тип</label>
        <select id="f-type">
          <option value="">все</option>
          <option value="ezpu">ЭЗПУ</option>
          <option value="tracker">трекер</option>
        </select>
      </div>
      <button onclick="loadDevices()">Найти</button>
      <button class="secondary" onclick="clearFilters()">Сбросить</button>
      <button class="secondary" onclick="openForm()" style="margin-left:auto">+ Новая операция</button>
    </div>
  </div>

  <div class="panel" id="form-panel">
    <button class="close-btn" onclick="closeForm()">×</button>
    <h3>Новая операция</h3>
    <div class="op-form">
      <div class="field"><label>Серийный номер устройства *</label><input id="op-serial"></div>
      <div class="field">
        <label>Тип операции *</label>
        <select id="op-type">
          <option value="навешивание">навешивание</option>
          <option value="снятие">снятие</option>
          <option value="передача">передача</option>
          <option value="возврат">возврат</option>
          <option value="ремонт">ремонт</option>
          <option value="списание">списание</option>
        </select>
      </div>
      <div class="field"><label>Откуда</label><input id="op-from"></div>
      <div class="field"><label>Куда</label><input id="op-to"></div>
      <div class="field"><label>Местонахождение</label><input id="op-location"></div>
      <div class="field"><label>Документ / основание</label><input id="op-doc"></div>
      <div class="full"><button onclick="submitOperation()">Сохранить операцию</button></div>
      <div class="full" id="msg"></div>
    </div>
  </div>

  <div class="panel" id="history-panel">
    <button class="close-btn" onclick="closeHistory()">×</button>
    <h3 id="hist-title">История</h3>
    <div id="hist-body"></div>
  </div>

  <div class="panel">
    <div class="count" id="count-label">Загрузка...</div>
    <div class="table-scroll">
    <table>
      <thead><tr>
        <th>Серийный номер</th><th>Тип</th><th>Статус</th><th>Местонахождение</th><th>Держатель</th><th>Последняя операция</th>
      </tr></thead>
      <tbody id="devices-body"></tbody>
    </table>
    </div>
  </div>

  </div>

  <div class="tab-view" id="tab-trips">

  <div class="panel">
    <div class="filters">
      <div class="field">
        <label>Статус рейса</label>
        <select id="tf-status">
          <option value="">активные (без завершённых)</option>
          <option value="запланирован">запланирован</option>
          <option value="в пути">в пути</option>
          <option value="снят">снят (завершённые)</option>
        </select>
      </div>
      <div class="field">
        <label>Клиент</label>
        <input id="tf-client" placeholder="JTI">
      </div>
      <button onclick="loadTrips()">Найти</button>
      <button class="secondary" onclick="openTripForm()" style="margin-left:auto">+ Новый рейс</button>
    </div>
  </div>

  <div class="panel" id="trip-form-panel" style="display:none">
    <button class="close-btn" onclick="closeTripForm()">×</button>
    <h3>Новый рейс</h3>
    <div class="op-form">
      <div class="field"><label>Клиент</label><input id="tr-client"></div>
      <div class="field">
        <label>Подрядчик</label>
        <select id="tr-contractor" onchange="onContractorChange()">
          <option value="">не указан</option>
          <option value="ТОО ТК Мегаполис Казахстан">ТОО ТК Мегаполис Казахстан</option>
          <option value="ТОО ТМЕ">ТОО ТМЕ</option>
          <option value="ТОО СОП ТЖК">ТОО СОП ТЖК</option>
        </select>
      </div>
      <div class="field"><label>№ борта</label><input id="tr-board"></div>
      <div class="field"><label id="tr-tracker-label">Серийный номер трекера</label><input id="tr-tracker"></div>
      <div class="field"><label>Серийный номер ЭЗПУ</label><input id="tr-ezpu"></div>
      <div class="field"><label>№ ЗПУ (разовая пломба)</label><input id="tr-zpu"></div>
      <div class="field"><label>Серийный номер закладки</label><input id="tr-lock"></div>
      <div class="field"><label>Город отправления</label><input id="tr-origin"></div>
      <div class="full">
        <label>Пункты погрузки (склады) *</label>
        <div id="pickups-list"></div>
        <button class="secondary" type="button" onclick="addStop('pickups-list')" style="margin-top:6px">+ Добавить склад</button>
      </div>
      <div class="full">
        <label>Пункты выгрузки *</label>
        <div id="dropoffs-list"></div>
        <button class="secondary" type="button" onclick="addStop('dropoffs-list')" style="margin-top:6px">+ Добавить пункт</button>
      </div>
      <div class="full"><label>Примечания</label><input id="tr-notes" style="width:100%"></div>
      <div class="full"><button onclick="submitTrip()">Создать рейс</button></div>
      <div class="full" id="trip-msg"></div>
    </div>
  </div>


  <div class="panel" id="trip-detail-panel" style="display:none">
    <button class="close-btn" onclick="closeTripDetail()">×</button>
    <h3 id="trip-detail-title">Рейс</h3>
    <div id="trip-detail-body"></div>
  </div>

  <div class="panel">
    <div class="count" id="trip-count-label">Загрузка...</div>
    <div class="table-scroll">
    <table>
      <thead><tr>
        <th>№</th><th>Номер борта</th><th>ЭЗПУ</th><th>№ ЗПУ</th><th>Трекер</th><th>Закладка</th><th>Отправление</th><th>Назначения</th><th>Навешена</th><th>Снято</th><th>Статус</th><th></th>
      </tr></thead>
      <tbody id="trips-body"></tbody>
    </table>
    </div>
  </div>

  </div>

  <div class="tab-view" id="tab-acts">

  <div class="panel" id="act-form-panel">
    <h3>Новый акт приёма-передачи</h3>
    <div class="op-form">
      <div class="field">
        <label>Направление</label>
        <select id="act-direction">
          <option value="передача">Передача (ТМА → Заказчику)</option>
          <option value="возврат">Возврат (Заказчик → ТМА)</option>
        </select>
      </div>
      <div class="field">
        <label>Заказчик (полное наименование)</label>
        <select id="act-counterparty-preset" onchange="onActCounterpartyPreset()">
          <option value="">выбрать из списка...</option>
          <option value="ТОО «ТК «Мегаполис-Казахстан»">ТОО «ТК «Мегаполис-Казахстан»</option>
          <option value="ТОО «ТМЕ»">ТОО «ТМЕ»</option>
          <option value="ТОО «СОП ТЖК»">ТОО «СОП ТЖК»</option>
        </select>
      </div>
      <div class="field"><label>&nbsp;</label><input id="act-counterparty" placeholder="или впишите вручную"></div>
      <div class="field"><label>Метка склада (напр. JTI)</label><input id="act-counterparty-label"></div>
      <div class="field"><label>Дата акта</label><input id="act-date" type="date"></div>
      <div class="field"><label>№ договора</label><input id="act-contract-number" value="07/04-2020"></div>
      <div class="field"><label>Дата договора</label><input id="act-contract-date" type="date" value="2020-07-01"></div>
      <div class="field"><label>Директор (Исполнитель)</label><input id="act-director" value="Архаров Н.Э."></div>
    </div>

    <h4 style="margin-top:20px">Поиск устройств для добавления в акт</h4>
    <div class="filters">
      <div class="field" style="flex:1">
        <label>Серийный номер</label>
        <input id="act-device-search" placeholder="начните вводить номер..." oninput="renderActDeviceSearch()">
      </div>
    </div>
    <div id="act-device-search-results"></div>

    <div class="op-form" style="margin-top:16px">
      <div class="full">
        <label>ЭЗПУ (серийные номера через запятую или с новой строки)</label>
        <textarea id="act-ezpu-list" rows="2" style="width:100%"></textarea>
      </div>
      <div class="field"><label>Цена за ед., тенге (ЭЗПУ)</label><input id="act-ezpu-price" type="number" value="190000"></div>

      <div class="full">
        <label>Трекеры (серийные номера через запятую или с новой строки)</label>
        <textarea id="act-tracker-list" rows="2" style="width:100%"></textarea>
      </div>
      <div class="field"><label>Цена за ед., тенге (трекеры)</label><input id="act-tracker-price" type="number" value="60000"></div>
    </div>

    <h4 style="margin-top:20px">Прочие позиции (ЗПУ и всё, что впишете вручную)</h4>
    <div id="act-custom-lines"></div>
    <button class="secondary" type="button" onclick="addActCustomLine()">+ Добавить произвольную позицию</button>

    <div style="margin-top:16px">
      <button onclick="submitAct()">Создать акт</button>
      <span id="act-msg" style="margin-left:10px"></span>
    </div>
  </div>

  <div class="panel">
    <div class="count" id="acts-count-label">Загрузка...</div>
    <div class="table-scroll">
    <table>
      <thead><tr>
        <th>№ акта</th><th>Направление</th><th>Заказчик</th><th>Дата</th><th>Позиций</th><th>Создан</th><th></th>
      </tr></thead>
      <tbody id="acts-body"></tbody>
    </table>
    </div>
  </div>

  </div>

  <div class="tab-view" id="tab-reports">

  <div class="panel" id="reports-list-panel">
    <h3>Отчёты по компаниям</h3>
    <div id="reports-company-list" class="filters"></div>
  </div>

  <div class="panel" id="reports-detail-panel" style="display:none">
    <button class="close-btn" onclick="closeReportDetail()">×</button>
    <h3 id="reports-detail-title">Отчёт</h3>
    <div class="filters">
      <div class="field">
        <label>Месяц</label>
        <select id="report-month"></select>
      </div>
      <div class="field">
        <label>Год</label>
        <input id="report-year" type="number" style="width:90px">
      </div>
      <div class="field">
        <label>Ставка, тенге/сутки</label>
        <input id="report-rate" type="number" value="4060" style="width:100px">
      </div>
      <button onclick="loadReportDetail()">Показать</button>
    </div>
    <div class="table-scroll">
    <table>
      <thead><tr>
        <th>№</th><th>Борт</th><th>ЭЗПУ</th><th>Откуда</th><th>Куда</th>
        <th>Навешивание</th><th>Снятие</th><th>Дней</th><th>Ставка</th><th>Сумма</th>
      </tr></thead>
      <tbody id="report-rows-body"></tbody>
    </table>
    </div>
    <div id="report-total" style="margin-top:14px; font-weight:600; font-size:15px"></div>
  </div>

  </div>

</div>

<script>
function fmtDate(s) {
  if (!s) return '—';
  const d = new Date(s);
  return d.toLocaleString('ru-RU', {day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit'});
}

async function loadDevices() {
  const search = document.getElementById('f-search').value.trim();
  const status = document.getElementById('f-status').value.trim();
  const type = document.getElementById('f-type').value;
  document.getElementById('count-label').textContent = 'Загрузка...';

  if (search) {
    const r = await fetch('/devices/' + encodeURIComponent(search));
    const body = document.getElementById('devices-body');
    body.innerHTML = '';
    if (r.status === 404) {
      document.getElementById('count-label').textContent = 'Устройство не найдено';
      return;
    }
    const d = await r.json();
    renderRows([d]);
    document.getElementById('count-label').textContent = '1 устройство';
    return;
  }

  const params = new URLSearchParams();
  if (status) params.set('status', status);
  if (type) params.set('device_type', type);
  const r = await fetch('/devices?' + params.toString());
  const data = await r.json();
  renderRows(data.devices);
  document.getElementById('count-label').textContent = data.count + ' устройств';
}

function renderRows(devices) {
  const body = document.getElementById('devices-body');
  body.innerHTML = '';
  devices.forEach(d => {
    const tr = document.createElement('tr');
    tr.className = 'device-row';
    tr.onclick = () => openHistory(d.serial_number);
    tr.innerHTML = `
      <td>${d.serial_number}</td>
      <td>${d.device_type === 'ezpu' ? 'ЭЗПУ' : 'трекер'}</td>
      <td><span class="status-pill status-${d.status}">${d.status}</span></td>
      <td>${d.current_location || '—'}</td>
      <td>${d.current_holder || '—'}</td>
      <td>${fmtDate(d.last_operation_at)}</td>
    `;
    body.appendChild(tr);
  });
}

function clearFilters() {
  document.getElementById('f-search').value = '';
  document.getElementById('f-status').value = '';
  document.getElementById('f-type').value = '';
  loadDevices();
}

async function openHistory(serial) {
  const panel = document.getElementById('history-panel');
  panel.classList.add('open');
  document.getElementById('hist-title').textContent = 'История: ' + serial;
  document.getElementById('hist-body').innerHTML = 'Загрузка...';
  const r = await fetch('/devices/' + encodeURIComponent(serial) + '/history');
  const rows = await r.json();
  if (!rows.length) {
    document.getElementById('hist-body').innerHTML = 'Операций пока нет';
    return;
  }
  document.getElementById('hist-body').innerHTML = rows.map(op => `
    <div class="hist-row">
      <b>${op.operation_type}</b> · ${op.from || '—'} → ${op.to || '—'} · ${op.location || '—'} · ${fmtDate(op.operation_dt)}
      ${op.document_ref ? '<br><span style="color:#888">' + op.document_ref + '</span>' : ''}
    </div>
  `).join('');
  panel.scrollIntoView({behavior: 'smooth'});
}
function closeHistory() { document.getElementById('history-panel').classList.remove('open'); }

function openForm() {
  document.getElementById('form-panel').classList.add('open');
  document.getElementById('form-panel').scrollIntoView({behavior: 'smooth'});
}
function closeForm() { document.getElementById('form-panel').classList.remove('open'); }

async function submitOperation() {
  const msg = document.getElementById('msg');
  msg.textContent = '';
  msg.className = '';
  const payload = {
    device_serial: document.getElementById('op-serial').value.trim(),
    operation_type: document.getElementById('op-type').value,
    from_party: document.getElementById('op-from').value.trim(),
    to_party: document.getElementById('op-to').value.trim(),
    location: document.getElementById('op-location').value.trim(),
    document_ref: document.getElementById('op-doc').value.trim(),
  };
  if (!payload.device_serial) {
    msg.textContent = 'Укажите серийный номер устройства';
    msg.className = 'err';
    return;
  }
  const r = await fetch('/operations', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload),
  });
  const data = await r.json();
  if (r.status === 201) {
    msg.textContent = 'Операция сохранена, id ' + data.id;
    msg.className = 'ok';
    document.getElementById('op-serial').value = '';
    document.getElementById('op-from').value = '';
    document.getElementById('op-to').value = '';
    document.getElementById('op-location').value = '';
    document.getElementById('op-doc').value = '';
    loadDevices();
  } else {
    msg.textContent = data.error || 'Ошибка сохранения';
    msg.className = 'err';
  }
}

function switchTab(name) {
  document.getElementById('tab-devices').classList.toggle('active', name === 'devices');
  document.getElementById('tab-trips').classList.toggle('active', name === 'trips');
  document.getElementById('tab-acts').classList.toggle('active', name === 'acts');
  document.getElementById('tab-reports').classList.toggle('active', name === 'reports');
  document.getElementById('tab-btn-devices').classList.toggle('active', name === 'devices');
  document.getElementById('tab-btn-trips').classList.toggle('active', name === 'trips');
  document.getElementById('tab-btn-acts').classList.toggle('active', name === 'acts');
  document.getElementById('tab-btn-reports').classList.toggle('active', name === 'reports');
  if (name === 'trips') loadTrips();
  if (name === 'acts') { loadAllDevicesCache(); loadActs(); if (!document.getElementById('act-date').value) document.getElementById('act-date').value = new Date().toISOString().slice(0,10); }
  if (name === 'reports') renderReportsCompanyList();
}

let editingTrips = new Set();

let editingZpuStops = new Set();

function statusClass(status) {
  return 'trip-status-' + String(status).toLowerCase().replace(/\s+/g, '-');
}

async function loadTrips() {
  const status = document.getElementById('tf-status').value;
  const client = document.getElementById('tf-client').value.trim();
  document.getElementById('trip-count-label').textContent = 'Загрузка...';
  const params = new URLSearchParams();
  if (status) params.set('status', status);
  if (client) params.set('client', client);
  const r = await fetch('/trips?' + params.toString());
  const data = await r.json();
  const hideCompleted = !status;
  const trips = hideCompleted ? data.trips.filter(t => t.status !== 'снят') : data.trips;
  const body = document.getElementById('trips-body');
  body.innerHTML = '';
  trips.forEach(t => {
    const combined = [...(t.pickups || []), ...(t.dropoffs || [])].sort((a, b) => a.sequence - b.sequence);
    let legs = [];
    for (let i = 0; i < combined.length - 1; i++) {
      legs.push({from: combined[i].location, to: combined[i + 1].location, fromStop: combined[i], toStop: combined[i + 1]});
    }
    if (legs.length === 0 && combined.length === 1) {
      legs.push({from: combined[0].location, to: '—', fromStop: combined[0], toStop: combined[0]});
    }
    if (legs.length === 0) {
      legs.push({from: '—', to: '—', fromStop: null, toStop: null});
    }

    const pickupsList = (t.pickups && t.pickups.length ? t.pickups : [])
      .slice().sort((a, b) => a.sequence - b.sequence)
      .map(s => `<div class="stop-line"><span class="stop-dot ${s.status === 'исполнено' ? 'done' : ''}"></span>${s.location}</div>`)
      .join('') || '—';
    const dropoffsList = (t.dropoffs && t.dropoffs.length ? t.dropoffs : [])
      .slice().sort((a, b) => a.sequence - b.sequence)
      .map(s => `<div class="stop-line"><span class="stop-dot ${s.status === 'исполнено' ? 'done' : ''}"></span>${s.location}</div>`)
      .join('') || '—';
    const dropoffs = t.dropoffs || [];
    const allDropoffsDone = dropoffs.length > 0 && dropoffs.every(s => s.status === 'исполнено');
    const editingDevice = editingTrips.has(t.id);

    // ---------- Главная строка рейса ----------
    const mainRow = document.createElement('tr');
    mainRow.className = 'device-row trip-main-row';
    mainRow.onclick = () => openTripDetail(t.id);

    const mainActions = [];
    if (editingDevice) {
      mainActions.push(`<button onclick="event.stopPropagation(); saveDevice(${t.id})">Сохранить</button>`);
      mainActions.push(`<button class="secondary" onclick="event.stopPropagation(); cancelEditDevice(${t.id})">Отмена</button>`);
    } else {
      const label = t.status === 'запланирован' ? 'Назначить устройство' : 'Изменить устройство';
      mainActions.push(`<button class="secondary" onclick="event.stopPropagation(); startEditDevice(${t.id})">${label}</button>`);
      if (t.status === 'в пути' && allDropoffsDone) {
        mainActions.push(`<button class="btn-success" onclick="event.stopPropagation(); closeTrip(${t.id})">Рейс завершён</button>`);
      } else if (t.status === 'в пути') {
        mainActions.push(`<button class="secondary" onclick="event.stopPropagation(); closeTrip(${t.id})">Закрыть вручную</button>`);
      }
    }

    const mainDeviceCells = editingDevice ? `
      <td><input id="edit-ezpu-${t.id}" value="${t.ezpu_serial || ''}" style="width:100px" onclick="event.stopPropagation()"></td>
      <td style="color:#bbb">—</td>
      <td><input id="edit-tracker-${t.id}" value="${t.tracker_serial || ''}" style="width:80px" onclick="event.stopPropagation()"></td>
      <td><input id="edit-lock-${t.id}" value="${t.lock_serial || ''}" style="width:80px" onclick="event.stopPropagation()"></td>
    ` : `
      <td>${t.ezpu_serial || '—'}</td>
      <td style="color:#bbb">—</td>
      <td>${t.tracker_serial || '—'}</td>
      <td>${t.lock_serial || '—'}</td>
    `;

    mainRow.innerHTML = `
      <td>${t.id}</td>
      <td>${t.board_number || '—'}</td>
      ${mainDeviceCells}
      <td>${pickupsList}</td>
      <td>${dropoffsList}</td>
      <td>${fmtDate(t.hang_datetime)}</td>
      <td>${t.removal_datetime ? fmtDate(t.removal_datetime) : '—'}</td>
      <td><span class="trip-status ${statusClass(t.status)}">${t.status}</span></td>
      <td class="actions-cell">${mainActions.join('')}</td>
    `;
    body.appendChild(mainRow);

    // ---------- Плечи маршрута ----------
    legs.forEach((leg, i) => {
      const tr = document.createElement('tr');
      tr.className = 'device-row trip-leg-row';
      tr.onclick = () => openTripDetail(t.id);
      const num = `${t.id}.${i + 1}`;
      const legDone = leg.toStop && leg.toStop.status === 'исполнено';
      const legStatus = leg.toStop ? leg.toStop.status : '—';
      const zpuStopId = leg.fromStop ? leg.fromStop.id : null;
      const editingZpu = zpuStopId && editingZpuStops.has(zpuStopId);

      const legActions = [];
      if (leg.toStop && !legDone) {
        legActions.push(`<button class="secondary" onclick="event.stopPropagation(); completeStop(${t.id}, ${leg.toStop.id})">Исполнено</button>`);
      }
      if (zpuStopId && !editingZpu) {
        const zpuLabel = leg.fromStop.zpu_number ? 'Изменить № ЗПУ' : 'Назначить № ЗПУ';
        legActions.push(`<button class="secondary" onclick="event.stopPropagation(); startEditZpu(${zpuStopId})">${zpuLabel}</button>`);
      }
      if (editingZpu) {
        legActions.push(`<button onclick="event.stopPropagation(); saveZpu(${t.id}, ${zpuStopId})">Сохранить ЗПУ</button>`);
        legActions.push(`<button class="secondary" onclick="event.stopPropagation(); cancelEditZpu(${zpuStopId})">Отмена</button>`);
      }

      tr.innerHTML = `
        <td>${num}</td>
        <td></td>
        <td>${t.ezpu_serial || '—'}</td>
        <td>${zpuCell(leg, editingZpu, zpuStopId)}</td>
        <td>${t.tracker_serial || '—'}</td>
        <td>${t.lock_serial || '—'}</td>
        <td>${leg.from}</td>
        <td>${leg.to}</td>
        <td></td>
        <td>${legDone && leg.toStop.completed_at ? fmtDate(leg.toStop.completed_at) : '—'}</td>
        <td>${leg.toStop
          ? '<span class="trip-status ' + statusClass(legStatus) + '">' + legStatus + '</span>'
          : ''}</td>
        <td class="actions-cell">${legActions.join('')}</td>
      `;
      body.appendChild(tr);
    });
  });
  document.getElementById('trip-count-label').textContent = trips.length + ' рейсов' + (hideCompleted ? ' (без завершённых)' : '');
}

function zpuCell(leg, editingZpu, zpuStopId) {
  if (!zpuStopId) return '—';
  if (editingZpu) {
    return `<input id="edit-zpu-stop-${zpuStopId}" value="${leg.fromStop.zpu_number || ''}" style="width:90px" onclick="event.stopPropagation()">`;
  }
  return leg.fromStop.zpu_number || '—';
}

function startEditZpu(stopId) {
  editingZpuStops.add(stopId);
  loadTrips();
}
function cancelEditZpu(stopId) {
  editingZpuStops.delete(stopId);
  loadTrips();
}
async function saveZpu(tripId, stopId) {
  const value = document.getElementById('edit-zpu-stop-' + stopId).value.trim();
  const r = await fetch(`/trips/${tripId}/stops/${stopId}/zpu`, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({zpu_number: value}),
  });
  if (r.status === 200) {
    editingZpuStops.delete(stopId);
    loadTrips();
  } else {
    const data = await r.json();
    alert(data.error || 'Ошибка сохранения ЗПУ');
  }
}

async function openTripDetail(id) {
  const panel = document.getElementById('trip-detail-panel');
  panel.style.display = 'block';
  document.getElementById('trip-detail-title').textContent = 'Рейс № ' + id;
  document.getElementById('trip-detail-body').innerHTML = 'Загрузка...';
  panel.scrollIntoView({behavior: 'smooth'});
  const r = await fetch('/trips/' + id);
  const t = await r.json();
  if (!t.stops) {
    document.getElementById('trip-detail-body').innerHTML = 'Не удалось загрузить рейс';
    return;
  }
  const rows = t.stops.map(s => {
    const done = s.status === 'исполнено';
    const arrived = !!s.arrived_at;
    let arriveBtn = '';
    if (s.stop_type === 'выгрузка' && !done) {
      arriveBtn = arrived
        ? '<span style="color:#1e40af; margin-right:8px">прибыл ' + fmtDate(s.arrived_at) + '</span>'
        : `<button class="secondary" onclick="markArrived(${id}, ${s.id})" style="margin-right:8px">Прибыл на базу</button>`;
    }
    const btn = done
      ? '<span style="color:#166534">✓ ' + (s.stop_type === 'выгрузка' ? 'снято ' : '') + fmtDate(s.completed_at) + '</span>'
      : `<button class="secondary" onclick="completeStop(${id}, ${s.id})">${s.stop_type === 'выгрузка' ? 'Снято' : 'Исполнено'}</button>`;
    return `<div class="hist-row"><b>${id}.${s.sequence}</b> ${s.stop_type} · ${s.location} ${arriveBtn}${btn}</div>`;
  }).join('');
  document.getElementById('trip-detail-body').innerHTML = `
    <div style="margin-bottom:10px">
      Клиент: ${t.client || '—'} · Подрядчик: ${t.contractor || '—'} · № борта: ${t.board_number || '—'}<br>
      ЭЗПУ: ${t.ezpu_serial || '—'} · № ЗПУ: ${t.zpu_number || '—'} · Трекер: ${t.tracker_serial || '—'} · Закладка: ${t.lock_serial || '—'} ·
      Статус: <span class="trip-status ${statusClass(t.status)}">${t.status}</span>
    </div>
    ${rows}
  `;
}
function closeTripDetail() { document.getElementById('trip-detail-panel').style.display = 'none'; }

async function markArrived(tripId, stopId) {
  const r = await fetch(`/trips/${tripId}/stops/${stopId}/arrive`, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({}),
  });
  if (r.status === 200) {
    openTripDetail(tripId);
  } else {
    const data = await r.json();
    alert(data.error || 'Ошибка');
  }
}

async function completeStop(tripId, stopId) {
  const r = await fetch(`/trips/${tripId}/stops/${stopId}/complete`, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({}),
  });
  if (r.status === 200) {
    const panel = document.getElementById('trip-detail-panel');
    if (panel.style.display === 'block' && document.getElementById('trip-detail-title').textContent === 'Рейс № ' + tripId) {
      openTripDetail(tripId);
    }
    loadTrips();
  } else {
    const data = await r.json();
    alert(data.error || 'Ошибка');
  }
}

const MEGAPOLIS_NAME = 'ТОО ТК Мегаполис Казахстан';
let stopCounter = 0;

function openTripForm() {
  document.getElementById('trip-form-panel').style.display = 'block';
  if (!document.querySelector('#pickups-list .waypoint-row')) addStop('pickups-list');
  if (!document.querySelector('#dropoffs-list .waypoint-row')) addStop('dropoffs-list');
  document.getElementById('trip-form-panel').scrollIntoView({behavior: 'smooth'});
}
function closeTripForm() { document.getElementById('trip-form-panel').style.display = 'none'; }

function onContractorChange() {
  const contractor = document.getElementById('tr-contractor').value;
  const label = document.getElementById('tr-tracker-label');
  if (contractor === MEGAPOLIS_NAME) {
    label.textContent = 'Серийный номер трекера * (обязателен для Мегаполис)';
  } else {
    label.textContent = 'Серийный номер трекера';
  }
}

function addStop(containerId, value) {
  stopCounter++;
  const id = 'stop-' + stopCounter;
  const placeholder = containerId === 'pickups-list' ? 'Склад погрузки' : 'Пункт выгрузки';
  const div = document.createElement('div');
  div.className = 'waypoint-row';
  div.id = id;
  div.innerHTML = `
    <input placeholder="${placeholder}" value="${value || ''}">
    <button class="secondary" type="button" onclick="document.getElementById('${id}').remove()">×</button>
  `;
  document.getElementById(containerId).appendChild(div);
}

function collectStops(containerId) {
  return Array.from(document.querySelectorAll('#' + containerId + ' .waypoint-row input'))
    .map(el => el.value.trim())
    .filter(v => v);
}

function resetTripForm() {
  ['tr-client','tr-contractor','tr-board','tr-ezpu','tr-zpu','tr-tracker','tr-lock','tr-origin','tr-notes'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('pickups-list').innerHTML = '';
  document.getElementById('dropoffs-list').innerHTML = '';
  addStop('pickups-list');
  addStop('dropoffs-list');
  onContractorChange();
}

async function submitTrip() {
  const msg = document.getElementById('trip-msg');
  msg.textContent = '';
  msg.className = '';

  const pickups = collectStops('pickups-list');
  const dropoffs = collectStops('dropoffs-list');
  const contractor = document.getElementById('tr-contractor').value;
  const ezpu = document.getElementById('tr-ezpu').value.trim();
  const zpu = document.getElementById('tr-zpu').value.trim();
  const tracker = document.getElementById('tr-tracker').value.trim();
  const lock = document.getElementById('tr-lock').value.trim();

  if (!pickups.length) {
    msg.textContent = 'Укажите хотя бы один склад погрузки';
    msg.className = 'err';
    return;
  }
  if (!dropoffs.length) {
    msg.textContent = 'Укажите хотя бы один пункт выгрузки';
    msg.className = 'err';
    return;
  }
  if (contractor === MEGAPOLIS_NAME && (ezpu || tracker) && !tracker) {
    msg.textContent = 'Для подрядчика «' + MEGAPOLIS_NAME + '» обязателен номер трекера';
    msg.className = 'err';
    return;
  }

  const payload = {
    client: document.getElementById('tr-client').value.trim(),
    contractor: contractor,
    board_number: document.getElementById('tr-board').value.trim(),
    pickups: pickups,
    dropoffs: dropoffs,
    ezpu_serial: ezpu,
    zpu_number: zpu,
    tracker_serial: tracker,
    lock_serial: lock,
    origin_city: document.getElementById('tr-origin').value.trim(),
    notes: document.getElementById('tr-notes').value.trim(),
  };
  const r = await fetch('/trips', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload),
  });
  const data = await r.json();
  if (r.status === 201) {
    msg.textContent = 'Рейс создан, id ' + data.id;
    msg.className = 'ok';
    resetTripForm();
    loadTrips();
  } else {
    msg.textContent = data.error || 'Ошибка сохранения';
    msg.className = 'err';
  }
}

async function closeTrip(id) {
  const location = prompt('Местонахождение при снятии пломбы:');
  if (location === null) return;
  const r = await fetch('/trips/' + id + '/close', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({location: location}),
  });
  if (r.status === 200) {
    loadTrips();
  } else {
    const data = await r.json();
    alert(data.error || 'Ошибка закрытия рейса');
  }
}

function startEditDevice(id) {
  editingTrips.add(id);
  loadTrips();
}

function cancelEditDevice(id) {
  editingTrips.delete(id);
  loadTrips();
}

async function saveDevice(id) {
  const ezpu = document.getElementById('edit-ezpu-' + id).value.trim();
  const tracker = document.getElementById('edit-tracker-' + id).value.trim();
  const lock = document.getElementById('edit-lock-' + id).value.trim();
  if (!ezpu && !tracker) {
    alert('Укажите хотя бы ЭЗПУ или трекер');
    return;
  }
  const r = await fetch('/trips/' + id + '/assign-device', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ezpu_serial: ezpu, tracker_serial: tracker, lock_serial: lock}),
  });
  if (r.status === 200) {
    editingTrips.delete(id);
    loadTrips();
  } else {
    const data = await r.json();
    alert(data.error || 'Ошибка сохранения устройства');
  }
}

let allDevicesCache = [];
let actCustomLineCounter = 0;

function onActCounterpartyPreset() {
  const preset = document.getElementById('act-counterparty-preset').value;
  if (preset) document.getElementById('act-counterparty').value = preset;
}

async function loadAllDevicesCache() {
  if (allDevicesCache.length) return;
  const r = await fetch('/devices?limit=2000');
  const data = await r.json();
  allDevicesCache = data.devices;
}

function renderActDeviceSearch() {
  const q = document.getElementById('act-device-search').value.trim().toLowerCase();
  const box = document.getElementById('act-device-search-results');
  if (!q) { box.innerHTML = ''; return; }
  const matches = allDevicesCache.filter(d => d.serial_number.toLowerCase().includes(q)).slice(0, 8);
  if (!matches.length) { box.innerHTML = '<div style="padding:6px 0; color:#888; font-size:13px">Ничего не найдено</div>'; return; }
  box.innerHTML = matches.map(d => `
    <div class="stop-line" style="justify-content:space-between; padding:4px 0">
      <span>${d.serial_number} <span style="color:#888; font-size:12px">(${d.device_type === 'ezpu' ? 'ЭЗПУ' : d.device_type === 'tracker' ? 'трекер' : 'закладка'}, ${d.status})</span></span>
      <button class="secondary" type="button" onclick="addDeviceToAct('${d.serial_number}', '${d.device_type}')">+ Добавить</button>
    </div>
  `).join('');
}

function addDeviceToAct(serial, deviceType) {
  const fieldId = deviceType === 'ezpu' ? 'act-ezpu-list' : 'act-tracker-list';
  const el = document.getElementById(fieldId);
  const current = el.value.split(/[\n,]+/).map(s => s.trim()).filter(s => s);
  if (!current.includes(serial)) current.push(serial);
  el.value = current.join(', ');
  document.getElementById('act-device-search').value = '';
  document.getElementById('act-device-search-results').innerHTML = '';
}

function addActCustomLine(name, qty, price) {
  actCustomLineCounter++;
  const id = 'act-custom-' + actCustomLineCounter;
  const div = document.createElement('div');
  div.className = 'waypoint-row';
  div.id = id;
  div.innerHTML = `
    <input placeholder="Наименование (напр. ЗПУ одноразовые)" id="${id}-name" value="${name || ''}" style="flex:2">
    <input placeholder="Кол-во" type="number" id="${id}-qty" value="${qty || 1}" style="width:70px">
    <input placeholder="Цена за ед." type="number" id="${id}-price" value="${price || ''}" style="width:100px">
    <button class="secondary" type="button" onclick="document.getElementById('${id}').remove()">×</button>
  `;
  document.getElementById('act-custom-lines').appendChild(div);
}

function parseSerialsField(id) {
  return document.getElementById(id).value.split(/[\n,]+/).map(s => s.trim()).filter(s => s);
}

async function submitAct() {
  const msg = document.getElementById('act-msg');
  msg.textContent = '';
  msg.className = '';

  const counterparty = document.getElementById('act-counterparty').value.trim();
  const actDate = document.getElementById('act-date').value;
  if (!counterparty) { msg.textContent = 'Укажите заказчика'; msg.className = 'err'; return; }
  if (!actDate) { msg.textContent = 'Укажите дату акта'; msg.className = 'err'; return; }

  const lines = [];
  const ezpuSerials = parseSerialsField('act-ezpu-list');
  const ezpuPrice = parseFloat(document.getElementById('act-ezpu-price').value) || null;
  ezpuSerials.forEach(s => lines.push({item_name: 'ЭЗПУ Сириус', serials: [s], qty: 1, unit_price: ezpuPrice, device_type: 'ezpu'}));

  const trackerSerials = parseSerialsField('act-tracker-list');
  const trackerPrice = parseFloat(document.getElementById('act-tracker-price').value) || null;
  if (trackerSerials.length) lines.push({item_name: 'Concox AT4', serials: trackerSerials, qty: trackerSerials.length, unit_price: trackerPrice, device_type: 'tracker'});

  document.querySelectorAll('#act-custom-lines .waypoint-row').forEach(row => {
    const id = row.id;
    const name = document.getElementById(id + '-name').value.trim();
    const qty = parseInt(document.getElementById(id + '-qty').value) || 0;
    const price = parseFloat(document.getElementById(id + '-price').value) || null;
    if (name && qty > 0) lines.push({item_name: name, serials: [], qty: qty, unit_price: price});
  });

  if (!lines.length) { msg.textContent = 'Добавьте хотя бы одну позицию'; msg.className = 'err'; return; }

  const payload = {
    direction: document.getElementById('act-direction').value,
    counterparty: counterparty,
    counterparty_label: document.getElementById('act-counterparty-label').value.trim(),
    act_date: actDate,
    contract_number: document.getElementById('act-contract-number').value.trim(),
    contract_date: document.getElementById('act-contract-date').value,
    director_name: document.getElementById('act-director').value.trim(),
    lines: lines,
  };

  const r = await fetch('/acts', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload),
  });
  const data = await r.json();
  if (r.status === 201) {
    msg.textContent = 'Акт № ' + data.act_number + ' создан';
    msg.className = 'ok';
    ['act-ezpu-list','act-tracker-list'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('act-custom-lines').innerHTML = '';
    loadActs();
  } else {
    msg.textContent = data.error || 'Ошибка создания акта';
    msg.className = 'err';
  }
}

async function loadActs() {
  document.getElementById('acts-count-label').textContent = 'Загрузка...';
  const r = await fetch('/acts');
  const data = await r.json();
  const body = document.getElementById('acts-body');
  body.innerHTML = '';
  data.acts.forEach(a => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${a.act_number}</td>
      <td>${a.direction}</td>
      <td>${a.counterparty || '—'}${a.counterparty_label ? ' (' + a.counterparty_label + ')' : ''}</td>
      <td>${a.act_date || '—'}</td>
      <td>${a.lines_count}</td>
      <td>${fmtDate(a.generated_at)}</td>
      <td><a href="/acts/${a.id}/download"><button class="secondary">Скачать .docx</button></a></td>
    `;
    body.appendChild(tr);
  });
  document.getElementById('acts-count-label').textContent = data.count + ' актов';
}

const REPORT_COMPANIES = ['ТОО ТК Мегаполис Казахстан', 'ТОО ТМЕ', 'ТОО СОП ТЖК'];
const REPORT_MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
let currentReportCompany = null;

function renderReportsCompanyList() {
  const box = document.getElementById('reports-company-list');
  box.innerHTML = REPORT_COMPANIES.map(c => `<button onclick="openReportDetail('${c}')">${c}</button>`).join('');
}

function openReportDetail(company) {
  currentReportCompany = company;
  document.getElementById('reports-detail-panel').style.display = 'block';
  document.getElementById('reports-detail-title').textContent = 'Отчёт по ЭЗПУ — ' + company;

  const monthSelect = document.getElementById('report-month');
  if (!monthSelect.options.length) {
    monthSelect.innerHTML = REPORT_MONTHS.map((m, i) => `<option value="${i + 1}">${m}</option>`).join('');
  }
  const now = new Date();
  monthSelect.value = now.getMonth() + 1;
  document.getElementById('report-year').value = now.getFullYear();

  document.getElementById('reports-detail-panel').scrollIntoView({behavior: 'smooth'});
  loadReportDetail();
}

function closeReportDetail() {
  document.getElementById('reports-detail-panel').style.display = 'none';
}

async function loadReportDetail() {
  if (!currentReportCompany) return;
  const year = document.getElementById('report-year').value;
  const month = document.getElementById('report-month').value;
  const rate = document.getElementById('report-rate').value;
  const params = new URLSearchParams({contractor: currentReportCompany, year, month, rate});
  const r = await fetch('/reports/ezpu-billing?' + params.toString());
  const data = await r.json();
  const body = document.getElementById('report-rows-body');
  body.innerHTML = '';
  if (!data.rows || !data.rows.length) {
    body.innerHTML = '<tr><td colspan="10" style="color:#888">Нет закрытых рейсов с ЭЗПУ за этот период</td></tr>';
    document.getElementById('report-total').textContent = '';
    return;
  }
  data.rows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.num}</td>
      <td>${row.board_number || '—'}</td>
      <td>${row.ezpu_serial}</td>
      <td>${row.origin || '—'}</td>
      <td>${row.destination || '—'}</td>
      <td>${fmtDate(row.hang_datetime)}</td>
      <td>${fmtDate(row.removal_datetime)}</td>
      <td>${row.days}</td>
      <td>${row.rate}</td>
      <td>${row.amount.toLocaleString('ru-RU')}</td>
    `;
    body.appendChild(tr);
  });
  document.getElementById('report-total').textContent =
    'Итого: ' + data.total_amount.toLocaleString('ru-RU') + ' тенге (' + data.rows.length + ' рейсов)';
}

loadDevices();
</script>
</body>
</html>"""


def _get_or_create_client(cur, name):
    if not name:
        return None
    cur.execute("SELECT id FROM clients WHERE name = %s", (name,))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute("INSERT INTO clients (name) VALUES (%s) RETURNING id", (name,))
    return cur.fetchone()[0]


def db_create_trip(payload):
    conn = get_connection()
    try:
        cur = conn.cursor()
        client_id = _get_or_create_client(cur, payload["client"])
        contractor_id = _get_or_create_party(cur, payload["contractor"]) if payload["contractor"] else None
        ezpu_id = _get_or_create_device(cur, payload["ezpu_serial"], "ezpu") if payload["ezpu_serial"] else None
        tracker_id = _get_or_create_device(cur, payload["tracker_serial"], "tracker") if payload["tracker_serial"] else None
        lock_id = _get_or_create_device(cur, payload["lock_serial"], "lock") if payload["lock_serial"] else None

        pickups = [p.strip() for p in payload["pickups"] if (p or "").strip()]
        dropoffs = [d.strip() for d in payload["dropoffs"] if (d or "").strip()]

        status = "в пути" if (ezpu_id or tracker_id or lock_id) else "запланирован"

        cur.execute(
            """
            INSERT INTO trips
                (client_id, contractor_id, board_number, warehouse, ezpu_device_id, tracker_device_id,
                 lock_device_id, zpu_number, origin_city, destination_city, hang_datetime, status, notes)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id
            """,
            (client_id, contractor_id, payload["board_number"], pickups[0] if pickups else None, ezpu_id, tracker_id,
             lock_id, payload["zpu_number"], payload["origin_city"], dropoffs[-1] if dropoffs else None,
             payload["hang_datetime"], status, payload["notes"]),
        )
        trip_id = cur.fetchone()[0]

        seq = 1
        first_pickup_id = None
        for loc in pickups:
            cur.execute(
                """
                INSERT INTO trip_stops (trip_id, stop_type, sequence, location, status)
                VALUES (%s, 'погрузка', %s, %s, 'ожидание') RETURNING id
                """,
                (trip_id, seq, loc),
            )
            sid = cur.fetchone()[0]
            if first_pickup_id is None:
                first_pickup_id = sid
            seq += 1
        for loc in dropoffs:
            cur.execute(
                "INSERT INTO trip_stops (trip_id, stop_type, sequence, location, status) VALUES (%s, 'выгрузка', %s, %s, 'ожидание')",
                (trip_id, seq, loc),
            )
            seq += 1

        if ezpu_id and first_pickup_id:
            cur.execute(
                "UPDATE trip_stops SET status = 'исполнено', completed_at = %s WHERE id = %s",
                (payload["hang_datetime"], first_pickup_id),
            )
            cur.execute(
                """
                INSERT INTO operations
                    (device_id, trip_id, operation_type, location, operation_dt, document_ref)
                VALUES (%s, %s, 'навешивание', %s, %s, %s)
                """,
                (ezpu_id, trip_id, pickups[0], payload["hang_datetime"], payload["board_number"]),
            )
            cur.execute(
                """
                UPDATE devices SET current_location = %s, last_operation_at = %s, updated_at = now()
                WHERE id = %s
                """,
                (pickups[0], payload["hang_datetime"], ezpu_id),
            )

        conn.commit()
        return trip_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_assign_trip_device(trip_id, ezpu_serial, tracker_serial, lock_serial, zpu_number, assign_dt):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT status FROM trips WHERE id = %s", (trip_id,))
        row = cur.fetchone()
        if not row:
            return None
        was_planned = row[0] == "запланирован"

        stop_row = None
        location = None
        if was_planned:
            cur.execute(
                """
                SELECT id, location FROM trip_stops
                WHERE trip_id = %s AND stop_type = 'погрузка' AND status = 'ожидание'
                ORDER BY sequence LIMIT 1
                """,
                (trip_id,),
            )
            stop_row = cur.fetchone()
            location = stop_row[1] if stop_row else None

        ezpu_id = _get_or_create_device(cur, ezpu_serial, "ezpu") if ezpu_serial else None
        tracker_id = _get_or_create_device(cur, tracker_serial, "tracker") if tracker_serial else None
        lock_id = _get_or_create_device(cur, lock_serial, "lock") if lock_serial else None

        cur.execute(
            """
            UPDATE trips SET
                ezpu_device_id = COALESCE(%s, ezpu_device_id),
                tracker_device_id = COALESCE(%s, tracker_device_id),
                lock_device_id = COALESCE(%s, lock_device_id),
                zpu_number = COALESCE(%s, zpu_number),
                status = 'в пути',
                hang_datetime = CASE WHEN %s THEN %s ELSE hang_datetime END,
                updated_at = now()
            WHERE id = %s
            """,
            (ezpu_id, tracker_id, lock_id, zpu_number, was_planned, assign_dt, trip_id),
        )

        if was_planned and stop_row:
            cur.execute(
                "UPDATE trip_stops SET status = 'исполнено', completed_at = %s WHERE id = %s",
                (assign_dt, stop_row[0]),
            )

        if was_planned and ezpu_id:
            cur.execute(
                """
                INSERT INTO operations
                    (device_id, trip_id, operation_type, location, operation_dt)
                VALUES (%s, %s, 'навешивание', %s, %s)
                """,
                (ezpu_id, trip_id, location, assign_dt),
            )
            cur.execute(
                """
                UPDATE devices SET current_location = %s, last_operation_at = %s, updated_at = now()
                WHERE id = %s
                """,
                (location, assign_dt, ezpu_id),
            )

        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_mark_stop_arrived(trip_id, stop_id, arrived_dt):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "UPDATE trip_stops SET arrived_at = %s WHERE id = %s AND trip_id = %s RETURNING id",
            (arrived_dt, stop_id, trip_id),
        )
        row = cur.fetchone()
        conn.commit()
        return row is not None
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_auto_mark_arrival_by_board(board_number, arrived_dt, location_hint=None):
    """Для автоматики с Wialon: по номеру борта находит активный рейс
    (status='в пути') и ближайшую ещё не пройденную точку выгрузки,
    проставляет ей arrived_at. Если указан location_hint (название
    геозоны/базы) - сначала пробует найти точку с похожим названием,
    иначе берёт просто самую раннюю необработанную точку выгрузки."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id FROM trips
            WHERE board_number = %s AND status = 'в пути'
            ORDER BY hang_datetime DESC LIMIT 1
            """,
            (board_number,),
        )
        trip = cur.fetchone()
        if not trip:
            return {"matched": False, "reason": f"Нет активного рейса с бортом {board_number}"}
        trip_id = trip[0]

        stop_id = None
        if location_hint:
            cur.execute(
                """
                SELECT id FROM trip_stops
                WHERE trip_id = %s AND stop_type = 'выгрузка' AND arrived_at IS NULL
                  AND location ILIKE %s
                ORDER BY sequence LIMIT 1
                """,
                (trip_id, f"%{location_hint}%"),
            )
            row = cur.fetchone()
            if row:
                stop_id = row[0]

        if stop_id is None:
            cur.execute(
                """
                SELECT id FROM trip_stops
                WHERE trip_id = %s AND stop_type = 'выгрузка' AND arrived_at IS NULL
                ORDER BY sequence LIMIT 1
                """,
                (trip_id,),
            )
            row = cur.fetchone()
            if row:
                stop_id = row[0]

        if stop_id is None:
            return {"matched": False, "reason": f"У рейса {trip_id} нет необработанных точек выгрузки"}

        cur.execute(
            "UPDATE trip_stops SET arrived_at = %s WHERE id = %s",
            (arrived_dt, stop_id),
        )
        conn.commit()
        return {"matched": True, "trip_id": trip_id, "stop_id": stop_id}
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_store_wialon_webhook_event(raw_body):
    conn = get_connection()
    try:
        cur = conn.cursor()
        external_id = None
        event_type = None
        if isinstance(raw_body, dict):
            event_type = raw_body.get("Type") or raw_body.get("type") or raw_body.get("eventType")
            raw_id = raw_body.get("Id") or raw_body.get("id")
            try:
                external_id = int(raw_id) if raw_id is not None else None
            except (ValueError, TypeError):
                external_id = None
        cur.execute(
            """
            INSERT INTO raw_events (source, event_type, external_id, event_dt, raw_payload, processed)
            VALUES ('wialon', %s, %s, now(), %s, false)
            """,
            (event_type, external_id, json.dumps(raw_body, ensure_ascii=False)),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_complete_stop(trip_id, stop_id, completed_dt, location_override):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT stop_type, location FROM trip_stops WHERE id = %s AND trip_id = %s",
            (stop_id, trip_id),
        )
        stop = cur.fetchone()
        if not stop:
            return None
        stop_type, default_location = stop
        location = location_override or default_location

        cur.execute(
            "UPDATE trip_stops SET status = 'исполнено', completed_at = %s, location = %s WHERE id = %s",
            (completed_dt, location, stop_id),
        )

        trip_closed = False
        if stop_type == "выгрузка":
            cur.execute(
                "SELECT COUNT(*) FROM trip_stops WHERE trip_id = %s AND stop_type = 'выгрузка' AND status != 'исполнено'",
                (trip_id,),
            )
            remaining = cur.fetchone()[0]
            if remaining == 0:
                cur.execute("SELECT ezpu_device_id FROM trips WHERE id = %s", (trip_id,))
                ezpu_id = cur.fetchone()[0]
                cur.execute(
                    "UPDATE trips SET status = 'снят', removal_datetime = %s, updated_at = now() WHERE id = %s",
                    (completed_dt, trip_id),
                )
                if ezpu_id:
                    cur.execute(
                        """
                        INSERT INTO operations (device_id, trip_id, operation_type, location, operation_dt)
                        VALUES (%s, %s, 'снятие', %s, %s)
                        """,
                        (ezpu_id, trip_id, location, completed_dt),
                    )
                    cur.execute(
                        "UPDATE devices SET current_location = %s, last_operation_at = %s, updated_at = now() WHERE id = %s",
                        (location, completed_dt, ezpu_id),
                    )
                trip_closed = True

        conn.commit()
        return {"trip_closed": trip_closed}
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_set_stop_zpu(trip_id, stop_id, zpu_number):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "UPDATE trip_stops SET zpu_number = %s WHERE id = %s AND trip_id = %s RETURNING id",
            (zpu_number, stop_id, trip_id),
        )
        row = cur.fetchone()
        conn.commit()
        return row is not None
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_get_trip(trip_id):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT t.id, c.name, ct.name, t.board_number, t.warehouse,
                   de.serial_number, dt.serial_number, dl.serial_number, t.zpu_number,
                   t.origin_city, t.destination_city, t.hang_datetime,
                   t.arrival_at_unload_datetime, t.removal_datetime, t.status, t.notes
            FROM trips t
            LEFT JOIN clients c ON c.id = t.client_id
            LEFT JOIN parties ct ON ct.id = t.contractor_id
            LEFT JOIN devices de ON de.id = t.ezpu_device_id
            LEFT JOIN devices dt ON dt.id = t.tracker_device_id
            LEFT JOIN devices dl ON dl.id = t.lock_device_id
            WHERE t.id = %s
            """,
            (trip_id,),
        )
        r = cur.fetchone()
        if not r:
            return None
        cur.execute(
            "SELECT id, stop_type, sequence, location, status, completed_at, zpu_number, arrived_at FROM trip_stops WHERE trip_id = %s ORDER BY sequence",
            (trip_id,),
        )
        stops = [
            {"id": s[0], "stop_type": s[1], "sequence": s[2], "location": s[3], "status": s[4],
             "completed_at": s[5], "zpu_number": s[6], "arrived_at": s[7]}
            for s in cur.fetchall()
        ]
        return {
            "id": r[0], "client": r[1], "contractor": r[2], "board_number": r[3], "warehouse": r[4],
            "ezpu_serial": r[5], "tracker_serial": r[6], "lock_serial": r[7], "zpu_number": r[8],
            "origin_city": r[9], "destination_city": r[10], "hang_datetime": r[11],
            "arrival_at_unload_datetime": r[12], "removal_datetime": r[13],
            "status": r[14], "notes": r[15], "stops": stops,
        }
    finally:
        conn.close()


def db_list_trips(status=None, client=None, limit=200):
    conn = get_connection()
    try:
        cur = conn.cursor()
        query = """
            SELECT t.id, c.name, ct.name, t.board_number, de.serial_number, dt.serial_number, dl.serial_number,
                   t.zpu_number, t.origin_city, t.destination_city, t.hang_datetime, t.status
            FROM trips t
            LEFT JOIN clients c ON c.id = t.client_id
            LEFT JOIN parties ct ON ct.id = t.contractor_id
            LEFT JOIN devices de ON de.id = t.ezpu_device_id
            LEFT JOIN devices dt ON dt.id = t.tracker_device_id
            LEFT JOIN devices dl ON dl.id = t.lock_device_id
            WHERE 1=1
        """
        params = []
        if status:
            query += " AND t.status = %s"
            params.append(status)
        if client:
            query += " AND c.name = %s"
            params.append(client)
        query += " ORDER BY t.hang_datetime DESC LIMIT %s"
        params.append(limit)
        cur.execute(query, params)
        rows = cur.fetchall()
        trip_ids = [r[0] for r in rows]

        stops_by_trip = {}
        if trip_ids:
            cur.execute(
                """
                SELECT id, trip_id, stop_type, sequence, location, status, zpu_number, completed_at, arrived_at FROM trip_stops
                WHERE trip_id = ANY(%s) ORDER BY sequence
                """,
                (trip_ids,),
            )
            for stop_id, trip_id, stop_type, sequence, location, st_status, zpu, completed_at, arrived_at in cur.fetchall():
                stops_by_trip.setdefault(trip_id, {"pickups": [], "dropoffs": []})
                key = "pickups" if stop_type == "погрузка" else "dropoffs"
                stops_by_trip[trip_id][key].append({
                    "id": stop_id, "location": location, "status": st_status,
                    "sequence": sequence, "zpu_number": zpu, "completed_at": completed_at,
                    "arrived_at": arrived_at,
                })

        return [
            {
                "id": r[0], "client": r[1], "contractor": r[2], "board_number": r[3],
                "ezpu_serial": r[4], "tracker_serial": r[5], "lock_serial": r[6], "zpu_number": r[7],
                "origin_city": r[8], "destination_city": r[9],
                "hang_datetime": r[10], "status": r[11],
                "pickups": stops_by_trip.get(r[0], {}).get("pickups", []),
                "dropoffs": stops_by_trip.get(r[0], {}).get("dropoffs", []),
            }
            for r in rows
        ]
    finally:
        conn.close()


def db_close_trip(trip_id, removal_dt, location):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT ezpu_device_id FROM trips WHERE id = %s", (trip_id,))
        row = cur.fetchone()
        if not row:
            return None
        ezpu_id = row[0]

        cur.execute(
            """
            UPDATE trips SET removal_datetime = %s, status = 'снят', updated_at = now()
            WHERE id = %s
            """,
            (removal_dt, trip_id),
        )

        if ezpu_id:
            cur.execute(
                """
                INSERT INTO operations
                    (device_id, trip_id, operation_type, location, operation_dt)
                VALUES (%s, %s, 'снятие', %s, %s)
                """,
                (ezpu_id, trip_id, location, removal_dt),
            )
            cur.execute(
                """
                UPDATE devices SET current_location = %s, last_operation_at = %s, updated_at = now()
                WHERE id = %s
                """,
                (location, removal_dt, ezpu_id),
            )

        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()






# ---------- Акты приёма-передачи ----------

_MONTHS_RU = {
    1: "Января", 2: "Февраля", 3: "Марта", 4: "Апреля", 5: "Мая", 6: "Июня",
    7: "Июля", 8: "Августа", 9: "Сентября", 10: "Октября", 11: "Ноября", 12: "Декабря",
}

_ITEM_NAME_BY_DEVICE_TYPE = {
    "ezpu": "ЭЗПУ Сириус",
    "tracker": "Concox AT4",
    "lock": "Трекер-закладка",
}


def db_create_act(payload):
    conn = get_connection()
    try:
        cur = conn.cursor()
        counterparty_id = _get_or_create_party(cur, payload["counterparty"])

        cur.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM acts")
        act_number = str(cur.fetchone()[0])

        cur.execute(
            """
            INSERT INTO acts
                (act_type, direction, counterparty_id, counterparty_label,
                 contract_number, contract_date, act_date, act_number, director_name, generated_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, now())
            RETURNING id
            """,
            (payload["direction"], payload["direction"], counterparty_id, payload["counterparty_label"],
             payload["contract_number"], payload["contract_date"], payload["act_date"],
             act_number, payload["director_name"]),
        )
        act_id = cur.fetchone()[0]

        for line in payload["lines"]:
            cur.execute(
                """
                INSERT INTO act_lines (act_id, item_name, serials, unit, qty, unit_price)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (act_id, line["item_name"], ", ".join(line["serials"]) if line["serials"] else None,
                 line.get("unit", "шт"), line["qty"], line.get("unit_price")),
            )

        # если позиция ссылается на реально отслеживаемые устройства (ЭЗПУ/трекер/закладка),
        # заодно фиксируем операцию передачи/возврата в журнале
        from_name, to_name = ("ТМА", payload["counterparty"]) if payload["direction"] == "передача" \
            else (payload["counterparty"], "ТМА")
        from_id = _get_or_create_party(cur, from_name)
        to_id = _get_or_create_party(cur, to_name)
        for line in payload["lines"]:
            if line.get("device_type") in ("ezpu", "tracker", "lock"):
                for serial in line["serials"]:
                    device_id = _get_or_create_device(cur, serial, line["device_type"])
                    cur.execute(
                        """
                        INSERT INTO operations (device_id, operation_type, from_party_id, to_party_id, operation_dt, document_ref)
                        VALUES (%s, %s, %s, %s, %s, %s)
                        """,
                        (device_id, payload["direction"], from_id, to_id, payload["act_date"], f"Акт №{act_number}"),
                    )
                    cur.execute(
                        "UPDATE devices SET current_holder_id = %s, last_operation_at = %s, updated_at = now() WHERE id = %s",
                        (to_id, payload["act_date"], device_id),
                    )

        conn.commit()
        return act_id, act_number
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def db_get_act(act_id):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT a.id, a.direction, cp.name, a.counterparty_label, a.contract_number,
                   a.contract_date, a.act_date, a.act_number, a.director_name, a.generated_at
            FROM acts a
            LEFT JOIN parties cp ON cp.id = a.counterparty_id
            WHERE a.id = %s
            """,
            (act_id,),
        )
        r = cur.fetchone()
        if not r:
            return None
        cur.execute(
            "SELECT item_name, serials, unit, qty, unit_price FROM act_lines WHERE act_id = %s ORDER BY id",
            (act_id,),
        )
        lines = [
            {"item_name": x[0], "serials": x[1], "unit": x[2], "qty": x[3], "unit_price": x[4]}
            for x in cur.fetchall()
        ]
        return {
            "id": r[0], "direction": r[1], "counterparty": r[2], "counterparty_label": r[3],
            "contract_number": r[4], "contract_date": r[5], "act_date": r[6],
            "act_number": r[7], "director_name": r[8], "generated_at": r[9], "lines": lines,
        }
    finally:
        conn.close()


def db_list_raw_events(source=None, limit=20):
    conn = get_connection()
    try:
        cur = conn.cursor()
        query = "SELECT id, source, event_type, external_id, event_dt, raw_payload, processed, created_at FROM raw_events WHERE 1=1"
        params = []
        if source:
            query += " AND source = %s"
            params.append(source)
        query += " ORDER BY created_at DESC LIMIT %s"
        params.append(limit)
        cur.execute(query, params)
        return [
            {
                "id": r[0], "source": r[1], "event_type": r[2], "external_id": r[3],
                "event_dt": r[4], "raw_payload": r[5], "processed": r[6], "created_at": r[7],
            }
            for r in cur.fetchall()
        ]
    finally:
        conn.close()


def db_list_acts(limit=100):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT a.id, a.direction, cp.name, a.counterparty_label, a.act_date, a.act_number, a.generated_at,
                   COUNT(al.id)
            FROM acts a
            LEFT JOIN parties cp ON cp.id = a.counterparty_id
            LEFT JOIN act_lines al ON al.act_id = a.id
            GROUP BY a.id, cp.name
            ORDER BY a.generated_at DESC
            LIMIT %s
            """,
            (limit,),
        )
        return [
            {
                "id": r[0], "direction": r[1], "counterparty": r[2], "counterparty_label": r[3],
                "act_date": r[4], "act_number": r[5], "generated_at": r[6], "lines_count": r[7],
            }
            for r in cur.fetchall()
        ]
    finally:
        conn.close()


# ---------- Клиент BigLock API (только stdlib, без aiohttp) ----------

BIGLOCK_BASE_URL = "https://www.biglock.pro"


def _biglock_opener():
    login = os.environ.get("BIGLOCK_LOGIN")
    password = os.environ.get("BIGLOCK_PASSWORD")
    if not login or not password:
        raise RuntimeError("BIGLOCK_LOGIN / BIGLOCK_PASSWORD не заданы в переменных окружения")

    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    login_body = json.dumps({"Login": login, "Password": password, "Remember": True}).encode("utf-8")
    login_req = urllib.request.Request(
        BIGLOCK_BASE_URL + "/api/auth/login", data=login_body,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with opener.open(login_req, timeout=15) as resp:
        resp.read()  # авторизация задаёт cookie в opener, тело ответа не нужно

    return opener


def _biglock_post(opener, path, payload):
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        BIGLOCK_BASE_URL + path, data=body,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with opener.open(req, timeout=30) as resp:
        return json.loads(resp.read())


def _biglock_search_with_opener(opener, payload):
    return _biglock_post(opener, "/api/usernotifications/my/search", payload)


def biglock_search_notifications(payload):
    opener = _biglock_opener()
    return _biglock_search_with_opener(opener, payload)


def biglock_events_for_object(native_id, limit=200, codes=None, page_size=200, max_pages=25):
    """Тянет события BigLock постранично (сервер, похоже, игнорирует
    большой Limit за один запрос - надёжнее пролистать несколько
    страниц) и фильтрует по номеру борта (GuardedObject.NativeId).
    Диагностический режим: пока неизвестны точные значения
    DeviceEvent.Type для навешивания/снятия, отдаём все события как
    есть, чтобы можно было их опознать вручную."""
    result = []
    total_count = None
    scanned = 0
    page = 1
    opener = _biglock_opener()
    while scanned < limit and page <= max_pages:
        payload = {
            "Page": page,
            "Limit": page_size,
            "OrderBy": "CreateTimeDesc",
            "MediaType": "System",
        }
        if codes != []:
            payload["Codes"] = codes or ["LockEvent"]
        data = _biglock_search_with_opener(opener, payload)
        if total_count is None:
            total_count = data.get("TotalCount")
        items = data.get("Items", [])
        if not items:
            break
        for item in items:
            params = item.get("Parameters", {})
            guarded = params.get("GuardedObject", {})
            if native_id and guarded.get("NativeId") != native_id:
                continue
            device_event = params.get("DeviceEvent", {})
            result.append({
                "id": item.get("Id"),
                "create_time": item.get("CreateTime"),
                "event_type": device_event.get("Type"),
                "event_time": device_event.get("CreateTime"),
                "native_id": guarded.get("NativeId"),
                "ezpu_serial": (params.get("ElectricDevice") or {}).get("CaseId"),
                "zpu_number": (params.get("MechanicalDevice") or {}).get("CaseId"),
                "lat": (params.get("DevicePoint") or {}).get("Latitude"),
                "lon": (params.get("DevicePoint") or {}).get("Longitude"),
            })
        scanned += len(items)
        page += 1
        if len(items) < page_size:
            break  # достигли конца выдачи

    return {
        "total_count": total_count, "scanned": scanned, "pages_fetched": page - 1,
        "matched": len(result), "events": result,
    }


def biglock_device_sessions(ezpu_serial, limit=5000, codes=None, page_size=200, max_pages=25):
    """Ищет все события по конкретному серийнику ЭЗПУ (не по борту) и
    группирует их в 'сессии' - непрерывные периоды, когда устройство
    числилось под одним и тем же номером борта. Первое событие сессии
    ~= навешивание, последнее событие перед сменой борта (или до
    текущего момента) ~= снятие. Это обходной способ, пока точный
    DeviceEvent.Type для навешивания/снятия не опознан вручную."""
    matched = []
    total_count = None
    scanned = 0
    page = 1
    opener = _biglock_opener()
    while scanned < limit and page <= max_pages:
        payload = {
            "Page": page,
            "Limit": page_size,
            "OrderBy": "CreateTimeDesc",
            "MediaType": "System",
        }
        if codes != []:
            payload["Codes"] = codes or ["LockEvent"]
        data = _biglock_search_with_opener(opener, payload)
        if total_count is None:
            total_count = data.get("TotalCount")
        items = data.get("Items", [])
        if not items:
            break
        for item in items:
            params = item.get("Parameters", {})
            electric = params.get("ElectricDevice") or {}
            if electric.get("CaseId") != ezpu_serial:
                continue
            guarded = params.get("GuardedObject", {})
            device_event = params.get("DeviceEvent", {})
            matched.append({
                "create_time": item.get("CreateTime"),
                "event_type": device_event.get("Type"),
                "native_id": guarded.get("NativeId"),
                "zpu_number": (params.get("MechanicalDevice") or {}).get("CaseId"),
            })
        scanned += len(items)
        page += 1
        if len(items) < page_size:
            break

    # события шли по убыванию времени (CreateTimeDesc) - развернём по возрастанию для группировки
    matched.sort(key=lambda e: e["create_time"])

    sessions = []
    current = None
    for e in matched:
        if current is None or current["native_id"] != e["native_id"]:
            if current:
                sessions.append(current)
            current = {
                "native_id": e["native_id"], "zpu_number": e["zpu_number"],
                "start_time": e["create_time"], "end_time": e["create_time"],
                "event_count": 1,
            }
        else:
            current["end_time"] = e["create_time"]
            current["event_count"] += 1
    if current:
        sessions.append(current)

    return {
        "ezpu_serial": ezpu_serial, "total_count": total_count, "scanned": scanned,
        "pages_fetched": page - 1, "matched_events": len(matched), "sessions": sessions,
    }


def biglock_device_status(case_id):
    """Проверяет текущий статус конкретной пломбы (по CaseId, напр.
    GNS10759) через связку electricdevices -> devicepackets - точная
    логика подтверждена самой поддержкой BigLock: находит DeviceId,
    смотрит последний пакет любого типа, и если в нём заполнен
    LockId ИЛИ DevicePointId - устройство на охране, оба пустые -
    свободно."""
    opener = _biglock_opener()

    devices_data = _biglock_post(opener, "/api/electricdevices/search", {
        "CaseId": case_id, "SkipCount": True,
    })
    items_dev = devices_data.get("Items", [])
    if not items_dev:
        return {"error": f"Устройство с CaseId={case_id} не найдено в BigLock", "case_id": case_id}
    device = items_dev[0]

    device_id = device.get("DeviceId")

    packets_data = _biglock_post(opener, "/api/devicepackets/search", {
        "DeviceId": device_id, "Limit": 1, "SkipCount": True, "OrderBy": "TimeDesc",
    })
    items = packets_data.get("Items", [])
    device_point_id = items[0].get("DevicePointId") if items else None
    lock_id = items[0].get("LockId") if items else None

    return {
        "case_id": case_id,
        "device_id": device_id,
        "client_name": device.get("ClientName"),
        "on_guard": bool(lock_id or device_point_id),
        "device_point_id": device_point_id,
        "lock_id": lock_id,
        "last_packet_raw": items[0] if items else None,
    }


def biglock_lock_status(mechanical_case_id=None, native_id=None, client_id=None,
                         is_released=None, limit=10, order_by="LockTimeDesc"):
    """Самый прямой способ узнать навешивание/снятие: ищет записи
    LockedDevice - там сразу есть LockTime (навешено), ReleaseTime
    (снято), IsReleased. Реально рабочие фильтры (подтверждено из
    структуры меню самого BigLock): MechanicalDeviceCaseId (номер
    разовой пломбы), ClientId (напр. 56 = ТОО Мегаполис КЗ),
    IsReleased (false = ещё в рейсе, не снята). NativeId (номер
    борта) и ElectricDeviceCaseId (серийник ЭЗПУ) BigLock тихо
    игнорирует - ими не пользуемся."""
    opener = _biglock_opener()
    payload = {"Limit": limit, "OrderBy": order_by}
    if mechanical_case_id:
        payload["MechanicalDeviceCaseId"] = mechanical_case_id
    if native_id:
        payload["NativeId"] = native_id
    if client_id is not None:
        payload["ClientId"] = client_id
    if is_released is not None:
        payload["IsReleased"] = is_released
    data = _biglock_post(opener, "/api/lockeddevices/search", payload)
    items = data.get("Items", [])
    return {
        "zpu_number": mechanical_case_id,
        "native_id": native_id,
        "client_id": client_id,
        "is_released": is_released,
        "total_count": data.get("TotalCount"),
        "records": [
            {
                "lock_time": it.get("LockTime"),
                "release_time": it.get("ReleaseTime"),
                "is_released": it.get("IsReleased"),
                "raw": it,
            }
            for it in items
        ],
    }


def biglock_guarded_object_status(native_id, obj_type="Auto"):
    """Статус объекта охраны (машины/борта) по NativeId напрямую:
    Free (свободен) / LockInProgress (постановка идёт) / Locked (на охране)."""
    opener = _biglock_opener()
    data = _biglock_post(opener, "/api/guardedobjects/search", {
        "Type": obj_type, "NativeId": native_id, "Limit": 10,
    })
    items = data.get("Items", [])
    return {
        "native_id": native_id,
        "total_count": data.get("TotalCount"),
        "objects": [
            {"status": it.get("Status"), "id": it.get("Id"), "raw": it}
            for it in items
        ],
    }


def db_get_active_trips_with_board():
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, board_number FROM trips WHERE status = 'в пути' AND board_number IS NOT NULL"
        )
        return [{"id": r[0], "board_number": r[1]} for r in cur.fetchall()]
    finally:
        conn.close()


def biglock_sync_trip_removal(trip_id, board_number):
    """Проверяет статус борта в BigLock; если объект свободен (Free) -
    закрывает ближайшую необработанную точку выгрузки рейса, используя
    реальное время смены статуса (StatusTime) от BigLock."""
    status_data = biglock_guarded_object_status(board_number)
    objects = status_data.get("objects", [])
    if not objects:
        return {"trip_id": trip_id, "board_number": board_number, "action": "not_found_in_biglock"}

    obj = objects[0]
    if obj.get("status") != "Free":
        return {"trip_id": trip_id, "board_number": board_number, "action": "still_locked"}

    status_time_raw = obj.get("raw", {}).get("StatusTime")
    try:
        completed_dt = datetime.datetime.fromisoformat(status_time_raw)
        if completed_dt.tzinfo is None:
            completed_dt = completed_dt.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=5)))
    except (ValueError, TypeError):
        completed_dt = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5)))

    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id FROM trip_stops
            WHERE trip_id = %s AND stop_type = 'выгрузка' AND status = 'ожидание'
            ORDER BY sequence LIMIT 1
            """,
            (trip_id,),
        )
        row = cur.fetchone()
    finally:
        conn.close()

    if not row:
        return {"trip_id": trip_id, "board_number": board_number, "action": "no_pending_stop"}

    stop_id = row[0]
    result = db_complete_stop(trip_id, stop_id, completed_dt, None)
    return {
        "trip_id": trip_id, "board_number": board_number, "action": "closed_stop",
        "stop_id": stop_id, "completed_at": completed_dt.isoformat(),
        "trip_closed": result["trip_closed"] if result else None,
    }


def biglock_sync_all_active_trips():
    trips = db_get_active_trips_with_board()
    results = []
    for t in trips:
        try:
            results.append(biglock_sync_trip_removal(t["id"], t["board_number"]))
        except Exception as e:
            results.append({"trip_id": t["id"], "board_number": t["board_number"], "action": "error", "error": str(e)})
    return results


def biglock_create_subscription(object_id, consumer_id="BigBlock", sub_type="LockGuardEvents"):
    """Подписывает нас на события конкретного охраняемого объекта -
    после этого BigLock должен сам присылать события навешивания/
    снятия (адрес назначения настраивается на стороне BigLock под
    этим ConsumerId, не через этот вызов)."""
    opener = _biglock_opener()
    return _biglock_post(opener, "/api/externalsubscriptions", {
        "Type": sub_type, "ConsumerId": consumer_id, "ObjectId": str(object_id),
    })


def db_store_biglock_webhook_event(raw_body):
    """Сохраняет любое входящее событие от BigLock как есть - формат
    ещё не подтверждён, поэтому просто складываем сырые данные,
    разбор сделаем по факту, глядя на реальные пришедшие данные."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        external_id = raw_body.get("Id") if isinstance(raw_body, dict) else None
        try:
            external_id = int(external_id) if external_id is not None else None
        except (ValueError, TypeError):
            external_id = None
        event_type = None
        if isinstance(raw_body, dict):
            event_type = raw_body.get("Type") or raw_body.get("EventType") or raw_body.get("TemplateCode")
        cur.execute(
            """
            INSERT INTO raw_events (source, event_type, external_id, event_dt, raw_payload, processed)
            VALUES ('biglock', %s, %s, now(), %s, false)
            """,
            (event_type, external_id, json.dumps(raw_body, ensure_ascii=False)),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ---------- Отчёт по ЭЗПУ для выставления счёта (расчёт суток) ----------

DEFAULT_EZPU_RATE = 4060


def _billing_days(hang_dt, removal_dt, cutoff_hour=11):
    """Сутки считаются 'от 11:00 до 11:00'. Если сняли до 11:00 - этот
    день не засчитывается (как ранний выезд из отеля)."""
    hang_date = hang_dt.date()
    removal_date = removal_dt.date()
    if removal_dt.time() < datetime.time(cutoff_hour, 0):
        removal_date = removal_date - datetime.timedelta(days=1)
    days = (removal_date - hang_date).days + 1
    return max(days, 1)


def db_get_ezpu_billing_report(contractor_name, year, month, rate=DEFAULT_EZPU_RATE):
    conn = get_connection()
    try:
        cur = conn.cursor()
        period_from = datetime.date(year, month, 1)
        period_to = datetime.date(year + 1, 1, 1) if month == 12 else datetime.date(year, month + 1, 1)

        cur.execute(
            """
            SELECT t.id, t.board_number, d.serial_number, t.hang_datetime, t.removal_datetime,
                   (SELECT location FROM trip_stops WHERE trip_id = t.id ORDER BY sequence LIMIT 1) AS origin,
                   (SELECT location FROM trip_stops WHERE trip_id = t.id ORDER BY sequence DESC LIMIT 1) AS destination
            FROM trips t
            JOIN devices d ON d.id = t.ezpu_device_id
            JOIN parties p ON p.id = t.contractor_id
            WHERE p.name = %s AND t.status = 'снят'
              AND t.hang_datetime >= %s AND t.hang_datetime < %s
              AND t.removal_datetime IS NOT NULL
            ORDER BY t.hang_datetime
            """,
            (contractor_name, period_from, period_to),
        )
        rows = []
        total_amount = 0
        for i, r in enumerate(cur.fetchall(), start=1):
            trip_id, board, serial, hang_dt, removal_dt, origin, destination = r
            days = _billing_days(hang_dt, removal_dt)
            amount = days * rate
            total_amount += amount
            rows.append({
                "num": i, "trip_id": trip_id, "board_number": board, "ezpu_serial": serial,
                "origin": origin, "destination": destination,
                "hang_datetime": hang_dt, "removal_datetime": removal_dt,
                "days": days, "rate": rate, "amount": amount,
            })
        return {
            "contractor": contractor_name, "year": year, "month": month, "rate": rate,
            "rows": rows, "total_amount": total_amount,
        }
    finally:
        conn.close()


_DOCX_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

_DOCX_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

_DOCX_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

_DOCX_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""


def _docx_p(text, bold=False, center=False, italic=False, size=22, justify=False, space_after=60):
    align = ""
    if center:
        align = '<w:jc w:val="center"/>'
    elif justify:
        align = '<w:jc w:val="both"/>'
    b = "<w:b/>" if bold else ""
    i = "<w:i/>" if italic else ""
    rpr = f'{b}{i}<w:sz w:val="{size}"/>'
    spacing = f'<w:spacing w:after="{space_after}" w:line="240" w:lineRule="auto"/>'
    return (
        f'<w:p><w:pPr>{align}{spacing}<w:rPr>{rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r></w:p>'
    )


def _docx_table(headers, rows, col_widths=None):
    def cell(text, bold=False, width=None):
        b = "<w:b/>" if bold else ""
        w = f'<w:tcW w:w="{width}" w:type="dxa"/>' if width else '<w:tcW w:w="0" w:type="auto"/>'
        return (
            f'<w:tc><w:tcPr>{w}<w:tcBorders>'
            f'<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
            f'<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
            f'</w:tcBorders></w:tcPr>'
            f'<w:p><w:r><w:rPr>{b}<w:sz w:val="18"/></w:rPr>'
            f'<w:t xml:space="preserve">{xml_escape(str(text))}</w:t></w:r></w:p></w:tc>'
        )

    def row(cells, bold=False):
        tds = "".join(cell(c, bold, col_widths[i] if col_widths else None) for i, c in enumerate(cells))
        return f"<w:tr>{tds}</w:tr>"

    tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
           '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
           '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
           '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>'
           '</w:tblBorders></w:tblPr>')
    tbl += row(headers, bold=True)
    for r in rows:
        tbl += row(r)
    tbl += "</w:tbl>"
    return tbl


def _fmt_ru_date(d):
    return f"«{d.day}» {_MONTHS_RU[d.month]} {d.year} г."


def _fmt_money(amount):
    return f"{amount:,.0f}".replace(",", " ")


def generate_act_docx(act):
    direction = act["direction"]
    counterparty = act["counterparty"] or "—"
    label = f" ({act['counterparty_label']})" if act["counterparty_label"] else ""

    if direction == "передача":
        verb_block = 'Исполнитель в соответствии с настоящим актом передал Заказчику, а Заказчик принял следующее оборудование:'
    else:
        verb_block = 'Заказчик в соответствии с настоящим актом возвратил Исполнителю, а Исполнитель принял следующее оборудование:'

    body_parts = []

    header_tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>'
        '<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>'
        '<w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="3500" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
        '<w:t xml:space="preserve">Приложение № 3</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
        f'<w:t xml:space="preserve">к Договору {xml_escape(act["contract_number"])}</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
        f'<w:t xml:space="preserve">от {act["contract_date"].strftime("%d.%m.%Y")} г.</w:t></w:r></w:p>'
        '</w:tc></w:tr></w:tbl>'
    )
    body_parts.append(header_tbl)

    body_parts.append(_docx_p("АКТ ПРИЕМА-ПЕРЕДАЧИ ОБОРУДОВАНИЯ", bold=True, center=True, size=26, space_after=20))
    body_parts.append(_docx_p(f"№{act['act_number']} {_fmt_ru_date(act['act_date'])}", bold=True, center=True, size=24, space_after=160))

    preamble = (
        f'ТОО «Транс Мониторинг Автоматизация», в лице Директора {act["director_name"]} '
        f'действующего на основании Устава, именуемое в дальнейшем Исполнитель, с одной стороны, '
        f'и {counterparty}{label}, именуемое в дальнейшем Заказчик, с другой стороны, '
        f'в дальнейшем совместно именуемые Стороны, а по отдельности — Сторона, подписали настоящий акт '
        f'приема-передачи Оборудования, согласно договора от {act["contract_date"].strftime("%d.%m.%Y")} '
        f'№ {act["contract_number"]}, заключенному между Сторонами, о нижеследующем:'
    )
    body_parts.append(_docx_p(preamble, justify=True, space_after=140))
    body_parts.append(_docx_p("1. " + verb_block, justify=True, space_after=100))

    headers = ["№ п/п", "Наименование оборудования", "Заводской (серийный) номер", "Ед. изм.", "Кол-во", "Стоимость, тенге."]
    rows = []
    total = 0
    for i, line in enumerate(act["lines"], start=1):
        qty = line["qty"] or 0
        price = float(line["unit_price"]) if line["unit_price"] is not None else 0
        total += qty * price
        rows.append([
            i, line["item_name"], line["serials"] or "—", line["unit"] or "шт", qty,
            _fmt_money(price) if line["unit_price"] is not None else "—",
        ])
    body_parts.append(_docx_table(headers, rows))

    if total > 0:
        body_parts.append(_docx_p(
            f"Общая стоимость Оборудования, {'передаваемого Исполнителем' if direction == 'передача' else 'возвращаемого Заказчиком'}, составляет:",
            space_after=20,
        ))
        body_parts.append(_docx_p(f"{_fmt_money(total)} тенге", bold=False, space_after=140))
    else:
        body_parts.append(_docx_p("", space_after=100))

    body_parts.append(_docx_p(
        "2. Претензий по внешнему виду, целостности и комплектации устройств, у Заказчика к Исполнителю "
        "по передаваемому Оборудованию не имеется.", justify=True, space_after=60
    ))
    body_parts.append(_docx_p(
        "3. Подписав настоящий акт, Стороны подтверждают, что обязательства Сторон по приему-передаче "
        "Оборудования по Договору исполнены надлежащим образом.", justify=True, space_after=60
    ))
    body_parts.append(_docx_p(
        "4. Настоящий акт подписан в 2 (двух) подлинных экземплярах на русском языке по одному "
        "для каждой из Сторон.", justify=True, space_after=240
    ))

    sig_tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>'
        '<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>'
        '<w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>'
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Исполнитель</w:t></w:r></w:p>'
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Директор</w:t></w:r></w:p>'
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ТОО «Транс Мониторинг Автоматизация»</w:t></w:r></w:p>'
        '<w:p/><w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">_____________________ </w:t></w:r>'
        f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{xml_escape(act["director_name"])}</w:t></w:r></w:p>'
        '</w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Заказчик</w:t></w:r></w:p>'
        f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{xml_escape(counterparty)}{xml_escape(label)}</w:t></w:r></w:p>'
        '<w:p/><w:p/>'
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>_____________________</w:t></w:r></w:p>'
        '</w:tc>'
        '</w:tr></w:tbl>'
    )
    body_parts.append(sig_tbl)

    sect_pr = (
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"/></w:sectPr>'
    )

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>" + "".join(body_parts) + sect_pr + "</w:body></w:document>"
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _DOCX_CONTENT_TYPES)
        z.writestr("_rels/.rels", _DOCX_RELS)
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", _DOCX_STYLES)
        z.writestr("word/_rels/document.xml.rels", _DOCX_DOC_RELS)
    return buf.getvalue()


class Handler(BaseHTTPRequestHandler):
    def _send_html(self, html, status=200):
        body = html.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, data, status=200):
        body = json.dumps(data, default=json_default, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        print("%s - %s" % (self.address_string(), format % args))

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if path in ("/dashboard", "/"):
            self._send_html(DASHBOARD_HTML)
            return

        if path == "/health":
            self._send_json({"status": "ok", "time": datetime.datetime.utcnow().isoformat()})
            return

        if path == "/devices":
            status = qs.get("status", [None])[0]
            device_type = qs.get("device_type", [None])[0]
            try:
                limit = int(qs.get("limit", ["200"])[0])
            except ValueError:
                limit = 200
            try:
                rows = db_list_devices(status=status, device_type=device_type, limit=limit)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json({"count": len(rows), "devices": rows})
            return

        if path == "/trips":
            status = qs.get("status", [None])[0]
            client = qs.get("client", [None])[0]
            try:
                limit = int(qs.get("limit", ["200"])[0])
            except ValueError:
                limit = 200
            try:
                rows = db_list_trips(status=status, client=client, limit=limit)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json({"count": len(rows), "trips": rows})
            return

        m = TRIP_RE.match(path)
        if m:
            trip_id = int(m.group(1))
            try:
                trip = db_get_trip(trip_id)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            if not trip:
                self._send_json({"error": f"Рейс {trip_id} не найден"}, status=404)
                return
            self._send_json(trip)
            return

        m = DEVICE_HISTORY_RE.match(path)
        if m:
            serial = m.group(1)
            try:
                rows = db_get_device_history(serial)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            if rows is None:
                self._send_json({"error": f"Устройство {serial} не найдено"}, status=404)
                return
            self._send_json(rows)
            return

        m = DEVICE_RE.match(path)
        if m:
            serial = m.group(1)
            try:
                row = db_get_device(serial)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            if not row:
                self._send_json({"error": f"Устройство {serial} не найдено"}, status=404)
                return
            self._send_json(row)
            return

        if path == "/acts":
            try:
                acts = db_list_acts()
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json({"count": len(acts), "acts": acts})
            return

        if path == "/raw-events":
            source = qs.get("source", [None])[0]
            try:
                limit = int(qs.get("limit", ["20"])[0])
            except ValueError:
                limit = 20
            try:
                events = db_list_raw_events(source=source, limit=limit)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json({"count": len(events), "events": events})
            return

        if path == "/reports/ezpu-billing":
            contractor = qs.get("contractor", [None])[0]
            year = qs.get("year", [None])[0]
            month = qs.get("month", [None])[0]
            rate = qs.get("rate", [None])[0]
            if not contractor or not year or not month:
                self._send_json({"error": "Укажите contractor, year, month"}, status=400)
                return
            try:
                year = int(year)
                month = int(month)
                rate = float(rate) if rate else DEFAULT_EZPU_RATE
            except ValueError:
                self._send_json({"error": "year/month/rate должны быть числами"}, status=400)
                return
            try:
                report = db_get_ezpu_billing_report(contractor, year, month, rate)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json(report)
            return

        if path == "/biglock/events":
            native_id = qs.get("native_id", [None])[0]
            try:
                limit = int(qs.get("limit", ["200"])[0])
            except ValueError:
                limit = 200
            codes_raw = qs.get("codes", [None])[0]
            if codes_raw and codes_raw.upper() == "ALL":
                codes = []
            elif codes_raw:
                codes = codes_raw.split(",")
            else:
                codes = None
            try:
                result = biglock_events_for_object(native_id, limit=limit, codes=codes)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json(result)
            return

        if path == "/biglock/sessions":
            ezpu_serial = qs.get("ezpu_serial", [None])[0]
            if not ezpu_serial:
                self._send_json({"error": "Укажите ezpu_serial"}, status=400)
                return
            try:
                limit = int(qs.get("limit", ["5000"])[0])
            except ValueError:
                limit = 5000
            codes_raw = qs.get("codes", [None])[0]
            if codes_raw and codes_raw.upper() == "ALL":
                codes = []
            elif codes_raw:
                codes = codes_raw.split(",")
            else:
                codes = None
            try:
                result = biglock_device_sessions(ezpu_serial, limit=limit, codes=codes)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json(result)
            return

        if path == "/wialon/webhook":
            board = qs.get("board", [None])[0]
            zone = qs.get("zone", [None])[0]
            time_raw = qs.get("time", [None])[0]

            if time_raw:
                try:
                    arrived_dt = datetime.datetime.fromtimestamp(
                        int(time_raw), tz=datetime.timezone(datetime.timedelta(hours=5))
                    )
                except (ValueError, TypeError):
                    arrived_dt = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5)))
            else:
                arrived_dt = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5)))

            try:
                db_store_wialon_webhook_event({"board": board, "zone": zone, "time": time_raw, "query": self.path})
            except Exception as e:
                print("Ошибка сохранения Wialon webhook:", e)

            if not board:
                self._send_json({"status": "logged_no_board"})
                return

            try:
                result = db_auto_mark_arrival_by_board(board, arrived_dt, location_hint=zone)
            except Exception as e:
                self._send_json({"status": "logged", "match_error": str(e)})
                return

            self._send_json({"status": "ok", "match": result})
            return

        if path == "/biglock/device-status":
            case_id = qs.get("case_id", [None])[0]
            if not case_id:
                self._send_json({"error": "Укажите case_id"}, status=400)
                return
            try:
                result = biglock_device_status(case_id)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json(result)
            return

        if path == "/biglock/lock-status":
            zpu_number = qs.get("zpu_number", [None])[0]
            native_id = qs.get("native_id", [None])[0]
            client_id_raw = qs.get("client_id", [None])[0]
            is_released_raw = qs.get("is_released", [None])[0]
            if not zpu_number and not native_id and not client_id_raw:
                self._send_json({"error": "Укажите zpu_number, native_id или client_id"}, status=400)
                return
            try:
                limit = int(qs.get("limit", ["10"])[0])
            except ValueError:
                limit = 10
            client_id = int(client_id_raw) if client_id_raw else None
            is_released = None
            if is_released_raw is not None:
                is_released = is_released_raw.lower() == "true"
            try:
                result = biglock_lock_status(
                    mechanical_case_id=zpu_number, native_id=native_id,
                    client_id=client_id, is_released=is_released, limit=limit,
                )
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json(result)
            return

        if path == "/biglock/guarded-object":
            native_id = qs.get("native_id", [None])[0]
            obj_type = qs.get("type", ["Auto"])[0]
            if not native_id:
                self._send_json({"error": "Укажите native_id"}, status=400)
                return
            try:
                result = biglock_guarded_object_status(native_id, obj_type)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json(result)
            return

        m = ACT_DOWNLOAD_RE.match(path)
        if m:
            act_id = int(m.group(1))
            try:
                act = db_get_act(act_id)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            if not act:
                self._send_json({"error": f"Акт {act_id} не найден"}, status=404)
                return
            docx_bytes = generate_act_docx(act)
            filename = f"Akt_{act['act_number']}.docx"
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(docx_bytes)))
            self.end_headers()
            self.wfile.write(docx_bytes)
            return

        m = ACT_RE.match(path)
        if m:
            act_id = int(m.group(1))
            try:
                act = db_get_act(act_id)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            if not act:
                self._send_json({"error": f"Акт {act_id} не найден"}, status=404)
                return
            self._send_json(act)
            return

        self._send_json({"error": "не найдено"}, status=404)

    def do_POST(self):
        path = urlparse(self.path).path

        if path == "/operations":
            self._handle_create_operation()
            return

        if path == "/trips":
            self._handle_create_trip()
            return

        if path == "/acts":
            self._handle_create_act()
            return

        if path == "/biglock/webhook":
            self._handle_biglock_webhook()
            return

        if path == "/biglock/subscribe":
            self._handle_biglock_subscribe()
            return

        if path == "/biglock/sync-now":
            try:
                results = biglock_sync_all_active_trips()
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)
                return
            self._send_json({"count": len(results), "results": results})
            return

        m = TRIP_CLOSE_RE.match(path)
        if m:
            self._handle_close_trip(int(m.group(1)))
            return

        m = TRIP_ASSIGN_RE.match(path)
        if m:
            self._handle_assign_device(int(m.group(1)))
            return

        m = TRIP_STOP_COMPLETE_RE.match(path)
        if m:
            self._handle_complete_stop(int(m.group(1)), int(m.group(2)))
            return

        m = TRIP_STOP_ARRIVE_RE.match(path)
        if m:
            self._handle_mark_arrived(int(m.group(1)), int(m.group(2)))
            return

        m = TRIP_STOP_ZPU_RE.match(path)
        if m:
            self._handle_set_stop_zpu(int(m.group(1)), int(m.group(2)))
            return

        self._send_json({"error": "не найдено"}, status=404)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b"{}"
        return json.loads(raw)

    def _handle_create_operation(self):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        device_serial = (body.get("device_serial") or "").strip()
        operation_type = (body.get("operation_type") or "").strip()

        if not device_serial:
            self._send_json({"error": "Поле device_serial обязательно"}, status=400)
            return
        if operation_type not in ALLOWED_OPERATION_TYPES:
            self._send_json(
                {"error": f"operation_type должен быть одним из: {sorted(ALLOWED_OPERATION_TYPES)}"},
                status=400,
            )
            return

        operation_dt_raw = body.get("operation_dt")
        if operation_dt_raw:
            try:
                operation_dt = datetime.datetime.fromisoformat(operation_dt_raw)
            except ValueError:
                self._send_json({"error": "operation_dt должен быть в формате ISO 8601"}, status=400)
                return
        else:
            operation_dt = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5)))

        payload = {
            "device_serial": device_serial,
            "device_type_hint": body.get("device_type", "ezpu"),
            "operation_type": operation_type,
            "from_party": (body.get("from_party") or "").strip() or None,
            "to_party": (body.get("to_party") or "").strip() or None,
            "location": body.get("location"),
            "document_ref": body.get("document_ref"),
            "recorded_by": body.get("recorded_by"),
            "operation_dt": operation_dt,
        }

        try:
            new_id = db_create_operation(payload)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        self._send_json({"id": new_id, "status": "created"}, status=201)

    def _parse_dt(self, raw, default_now=True):
        if raw:
            return datetime.datetime.fromisoformat(raw)
        if default_now:
            return datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5)))
        return None

    MEGAPOLIS_NAME = "ТОО ТК Мегаполис Казахстан"

    def _handle_create_trip(self):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        ezpu_serial = (body.get("ezpu_serial") or "").strip() or None
        tracker_serial = (body.get("tracker_serial") or "").strip() or None
        lock_serial = (body.get("lock_serial") or "").strip() or None
        contractor = (body.get("contractor") or "").strip() or None

        pickups = body.get("pickups")
        dropoffs = body.get("dropoffs")
        if pickups is None and body.get("warehouse"):
            pickups = [body.get("warehouse")]
        if dropoffs is None:
            dropoffs = list(body.get("waypoints") or [])
            if body.get("destination_city"):
                dropoffs.append(body.get("destination_city"))

        if not isinstance(pickups, list) or not any((p or "").strip() for p in pickups):
            self._send_json({"error": "Укажите хотя бы один пункт погрузки (склад отгрузки)"}, status=400)
            return
        if not isinstance(dropoffs, list) or not any((d or "").strip() for d in dropoffs):
            self._send_json({"error": "Укажите хотя бы один пункт выгрузки"}, status=400)
            return

        if contractor == self.MEGAPOLIS_NAME and (ezpu_serial or tracker_serial) and not tracker_serial:
            self._send_json(
                {"error": f"Для подрядчика «{self.MEGAPOLIS_NAME}» обязателен номер трекера (tracker_serial)"},
                status=400,
            )
            return

        try:
            hang_dt = self._parse_dt(body.get("hang_datetime"))
        except ValueError:
            self._send_json({"error": "hang_datetime должен быть в формате ISO 8601"}, status=400)
            return

        payload = {
            "client": (body.get("client") or "").strip() or None,
            "contractor": contractor,
            "board_number": body.get("board_number"),
            "pickups": pickups,
            "dropoffs": dropoffs,
            "ezpu_serial": ezpu_serial,
            "tracker_serial": tracker_serial,
            "lock_serial": lock_serial,
            "zpu_number": (body.get("zpu_number") or "").strip() or None,
            "origin_city": body.get("origin_city"),
            "hang_datetime": hang_dt,
            "notes": body.get("notes"),
        }

        try:
            trip_id = db_create_trip(payload)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        self._send_json({"id": trip_id, "status": "created"}, status=201)

    def _handle_close_trip(self, trip_id):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        try:
            removal_dt = self._parse_dt(body.get("removal_datetime"))
        except ValueError:
            self._send_json({"error": "removal_datetime должен быть в формате ISO 8601"}, status=400)
            return

        location = body.get("location")

        try:
            result = db_close_trip(trip_id, removal_dt, location)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        if result is None:
            self._send_json({"error": f"Рейс {trip_id} не найден"}, status=404)
            return

        self._send_json({"id": trip_id, "status": "closed"})

    def _handle_assign_device(self, trip_id):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        ezpu_serial = (body.get("ezpu_serial") or "").strip() or None
        tracker_serial = (body.get("tracker_serial") or "").strip() or None
        lock_serial = (body.get("lock_serial") or "").strip() or None
        contractor = (body.get("contractor") or "").strip() or None

        if not ezpu_serial and not tracker_serial:
            self._send_json({"error": "Укажите ezpu_serial или tracker_serial"}, status=400)
            return

        if contractor == self.MEGAPOLIS_NAME and not tracker_serial:
            self._send_json(
                {"error": f"Для подрядчика «{self.MEGAPOLIS_NAME}» обязателен номер трекера (tracker_serial)"},
                status=400,
            )
            return

        try:
            assign_dt = self._parse_dt(body.get("assign_datetime"))
        except ValueError:
            self._send_json({"error": "assign_datetime должен быть в формате ISO 8601"}, status=400)
            return

        try:
            result = db_assign_trip_device(trip_id, ezpu_serial, tracker_serial, lock_serial,
                                            (body.get("zpu_number") or "").strip() or None, assign_dt)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        if result is None:
            self._send_json({"error": f"Рейс {trip_id} не найден"}, status=404)
            return

        self._send_json({"id": trip_id, "status": "device_assigned"})

    def _handle_mark_arrived(self, trip_id, stop_id):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        try:
            arrived_dt = self._parse_dt(body.get("arrived_at"))
        except ValueError:
            self._send_json({"error": "arrived_at должен быть в формате ISO 8601"}, status=400)
            return

        try:
            ok = db_mark_stop_arrived(trip_id, stop_id, arrived_dt)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        if not ok:
            self._send_json({"error": "Остановка не найдена"}, status=404)
            return

        self._send_json({"stop_id": stop_id, "status": "arrived"})

    def _handle_complete_stop(self, trip_id, stop_id):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        try:
            completed_dt = self._parse_dt(body.get("completed_at"))
        except ValueError:
            self._send_json({"error": "completed_at должен быть в формате ISO 8601"}, status=400)
            return

        location_override = (body.get("location") or "").strip() or None

        try:
            result = db_complete_stop(trip_id, stop_id, completed_dt, location_override)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        if result is None:
            self._send_json({"error": "Остановка не найдена"}, status=404)
            return

        self._send_json({"stop_id": stop_id, "status": "completed", "trip_closed": result["trip_closed"]})

    def _handle_set_stop_zpu(self, trip_id, stop_id):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        zpu_number = (body.get("zpu_number") or "").strip() or None

        try:
            ok = db_set_stop_zpu(trip_id, stop_id, zpu_number)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        if not ok:
            self._send_json({"error": "Остановка не найдена"}, status=404)
            return

        self._send_json({"stop_id": stop_id, "zpu_number": zpu_number, "status": "updated"})

    def _handle_create_act(self):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        direction = (body.get("direction") or "").strip()
        counterparty = (body.get("counterparty") or "").strip()
        lines_in = body.get("lines")

        if direction not in ("передача", "возврат"):
            self._send_json({"error": "direction должен быть 'передача' или 'возврат'"}, status=400)
            return
        if not counterparty:
            self._send_json({"error": "Укажите counterparty (заказчика)"}, status=400)
            return
        if not isinstance(lines_in, list) or not lines_in:
            self._send_json({"error": "Укажите хотя бы одну позицию в lines"}, status=400)
            return

        try:
            act_date = datetime.datetime.strptime(body.get("act_date"), "%Y-%m-%d").date()
        except (ValueError, TypeError):
            self._send_json({"error": "act_date должен быть в формате ГГГГ-ММ-ДД"}, status=400)
            return

        contract_date_raw = body.get("contract_date") or "2020-07-01"
        try:
            contract_date = datetime.datetime.strptime(contract_date_raw, "%Y-%m-%d").date()
        except ValueError:
            self._send_json({"error": "contract_date должен быть в формате ГГГГ-ММ-ДД"}, status=400)
            return

        lines = []
        for i, ln in enumerate(lines_in):
            item_name = (ln.get("item_name") or "").strip()
            if not item_name:
                self._send_json({"error": f"Позиция {i + 1}: укажите item_name"}, status=400)
                return
            serials = ln.get("serials") or []
            if not isinstance(serials, list):
                self._send_json({"error": f"Позиция {i + 1}: serials должен быть списком"}, status=400)
                return
            serials = [s.strip() for s in serials if (s or "").strip()]
            qty = ln.get("qty")
            if qty is None:
                qty = len(serials) if serials else 0
            try:
                qty = int(qty)
            except (ValueError, TypeError):
                self._send_json({"error": f"Позиция {i + 1}: qty должен быть числом"}, status=400)
                return
            unit_price = ln.get("unit_price")
            if unit_price is not None:
                try:
                    unit_price = float(unit_price)
                except (ValueError, TypeError):
                    self._send_json({"error": f"Позиция {i + 1}: unit_price должен быть числом"}, status=400)
                    return
            device_type = ln.get("device_type")
            if device_type not in ("ezpu", "tracker", "lock", None):
                self._send_json({"error": f"Позиция {i + 1}: device_type должен быть ezpu/tracker/lock"}, status=400)
                return

            lines.append({
                "item_name": item_name, "serials": serials, "unit": ln.get("unit") or "шт",
                "qty": qty, "unit_price": unit_price, "device_type": device_type,
            })

        payload = {
            "direction": direction,
            "counterparty": counterparty,
            "counterparty_label": (body.get("counterparty_label") or "").strip() or None,
            "contract_number": (body.get("contract_number") or "").strip() or "07/04-2020",
            "contract_date": contract_date,
            "act_date": act_date,
            "director_name": (body.get("director_name") or "").strip() or "Архаров Н.Э.",
            "lines": lines,
        }

        try:
            act_id, act_number = db_create_act(payload)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        self._send_json({"id": act_id, "act_number": act_number, "status": "created"}, status=201)

    def _handle_biglock_webhook(self):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            # даже если тело не JSON - отвечаем 200, чтобы BigLock не считал вебхук сломанным
            self._send_json({"status": "ignored_not_json"}, status=200)
            return

        try:
            db_store_biglock_webhook_event(body)
        except Exception as e:
            # логируем, но всё равно отвечаем 200 - иначе BigLock может отключить подписку
            print("Ошибка сохранения BigLock webhook:", e)
            self._send_json({"status": "error_logged"}, status=200)
            return

        self._send_json({"status": "ok"}, status=200)

    def _handle_biglock_subscribe(self):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        object_id = body.get("object_id")
        if not object_id:
            self._send_json({"error": "Укажите object_id"}, status=400)
            return

        consumer_id = body.get("consumer_id") or "BigBlock"

        try:
            result = biglock_create_subscription(object_id, consumer_id=consumer_id)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)
            return

        self._send_json(result)


def _biglock_background_sync_loop(interval_seconds=300):
    while True:
        time.sleep(interval_seconds)
        try:
            if os.environ.get("BIGLOCK_LOGIN") and os.environ.get("BIGLOCK_PASSWORD"):
                results = biglock_sync_all_active_trips()
                closed = [r for r in results if r.get("action") == "closed_stop"]
                if closed:
                    print(f"[biglock-sync] обработано рейсов: {len(results)}, закрыто точек: {len(closed)}")
        except Exception as e:
            print("[biglock-sync] ошибка фоновой синхронизации:", e)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    sync_thread = threading.Thread(target=_biglock_background_sync_loop, daemon=True)
    sync_thread.start()
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Сервер запущен на порту {port}")
    server.serve_forever()
