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
from urllib.parse import urlparse, parse_qs
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pg8000

DATABASE_URL = os.environ.get("DATABASE_URL")
ALLOWED_OPERATION_TYPES = {"навешивание", "снятие", "передача", "возврат", "ремонт", "списание"}

DEVICE_RE = re.compile(r"^/devices/([^/]+)$")
DEVICE_HISTORY_RE = re.compile(r"^/devices/([^/]+)/history$")
TRIP_RE = re.compile(r"^/trips/(\d+)$")
TRIP_CLOSE_RE = re.compile(r"^/trips/(\d+)/close$")


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


DASHBOARD_HTML = """<!DOCTYPE html>
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
  .filters { display: flex; gap: 8px; flex-wrap: wrap; align-items: end; }
  .field { display: flex; flex-direction: column; gap: 4px; }
  .field label { font-size: 12px; color: #666; }
  input, select, button { padding: 8px 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px; }
  button { background: #1f2937; color: #fff; border: none; cursor: pointer; }
  button:hover { background: #374151; }
  button.secondary { background: #e5e7eb; color: #222; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { text-align: left; padding: 8px 10px; border-bottom: 1px solid #eee; }
  th { background: #f9fafb; font-weight: 600; color: #444; }
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
  .trip-status { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; background: #dbeafe; color: #1e40af; }
  .trip-status-снят { background: #dcfce7; color: #166534; }
</style>
</head>
<body>
<header><h1>ТМА — учёт ЭЗПУ и трекеров</h1></header>
<div class="wrap">

  <div class="tabs">
    <button class="tab-btn active" id="tab-btn-devices" onclick="switchTab('devices')">Устройства</button>
    <button class="tab-btn" id="tab-btn-trips" onclick="switchTab('trips')">Рейсы</button>
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
    <table>
      <thead><tr>
        <th>Серийный номер</th><th>Тип</th><th>Статус</th><th>Местонахождение</th><th>Держатель</th><th>Последняя операция</th>
      </tr></thead>
      <tbody id="devices-body"></tbody>
    </table>
  </div>

  </div>

  <div class="tab-view" id="tab-trips">

  <div class="panel">
    <div class="filters">
      <div class="field">
        <label>Статус рейса</label>
        <select id="tf-status">
          <option value="">все</option>
          <option value="в пути">в пути</option>
          <option value="снят">снят</option>
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
      <div class="field"><label>№ борта</label><input id="tr-board"></div>
      <div class="field"><label>Серийный номер ЭЗПУ</label><input id="tr-ezpu"></div>
      <div class="field"><label>Серийный номер трекера</label><input id="tr-tracker"></div>
      <div class="field"><label>Город отправления</label><input id="tr-origin"></div>
      <div class="field"><label>Город назначения</label><input id="tr-dest"></div>
      <div class="field"><label>Склад</label><input id="tr-warehouse"></div>
      <div class="full"><label>Примечания</label><input id="tr-notes" style="width:100%"></div>
      <div class="full"><button onclick="submitTrip()">Создать рейс</button></div>
      <div class="full" id="trip-msg"></div>
    </div>
  </div>

  <div class="panel">
    <div class="count" id="trip-count-label">Загрузка...</div>
    <table>
      <thead><tr>
        <th>№</th><th>Клиент</th><th>ЭЗПУ</th><th>Трекер</th><th>Маршрут</th><th>Навешена</th><th>Статус</th><th></th>
      </tr></thead>
      <tbody id="trips-body"></tbody>
    </table>
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
  document.getElementById('tab-btn-devices').classList.toggle('active', name === 'devices');
  document.getElementById('tab-btn-trips').classList.toggle('active', name === 'trips');
  if (name === 'trips') loadTrips();
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
  const body = document.getElementById('trips-body');
  body.innerHTML = '';
  data.trips.forEach(t => {
    const tr = document.createElement('tr');
    const route = (t.origin_city || '—') + ' → ' + (t.destination_city || '—');
    const closeBtn = t.status === 'в пути'
      ? `<button class="secondary" onclick="closeTrip(${t.id})">Закрыть</button>`
      : '';
    tr.innerHTML = `
      <td>${t.id}</td>
      <td>${t.client || '—'}</td>
      <td>${t.ezpu_serial || '—'}</td>
      <td>${t.tracker_serial || '—'}</td>
      <td>${route}</td>
      <td>${fmtDate(t.hang_datetime)}</td>
      <td><span class="trip-status trip-status-${t.status}">${t.status}</span></td>
      <td>${closeBtn}</td>
    `;
    body.appendChild(tr);
  });
  document.getElementById('trip-count-label').textContent = data.count + ' рейсов';
}

function openTripForm() {
  document.getElementById('trip-form-panel').style.display = 'block';
  document.getElementById('trip-form-panel').scrollIntoView({behavior: 'smooth'});
}
function closeTripForm() { document.getElementById('trip-form-panel').style.display = 'none'; }

async function submitTrip() {
  const msg = document.getElementById('trip-msg');
  msg.textContent = '';
  msg.className = '';
  const payload = {
    client: document.getElementById('tr-client').value.trim(),
    board_number: document.getElementById('tr-board').value.trim(),
    ezpu_serial: document.getElementById('tr-ezpu').value.trim(),
    tracker_serial: document.getElementById('tr-tracker').value.trim(),
    origin_city: document.getElementById('tr-origin').value.trim(),
    destination_city: document.getElementById('tr-dest').value.trim(),
    warehouse: document.getElementById('tr-warehouse').value.trim(),
    notes: document.getElementById('tr-notes').value.trim(),
  };
  if (!payload.ezpu_serial && !payload.tracker_serial) {
    msg.textContent = 'Укажите серийный номер ЭЗПУ или трекера';
    msg.className = 'err';
    return;
  }
  const r = await fetch('/trips', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload),
  });
  const data = await r.json();
  if (r.status === 201) {
    msg.textContent = 'Рейс создан, id ' + data.id;
    msg.className = 'ok';
    ['tr-client','tr-board','tr-ezpu','tr-tracker','tr-origin','tr-dest','tr-warehouse','tr-notes'].forEach(id => {
      document.getElementById(id).value = '';
    });
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
        ezpu_id = _get_or_create_device(cur, payload["ezpu_serial"], "ezpu") if payload["ezpu_serial"] else None
        tracker_id = _get_or_create_device(cur, payload["tracker_serial"], "tracker") if payload["tracker_serial"] else None

        cur.execute(
            """
            INSERT INTO trips
                (client_id, board_number, warehouse, ezpu_device_id, tracker_device_id,
                 origin_city, destination_city, hang_datetime, status, notes)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'в пути', %s)
            RETURNING id
            """,
            (client_id, payload["board_number"], payload["warehouse"], ezpu_id, tracker_id,
             payload["origin_city"], payload["destination_city"], payload["hang_datetime"], payload["notes"]),
        )
        trip_id = cur.fetchone()[0]

        if ezpu_id:
            cur.execute(
                """
                INSERT INTO operations
                    (device_id, trip_id, operation_type, location, operation_dt, document_ref)
                VALUES (%s, %s, 'навешивание', %s, %s, %s)
                """,
                (ezpu_id, trip_id, payload["origin_city"], payload["hang_datetime"], payload["board_number"]),
            )
            cur.execute(
                """
                UPDATE devices SET current_location = %s, last_operation_at = %s, updated_at = now()
                WHERE id = %s
                """,
                (payload["origin_city"], payload["hang_datetime"], ezpu_id),
            )

        conn.commit()
        return trip_id
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
            SELECT t.id, c.name, t.board_number, t.warehouse,
                   de.serial_number, dt.serial_number,
                   t.origin_city, t.destination_city, t.hang_datetime,
                   t.arrival_at_unload_datetime, t.removal_datetime, t.status, t.notes
            FROM trips t
            LEFT JOIN clients c ON c.id = t.client_id
            LEFT JOIN devices de ON de.id = t.ezpu_device_id
            LEFT JOIN devices dt ON dt.id = t.tracker_device_id
            WHERE t.id = %s
            """,
            (trip_id,),
        )
        r = cur.fetchone()
        if not r:
            return None
        return {
            "id": r[0], "client": r[1], "board_number": r[2], "warehouse": r[3],
            "ezpu_serial": r[4], "tracker_serial": r[5],
            "origin_city": r[6], "destination_city": r[7], "hang_datetime": r[8],
            "arrival_at_unload_datetime": r[9], "removal_datetime": r[10],
            "status": r[11], "notes": r[12],
        }
    finally:
        conn.close()


def db_list_trips(status=None, client=None, limit=200):
    conn = get_connection()
    try:
        cur = conn.cursor()
        query = """
            SELECT t.id, c.name, t.board_number, de.serial_number, dt.serial_number,
                   t.origin_city, t.destination_city, t.hang_datetime, t.status
            FROM trips t
            LEFT JOIN clients c ON c.id = t.client_id
            LEFT JOIN devices de ON de.id = t.ezpu_device_id
            LEFT JOIN devices dt ON dt.id = t.tracker_device_id
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
        return [
            {
                "id": r[0], "client": r[1], "board_number": r[2],
                "ezpu_serial": r[3], "tracker_serial": r[4],
                "origin_city": r[5], "destination_city": r[6],
                "hang_datetime": r[7], "status": r[8],
            }
            for r in cur.fetchall()
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

        self._send_json({"error": "не найдено"}, status=404)

    def do_POST(self):
        path = urlparse(self.path).path

        if path == "/operations":
            self._handle_create_operation()
            return

        if path == "/trips":
            self._handle_create_trip()
            return

        m = TRIP_CLOSE_RE.match(path)
        if m:
            self._handle_close_trip(int(m.group(1)))
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

    def _handle_create_trip(self):
        try:
            body = self._read_json_body()
        except json.JSONDecodeError:
            self._send_json({"error": "Тело запроса должно быть JSON"}, status=400)
            return

        ezpu_serial = (body.get("ezpu_serial") or "").strip() or None
        tracker_serial = (body.get("tracker_serial") or "").strip() or None

        if not ezpu_serial and not tracker_serial:
            self._send_json({"error": "Укажите хотя бы ezpu_serial или tracker_serial"}, status=400)
            return

        try:
            hang_dt = self._parse_dt(body.get("hang_datetime"))
        except ValueError:
            self._send_json({"error": "hang_datetime должен быть в формате ISO 8601"}, status=400)
            return

        payload = {
            "client": (body.get("client") or "").strip() or None,
            "board_number": body.get("board_number"),
            "warehouse": body.get("warehouse"),
            "ezpu_serial": ezpu_serial,
            "tracker_serial": tracker_serial,
            "origin_city": body.get("origin_city"),
            "destination_city": body.get("destination_city"),
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Сервер запущен на порту {port}")
    server.serve_forever()
