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
from urllib.parse import urlparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pg8000

DATABASE_URL = os.environ.get("DATABASE_URL")
ALLOWED_OPERATION_TYPES = {"навешивание", "снятие", "передача", "возврат", "ремонт", "списание"}

DEVICE_RE = re.compile(r"^/devices/([^/]+)$")
DEVICE_HISTORY_RE = re.compile(r"^/devices/([^/]+)/history$")


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


# ---------- HTTP-сервер ----------

class Handler(BaseHTTPRequestHandler):
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
        if self.path == "/health":
            self._send_json({"status": "ok", "time": datetime.datetime.utcnow().isoformat()})
            return

        m = DEVICE_HISTORY_RE.match(self.path)
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

        m = DEVICE_RE.match(self.path)
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
        if self.path != "/operations":
            self._send_json({"error": "не найдено"}, status=404)
            return

        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b"{}"
        try:
            body = json.loads(raw)
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Сервер запущен на порту {port}")
    server.serve_forever()
