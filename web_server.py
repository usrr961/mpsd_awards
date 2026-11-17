"""LAN-only read-and-override web portal for the CSA desktop application."""

import base64
import hashlib
import hmac
import html
import json
import math
import secrets
import socket
import threading
import time
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

from award_engine import AWARD_ORDER, calculate_awards
from database import AwardOverride, SessionLocal, Student


def local_ip_addresses():
    """Return a useful LAN address for display (does not send network traffic)."""
    addresses = set()
    for family in (socket.AF_INET,):
        try:
            sock = socket.socket(family, socket.SOCK_DGRAM)
            sock.connect(("10.255.255.255", 80))
            addresses.add(sock.getsockname()[0])
            sock.close()
        except OSError:
            pass
    return sorted(addresses)


class LANPortal:
    def __init__(self, host="0.0.0.0", port=8088, pin=None):
        self.host = host
        self.port = int(port)
        self.pin = str(pin) if pin is not None else secrets.token_urlsafe(4)[:6]
        self.secret = secrets.token_bytes(32)
        self.server = None
        self.thread = None
        self.activity = []
        self.lock = threading.RLock()

    def start(self):
        if self.server is not None:
            return self
        owner = self

        class Handler(BaseHTTPRequestHandler):
            server_version = "CSAPortal/1.0"

            def _lan_request(self):
                peer = self.client_address[0]
                if peer.startswith("::ffff:"):
                    peer = peer.split("::ffff:", 1)[1]
                if peer == "127.0.0.1" or peer == "::1":
                    return True
                if peer.startswith("10.") or peer.startswith("192.168."):
                    return True
                if peer.startswith("172."):
                    parts = peer.split(".")
                    if len(parts) >= 2:
                        try:
                            return 16 <= int(parts[1]) <= 31
                        except ValueError:
                            return False
                return False

            def _json(self, status, payload):
                data = json.dumps(payload).encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(data)

            def _actor(self):
                cookie = next((part.strip()[5:] for part in self.headers.get("Cookie", "").split(";")
                               if part.strip().startswith("csa=")), None)
                authorization = self.headers.get("Authorization", "")
                token = cookie
                if authorization.startswith("Bearer "):
                    token = authorization[7:].strip()
                if not token:
                    return None
                try:
                    encoded, timestamp, signature = token.split(".")
                    if time.time() - float(timestamp) > 12 * 60 * 60:
                        return None
                    expected = hmac.new(owner.secret, f"{encoded}.{timestamp}".encode(), hashlib.sha256).hexdigest()
                    if not hmac.compare_digest(signature, expected):
                        return None
                    name = base64.urlsafe_b64decode(encoded + "==").decode("utf-8")
                    return name[:80] if name else None
                except Exception:
                    return None

            def do_GET(self):
                if not self._lan_request():
                    self._json(HTTPStatus.FORBIDDEN, {"error": "LAN access only"})
                    return
                path = urlparse(self.path).path
                if path == "/" or path == "/portal":
                    self._html()
                elif path == "/api/session":
                    actor = self._actor()
                    self._json(200, {"authenticated": actor is not None, "actor": actor,
                                     "award_options": list(AWARD_ORDER)})
                elif path == "/api/students":
                    actor = self._actor()
                    if not actor:
                        self._json(HTTPStatus.UNAUTHORIZED, {"error": "Sign in required"})
                        return
                    self._json(200, owner.student_payload())
                elif path == "/api/activity":
                    if not self._actor():
                        self._json(HTTPStatus.UNAUTHORIZED, {"error": "Sign in required"})
                        return
                    self._json(200, {"activity": list(owner.activity)})
                else:
                    self._json(HTTPStatus.NOT_FOUND, {"error": "Not found"})

            def do_POST(self):
                if not self._lan_request():
                    self._json(HTTPStatus.FORBIDDEN, {"error": "LAN access only"})
                    return
                path = urlparse(self.path).path
                try:
                    length = int(self.headers.get("Content-Length", "0"))
                    body = json.loads(self.rfile.read(length) or "{}")
                except (ValueError, json.JSONDecodeError):
                    self._json(HTTPStatus.BAD_REQUEST, {"error": "Invalid JSON"})
                    return

                if path == "/api/login":
                    pin = str(body.get("pin", ""))
                    name = str(body.get("name", "")).strip()[:80]
                    if not hmac.compare_digest(pin.encode("utf-8"), owner.pin.encode("utf-8")):
                        self._json(HTTPStatus.UNAUTHORIZED, {"error": "Incorrect access PIN"})
                        return
                    if not name:
                        self._json(HTTPStatus.BAD_REQUEST, {"error": "Display name required"})
                        return
                    encoded = base64.urlsafe_b64encode(name.encode("utf-8")).decode("ascii").rstrip("=")
                    timestamp = str(int(time.time()))
                    signature = hmac.new(owner.secret, f"{encoded}.{timestamp}".encode(), hashlib.sha256).hexdigest()
                    token = f"{encoded}.{timestamp}.{signature}"
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.send_header("Set-Cookie", f"csa={token}; HttpOnly; SameSite=Strict; Path=/; Max-Age=43200")
                    self.end_headers()
                    self.wfile.write(json.dumps({"ok": True, "actor": name, "token": token}).encode())
                    return

                if path == "/api/override":
                    actor = self._actor()
                    if not actor:
                        self._json(HTTPStatus.UNAUTHORIZED, {"error": "Sign in required"})
                        return
                    student_id = str(body.get("student_id", "")).strip()
                    standard = str(body.get("standard", "")).strip()
                    owl = bool(body.get("owl", False))
                    if standard not in ("", *AWARD_ORDER):
                        self._json(HTTPStatus.BAD_REQUEST, {"error": "Invalid standard award"})
                        return
                    if not student_id:
                        self._json(HTTPStatus.BAD_REQUEST, {"error": "Student required"})
                        return
                    try:
                        result = owner.set_override(student_id, standard, owl, actor)
                        self._json(200, result)
                    except LookupError:
                        self._json(404, {"error": "Student not found"})
                    except Exception as exc:
                        self._json(500, {"error": str(exc)})
                    return

                if path == "/api/batch_override":
                    actor = self._actor()
                    if not actor:
                        self._json(HTTPStatus.UNAUTHORIZED, {"error": "Sign in required"})
                        return
                    student_ids = body.get("student_ids", [])
                    standard = str(body.get("standard", "")).strip()
                    owl_mode = str(body.get("owl", "keep")).strip().lower()
                    if standard not in ("", *AWARD_ORDER):
                        self._json(HTTPStatus.BAD_REQUEST, {"error": "Invalid standard award"})
                        return
                    if owl_mode not in ("keep", "on", "off", "reset"):
                        self._json(HTTPStatus.BAD_REQUEST, {"error": "Invalid Owl operation"})
                        return
                    if not isinstance(student_ids, list) or not student_ids:
                        self._json(HTTPStatus.BAD_REQUEST, {"error": "Students required"})
                        return
                    try:
                        result = owner.set_batch_override(student_ids, standard, owl_mode, actor)
                        self._json(200, result)
                    except LookupError:
                        self._json(404, {"error": "Student not found"})
                    except Exception as exc:
                        self._json(500, {"error": str(exc)})
                    return

                if path == "/api/owl/simulate":
                    actor = self._actor()
                    if not actor:
                        self._json(HTTPStatus.UNAUTHORIZED, {"error": "Sign in required"})
                        return
                    try:
                        percentage = int(body.get("percentage", 0))
                    except (TypeError, ValueError):
                        self._json(HTTPStatus.BAD_REQUEST, {"error": "Invalid percentage"})
                        return
                    reserve = bool(body.get("reserve_overrides", True))
                    if percentage < 0 or percentage > 100:
                        self._json(HTTPStatus.BAD_REQUEST, {"error": "Percentage must be from 0 to 100"})
                        return
                    try:
                        result = owner.simulate_owl(percentage, reserve, actor)
                        self._json(200, result)
                    except Exception as exc:
                        self._json(500, {"error": str(exc)})
                    return

                self._json(404, {"error": "Not found"})

            def log_message(self, format, *args):
                return

            def _html(self):
                page = HTML_PAGE.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(page)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(page)

        self.server = ThreadingHTTPServer((self.host, self.port), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True,
                                       name="CSA-LAN-Portal")
        self.thread.start()
        return self

    def stop(self):
        if self.server:
            server = self.server
            thread = self.thread
            self.server = None
            if thread and thread.is_alive():
                server.shutdown()
            server.server_close()
            if thread:
                thread.join(timeout=2)
            self.thread = None

    def set_pin(self, pin):
        pin = str(pin)
        if not pin:
            raise ValueError("Access password cannot be empty")
        self.pin = pin

    def student_payload(self):
        with self.lock:
            session = SessionLocal()
            try:
                calculate_awards(session)
                session.commit()
                summaries = calculate_awards(session)
                rows = []
                for summary in summaries:
                    student = summary["student"]
                    rows.append({
                        "id": student.id,
                        "student_id": student.student_id or "",
                        "name": student.full_name,
                        "email": student.email,
                        "grade": student.grade_level or "",
                        "status": student.status,
                        "hours": round(summary["total_hours"], 1),
                        "current_award": summary["award_label"] or "None",
                        "standard_override": summary["admin_award_override"] or "",
                        "owl_override": bool(summary["admin_owl_override"]),
                        "counselor_writeup": any(p.owl_writeup for p in student.pin_history if p.pin_awarded == "Owl"),
                    })
                return {"students": rows, "updated_at": int(time.time())}
            finally:
                session.close()

    def set_override(self, student_id, standard, owl, actor):
        with self.lock:
            session = SessionLocal()
            try:
                student = session.query(Student).filter(
                    (Student.id == int(student_id))
                    if student_id.isdigit()
                    else (Student.student_id == student_id)
                ).one_or_none()
                if student is None:
                    raise LookupError(student_id)
                override = session.query(AwardOverride).filter_by(student_id=student.id).first()
                if not standard and not owl:
                    if override:
                        session.delete(override)
                    student.admin_award_override = None
                    student.owl_override = None
                    student.owl_selected = False
                    action = "reset override"
                else:
                    if override is None:
                        override = AwardOverride(student_id=student.id, award=standard or "Owl")
                        session.add(override)
                    override.standard_award = standard or None
                    override.owl_awarded = owl
                    override.award = standard or "Owl"
                    student.admin_award_override = None
                    student.owl_override = True if owl else None
                    student.owl_selected = owl
                    action = f"set {standard + ' + ' if standard else ''}{'Owl' if owl else ''}"
                session.commit()
                name = student.full_name
                entry = {"time": int(time.time()), "actor": actor,
                         "student": name, "action": action}
                self.activity.insert(0, entry)
                del self.activity[100:]
                return {"ok": True, "student": name, "action": action}
            except Exception:
                session.rollback()
                raise
            finally:
                session.close()

    def set_batch_override(self, student_ids, standard, owl_mode, actor):
        count = 0
        for raw_id in student_ids:
            student_id = str(raw_id).strip()
            if not student_id:
                continue
            session = SessionLocal()
            try:
                student = session.query(Student).filter(
                    (Student.id == int(student_id))
                    if student_id.isdigit()
                    else (Student.student_id == student_id)
                ).one_or_none()
                if student is None:
                    raise LookupError(student_id)
                override = session.query(AwardOverride).filter_by(student_id=student.id).first()
                if owl_mode == "reset":
                    owl = False
                    standard = ""
                elif owl_mode == "on":
                    owl = True
                elif owl_mode == "off":
                    owl = False
                else:
                    owl = bool(override.owl_awarded or override.award == "Owl") if override else False
                if not standard and not owl:
                    if override:
                        session.delete(override)
                    student.admin_award_override = None
                    student.owl_override = None
                    student.owl_selected = False
                else:
                    if override is None:
                        override = AwardOverride(student_id=student.id, award=standard or "Owl")
                        session.add(override)
                    override.standard_award = standard or None
                    override.owl_awarded = owl
                    override.award = standard or "Owl"
                    student.admin_award_override = None
                    student.owl_override = True if owl else None
                    student.owl_selected = owl
                session.commit()
                count += 1
            except Exception:
                session.rollback()
                raise
            finally:
                session.close()
        action = "reset overrides" if owl_mode == "reset" and not standard else f"batch set {standard or ('Owl' if owl_mode == 'on' else 'selection')}"
        self.activity.insert(0, {"time": int(time.time()), "actor": actor,
                                  "student": f"{count} students", "action": action})
        del self.activity[100:]
        return {"ok": True, "count": count, "action": action}

    def simulate_owl(self, percentage, reserve_overrides, actor):
        with self.lock:
            session = SessionLocal()
            try:
                summaries = calculate_awards(session)
                eligible = [summary for summary in summaries
                            if summary.get("hour_pin") == "Gold" and summary["student"].status == "ACTIVE"]
                eligible.sort(key=lambda summary: (
                    summary["student"].owl_rubric_score is None,
                    -(summary["student"].owl_rubric_score or 0),
                    summary["student"].full_name.lower()))
                manual_ids = {summary["student"].id for summary in eligible
                              if summary.get("admin_owl_override")}
                target = math.ceil(len(eligible) * percentage / 100) if eligible else 0
                available = max(0, target - len(manual_ids)) if reserve_overrides else target
                candidates = [summary for summary in eligible
                              if summary["student"].id not in manual_ids
                              and summary["student"].owl_rubric_score is not None]
                selected_ids = {summary["student"].id for summary in candidates[:available]}
                for student in session.query(Student).filter_by(status="ACTIVE").all():
                    if student.id not in manual_ids:
                        student.owl_selected = student.id in selected_ids
                session.commit()
                result = {"ok": True, "eligible": len(eligible), "target": target,
                          "selected": len(selected_ids)}
                self.activity.insert(0, {"time": int(time.time()), "actor": actor,
                                         "student": "Owl slots", "action": f"simulated {len(selected_ids)} at {percentage}%"})
                del self.activity[100:]
                return result
            except Exception:
                session.rollback()
                raise
            finally:
                session.close()


HTML_PAGE = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CSA Award Web Portal</title><style>
body{font-family:system-ui,-apple-system,sans-serif;background:#edf2f6;color:#17202a;margin:0}header{position:relative;background:#183b56;color:white;padding:16px 20px}header h1{margin:0;font-size:22px}.top-about{position:absolute;top:12px;right:16px;background:#315d7c;border:1px solid #9bc1d8!important;padding:8px 14px!important}.wrap{max-width:1100px;margin:auto;padding:16px}.card{background:white;border:1px solid #ccd7df;border-radius:8px;padding:14px;margin-bottom:12px}.toolbar{display:flex;gap:8px;flex-wrap:wrap}.toolbar input,.toolbar select{padding:9px;border:1px solid #abbac5;border-radius:5px}.toolbar input{flex:1;min-width:220px}.table-wrap{overflow:auto}table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:10px;border-bottom:1px solid #dde4e9;white-space:nowrap}th{background:#eef3f6}tr:hover td{background:#f8fafb}.muted{color:#657681;font-size:13px}.badge{font-weight:700}.override{background:#fff4c7}.controls{display:flex;gap:8px;align-items:center}.controls select{padding:7px}button{border:0;border-radius:5px;background:#2f6f9f;color:white;padding:8px 12px;cursor:pointer}button:hover{background:#4b91c1}.secondary{background:#657681}dialog{border:0;border-radius:8px;padding:20px;width:min(420px,calc(100% - 40px));box-shadow:0 12px 40px #0005}dialog input{width:100%;box-sizing:border-box;padding:10px;margin:7px 0 12px}#loginStatus,#saveStatus{min-height:20px}.activity{max-height:130px;overflow:auto}.activity div{padding:4px 0}@media(max-width:700px){header{padding-right:100px}th:nth-child(2),td:nth-child(2),th:nth-child(5),td:nth-child(5){display:none}.wrap{padding:8px}}
</style></head><body><header><h1>CSA Award Web Portal</h1><button class="top-about" id="aboutButton" type="button">About</button></header>
<div class="wrap"><div class="card toolbar"><input id="search" placeholder="Search name, student ID, or email"><select id="grade"><option value="">All grades</option><option>9</option><option>10</option><option>11</option><option>12</option></select><select id="status"><option value="">All statuses</option><option value="ACTIVE">Active</option><option value="GRADUATED">Inactive/Graduated</option></select><select id="award"><option value="">All awards</option><option>First Year</option><option>Bronze</option><option>Silver</option><option>Gold</option><option>Owl</option><option>None</option></select><span id="count" class="muted"></span></div>
<div class="card toolbar"><b>Batch selection</b><button class="secondary" type="button" onclick="toggleVisible(true)">Select visible</button><button class="secondary" type="button" onclick="toggleVisible(false)">Clear</button><select id="batch-standard"><option value="">Keep/Calculated</option><option value="__reset">Reset overrides</option><option>First Year</option><option>Bronze</option><option>Silver</option><option>Gold</option></select><select id="batch-owl"><option value="keep">Keep Owl</option><option value="on">Award Owl</option><option value="off">Remove Owl</option></select><button type="button" onclick="batchSave()">Apply to selected</button></div>
<div class="card toolbar"><b>Award Owl to top</b><label><input id="owl-percent" type="number" min="0" max="100" value="20"> % of eligible Gold students</label><button type="button" onclick="simulateOwl(true)">Simulate, reserve overrides</button><button type="button" onclick="simulateOwl(false)">Simulate full slots</button></div>
<div id="saveStatus" class="muted"></div><div class="card table-wrap"><table><thead><tr><th>Student</th><th>ID</th><th>Grade</th><th>Hours</th><th>Status</th><th>Award</th><th>Override</th></tr></thead><tbody id="rows"></tbody></table></div>
<div class="card"><b>Recent activity</b><div id="activity" class="activity muted">No changes yet</div></div></div>
<dialog id="login"><form id="loginForm"><h2>Shared administrator sign-in</h2><label for="name">Name shown in activity</label><input id="name" name="name" type="text" autocomplete="name" required><label for="pin">Shared portal password</label><input id="pin" name="pin" type="password" autocomplete="current-password" required><div id="loginStatus" class="muted"></div><button type="submit">Enter portal</button></form></dialog>
<dialog id="aboutDialog"><h2>MPS CSA Email Blast Manager</h2><p><b>Developer:</b> Rayan Nasir (<a href="https://github.com/usrr961" target="_blank" rel="noopener">@usrr961</a>)</p><p><b>Version:</b> 2.4.0</p><p>Automated utility for managing individualized Community Service Awards communication.</p><form method="dialog"><button>Close</button></form></dialog>
<script>
let rows=[],auth=null,timer;
let sessionToken=sessionStorage.getItem('csa_session_token')||'';
async function api(path,opt={}){opt=Object.assign({},opt);opt.headers=Object.assign({},opt.headers||{},sessionToken?{'Authorization':'Bearer '+sessionToken}:{}, {'Cache-Control':'no-store'});const r=await fetch(path,opt);const d=await r.json().catch(()=>({}));if(!r.ok)throw new Error(d.error||'Request failed');return d}
async function session(){try{const s=await api('/api/session');auth=s;render()}catch(e){}}
async function load(){try{const d=await api('/api/students');rows=d.students;render();await activity()}catch(e){if(e.message.includes('Sign in')){document.querySelector('#login').showModal()}}}
async function activity(){try{const d=await api('/api/activity');const el=document.querySelector('#activity');el.innerHTML=d.activity.length?d.activity.map(x=>`<div>${escapeHtml(x.actor)} · ${escapeHtml(x.action)} · ${escapeHtml(x.student)} · ${new Date(x.time*1000).toLocaleTimeString()}</div>`).join(''):'No changes yet'}catch(e){}}
function escapeHtml(v){const d=document.createElement('div');d.textContent=v==null?'':v;return d.innerHTML}
function render(){const q=document.querySelector('#search').value.toLowerCase(),grade=document.querySelector('#grade').value,status=document.querySelector('#status').value,award=document.querySelector('#award').value;const visible=rows.filter(s=>(!q||`${s.name} ${s.student_id} ${s.email}`.toLowerCase().includes(q))&&(!grade||s.grade===grade)&&(!status||s.status===status)&&(!award||s.current_award===award||s.current_award.split(' + ').includes(award)));document.querySelector('#count').textContent=`${visible.length} of ${rows.length} students`;document.querySelector('#rows').innerHTML=visible.map(s=>`<tr class="${s.standard_override||s.owl_override?'override':''}"><td><input class="row-check" type="checkbox" value="${s.id}" aria-label="Select ${escapeHtml(s.name)}"> <b>${escapeHtml(s.name)}</b><br><span class="muted">${escapeHtml(s.email)}</span></td><td>${escapeHtml(s.student_id)}</td><td>${escapeHtml(s.grade)}</td><td>${s.hours.toFixed(1)}</td><td>${escapeHtml(s.status)}</td><td class="badge">${escapeHtml(s.current_award)}</td><td><div class="controls"><select id="std-${s.id}" aria-label="Standard award"><option value="">Calculated</option>${['First Year','Bronze','Silver','Gold'].map(x=>`<option ${s.standard_override===x?'selected':''}>${x}</option>`).join('')}</select><label><input id="owl-${s.id}" type="checkbox" ${s.owl_override?'checked':''}> Owl</label><button onclick="save(${s.id})">Save</button></div></td></tr>`).join('')}
async function save(id){try{document.querySelector('#saveStatus').textContent='Saving…';const standard=document.querySelector('#std-'+id).value,owl=document.querySelector('#owl-'+id).checked;await api('/api/override',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({student_id:id,standard,owl})});document.querySelector('#saveStatus').textContent='Saved and shared with all connected administrators.';await load()}catch(e){document.querySelector('#saveStatus').textContent=e.message}}
function selectedIds(){return [...document.querySelectorAll('.row-check:checked')].map(x=>x.value)}
function toggleVisible(checked){document.querySelectorAll('.row-check').forEach(x=>x.checked=checked)}
async function batchSave(){const ids=selectedIds();if(!ids.length){document.querySelector('#saveStatus').textContent='Select at least one student.';return}const raw=document.querySelector('#batch-standard').value,owl=document.querySelector('#batch-owl').value;const standard=raw==='__reset'?'':raw,owlMode=raw==='__reset'?'reset':owl;try{const result=await api('/api/batch_override',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({student_ids:ids,standard,owl:owlMode})});document.querySelector('#saveStatus').textContent=`Updated ${result.count} selected students.`;await load()}catch(e){document.querySelector('#saveStatus').textContent=e.message}}
async function simulateOwl(reserveOverrides){const percentage=Number(document.querySelector('#owl-percent').value||0);try{const result=await api('/api/owl/simulate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({percentage,reserve_overrides:reserveOverrides})});document.querySelector('#saveStatus').textContent=`Owl experiment saved: ${result.selected} of ${result.eligible} eligible students selected (${percentage}%).`;await load()}catch(e){document.querySelector('#saveStatus').textContent=e.message}}
document.querySelector('#loginForm').addEventListener('submit',async e=>{e.preventDefault();const submit=e.submitter||e.currentTarget.querySelector('button[type="submit"]');if(submit)submit.disabled=true;try{const name=document.querySelector('#name').value.trim();const pin=document.querySelector('#pin').value;if(!name)throw new Error('Enter a name for activity records');if(!pin)throw new Error('Enter the shared portal password');const loginResult=await api('/api/login',{method:'POST',headers:{'Content-Type':'application/json','Cache-Control':'no-store'},body:JSON.stringify({name,pin})});if(loginResult.token){sessionToken=loginResult.token;sessionStorage.setItem('csa_session_token',sessionToken)}document.querySelector('#login').close();await session();await load()}catch(error){document.querySelector('#loginStatus').textContent=error.message||'Sign-in failed'}finally{if(submit)submit.disabled=false}});
['search','grade','status','award'].forEach(id=>document.querySelector('#'+id).addEventListener('input',render));
document.querySelector('#aboutButton').addEventListener('click',()=>document.querySelector('#aboutDialog').showModal());
session().then(load);timer=setInterval(load,2000);
</script></body></html>'''
