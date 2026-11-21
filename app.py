from pathlib import Path
import base64
import io
import json
import math
import os
import sys
import threading
import hashlib
import subprocess
import shutil
import shutil
import zipfile
import tempfile
import uuid
import secrets
from collections import defaultdict
from datetime import datetime, time, timedelta
from flask import Flask, render_template_string, request, Response, send_file, session, redirect, url_for, flash, send_from_directory
from github import Github
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename
from zk import ZK, const

app = Flask(__name__)

# Secret configurations using Environment Variables with fallbacks
app.secret_key = os.getenv('SECRET_KEY') or secrets.token_urlsafe(32)

DATA_DIR = os.getenv('DATA_DIR', os.getcwd())
os.makedirs(DATA_DIR, exist_ok=True)
def data_path(name): return os.path.join(DATA_DIR, name)
UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', data_path('uploads'))
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
BACKUP_FOLDER = os.getenv('BACKUP_FOLDER', os.path.join(os.getcwd(), 'portal_backups'))
AUTO_EXPORT_FOLDER = os.getenv('AUTO_EXPORT_FOLDER', os.path.join(os.getcwd(), 'portal_auto_exports'))
os.makedirs(BACKUP_FOLDER, exist_ok=True)
os.makedirs(AUTO_EXPORT_FOLDER, exist_ok=True)
BACKUP_RETENTION_DAYS = int(os.getenv('BACKUP_RETENTION_DAYS', '30'))
BACKUP_STATE = {'running': False, 'percent': 0, 'message': 'Ready', 'file': None, 'error': None}
BACKUP_LOCK = threading.Lock()
BACKUP_SCHEDULER_STARTED = False
BACKUP_SCHEDULER_LOCK = threading.Lock()
LAST_SCHEDULED_BACKUP_KEYS = set()

ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'doc', 'docx'}

MACHINE_IP = os.getenv('MACHINE_IP', '192.168.1.153')
PORT = int(os.getenv('MACHINE_PORT', 4370))

# GitHub Configurations via Environment Variables
GITHUB_TOKEN = os.getenv('GITHUB_TOKEN', '')
GITHUB_REPO_NAME = os.getenv('GITHUB_REPO_NAME', 'fresmartgamek-hue/Biometric')
GITHUB_BRANCH = os.getenv('GITHUB_BRANCH', 'main')

# Persistent Storage Files
LEAVE_JSON_FILE = os.getenv('LEAVE_JSON_FILE', data_path('leave_requests.json'))
LEAVE_EXCEL_FILE = os.getenv('LEAVE_EXCEL_FILE', data_path('leave_records.xlsx'))
MONTHLY_ROSTER_FILE = os.getenv('MONTHLY_ROSTER_FILE', data_path('monthly_rosters.json'))
OFF_CHANGE_JSON_FILE = os.getenv('OFF_CHANGE_JSON_FILE', data_path('off_change_requests.json'))
SHIFT_CHANGE_JSON_FILE = os.getenv('SHIFT_CHANGE_JSON_FILE', data_path('shift_change_requests.json'))
EXAMS_JSON_FILE = os.getenv('EXAMS_JSON_FILE', data_path('employee_exams.json'))
EXAM_SUBMISSIONS_JSON_FILE = os.getenv('EXAM_SUBMISSIONS_JSON_FILE', data_path('exam_submissions.json'))
STORES_JSON_FILE = os.getenv('STORES_JSON_FILE', data_path('stores.json'))
DEFAULT_STORE_ID = os.getenv('DEFAULT_STORE_ID', 'LM11')
PORTAL_PUBLIC_URL = os.getenv('PORTAL_PUBLIC_URL', '').strip().rstrip('/')
PORTAL_PORT = int(os.getenv('PORT', '5000'))

# Global storage for synced biometric logs
ATTENDANCE_LOGS_FILE = os.getenv('ATTENDANCE_LOGS_FILE', data_path('attendance_logs.json'))
def load_synced_attendance_logs():
    try:
        if os.path.isfile(ATTENDANCE_LOGS_FILE):
            with open(ATTENDANCE_LOGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
    except Exception as exc:
        print(f'Attendance log load error: {exc}')
    return []
SYNCED_ATTENDANCE_LOGS = load_synced_attendance_logs()
LAST_DEVICE_SYNC_TIME = None
ATTENDANCE_SYNC_INTERVAL_SECONDS = 30 * 60
ATTENDANCE_SYNC_LOCK = threading.Lock()
BACKGROUND_SYNC_START_LOCK = threading.Lock()
BACKGROUND_SYNC_THREAD = None

MASTER_EMPLOYEES = {
    'NWC2981': {'name': 'ANTONIO JOSE BANDOLA', 'off': 'SUNDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC3127': {'name': 'MATEUS ANTONIO DA COSTA BALMIRO', 'off': 'FRIDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC1525': {'name': 'ETY JOSÉ BANDUA MONTEIRO', 'off': 'MONDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC8328': {'name': 'JOAO MATIAS DOMINGOS', 'off': 'SATURDAY', 'dept': 'ADMIN - CCTV', 'shift': 'morning'},
    'NWC6661': {'name': 'FRANCISCO MUNDELE CHIVELA', 'off': 'WEDNESDAY', 'dept': 'ADMIN - CCTV', 'shift': 'morning'},
    'NWC1553': {'name': 'TIAGO SANDALA CHISSANHA', 'off': 'SUNDAY', 'dept': 'ADMIN - AUDITOR', 'shift': 'morning'},
    'NWC8350': {'name': 'LOLIVALDO ALBERTO MADEIRA', 'off': 'SUNDAY', 'dept': 'ADMIN - EDP', 'shift': 'morning'},
    'NWC5187': {'name': 'VICTOR NSOSI JOAO', 'off': 'MONDAY', 'dept': 'CASH - HEAD', 'shift': 'morning'},
    'NWC1168': {'name': 'ADELIA MBALOMBO CHIPEPI', 'off': 'SUNDAY', 'dept': 'CASH - HEAD', 'shift': 'morning'},
    'NWC2652': {'name': 'DULCE DOROTEIA GARCIA LUSITANO', 'off': 'MONDAY', 'dept': 'CASH - HEAD', 'shift': 'morning'},
    'NWC3381': {'name': 'ANDRE DE JESUS NGOLA JOSE', 'off': 'TUESDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC1983': {'name': 'PATRICIA SOLANGE FRANCISCO', 'off': 'THURSDAY', 'dept': 'CASH', 'shift': 'second'},
    'NWC8364': {'name': 'DIELUMBAKA AUGUSTO', 'off': 'WEDNESDAY', 'dept': 'CASH', 'shift': 'second'},
    'NWC2788': {'name': 'INES NACHINGOLO FELICIANO NAMBELO', 'off': 'TUESDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC1010': {'name': 'TERESA PEDRO LEAO', 'off': 'FRIDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC6638': {'name': 'CLAUDIO JANUARIO MANUEL AVELINO', 'off': 'SUNDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC5830': {'name': 'REGINA DE FATIMA VIDAL', 'off': 'MONDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC5529': {'name': 'ALEXANDRE LUIS CORREIA', 'off': 'FRIDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC5713': {'name': 'ROSA GARNEIRA BUMBA', 'off': 'WEDNESDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC5396': {'name': 'COSTA BEBIANO HEBO', 'off': 'THURSDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC8361': {'name': 'AGOSTINHO JOAQUIM KUANGO DA COSTA', 'off': 'SUNDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC2300': {'name': 'JOANA CARDOSO JOAQUIM AFONSO', 'off': 'MONDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC5168': {'name': 'JOAO NVUNDA DALA', 'off': 'THURSDAY', 'dept': 'F & V', 'shift': 'morning'},
    'NWC3711': {'name': 'ANGELA MARIA BUMBA', 'off': 'FRIDAY', 'dept': 'F & V', 'shift': 'morning'},
    'NWC5186': {'name': 'HELIA DOMINGOS DE CARVALHO', 'off': 'WEDNESDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC6702': {'name': 'DOMINGOS GAMA PEREIRA', 'off': 'FRIDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC3596': {'name': 'ALDAIR FERNANDES FERREIRA', 'off': 'TUESDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC2757': {'name': 'JOSEFA KUELUNGA MUASSOKA', 'off': 'THURSDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC4554': {'name': 'DOMINGOS ANTONIO FERNANDO', 'off': 'THURSDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC2624': {'name': 'CECILIA JORGE FAMOSO', 'off': 'FRIDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC3318': {'name': 'JOSE MANUEL KAZOLA', 'off': 'SUNDAY', 'dept': 'FRESCO', 'shift': 'morning'},
    'NWC7347': {'name': 'ARMANDO CHICOVO SAMBA', 'off': 'TUESDAY', 'dept': 'FRESCO', 'shift': 'morning'},
    'NWC8362': {'name': 'ARAUJO PAULOMENDES', 'off': 'FRIDAY', 'dept': 'STOCK', 'shift': 'morning'},
    'NWC6715': {'name': 'RIBEIRO ANTONIO FRANCISCO', 'off': 'THURSDAY', 'dept': 'STOCK', 'shift': 'morning'},
    'NWC6444': {'name': 'HENRIQUES BRANDAO', 'off': 'WEDNESDAY', 'dept': 'STOCK', 'shift': 'morning'}
}

EMPLOYEE_OVERRIDES = {
    '8364': {'code': 'NWC8364', 'name': 'DIELUMBAKA AUGUSTO'},
    '1': {'code': 'NWC8350', 'name': 'LOLIVALDO ALBERTO MADEIRA'},
    '8362': {'code': 'NWC8362', 'name': 'ARAUJO PAULOMENDES'},
    '6661': {'code': 'NWC6661', 'name': 'FRANCISCO MUNDELE CHIVELA'}
}

# ===== BIOMETRIC EMPLOYEE ENROLLMENT - COMPLETE CODE ID =====
EMPLOYEES_JSON_FILE = os.getenv('EMPLOYEES_JSON_FILE', data_path('employees.json'))

def load_dynamic_employees():
    try:
        if os.path.exists(EMPLOYEES_JSON_FILE):
            with open(EMPLOYEES_JSON_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
    except Exception as exc:
        print(f'Employee data load error: {exc}')
    return {}

def save_dynamic_employees():
    temp = EMPLOYEES_JSON_FILE + '.tmp'
    with open(temp, 'w', encoding='utf-8') as f:
        json.dump(DYNAMIC_EMPLOYEES, f, indent=2, ensure_ascii=False)
    os.replace(temp, EMPLOYEES_JSON_FILE)

DYNAMIC_EMPLOYEES = load_dynamic_employees()
MASTER_EMPLOYEES.update(DYNAMIC_EMPLOYEES)

def next_employee_code():
    nums=[]
    for code in MASTER_EMPLOYEES:
        raw=str(code).upper().replace('NWC','',1)
        if raw.isdigit(): nums.append(int(raw))
    return f"NWC{(max(nums) if nums else 0)+1:04d}"

def biometric_user_id(employee_code):
    code=str(employee_code).strip().upper()
    if not code.startswith('NWC') or not code[3:].isdigit():
        raise ValueError('Employee code NWC + number format me hona chahiye, jaise NWC0608.')
    return code

def push_employee_to_biometric(employee_code, employee_name):
    # ZKTeco: uid internal numeric slot hai; user_id device par visible complete code hai.
    full_code=biometric_user_id(employee_code)
    numeric_part=full_code[3:]
    default_uid=int(numeric_part)
    zk=ZK(MACHINE_IP, port=PORT, timeout=10, password=0, force_udp=False, ommit_ping=False)
    conn=zk.connect()
    if not conn: raise ConnectionError('Biometric machine connection failed.')
    try:
        conn.disable_device()
        users=conn.get_users() or []
        by_user_id={str(u.user_id).strip().upper():u for u in users}
        if full_code in by_user_id:
            raise ValueError(f'{full_code} machine par already exists.')
        # Old patch ne 0608 bheja tha. Us entry ka same internal UID reuse karke
        # visible ID ko NWC0608 banaya jata hai, taaki linked template preserve rahe.
        legacy=by_user_id.get(numeric_part) or by_user_id.get(str(int(numeric_part)))
        target_uid=int(getattr(legacy,'uid',default_uid)) if legacy else default_uid
        uid_owner=next((u for u in users if int(getattr(u,'uid',-1))==target_uid),None)
        if uid_owner and not legacy:
            raise ValueError(f'Internal UID {target_uid} already used by {getattr(uid_owner,"user_id","another employee")}.')
        conn.set_user(uid=target_uid, name=employee_name[:24], privilege=const.USER_DEFAULT,
                      password='', group_id='', user_id=full_code, card=0)
        return full_code
    finally:
        try: conn.enable_device()
        except Exception: pass
        conn.disconnect()

BIOMETRIC_ENROLL_TEMPLATE=r'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><script src="https://cdn.tailwindcss.com"></script><title>Biometric Employee Enrollment</title></head><body class="min-h-screen bg-slate-100 text-slate-800"><main class="mx-auto max-w-6xl p-5 space-y-5"><header class="rounded-3xl bg-slate-950 p-6 text-white flex justify-between gap-3"><div><p class="text-xs font-black tracking-widest text-emerald-300">DEVELOPER CONTROL</p><h1 class="text-3xl font-black">Biometric Employee Enrollment</h1><p class="mt-2 text-sm text-slate-400">Portal me employee banayein aur complete code device par push karein.</p></div><a href="/" class="h-fit rounded-xl bg-white/10 px-4 py-2 font-bold">Dashboard</a></header>
{% for c,m in get_flashed_messages(with_categories=true) %}<div class="rounded-xl p-4 font-bold {{'bg-emerald-50 text-emerald-700' if c=='success' else 'bg-rose-50 text-rose-700'}}">{{m}}</div>{% endfor %}
<section class="grid lg:grid-cols-3 gap-5"><form method="post" class="lg:col-span-2 rounded-3xl bg-white p-6 shadow space-y-4"><h2 class="text-2xl font-black">Generate and Push</h2><div class="grid md:grid-cols-2 gap-3"><label class="text-xs font-black">EMPLOYEE NAME<input name="employee_name" required maxlength="80" class="mt-1 w-full rounded-xl border p-3 text-sm"></label><label class="text-xs font-black">EMPLOYEE CODE<input name="employee_code" value="{{suggested_code}}" required pattern="NWC[0-9]+" class="mt-1 w-full rounded-xl border p-3 text-sm font-mono"></label><label class="text-xs font-black">STORE<select name="store_id" required class="mt-1 w-full rounded-xl border p-3 text-sm">{% for sid,x in stores %}<option value="{{sid}}">{{x.store_name}} ({{sid}})</option>{% endfor %}</select></label><label class="text-xs font-black">DEPARTMENT / DESIGNATION<input name="department" required class="mt-1 w-full rounded-xl border p-3 text-sm"></label><label class="text-xs font-black">WEEKLY OFF<select name="weekly_off" class="mt-1 w-full rounded-xl border p-3 text-sm">{% for d in weekdays %}<option>{{d}}</option>{% endfor %}</select></label><label class="text-xs font-black">DEFAULT SHIFT<select name="shift" class="mt-1 w-full rounded-xl border p-3 text-sm"><option value="morning">Shift A</option><option value="second">Shift B</option></select></label></div><label class="flex gap-2 rounded-xl bg-amber-50 p-3 text-sm font-bold text-amber-800"><input type="checkbox" name="push_now" value="yes" checked> Save ke saath biometric machine par abhi push karein</label><button class="w-full rounded-xl bg-gradient-to-r from-emerald-600 to-cyan-600 p-4 font-black text-white">Generate Employee and Push</button></form><aside class="rounded-3xl bg-white p-6 shadow"><p class="text-xs font-black text-emerald-600">DEVICE TARGET</p><p class="mt-2 text-xl font-black">{{machine_ip}}:{{machine_port}}</p><div class="mt-5 rounded-2xl bg-slate-100 p-4 text-xs"><b>Next code:</b> {{suggested_code}}<br><b>Biometric ID:</b> Complete employee code, jaise NWC0608</div></aside></section>
<section class="rounded-3xl bg-white p-6 shadow"><h2 class="text-xl font-black mb-4">Recently Created Employees</h2><div class="overflow-auto"><table class="w-full text-sm"><thead><tr class="bg-slate-900 text-white"><th class="p-3 text-left">Code</th><th>Name</th><th>Department</th><th>Off</th><th>Shift</th><th>Device ID</th></tr></thead><tbody>{% for code,e in employees %}<tr class="border-b"><td class="p-3 font-mono font-bold">{{code}}</td><td class="text-center">{{e.name}}</td><td class="text-center">{{e.dept}}</td><td class="text-center">{{e.off}}</td><td class="text-center">{{'Shift A' if e.shift=='morning' else 'Shift B'}}</td><td class="text-center font-mono">{{e.biometric_user_id or code}}</td></tr>{% else %}<tr><td colspan="6" class="p-8 text-center text-slate-400">No developer-created employees yet.</td></tr>{% endfor %}</tbody></table></div></section></main></body></html>'''

@app.route('/developer/biometric-employees', methods=['GET','POST'])
def developer_biometric_employees():
    if not session.get('logged_in') or session.get('role')!='developer': return 'Developer access required',403
    if request.method=='POST':
        name=' '.join(request.form.get('employee_name','').strip().upper().split())
        code=normalized_code(request.form.get('employee_code','').strip().upper())
        dept=' '.join(request.form.get('department','').strip().upper().split()); store_id=safe_store_id(request.form.get('store_id'))
        weekly_off=request.form.get('weekly_off','SUNDAY').upper(); shift=request.form.get('shift','morning')
        days=['MONDAY','TUESDAY','WEDNESDAY','THURSDAY','FRIDAY','SATURDAY','SUNDAY']
        if not name or not dept or store_id not in STORES or weekly_off not in days or shift not in ['morning','second']:
            flash('Employee details invalid hain.','error'); return redirect(url_for('developer_biometric_employees'))
        try: biometric_user_id(code)
        except ValueError as exc:
            flash(str(exc),'error'); return redirect(url_for('developer_biometric_employees'))
        if code in MASTER_EMPLOYEES:
            flash(f'{code} portal me already exists.','error'); return redirect(url_for('developer_biometric_employees'))
        pushed_id=None
        if request.form.get('push_now')=='yes':
            try: pushed_id=push_employee_to_biometric(code,name)
            except Exception as exc:
                flash('Biometric push failed; employee save nahi hua: '+str(exc),'error'); return redirect(url_for('developer_biometric_employees'))
        record={'name':name,'off':weekly_off,'dept':dept,'shift':shift,'store_id':store_id,'created_at':datetime.now().isoformat(timespec='seconds'),'created_by':session.get('user_id'),'biometric_user_id':pushed_id,'device_ip':MACHINE_IP if pushed_id else None}
        DYNAMIC_EMPLOYEES[code]=record; MASTER_EMPLOYEES[code]=record
        try: save_dynamic_employees()
        except Exception as exc:
            DYNAMIC_EMPLOYEES.pop(code,None); MASTER_EMPLOYEES.pop(code,None)
            flash('Employee save failed: '+str(exc),'error'); return redirect(url_for('developer_biometric_employees'))
        flash(f'{code} - {name} created'+(f' and pushed as biometric ID {pushed_id}.' if pushed_id else '.'),'success')
        return redirect(url_for('developer_biometric_employees'))
    employees=sorted(DYNAMIC_EMPLOYEES.items(),key=lambda x:x[1].get('created_at',''),reverse=True)
    return render_template_string(BIOMETRIC_ENROLL_TEMPLATE,suggested_code=next_employee_code(),employees=employees,weekdays=['MONDAY','TUESDAY','WEDNESDAY','THURSDAY','FRIDAY','SATURDAY','SUNDAY'],machine_ip=MACHINE_IP,machine_port=PORT,stores=sorted(STORES.items()))
# ===== END BIOMETRIC EMPLOYEE ENROLLMENT =====


# ===== ATTENDANCE REGULARISATION MODULE =====
REGULARISATION_FILE = os.getenv('REGULARISATION_FILE', data_path('attendance_regularisation.json'))
BIOMETRIC_PUSH_QUEUE_FILE = os.getenv('BIOMETRIC_PUSH_QUEUE_FILE', data_path('biometric_attendance_push_queue.json'))

def _load_json_records(path):
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data=json.load(f)
                return data if isinstance(data,list) else []
    except Exception as exc: print(f'JSON load error {path}: {exc}')
    return []

def _atomic_save_json(path, data):
    temp=path+'.tmp'
    with open(temp,'w',encoding='utf-8') as f: json.dump(data,f,indent=2,ensure_ascii=False)
    os.replace(temp,path)

REGULARISATION_REQUESTS=_load_json_records(REGULARISATION_FILE)

def save_regularisation_requests(): _atomic_save_json(REGULARISATION_FILE, REGULARISATION_REQUESTS)

def approved_regularisation_punches():
    punches=[]
    for item in REGULARISATION_REQUESTS:
        if item.get('status')!='Approved': continue
        for key in ('in_time','out_time'):
            value=item.get(key)
            if value:
                try: timestamp=datetime.strptime(item['date']+' '+value,'%Y-%m-%d %H:%M')
                except ValueError: continue
                punches.append({'user_id':item['employee_code'],'timestamp':timestamp,'regularisation_id':item['id']})
    return punches

def queue_biometric_regularisation(item, error=None):
    queue=_load_json_records(BIOMETRIC_PUSH_QUEUE_FILE)
    queue=[x for x in queue if x.get('request_id')!=item['id']]
    queue.append({'request_id':item['id'],'employee_code':item['employee_code'],'date':item['date'],
                  'in_time':item['in_time'],'out_time':item['out_time'],'status':'Pending',
                  'attempts':item.get('device_push_attempts',0),'last_error':error,
                  'queued_at':datetime.now().isoformat(timespec='seconds')})
    _atomic_save_json(BIOMETRIC_PUSH_QUEUE_FILE,queue)

def push_regularisation_to_biometric(item):
    # Standard pyzk/ZKTeco protocol can read attendance logs but usually cannot
    # create historical punch logs on the terminal. Hand off to the local bridge
    # queue without treating this device limitation as an application error.
    queue_biometric_regularisation(item, None)
    return 'Bridge Queued'

REGULARISATION_TEMPLATE=r'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><script src="https://cdn.tailwindcss.com"></script><title>Attendance Regularisation</title></head><body class="min-h-screen bg-slate-100 text-slate-800"><main class="mx-auto max-w-7xl p-4 md:p-7 space-y-5"><header class="rounded-[28px] bg-slate-950 p-6 text-white flex flex-wrap justify-between gap-4"><div><p class="text-xs font-black tracking-widest text-cyan-300">ATTENDANCE CONTROL</p><h1 class="text-3xl font-black">Attendance Regularisation</h1><p class="mt-2 text-sm text-slate-400">Employee request → Admin approval → Portal attendance correction → local bridge sync queue.</p></div><a href="/" class="h-fit rounded-xl bg-white/10 px-4 py-2 font-bold">Dashboard</a></header>
{% for c,m in get_flashed_messages(with_categories=true) %}<div class="rounded-xl p-4 font-bold {{'bg-emerald-50 text-emerald-700' if c=='success' else 'bg-rose-50 text-rose-700'}}">{{m}}</div>{% endfor %}
{% if role=='employee' %}<section class="rounded-3xl bg-white p-6 shadow"><h2 class="text-2xl font-black">Submit Correction Request</h2><form method="post" action="/attendance-regularisation/submit" class="mt-5 grid md:grid-cols-5 gap-3 items-end"><label class="text-xs font-black">DATE<input type="date" name="date" max="{{today}}" required class="mt-1 w-full rounded-xl border p-3"></label><label class="text-xs font-black">IN TIME<input type="time" name="in_time" required class="mt-1 w-full rounded-xl border p-3"></label><label class="text-xs font-black">OUT TIME<input type="time" name="out_time" required class="mt-1 w-full rounded-xl border p-3"></label><label class="text-xs font-black">REASON<input name="reason" required maxlength="250" class="mt-1 w-full rounded-xl border p-3" placeholder="Missed punch / device issue"></label><button class="rounded-xl bg-indigo-600 p-3 font-black text-white">Submit Request</button></form></section>{% endif %}
<section class="rounded-3xl bg-white p-5 shadow"><div class="flex justify-between"><h2 class="text-xl font-black">{{'Pending Admin Actions' if role in ['admin','developer'] else 'My Requests'}}</h2><span class="rounded-full bg-slate-100 px-3 py-1 text-xs font-bold">{{rows|length}} requests</span></div><div class="mt-4 overflow-auto"><table class="w-full min-w-[1050px] text-sm"><thead><tr class="bg-slate-900 text-white"><th class="p-3 text-left">Employee</th><th>Date</th><th>In</th><th>Out</th><th>Reason</th><th>Status</th><th>Admin Note</th><th>Device Push</th>{% if role in ['admin','developer'] %}<th>Action</th>{% endif %}</tr></thead><tbody>{% for r in rows %}<tr class="border-b"><td class="p-3"><b>{{r.employee_name}}</b><br><small>{{r.employee_code}}</small></td><td class="text-center">{{r.date}}</td><td class="text-center font-mono">{{r.in_time}}</td><td class="text-center font-mono">{{r.out_time}}</td><td class="max-w-xs">{{r.reason}}</td><td class="text-center font-black">{{r.status}}</td><td class="text-center">{{r.admin_note or '-'}}</td><td class="text-center"><b>{{r.device_push_status or 'Not attempted'}}</b>{% if r.device_push_error %}<br><small class="text-rose-600">{{r.device_push_error}}</small>{% endif %}</td>{% if role in ['admin','developer'] %}<td class="p-2">{% if r.status=='Pending' %}<form method="post" action="/attendance-regularisation/{{r.id}}/decision" class="flex gap-2"><input name="admin_note" class="min-w-32 rounded-lg border p-2 text-xs" placeholder="Optional note"><button name="decision" value="approve" class="rounded-lg bg-emerald-600 px-3 py-2 font-bold text-white">Approve</button><button name="decision" value="reject" class="rounded-lg bg-rose-600 px-3 py-2 font-bold text-white">Reject</button></form>{% elif r.status=='Approved' and r.device_push_status=='Queued' %}<form method="post" action="/attendance-regularisation/{{r.id}}/retry-device"><button class="rounded-lg bg-cyan-600 px-3 py-2 font-bold text-white">Queue Again</button></form>{% else %}-{% endif %}</td>{% endif %}</tr>{% else %}<tr><td colspan="9" class="p-10 text-center text-slate-400">No regularisation requests.</td></tr>{% endfor %}</tbody></table></div></section></main></body></html>'''

@app.route('/attendance-regularisation')
def attendance_regularisation():
    if not session.get('logged_in'): return redirect(url_for('login'))
    role=session.get('role'); code=session.get('user_id')
    rows=REGULARISATION_REQUESTS if role in ['admin','developer'] else [r for r in REGULARISATION_REQUESTS if r.get('employee_code')==code]
    rows=sorted(rows,key=lambda r:r.get('submitted_at',''),reverse=True)
    return render_template_string(REGULARISATION_TEMPLATE,role=role,rows=rows,today=datetime.now().strftime('%Y-%m-%d'))

@app.route('/attendance-regularisation/submit',methods=['POST'])
def submit_attendance_regularisation():
    if session.get('role')!='employee': return 'Employee access required',403
    code=session.get('user_id'); date_value=request.form.get('date',''); in_time=request.form.get('in_time',''); out_time=request.form.get('out_time',''); reason=request.form.get('reason','').strip()
    try:
        day=datetime.strptime(date_value,'%Y-%m-%d').date(); start=datetime.strptime(in_time,'%H:%M'); end=datetime.strptime(out_time,'%H:%M')
        if day>datetime.now().date(): raise ValueError('Future date allowed nahi hai.')
        if end<=start: raise ValueError('Out time, In time ke baad hona chahiye.')
    except ValueError as exc:
        flash('Invalid request: '+str(exc),'error'); return redirect(url_for('attendance_regularisation'))
    if any(r for r in REGULARISATION_REQUESTS if r.get('employee_code')==code and r.get('date')==date_value and r.get('status')=='Pending'):
        flash('Is date ka request already pending hai.','error'); return redirect(url_for('attendance_regularisation'))
    info=get_emp_info(code)
    REGULARISATION_REQUESTS.append({'id':uuid.uuid4().hex[:12],'employee_code':code,'employee_name':info['name'],'date':date_value,'in_time':in_time,'out_time':out_time,'reason':reason,'status':'Pending','submitted_at':datetime.now().isoformat(timespec='seconds'),'admin_note':'','approved_by':None,'approved_at':None,'device_push_status':'Not attempted','device_push_error':None,'device_push_attempts':0})
    save_regularisation_requests(); flash('Attendance regularisation admin ko approval ke liye bhej diya gaya.','success')
    return redirect(url_for('attendance_regularisation'))

@app.route('/attendance-regularisation/<request_id>/decision',methods=['POST'])
def decide_attendance_regularisation(request_id):
    if session.get('role') not in ['admin','developer']: return 'Admin access required',403
    item=next((r for r in REGULARISATION_REQUESTS if r.get('id')==request_id),None)
    if not item: return 'Request not found',404
    if item.get('status')!='Pending': flash('Request already processed.','error'); return redirect(url_for('attendance_regularisation'))
    decision=request.form.get('decision'); item['admin_note']=request.form.get('admin_note','').strip(); item['approved_by']=session.get('user_id'); item['approved_at']=datetime.now().isoformat(timespec='seconds')
    if decision=='reject':
        item['status']='Rejected'; save_regularisation_requests(); flash('Regularisation rejected.','success'); return redirect(url_for('attendance_regularisation'))
    item['status']='Approved'; item['device_push_attempts']=1
    try:
        result=push_regularisation_to_biometric(item); item['device_push_status']=result; item['device_push_error']=None; flash('Approved: portal attendance corrected. Device sync local bridge queue me add ho gaya.','success')
    except Exception as exc:
        item['device_push_status']='Queued'; item['device_push_error']=str(exc); queue_biometric_regularisation(item,str(exc)); flash('Approved: portal attendance corrected. Device push queue me hai.','success')
    save_regularisation_requests(); return redirect(url_for('attendance_regularisation'))

@app.route('/attendance-regularisation/<request_id>/retry-device',methods=['POST'])
def retry_regularisation_device(request_id):
    if session.get('role') not in ['admin','developer']: return 'Admin access required',403
    item=next((r for r in REGULARISATION_REQUESTS if r.get('id')==request_id and r.get('status')=='Approved'),None)
    if not item: return 'Approved request not found',404
    item['device_push_attempts']=int(item.get('device_push_attempts',0))+1
    try:
        result=push_regularisation_to_biometric(item); item['device_push_status']=result; item['device_push_error']=None; flash('Device sync request local bridge queue me add ho gaya.','success')
    except Exception as exc:
        item['device_push_status']='Queued'; item['device_push_error']=str(exc); queue_biometric_regularisation(item,str(exc)); flash('Device push abhi queued hai: '+str(exc),'error')
    save_regularisation_requests(); return redirect(url_for('attendance_regularisation'))
# ===== END ATTENDANCE REGULARISATION MODULE =====

# Helper function to check allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Persistent JSON storage helpers for leave requests
def load_leave_requests():
    if os.path.exists(LEAVE_JSON_FILE):
        try:
            with open(LEAVE_JSON_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading {LEAVE_JSON_FILE}: {e}")
            return []
    return []

def save_leave_requests(leave_list):
    try:
        with open(LEAVE_JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump(leave_list, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error saving {LEAVE_JSON_FILE}: {e}")

LEAVE_REQUESTS = load_leave_requests()

def load_off_change_requests():
    if os.path.exists(OFF_CHANGE_JSON_FILE):
        try:
            with open(OFF_CHANGE_JSON_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading {OFF_CHANGE_JSON_FILE}: {e}")
    return []

def save_off_change_requests():
    try:
        with open(OFF_CHANGE_JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump(OFF_CHANGE_REQUESTS, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving {OFF_CHANGE_JSON_FILE}: {e}")
        return False

OFF_CHANGE_REQUESTS = load_off_change_requests()

def load_shift_change_requests():
    if os.path.exists(SHIFT_CHANGE_JSON_FILE):
        try:
            with open(SHIFT_CHANGE_JSON_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading {SHIFT_CHANGE_JSON_FILE}: {e}")
    return []

def save_shift_change_requests():
    try:
        with open(SHIFT_CHANGE_JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump(SHIFT_CHANGE_REQUESTS, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving {SHIFT_CHANGE_JSON_FILE}: {e}")
        return False

SHIFT_CHANGE_REQUESTS = load_shift_change_requests()

def normalized_code(code):
    code = str(code or '').strip()
    return code if code.startswith('NWC') else f"NWC{code}"

def effective_roster_assignment(emp_code, date_str):
    code = normalized_code(emp_code)
    assignment = get_monthly_roster_assignment(code, date_str)
    if not assignment:
        info = get_emp_info(code)
        assignment = 'Shift A' if info.get('shift') == 'morning' else 'Shift B'
    for req in SHIFT_CHANGE_REQUESTS:
        if req.get('status') == 'Approved' and req.get('swap_date') == date_str:
            if req.get('requester_id') == code:
                return req.get('partner_shift')
            if req.get('partner_id') == code:
                return req.get('requester_shift')
    return assignment

def refresh_dual_approval(req):
    if req.get('employee_approval') == 'Approved' and req.get('admin_approval') == 'Approved':
        req['status'] = 'Approved'
        req['approved_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    elif req.get('employee_approval') == 'Rejected' or req.get('admin_approval') == 'Rejected':
        req['status'] = 'Rejected'
    else:
        req['status'] = 'Pending'

def get_approved_off_change(emp_code, date_str):
    code = str(emp_code).strip()
    if not code.startswith('NWC'):
        code = f"NWC{code}"
    for req in OFF_CHANGE_REQUESTS:
        if req.get('status') == 'Approved':
            # Requester gives the old off to the partner and receives the partner's off date.
            if req.get('user_id') == code:
                if req.get('old_off_date') == date_str:
                    return 'WORKING'
                if req.get('new_off_date') == date_str:
                    return 'OFF'
            # Partner receives requester's old off and gives the selected new off date.
            if req.get('partner_id') == code:
                if req.get('new_off_date') == date_str:
                    return 'WORKING'
                if req.get('old_off_date') == date_str:
                    return 'OFF'
    return None

def load_monthly_rosters():
    if os.path.exists(MONTHLY_ROSTER_FILE):
        try:
            with open(MONTHLY_ROSTER_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading {MONTHLY_ROSTER_FILE}: {e}")
    return {}

def save_monthly_rosters():
    try:
        with open(MONTHLY_ROSTER_FILE, 'w', encoding='utf-8') as f:
            json.dump(MONTHLY_ROSTERS, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving {MONTHLY_ROSTER_FILE}: {e}")
        return False

MONTHLY_ROSTERS = load_monthly_rosters()

def get_emp_info(emp_code):
    emp_str = str(emp_code).strip()
    if not emp_str.startswith('NWC') and f"NWC{emp_str}" in MASTER_EMPLOYEES:
        emp_str = f"NWC{emp_str}"
    
    val = MASTER_EMPLOYEES.get(emp_str, {'name': f'Employee {emp_code}', 'off': 'SUNDAY', 'dept': 'General', 'shift': 'morning'})
    if isinstance(val, str):
        return {'name': val, 'off': 'SUNDAY', 'dept': 'General', 'shift': 'morning'}
    return val

def get_default_roster_patterns(emp_code):
    info = get_emp_info(emp_code)
    shift = 'Shift B' if info['shift'] == 'second' else 'Shift A'
    weekdays = ['Off' if day == info['off'].upper() else shift for day in (
        'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'
    )]
    return {'week1': weekdays[:], 'week2': weekdays[:]}

def get_monthly_roster_patterns(month):
    saved_roster = MONTHLY_ROSTERS.get(month, {})
    patterns = {}
    for emp_code in MASTER_EMPLOYEES:
        final_code = f"NWC{emp_code}" if not emp_code.startswith('NWC') else emp_code
        defaults = get_default_roster_patterns(emp_code)
        saved = saved_roster.get(final_code, {})
        patterns[final_code] = {
            week: saved.get(week, defaults[week])
            for week in ('week1', 'week2')
        }
    return patterns

def get_monthly_roster_assignment(emp_code, date_str):
    month = date_str[:7]
    saved = MONTHLY_ROSTERS.get(month, {}).get(emp_code, {})
    # New format stores exact YYYY-MM-DD assignments.
    if isinstance(saved, dict) and 'dates' in saved:
        return saved.get('dates', {}).get(date_str)
    # Backward compatibility with old week1/week2 pattern files.
    if isinstance(saved, dict) and 'week1' in saved and 'week2' in saved:
        return get_roster_assignment_from_patterns(saved, datetime.strptime(date_str, '%Y-%m-%d').date())
    return None

def get_roster_assignment_from_patterns(patterns, current_date):
    # Monday is always the first day of a roster week. The week containing
    # the first day of the month is Week 1; following weeks alternate A/B.
    month_first = current_date.replace(day=1)
    first_monday = month_first - timedelta(days=month_first.weekday())
    week_index = (current_date - first_monday).days // 7
    pattern_name = 'week1' if week_index % 2 == 0 else 'week2'
    return patterns[pattern_name][current_date.weekday()]

def build_month_dates(month):
    first = datetime.strptime(month + '-01', '%Y-%m-%d').date()
    next_month = (first + timedelta(days=32)).replace(day=1)
    return [first + timedelta(days=i) for i in range((next_month-first).days)]

def generate_alternating_month_roster(month):
    dates = build_month_dates(month)
    first_monday = dates[0] - timedelta(days=dates[0].weekday())
    generated = {}
    for raw_code in MASTER_EMPLOYEES:
        code = normalized_code(raw_code)
        info = get_emp_info(code)
        first_shift = 'Shift B' if info.get('shift') == 'second' else 'Shift A'
        second_shift = 'Shift A' if first_shift == 'Shift B' else 'Shift B'
        off_day = str(info.get('off', 'SUNDAY')).upper()
        date_map = {}
        for d in dates:
            week_index = (d - first_monday).days // 7
            shift = first_shift if week_index % 2 == 0 else second_shift
            date_map[d.strftime('%Y-%m-%d')] = 'Off' if d.strftime('%A').upper() == off_day else shift
        generated[code] = {'dates': date_map, 'weekly_off': off_day, 'first_week_shift': first_shift}
    return generated

def upload_file_to_github(file_path, github_destination_path):
    """Local file ko GitHub repository ke folder me upload karta hai"""
    if not GITHUB_TOKEN:
        print("GitHub token default or missing, skipping GitHub upload.")
        return False
    try:
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(GITHUB_REPO_NAME)
        
        with open(file_path, "rb") as f:
            content = f.read()
            
        try:
            file = repo.get_contents(github_destination_path, ref=GITHUB_BRANCH)
            repo.update_file(
                path=github_destination_path,
                message=f"Update leave document: {github_destination_path}",
                content=content,
                sha=file.sha,
                branch=GITHUB_BRANCH
            )
        except Exception:
            repo.create_file(
                path=github_destination_path,
                message=f"Upload leave document: {github_destination_path}",
                content=content,
                branch=GITHUB_BRANCH
            )
        return True
    except Exception as e:
        print(f"GitHub Upload Error: {e}")
        return False

def save_leave_to_excel(emp_code, emp_name, start_date, end_date, leave_type, doc_filename):
    """Excel file me employee ki leave details save karta hai"""
    if os.path.exists(LEAVE_EXCEL_FILE):
        wb = openpyxl.load_workbook(LEAVE_EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Leave Records"
        ws.append(["Employee Code", "Employee Name", "Start Date", "End Date", "Leave Type", "Document Name"])
        
    ws.append([emp_code, emp_name, start_date, end_date, leave_type, doc_filename])
    wb.save(LEAVE_EXCEL_FILE)

def check_device_connectivity():
    try:
        if selected_store_id and selected_store_id != DEFAULT_STORE_ID:
            raise ConnectionError('Cloud bridge mode for selected store')
        zk = ZK(MACHINE_IP, port=PORT, timeout=2, password=0, force_udp=False, ommit_ping=False)
        conn = zk.connect()
        if conn:
            conn.disconnect()
            return True
    except Exception:
        pass
    
    if LAST_DEVICE_SYNC_TIME:
        time_diff = (datetime.now() - LAST_DEVICE_SYNC_TIME).total_seconds()
        if time_diff < 300: # Active in last 5 mins
            return True
            
    return False

def fetch_attendance_data(start_date_str, end_date_str, filter_user_id, selected_store_id=None):
    selected_store_id = safe_store_id(selected_store_id) if selected_store_id else None
    device_online = check_device_connectivity()
    period_data = {}
    raw_punches_list = []
    users_map_temp = {}
    
    for k, v in MASTER_EMPLOYEES.items():
        info = get_emp_info(k)
        code_formatted = f"NWC{k}" if not k.startswith('NWC') else k
        users_map_temp[str(k)] = {'code': code_formatted, 'name': info['name']}
        
    for uid_override, over_data in EMPLOYEE_OVERRIDES.items():
        users_map_temp[str(uid_override)] = {'code': over_data['code'], 'name': over_data['name']}

    attendance_records = []
    
    try:
        zk = ZK(MACHINE_IP, port=PORT, timeout=2, password=0, force_udp=False, ommit_ping=False)
        conn = zk.connect()
        if conn:
            users = conn.get_users()
            for user in users:
                uid_str = str(user.user_id)
                if uid_str in EMPLOYEE_OVERRIDES:
                    emp_name = EMPLOYEE_OVERRIDES[uid_str]['name']
                    emp_code = EMPLOYEE_OVERRIDES[uid_str]['code']
                else:
                    info = get_emp_info(uid_str)
                    emp_name = user.name if user.name else info['name']
                    emp_code = f"NWC{uid_str}" if not uid_str.startswith('NWC') else uid_str
                users_map_temp[uid_str] = {'code': emp_code, 'name': emp_name}
                
            attendance = conn.get_attendance()
            for att in attendance:
                attendance_records.append({
                    'user_id': str(att.user_id),
                    'timestamp': att.timestamp
                })
            conn.disconnect()
    except Exception as e:
        print(f"Direct connection to local device failed (cloud fallback active): {e}")

    if not attendance_records and SYNCED_ATTENDANCE_LOGS:
        for log in SYNCED_ATTENDANCE_LOGS:
            if selected_store_id and safe_store_id(log.get('store_id', DEFAULT_STORE_ID)) != selected_store_id:
                continue
            ts = log['timestamp']
            if isinstance(ts, str):
                try:
                    ts = datetime.strptime(ts, '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    try:
                        ts = datetime.fromisoformat(ts)
                    except Exception:
                        continue
            attendance_records.append({
                'user_id': str(log['user_id']),
                'timestamp': ts
            })

    # Approved regularisation punches become part of the official attendance engine.
    existing_keys={(str(x.get('user_id')),x.get('timestamp')) for x in attendance_records}
    for correction in approved_regularisation_punches():
        key=(str(correction['user_id']),correction['timestamp'])
        if key not in existing_keys:
            attendance_records.append({'user_id':correction['user_id'],'timestamp':correction['timestamp']})
            existing_keys.add(key)

    for att in attendance_records:
        att_ts = att['timestamp']
        att_date_str = att_ts.strftime('%Y-%m-%d')
        if start_date_str <= att_date_str <= end_date_str:
            raw_uid = str(att['user_id'])
            
            if raw_uid in EMPLOYEE_OVERRIDES:
                emp_code = EMPLOYEE_OVERRIDES[raw_uid]['code']
                emp_name = EMPLOYEE_OVERRIDES[raw_uid]['name']
            elif raw_uid in users_map_temp:
                emp_code = users_map_temp[raw_uid]['code']
                emp_name = users_map_temp[raw_uid]['name']
            else:
                clean_uid = raw_uid.replace('NWC', '')
                info = get_emp_info(clean_uid)
                emp_name = info['name']
                emp_code = f"NWC{clean_uid}" if not clean_uid.startswith('NWC') else clean_uid
                
            if filter_user_id and filter_user_id != 'ALL' and emp_code != filter_user_id and raw_uid != filter_user_id:
                continue
            
            raw_punches_list.append({
                'date': att_date_str, 'time': att_ts.strftime('%H:%M:%S'),
                'user_id': emp_code, 'name': emp_name, 'timestamp': att_ts
            })
            
            if att_date_str not in period_data:
                period_data[att_date_str] = {}
            if emp_code not in period_data[att_date_str]:
                period_data[att_date_str][emp_code] = {'name': emp_name, 'timestamps': []}
            period_data[att_date_str][emp_code]['timestamps'].append(att_ts)
            
    users_list = []
    for k, v in sorted(MASTER_EMPLOYEES.items(), key=lambda x: get_emp_info(x[0])['name']):
        if selected_store_id and employee_store_id(k) != selected_store_id:
            continue
        info = get_emp_info(k)
        code_formatted = f"NWC{k}" if not k.startswith('NWC') else k
        users_list.append({'user_id': code_formatted, 'name': info['name'], 'dept': info['dept'], 'off': info['off']})

    final_data = []
    total_duration_seconds = 0
    total_lunch_seconds = 0
    total_net_variance_seconds = 0
    present_count, absent_count, off_count, mis_punch_count, late_arrival_count, ml_count = 0, 0, 0, 0, 0, 0
    shift_a_count, shift_b_count = 0, 0
    
    try:
        range_start = datetime.strptime(start_date_str, '%Y-%m-%d')
        range_end = datetime.strptime(end_date_str, '%Y-%m-%d')
        dates_to_process = []
        cursor = range_end
        while cursor >= range_start:
            dates_to_process.append(cursor.strftime('%Y-%m-%d'))
            cursor -= timedelta(days=1)
    except ValueError:
        dates_to_process = sorted(period_data.keys(), reverse=True)

    for date_str in dates_to_process:
        day_users_dict = period_data.get(date_str, {})
        current_dt = datetime.strptime(date_str, '%Y-%m-%d')
        current_day_name = current_dt.strftime('%A').upper()
        is_weekend = current_dt.weekday() >= 5

        present_records, absent_records, off_records, mispunch_records, ml_records = [], [], [], [], []

        for emp_code, emp_data_val in MASTER_EMPLOYEES.items():
            if selected_store_id and employee_store_id(emp_code) != selected_store_id:
                continue
            final_emp_code = f"NWC{emp_code}" if not emp_code.startswith('NWC') else emp_code
            emp_info = get_emp_info(emp_code)
            emp_name, emp_off, emp_dept, emp_shift = emp_info['name'], emp_info['off'].upper(), emp_info['dept'], emp_info['shift']
            roster_assignment = effective_roster_assignment(final_emp_code, date_str)
            if roster_assignment == 'Shift A':
                emp_shift = 'morning'
            elif roster_assignment == 'Shift B':
                emp_shift = 'second'
            scheduled_off = roster_assignment == 'Off' if roster_assignment else current_day_name == emp_off
            approved_off_change = get_approved_off_change(final_emp_code, date_str)
            if approved_off_change == 'WORKING':
                scheduled_off = False
            elif approved_off_change == 'OFF':
                scheduled_off = True

            if filter_user_id and filter_user_id != 'ALL' and final_emp_code != filter_user_id and emp_code != filter_user_id:
                continue

            approved_leave_obj = next(
                (l for l in LEAVE_REQUESTS if l['user_id'] == final_emp_code and l['status'] == 'Approved' and l['start_date'] <= date_str <= l['end_date']),
                None
            )

            if approved_leave_obj:
                ml_count += 1
                leave_type_code = approved_leave_obj.get('leave_type', 'F10;1')
                ml_records.append({
                    'date': date_str, 'user_id': final_emp_code, 'name': emp_name, 'dept': emp_dept,
                    'store_in': f'Approved Leave ({leave_type_code})', 'lunch_out': '-', 'lunch_in': '-', 'out_time': '-',
                    'total_lunch': '-', 'lunch_seconds': 3600, 'net_duration_seconds': 0, 'total_hours': '-', 'net_variance': '-', 'variance_type': 'neutral',
                    'status': f'{leave_type_code} (Leave)', 'is_late': 'No', 'shift_type': '-', 'scheduled_shift': roster_assignment or '-', 'shift_exception': ''
                })
                continue

            matched_key = emp_code if emp_code in day_users_dict else (final_emp_code if final_emp_code in day_users_dict else None)

            if matched_key:
                present_count += 1
                data_obj = day_users_dict[matched_key]
                sorted_times = sorted(data_obj['timestamps'])
                total_punches = len(sorted_times)
                
                store_in, lunch_out, lunch_in, out_time = '-', '-', '-', '-'
                lunch_seconds, net_duration_seconds = 0, 0
                total_lunch_str, total_hours_str, net_variance_str = '-', '-', '-'
                variance_type, status, is_late = 'neutral', 'Present', False
                shift_type = '-'

                first_punch_time = sorted_times[0].time()
                if first_punch_time <= time(10, 0, 0):
                    shift_type = 'Shift A'
                    shift_a_count += 1
                else:
                    shift_type = 'Shift B'
                    shift_b_count += 1
                
                if total_punches == 1:
                    store_in = sorted_times[0].strftime('%H:%M:%S')
                    status = 'Mis Punch'
                    mis_punch_count += 1
                else:
                    store_in = sorted_times[0].strftime('%H:%M:%S')
                    out_time = sorted_times[-1].strftime('%H:%M:%S')
                    store_in_time = sorted_times[0].time()
                    limit_time = time(13, 10, 0) if emp_shift == 'second' else time(7, 0, 0)
                    if store_in_time > limit_time:
                        late_arrival_count += 1
                        is_late = True

                    if total_punches >= 3:
                        lunch_out = sorted_times[1].strftime('%H:%M:%S')
                        lunch_in = sorted_times[2].strftime('%H:%M:%S')
                        actual_lunch_seconds = (sorted_times[2] - sorted_times[1]).seconds
                        lunch_seconds = 3600 if actual_lunch_seconds < 3600 else actual_lunch_seconds
                    else:
                        lunch_seconds = 3600
                        
                    total_lunch_seconds += lunch_seconds
                    l_hrs = divmod(lunch_seconds, 3600)
                    total_lunch_str = f"{l_hrs[0]}h {l_hrs[1]//60}m"
                    
                    gross_seconds = (sorted_times[-1] - sorted_times[0]).seconds
                    net_duration_seconds = max(0, gross_seconds - lunch_seconds)
                    total_duration_seconds += net_duration_seconds
                    
                    hours = divmod(net_duration_seconds, 3600)
                    total_hours_str = f"{hours[0]}h {hours[1]//60}m"
                    
                    diff_from_target = net_duration_seconds - (7 * 3600)
                    total_net_variance_seconds += diff_from_target
                    
                    if diff_from_target > 0:
                        e_hrs = divmod(diff_from_target, 3600)
                        extra_hours_val = e_hrs[0] + (1 if e_hrs[1] > 0 else 0)
                        code_prefix = 'H07' if is_weekend else 'H06'
                        net_variance_str, variance_type = f"{code_prefix};{extra_hours_val}", 'positive'
                    elif diff_from_target < 0:
                        short_sec = abs(diff_from_target)
                        s_hrs = divmod(short_sec, 3600)
                        net_variance_str, variance_type = f"-{s_hrs[0]}h {s_hrs[1]//60}m", 'negative'
                    else:
                        net_variance_str, variance_type = "0h 0m", 'neutral'
                    
                    status = 'Present'
                
                if scheduled_off:
                    status = 'Weekly Off'

                record = {
                    'date': date_str, 'user_id': final_emp_code, 'name': emp_name, 'dept': emp_dept,
                    'store_in': store_in, 'lunch_out': lunch_out, 'lunch_in': lunch_in,
                    'out_time': out_time, 'total_lunch': total_lunch_str, 'lunch_seconds': lunch_seconds,
                    'net_duration_seconds': net_duration_seconds, 'total_hours': total_hours_str,
                    'net_variance': net_variance_str, 'variance_type': variance_type,
                    'status': status, 'is_late': 'Yes' if is_late else 'No', 'shift_type': shift_type,
                    'scheduled_shift': 'Off' if scheduled_off else ('Shift A' if emp_shift == 'morning' else 'Shift B'),
                    'shift_exception': (
                        'Weekly Off Not Taken' if scheduled_off else
                        ('Unapproved Shift Change' if shift_type in ('Shift A', 'Shift B') and shift_type != ('Shift A' if emp_shift == 'morning' else 'Shift B') else '')
                    )
                }
                
                if status == 'Weekly Off': off_records.append(record)
                elif status == 'Mis Punch': mispunch_records.append(record)
                else: present_records.append(record)
            else:
                if scheduled_off:
                    off_count += 1
                    off_records.append({
                        'date': date_str, 'user_id': final_emp_code, 'name': emp_name, 'dept': emp_dept,
                        'store_in': '-', 'lunch_out': '-', 'lunch_in': '-', 'out_time': '-',
                        'total_lunch': '-', 'lunch_seconds': 3600, 'net_duration_seconds': 0, 'total_hours': '-', 'net_variance': 'Off', 'variance_type': 'neutral',
                        'status': 'Weekly Off', 'is_late': 'No', 'shift_type': '-', 'scheduled_shift': 'Off', 'shift_exception': ''
                    })
                else:
                    absent_count += 1
                    absent_records.append({
                        'date': date_str, 'user_id': final_emp_code, 'name': emp_name, 'dept': emp_dept,
                        'store_in': '-', 'lunch_out': '-', 'lunch_in': '-', 'out_time': '-',
                        'total_lunch': '-', 'lunch_seconds': 3600, 'net_duration_seconds': 0, 'total_hours': '-', 'net_variance': '-', 'variance_type': 'neutral',
                        'status': 'Absent', 'is_late': 'No', 'shift_type': '-', 'scheduled_shift': roster_assignment or ('Shift A' if emp_shift == 'morning' else 'Shift B'), 'shift_exception': ''
                    })
        
        final_data.extend(present_records + mispunch_records + off_records + absent_records + ml_records)
            
    tot_hrs = divmod(total_duration_seconds, 3600)
    grand_total_hours = f"{tot_hrs[0]}h {tot_hrs[1]//60}m"
    tot_l_hrs = divmod(total_lunch_seconds, 3600)
    grand_total_lunch_hours = f"{tot_l_hrs[0]}h {tot_l_hrs[1]//60}m"

    v_sec = total_net_variance_seconds
    if v_sec >= 0:
        v_hrs = divmod(v_sec, 3600)
        grand_total_variance, grand_variance_type = f"+{v_hrs[0]}h {v_hrs[1]//60}m", 'positive'
    else:
        v_hrs = divmod(abs(v_sec), 3600)
        grand_total_variance, grand_variance_type = f"-{v_hrs[0]}h {v_hrs[1]//60}m", 'negative'
    
    stats_summary = {
        'present': present_count, 'absent': absent_count, 'off': off_count, 'mispunch': mis_punch_count,
        'late_arrival': late_arrival_count, 'ml': ml_count,
        'shift_a': shift_a_count, 'shift_b': shift_b_count,
        'total_hrs': grand_total_hours, 'total_lunch_hrs': grand_total_lunch_hours,
        'total_variance': grand_total_variance, 'variance_type': grand_variance_type, 'device_online': device_online
    }
    
    raw_punches_list = sorted(raw_punches_list, key=lambda x: x['timestamp'], reverse=True)
    return final_data, users_list, grand_total_hours, grand_total_lunch_hours, grand_total_variance, raw_punches_list, stats_summary


def _load_json_list(path):
    try:
        if os.path.exists(path):
            with open(path,'r',encoding='utf-8') as f: return json.load(f)
    except Exception as exc: print(f'Unable to load {path}: {exc}')
    return []
def _save_json_list(path,data):
    with open(path,'w',encoding='utf-8') as f: json.dump(data,f,indent=2,ensure_ascii=False)
EXAMS=_load_json_list(EXAMS_JSON_FILE)
EXAM_SUBMISSIONS=_load_json_list(EXAM_SUBMISSIONS_JSON_FILE)

EXAM_TEMPLATE="""
<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Projects, Tasks & Exams</title><script src="https://cdn.tailwindcss.com"></script>
<style>
:root{--ink:#07111f;--mint:#10b981;--cyan:#22d3ee;--violet:#7c3aed}*{box-sizing:border-box}body{background:radial-gradient(circle at 6% 2%,rgba(16,185,129,.18),transparent 23%),radial-gradient(circle at 94% 8%,rgba(59,130,246,.18),transparent 28%),linear-gradient(145deg,#f8fafc,#edf3ff);min-height:100vh}.gridbg{position:fixed;inset:0;z-index:-1;background-image:linear-gradient(rgba(15,23,42,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(15,23,42,.025) 1px,transparent 1px);background-size:38px 38px}.glass{background:rgba(255,255,255,.76);border:1px solid rgba(255,255,255,.9);backdrop-filter:blur(22px);box-shadow:0 24px 70px rgba(15,23,42,.10)}.hero{background:linear-gradient(125deg,#06101f 0%,#101d33 55%,#0b2440 100%);position:relative;overflow:hidden}.hero:after{content:'';position:absolute;width:320px;height:320px;border-radius:50%;right:-80px;top:-140px;background:radial-gradient(circle,rgba(34,211,238,.28),transparent 68%)}.field{background:rgba(255,255,255,.9);border:1px solid #dbe4ef;transition:.2s}.field:focus{outline:none;border-color:#22d3ee;box-shadow:0 0 0 4px rgba(34,211,238,.13);transform:translateY(-1px)}.question-card{transition:.25s}.question-card:hover{transform:translateY(-2px);box-shadow:0 16px 38px rgba(15,23,42,.09)}.exam-card{transition:.25s}.exam-card:hover{transform:translateY(-5px);box-shadow:0 24px 55px rgba(15,23,42,.12)}button,a{transition:.2s}button:hover,a:hover{transform:translateY(-2px)}.step{display:inline-flex;height:28px;width:28px;align-items:center;justify-content:center;border-radius:999px;background:#07111f;color:white;font-size:11px;font-weight:900}@keyframes enter{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}main>*{animation:enter .45s ease both}@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style></head>
<body class="text-slate-800"><div class="gridbg"></div><main class="mx-auto max-w-[1800px] p-4 md:p-7 space-y-5">
<header class="hero rounded-[32px] p-6 md:p-8 text-white shadow-2xl"><div class="relative z-10 flex flex-wrap items-end justify-between gap-5"><div><div class="inline-flex items-center gap-2 rounded-full border border-emerald-300/20 bg-emerald-300/10 px-3 py-1.5 text-[10px] font-black tracking-[.18em] text-emerald-300">◆ PROJECTS & TASKS</div><h1 class="mt-4 text-3xl md:text-5xl font-black tracking-tight">Employee Exam Studio</h1><p class="mt-2 max-w-3xl text-sm leading-6 text-slate-300">Create professional single or multiple-choice assessments, schedule precise exam windows, assign employees and review answer-level analytics.</p></div><div class="grid grid-cols-3 gap-2 text-center"><div class="rounded-2xl border border-white/10 bg-white/5 px-4 py-3"><b class="block text-xl text-emerald-300">{{ exams|length }}</b><span class="text-[9px] uppercase text-slate-400">Exams</span></div><div class="rounded-2xl border border-white/10 bg-white/5 px-4 py-3"><b class="block text-xl text-cyan-300">AI</b><span class="text-[9px] uppercase text-slate-400">Scoring</span></div><div class="rounded-2xl border border-white/10 bg-white/5 px-4 py-3"><b class="block text-xl text-violet-300">360°</b><span class="text-[9px] uppercase text-slate-400">Reports</span></div></div></div></header>
{% with messages=get_flashed_messages(with_categories=true) %}{% for c,m in messages %}<div class="glass rounded-2xl border-l-4 border-l-emerald-500 p-4 font-bold text-emerald-800">{{m}}</div>{% endfor %}{% endwith %}
{% if role in ['admin','developer'] %}<section class="glass rounded-[30px] p-5 md:p-7"><div class="flex flex-wrap items-center justify-between gap-3"><div><p class="text-[10px] font-black uppercase tracking-wider text-indigo-600">Assessment Builder</p><h2 class="mt-1 text-2xl font-black">Create Question Paper</h2></div><span class="rounded-full bg-slate-100 px-3 py-1.5 text-xs font-bold text-slate-500">Auto-save after schedule</span></div>
<form method="post" action="/projects-tasks/exam/create" id="examForm" class="mt-6 space-y-5"><div class="grid lg:grid-cols-4 gap-3"><label class="block"><span class="mb-2 flex items-center gap-2 text-[10px] font-black uppercase text-slate-500"><span class="step">1</span> Exam title</span><input name="title" required placeholder="Example: Cash SOP Certification" class="field w-full rounded-2xl p-4 font-bold"></label><label class="block"><span class="mb-2 flex items-center gap-2 text-[10px] font-black uppercase text-slate-500"><span class="step">2</span> Opens at</span><input type="datetime-local" name="start_at" required class="field w-full rounded-2xl p-4 font-bold"></label><label class="block"><span class="mb-2 flex items-center gap-2 text-[10px] font-black uppercase text-slate-500"><span class="step">3</span> Closes at</span><input type="datetime-local" name="end_at" required class="field w-full rounded-2xl p-4 font-bold"></label><div><span class="mb-2 flex items-center gap-2 text-[10px] font-black uppercase text-slate-500"><span class="step">4</span> Assignment</span><select id="assignmentMode" name="assignment_mode" class="field mb-2 w-full rounded-2xl p-3 font-black"><option value="multiple">Assign Multiple Employees</option><option value="single">Assign Single Employee</option></select><select id="assigneeSelect" name="assignees" multiple required class="field h-28 w-full rounded-2xl p-3 text-sm">{% for e in employees %}<option value="{{e.user_id}}">{{e.name}} · {{e.dept}}</option>{% endfor %}</select><p id="assignHelp" class="mt-1 text-[10px] font-bold text-slate-400">Select one or more employees.</p></div></div>
<div class="flex items-center justify-between"><div><p class="text-[10px] font-black uppercase text-indigo-600">Question Section</p><h3 class="text-lg font-black">Build the assessment</h3></div><button type="button" onclick="addQuestion()" class="rounded-xl border border-indigo-200 bg-indigo-50 px-4 py-2 text-sm font-black text-indigo-700">＋ Add Question</button></div><div id="questions" class="space-y-4"></div><div class="sticky bottom-3 z-20 flex flex-wrap justify-between gap-3 rounded-2xl border border-white/80 bg-slate-950/92 p-3 text-white shadow-2xl backdrop-blur-xl"><div class="px-2 py-2 text-xs text-slate-300">Review question types, correct options and marks before scheduling.</div><button class="rounded-xl bg-gradient-to-r from-emerald-500 to-cyan-500 px-6 py-3 font-black text-slate-950 shadow-lg">Schedule Exam →</button></div></form></section>{% endif %}
<section><div class="mb-3 flex items-center justify-between"><div><p class="text-[10px] font-black uppercase text-indigo-600">Assessment Library</p><h2 class="text-2xl font-black">Scheduled Exams</h2></div></div><div class="grid xl:grid-cols-2 gap-4">{% for exam in exams %}<article class="exam-card glass rounded-[28px] p-5"><div class="flex justify-between gap-4"><div><div class="mb-2 inline-flex rounded-full bg-indigo-50 px-2.5 py-1 text-[9px] font-black text-indigo-700">{{exam.questions|length}} QUESTIONS · {{exam.assignees|length}} EMPLOYEE(S)</div><h3 class="text-xl font-black">{{exam.title}}</h3><p class="mt-1 text-xs text-slate-500">{{exam.start_at}} → {{exam.end_at}}</p></div><span class="h-fit rounded-full px-3 py-1.5 text-[10px] font-black {% if exam.status=='Open' %}bg-emerald-100 text-emerald-700{% elif exam.status=='Upcoming' %}bg-blue-100 text-blue-700{% else %}bg-slate-200 text-slate-600{% endif %}">{{exam.status}}</span></div><div class="mt-5 flex flex-wrap gap-2">{% if role=='employee' and exam.status=='Open' and not exam.submitted %}<a href="/projects-tasks/exam/{{exam.id}}" class="rounded-xl bg-indigo-600 px-4 py-2.5 text-sm font-black text-white">Open Exam</a>{% endif %}{% if role in ['admin','developer'] %}<a href="/projects-tasks/exam/{{exam.id}}/report" class="rounded-xl bg-slate-950 px-4 py-2.5 text-sm font-black text-white">View Analytics</a><a href="/projects-tasks/exam/{{exam.id}}/report.xlsx" class="rounded-xl bg-emerald-600 px-4 py-2.5 text-sm font-black text-white">Download Excel</a>{% endif %}</div></article>{% else %}<div class="glass col-span-full rounded-3xl p-10 text-center"><p class="text-4xl">◫</p><p class="mt-3 font-black">No exams created yet</p><p class="text-sm text-slate-400">Create the first professional assessment above.</p></div>{% endfor %}</div></section>
</main><script>
let q=0;function addQuestion(){q++;document.getElementById('questions').insertAdjacentHTML('beforeend',`<div class="question-card rounded-[24px] border border-slate-200 bg-slate-50/85 p-4 md:p-5"><div class="flex items-center justify-between"><div class="flex items-center gap-3"><span class="step">${q}</span><div><p class="text-[9px] font-black uppercase text-indigo-600">Question ${q}</p><b class="text-sm">Configure prompt and scoring</b></div></div><button type="button" onclick="this.closest('.question-card').remove()" class="rounded-lg border border-rose-200 bg-rose-50 px-3 py-1.5 text-xs font-black text-rose-600">Remove</button></div><div class="mt-4 grid md:grid-cols-[1fr_220px] gap-3"><input name="q_text" required placeholder="Write the question clearly" class="field w-full rounded-xl p-3 font-bold"><select name="q_${q}_type" onchange="updateCorrectType(${q},this.value)" class="field rounded-xl p-3 font-black"><option value="single">Single Selection</option><option value="multiple">Multiple Selection</option></select></div><div class="grid md:grid-cols-4 gap-2 mt-3">${[0,1,2,3].map(i=>`<input name="q_${q}_option" required placeholder="Option ${i+1}" class="field rounded-xl p-3">`).join('')}</div><div class="mt-3 flex flex-wrap items-center gap-2"><span class="text-[10px] font-black uppercase text-slate-500">Correct answer:</span><div id="correct_${q}" class="flex flex-wrap gap-2">${[0,1,2,3].map((i)=>`<label class="cursor-pointer rounded-xl border bg-white px-3 py-2 text-xs font-bold"><input type="radio" name="q_${q}_correct" value="${i}" ${i===0?'required':''}> Option ${i+1}</label>`).join('')}</div><label class="ml-auto flex items-center gap-2 text-xs font-black text-slate-500">Marks <input type="number" min="1" value="1" name="q_${q}_marks" class="field w-20 rounded-xl p-2 text-center"></label></div><input type="hidden" name="q_index" value="${q}"></div>`)}
function updateCorrectType(i,type){document.querySelectorAll(`#correct_${i} input`).forEach((x,n)=>{x.type=type==='multiple'?'checkbox':'radio';x.required=type==='single'&&n===0})}
const mode=document.getElementById('assignmentMode'),assignees=document.getElementById('assigneeSelect');if(mode){mode.addEventListener('change',()=>{assignees.multiple=mode.value==='multiple';if(mode.value==='single'){let kept=false;[...assignees.options].forEach(o=>{if(o.selected&&!kept)kept=true;else o.selected=false})}document.getElementById('assignHelp').textContent=mode.value==='multiple'?'Select one or more employees.':'Select exactly one employee.'})}addQuestion();
</script></body></html>
"""
EXAM_TAKE_TEMPLATE="""<!doctype html><html><head><meta charset="utf-8"><script src="https://cdn.tailwindcss.com"></script><title>{{exam.title}}</title></head><body class="bg-slate-50"><main class="mx-auto max-w-3xl p-5"><header class="rounded-3xl bg-indigo-700 p-6 text-white"><h1 class="text-3xl font-black">{{exam.title}}</h1><p>Submit before {{exam.end_at}}</p></header><form method="post" class="mt-4 space-y-4">{% for q in exam.questions %}{% set qi=loop.index0 %}<section class="rounded-2xl bg-white p-5 shadow"><p class="font-black">{{loop.index}}. {{q.text}} <span class="text-slate-400">({{q.marks}} mark)</span></p><p class="text-[10px] font-bold text-indigo-600">{% if q.type=='multiple' %}Select all correct choices{% else %}Select one choice{% endif %}</p>{% for opt in q.options %}<label class="mt-2 flex gap-2 rounded-xl border p-3"><input type="{% if q.type=='multiple' %}checkbox{% else %}radio{% endif %}" name="answer_{{qi}}" value="{{loop.index0}}" {% if q.type!='multiple' %}required{% endif %}>{{opt}}</label>{% endfor %}</section>{% endfor %}<button class="w-full rounded-xl bg-emerald-600 p-4 font-black text-white">Submit Exam</button></form></main></body></html>"""
EXAM_REPORT_TEMPLATE="""<!doctype html><html><head><meta charset="utf-8"><script src="https://cdn.tailwindcss.com"></script><title>Exam Report</title></head><body class="bg-slate-50"><main class="p-5"><div class="flex flex-wrap justify-between gap-3"><h1 class="text-3xl font-black">{{exam.title}} · Submission Report</h1><a href="/projects-tasks/exam/{{exam.id}}/report.xlsx" class="rounded-xl bg-emerald-600 px-5 py-3 font-black text-white">Download Detailed Excel</a></div><div class="mt-4 grid grid-cols-4 gap-3">{% for l,v in cards %}<div class="rounded-2xl bg-white p-4 shadow"><p class="text-xs text-slate-500">{{l}}</p><p class="text-2xl font-black">{{v}}</p></div>{% endfor %}</div><div class="mt-4 overflow-auto rounded-3xl bg-white shadow"><table class="w-full text-sm"><thead class="bg-slate-950 text-white"><tr><th class="p-3 text-left">Employee</th><th>Submitted</th><th>Score</th><th>Percentage</th><th>Result</th><th>Answers</th></tr></thead><tbody>{% for r in rows %}<tr class="border-b"><td class="p-3"><b>{{r.name}}</b><br><small>{{r.user_id}} · {{r.dept}}</small></td><td class="text-center">{{r.submitted_at}}</td><td class="text-center">{{r.score}} / {{r.total}}</td><td class="text-center font-black">{{r.percentage}}%</td><td class="text-center">{{r.result}}</td><td class="text-center">{% if r.submitted %}<button onclick="openAnswers('{{r.user_id}}')" class="rounded-lg bg-indigo-600 px-3 py-2 text-xs font-black text-white">View Selected Answers</button>{% else %}-{% endif %}</td></tr>{% endfor %}</tbody></table></div></main><div id="answerModal" class="fixed inset-0 hidden items-center justify-center bg-slate-950/60 p-4"><div class="max-h-[85vh] w-full max-w-5xl overflow-hidden rounded-3xl bg-white"><header class="flex justify-between border-b p-5"><div><p id="answerEmployee" class="text-xs font-bold text-indigo-700"></p><h2 class="text-xl font-black">Question-wise Selected Answers</h2></div><button onclick="closeAnswers()" class="rounded-xl bg-slate-950 px-4 py-2 text-white">Close</button></header><div id="answerBody" class="max-h-[70vh] overflow-auto p-5 space-y-3"></div></div></div><script>const answerDetails={{answer_details|tojson}};function openAnswers(uid){const d=answerDetails[uid];document.getElementById('answerEmployee').textContent=d.employee;document.getElementById('answerBody').innerHTML=d.answers.map((a,i)=>`<div class="rounded-2xl border p-4 ${a.correct?'bg-emerald-50 border-emerald-200':'bg-rose-50 border-rose-200'}"><p class="font-black">${i+1}. ${a.question}</p><p class="mt-2 text-sm"><b>Selected:</b> ${a.selected.join(', ')||'No answer'}</p><p class="text-sm"><b>Correct:</b> ${a.correct_answers.join(', ')}</p><p class="mt-1 font-black">${a.correct?'Correct':'Incorrect'} · ${a.awarded}/${a.marks} marks</p></div>`).join('');const m=document.getElementById('answerModal');m.classList.remove('hidden');m.classList.add('flex')}function closeAnswers(){const m=document.getElementById('answerModal');m.classList.add('hidden');m.classList.remove('flex')}</script></body></html>"""

def _exam_status(exam):
    now=datetime.now(); start=datetime.fromisoformat(exam['start_at']); end=datetime.fromisoformat(exam['end_at'])
    return 'Upcoming' if now<start else ('Closed' if now>end else 'Open')
def _correct_indexes(question):
    value=question.get('correct',[])
    return sorted([int(x) for x in (value if isinstance(value,list) else [value])])
def _submission_detail(exam,submission):
    answers=[]
    for i,q in enumerate(exam['questions']):
        raw=(submission.get('answers') or [])[i] if i<len(submission.get('answers') or []) else []
        selected=sorted([int(x) for x in (raw if isinstance(raw,list) else [raw])])
        correct=_correct_indexes(q); ok=selected==correct
        answers.append({'question':q['text'],'selected':[q['options'][x] for x in selected if x<len(q['options'])],'correct_answers':[q['options'][x] for x in correct if x<len(q['options'])],'correct':ok,'marks':q['marks'],'awarded':q['marks'] if ok else 0})
    return answers
@app.route('/projects-tasks')
def projects_tasks():
    if not session.get('logged_in'): return redirect(url_for('login'))
    role=session.get('role'); uid=session.get('user_id'); employees=[{'user_id':normalized_code(c),'name':get_emp_info(c)['name'],'dept':get_emp_info(c)['dept']} for c in MASTER_EMPLOYEES]
    visible=[]
    for x in EXAMS:
        if role in ['admin','developer'] or uid in x['assignees']:
            y=dict(x); y['status']=_exam_status(x); y['submitted']=any(r['exam_id']==x['id'] and r['user_id']==uid for r in EXAM_SUBMISSIONS); visible.append(y)
    return render_template_string(EXAM_TEMPLATE,role=role,employees=employees,exams=visible)
@app.route('/projects-tasks/exam/create',methods=['POST'])
def create_exam():
    if session.get('role') not in ['admin','developer']: return 'Forbidden',403
    indices=request.form.getlist('q_index'); texts=request.form.getlist('q_text'); questions=[]
    for pos,i in enumerate(indices):
        qtype=request.form.get(f'q_{i}_type','single'); correct=[int(x) for x in request.form.getlist(f'q_{i}_correct')]
        if not correct: return 'Every question needs at least one correct answer.',400
        questions.append({'text':texts[pos],'type':qtype,'options':request.form.getlist(f'q_{i}_option'),'correct':correct,'marks':int(request.form[f'q_{i}_marks'])})
    assignees=request.form.getlist('assignees')
    if request.form.get('assignment_mode')=='single': assignees=assignees[:1]
    if not assignees:return 'Select at least one employee.',400
    EXAMS.append({'id':uuid.uuid4().hex[:10],'title':request.form['title'],'start_at':request.form['start_at'],'end_at':request.form['end_at'],'assignees':assignees,'assignment_mode':request.form.get('assignment_mode','multiple'),'questions':questions,'created_at':datetime.now().isoformat()});_save_json_list(EXAMS_JSON_FILE,EXAMS);flash('Exam scheduled successfully.','success');return redirect(url_for('projects_tasks'))
@app.route('/projects-tasks/exam/<exam_id>',methods=['GET','POST'])
def take_exam(exam_id):
    exam=next((x for x in EXAMS if x['id']==exam_id),None);uid=session.get('user_id')
    if not exam or uid not in exam['assignees'] or _exam_status(exam)!='Open': return 'Exam unavailable',403
    if any(r['exam_id']==exam_id and r['user_id']==uid for r in EXAM_SUBMISSIONS): return 'Already submitted',409
    if request.method=='POST':
        score=0;total=sum(q['marks'] for q in exam['questions']);answers=[]
        for i,q in enumerate(exam['questions']):
            selected=sorted([int(x) for x in request.form.getlist(f'answer_{i}')]);answers.append(selected)
            if selected==_correct_indexes(q):score+=q['marks']
        pct=round(score*100/total,2) if total else 0;EXAM_SUBMISSIONS.append({'exam_id':exam_id,'user_id':uid,'submitted_at':datetime.now().isoformat(timespec='seconds'),'answers':answers,'score':score,'total':total,'percentage':pct});_save_json_list(EXAM_SUBMISSIONS_JSON_FILE,EXAM_SUBMISSIONS);flash(f'Exam submitted. Score: {pct}%','success');return redirect(url_for('projects_tasks'))
    normalized=dict(exam);normalized['questions']=[dict(q,type=q.get('type','single')) for q in exam['questions']]
    return render_template_string(EXAM_TAKE_TEMPLATE,exam=normalized)
def _exam_report_rows(exam):
    rows=[];details={}
    for uid in exam['assignees']:
        info=get_emp_info(uid);sub=next((r for r in EXAM_SUBMISSIONS if r['exam_id']==exam['id'] and r['user_id']==uid),None);pct=sub['percentage'] if sub else 0
        rows.append({'user_id':uid,'name':info['name'],'dept':info['dept'],'submitted':bool(sub),'submitted_at':sub['submitted_at'] if sub else 'Pending','score':sub['score'] if sub else 0,'total':sub['total'] if sub else sum(q['marks'] for q in exam['questions']),'percentage':pct,'result':'Pass' if sub and pct>=60 else ('Fail' if sub else 'Pending')})
        if sub:details[uid]={'employee':f"{info['name']} · {uid} · {info['dept']}",'answers':_submission_detail(exam,sub)}
    return rows,details
@app.route('/projects-tasks/exam/<exam_id>/report')
def exam_report(exam_id):
    if session.get('role') not in ['admin','developer']: return 'Forbidden',403
    exam=next((x for x in EXAMS if x['id']==exam_id),None)
    if not exam:return 'Not found',404
    rows,details=_exam_report_rows(exam);submitted=sum(1 for r in rows if r['submitted']);avg=round(sum(r['percentage'] for r in rows if r['submitted'])/submitted,2) if submitted else 0
    return render_template_string(EXAM_REPORT_TEMPLATE,exam=exam,rows=rows,answer_details=details,cards=[('Assigned',len(rows)),('Submitted',submitted),('Pending',len(rows)-submitted),('Average',f'{avg}%')])
@app.route('/projects-tasks/exam/<exam_id>/report.xlsx')
def exam_report_excel(exam_id):
    if session.get('role') not in ['admin','developer']: return 'Forbidden',403
    exam=next((x for x in EXAMS if x['id']==exam_id),None)
    if not exam:return 'Not found',404
    rows,details=_exam_report_rows(exam);wb=openpyxl.Workbook();summary=wb.active;summary.title='Summary';summary.append(['Employee ID','Employee Name','Designation','Submitted At','Score','Total','Percentage','Result'])
    for r in rows:summary.append([r['user_id'],r['name'],r['dept'],r['submitted_at'],r['score'],r['total'],r['percentage'],r['result']])
    answers=wb.create_sheet('Selected Answers');answers.append(['Employee ID','Employee Name','Question No.','Question','Selection Type','Selected Answer(s)','Correct Answer(s)','Status','Marks Awarded','Total Marks'])
    for r in rows:
        if not r['submitted']:continue
        for i,a in enumerate(details[r['user_id']]['answers']):
            q=exam['questions'][i];answers.append([r['user_id'],r['name'],i+1,a['question'],q.get('type','single'),', '.join(a['selected']),', '.join(a['correct_answers']),'Correct' if a['correct'] else 'Incorrect',a['awarded'],a['marks']])
    for ws in wb.worksheets:
        ws.freeze_panes='A2'
        for c in ws[1]:c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='0F172A')
        for col in ws.columns:ws.column_dimensions[col[0].column_letter].width=min(55,max(12,max(len(str(c.value or '')) for c in col)+2))
    output=io.BytesIO();wb.save(output);output.seek(0);name=secure_filename(exam['title']) or 'exam'
    return send_file(output,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name=f'{name}_detailed_report.xlsx')

# ===== MULTI STORE MANAGEMENT V1 =====
def _load_store_registry():
    default = {
        DEFAULT_STORE_ID: {
            'store_id': DEFAULT_STORE_ID,
            'store_name': 'Fresmart Express LM11',
            'location': 'Luanda',
            'machine_ip': MACHINE_IP,
            'machine_port': PORT,
            'bridge_token': secrets.token_urlsafe(32),
            'active': True,
            'created_at': datetime.now().isoformat(timespec='seconds')
        }
    }
    try:
        if os.path.isfile(STORES_JSON_FILE):
            with open(STORES_JSON_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict) and data:
                    return data
    except Exception as exc:
        print(f'Store registry load error: {exc}')
    _atomic_save_json(STORES_JSON_FILE, default)
    return default

STORES = _load_store_registry()

def save_stores():
    _atomic_save_json(STORES_JSON_FILE, STORES)

def safe_store_id(value):
    clean = ''.join(ch for ch in str(value or '').strip().upper() if ch.isalnum() or ch in ('-', '_'))
    return clean[:24]

def employee_store_id(employee_code):
    return str(get_emp_info(employee_code).get('store_id') or DEFAULT_STORE_ID).upper()

def store_employees(store_id):
    sid = safe_store_id(store_id)
    return {normalized_code(code) for code in MASTER_EMPLOYEES if employee_store_id(code) == sid}

def detect_tailscale_portal_url():
    """Prefer explicit URL, otherwise auto-detect this PC's Tailscale IPv4."""
    if PORTAL_PUBLIC_URL:
        return PORTAL_PUBLIC_URL, 'environment'
    commands = [
        [r'C:\Program Files\Tailscale\tailscale.exe', 'ip', '-4'],
        ['tailscale', 'ip', '-4']
    ]
    for command in commands:
        try:
            result = subprocess.run(command, capture_output=True, text=True, timeout=5, check=False)
            ip = (result.stdout or '').strip().splitlines()[0].strip()
            if ip.startswith('100.'):
                return f'http://{ip}:{PORTAL_PORT}', 'tailscale-auto'
        except Exception:
            pass
    return request.url_root.rstrip('/'), 'request-fallback'

STORE_MANAGER_TEMPLATE = r"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><script src="https://cdn.tailwindcss.com"></script><title>Multi Store & Bridge</title></head><body class="min-h-screen bg-slate-100 text-slate-800"><main class="mx-auto max-w-7xl p-4 md:p-7 space-y-5"><header class="rounded-3xl bg-slate-950 p-6 text-white flex flex-wrap justify-between gap-4"><div><p class="text-xs font-black tracking-widest text-cyan-300">DEVELOPER CONTROL</p><h1 class="text-3xl font-black">Multi Store & Bridge Manager</h1><p class="mt-2 text-sm text-slate-400">Create stores and download a dedicated biometric bridge for each location.</p></div><a href="/" class="h-fit rounded-xl bg-white/10 px-4 py-2 font-bold">Dashboard</a></header>{% for c,m in get_flashed_messages(with_categories=true) %}<div class="rounded-xl p-4 font-bold {{'bg-emerald-50 text-emerald-700' if c=='success' else 'bg-rose-50 text-rose-700'}}">{{m}}</div>{% endfor %}<section class="grid lg:grid-cols-3 gap-5"><form method="post" class="rounded-3xl bg-white p-6 shadow space-y-4"><h2 class="text-xl font-black">Create New Store</h2><label class="block text-xs font-black">STORE CODE<input name="store_id" required placeholder="LM12" class="mt-1 w-full rounded-xl border p-3 uppercase"></label><label class="block text-xs font-black">STORE NAME<input name="store_name" required placeholder="Fresmart Express LM12" class="mt-1 w-full rounded-xl border p-3"></label><label class="block text-xs font-black">LOCATION<input name="location" placeholder="Luanda" class="mt-1 w-full rounded-xl border p-3"></label><div class="grid grid-cols-2 gap-2"><label class="text-xs font-black">MACHINE IP<input name="machine_ip" value="192.168.1.201" required class="mt-1 w-full rounded-xl border p-3"></label><label class="text-xs font-black">PORT<input name="machine_port" type="number" value="4370" min="1" max="65535" required class="mt-1 w-full rounded-xl border p-3"></label></div><button class="w-full rounded-xl bg-gradient-to-r from-emerald-600 to-cyan-600 p-3 font-black text-white">Create Store</button></form><section class="lg:col-span-2 rounded-3xl bg-white p-6 shadow"><h2 class="text-xl font-black mb-4">Created Stores</h2><div class="overflow-auto"><table class="w-full min-w-[780px] text-sm"><thead><tr class="bg-slate-900 text-white"><th class="p-3 text-left">Store</th><th>Location</th><th>Device</th><th>Employees</th><th>Bridge</th></tr></thead><tbody>{% for sid,x in stores %}<tr class="border-b"><td class="p-3"><b>{{x.store_name}}</b><br><span class="font-mono text-xs">{{sid}}</span></td><td class="text-center">{{x.location or '-'}}</td><td class="text-center font-mono">{{x.machine_ip}}:{{x.machine_port}}</td><td class="text-center">{{x.employee_count}}</td><td class="text-center"><a href="/developer/stores/{{sid}}/bridge.zip" class="inline-flex rounded-xl bg-indigo-600 px-4 py-2 font-black text-white">Download Bridge ZIP</a></td></tr>{% endfor %}</tbody></table></div></section></section></main></body></html>"""

@app.route('/developer/stores', methods=['GET', 'POST'])
def developer_stores():
    if not session.get('logged_in') or session.get('role') != 'developer':
        return 'Developer access required', 403
    if request.method == 'POST':
        sid = safe_store_id(request.form.get('store_id'))
        name = ' '.join(request.form.get('store_name', '').strip().split())
        if not sid or not name:
            flash('Store code aur store name required hain.', 'error')
            return redirect(url_for('developer_stores'))
        if sid in STORES:
            flash(f'{sid} already exists.', 'error')
            return redirect(url_for('developer_stores'))
        try:
            machine_port = int(request.form.get('machine_port', '4370'))
            if not 1 <= machine_port <= 65535:
                raise ValueError()
        except ValueError:
            flash('Machine port invalid hai.', 'error')
            return redirect(url_for('developer_stores'))
        STORES[sid] = {
            'store_id': sid,
            'store_name': name,
            'location': request.form.get('location', '').strip(),
            'machine_ip': request.form.get('machine_ip', '').strip(),
            'machine_port': machine_port,
            'bridge_token': secrets.token_urlsafe(32),
            'active': True,
            'created_at': datetime.now().isoformat(timespec='seconds'),
            'created_by': session.get('user_id')
        }
        save_stores()
        flash(f'{sid} store create ho gaya. Bridge ZIP isi row se download karein.', 'success')
        return redirect(url_for('developer_stores'))
    rows = []
    for sid, record in sorted(STORES.items()):
        item = dict(record)
        item['employee_count'] = len(store_employees(sid))
        rows.append((sid, item))
    return render_template_string(STORE_MANAGER_TEMPLATE, stores=rows)

@app.route('/developer/stores/<store_id>/bridge.zip')
def download_store_bridge(store_id):
    if not session.get('logged_in') or session.get('role') != 'developer':
        return redirect(url_for('login'))
    sid = safe_store_id(store_id)
    store = STORES.get(sid)
    if not store:
        return 'Store not found', 404
    portal_url, portal_url_source = detect_tailscale_portal_url()
    config = {
        'store_id': sid,
        'store_name': store['store_name'],
        'portal_url': portal_url,
        'portal_url_source': portal_url_source,
        'bridge_token': store['bridge_token'],
        'machine_ip': store['machine_ip'],
        'machine_port': store['machine_port'],
        'sync_seconds': 300,
        'generated_at': datetime.now().isoformat(timespec='seconds')
    }
    bridge_source = """import json, time, requests
from datetime import datetime
from pathlib import Path
from zk import ZK
CFG = json.loads(Path('bridge_config.json').read_text(encoding='utf-8'))
def sync_once():
    zk = ZK(CFG['machine_ip'], port=int(CFG['machine_port']), timeout=10, password=0, force_udp=False, ommit_ping=False)
    conn = zk.connect()
    try:
        logs = [{'user_id': str(x.user_id), 'timestamp': x.timestamp.strftime('%Y-%m-%d %H:%M:%S')} for x in (conn.get_attendance() or [])]
    finally:
        conn.disconnect()
    response = requests.post(CFG['portal_url'] + '/api/attendance/sync', json={'store_id': CFG['store_id'], 'bridge_token': CFG['bridge_token'], 'logs': logs}, timeout=60)
    response.raise_for_status()
    print(datetime.now(), response.json())
while True:
    try:
        sync_once()
    except Exception as exc:
        print(datetime.now(), 'sync failed:', exc)
    time.sleep(max(30, int(CFG.get('sync_seconds', 300))))
"""
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('bridge_config.json', json.dumps(config, indent=2))
        zf.writestr('store_bridge.py', bridge_source)
        zf.writestr('requirements.txt', 'requests>=2.31.0\npyzk>=0.9\n')
        zf.writestr('run_bridge.bat', '@echo off\ncd /d %~dp0\npython -m pip install -r requirements.txt\npython test_connection.py\nif errorlevel 1 pause & exit /b 1\npython store_bridge.py\npause\n')
        zf.writestr('test_connection.py', """import json, socket, sys, requests
from pathlib import Path
c=json.loads(Path('bridge_config.json').read_text(encoding='utf-8'))
ok=True
try:
 socket.create_connection((c['machine_ip'],int(c['machine_port'])),5).close();print('[OK] Biometric reachable')
except Exception as e: print('[FAIL] Biometric:',e);ok=False
try:
 r=requests.get(c['portal_url']+'/login',timeout=15);print('[OK] Portal reachable:',r.status_code)
except Exception as e: print('[FAIL] Portal:',e);ok=False
sys.exit(0 if ok else 1)
""")
        zf.writestr('README.txt', f"Store: {sid}\nDevice: {store['machine_ip']}:{store['machine_port']}\nPortal: {portal_url}\nURL source: {portal_url_source}\n\n1. Install Tailscale and join the same tailnet as the portal PC.\n2. Enable Preferences > Run unattended.\n3. Extract this ZIP and run run_bridge.bat.\n4. Keep this bridge PC and biometric machine on the same local network.\n")
    output.seek(0)
    return send_file(output, mimetype='application/zip', as_attachment=True, download_name=f'{sid}_biometric_bridge.zip')
# ===== END MULTI STORE MANAGEMENT V1 =====

LOGIN_TEMPLATE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Secure Access | Gamek AI Workforce</title>
<style id="gamek-neo-ui">
:root{--neo:#07111f;--neo2:#0d1b2e;--cyan:#22d3ee;--mint:#34d399;--violet:#8b5cf6;--line:rgba(148,163,184,.18);--glass:rgba(255,255,255,.72)}
html{scroll-behavior:smooth}body{background:radial-gradient(circle at 8% 4%,rgba(52,211,153,.20),transparent 25%),radial-gradient(circle at 92% 5%,rgba(59,130,246,.20),transparent 28%),linear-gradient(145deg,#f8fafc,#eef4ff)!important;color:#142033}
body:before{content:'';position:fixed;inset:0;pointer-events:none;z-index:-1;background-image:linear-gradient(rgba(15,23,42,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(15,23,42,.025) 1px,transparent 1px);background-size:42px 42px;mask-image:linear-gradient(to bottom,black,transparent 82%)}
.neo-shell{border:1px solid rgba(255,255,255,.75);background:linear-gradient(145deg,rgba(255,255,255,.86),rgba(255,255,255,.62));backdrop-filter:blur(24px);box-shadow:0 22px 70px rgba(15,23,42,.10)}
header,.card,.bg-white{backdrop-filter:blur(20px)}
button,a,input,select{transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease}button:hover,a:hover{transform:translateY(-2px)}button:active,a:active{transform:scale(.97)}
input:focus,select:focus{outline:none!important;border-color:var(--cyan)!important;box-shadow:0 0 0 4px rgba(34,211,238,.13)!important}
::-webkit-scrollbar{width:10px;height:10px}::-webkit-scrollbar-track{background:#e8eef7}::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#34d399,#3b82f6);border-radius:20px;border:2px solid #e8eef7}
table th{letter-spacing:.035em}tbody tr{transition:background .18s ease,transform .18s ease}tbody tr:hover{background:rgba(224,242,254,.48)!important}
@keyframes neoIn{from{opacity:0;transform:translateY(12px) scale(.99)}to{opacity:1;transform:none}}main>*,body>main>*{animation:neoIn .45s ease both}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;scroll-behavior:auto!important}}
</style>
<script src="https://cdn.tailwindcss.com"></script>
<style>
*{box-sizing:border-box}body{font-family:Inter,system-ui,sans-serif;background:#030712;overflow:hidden}.grid-bg{position:fixed;inset:0;background-image:linear-gradient(rgba(16,185,129,.06) 1px,transparent 1px),linear-gradient(90deg,rgba(59,130,246,.06) 1px,transparent 1px);background-size:42px 42px;mask-image:linear-gradient(to bottom,black,transparent)}.orb{position:fixed;border-radius:999px;filter:blur(70px);opacity:.38;animation:float 10s ease-in-out infinite}.glass{background:linear-gradient(135deg,rgba(15,23,42,.82),rgba(15,23,42,.58));backdrop-filter:blur(26px);border:1px solid rgba(148,163,184,.18);box-shadow:0 40px 110px rgba(0,0,0,.5)}.scanner{position:relative;overflow:hidden}.scanner:after{content:'';position:absolute;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,#34d399,transparent);box-shadow:0 0 18px #34d399;animation:scan 3.2s linear infinite}.field{background:rgba(2,6,23,.62);border:1px solid rgba(148,163,184,.22);transition:.25s}.field:focus-within{border-color:#34d399;box-shadow:0 0 0 4px rgba(52,211,153,.1);transform:translateY(-2px)}.enter{animation:enter .7s cubic-bezier(.2,.8,.2,1) both}.enter2{animation:enter .7s .12s cubic-bezier(.2,.8,.2,1) both}.pulse-dot{animation:pulse 1.8s infinite}@keyframes enter{from{opacity:0;transform:translateY(25px) scale(.97)}to{opacity:1;transform:none}}@keyframes float{50%{transform:translate(45px,-30px) scale(1.1)}}@keyframes scan{from{top:-5%}to{top:105%}}@keyframes pulse{50%{opacity:.35;box-shadow:0 0 18px #34d399}}@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}@media(max-width:800px){body{overflow:auto}}
</style></head><body class="min-h-screen text-slate-100"><div class="grid-bg"></div><div class="orb h-80 w-80 bg-emerald-500 -left-20 -top-20"></div><div class="orb h-96 w-96 bg-blue-600 right-0 bottom-0" style="animation-delay:-4s"></div>
<main class="relative min-h-screen flex items-center justify-center p-4 md:p-8"><section class="glass enter grid w-full max-w-6xl overflow-hidden rounded-[32px] md:grid-cols-[1.15fr_.85fr]">
<div class="scanner hidden md:flex min-h-[650px] flex-col justify-between p-10 lg:p-14 border-r border-white/10"><div><div class="inline-flex items-center gap-2 rounded-full border border-emerald-400/20 bg-emerald-400/10 px-3 py-1.5 text-xs font-black text-emerald-300"><span class="pulse-dot h-2 w-2 rounded-full bg-emerald-400"></span> SYSTEM ONLINE</div><h1 class="mt-8 text-5xl font-black leading-[1.05]">Workforce intelligence,<br><span class="bg-gradient-to-r from-emerald-300 to-cyan-300 bg-clip-text text-transparent">reimagined.</span></h1><p class="mt-5 max-w-lg text-sm leading-7 text-slate-400">Biometric attendance, AI roster planning, shift exceptions, payroll-ready reports and secure workforce insights in one command center.</p></div><div class="grid grid-cols-3 gap-3"><div class="rounded-2xl border border-white/10 bg-white/5 p-4"><p class="text-2xl font-black text-emerald-300">24/7</p><p class="mt-1 text-[10px] uppercase tracking-wider text-slate-400">Live monitoring</p></div><div class="rounded-2xl border border-white/10 bg-white/5 p-4"><p class="text-2xl font-black text-cyan-300">AI</p><p class="mt-1 text-[10px] uppercase tracking-wider text-slate-400">Roster engine</p></div><div class="rounded-2xl border border-white/10 bg-white/5 p-4"><p class="text-2xl font-black text-violet-300">360°</p><p class="mt-1 text-[10px] uppercase tracking-wider text-slate-400">Reports</p></div></div></div>
<div class="enter2 flex items-center p-7 md:p-10 lg:p-14"><div class="w-full"><div class="flex items-center gap-3"><div class="rounded-2xl bg-white p-2 shadow-2xl"><img src="{{ url_for('static',filename='fresmart.png') }}" class="h-11 w-auto" alt="Fresmart"></div><div><p class="text-xs font-black tracking-[.18em] text-emerald-300">GAMEK FRESMART EXPRESS</p><p class="text-sm text-slate-500">AI Workforce OS</p></div></div><h2 class="mt-10 text-3xl font-black">Secure portal access</h2><p class="mt-2 text-sm text-slate-400">Authorized administrators and developers only.</p>
{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="mt-5 rounded-xl border {% if category=='success' %}border-emerald-500/30 bg-emerald-500/10 text-emerald-200{% else %}border-rose-500/30 bg-rose-500/10 text-rose-200{% endif %} p-3 text-xs font-bold">{{ message }}</div>{% endfor %}{% endwith %}{% if error %}<div class="mt-5 rounded-xl border border-rose-500/30 bg-rose-500/10 p-3 text-xs font-bold text-rose-200">{{ error }}</div>{% endif %}
<form method="POST" action="/login" class="mt-7 space-y-4"><label class="block"><span class="mb-2 block text-[11px] font-black uppercase tracking-wider text-slate-400">Access ID</span><div class="field flex items-center gap-3 rounded-2xl px-4"><span class="text-emerald-300">◈</span><input name="user_id" required autocomplete="username" class="w-full bg-transparent py-4 text-sm font-bold outline-none placeholder:text-slate-600" placeholder="Enter secure user ID"></div></label><label class="block"><span class="mb-2 block text-[11px] font-black uppercase tracking-wider text-slate-400">Security key</span><div class="field flex items-center gap-3 rounded-2xl px-4"><span class="text-cyan-300">◆</span><input id="loginPassword" name="password" type="password" required autocomplete="current-password" class="w-full bg-transparent py-4 text-sm font-bold outline-none placeholder:text-slate-600" placeholder="Enter password"><button type="button" id="togglePassword" class="text-xs font-black text-slate-400 hover:text-white">SHOW</button></div></label><button class="group relative w-full overflow-hidden rounded-2xl bg-gradient-to-r from-emerald-500 to-cyan-500 p-4 font-black text-slate-950 shadow-[0_18px_45px_rgba(16,185,129,.22)] transition hover:-translate-y-1"><span class="relative z-10">ENTER COMMAND CENTER →</span><span class="absolute inset-0 translate-x-[-110%] bg-white/30 transition duration-500 group-hover:translate-x-[110%]"></span></button></form><div class="mt-7 flex items-center justify-between text-[10px] uppercase tracking-wider text-slate-600"><span>Encrypted session</span><span>Developer: Sonu Kumar · NCSA0608</span></div></div></div></section></main>
<script>const t=document.getElementById('togglePassword'),p=document.getElementById('loginPassword');t.addEventListener('click',()=>{p.type=p.type==='password'?'text':'password';t.textContent=p.type==='password'?'SHOW':'HIDE'});
</script></body></html>
"""

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gamek AI Workforce Command Center</title>
<style id="gamek-neo-ui">
:root{--neo:#07111f;--neo2:#0d1b2e;--cyan:#22d3ee;--mint:#34d399;--violet:#8b5cf6;--line:rgba(148,163,184,.18);--glass:rgba(255,255,255,.72)}
html{scroll-behavior:smooth}body{background:radial-gradient(circle at 8% 4%,rgba(52,211,153,.20),transparent 25%),radial-gradient(circle at 92% 5%,rgba(59,130,246,.20),transparent 28%),linear-gradient(145deg,#f8fafc,#eef4ff)!important;color:#142033}
body:before{content:'';position:fixed;inset:0;pointer-events:none;z-index:-1;background-image:linear-gradient(rgba(15,23,42,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(15,23,42,.025) 1px,transparent 1px);background-size:42px 42px;mask-image:linear-gradient(to bottom,black,transparent 82%)}
.neo-shell{border:1px solid rgba(255,255,255,.75);background:linear-gradient(145deg,rgba(255,255,255,.86),rgba(255,255,255,.62));backdrop-filter:blur(24px);box-shadow:0 22px 70px rgba(15,23,42,.10)}
header,.card,.bg-white{backdrop-filter:blur(20px)}
button,a,input,select{transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease}button:hover,a:hover{transform:translateY(-2px)}button:active,a:active{transform:scale(.97)}
input:focus,select:focus{outline:none!important;border-color:var(--cyan)!important;box-shadow:0 0 0 4px rgba(34,211,238,.13)!important}
::-webkit-scrollbar{width:10px;height:10px}::-webkit-scrollbar-track{background:#e8eef7}::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#34d399,#3b82f6);border-radius:20px;border:2px solid #e8eef7}
table th{letter-spacing:.035em}tbody tr{transition:background .18s ease,transform .18s ease}tbody tr:hover{background:rgba(224,242,254,.48)!important}
@keyframes neoIn{from{opacity:0;transform:translateY(12px) scale(.99)}to{opacity:1;transform:none}}main>*,body>main>*{animation:neoIn .45s ease both}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;scroll-behavior:auto!important}}
</style>

    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; }
        td.nowrap-cell { white-space: nowrap; }
        .excel-table th, .excel-table td { border: 1px solid #e2e8f0 !important; }
        th.sortable { cursor: pointer; user-select: none; transition: background-color 0.15s ease; }
        th.sortable:hover { background-color: #cbd5e1; }
        .stat-card { cursor: pointer; transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1); position: relative; overflow: hidden; }
        .stat-card:hover { transform: translateY(-2px); box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05); }
        body { background: radial-gradient(circle at 8% 8%, #d1fae5 0, transparent 24%), radial-gradient(circle at 92% 12%, #dbeafe 0, transparent 28%), #f8fafc; }
        main > * { animation: portalRise .5s ease both; }
        main > *:nth-child(2){animation-delay:.05s} main > *:nth-child(3){animation-delay:.1s} main > *:nth-child(4){animation-delay:.15s}
        .bg-white, .stat-card { backdrop-filter: blur(14px); }
        button, a, select, input { transition: transform .2s ease, box-shadow .2s ease, background-color .2s ease; }
        button:hover, a:hover { transform: translateY(-2px); }
        button:active, a:active { transform: scale(.97); }
        .stat-card { animation: portalRise .55s ease both; }
        @keyframes portalRise { from { opacity:0; transform:translateY(12px) } to { opacity:1; transform:none } }

        :root { --panel: rgba(255,255,255,.72); --stroke: rgba(148,163,184,.2); }
        body { background: radial-gradient(circle at 5% 5%,#d1fae5 0,transparent 24%),radial-gradient(circle at 95% 8%,#dbeafe 0,transparent 27%),linear-gradient(145deg,#f8fafc,#eef2ff)!important; }
        aside { background: rgba(255,255,255,.72)!important; backdrop-filter: blur(24px); border-right:1px solid var(--stroke)!important; box-shadow:20px 0 55px rgba(15,23,42,.05); }
        header { background: rgba(255,255,255,.68)!important; backdrop-filter: blur(20px); }
        main { background: transparent!important; }
        .stat-card { border:1px solid rgba(255,255,255,.7)!important; background:linear-gradient(145deg,rgba(255,255,255,.88),rgba(255,255,255,.57))!important; backdrop-filter:blur(18px); box-shadow:0 16px 45px rgba(15,23,42,.07)!important; border-radius:22px!important; }
        .stat-card:before { content:''; position:absolute; inset:0; border-radius:inherit; background:linear-gradient(120deg,transparent,rgba(255,255,255,.65),transparent); transform:translateX(-110%); transition:.65s; }
        .stat-card:hover:before { transform:translateX(110%); }
        aside a { border:1px solid transparent; }
        aside a:hover { background:linear-gradient(90deg,rgba(16,185,129,.12),rgba(59,130,246,.08))!important; border-color:rgba(16,185,129,.14); transform:translateX(4px); }
        table { border-collapse:separate!important; border-spacing:0!important; }
        .excel-table { border-radius:20px; overflow:hidden; box-shadow:0 18px 55px rgba(15,23,42,.06); }
        input,select { background-color:rgba(255,255,255,.78)!important; }
        .future-glow { position:fixed; width:340px; height:340px; border-radius:50%; filter:blur(90px); opacity:.14; pointer-events:none; animation:drift 12s ease-in-out infinite; }
        @keyframes drift { 50% { transform:translate(70px,-45px) scale(1.15); } }


        aside{background:linear-gradient(180deg,rgba(5,15,29,.97),rgba(12,25,44,.95))!important;color:#dbeafe!important;border-right:1px solid rgba(34,211,238,.15)!important}
        aside a{color:#cbd5e1!important;border-radius:14px!important}aside a:hover{color:white!important;background:linear-gradient(90deg,rgba(52,211,153,.22),rgba(59,130,246,.18))!important;box-shadow:inset 3px 0 #34d399}
        aside .text-slate-900,aside .text-slate-800,aside .text-slate-700{color:#f8fafc!important}aside .text-slate-500,aside .text-slate-400{color:#94a3b8!important}
        header{border-bottom:1px solid rgba(148,163,184,.16)!important;box-shadow:0 12px 35px rgba(15,23,42,.05)}
        .stat-card{min-height:96px;border-radius:24px!important;box-shadow:0 18px 55px rgba(15,23,42,.08)!important}.stat-card:hover{transform:translateY(-6px) scale(1.015)!important}
        .excel-table thead th{background:linear-gradient(180deg,#101c30,#17263d)!important;color:#e2e8f0!important;position:sticky;top:0;z-index:10}.excel-table tbody tr:nth-child(even){background:rgba(248,250,252,.76)}


        #dashboardWorkspace.workspace-mode{padding:0!important;overflow:hidden!important;background:#f3f7fc!important}
        #integratedPageFrame{display:block;width:100%;height:100%;min-height:calc(100vh - 64px);border:0;background:#f3f7fc}
        .workspace-active-link{background:linear-gradient(90deg,rgba(52,211,153,.24),rgba(59,130,246,.18))!important;color:#fff!important;box-shadow:inset 3px 0 #34d399}
        #workspaceLoadBar{position:absolute;left:0;right:0;top:0;z-index:80;height:3px;overflow:hidden;background:rgba(15,23,42,.08)}
        #workspaceLoadBar:after{content:'';display:block;width:35%;height:100%;background:linear-gradient(90deg,#34d399,#22d3ee,#6366f1);animation:workspaceLoad 1s ease-in-out infinite}
        @keyframes workspaceLoad{from{transform:translateX(-110%)}to{transform:translateX(320%)}}
        @media (prefers-reduced-motion: reduce) { *, *:before, *:after { animation:none!important; transition:none!important; } }

    </style>
    <script>
        let inactivityTimer;
        const INACTIVITY_LIMIT = 10 * 60 * 1000;

        function resetInactivityTimer() {
            clearTimeout(inactivityTimer);
            inactivityTimer = setTimeout(() => {
                alert("Session timeout ho gaya hai. Dobara login karein.");
                window.location.href = "/logout";
            }, INACTIVITY_LIMIT);
        }

        window.onload = function() {
            const events = ['mousemove', 'keypress', 'click', 'scroll', 'touchstart'];
            events.forEach(eventName => {
                document.addEventListener(eventName, resetInactivityTimer, true);
            });
            resetInactivityTimer();
        };

        function updateLiveClock() {
            const now = new Date();
            const options = { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' };
            const dateStr = now.toLocaleDateString('en-US', options);
            let hours = String(now.getHours()).padStart(2, '0');
            let minutes = String(now.getMinutes()).padStart(2, '0');
            let seconds = String(now.getSeconds()).padStart(2, '0');
            const clockEl = document.getElementById('live-digital-clock');
            if (clockEl) clockEl.innerText = dateStr + ' | ' + hours + ':' + minutes + ':' + seconds;
        }

        let sortDirections = {};
        function sortTable(columnIndex, isNumeric = false) {
            const table = document.getElementById("attendance-table");
            if (!table) return;
            const tbody = table.tBodies[0];
            const rows = Array.from(tbody.querySelectorAll("tr"));
            if (rows.length <= 1 && rows[0].cells.length <= 1) return;

            let dir = sortDirections[columnIndex] || 'asc';
            sortDirections[columnIndex] = (dir === 'asc') ? 'desc' : 'asc';

            rows.sort((rowA, rowB) => {
                let cellA = rowA.cells[columnIndex].innerText.trim();
                let cellB = rowB.cells[columnIndex].innerText.trim();
                if (isNumeric) {
                    let valA = parseFloat(cellA.replace(/[^0-9.-]+/g,"")) || 0;
                    let valB = parseFloat(cellB.replace(/[^0-9.-]+/g,"")) || 0;
                    return (dir === 'asc') ? valA - valB : valB - valA;
                } else {
                    return (dir === 'asc') ? cellA.localeCompare(cellB) : cellB.localeCompare(cellA);
                }
            });

            tbody.innerHTML = "";
            rows.forEach((row, index) => {
                if (row.cells[0]) row.cells[0].innerText = index + 1;
                tbody.appendChild(row);
            });
        }

        function filterByStatus(statusVal) {
            let table = document.getElementById('attendance-table');
            if (!table) return;
            let trs = table.tBodies[0].getElementsByTagName('tr');
            for (let i = 0; i < trs.length; i++) {
                if (statusVal === 'ALL') {
                    trs[i].style.display = "";
                } else if (statusVal === 'Late Arrival') {
                    trs[i].style.display = (trs[i].getAttribute('data-late') === 'Yes') ? "" : "none";
                } else if (statusVal === 'Shift A' || statusVal === 'Shift B') {
                    trs[i].style.display = (trs[i].getAttribute('data-shift') === statusVal) ? "" : "none";
                } else if (statusVal === 'ML') {
                    let statusCell = trs[i].getElementsByTagName('td')[12];
                    if (statusCell) {
                        let text = statusCell.textContent || statusCell.innerText;
                        trs[i].style.display = (text.includes('F01;1') || text.includes('F05;1') || text.includes('F10;1') || text.includes('F51;1') || text.includes('F60;1') || text.includes('F61;1') || text.includes('F62;1') || text.includes('Leave')) ? "" : "none";
                    }
                } else {
                    let statusCell = trs[i].getElementsByTagName('td')[12];
                    if (statusCell) {
                        trs[i].style.display = (statusCell.textContent || statusCell.innerText).includes(statusVal) ? "" : "none";
                    }
                }
            }
        }

        function filterTableSearch() {
            let input = document.getElementById('table-search-input').value.toLowerCase();
            let table = document.getElementById('attendance-table');
            if (!table) return;
            let trs = table.tBodies[0].getElementsByTagName('tr');
            for (let i = 0; i < trs.length; i++) {
                let idCell = trs[i].getElementsByTagName('td')[2];
                let nameCell = trs[i].getElementsByTagName('td')[3];
                let deptCell = trs[i].getElementsByTagName('td')[4];
                if (idCell && nameCell && deptCell) {
                    let text = (idCell.textContent + ' ' + nameCell.textContent + ' ' + deptCell.textContent).toLowerCase();
                    trs[i].style.display = (text.indexOf(input) > -1) ? "" : "none";
                }
            }
        }

        async function syncBiometricNow() {
            const button = document.getElementById('sync-button');
            const status = document.getElementById('sync-status');
            if (!button || button.disabled) return;

            button.disabled = true;
            button.textContent = 'Syncing...';
            status.textContent = 'Connecting to biometric device';
            try {
                const response = await fetch('/api/attendance/sync', {
                    method: 'POST',
                    credentials: 'same-origin'
                });
                const result = await response.json();
                if (!response.ok) throw new Error(result.message || 'Sync failed');
                status.textContent = result.message;
            } catch (error) {
                status.textContent = error.message || 'Unable to sync biometric data';
            } finally {
                button.disabled = false;
                button.textContent = 'Sync Now';
            }
        }

        function setQuickDate(type) {
            let today = new Date();
            let startInput = document.querySelector('input[name="start_date"]');
            let endInput = document.querySelector('input[name="end_date"]');
            let formatDate = (d) => {
                let month = '' + (d.getMonth() + 1), day = '' + d.getDate(), year = d.getFullYear();
                if (month.length < 2) month = '0' + month;
                if (day.length < 2) day = '0' + day;
                return [year, month, day].join('-');
            };

            if (type === 'today') { startInput.value = formatDate(today); endInput.value = formatDate(today); }
            else if (type === 'yesterday') { let yest = new Date(); yest.setDate(today.getDate() - 1); startInput.value = formatDate(yest); endInput.value = formatDate(yest); }
            else if (type === 'week') { let firstDay = new Date(today.setDate(today.getDate() - today.getDay())); startInput.value = formatDate(firstDay); endInput.value = formatDate(new Date()); }
            else if (type === 'month') { let firstDay = new Date(today.getFullYear(), today.getMonth(), 1); startInput.value = formatDate(firstDay); endInput.value = formatDate(new Date()); }
            else if (type === 'last_month') { let firstDay = new Date(today.getFullYear(), today.getMonth() - 1, 1); let lastDay = new Date(today.getFullYear(), today.getMonth(), 0); startInput.value = formatDate(firstDay); endInput.value = formatDate(lastDay); }
            
            document.getElementById('filter-form').submit();
        }

        function openExportModal() { document.getElementById('export-modal').classList.remove('hidden'); }
        function closeExportModal() { document.getElementById('export-modal').classList.add('hidden'); }

        function openRosterModal() { document.getElementById('roster-modal').classList.remove('hidden'); }
        function closeRosterModal() { document.getElementById('roster-modal').classList.add('hidden'); }

        function openCalendarModal() { document.getElementById('calendar-modal').classList.remove('hidden'); }
        function closeCalendarModal() { document.getElementById('calendar-modal').classList.add('hidden'); }

        function openLeaveModal() { document.getElementById('leave-modal').classList.remove('hidden'); }
        function closeLeaveModal() { document.getElementById('leave-modal').classList.add('hidden'); }

        function secureShutdown() {
            let pwd = prompt("Server band karne ke liye password enter karein:");
            if (pwd) window.location.href = "/shutdown?pwd=" + encodeURIComponent(pwd);
        }

        document.addEventListener('DOMContentLoaded', function() {
            setInterval(updateLiveClock, 1000);
            updateLiveClock();
        });
    </script>
</head>
<body class="bg-slate-50 text-slate-800 antialiased flex h-screen overflow-hidden"><div class="future-glow bg-emerald-400 -left-28 top-20"></div><div class="future-glow bg-blue-500 right-0 bottom-0" style="animation-delay:-5s"></div>
    
    <!-- Sidebar Navigation -->
    <aside class="w-64 bg-white border-r border-slate-200 flex flex-col justify-between hidden lg:flex z-20">
        <div>
            <!-- Logo Header -->
            <div class="p-5 flex items-center space-x-3 border-b border-slate-100">
                <div class="bg-[#78b13f] p-2 rounded-xl shadow-sm">
                    <img src="{{ url_for('static', filename='fresmart.png') }}" alt="Logo" class="h-6 object-contain">
                </div>
                <div>
                    <h2 class="text-sm font-bold text-slate-900 leading-tight">Gamek HRM</h2>
                    <p class="text-[10px] text-slate-400 font-medium">Fresmart Express LM11</p>
                </div>
            </div>

            <!-- Menu Links -->
            <div class="p-4 space-y-1">
                <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400 px-3 mb-2">Main Menu</p>
                <a href="/" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl bg-slate-900 text-white font-semibold text-xs shadow-sm">
                    <span>📊</span>
                    <span>AI Command Center</span>
                </a>
                <a href="{{ url_for('projects_tasks') }}" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition"><span>📁</span><span>Projects & Tasks</span></a>
                <a href="{{ url_for('attendance_regularisation') }}" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition"><span>🕒</span><span>Attendance Regularisation</span></a>
                {% if role == 'developer' %}<a href="{{ url_for('developer_stores') }}" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition"><span>🏬</span><span>Multi Store & Bridge</span></a><a href="{{ url_for('developer_biometric_employees') }}" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition"><span>🧬</span><span>Biometric Enrollment</span></a>{% endif %}
                {% if role in ['admin', 'developer'] %}
                <a href="{{ url_for('monthly_roster', month=start_date[:7]) }}" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>📅</span>
                    <span>Monthly Roster</span>
                </a>
                {% else %}
                <a href="#" onclick="openCalendarModal(); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>📅</span>
                    <span>Calendar & Rota</span>
                </a>
                {% endif %}
                <a href="#" onclick="openLeaveModal(); return false;" class="flex items-center justify-between px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <div class="flex items-center space-x-3">
                        <span>🏖️</span>
                        <span>Leave Management</span>
                    </div>
                    {% if role in ['admin', 'developer'] and pending_leaves_count > 0 %}
                    <span class="bg-rose-500 text-white text-[10px] font-bold px-2 py-0.5 rounded-full animate-pulse">{{ pending_leaves_count }}</span>
                    {% endif %}
                </a>
                <a href="#" onclick="alert('Module under preparation.'); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>⚙️</span>
                    <span>Settings</span>
                </a>

                <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400 px-3 mt-6 mb-2">Team Management</p>
                <a href="{{ url_for('performance_dashboard') }}" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>📈</span>
                    <span>Performance</span>
                </a>
                <a href="{{ url_for('day_wise_reports', start_date=start_date, end_date=end_date, employee=selected_emp) }}" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>📋</span><span>Day-wise Reports</span>
                </a>
                <a href="#" onclick="openExportModal(); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>💰</span><span>Payroll Intelligence</span>
                </a>
                <a href="#" onclick="openRosterModal(); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>👥</span>
                    <span>Workforce Directory</span>
                </a>
            </div>
        </div>

        <!-- Sidebar Footer -->
        <div class="p-4 border-t border-slate-100">
            <div class="bg-emerald-50 border border-emerald-100 rounded-2xl p-3.5 space-y-2">
                <div class="flex items-center space-x-2 text-emerald-800 font-bold text-xs">
                    <span>📢</span>
                    <span>Announcements</span>
                </div>
                <p class="text-[11px] text-slate-600 leading-tight">Biometric live tracking active for Gamek Fresmart Express LM11.</p>
                <div class="text-[10px] text-emerald-600 font-bold pt-1">Dev: Sonu Kumar (NCSA0608)</div>
            </div>
        </div>
    </aside>

    <!-- Main Wrapper -->
    <div class="flex-1 flex flex-col h-screen overflow-hidden">
        
        <!-- Top Navbar -->
        <header class="bg-white border-b border-slate-200 px-6 py-3.5 flex justify-between items-center z-10">
            <div class="flex items-center space-x-3">
                <h1 class="text-base font-black text-slate-900 tracking-tight" id="persistentPageTitle">Dashboard</h1>
                <span class="text-xs text-slate-400 font-medium">| Good day, {{ logged_user_name }}</span>
            </div>

            <div class="flex items-center space-x-3 flex-wrap">
                <button onclick="openLeaveModal()" class="relative bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-bold px-3 py-2 rounded-xl transition border border-emerald-200 flex items-center space-x-1.5">
                    <span>🏖️ Leave Portal</span>
                    {% if role in ['admin', 'developer'] and pending_leaves_count > 0 %}
                    <span class="bg-rose-500 text-white text-[10px] px-1.5 py-0.2 rounded-full font-black animate-bounce">{{ pending_leaves_count }}</span>
                    {% endif %}
                </button>

                <div class="text-xs bg-slate-50 px-3 py-2 rounded-xl border border-slate-200 flex items-center space-x-2">
                    <span class="h-2 w-2 {% if stats.device_online %}bg-emerald-500{% else %}bg-red-500{% endif %} rounded-full animate-pulse"></span>
                    <span class="text-slate-600 font-medium">Device: <strong class="{% if stats.device_online %}text-emerald-600{% else %}text-red-600{% endif %}">{% if stats.device_online %}Online{% else %}Offline{% endif %}</strong></span>
                    <span class="text-slate-300">|</span>
                    <span id="live-digital-clock" class="text-slate-700 font-semibold"></span>
                    {% if role in ['admin', 'developer'] %}
                    <span class="text-slate-300">|</span>
                    <button id="sync-button" type="button" onclick="syncBiometricNow()" class="rounded-lg bg-emerald-600 px-3 py-1.5 font-bold text-white hover:bg-emerald-700 disabled:cursor-wait disabled:opacity-60">Sync Now</button>
                    <span id="sync-status" role="status" class="max-w-56 text-slate-500"></span>
                    {% endif %}
                </div>

                <div class="flex items-center space-x-2 bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-xl text-xs font-bold text-slate-700">
                    <span>👤 {{ logged_user_name }}</span>
                    <a href="/logout" class="text-rose-600 hover:text-rose-700 ml-2 font-semibold">Logout 🔒</a>
                </div>

                {% if role == 'admin' or role == 'developer' %}
                {% if role == 'developer' and session.get('user_id') == 'NCSA0608' %}<a href="{{ url_for('portal_update_manager') }}" class="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-bold px-3 py-2 rounded-xl transition border border-indigo-200">🚀 Portal Update</a>{% endif %}
                <button onclick="secureShutdown()" class="bg-rose-50 hover:bg-rose-100 text-rose-700 text-xs font-bold px-3 py-2 rounded-xl transition border border-rose-200">
                    🛑 Shutdown
                </button>
                {% endif %}
            </div>
        </header>

        <!-- Main Content Area -->
        <main id="dashboardWorkspace" class="flex-1 overflow-y-auto p-6 space-y-6">
            
            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                    <div class="{% if category == 'success' %}bg-emerald-50 border-emerald-200 text-emerald-800{% else %}bg-rose-50 border-rose-200 text-rose-700{% endif %} border text-xs font-bold p-4 rounded-2xl shadow-sm flex items-center justify-between">
                        <span>{{ message }}</span>
                        <span class="cursor-pointer" onclick="this.parentElement.style.display='none'">✕</span>
                    </div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            <!-- Quick Top Cards / Stat Overview -->
            <div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 lg:grid-cols-9 gap-3">
                <div onclick="filterByStatus('Present')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-emerald-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Present</p>
                        <h3 class="text-xl font-black text-emerald-600 mt-0.5">{{ stats.present }}</h3>
                    </div>
                    <div class="p-2 bg-emerald-50 text-emerald-600 rounded-xl">✅</div>
                </div>
                <div onclick="filterByStatus('Absent')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-rose-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Absent</p>
                        <h3 class="text-xl font-black text-rose-600 mt-0.5">{{ stats.absent }}</h3>
                    </div>
                    <div class="p-2 bg-rose-50 text-rose-600 rounded-xl">❌</div>
                </div>
                <div onclick="filterByStatus('ML')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-cyan-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Medical/Leave</p>
                        <h3 class="text-xl font-black text-cyan-600 mt-0.5">{{ stats.ml }}</h3>
                    </div>
                    <div class="p-2 bg-cyan-50 text-cyan-600 rounded-xl">🏥</div>
                </div>
                <div onclick="filterByStatus('Weekly Off')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-slate-400">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Weekly Off</p>
                        <h3 class="text-xl font-black text-slate-700 mt-0.5">{{ stats.off }}</h3>
                    </div>
                    <div class="p-2 bg-slate-100 text-slate-600 rounded-xl">🏖️</div>
                </div>
                <div onclick="filterByStatus('Late Arrival')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-amber-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Late Arrival</p>
                        <h3 class="text-xl font-black text-amber-600 mt-0.5">{{ stats.late_arrival }}</h3>
                    </div>
                    <div class="p-2 bg-amber-50 text-amber-600 rounded-xl">⏰</div>
                </div>
                <div onclick="filterByStatus('Mis Punch')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-orange-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Mis-Punches</p>
                        <h3 class="text-xl font-black text-orange-600 mt-0.5">{{ stats.mispunch }}</h3>
                    </div>
                    <div class="p-2 bg-orange-50 text-orange-600 rounded-xl">⚠</div>
                </div>
                <div onclick="filterByStatus('Shift A')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-blue-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Shift A (06-10)</p>
                        <h3 class="text-xl font-black text-blue-600 mt-0.5">{{ stats.shift_a }}</h3>
                    </div>
                    <div class="p-2 bg-blue-50 text-blue-600 rounded-xl">☀️</div>
                </div>
                <div onclick="filterByStatus('Shift B')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-indigo-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Shift B (&gt;10:00)</p>
                        <h3 class="text-xl font-black text-indigo-600 mt-0.5">{{ stats.shift_b }}</h3>
                    </div>
                    <div class="p-2 bg-indigo-50 text-indigo-600 rounded-xl">🌙</div>
                </div>
                <div onclick="filterByStatus('ALL')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-emerald-600">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Total Hours</p>
                        <h3 class="text-xl font-black text-emerald-600 mt-0.5">{{ stats.total_hrs }}</h3>
                    </div>
                    <div class="p-2 bg-emerald-50 text-emerald-600 rounded-xl">⏱</div>
                </div>
            </div>

            <!-- Filter Controls Bar -->
            <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
                <form id="filter-form" method="GET" action="/" class="grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
                    <div>
                        <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Start Date</label>
                        <input type="date" name="start_date" value="{{ start_date }}" onchange="this.form.submit()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">End Date</label>
                        <input type="date" name="end_date" value="{{ end_date }}" onchange="this.form.submit()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                    </div>
                    {% if role == 'developer' %}<div><label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Store Filter</label><select name="store" onchange="this.form.submit()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium"><option value="">-- Select Store to View --</option>{% for sid,x in stores %}<option value="{{sid}}" {% if selected_store==sid %}selected{% endif %}>{{x.store_name}} ({{sid}})</option>{% endfor %}</select></div>{% endif %}
                    {% if role == 'admin' or role == 'developer' %}
                    <div>
                        <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Employee Filter</label>
                        <select name="employee" onchange="this.form.submit()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                            <option value="ALL">-- All Personnel --</option>
                            {% for emp in all_users %}
                                <option value="{{ emp.user_id }}" {% if selected_emp == emp.user_id %}selected{% endif %}>{{ emp.name }} ({{ emp.user_id }})</option>
                            {% endfor %}
                        </select>
                    </div>
                    {% else %}
                    <div>
                        <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Logged In As</label>
                        <input type="hidden" name="employee" value="{{ selected_emp }}">
                        <input type="text" disabled value="{{ selected_emp }}" class="w-full bg-slate-100 border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm font-bold text-emerald-700 cursor-not-allowed">
                    </div>
                    {% endif %}
                    <div class="flex space-x-2">
                        {% if role == 'admin' or role == 'developer' %}
                        <button type="button" onclick="openExportModal()" class="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs uppercase tracking-wider py-3 px-3 rounded-xl text-center shadow-md transition">Export 📥</button>
                        <a href="{{ url_for('matrix_preview', start_date=start_date, end_date=end_date, employee=selected_emp) }}" class="flex-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs uppercase tracking-wider py-3 px-3 rounded-xl text-center shadow-md transition">Matrix Preview</a>
                        {% endif %}
                        {% if role in ['admin', 'developer'] %}
                        <a href="{{ url_for('monthly_roster', month=start_date[:7]) }}" class="bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs uppercase tracking-wider py-3 px-3 rounded-xl text-center shadow-md transition">📅 Monthly Roster</a>
                        {% else %}
                        <button type="button" onclick="openCalendarModal()" class="bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs uppercase tracking-wider py-3 px-3 rounded-xl text-center shadow-md transition">📅 Rota</button>
                        {% endif %}
                    </div>
                </form>
                
                <div class="flex items-center space-x-2 mt-4 pt-4 border-t border-slate-100 text-xs flex-wrap gap-y-2">
                    <span class="text-slate-400 font-bold uppercase tracking-wide mr-1">Quick Range:</span>
                    <button type="button" onclick="setQuickDate('today')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">Today</button>
                    <button type="button" onclick="setQuickDate('yesterday')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">Yesterday</button>
                    <button type="button" onclick="setQuickDate('week')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">This Week</button>
                    <button type="button" onclick="setQuickDate('month')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">This Month</button>
                    <button type="button" onclick="setQuickDate('last_month')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">Last Month</button>
                </div>
            </div>

            <!-- Attendance Data Table -->
            <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="p-4 bg-slate-50 border-b border-slate-200 flex flex-col sm:flex-row justify-between items-center gap-3">
                    <div class="text-xs text-slate-500 font-semibold">
                        💡 Shift A = 1st Punch 06:00-10:00 AM | Shift B = 1st Punch after 10:00 AM.
                    </div>
                    <div>
                        <input type="text" id="table-search-input" onkeyup="filterTableSearch()" placeholder="🔍 Search employee name, ID or department..." class="bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-xs w-72 shadow-sm focus:outline-none focus:border-emerald-500 font-medium">
                    </div>
                </div>
                <div class="overflow-x-auto">
                    <table id="attendance-table" class="w-full text-left border-collapse excel-table">
                        <thead>
                            <tr class="bg-slate-100 text-slate-600 uppercase text-[11px] font-bold tracking-wider">
                                <th class="py-3 px-4">Sr. No.</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(1)">Date ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(2)">ID ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(3)">Employee Name ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(4)">Dept ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(5)">Store In ↕</th>
                                <th class="py-3 px-4">Lunch Out</th>
                                <th class="py-3 px-4">Lunch In</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(8)">Out Time ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(9)">Total Lunch ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(10)">Working Hours ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(11)">Total Hora Extra ↕</th>
                                <th class="py-3 px-4 text-center sortable" onclick="sortTable(12)">Status ↕</th>
                            </tr>
                        </thead>
                        <tbody class="text-sm text-slate-700 divide-y divide-slate-100">
                            {% if logs %}
                                {% for log in logs %}
                                <tr class="hover:bg-slate-50 transition-colors" data-late="{{ log.is_late }}" data-shift="{{ log.shift_type }}">
                                    <td class="py-3 px-4 nowrap-cell font-medium text-slate-400">{{ loop.index }}</td>
                                    <td class="py-3 px-4 nowrap-cell font-medium">{{ log.date }}</td>
                                    <td class="py-3 px-4 nowrap-cell text-slate-500 font-mono text-xs">{{ log.user_id }}</td>
                                    <td class="py-3 px-4 font-bold text-slate-900 nowrap-cell">{{ log.name }}</td>
                                    <td class="py-3 px-4 nowrap-cell"><span class="px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-700 text-[11px] font-bold">{{ log.dept }}</span></td>
                                    <td class="py-3 px-4 nowrap-cell font-mono text-xs">
                                        {{ log.store_in }}
                                        {% if log.shift_type == 'Shift A' %}
                                            <span class="text-[10px] bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded font-sans font-bold ml-1">Shift A</span>
                                        {% elif log.shift_type == 'Shift B' %}
                                            <span class="text-[10px] bg-indigo-100 text-indigo-700 px-1.5 py-0.5 rounded font-sans font-bold ml-1">Shift B</span>
                                        {% endif %}
                                    </td>
                                    <td class="py-3 px-4 nowrap-cell font-mono text-xs text-slate-500">{{ log.lunch_out }}</td>
                                    <td class="py-3 px-4 nowrap-cell font-mono text-xs text-slate-500">{{ log.lunch_in }}</td>
                                    <td class="py-3 px-4 nowrap-cell font-mono text-xs">{{ log.out_time }}</td>
                                    <td class="py-3 px-4 font-bold nowrap-cell font-mono text-xs {% if log.lunch_seconds > 3600 %}text-rose-600 bg-rose-50/50{% else %}text-slate-700{% endif %}">{{ log.total_lunch }}</td>
                                    <td class="py-3 px-4 font-bold text-slate-900 nowrap-cell font-mono text-xs">{{ log.total_hours }}</td>
                                    <td class="py-3 px-4 font-bold nowrap-cell font-mono text-xs {% if log.variance_type == 'positive' %}text-emerald-600{% elif log.variance_type == 'negative' %}text-rose-600{% else %}text-slate-600{% endif %}">{{ log.net_variance }}</td>
                                    <td class="py-3 px-4 text-center nowrap-cell">
                                        {% if log.status == 'Weekly Off' %}
                                            <span class="px-3 py-1 rounded-full text-xs font-bold bg-slate-100 text-slate-600 border border-slate-200">Weekly Off</span>
                                        {% elif log.status == 'Absent' %}
                                            <span class="px-3 py-1 rounded-full text-xs font-bold bg-rose-50 text-rose-600 border border-rose-200">Absent</span>
                                        {% elif 'F01;1' in log.status or 'F05;1' in log.status or 'F10;1' in log.status or 'F51;1' in log.status or 'F60;1' in log.status or 'F61;1' in log.status or 'F62;1' in log.status %}
                                            <span class="px-3 py-1 rounded-full text-xs font-bold bg-cyan-50 text-cyan-700 border border-cyan-200">{{ log.status }}</span>
                                        {% elif log.status == 'ML' %}
                                            <span class="px-3 py-1 rounded-full text-xs font-bold bg-cyan-50 text-cyan-700 border border-cyan-200">Medical/Leave</span>
                                        {% elif log.status == 'Present' %}
                                            <span class="px-3 py-1 rounded-full text-xs font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">Present</span>
                                        {% elif log.status == 'Mis Punch' %}
                                            <span class="px-3 py-1 rounded-full text-xs font-bold bg-orange-50 text-orange-700 border border-orange-200">Mis Punch</span>
                                        {% endif %}
                                    </td>
                                </tr>
                                {% endfor %}
                            {% else %}
                                <tr><td colspan="13" class="text-center py-16 text-slate-400 font-medium">No attendance records found for this selection.</td></tr>
                            {% endif %}
                        </tbody>
                        <tfoot class="bg-slate-100 font-bold text-slate-900 text-sm border-t border-slate-200">
                            <tr>
                                <td colspan="9" class="py-4 px-4 text-right uppercase text-xs tracking-wider text-slate-500">Total Summary:</td>
                                <td class="py-4 px-4 text-emerald-700 font-mono">{{ grand_total_lunch_hours }}</td>
                                <td class="py-4 px-4 text-slate-900 font-mono">{{ grand_total_hours }}</td>
                                <td class="py-4 px-4 font-mono {% if stats.variance_type == 'positive' %}text-emerald-700{% else %}text-rose-700{% endif %}" colspan="2">{{ grand_total_variance }}</td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </main>
    </div>

    {% if role == 'developer' %}
    <div id="developer-backup-panel" class="fixed bottom-4 right-4 z-40 w-[min(430px,calc(100vw-2rem))] bg-white border border-slate-200 shadow-2xl rounded-2xl p-4">
      <div class="flex justify-between items-center mb-3"><div><h3 class="font-bold text-slate-900">Developer Backup & Restore</h3><p class="text-[10px] text-slate-500">Auto backup daily at 12:00 PM and 10:00 PM</p></div><button type="button" onclick="document.getElementById('developer-backup-panel').classList.toggle('max-h-14')" class="text-slate-400">▾</button></div>
      <button type="button" id="manual-backup-btn" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl px-4 py-2">Start Full Backup</button>
      <div class="mt-3"><div class="flex justify-between text-xs mb-1"><span id="backup-message">Ready</span><b id="backup-percent">0%</b></div><div class="h-3 bg-slate-100 rounded-full overflow-hidden"><div id="backup-progress" class="h-full bg-emerald-500 transition-all" style="width:0%"></div></div></div>
      <div class="mt-3 grid grid-cols-1 gap-2"><select id="backup-list" class="border rounded-xl px-3 py-2 text-xs"><option value="">Loading backups...</option></select><a id="download-backup-link" href="#" class="text-center bg-slate-100 rounded-xl px-3 py-2 text-xs font-bold">Download selected backup</a></div>
      <form id="restore-backup-form" class="mt-3 space-y-2" enctype="multipart/form-data"><input id="restore-file" name="backup_file" type="file" accept=".zip" class="w-full text-xs border rounded-xl p-2"><button type="submit" class="w-full bg-rose-600 text-white font-bold rounded-xl px-4 py-2">Upload / Restore Backup</button><p class="text-[10px] text-slate-500">Restore se pehle automatic safety backup banega.</p></form>
      <div id="backup-result" class="text-xs mt-2" aria-live="polite"></div>
    </div>
    <script>
    (()=>{const root=document.getElementById('developer-backup-panel'),btn=document.getElementById('manual-backup-btn'),bar=document.getElementById('backup-progress'),pct=document.getElementById('backup-percent'),msg=document.getElementById('backup-message'),list=document.getElementById('backup-list'),link=document.getElementById('download-backup-link'),result=document.getElementById('backup-result');
    async function loadList(){const r=await fetch('/developer/backups');const d=await r.json();list.innerHTML='<option value="">Select saved backup</option>';(d.backups||[]).forEach(x=>{const o=document.createElement('option');o.value=x.name;o.textContent=`${x.created} - ${x.name}`;list.appendChild(o);});}
    async function poll(){const r=await fetch('/developer/backup/status');const d=await r.json();const n=Math.max(0,Math.min(100,Number(d.percent)||0));bar.style.width=n+'%';pct.textContent=n+'%';msg.textContent=d.error||d.message||'Ready';btn.disabled=!!d.running;if(d.running)setTimeout(poll,700);else loadList();}
    btn.addEventListener('click',async()=>{result.textContent='';await fetch('/developer/backup/start',{method:'POST'});poll();});
    list.addEventListener('change',()=>{link.href=list.value?'/developer/backup/download/'+encodeURIComponent(list.value):'#';});
    document.getElementById('restore-backup-form').addEventListener('submit',async e=>{e.preventDefault();if(!confirm('Selected backup restore karna hai? Current data ka safety backup pehle banega.'))return;const fd=new FormData(e.target);if(!document.getElementById('restore-file').files.length && list.value)fd.append('existing_backup',list.value);result.textContent='Restoring backup...';const r=await fetch('/developer/backup/restore',{method:'POST',body:fd});const d=await r.json();result.textContent=d.message||'Done';result.className='text-xs mt-2 '+(r.ok?'text-emerald-700':'text-rose-700');loadList();});loadList();poll();})();
    </script>
    {% endif %}

    <!-- Leave Management Modal -->
    <div id="leave-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center hidden">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-6 w-full max-w-6xl mx-4 space-y-4 max-h-[95vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-4">
                <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">🏖️ Leave Management & Complete History (Never Deleted)</h3>
                <button onclick="closeLeaveModal()" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            
                <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5 mb-6">
                    <h4 class="text-sm font-bold text-slate-900 mb-3">Shift Swap Request</h4>
                    {% if role == 'employee' %}<form action="/submit_shift_change" method="POST" class="grid grid-cols-1 md:grid-cols-4 gap-3 items-end">
                    <div><label class="block text-xs font-bold mb-1">Swap Date</label><input type="date" name="swap_date" required class="w-full border rounded-xl px-3 py-2"></div>
                    <div><label class="block text-xs font-bold mb-1">Opposite Shift Employee</label><select name="partner_id" required class="w-full border rounded-xl px-3 py-2"><option value="">Select employee</option>{% for emp in swap_partners %}<option value="{{ emp.code }}">{{ emp.name }} - {{ emp.shift }}</option>{% endfor %}</select></div>
                    <div><label class="block text-xs font-bold mb-1">Reason</label><input name="reason" class="w-full border rounded-xl px-3 py-2"></div><button class="bg-indigo-600 text-white font-bold rounded-xl px-4 py-2">Submit Shift Swap</button></form>{% endif %}
                    <div class="overflow-x-auto mt-4"><table class="w-full text-xs"><thead><tr class="bg-slate-50"><th>Requester</th><th>Partner</th><th>Date</th><th>Employee</th><th>Admin</th><th>Status</th><th>Action</th></tr></thead><tbody>{% for req in shift_change_requests %}<tr class="border-t"><td class="p-2">{{ req.requester_name }} ({{ req.requester_shift }})</td><td class="p-2">{{ req.partner_name }} ({{ req.partner_shift }})</td><td class="p-2">{{ req.swap_date }}</td><td class="p-2">{{ req.employee_approval }}</td><td class="p-2">{{ req.admin_approval }}</td><td class="p-2 font-bold">{{ req.status }}</td><td class="p-2">{% if req.status == 'Pending' %}{% if role in ['admin','developer'] or logged_user_id == req.partner_id %}<a class="text-emerald-700 font-bold mr-2" href="/update_shift_change/{{ req.id }}/approve">Approve</a><a class="text-rose-700 font-bold" href="/update_shift_change/{{ req.id }}/reject">Reject</a>{% endif %}{% endif %}</td></tr>{% else %}<tr><td colspan="7" class="p-3 text-center">No shift swap requests</td></tr>{% endfor %}</tbody></table></div>
                </div>

                <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5 mb-6">
                    <h4 class="text-sm font-bold text-slate-900 mb-3">Weekly Off Change Request</h4>
                    {% if role == 'employee' %}
                    <form action="/submit_off_change" method="POST" class="grid grid-cols-1 md:grid-cols-4 gap-3 items-end">
                        <div><label class="block text-xs font-bold text-slate-600 mb-1">Current Off Date</label><input type="date" name="old_off_date" required class="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm"></div>
                        <div><label class="block text-xs font-bold text-slate-600 mb-1">New Off Date</label><input id="off-new-date" type="date" name="new_off_date" required class="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm"></div>
                        <div><label class="block text-xs font-bold text-slate-600 mb-1">Swap Employee</label><select id="off-partner-select" name="partner_id" required disabled class="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm"><option value="">Pehle New Off Date select karein</option></select></div><div><label class="block text-xs font-bold text-slate-600 mb-1">Reason</label><input type="text" name="reason" class="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm" placeholder="Optional reason"></div>
                        <button type="submit" class="bg-amber-500 hover:bg-amber-600 text-white font-bold rounded-xl px-4 py-2">Submit Off Change</button><p id="off-partner-help" class="text-[10px] text-slate-500 md:col-span-4">New Off Date select karne par sirf us date ke Weekly Off employees dikhenge.</p>
                    </form>
                    {% endif %}
                    <div class="overflow-x-auto mt-4">
                    <table class="w-full text-xs"><thead><tr class="bg-slate-50"><th class="p-2 text-left">Employee</th><th class="p-2">Old Off</th><th class="p-2">New Off</th><th class="p-2">Status</th>{% if role in ['admin','developer'] %}<th class="p-2">Action</th>{% endif %}</tr></thead><tbody>
                    {% for req in off_change_requests %}<tr class="border-t"><td class="p-2"><b>{{ req.name }}</b><br><span class="text-slate-400">{{ req.user_id }}</span></td><td class="p-2 text-center">{{ req.old_off_date }}</td><td class="p-2 text-center">{{ req.new_off_date }}</td><td class="p-2 text-center"><b>{{ req.status }}</b></td><td class="p-2 text-center">{% if req.status == 'Pending' and (role in ['admin','developer'] or logged_user_id == req.partner_id) %}<a class="text-emerald-700 font-bold mr-2" href="/update_off_change/{{ req.id }}/approve">Approve</a><a class="text-rose-700 font-bold" href="/update_off_change/{{ req.id }}/reject">Reject</a>{% endif %}</td></tr>{% else %}<tr><td colspan="5" class="p-3 text-center text-slate-400">No off change requests</td></tr>{% endfor %}
                    </tbody></table></div>
                </div>

            <div class="overflow-y-auto flex-1 space-y-6">
                {% if role == 'employee' %}
                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-5 space-y-4">
                    <h4 class="text-sm font-bold text-slate-900">Apply for Leave Request</h4>
                    <form method="POST" action="/apply_leave" enctype="multipart/form-data" class="grid grid-cols-1 sm:grid-cols-3 gap-4">
                        <div>
                            <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">Start Date</label>
                            <input type="date" name="start_date" required class="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">End Date</label>
                            <input type="date" name="end_date" required class="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">Leave Type (Required)</label>
                            <select name="leave_type" required class="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                                <option value="F01;1">F01;1 - Baixa Médica</option>
                                <option value="F03;1">F03;1 - Falta Injustificada</option>
                                <option value="F05;1">F05;1 - Licença sem vencimento</option>
                                <option value="F10;1" selected>F10;1 - Falta Justificada</option>
                                <option value="F51;1">F51;1 - Casamento</option>
                                <option value="F60;1">F60;1 - Nascimento</option>
                                <option value="F61;1">F61;1 - Obito</option>
                                <option value="F62;1">F62;1 - Gravidez</option>
                            </select>
                        </div>
                        <div class="sm:col-span-3">
                            <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">Supporting Document Upload (Optional - PDF/Image)</label>
                            <input type="file" name="supporting_doc" accept=".pdf,image/*" class="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none file:mr-4 file:py-1 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-emerald-50 file:text-emerald-700 hover:file:bg-emerald-100">
                        </div>
                        <div class="sm:col-span-3 flex justify-end">
                            <button type="submit" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs uppercase tracking-wider px-5 py-2.5 rounded-xl shadow transition">Submit Request 🚀</button>
                        </div>
                    </form>
                </div>
                {% endif %}

                <div class="space-y-3">
                    <h4 class="text-sm font-bold text-slate-900 flex items-center justify-between">
                        <span>{% if role == 'employee' %}My Leave History Archive{% else %}All Employees Leave History Archive (Persistent & Permanent){% endif %}</span>
                        {% if role in ['admin', 'developer'] and pending_leaves_count > 0 %}
                        <span class="bg-rose-500 text-white text-[10px] px-2 py-0.5 rounded-full font-bold">{{ pending_leaves_count }} Pending Actions</span>
                        {% endif %}
                    </h4>
                    <div class="border border-slate-200 rounded-xl overflow-hidden">
                        <table class="w-full text-left border-collapse">
                            <thead>
                                <tr class="bg-slate-100 text-slate-600 uppercase text-[10px] font-bold tracking-wider">
                                    <th class="py-2.5 px-3 border-b border-slate-200">ID & Name</th>
                                    <th class="py-2.5 px-3 border-b border-slate-200">From - To Dates</th>
                                    <th class="py-2.5 px-3 border-b border-slate-200">Leave Type</th>
                                    <th class="py-2.5 px-3 border-b border-slate-200">Support Document</th>
                                    <th class="py-2.5 px-3 border-b border-slate-200 text-center">Status</th>
                                    {% if role in ['admin', 'developer'] %}
                                    <th class="py-2.5 px-3 border-b border-slate-200 text-center">Action</th>
                                    {% endif %}
                                </tr>
                            </thead>
                            <tbody class="text-xs text-slate-700 divide-y divide-slate-100">
                                {% if leave_requests %}
                                    {% for req in leave_requests %}
                                    <tr class="hover:bg-slate-50">
                                        <td class="py-2.5 px-3">
                                            <div class="font-bold text-slate-900">{{ req.name }}</div>
                                            <div class="text-[10px] text-slate-400 font-mono">{{ req.user_id }}</div>
                                        </td>
                                        <td class="py-2.5 px-3 font-mono text-[11px]">{{ req.start_date }} to {{ req.end_date }}</td>
                                        <td class="py-2.5 px-3 font-bold text-cyan-700">{{ req.leave_type }}</td>
                                        <td class="py-2.5 px-3 text-slate-600">
                                            {% if req.filename %}
                                            <a href="/uploads/{{ req.filename }}" target="_blank" class="inline-flex items-center space-x-1 mt-1 text-[11px] text-emerald-600 hover:text-emerald-700 font-bold bg-emerald-50 px-2 py-0.5 rounded-lg border border-emerald-200 transition">
                                                <span>📎 View Document</span>
                                            </a>
                                            {% else %}
                                            <span class="text-[10px] text-slate-400 italic">No attachment</span>
                                            {% endif %}
                                        </td>
                                        <td class="py-2.5 px-3 text-center">
                                            {% if req.status == 'Approved' %}
                                                <span class="px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 font-bold text-[10px]">Approved ({{ req.leave_type }}) ✅</span>
                                            {% elif req.status == 'Rejected' %}
                                                <span class="px-2.5 py-1 rounded-full bg-rose-50 text-rose-700 font-bold text-[10px]">Rejected ❌</span>
                                            {% else %}
                                                <span class="px-2.5 py-1 rounded-full bg-amber-50 text-amber-700 font-bold text-[10px] animate-pulse">Pending ⏳</span>
                                            {% endif %}
                                        </td>
                                        {% if role in ['admin', 'developer'] %}
                                        <td class="py-2.5 px-3 text-center">
                                            {% if req.status == 'Pending' %}
                                            <div class="flex items-center justify-center space-x-1.5">
                                                <a href="/update_leave/{{ req.id }}/approve" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-2.5 py-1 rounded-lg text-[10px] transition">Accept</a>
                                                <a href="/update_leave/{{ req.id }}/reject" class="bg-rose-600 hover:bg-rose-700 text-white font-bold px-2.5 py-1 rounded-lg text-[10px] transition">Reject</a>
                                            </div>
                                            {% else %}
                                            <span class="text-slate-400 text-[10px] italic">Processed</span>
                                            {% endif %}
                                        </td>
                                        {% endif %}
                                    </tr>
                                    {% endfor %}
                                {% else %}
                                    <tr><td colspan="6" class="text-center py-8 text-slate-400">No leave history records found.</td></tr>
                                {% endif %}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <div class="pt-2 flex justify-end">
                <button onclick="closeLeaveModal()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition">Close</button>
            </div>
        </div>
    </div>

    <!-- Export Modal -->
    <div id="export-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center hidden">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-6 w-full max-w-md mx-4 space-y-6">
            <div class="flex justify-between items-center border-b border-slate-100 pb-4">
                <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">📊 Choose Excel Export Format</h3>
                <button onclick="closeExportModal()" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            <div class="space-y-4">
                <a href="/export?start_date={{ start_date }}&end_date={{ end_date }}&employee={{ selected_emp }}" onclick="closeExportModal()" class="block p-4 rounded-xl border border-slate-200 hover:border-emerald-500 hover:bg-emerald-50/50 transition group">
                    <div class="font-bold text-slate-900 group-hover:text-emerald-700 flex items-center justify-between">
                        <span>Option 1: Standard Row Export</span>
                        <span>📄</span>
                    </div>
                    <p class="text-xs text-slate-500 mt-1">Same as current view with individual record rows for each employee per date.</p>
                </a>
                
                <a href="/export_matrix?start_date={{ start_date }}&end_date={{ end_date }}&employee={{ selected_emp }}" onclick="closeExportModal()" class="block p-4 rounded-xl border border-slate-200 hover:border-emerald-500 hover:bg-emerald-50/50 transition group">
                    <div class="font-bold text-slate-900 group-hover:text-emerald-700 flex items-center justify-between">
                        <span>Option 2: Employee Matrix (Leave Codes)</span>
                        <span>📅</span>
                    </div>
                    <p class="text-xs text-slate-500 mt-1">Employees in rows, Dates in columns. Specific codes (F01;1, F10;1, H06, H07, Off) for attendance statuses.</p>
                </a>
            </div>
            <div class="pt-2 flex justify-end">
                <button onclick="closeExportModal()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition">Cancel</button>
            </div>
        </div>
    </div>

    <!-- Employees Roster Modal -->
    <div id="roster-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center hidden">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-6 w-full max-w-4xl mx-4 space-y-6 max-h-[85vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-4">
                <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">👥 Gamek Fresmart Express - Employees Roster</h3>
                <button onclick="closeRosterModal()" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            <div class="overflow-y-auto flex-1">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-100 text-slate-600 uppercase text-[11px] font-bold tracking-wider">
                            <th class="py-3 px-4 border border-slate-200">Sr.</th>
                            <th class="py-3 px-4 border border-slate-200">ID</th>
                            <th class="py-3 px-4 border border-slate-200">Name</th>
                            <th class="py-3 px-4 border border-slate-200">Department</th>
                            <th class="py-3 px-4 border border-slate-200">Weekly Off</th>
                        </tr>
                    </thead>
                    <tbody class="text-xs text-slate-700 divide-y divide-slate-100">
                        {% for emp in all_users %}
                        <tr class="hover:bg-slate-50">
                            <td class="py-2.5 px-4 border border-slate-200 text-slate-400 font-medium">{{ loop.index }}</td>
                            <td class="py-2.5 px-4 border border-slate-200 font-mono text-slate-600 font-semibold">{{ emp.user_id }}</td>
                            <td class="py-2.5 px-4 border border-slate-200 font-bold text-slate-900">{{ emp.name }}</td>
                            <td class="py-2.5 px-4 border border-slate-200"><span class="px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold">{{ emp.dept }}</span></td>
                            <td class="py-2.5 px-4 border border-slate-200 font-bold text-indigo-600">{{ emp.off }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
            <div class="pt-2 flex justify-end">
                <button onclick="closeRosterModal()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition">Close</button>
            </div>
        </div>
    </div>


    <script>
    document.addEventListener('DOMContentLoaded', function () {
        const dateInput = document.getElementById('off-new-date');
        const partnerSelect = document.getElementById('off-partner-select');
        const help = document.getElementById('off-partner-help');
        if (!dateInput || !partnerSelect) return;
        async function loadOffEmployees() {
            const selectedDate = dateInput.value;
            partnerSelect.disabled = true;
            partnerSelect.innerHTML = '<option value="">Loading...</option>';
            if (!selectedDate) {
                partnerSelect.innerHTML = '<option value="">Pehle New Off Date select karein</option>';
                help.textContent = 'New Off Date select karne par sirf us date ke Weekly Off employees dikhenge.';
                return;
            }
            try {
                const response = await fetch('/api/employees-off-on-date?date=' + encodeURIComponent(selectedDate));
                const data = await response.json();
                const employees = data.employees || [];
                partnerSelect.innerHTML = '<option value="">Select employee</option>';
                employees.forEach(emp => {
                    const option = document.createElement('option');
                    option.value = emp.code;
                    option.textContent = emp.name + ' (' + emp.code + ')';
                    partnerSelect.appendChild(option);
                });
                partnerSelect.disabled = employees.length === 0;
                help.textContent = employees.length ? employees.length + ' employee(s) ka Weekly Off is date par hai.' : 'Is date par kisi doosre employee ka Weekly Off nahi mila.';
            } catch (error) {
                partnerSelect.innerHTML = '<option value="">Employees load nahi hue</option>';
                help.textContent = 'Employee list load karte samay error hua.';
            }
        }
        dateInput.addEventListener('change', loadOffEmployees);
    });
    
(function(){
 const workspace=document.getElementById('dashboardWorkspace');
 const title=document.getElementById('persistentPageTitle');
 if(!workspace)return;
 const routes={
   '/projects-tasks':'Projects, Tasks & Exams',
   '/monthly_roster':'Monthly Roster',
   '/performance':'Employee Performance',
   '/reports/day-wise':'Day-wise Reports',
   '/matrix_preview':'Attendance Matrix Preview'
 };
 const originalDashboard=workspace.innerHTML;
 let activeLink=null;
 function isWorkspaceLink(link){
   if(!link||!link.href)return false;
   const url=new URL(link.href,location.origin);
   return url.origin===location.origin && Object.prototype.hasOwnProperty.call(routes,url.pathname);
 }
 function setActive(link){
   if(activeLink)activeLink.classList.remove('workspace-active-link');
   activeLink=link||null;
   if(activeLink)activeLink.classList.add('workspace-active-link');
 }
 function openInsideDashboard(url,label,link,pushState=true){
   const parsed=new URL(url,location.origin);
   workspace.classList.add('workspace-mode');
   workspace.innerHTML='<div id="workspaceLoadBar"></div><iframe id="integratedPageFrame" title="'+label+'"></iframe>';
   if(title)title.textContent=label;
   setActive(link);
   const frame=document.getElementById('integratedPageFrame');
   frame.src=parsed.pathname+parsed.search;
   frame.addEventListener('load',()=>{
     document.getElementById('workspaceLoadBar')?.remove();
     try{
       const doc=frame.contentDocument;
       if(doc){
         const style=doc.createElement('style');
         style.textContent='body{min-height:auto!important} main{max-width:none!important;padding:18px!important} a[href="/"]{display:none!important}';
         doc.head.appendChild(style);
       }
     }catch(err){}
   });
   if(pushState)history.pushState({workspace:parsed.pathname+parsed.search},'',`#page=${encodeURIComponent(parsed.pathname+parsed.search)}`);
 }
 function showDashboard(pushState=true){
   workspace.classList.remove('workspace-mode');
   workspace.innerHTML=originalDashboard;
   if(title)title.textContent='Dashboard';
   setActive(document.querySelector('aside a[href="/"]'));
   if(pushState)history.pushState({workspace:null},'',location.pathname+location.search);
 }
 document.addEventListener('click',event=>{
   const link=event.target.closest('a');
   if(!link||event.ctrlKey||event.metaKey||event.shiftKey||event.altKey)return;
   const url=new URL(link.href||'',location.origin);
   if(isWorkspaceLink(link)){
     event.preventDefault();
     openInsideDashboard(url.href,routes[url.pathname],link,true);
   }else if(url.origin===location.origin && url.pathname==='/' && link.closest('aside')){
     event.preventDefault();
     showDashboard(true);
   }
 });
 window.addEventListener('popstate',event=>{
   if(event.state&&event.state.workspace){
     const parsed=new URL(event.state.workspace,location.origin);
     openInsideDashboard(parsed.href,routes[parsed.pathname]||'Workspace',null,false);
   }else showDashboard(false);
 });
 const initial=new URLSearchParams(location.hash.replace(/^#/,'' )).get('page');
 if(initial){
   const parsed=new URL(decodeURIComponent(initial),location.origin);
   if(routes[parsed.pathname])openInsideDashboard(parsed.href,routes[parsed.pathname],null,false);
 }
})();
</script>

    <!-- Weekly Off Calendar Modal -->
    <div id="calendar-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center hidden">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-6 w-full max-w-5xl mx-4 space-y-6 max-h-[85vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-4">
                <h3 class="text-lg font-bold text-slate-900 flex items-center gap-2">📅 All Employees Weekly Off Rota Calendar</h3>
                <button onclick="closeCalendarModal()" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            <div class="overflow-y-auto flex-1">
                <form method="GET" action="/" class="mb-4 rounded-2xl border border-slate-200 bg-slate-50 p-4">
                    <div class="flex flex-col md:flex-row md:items-end gap-3">
                        <div><label class="block text-[11px] font-bold uppercase text-slate-500 mb-1">Dashboard Date</label><input type="date" name="rota_date" value="{{ rota_dashboard_date }}" class="border border-slate-300 rounded-xl px-3 py-2 text-sm"></div>
                        <input type="hidden" name="start_date" value="{{ start_date }}"><input type="hidden" name="end_date" value="{{ end_date }}">
                        <button class="bg-indigo-600 text-white font-bold rounded-xl px-5 py-2">Show Rota Dashboard</button>
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-3 mt-4">
                        <div class="rounded-xl bg-blue-50 border border-blue-200 p-3"><p class="text-[10px] font-bold uppercase text-blue-600">Shift A</p><p class="text-2xl font-black text-blue-800">{{ rota_shift_counts['Shift A'] }}</p><p class="text-xs text-blue-600">employees on {{ rota_dashboard_date }}</p></div>
                        <div class="rounded-xl bg-violet-50 border border-violet-200 p-3"><p class="text-[10px] font-bold uppercase text-violet-600">Shift B</p><p class="text-2xl font-black text-violet-800">{{ rota_shift_counts['Shift B'] }}</p><p class="text-xs text-violet-600">employees on {{ rota_dashboard_date }}</p></div>
                        <div class="rounded-xl bg-emerald-50 border border-emerald-200 p-3"><p class="text-[10px] font-bold uppercase text-emerald-600">Weekly Off</p><p class="text-2xl font-black text-emerald-800">{{ rota_shift_counts['Weekly Off'] }}</p><p class="text-xs text-emerald-600">employees on {{ rota_dashboard_date }}</p></div>
                    </div>
                    <div class="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2 mt-3">{% for day,count in weekly_off_by_day.items() %}<div class="rounded-lg bg-white border border-slate-200 p-2 text-center"><p class="text-[10px] font-bold text-slate-500">{{ day[:3] }}</p><p class="text-lg font-black text-indigo-700">{{ count }}</p></div>{% endfor %}</div>
                    <div class="mt-3 rounded-xl bg-white border border-slate-200 p-3"><p class="text-xs font-bold text-slate-700 mb-2">Weekly Off employees on {{ rota_dashboard_date }}</p><div class="flex flex-wrap gap-2">{% for emp in rota_off_employees %}<span class="px-2 py-1 rounded-lg bg-emerald-50 text-emerald-700 text-[10px] font-bold">{{ emp.name }} · {{ emp.dept }}</span>{% else %}<span class="text-xs text-slate-400">No Weekly Off employees</span>{% endfor %}</div></div>
                </form>
                <p class="text-xs text-slate-500 mb-3">Yahan sabhi employees ke designated Weekly Off days ki complete list di gayi hai:</p>
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-100 text-slate-600 uppercase text-[11px] font-bold tracking-wider">
                            <th class="py-3 px-4 border border-slate-200">Sr.</th>
                            <th class="py-3 px-4 border border-slate-200">ID</th>
                            <th class="py-3 px-4 border border-slate-200">Employee Name</th>
                            <th class="py-3 px-4 border border-slate-200">Department</th>
                            <th class="py-3 px-4 text-center border border-slate-200 bg-indigo-50 text-indigo-800">Weekly Off Day</th>
                        </tr>
                    </thead>
                    <tbody class="text-xs text-slate-700 divide-y divide-slate-100">
                        {% for emp in all_users %}
                        <tr class="hover:bg-slate-50">
                            <td class="py-2.5 px-4 border border-slate-200 text-slate-400 font-medium">{{ loop.index }}</td>
                            <td class="py-2.5 px-4 border border-slate-200 font-mono text-slate-600 font-semibold">{{ emp.user_id }}</td>
                            <td class="py-2.5 px-4 border border-slate-200 font-bold text-slate-900">{{ emp.name }}</td>
                            <td class="py-2.5 px-4 border border-slate-200"><span class="px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold">{{ emp.dept }}</span></td>
                            <td class="py-2.5 px-4 border border-slate-200 text-center font-bold text-indigo-700 bg-indigo-50/50 text-sm">{{ emp.off }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
            <div class="pt-2 flex justify-end">
                <button onclick="closeCalendarModal()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition">Close</button>
            </div>
        </div>
    </div>
</body>
</html>
"""

MONTHLY_ROSTER_TEMPLATE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>AI Monthly Roster</title>
<style id="gamek-neo-ui">
:root{--neo:#07111f;--neo2:#0d1b2e;--cyan:#22d3ee;--mint:#34d399;--violet:#8b5cf6;--line:rgba(148,163,184,.18);--glass:rgba(255,255,255,.72)}
html{scroll-behavior:smooth}body{background:radial-gradient(circle at 8% 4%,rgba(52,211,153,.20),transparent 25%),radial-gradient(circle at 92% 5%,rgba(59,130,246,.20),transparent 28%),linear-gradient(145deg,#f8fafc,#eef4ff)!important;color:#142033}
body:before{content:'';position:fixed;inset:0;pointer-events:none;z-index:-1;background-image:linear-gradient(rgba(15,23,42,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(15,23,42,.025) 1px,transparent 1px);background-size:42px 42px;mask-image:linear-gradient(to bottom,black,transparent 82%)}
.neo-shell{border:1px solid rgba(255,255,255,.75);background:linear-gradient(145deg,rgba(255,255,255,.86),rgba(255,255,255,.62));backdrop-filter:blur(24px);box-shadow:0 22px 70px rgba(15,23,42,.10)}
header,.card,.bg-white{backdrop-filter:blur(20px)}
button,a,input,select{transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease}button:hover,a:hover{transform:translateY(-2px)}button:active,a:active{transform:scale(.97)}
input:focus,select:focus{outline:none!important;border-color:var(--cyan)!important;box-shadow:0 0 0 4px rgba(34,211,238,.13)!important}
::-webkit-scrollbar{width:10px;height:10px}::-webkit-scrollbar-track{background:#e8eef7}::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#34d399,#3b82f6);border-radius:20px;border:2px solid #e8eef7}
table th{letter-spacing:.035em}tbody tr{transition:background .18s ease,transform .18s ease}tbody tr:hover{background:rgba(224,242,254,.48)!important}
@keyframes neoIn{from{opacity:0;transform:translateY(12px) scale(.99)}to{opacity:1;transform:none}}main>*,body>main>*{animation:neoIn .45s ease both}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;scroll-behavior:auto!important}}
</style>
<script src="https://cdn.tailwindcss.com"></script>
<style>
:root{--glass:rgba(255,255,255,.78);--line:rgba(148,163,184,.25)}*{box-sizing:border-box}body{background:radial-gradient(circle at 10% 10%,#d1fae5 0,transparent 28%),radial-gradient(circle at 90% 15%,#dbeafe 0,transparent 30%),linear-gradient(135deg,#f8fafc,#eef2ff);min-height:100vh}
.card{background:var(--glass);backdrop-filter:blur(18px);border:1px solid var(--line);box-shadow:0 20px 60px rgba(15,23,42,.08)}.lift{transition:transform .25s,box-shadow .25s}.lift:hover{transform:translateY(-3px);box-shadow:0 18px 38px rgba(15,23,42,.13)}
.fade{animation:fadeUp .55s ease both}.pulse-soft{animation:pulseSoft 2.4s ease-in-out infinite}.shine{position:relative;overflow:hidden}.shine:after{content:'';position:absolute;inset:-100% auto -100% -40%;width:25%;transform:rotate(18deg);background:linear-gradient(90deg,transparent,rgba(255,255,255,.5),transparent);animation:shine 4s infinite}
.roster-select{transition:.2s;border:1px solid #cbd5e1}.roster-select:focus{transform:scale(1.04);box-shadow:0 0 0 4px rgba(99,102,241,.15)}.sticky-name{position:sticky;left:0;z-index:5}.week-a{background:#eff6ff}.week-b{background:#faf5ff}.off-cell{background:#ecfdf5!important}.today-col{box-shadow:inset 0 3px #f59e0b} .employee-row[hidden]{display:none!important} #designationMenu.hidden{display:none!important}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}@keyframes pulseSoft{50%{transform:scale(1.025)}}@keyframes shine{0%,65%{left:-50%}100%{left:140%}}@media(prefers-reduced-motion:reduce){*,*:before,*:after{animation:none!important;transition:none!important}}
</style></head>
<body class="text-slate-800"><main class="mx-auto max-w-[1900px] p-3 md:p-7 space-y-5">
<header class="card fade rounded-3xl p-5 md:p-7 flex flex-wrap justify-between gap-4 items-center"><div><div class="inline-flex items-center gap-2 rounded-full bg-emerald-100 text-emerald-800 px-3 py-1 text-xs font-black">✦ AI WORKFORCE PLANNER</div><h1 class="mt-3 text-3xl md:text-4xl font-black tracking-tight text-slate-950">Monthly Date-wise Roster</h1><p class="mt-2 text-sm text-slate-600">Monday se week start. Week 1 Shift A ho to Week 2 Shift B, phir poore month alternate. Weekly off har week same weekday par.</p></div><a href="/" class="lift rounded-2xl border bg-white/80 px-5 py-3 text-sm font-bold">← Dashboard</a></header>
{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="fade rounded-2xl border border-emerald-200 bg-emerald-50 p-4 font-bold text-emerald-800">{{ message }}</div>{% endfor %}{% endwith %}
<section class="card fade rounded-3xl p-4 md:p-5"><form method="GET" action="/monthly_roster" class="flex flex-wrap items-end gap-3"><div><label class="block text-[11px] font-black uppercase text-slate-500 mb-1">Roster month</label><input name="month" type="month" value="{{ month }}" class="rounded-xl border bg-white px-4 py-3 font-bold"></div><button class="lift rounded-xl bg-slate-900 px-5 py-3 text-sm font-bold text-white">Load Month</button></form>
<form method="POST" action="/monthly_roster/auto_generate" class="mt-3 flex flex-wrap items-end gap-3"><input type="hidden" name="month" value="{{ month }}"><button class="shine pulse-soft rounded-xl bg-gradient-to-r from-indigo-600 to-violet-600 px-5 py-3 text-sm font-black text-white">✦ AI Auto Generate</button><span class="text-xs text-slate-500">Existing month data regenerate hoga. Baad mein har date manually edit kar sakte hain.</span></form></section>
<section class="card fade rounded-3xl p-4 md:p-5"><div class="flex flex-wrap items-center justify-between gap-3"><div><p class="text-[11px] font-black uppercase tracking-wider text-indigo-600">Daily manpower snapshot</p><h2 class="text-lg font-black text-slate-950">Day-wise Shift & Designation Summary</h2></div><div class="flex gap-2"><button type="button" id="summaryPrev" class="lift rounded-xl border bg-white px-3 py-2 font-black">‹</button><select id="summaryDate" class="rounded-xl border bg-white px-3 py-2 font-bold">{% for d in dates %}<option value="{{ d.iso }}">{{ d.date_label }} · {{ d.day }}</option>{% endfor %}</select><button type="button" id="summaryNext" class="lift rounded-xl border bg-white px-3 py-2 font-black">›</button></div></div>
<div class="mt-4 grid grid-cols-3 gap-2 md:max-w-xl"><div class="rounded-2xl border border-blue-200 bg-blue-50 p-3"><p class="text-[10px] font-black uppercase text-blue-600">Shift A</p><p id="sumA" class="text-2xl font-black text-blue-900">0</p></div><div class="rounded-2xl border border-violet-200 bg-violet-50 p-3"><p class="text-[10px] font-black uppercase text-violet-600">Shift B</p><p id="sumB" class="text-2xl font-black text-violet-900">0</p></div><div class="rounded-2xl border border-emerald-200 bg-emerald-50 p-3"><p class="text-[10px] font-black uppercase text-emerald-600">Off</p><p id="sumOff" class="text-2xl font-black text-emerald-900">0</p></div></div>
<div class="mt-4 overflow-auto"><table class="w-full min-w-[620px] text-xs"><thead><tr class="bg-slate-900 text-white"><th class="rounded-l-xl p-3 text-left">Designation / Department</th><th class="p-3">Shift A</th><th class="p-3">Shift B</th><th class="p-3">Off</th><th class="rounded-r-xl p-3">Total</th></tr></thead><tbody id="designationSummary"></tbody></table></div></section>
<section id="rosterFilterPanel" class="card fade relative z-[80] overflow-visible rounded-3xl p-4"><div class="flex flex-wrap items-end gap-3"><div class="min-w-64"><label class="mb-1 block text-[11px] font-black uppercase text-slate-500">Search employee</label><input id="employeeSearch" type="search" placeholder="Name or employee code" class="w-full rounded-xl border px-3 py-2.5"></div><div><label class="mb-1 block text-[11px] font-black uppercase text-slate-500">Sort by designation</label><select id="designationSort" class="rounded-xl border px-3 py-2.5 font-bold"><option value="az">Designation A–Z</option><option value="za">Designation Z–A</option><option value="name">Employee Name A–Z</option></select></div><div class="relative"><label class="mb-1 block text-[11px] font-black uppercase text-slate-500">Designation filter</label><button id="designationToggle" type="button" class="rounded-xl border bg-white px-4 py-2.5 font-bold">All designations ▾</button><div id="designationMenu" class="fixed z-[9999] hidden max-h-[360px] min-w-72 overflow-auto rounded-2xl border border-slate-200 bg-white p-3 shadow-[0_24px_70px_rgba(15,23,42,.28)]"><div class="mb-2 flex gap-2"><button id="selectAllDesignations" type="button" class="rounded-lg bg-slate-900 px-3 py-1.5 text-xs font-bold text-white">Select all</button><button id="clearDesignations" type="button" class="rounded-lg border px-3 py-1.5 text-xs font-bold">Clear</button></div>{% for designation in designations %}<label class="flex cursor-pointer items-center gap-2 rounded-lg px-2 py-2 hover:bg-slate-50"><input type="checkbox" class="designation-filter" value="{{ designation }}" checked><span class="font-bold">{{ designation }}</span></label>{% endfor %}</div></div><div class="pb-1 text-xs font-bold text-slate-500"><span id="visibleEmployeeCount">{{ employees|length }}</span> employee visible</div></div></section>
<form method="POST" action="/monthly_roster" class="space-y-4"><input type="hidden" name="month" value="{{ month }}">
<div class="card fade relative z-0 rounded-3xl overflow-hidden"><div id="rosterTableViewport" class="relative overflow-auto h-[62vh] min-h-[420px] max-h-[70vh]"><div id="rosterEmptyState" class="pointer-events-none absolute inset-x-0 top-24 z-10 hidden items-center justify-center p-8"><div class="rounded-2xl border border-slate-200 bg-white/95 px-8 py-6 text-center shadow-xl backdrop-blur"><p class="text-2xl">⌕</p><p class="mt-2 font-black text-slate-800">No designation selected</p><p class="mt-1 text-xs text-slate-500">Select one or more designations to show employees.</p></div></div><table class="border-collapse min-w-max text-xs"><thead class="sticky top-0 z-20"><tr class="bg-slate-900 text-white"><th class="sticky-name bg-slate-900 min-w-72 p-3 text-left">Employee / Weekly Off</th>{% for w in week_groups %}<th colspan="{{ w.dates|length }}" class="border-l border-slate-700 p-2 text-center">Week {{ w.number }} · {{ w.shift_label }}</th>{% endfor %}</tr><tr class="bg-slate-800 text-white"><th class="sticky-name bg-slate-800 p-3 text-left">Monday-based calendar</th>{% for d in dates %}<th class="min-w-24 border-l border-slate-700 p-2 text-center {% if d.is_today %}today-col{% endif %}">{{ d.day }}<br><b>{{ d.date_label }}</b></th>{% endfor %}</tr></thead>
<tbody id="rosterBody">{% for emp in employees %}<tr class="employee-row border-b border-slate-200 hover:bg-white/60 transition" data-name="{{ emp.name|lower }}" data-code="{{ emp.user_id|lower }}" data-designation="{{ emp.dept }}"><th class="sticky-name bg-white/95 p-3 text-left"><span class="font-mono text-[10px] text-slate-400">{{ emp.user_id }}</span><span class="block font-black text-slate-950">{{ emp.name }} <span class="ml-1 inline-flex rounded-full bg-indigo-50 px-2 py-0.5 text-[9px] font-black text-indigo-700">{{ emp.dept }}</span></span><select name="off_{{ emp.user_id }}" class="weekly-off mt-2 w-full rounded-lg border px-2 py-1.5 font-bold text-emerald-700" data-employee="{{ emp.user_id }}">{% for day in weekdays %}<option value="{{ day.upper() }}" {% if roster[emp.user_id].weekly_off==day.upper() %}selected{% endif %}>Off: {{ day }}</option>{% endfor %}</select></th>
{% for d in dates %}{% set value=roster[emp.user_id].dates[d.iso] %}<td class="border-l border-slate-200 p-1.5 {% if d.week_index%2==0 %}week-a{% else %}week-b{% endif %} {% if value=='Off' %}off-cell{% endif %}"><select name="date_{{ emp.user_id }}_{{ d.iso }}" data-employee="{{ emp.user_id }}" data-designation="{{ emp.dept }}" data-date="{{ d.iso }}" data-weekday="{{ d.day_upper }}" class="roster-select w-full rounded-lg bg-white/90 px-1.5 py-2 font-bold"><option {% if value=='Shift A' %}selected{% endif %}>Shift A</option><option {% if value=='Shift B' %}selected{% endif %}>Shift B</option><option {% if value=='Off' %}selected{% endif %}>Off</option></select></td>{% endfor %}</tr>{% endfor %}</tbody></table></div></div>
<div class="sticky bottom-3 z-30 card rounded-2xl p-3 flex flex-wrap justify-between items-center gap-3"><div class="text-xs text-slate-600"><b>Blue</b> odd weeks · <b>Violet</b> even weeks · <b>Green</b> weekly off</div><div class="flex gap-2"><a href="/monthly_roster/export?month={{ month }}" class="lift rounded-xl border bg-white px-5 py-3 font-bold">Download Excel</a><button class="shine lift rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 px-6 py-3 font-black text-white">Save Complete Month</button></div></div></form></main>
<script>
document.querySelectorAll('.weekly-off').forEach(off=>off.addEventListener('change',()=>{const emp=off.dataset.employee, day=off.value;document.querySelectorAll(`select[data-employee="${emp}"][data-weekday]`).forEach(sel=>{const td=sel.closest('td');if(sel.dataset.weekday===day){sel.value='Off';td.classList.add('off-cell')}else if(sel.value==='Off'){const all=[...document.querySelectorAll(`select[data-employee="${emp}"][data-weekday]`)];const idx=all.indexOf(sel), mondayIndex=Math.floor(idx/7);sel.value=mondayIndex%2===0?'Shift A':'Shift B';td.classList.remove('off-cell')}})}));
document.querySelectorAll('.roster-select').forEach(sel=>sel.addEventListener('change',()=>sel.closest('td').classList.toggle('off-cell',sel.value==='Off')));

function refreshSummary(){
 const date=document.getElementById('summaryDate').value, counts={A:0,B:0,Off:0}, groups={};
 document.querySelectorAll(`.roster-select[data-date="${date}"]`).forEach(sel=>{const key=sel.value==='Shift A'?'A':sel.value==='Shift B'?'B':'Off', dept=sel.dataset.designation||'Other';counts[key]++;if(!groups[dept])groups[dept]={A:0,B:0,Off:0};groups[dept][key]++;});
 document.getElementById('sumA').textContent=counts.A;document.getElementById('sumB').textContent=counts.B;document.getElementById('sumOff').textContent=counts.Off;
 document.getElementById('designationSummary').innerHTML=Object.entries(groups).sort().map(([name,c])=>`<tr class="border-b border-slate-200 hover:bg-white/70"><th class="p-3 text-left font-black">${name}</th><td class="p-3 text-center font-black text-blue-700">${c.A}</td><td class="p-3 text-center font-black text-violet-700">${c.B}</td><td class="p-3 text-center font-black text-emerald-700">${c.Off}</td><td class="p-3 text-center font-black">${c.A+c.B+c.Off}</td></tr>`).join('');
}
const summaryDate=document.getElementById('summaryDate');summaryDate.addEventListener('change',refreshSummary);document.getElementById('summaryPrev').addEventListener('click',()=>{summaryDate.selectedIndex=Math.max(0,summaryDate.selectedIndex-1);refreshSummary()});document.getElementById('summaryNext').addEventListener('click',()=>{summaryDate.selectedIndex=Math.min(summaryDate.options.length-1,summaryDate.selectedIndex+1);refreshSummary()});
document.querySelectorAll('.roster-select,.weekly-off').forEach(x=>x.addEventListener('change',()=>setTimeout(refreshSummary,0)));refreshSummary();
const rosterBody=document.getElementById('rosterBody');
if(rosterBody){
const employeeSearch=document.getElementById('employeeSearch'), designationSort=document.getElementById('designationSort'), designationToggle=document.getElementById('designationToggle'), designationMenu=document.getElementById('designationMenu');
function selectedDesignations(){return new Set([...document.querySelectorAll('.designation-filter:checked')].map(x=>x.value))}
function applyRosterView(){const q=employeeSearch.value.trim().toLowerCase(), chosen=selectedDesignations(), rows=[...document.querySelectorAll('.employee-row')];rows.forEach(row=>{row.hidden=!chosen.has(row.dataset.designation)||!(row.dataset.name.includes(q)||row.dataset.code.includes(q))});const mode=designationSort.value;rows.sort((a,b)=>{if(mode==='name')return a.dataset.name.localeCompare(b.dataset.name);const cmp=a.dataset.designation.localeCompare(b.dataset.designation)||a.dataset.name.localeCompare(b.dataset.name);return mode==='za'?-cmp:cmp});rows.forEach(row=>rosterBody.appendChild(row));const visible=rows.filter(r=>!r.hidden).length;document.getElementById('visibleEmployeeCount').textContent=visible;const empty=document.getElementById('rosterEmptyState');empty.classList.toggle('hidden',visible!==0);empty.classList.toggle('flex',visible===0);designationToggle.textContent=chosen.size===document.querySelectorAll('.designation-filter').length?'All designations ▾':`${chosen.size} selected ▾`;}
function positionDesignationMenu(){
 const r=designationToggle.getBoundingClientRect(), gap=8, menuH=Math.min(360,designationMenu.scrollHeight||360), below=window.innerHeight-r.bottom-gap, above=r.top-gap;
 designationMenu.style.left=Math.max(8,Math.min(r.left,window.innerWidth-304))+'px';
 designationMenu.style.width=Math.max(288,r.width)+'px';
 if(below>=Math.min(menuH,220)||below>=above){designationMenu.style.top=(r.bottom+gap)+'px';designationMenu.style.bottom='auto';designationMenu.style.maxHeight=Math.max(160,below-8)+'px'}else{designationMenu.style.top='auto';designationMenu.style.bottom=(window.innerHeight-r.top+gap)+'px';designationMenu.style.maxHeight=Math.max(160,above-8)+'px'}
}
designationToggle.addEventListener('click',()=>{designationMenu.classList.toggle('hidden');if(!designationMenu.classList.contains('hidden')){positionDesignationMenu()}});
window.addEventListener('resize',()=>{if(!designationMenu.classList.contains('hidden'))positionDesignationMenu()});
employeeSearch.addEventListener('input',applyRosterView);designationSort.addEventListener('change',applyRosterView);document.querySelectorAll('.designation-filter').forEach(x=>x.addEventListener('change',applyRosterView));document.getElementById('selectAllDesignations').addEventListener('click',()=>{document.querySelectorAll('.designation-filter').forEach(x=>x.checked=true);applyRosterView();designationMenu.classList.add('hidden')});document.getElementById('clearDesignations').addEventListener('click',()=>{document.querySelectorAll('.designation-filter').forEach(x=>x.checked=false);applyRosterView();designationMenu.classList.add('hidden')});document.addEventListener('click',e=>{if(!designationMenu.contains(e.target)&&e.target!==designationToggle)designationMenu.classList.add('hidden')});applyRosterView();

}
</script></body></html>
"""

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        uid = (request.form.get('user_id') or request.form.get('username') or '').strip().upper()
        pwd = (request.form.get('password') or '').strip()
        
        if not uid or not pwd:
            return render_template_string(LOGIN_TEMPLATE, error='User ID aur password dono enter karein.')

        admin_pass = os.getenv('ADMIN_PWD', 'Gamek@789')
        dev_pass = os.getenv('DEV_PWD', 'Shama@8577')
        
        if uid == 'LM11' and pwd == admin_pass:
            session['logged_in'] = True
            session['role'] = 'admin'
            session['user_id'] = 'LM11'
            session['user_name'] = 'Admin (LM11)'
            return redirect(url_for('index'))
            
        if uid == 'NCSA0608' and pwd == dev_pass:
            session['logged_in'] = True
            session['role'] = 'developer'
            session['user_id'] = 'NCSA0608'
            session['user_name'] = 'Sonu Kumar (Developer)'
            return redirect(url_for('index'))
            
        emp_key = f"NWC{uid}" if not uid.startswith('NWC') else uid
        if emp_key in MASTER_EMPLOYEES and pwd == '123':
            session['logged_in'] = True
            session['role'] = 'employee'
            session['user_id'] = emp_key
            session['user_name'] = MASTER_EMPLOYEES[emp_key]['name']
            return redirect(url_for('index'))
        else:
            return render_template_string(LOGIN_TEMPLATE, error="Galat User ID ya Password!")
    return render_template_string(LOGIN_TEMPLATE, error=None)

@app.route('/logout')
def logout():
    role = session.get('role')
    user_name = session.get('user_name', '')
    
    msg_text = "Thank you admin" if role == 'admin' else f"Thank you {user_name}"
    session.clear()
    flash(msg_text, 'success')
    return redirect(url_for('login'))

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    role = session.get('role')
    logged_user_id = session.get('user_id')
    logged_user_name = session.get('user_name')
    
    if role == 'employee':
        selected_emp = logged_user_id
    else:
        selected_emp = request.args.get('employee', 'ALL')
        
    today_str = datetime.now().strftime('%Y-%m-%d')
    start_date = request.args.get('start_date', today_str)
    end_date = request.args.get('end_date', today_str)
    
    selected_store = request.args.get('store', '').strip().upper() if role == 'developer' else DEFAULT_STORE_ID
    if role == 'developer' and not selected_store:
        logs, all_users, raw_punches = [], [], []
        g_hrs, g_l_hrs, g_var = '0h 0m', '0h 0m', '0h 0m'
        stats = {'present':0,'absent':0,'off':0,'mispunch':0,'late_arrival':0,'ml':0,'shift_a':0,'shift_b':0,'total_hrs':'0h 0m','total_lunch_hrs':'0h 0m','total_variance':'0h 0m','variance_type':'neutral','device_online':False}
    else:
        logs, all_users, g_hrs, g_l_hrs, g_var, raw_punches, stats = fetch_attendance_data(start_date, end_date, selected_emp, selected_store)
        allowed_codes = store_employees(selected_store)
        logs = [x for x in logs if x.get('user_id') in allowed_codes]
        all_users = [x for x in all_users if x.get('user_id') in allowed_codes]
        raw_punches = [x for x in raw_punches if x.get('user_id') in allowed_codes]
    
    pending_leaves_count = sum(1 for req in LEAVE_REQUESTS if req['status'] == 'Pending')
    pending_off_changes_count = sum(1 for req in OFF_CHANGE_REQUESTS if req.get('status') == 'Pending')
    pending_shift_changes_count = sum(1 for req in SHIFT_CHANGE_REQUESTS if req.get('status') == 'Pending')
    
    if role == 'employee':
        current_user_leave_requests = [req for req in LEAVE_REQUESTS if req['user_id'] == logged_user_id]
        current_user_off_change_requests = [req for req in OFF_CHANGE_REQUESTS if req.get('user_id') == logged_user_id or req.get('partner_id') == logged_user_id]
        current_user_shift_change_requests = [req for req in SHIFT_CHANGE_REQUESTS if req.get('requester_id') == logged_user_id or req.get('partner_id') == logged_user_id]
    else:
        current_user_leave_requests = LEAVE_REQUESTS
        current_user_off_change_requests = OFF_CHANGE_REQUESTS
        current_user_shift_change_requests = SHIFT_CHANGE_REQUESTS

    dashboard_date = request.args.get('rota_date', today_str)
    try:
        dashboard_dt = datetime.strptime(dashboard_date, '%Y-%m-%d')
    except ValueError:
        dashboard_date = today_str
        dashboard_dt = datetime.strptime(today_str, '%Y-%m-%d')
    weekly_off_by_day = {day: 0 for day in ['MONDAY','TUESDAY','WEDNESDAY','THURSDAY','FRIDAY','SATURDAY','SUNDAY']}
    shift_counts = {'Shift A': 0, 'Shift B': 0, 'Weekly Off': 0}
    selected_date_off_employees = []
    for emp_code in MASTER_EMPLOYEES:
        code = normalized_code(emp_code)
        info = get_emp_info(code)
        base_off = str(info.get('off', '')).upper()
        if base_off in weekly_off_by_day:
            weekly_off_by_day[base_off] += 1
        assignment = effective_roster_assignment(code, dashboard_date)
        off_change = get_approved_off_change(code, dashboard_date)
        is_off = assignment == 'Off' if assignment else dashboard_dt.strftime('%A').upper() == base_off
        if off_change == 'WORKING': is_off = False
        elif off_change == 'OFF': is_off = True
        if is_off:
            shift_counts['Weekly Off'] += 1
            selected_date_off_employees.append({'code': code, 'name': info.get('name', code), 'dept': info.get('dept', 'N/A')})
        elif assignment == 'Shift A':
            shift_counts['Shift A'] += 1
        else:
            shift_counts['Shift B'] += 1
    selected_date_off_employees.sort(key=lambda x: x['name'])
    
    return render_template_string(
        HTML_TEMPLATE,
        logs=logs,
        all_users=all_users,
        start_date=start_date,
        end_date=end_date,
        selected_store=selected_store,
        stores=sorted(STORES.items()),
        selected_emp=selected_emp,
        grand_total_hours=g_hrs,
        grand_total_lunch_hours=g_l_hrs,
        grand_total_variance=g_var,
        stats=stats,
        raw_punches=raw_punches,
        role=role,
        logged_user_name=logged_user_name,
        logged_user_id=logged_user_id,
        leave_requests=current_user_leave_requests,
        pending_leaves_count=pending_leaves_count,
        off_change_requests=current_user_off_change_requests,
        pending_off_changes_count=pending_off_changes_count,
        shift_change_requests=current_user_shift_change_requests,
        pending_shift_changes_count=pending_shift_changes_count,
        rota_dashboard_date=dashboard_date,
        rota_shift_counts=shift_counts,
        weekly_off_by_day=weekly_off_by_day,
        rota_off_employees=selected_date_off_employees,
        swap_partners=[{'code': normalized_code(c), 'name': get_emp_info(c)['name'], 'shift': ('Shift A' if get_emp_info(c).get('shift') == 'morning' else 'Shift B')} for c in MASTER_EMPLOYEES]
    )


PERFORMANCE_TEMPLATE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Monthly Employee Performance</title><script src="https://cdn.tailwindcss.com"></script>
<style>
body{background:radial-gradient(circle at 8% 5%,rgba(52,211,153,.18),transparent 25%),radial-gradient(circle at 92% 8%,rgba(59,130,246,.18),transparent 28%),linear-gradient(145deg,#f8fafc,#eef4ff);min-height:100vh}.glass{background:rgba(255,255,255,.76);border:1px solid rgba(255,255,255,.8);backdrop-filter:blur(22px);box-shadow:0 22px 65px rgba(15,23,42,.09)}.metric{transition:.2s}.metric:hover{transform:translateY(-3px);box-shadow:0 12px 28px rgba(15,23,42,.12)}#detailModal{backdrop-filter:blur(10px)}tbody tr{transition:.18s}tbody tr:hover{background:#f0fdfa}
</style></head><body class="text-slate-800"><main class="mx-auto max-w-[1900px] p-4 md:p-7 space-y-5">
<header class="glass rounded-3xl p-5 md:p-7 flex flex-wrap justify-between gap-4"><div><div class="inline-flex rounded-full bg-cyan-100 px-3 py-1 text-xs font-black text-cyan-800">AI PERFORMANCE INTELLIGENCE</div><h1 class="mt-3 text-3xl font-black">Monthly Employee Performance</h1><p class="mt-1 text-sm text-slate-500">Monday-Sunday weekly-off adjustment, shift-specific punctuality, justified/unjustified absence and achievement scoring.</p><p class="mt-2 inline-flex rounded-full border border-amber-200 bg-amber-50 px-3 py-1 text-[11px] font-black text-amber-700">Calculated only through {{ as_of_date }} · Future dates excluded</p></div><a href="/" class="rounded-xl border bg-white px-5 py-3 text-sm font-black">← Command Center</a></header>
<section class="glass rounded-3xl p-5"><form class="flex flex-wrap items-end gap-3"><div><label class="block text-[11px] font-black uppercase text-slate-500 mb-1">Performance month</label><input type="month" name="month" value="{{ month }}" class="rounded-xl border bg-white px-4 py-3 font-bold"></div><div><label class="block text-[11px] font-black uppercase text-slate-500 mb-1">Employee</label><select name="employee" class="min-w-72 rounded-xl border bg-white px-4 py-3 font-bold"><option value="ALL">All Employees</option>{% for e in employees %}<option value="{{ e.user_id }}" {% if selected_emp==e.user_id %}selected{% endif %}>{{ e.name }} · {{ e.dept }}</option>{% endfor %}</select></div><button class="rounded-xl bg-slate-950 px-6 py-3 font-black text-white">Open Performance</button></form></section>
<section class="grid grid-cols-2 md:grid-cols-5 gap-3">{% for label,value,color in overview %}<div class="glass rounded-2xl p-4 border-l-4 {{ color }}"><p class="text-[10px] font-black uppercase text-slate-500">{{ label }}</p><p class="mt-1 text-2xl font-black">{{ value }}</p></div>{% endfor %}</section>
<section class="glass rounded-3xl p-4"><div class="flex flex-wrap items-end gap-3"><div class="min-w-64"><label class="mb-1 block text-[10px] font-black uppercase text-slate-500">Search employee</label><input id="perfSearch" type="search" placeholder="Name, ID or designation" class="w-full rounded-xl border bg-white px-3 py-2.5"></div><div><label class="mb-1 block text-[10px] font-black uppercase text-slate-500">Designation</label><select id="perfDesignation" class="rounded-xl border bg-white px-3 py-2.5 font-bold"><option value="ALL">All Designations</option>{% for d in designations %}<option value="{{ d }}">{{ d }}</option>{% endfor %}</select></div><div><label class="mb-1 block text-[10px] font-black uppercase text-slate-500">Warning status</label><select id="perfWarning" class="rounded-xl border bg-white px-3 py-2.5 font-bold"><option value="ALL">All Employees</option><option value="WARNING">Warning Letter</option><option value="CLEAR">No Warning</option></select></div><div><label class="mb-1 block text-[10px] font-black uppercase text-slate-500">Sort performance</label><select id="perfSort" class="rounded-xl border bg-white px-3 py-2.5 font-bold"><option value="best">Best Employee First</option><option value="low">Lowest Percentage First</option><option value="name">Employee Name A-Z</option><option value="late">Most Late First</option><option value="absent">Most Unjustified First</option></select></div><div class="pb-2 text-xs font-black text-slate-500"><span id="perfVisible">{{ rows|length }}</span> employees</div></div></section>
<section class="glass rounded-3xl overflow-hidden"><div class="overflow-auto"><table class="w-full min-w-[1650px] text-xs"><thead class="sticky top-0 z-10"><tr class="bg-slate-950 text-white"><th class="p-4 text-left">Employee</th><th>Working Days</th><th>On Time</th><th>Late</th><th>Shift Mismatch</th><th>Overtime</th><th>Weekly Off</th><th>Off Not Taken</th><th>Justified Absent</th><th>Unjustified Absent</th><th>Best Employee<br><span class="text-[9px] font-medium text-slate-300">Good Achievement %</span></th></tr></thead><tbody>{% for r in rows %}<tr class="performance-row border-b border-slate-200" data-name="{{ r.name|lower }}" data-code="{{ r.user_id|lower }}" data-designation="{{ r.dept }}" data-warning="{% if r.warning %}WARNING{% else %}CLEAR{% endif %}" data-achievement="{{ r.achievement }}" data-late="{{ r.late }}" data-absent="{{ r.unjustified_absent }}"><td class="p-4"><div class="flex flex-wrap items-center gap-2"><b class="text-sm">{{ r.name }}</b>{% if r.warning %}<button type="button" onclick="openWarning('{{ r.user_id }}')" class="animate-pulse rounded-full border border-rose-300 bg-rose-100 px-2 py-1 text-[9px] font-black text-rose-700">⚠ WARNING LETTER</button>{% endif %}</div><span class="text-slate-400">{{ r.user_id }} · {{ r.dept }}</span></td><td class="text-center font-black">{{ r.working_days }}</td>
{% for key,color in [('on_time','emerald'),('late','amber'),('shift_mismatch','rose'),('overtime','violet'),('weekly_off','blue'),('off_not_taken','cyan'),('justified_absent','teal'),('unjustified_absent','red')] %}<td class="p-2 text-center"><button type="button" onclick="openMetric('{{ r.user_id }}','{{ key }}')" class="metric min-w-12 rounded-xl bg-{{ color }}-50 px-3 py-2 font-black text-{{ color }}-700 border border-{{ color }}-200">{{ r[key] }}</button></td>{% endfor %}
<td class="p-3"><div class="flex items-center gap-2"><div class="h-2 w-28 rounded-full bg-slate-200 overflow-hidden"><div class="h-full rounded-full {% if r.achievement >= 85 %}bg-gradient-to-r from-emerald-500 to-cyan-500{% elif r.achievement >= 70 %}bg-gradient-to-r from-amber-400 to-orange-500{% else %}bg-gradient-to-r from-rose-500 to-red-700{% endif %}" style="width:{{ r.achievement }}%"></div></div><div><b class="block">{{ r.achievement }}%</b><span class="text-[9px] font-black {% if r.achievement >= 90 %}text-emerald-700{% elif r.achievement >= 75 %}text-blue-700{% elif r.achievement >= 60 %}text-amber-700{% else %}text-rose-700{% endif %}">{% if r.achievement >= 90 %}BEST EMPLOYEE{% elif r.achievement >= 75 %}GOOD ACHIEVEMENT{% elif r.achievement >= 60 %}NEEDS IMPROVEMENT{% else %}LOW PERFORMANCE{% endif %}</span></div></div></td></tr>{% endfor %}</tbody></table></div></section>
</main>
<div id="detailModal" class="fixed inset-0 z-[9999] hidden items-center justify-center bg-slate-950/55 p-4"><div class="glass w-full max-w-5xl max-h-[82vh] overflow-hidden rounded-3xl"><header class="flex items-center justify-between border-b p-5"><div><p id="modalEmployee" class="text-xs font-black text-cyan-700"></p><h2 id="modalTitle" class="text-xl font-black"></h2></div><button type="button" onclick="closeMetric()" class="rounded-xl bg-slate-950 px-4 py-2 font-black text-white">Close</button></header><div class="max-h-[65vh] overflow-auto"><table class="w-full text-xs"><thead class="sticky top-0 bg-slate-100"><tr><th class="p-3 text-left">Date</th><th>Scheduled</th><th>Actual</th><th>Store In</th><th>Out</th><th>Hours</th><th>Extra</th><th>Finding</th></tr></thead><tbody id="modalRows"></tbody></table></div></div></div>
<script>
const performanceDetails={{ details|tojson }};const labels={on_time:'On-time attendance details',late:'Late arrival details',shift_mismatch:'Shift mismatch details',overtime:'Overtime details',weekly_off:'Weekly off details',off_not_taken:'Weekly off not taken details',justified_absent:'Justified absence details',unjustified_absent:'Unjustified absence details'};
function openMetric(uid,key){const set=(performanceDetails[uid]||{})[key]||[],emp=(performanceDetails[uid]||{})._employee||uid;document.getElementById('modalEmployee').textContent=emp;document.getElementById('modalTitle').textContent=`${labels[key]} · ${set.length} record(s)`;document.getElementById('modalRows').innerHTML=set.length?set.map(x=>`<tr class="border-b"><td class="p-3 font-bold">${x.date}</td><td class="text-center">${x.scheduled_shift}</td><td class="text-center font-bold">${x.shift_type}</td><td class="text-center">${x.store_in}</td><td class="text-center">${x.out_time}</td><td class="text-center">${x.total_hours}</td><td class="text-center font-bold">${x.net_variance}</td><td class="p-3 font-bold">${x.finding}</td></tr>`).join(''):'<tr><td colspan="8" class="p-10 text-center text-slate-400">No records for this category.</td></tr>';const m=document.getElementById('detailModal');m.classList.remove('hidden');m.classList.add('flex')}
function openWarning(uid){const emp=(performanceDetails[uid]||{})._employee||uid,count=((performanceDetails[uid]||{}).unjustified_absent||[]).length;document.getElementById('modalEmployee').textContent=emp;document.getElementById('modalTitle').textContent='Warning Letter Recommended';document.getElementById('modalRows').innerHTML=`<tr><td colspan="8" class="p-8"><div class="rounded-2xl border border-rose-200 bg-rose-50 p-6 text-center"><p class="text-3xl">⚠</p><p class="mt-2 text-lg font-black text-rose-800">${count} unjustified absences recorded</p><p class="mt-2 text-sm text-rose-700">Monthly unjustified absence 3 ya usse zyada hai. HR/Admin ko warning letter issue karna chahiye.</p></div></td></tr>`;const m=document.getElementById('detailModal');m.classList.remove('hidden');m.classList.add('flex')}
function closeMetric(){const m=document.getElementById('detailModal');m.classList.add('hidden');m.classList.remove('flex')}document.getElementById('detailModal').addEventListener('click',e=>{if(e.target.id==='detailModal')closeMetric()});

const perfBody=document.querySelector('.performance-row')?.parentElement,perfSearch=document.getElementById('perfSearch'),perfDesignation=document.getElementById('perfDesignation'),perfWarning=document.getElementById('perfWarning'),perfSort=document.getElementById('perfSort');
function applyPerformanceView(){if(!perfBody)return;const q=perfSearch.value.trim().toLowerCase(),designation=perfDesignation.value,warning=perfWarning.value,rows=[...document.querySelectorAll('.performance-row')];rows.forEach(r=>{const text=`${r.dataset.name} ${r.dataset.code} ${r.dataset.designation.toLowerCase()}`;r.hidden=!text.includes(q)||(designation!=='ALL'&&r.dataset.designation!==designation)||(warning!=='ALL'&&r.dataset.warning!==warning)});const mode=perfSort.value;rows.sort((a,b)=>mode==='name'?a.dataset.name.localeCompare(b.dataset.name):mode==='low'?Number(a.dataset.achievement)-Number(b.dataset.achievement):mode==='late'?Number(b.dataset.late)-Number(a.dataset.late):mode==='absent'?Number(b.dataset.absent)-Number(a.dataset.absent):Number(b.dataset.achievement)-Number(a.dataset.achievement)||a.dataset.name.localeCompare(b.dataset.name));rows.forEach(r=>perfBody.appendChild(r));document.getElementById('perfVisible').textContent=rows.filter(r=>!r.hidden).length}
[perfSearch,perfDesignation,perfWarning,perfSort].forEach(x=>x&&x.addEventListener(x===perfSearch?'input':'change',applyPerformanceView));applyPerformanceView();
</script></body></html>
"""

@app.route('/performance')
def performance_dashboard():
    if not session.get('logged_in') or session.get('role') not in ['admin','developer']:
        return redirect(url_for('login'))
    month=request.args.get('month') or datetime.now().strftime('%Y-%m')
    selected_emp=request.args.get('employee','ALL')
    try:
        first=datetime.strptime(month+'-01','%Y-%m-%d')
    except ValueError:
        return 'Invalid month. Use YYYY-MM.',400
    next_month=(first+timedelta(days=32)).replace(day=1)
    today = datetime.now().date()
    month_last = (next_month-timedelta(days=1)).date()
    # Do not create Absent/Weekly Off/performance records for future dates.
    effective_end = min(month_last, today)
    start_date = first.strftime('%Y-%m-%d')
    if first.date() > today:
        end_date = start_date
        logs = []
        employees = [
            {'user_id': normalized_code(code), 'name': get_emp_info(normalized_code(code))['name'], 'dept': get_emp_info(normalized_code(code))['dept']}
            for code in sorted(MASTER_EMPLOYEES, key=lambda c: get_emp_info(c)['name'])
        ]
        if selected_emp != 'ALL':
            employees = [e for e in employees if e['user_id'] == selected_emp]
    else:
        end_date = effective_end.strftime('%Y-%m-%d')
        logs, employees, _, _, _, _, _ = fetch_attendance_data(start_date,end_date,selected_emp)
    by_emp={e['user_id']:[] for e in employees if selected_emp=='ALL' or e['user_id']==selected_emp}
    for x in logs:
        if x['user_id'] in by_emp: by_emp[x['user_id']].append(dict(x))
    bucket_names=['on_time','late','shift_mismatch','overtime','weekly_off','off_not_taken','justified_absent','unjustified_absent']
    rows=[]; details={}; totals={k:0 for k in bucket_names}
    def compact(x,finding):
        return {k:x.get(k,'-') for k in ['date','scheduled_shift','shift_type','store_in','out_time','total_hours','net_variance']} | {'finding':finding}
    def punch_minutes(value):
        try:
            t=datetime.strptime(str(value).split()[0],'%H:%M:%S')
            return t.hour*60+t.minute+(1 if t.second else 0)
        except (ValueError,TypeError,IndexError):
            return None
    def week_key(x):
        d=datetime.strptime(x['date'],'%Y-%m-%d').date()
        monday=d-timedelta(days=d.weekday())
        return monday.isoformat()
    for e in employees:
        uid=e['user_id']
        if uid not in by_emp: continue
        records=sorted(by_emp[uid],key=lambda x:x['date'])
        buckets={k:[] for k in bucket_names}; working=0
        # Approved leave is justified. Raw Absents are adjusted Monday-Sunday:
        # first absent becomes weekly off; remaining absent days are unjustified.
        weeks={}
        for x in records: weeks.setdefault(week_key(x),[]).append(x)
        converted_off_dates=set(); unjustified_dates=set()
        for week_records in weeks.values():
            absent_days=[x for x in week_records if x.get('status')=='Absent']
            if absent_days:
                converted_off_dates.add(absent_days[0]['date'])
                unjustified_dates.update(x['date'] for x in absent_days[1:])
        for x in records:
            status=x.get('status',''); present=x.get('store_in')!='-'; date=x.get('date')
            if status=='Weekly Off' or date in converted_off_dates:
                finding='Automatic Weekly Off (first absence of Monday-Sunday week)' if date in converted_off_dates else 'Roster Weekly Off'
                buckets['weekly_off'].append(compact(x,finding))
                if not present: continue
            if 'Leave' in status:
                buckets['justified_absent'].append(compact(x,status))
                continue
            if date in unjustified_dates:
                buckets['unjustified_absent'].append(compact(x,'Unjustified Absence'))
                continue
            if status=='Absent':
                # Defensive fallback for malformed dates/weeks.
                buckets['unjustified_absent'].append(compact(x,'Unjustified Absence'))
                continue
            if present and x.get('scheduled_shift')!='Off': working+=1
            mins=punch_minutes(x.get('store_in'))
            actual='Shift A' if mins is not None and mins<720 else 'Shift B'
            is_late=(actual=='Shift A' and mins is not None and mins>420) or (actual=='Shift B' and mins is not None and mins>790)
            is_on_time=(actual=='Shift A' and mins is not None and mins<=420) or (actual=='Shift B' and mins is not None and 720<=mins<=790)
            x['shift_type']=actual if present else x.get('shift_type','-')
            if present and is_on_time: buckets['on_time'].append(compact(x,'On Time'))
            if present and is_late: buckets['late'].append(compact(x,'Late Arrival'))
            scheduled=x.get('scheduled_shift')
            if present and scheduled in ('Shift A','Shift B') and actual!=scheduled:
                buckets['shift_mismatch'].append(compact(x,'Shift Mismatch'))
            if _extra_minutes(x.get('net_variance'))>0: buckets['overtime'].append(compact(x,'Overtime'))
            if x.get('scheduled_shift')=='Off' and present: buckets['off_not_taken'].append(compact(x,'Weekly Off Worked'))
        counts={k:len(v) for k,v in buckets.items()}
        for k,v in counts.items(): totals[k]+=v
        # Achievement rewards punctual work and removes points for discipline exceptions.
        achievement=max(0,min(100,round(100-(counts['late']*2+counts['shift_mismatch']*5+counts['unjustified_absent']*10+counts['off_not_taken']*2))))
        rows.append({'user_id':uid,'name':e['name'],'dept':e['dept'],'working_days':working,'achievement':achievement,'warning':counts['unjustified_absent']>=3,**counts})
        details[uid]={'_employee':f"{e['name']} · {uid} · {e['dept']}",**buckets}
    overview=[('Employees',len(rows),'border-l-cyan-500'),('On Time',totals['on_time'],'border-l-emerald-500'),('Late',totals['late'],'border-l-amber-500'),('Unjustified',totals['unjustified_absent'],'border-l-rose-500'),('Warning Letters',sum(1 for r in rows if r['warning']),'border-l-violet-500')]
    return render_template_string(PERFORMANCE_TEMPLATE,month=month,selected_emp=selected_emp,employees=employees,rows=rows,details=details,overview=overview,designations=sorted({e['dept'] for e in employees}),as_of_date=(effective_end if first.date() <= today else today).strftime('%d %b %Y'))

MATRIX_PREVIEW_TEMPLATE = """
<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Attendance Matrix Preview</title>
<style id="gamek-neo-ui">
:root{--neo:#07111f;--neo2:#0d1b2e;--cyan:#22d3ee;--mint:#34d399;--violet:#8b5cf6;--line:rgba(148,163,184,.18);--glass:rgba(255,255,255,.72)}
html{scroll-behavior:smooth}body{background:radial-gradient(circle at 8% 4%,rgba(52,211,153,.20),transparent 25%),radial-gradient(circle at 92% 5%,rgba(59,130,246,.20),transparent 28%),linear-gradient(145deg,#f8fafc,#eef4ff)!important;color:#142033}
body:before{content:'';position:fixed;inset:0;pointer-events:none;z-index:-1;background-image:linear-gradient(rgba(15,23,42,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(15,23,42,.025) 1px,transparent 1px);background-size:42px 42px;mask-image:linear-gradient(to bottom,black,transparent 82%)}
.neo-shell{border:1px solid rgba(255,255,255,.75);background:linear-gradient(145deg,rgba(255,255,255,.86),rgba(255,255,255,.62));backdrop-filter:blur(24px);box-shadow:0 22px 70px rgba(15,23,42,.10)}
header,.card,.bg-white{backdrop-filter:blur(20px)}
button,a,input,select{transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease}button:hover,a:hover{transform:translateY(-2px)}button:active,a:active{transform:scale(.97)}
input:focus,select:focus{outline:none!important;border-color:var(--cyan)!important;box-shadow:0 0 0 4px rgba(34,211,238,.13)!important}
::-webkit-scrollbar{width:10px;height:10px}::-webkit-scrollbar-track{background:#e8eef7}::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#34d399,#3b82f6);border-radius:20px;border:2px solid #e8eef7}
table th{letter-spacing:.035em}tbody tr{transition:background .18s ease,transform .18s ease}tbody tr:hover{background:rgba(224,242,254,.48)!important}
@keyframes neoIn{from{opacity:0;transform:translateY(12px) scale(.99)}to{opacity:1;transform:none}}main>*,body>main>*{animation:neoIn .45s ease both}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;scroll-behavior:auto!important}}
</style>
<script src="https://cdn.tailwindcss.com"></script>
<style>.sticky-col{position:sticky;left:0;z-index:2}.cell{min-width:92px}</style></head>
<body class="bg-slate-50 text-slate-800"><main class="mx-auto max-w-[1900px] p-4 md:p-7 space-y-5">
<header class="flex flex-wrap justify-between gap-3"><div><h1 class="text-2xl font-black">Attendance Matrix Preview</h1><p class="text-sm text-slate-500">Excel-style employee × date view with working hours and extra hours.</p></div><div class="flex gap-2"><a href="/" class="px-4 py-2 rounded-lg border bg-white font-bold text-sm">Dashboard</a><a href="/export_matrix?start_date={{ start_date }}&end_date={{ end_date }}&employee={{ selected_emp }}" class="px-4 py-2 rounded-lg bg-emerald-700 text-white font-bold text-sm">Export Excel</a></div></header>
<form class="neo-shell flex flex-wrap gap-3 items-end rounded-3xl p-5"><div><label class="block text-xs font-bold">Start</label><input type="date" name="start_date" value="{{ start_date }}" class="border rounded-lg px-3 py-2"></div><div><label class="block text-xs font-bold">End</label><input type="date" name="end_date" value="{{ end_date }}" class="border rounded-lg px-3 py-2"></div><div><label class="block text-xs font-bold">Employee</label><select name="employee" class="border rounded-lg px-3 py-2"><option value="ALL">All Employees</option>{% for emp in employees %}<option value="{{ emp.user_id }}" {% if selected_emp==emp.user_id %}selected{% endif %}>{{ emp.name }}</option>{% endfor %}</select></div><button class="bg-slate-900 text-white rounded-lg px-5 py-2 font-bold">Preview</button></form>
<div class="neo-shell overflow-auto rounded-3xl"><table class="border-collapse text-xs min-w-max"><thead class="sticky top-0 z-10"><tr class="bg-slate-800 text-white"><th class="sticky-col bg-slate-800 border p-2 min-w-64 text-left">Employee</th>{% for d in dates %}<th class="cell border border-slate-600 p-2">{{ d[5:] }}</th>{% endfor %}<th class="border p-2">Extra Total</th></tr></thead><tbody>
{% for row in rows %}<tr><th class="sticky-col bg-white border p-2 text-left"><b>{{ row.name }}</b><br><span class="text-slate-400">{{ row.code }} · {{ row.dept }}</span></th>{% for c in row.cells %}<td class="cell border p-2 text-center {{ c.color }}"><b>{{ c.status }}</b><br><span>{{ c.hours }}</span><br><span class="font-bold">{{ c.extra }}</span></td>{% endfor %}<td class="border p-2 text-center font-black text-emerald-700">{{ row.extra_total }}</td></tr>{% endfor %}
</tbody></table></div><div class="text-xs text-slate-500">Green: Present · Red: Absent · Yellow: Approved Leave · Slate: Weekly Off · Orange: shift exception</div>
</main></body></html>
"""

REPORTS_TEMPLATE = """
<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Day-wise Attendance Reports</title>
<style id="gamek-neo-ui">
:root{--neo:#07111f;--neo2:#0d1b2e;--cyan:#22d3ee;--mint:#34d399;--violet:#8b5cf6;--line:rgba(148,163,184,.18);--glass:rgba(255,255,255,.72)}
html{scroll-behavior:smooth}body{background:radial-gradient(circle at 8% 4%,rgba(52,211,153,.20),transparent 25%),radial-gradient(circle at 92% 5%,rgba(59,130,246,.20),transparent 28%),linear-gradient(145deg,#f8fafc,#eef4ff)!important;color:#142033}
body:before{content:'';position:fixed;inset:0;pointer-events:none;z-index:-1;background-image:linear-gradient(rgba(15,23,42,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(15,23,42,.025) 1px,transparent 1px);background-size:42px 42px;mask-image:linear-gradient(to bottom,black,transparent 82%)}
.neo-shell{border:1px solid rgba(255,255,255,.75);background:linear-gradient(145deg,rgba(255,255,255,.86),rgba(255,255,255,.62));backdrop-filter:blur(24px);box-shadow:0 22px 70px rgba(15,23,42,.10)}
header,.card,.bg-white{backdrop-filter:blur(20px)}
button,a,input,select{transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease}button:hover,a:hover{transform:translateY(-2px)}button:active,a:active{transform:scale(.97)}
input:focus,select:focus{outline:none!important;border-color:var(--cyan)!important;box-shadow:0 0 0 4px rgba(34,211,238,.13)!important}
::-webkit-scrollbar{width:10px;height:10px}::-webkit-scrollbar-track{background:#e8eef7}::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#34d399,#3b82f6);border-radius:20px;border:2px solid #e8eef7}
table th{letter-spacing:.035em}tbody tr{transition:background .18s ease,transform .18s ease}tbody tr:hover{background:rgba(224,242,254,.48)!important}
@keyframes neoIn{from{opacity:0;transform:translateY(12px) scale(.99)}to{opacity:1;transform:none}}main>*,body>main>*{animation:neoIn .45s ease both}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;scroll-behavior:auto!important}}
</style>
<script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-slate-50 text-slate-800"><main class="mx-auto max-w-[1900px] p-4 md:p-7 space-y-5"><header class="flex flex-wrap justify-between gap-3"><div><h1 class="text-2xl font-black">Day-wise Reports</h1><p class="text-sm text-slate-500">Weekly off worked, scheduled-versus-actual shift, and unapproved shift arrival alerts.</p></div><a href="/" class="px-4 py-2 rounded-lg border bg-white font-bold text-sm">Dashboard</a></header>
{% if flash_count %}<div class="animate-pulse border-2 border-rose-500 bg-rose-50 text-rose-800 rounded-xl p-4 font-black">⚠ {{ flash_count }} unapproved shift mismatch alert(s) in selected period.</div>{% endif %}
<form class="neo-shell flex flex-wrap gap-3 items-end rounded-3xl p-5"><div><label class="block text-xs font-bold">Start</label><input type="date" name="start_date" value="{{ start_date }}" class="border rounded-lg px-3 py-2"></div><div><label class="block text-xs font-bold">End</label><input type="date" name="end_date" value="{{ end_date }}" class="border rounded-lg px-3 py-2"></div><div><label class="block text-xs font-bold">Employee</label><select name="employee" class="border rounded-lg px-3 py-2"><option value="ALL">All Employees</option>{% for emp in employees %}<option value="{{ emp.user_id }}" {% if selected_emp==emp.user_id %}selected{% endif %}>{{ emp.name }}</option>{% endfor %}</select></div><button class="bg-slate-900 text-white rounded-lg px-5 py-2 font-bold">Run Report</button></form>
<div class="grid grid-cols-2 md:grid-cols-4 gap-3">{% for label,value,color in cards %}<div class="bg-white border-l-4 {{ color }} rounded-xl p-4 shadow-sm"><p class="text-xs uppercase font-bold text-slate-400">{{ label }}</p><p class="text-2xl font-black">{{ value }}</p></div>{% endfor %}</div>
<div class="neo-shell overflow-auto rounded-3xl"><table class="w-full min-w-[1100px] text-xs"><thead><tr class="bg-slate-800 text-white"><th class="p-3 text-left">Date</th><th class="p-3 text-left">Employee</th><th>Scheduled</th><th>Actual</th><th>In</th><th>Out</th><th>Hours</th><th>Extra</th><th>Status / Finding</th></tr></thead><tbody>{% for r in report_rows %}<tr class="border-t {% if r.shift_exception %}bg-rose-50{% endif %}"><td class="p-3">{{ r.date }}</td><td class="p-3"><b>{{ r.name }}</b><br><span class="text-slate-400">{{ r.user_id }} · {{ r.dept }}</span></td><td class="text-center">{{ r.scheduled_shift }}</td><td class="text-center font-bold">{{ r.shift_type }}</td><td class="text-center">{{ r.store_in }}</td><td class="text-center">{{ r.out_time }}</td><td class="text-center">{{ r.total_hours }}</td><td class="text-center font-bold">{{ r.net_variance }}</td><td class="p-3 font-bold {% if r.shift_exception %}text-rose-700{% endif %}">{{ r.shift_exception or r.status }}</td></tr>{% else %}<tr><td colspan="9" class="p-10 text-center text-slate-400">No report records.</td></tr>{% endfor %}</tbody></table></div>
</main></body></html>
"""

def _extra_minutes(value):
    if not value or not str(value).startswith(('H06;', 'H07;')):
        return 0
    try: return int(str(value).split(';', 1)[1]) * 60
    except (ValueError, IndexError): return 0

@app.route('/matrix_preview')
def matrix_preview():
    if not session.get('logged_in') or session.get('role') not in ['admin', 'developer']:
        return redirect(url_for('login'))
    today = datetime.now().strftime('%Y-%m-%d')
    start_date = request.args.get('start_date', today[:8] + '01')
    end_date = request.args.get('end_date', today)
    selected_emp = request.args.get('employee', 'ALL')
    logs, employees, _, _, _, _, _ = fetch_attendance_data(start_date, end_date, selected_emp)
    start_dt, end_dt = datetime.strptime(start_date, '%Y-%m-%d'), datetime.strptime(end_date, '%Y-%m-%d')
    dates=[]; cur=start_dt
    while cur<=end_dt: dates.append(cur.strftime('%Y-%m-%d')); cur += timedelta(days=1)
    by_key={(x['user_id'],x['date']):x for x in logs}
    rows=[]
    for emp in employees:
        if selected_emp != 'ALL' and emp['user_id'] != selected_emp: continue
        cells=[]; extra_total=0
        for d in dates:
            x=by_key.get((emp['user_id'],d), {})
            status=x.get('status','Absent'); exception=x.get('shift_exception','')
            color='bg-orange-100 text-orange-800' if exception else ('bg-emerald-50 text-emerald-800' if status in ('Present','Mis Punch') else 'bg-slate-100 text-slate-700' if status=='Weekly Off' else 'bg-yellow-100 text-yellow-800' if 'Leave' in status else 'bg-rose-50 text-rose-700')
            extra=_extra_minutes(x.get('net_variance')); extra_total += extra
            cells.append({'status': exception or status, 'hours':x.get('total_hours','-'),'extra':x.get('net_variance','-'),'color':color})
        rows.append({'code':emp['user_id'],'name':emp['name'],'dept':emp['dept'],'cells':cells,'extra_total':f'{extra_total//60}h {extra_total%60}m'})
    return render_template_string(MATRIX_PREVIEW_TEMPLATE,start_date=start_date,end_date=end_date,selected_emp=selected_emp,employees=employees,dates=dates,rows=rows)

@app.route('/reports/day-wise')
def day_wise_reports():
    if not session.get('logged_in') or session.get('role') not in ['admin', 'developer']:
        return redirect(url_for('login'))
    today=datetime.now().strftime('%Y-%m-%d'); start_date=request.args.get('start_date',today[:8]+'01'); end_date=request.args.get('end_date',today); selected_emp=request.args.get('employee','ALL')
    logs, employees, _, _, _, _, _=fetch_attendance_data(start_date,end_date,selected_emp)
    report_rows=[]
    for x in logs:
        worked_off=x.get('scheduled_shift')=='Off' and x.get('store_in')!='-'
        mismatch=x.get('shift_exception')=='Unapproved Shift Change'
        if worked_off: x=dict(x,shift_exception='Weekly Off Not Taken / Employee Worked')
        if worked_off or mismatch: report_rows.append(x)
    flash_count=sum(1 for x in report_rows if 'Unapproved' in x.get('shift_exception',''))
    cards=[('Exceptions',len(report_rows),'border-l-orange-500'),('Shift Mismatch',flash_count,'border-l-rose-500'),('Off Worked',sum(1 for x in report_rows if 'Off Not Taken' in x.get('shift_exception','')),'border-l-indigo-500'),('Employees',len({x['user_id'] for x in report_rows}),'border-l-emerald-500')]
    return render_template_string(REPORTS_TEMPLATE,start_date=start_date,end_date=end_date,selected_emp=selected_emp,employees=employees,report_rows=report_rows,flash_count=flash_count,cards=cards)

@app.route('/monthly_roster/auto_generate', methods=['POST'])
def auto_generate_monthly_roster():
    if not session.get('logged_in') or session.get('role') not in ['admin','developer']:
        return redirect(url_for('login'))
    month=request.form.get('month') or datetime.now().strftime('%Y-%m')
    try: datetime.strptime(month,'%Y-%m')
    except ValueError: return 'Invalid roster month.',400
    MONTHLY_ROSTERS[month]=generate_alternating_month_roster(month)
    if not save_monthly_rosters(): return 'Unable to save generated roster.',500
    flash('Date-wise alternating roster generated. Monday se week count aur fixed weekly off apply ho gaya.','success')
    return redirect(url_for('monthly_roster',month=month))

@app.route('/monthly_roster', methods=['GET','POST'])
def monthly_roster():
    if not session.get('logged_in') or session.get('role') not in ['admin','developer']:
        return redirect(url_for('login'))
    month=(request.form.get('month') if request.method=='POST' else request.args.get('month')) or datetime.now().strftime('%Y-%m')
    try: datetime.strptime(month,'%Y-%m')
    except ValueError: return 'Invalid roster month. Use YYYY-MM.',400
    month_dates=build_month_dates(month)
    if month not in MONTHLY_ROSTERS or not all('dates' in v for v in MONTHLY_ROSTERS.get(month,{}).values() if isinstance(v,dict)):
        MONTHLY_ROSTERS[month]=generate_alternating_month_roster(month)
        save_monthly_rosters()
    if request.method=='POST':
        allowed={'Shift A','Shift B','Off'}; saved={}
        for raw_code in MASTER_EMPLOYEES:
            code=normalized_code(raw_code); off=request.form.get('off_'+code,get_emp_info(code).get('off','SUNDAY')).upper()
            date_map={}
            for d in month_dates:
                ds=d.strftime('%Y-%m-%d'); value=request.form.get(f'date_{code}_{ds}')
                if value not in allowed: return 'Roster contains an invalid assignment.',400
                date_map[ds]=value
            saved[code]={'dates':date_map,'weekly_off':off,'first_week_shift':date_map[month_dates[0].strftime('%Y-%m-%d')] if date_map else 'Shift A'}
        MONTHLY_ROSTERS[month]=saved
        if not save_monthly_rosters(): return 'Unable to save roster.',500
        flash('Complete monthly date-wise roster saved successfully.','success')
        return redirect(url_for('monthly_roster',month=month))
    employees=[]
    for raw_code in sorted(MASTER_EMPLOYEES,key=lambda c:get_emp_info(c)['name']):
        code=normalized_code(raw_code); info=get_emp_info(code); employees.append({'user_id':code,'name':info['name'],'dept':info['dept']})
    first_monday=month_dates[0]-timedelta(days=month_dates[0].weekday()); dates=[]; groups=[]
    for d in month_dates:
        wi=(d-first_monday).days//7
        dates.append({'iso':d.strftime('%Y-%m-%d'),'day':d.strftime('%a'),'day_upper':d.strftime('%A').upper(),'date_label':d.strftime('%d %b'),'week_index':wi,'is_today':d==datetime.now().date()})
    for wi in sorted(set(x['week_index'] for x in dates)):
        group_dates=[x for x in dates if x['week_index']==wi]
        groups.append({'number':wi+1,'dates':group_dates,'shift_label':'Alternating Shift A/B'})
    return render_template_string(MONTHLY_ROSTER_TEMPLATE,month=month,employees=employees,designations=sorted({e['dept'] for e in employees}),roster=MONTHLY_ROSTERS[month],dates=dates,week_groups=groups,weekdays=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'])

@app.route('/monthly_roster/export')
def export_monthly_roster():
    if not session.get('logged_in') or session.get('role') not in ['admin', 'developer']:
        return redirect(url_for('login'))

    month = request.args.get('month', datetime.now().strftime('%Y-%m'))
    try:
        month_start = datetime.strptime(f'{month}-01', '%Y-%m-%d').date()
    except (TypeError, ValueError):
        return 'Invalid roster month. Use YYYY-MM.', 400

    next_month = (month_start + timedelta(days=32)).replace(day=1)
    month_dates = [
        month_start + timedelta(days=day_offset)
        for day_offset in range((next_month - month_start).days)
    ]
    patterns = get_monthly_roster_patterns(month)

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = 'Monthly Roster'
    worksheet.append(['Gamek Fresmart Express - Monthly Employee Roster'])
    worksheet.append([f'Month: {month}'])
    worksheet.append(['A = Shift A | B = Shift B | Off = Weekly Off'])
    worksheet.append(['ID', 'Employee Name', 'Department'] + [date.strftime('%a %d') for date in month_dates])
    worksheet.freeze_panes = 'D5'
    worksheet.column_dimensions['A'].width = 14
    worksheet.column_dimensions['B'].width = 34
    worksheet.column_dimensions['C'].width = 24

    for date_index in range(len(month_dates)):
        worksheet.column_dimensions[openpyxl.utils.get_column_letter(date_index + 4)].width = 12

    for emp_code, emp_data in sorted(MASTER_EMPLOYEES.items(), key=lambda item: get_emp_info(item[0])['name']):
        final_code = f"NWC{emp_code}" if not emp_code.startswith('NWC') else emp_code
        info = get_emp_info(emp_code)
        row = [final_code, info['name'], info['dept']]
        for current_date in month_dates:
            assignment = get_roster_assignment_from_patterns(patterns[final_code], current_date)
            row.append({'Shift A': 'A', 'Shift B': 'B', 'Off': 'Off'}[assignment])
        worksheet.append(row)

    for cell in worksheet[4]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(fill_type='solid', fgColor='334155')
        cell.alignment = Alignment(horizontal='center')

    excel_io = io.BytesIO()
    workbook.save(excel_io)
    excel_io.seek(0)
    return send_file(
        excel_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'Employee_Roster_{month}.xlsx'
    )

@app.route('/apply_leave', methods=['POST'])
def apply_leave():
    if not session.get('logged_in') or session.get('role') != 'employee':
        return redirect(url_for('login'))
        
    user_id = session.get('user_id')
    name = session.get('user_name')
    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')
    leave_type = request.form.get('leave_type', 'F10;1')
    
    filename = None
    file = request.files.get('supporting_doc')
    if file and file.filename != '':
        if not allowed_file(file.filename):
            flash('Invalid file format! Sirf PDF, JPG, PNG ya DOC files allowed hain.', 'danger')
            return redirect(url_for('index'))
            
        filename = secure_filename(file.filename)
        local_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(local_path)
        
        # 1. GitHub par upload karein (Folder ka naam: leave_documents/)
        github_path = f"leave_documents/{user_id}_{filename}"
        upload_file_to_github(local_path, github_path)

    # 2. Excel me record save karein
    save_leave_to_excel(user_id, name, start_date, end_date, leave_type, filename if filename else "No Document")

    # 3. Global list me add karein aur JSON persistence save karein
    leave_req = {
        'id': len(LEAVE_REQUESTS) + 1,
        'user_id': user_id,
        'name': name,
        'start_date': start_date,
        'end_date': end_date,
        'leave_type': leave_type,
        'filename': filename,
        'status': 'Pending'
    }
    LEAVE_REQUESTS.append(leave_req)
    save_leave_requests(LEAVE_REQUESTS)
    
    flash('Aapki leave request successfully submit ho gayi hai aur document GitHub par sync ho gaya hai!', 'success')
    return redirect(url_for('index'))

@app.route('/submit_shift_change', methods=['POST'])
def submit_shift_change():
    if not session.get('logged_in') or session.get('role') != 'employee':
        return redirect(url_for('login'))
    requester_id = normalized_code(session.get('user_id'))
    partner_id = normalized_code(request.form.get('partner_id'))
    swap_date = request.form.get('swap_date', '').strip()
    try:
        datetime.strptime(swap_date, '%Y-%m-%d')
    except ValueError:
        flash('Valid shift swap date select karein.', 'error'); return redirect(url_for('index'))
    if partner_id == requester_id or partner_id not in MASTER_EMPLOYEES:
        flash('Valid opposite-shift employee select karein.', 'error'); return redirect(url_for('index'))
    requester_shift = effective_roster_assignment(requester_id, swap_date)
    partner_shift = effective_roster_assignment(partner_id, swap_date)
    if requester_shift not in ['Shift A','Shift B'] or partner_shift not in ['Shift A','Shift B'] or requester_shift == partner_shift:
        flash('Shift swap sirf Shift A aur opposite Shift B employee ke beech ho sakta hai.', 'error'); return redirect(url_for('index'))
    if any(x.get('swap_date') == swap_date and x.get('status') == 'Pending' and requester_id in [x.get('requester_id'),x.get('partner_id')] for x in SHIFT_CHANGE_REQUESTS):
        flash('Is date ke liye shift swap request already pending hai.', 'error'); return redirect(url_for('index'))
    req={'id':max([x.get('id',0) for x in SHIFT_CHANGE_REQUESTS] or [0])+1,'requester_id':requester_id,'requester_name':get_emp_info(requester_id)['name'],'partner_id':partner_id,'partner_name':get_emp_info(partner_id)['name'],'swap_date':swap_date,'requester_shift':requester_shift,'partner_shift':partner_shift,'reason':request.form.get('reason','').strip(),'employee_approval':'Pending','admin_approval':'Pending','status':'Pending','requested_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    SHIFT_CHANGE_REQUESTS.append(req); save_shift_change_requests()
    flash('Shift swap request partner aur admin approval ke liye bhej di gayi.', 'success')
    return redirect(url_for('index'))

@app.route('/update_shift_change/<int:req_id>/<action>')
def update_shift_change(req_id, action):
    if not session.get('logged_in'): return redirect(url_for('login'))
    actor=normalized_code(session.get('user_id')); role=session.get('role')
    for req in SHIFT_CHANGE_REQUESTS:
        if req.get('id') != req_id: continue
        decision='Approved' if action=='approve' else 'Rejected'
        if role in ['admin','developer']: req['admin_approval']=decision
        elif actor == req.get('partner_id'): req['employee_approval']=decision
        else: flash('Aap is request ko approve nahi kar sakte.','error'); return redirect(url_for('index'))
        refresh_dual_approval(req); break
    save_shift_change_requests(); flash('Shift swap approval update ho gaya.','success')
    return redirect(url_for('index'))

@app.route('/api/employees-off-on-date')
def employees_off_on_date():
    if not session.get('logged_in'):
        return {'employees': []}, 401
    date_str = request.args.get('date', '').strip()
    try:
        target = datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        return {'employees': []}
    requester = normalized_code(session.get('user_id'))
    employees = []
    for emp_code in MASTER_EMPLOYEES:
        code = normalized_code(emp_code)
        if code == requester:
            continue
        info = get_emp_info(code)
        assignment = get_monthly_roster_assignment(code, date_str)
        scheduled_off = assignment == 'Off' if assignment else target.strftime('%A').upper() == info.get('off', '').upper()
        approved_change = get_approved_off_change(code, date_str)
        if approved_change == 'WORKING':
            scheduled_off = False
        elif approved_change == 'OFF':
            scheduled_off = True
        if scheduled_off:
            employees.append({'code': code, 'name': info.get('name', code)})
    employees.sort(key=lambda x: x['name'])
    return {'employees': employees}

@app.route('/submit_off_change', methods=['POST'])
def submit_off_change():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    user_id = session.get('user_id')
    if not str(user_id).startswith('NWC'):
        user_id = f"NWC{user_id}"
    old_off_date = request.form.get('old_off_date', '').strip()
    new_off_date = request.form.get('new_off_date', '').strip()
    reason = request.form.get('reason', '').strip()
    partner_id = normalized_code(request.form.get('partner_id', ''))
    if not partner_id or partner_id == user_id or partner_id not in MASTER_EMPLOYEES:
        flash('Valid dusra employee select karein.', 'error')
        return redirect(url_for('index'))
    try:
        old_dt = datetime.strptime(old_off_date, '%Y-%m-%d').date()
        new_dt = datetime.strptime(new_off_date, '%Y-%m-%d').date()
    except ValueError:
        flash('Valid old off date aur new off date select karein.', 'error')
        return redirect(url_for('index'))
    if old_dt == new_dt:
        flash('Old off date aur new off date alag honi chahiye.', 'error')
        return redirect(url_for('index'))
    if any(r.get('user_id') == user_id and r.get('status') == 'Pending' and (r.get('old_off_date') == old_off_date or r.get('new_off_date') == new_off_date) for r in OFF_CHANGE_REQUESTS):
        flash('Is date ke liye off change request already pending hai.', 'error')
        return redirect(url_for('index'))
    info = get_emp_info(user_id)
    req = {
        'id': max([r.get('id', 0) for r in OFF_CHANGE_REQUESTS] or [0]) + 1,
        'user_id': user_id, 'name': info['name'],
        'partner_id': partner_id, 'partner_name': get_emp_info(partner_id)['name'],
        'employee_approval': 'Pending', 'admin_approval': 'Pending',
        'old_off_date': old_off_date, 'new_off_date': new_off_date,
        'reason': reason, 'status': 'Pending',
        'requested_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    OFF_CHANGE_REQUESTS.append(req)
    save_off_change_requests()
    flash('Weekly off change request submit ho gayi hai.', 'success')
    return redirect(url_for('index'))

@app.route('/update_off_change/<int:req_id>/<action>')
def update_off_change(req_id, action):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    actor = normalized_code(session.get('user_id'))
    role = session.get('role')
    for req in OFF_CHANGE_REQUESTS:
        if req.get('id') != req_id:
            continue
        decision = 'Approved' if action == 'approve' else 'Rejected'
        if role in ['admin', 'developer']:
            req['admin_approval'] = decision
        elif actor == req.get('partner_id'):
            req['employee_approval'] = decision
        else:
            flash('Aap is request ko approve nahi kar sakte.', 'error')
            return redirect(url_for('index'))
        refresh_dual_approval(req)
        break
    save_off_change_requests()
    flash('Weekly off swap approval update ho gaya.', 'success')
    return redirect(url_for('index'))

@app.route('/update_leave/<int:req_id>/<action>')
def update_leave(req_id, action):
    if not session.get('logged_in') or session.get('role') not in ['admin', 'developer']:
        return redirect(url_for('login'))
        
    for req in LEAVE_REQUESTS:
        if req['id'] == req_id:
            if action == 'approve':
                req['status'] = 'Approved'
                flash(f"Leave request for {req['name']} approved successfully!", 'success')
            elif action == 'reject':
                req['status'] = 'Rejected'
                flash(f"Leave request for {req['name']} rejected.", 'success')
            break
            
    save_leave_requests(LEAVE_REQUESTS)
    return redirect(url_for('index'))

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/export')
def export_excel():
    if not session.get('logged_in') or session.get('role') not in ['admin', 'developer']:
        return redirect(url_for('login'))
        
    start_date = request.args.get('start_date', datetime.now().strftime('%Y-%m-%d'))
    end_date = request.args.get('end_date', datetime.now().strftime('%Y-%m-%d'))
    selected_emp = request.args.get('employee', 'ALL')
    
    logs, _, g_hrs, g_l_hrs, g_var, _, _ = fetch_attendance_data(start_date, end_date, selected_emp)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance Report"
    ws.sheet_view.showGridLines = True
    
    headers = ["Sr. No.", "Date", "ID", "Employee Name", "Department", "Store In", "Lunch Out", "Lunch In", "Out Time", "Total Lunch", "Working Hours", "Total Hora Extra", "Status"]
    ws.append([])
    ws.append(["Gamek Fresmart Express - Attendance Report"])
    ws.append([f"Period: {start_date} to {end_date}"])
    ws.append([])
    ws.append(headers)
    
    for idx, log in enumerate(logs, 1):
        ws.append([
            idx, log['date'], log['user_id'], log['name'], log['dept'],
            log['store_in'], log['lunch_out'], log['lunch_in'], log['out_time'],
            log['total_lunch'], log['total_hours'], log['net_variance'], log['status']
        ])
        
    ws.append([])
    ws.append(["", "", "", "", "", "", "", "", "Total Summary:", g_l_hrs, g_hrs, g_var])
    
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    
    filename = f"Attendance_Report_{start_date}_to_{end_date}.xlsx"
    return send_file(excel_io, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)

@app.route('/export_matrix')
def export_matrix():
    if not session.get('logged_in') or session.get('role') not in ['admin', 'developer']:
        return redirect(url_for('login'))
        
    start_date = request.args.get('start_date', datetime.now().strftime('%Y-%m-%d'))
    end_date = request.args.get('end_date', datetime.now().strftime('%Y-%m-%d'))

    logs, _, _, _, _, _, _ = fetch_attendance_data(start_date, end_date, 'ALL')
    logs_by_employee_and_date = {
        (log['user_id'], log['date']): log
        for log in logs
    }
    
    start_dt = datetime.strptime(start_date, '%Y-%m-%d')
    end_dt = datetime.strptime(end_date, '%Y-%m-%d')
    date_list = []
    curr = start_dt
    while curr <= end_dt:
        date_list.append(curr.strftime('%Y-%m-%d'))
        curr += timedelta(days=1)
        
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employee Matrix"
    ws.sheet_view.showGridLines = True
    
    headers = ["ID", "Employee Name", "Department"] + date_list + [
        "Total Absent", "Total Weekly Off", "Total Present"
    ]
    ws.append(["Gamek Fresmart Express - Employee Matrix Attendance Report"])
    ws.append([f"Period: {start_date} to {end_date}"])
    ws.append([])
    ws.append(headers)
    
    for emp_code, emp_data in sorted(MASTER_EMPLOYEES.items(), key=lambda x: get_emp_info(x[0])['name']):
        final_code = f"NWC{emp_code}" if not emp_code.startswith('NWC') else emp_code
        emp_info = get_emp_info(emp_code)
        row = [final_code, emp_info['name'], emp_info['dept']]
        
        for d_str in date_list:
            st = logs_by_employee_and_date.get((final_code, d_str))
            if st:
                status = st['status']
                if '(Leave)' in status:
                    # Show approved leave code in Excel, e.g. F10;1
                    row.append(status.split(' (Leave)', 1)[0].strip())
                elif status == 'Weekly Off':
                    row.append('Weekly Off')
                elif status == 'Absent':
                    row.append('F03;1')
                elif status in ('Present', 'Mis Punch'):
                    row.append('Present')
                else:
                    row.append('F03;1')
            else:
                curr_dt_obj = datetime.strptime(d_str, '%Y-%m-%d')
                if curr_dt_obj.strftime('%A').upper() == emp_info['off'].upper():
                    row.append('Weekly Off')
                else:
                    row.append('F03;1')
        date_statuses = row[3:]
        row.extend([
            date_statuses.count('F03;1'),
            date_statuses.count('Weekly Off'),
            date_statuses.count('Present')
        ])
        ws.append(row)

    for cell in ws[4]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(fill_type='solid', fgColor='334155')
        cell.alignment = Alignment(horizontal='center')

    absent_fill = PatternFill(fill_type='solid', fgColor='FF0000')
    absent_font = Font(color='FFFFFF', bold=True)
    weekly_off_fill = PatternFill(fill_type='solid', fgColor='16A34A')
    weekly_off_font = Font(color='FFFFFF', bold=True)
    approved_leave_fill = PatternFill(fill_type='solid', fgColor='FFFF00')
    approved_leave_font = Font(color='000000', bold=True)
    approved_leave_codes = {'F01;1', 'F05;1', 'F10;1', 'F51;1', 'F60;1', 'F61;1', 'F62;1'}
    for row in ws.iter_rows(min_row=5, min_col=4, max_col=3 + len(date_list)):
        for cell in row:
            if cell.value == 'F03;1':
                cell.fill = absent_fill
                cell.font = absent_font
            elif cell.value == 'Weekly Off':
                cell.fill = weekly_off_fill
                cell.font = weekly_off_font
            elif cell.value in approved_leave_codes:
                cell.fill = approved_leave_fill
                cell.font = approved_leave_font
        
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    filename = f"Employee_Matrix_{start_date}_to_{end_date}.xlsx"
    return send_file(excel_io, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)

def _backup_source_files():
    files = []
    fixed = [
        'app.py', LEAVE_JSON_FILE, LEAVE_EXCEL_FILE, MONTHLY_ROSTER_FILE,
        OFF_CHANGE_JSON_FILE, SHIFT_CHANGE_JSON_FILE,
        'attendance_logs.json', 'users.json', 'employee_data.json', EMPLOYEES_JSON_FILE, STORES_JSON_FILE, REGULARISATION_FILE, BIOMETRIC_PUSH_QUEUE_FILE,
        'requirements.txt', '.env', STORES_JSON_FILE
    ]
    for item in fixed:
        path = os.path.abspath(item)
        if os.path.isfile(path):
            files.append((path, os.path.basename(path)))
    if os.path.isdir(UPLOAD_FOLDER):
        for root, _, names in os.walk(UPLOAD_FOLDER):
            for name in names:
                full = os.path.join(root, name)
                arc = os.path.join('uploads', os.path.relpath(full, UPLOAD_FOLDER))
                files.append((full, arc))
    return files

def export_live_portal_state(reason='update'):
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    payload = {
        'exported_at': datetime.now().isoformat(), 'reason': reason,
        'leave_requests': LEAVE_REQUESTS,
        'off_change_requests': OFF_CHANGE_REQUESTS,
        'shift_change_requests': SHIFT_CHANGE_REQUESTS,
        'monthly_rosters': MONTHLY_ROSTERS,
        'synced_attendance_logs': SYNCED_ATTENDANCE_LOGS
    }
    json_path = os.path.join(AUTO_EXPORT_FOLDER, f'portal_state_{stamp}.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, indent=2, ensure_ascii=False, default=str)
    latest = os.path.join(AUTO_EXPORT_FOLDER, 'portal_state_latest.json')
    shutil.copy2(json_path, latest)
    return json_path

def cleanup_old_backups():
    cutoff = datetime.now() - timedelta(days=BACKUP_RETENTION_DAYS)
    for name in os.listdir(BACKUP_FOLDER):
        path = os.path.join(BACKUP_FOLDER, name)
        if os.path.isfile(path) and name.endswith('.zip'):
            if datetime.fromtimestamp(os.path.getmtime(path)) < cutoff:
                try: os.remove(path)
                except OSError: pass

def create_full_backup(trigger='manual'):
    with BACKUP_LOCK:
        if BACKUP_STATE['running']:
            return None
        BACKUP_STATE.update(running=True, percent=1, message='Backup preparing...', file=None, error=None)
    try:
        export_live_portal_state('backup')
        files = _backup_source_files()
        stamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        filename = f'portal_full_backup_{stamp}_{trigger}.zip'
        final_path = os.path.join(BACKUP_FOLDER, filename)
        temp_path = final_path + '.tmp'
        total = max(len(files), 1)
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            manifest = {'created_at': datetime.now().isoformat(), 'trigger': trigger, 'files': [arc for _, arc in files]}
            zf.writestr('backup_manifest.json', json.dumps(manifest, indent=2))
            for idx, (src, arc) in enumerate(files, 1):
                zf.write(src, arc)
                pct = 5 + int((idx / total) * 90)
                BACKUP_STATE.update(percent=min(pct, 95), message=f'Backing up {arc}')
        os.replace(temp_path, final_path)
        cleanup_old_backups()
        BACKUP_STATE.update(running=False, percent=100, message='Backup completed', file=filename)
        return final_path
    except Exception as e:
        BACKUP_STATE.update(running=False, percent=0, message='Backup failed', error=str(e))
        return None

def backup_worker(trigger='manual'):
    create_full_backup(trigger)

def start_backup_async(trigger='manual'):
    if BACKUP_STATE['running']:
        return False
    threading.Thread(target=backup_worker, args=(trigger,), daemon=True, name='portal-full-backup').start()
    return True

def backup_scheduler_loop():
    while True:
        now = datetime.now()
        # Run once in the 12:00 and 22:00 minute windows.
        if now.hour in (12, 22) and now.minute < 2:
            key = now.strftime('%Y-%m-%d-%H')
            if key not in LAST_SCHEDULED_BACKUP_KEYS:
                LAST_SCHEDULED_BACKUP_KEYS.add(key)
                start_backup_async('scheduled')
        # Keep the marker set small.
        if len(LAST_SCHEDULED_BACKUP_KEYS) > 10:
            today = now.strftime('%Y-%m-%d')
            LAST_SCHEDULED_BACKUP_KEYS.intersection_update({k for k in LAST_SCHEDULED_BACKUP_KEYS if k.startswith(today)})
        threading.Event().wait(30)

def start_backup_scheduler():
    global BACKUP_SCHEDULER_STARTED
    with BACKUP_SCHEDULER_LOCK:
        if BACKUP_SCHEDULER_STARTED:
            return
        BACKUP_SCHEDULER_STARTED = True
        threading.Thread(target=backup_scheduler_loop, daemon=True, name='portal-backup-scheduler').start()

@app.after_request
def auto_export_after_update(response):
    # Save a local state export after successful portal changes.
    if request.method in ('POST', 'PUT', 'PATCH', 'DELETE') and response.status_code < 400:
        if not request.path.startswith('/developer/backup') and request.path != '/api/attendance/sync':
            try:
                export_live_portal_state(f'update:{request.path}')
            except Exception as e:
                print(f'Automatic local export failed: {e}')
    return response

@app.route('/developer/backup/start', methods=['POST'])
def developer_backup_start():
    if not session.get('logged_in') or session.get('role') != 'developer':
        return {'status':'error','message':'Developer access required'}, 403
    started = start_backup_async('manual')
    return {'status':'started' if started else 'running', 'message':'Backup started' if started else 'Backup already running'}

@app.route('/developer/backup/status')
def developer_backup_status():
    if not session.get('logged_in') or session.get('role') != 'developer':
        return {'status':'error'}, 403
    return dict(BACKUP_STATE)

@app.route('/developer/backups')
def developer_backups():
    if not session.get('logged_in') or session.get('role') != 'developer':
        return {'status':'error'}, 403
    items=[]
    for name in sorted(os.listdir(BACKUP_FOLDER), reverse=True):
        path=os.path.join(BACKUP_FOLDER,name)
        if os.path.isfile(path) and name.endswith('.zip'):
            items.append({'name':name,'size':os.path.getsize(path),'created':datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')})
    return {'backups':items[:100]}

@app.route('/developer/backup/download/<path:filename>')
def developer_backup_download(filename):
    if not session.get('logged_in') or session.get('role') != 'developer':
        return redirect(url_for('login'))
    return send_from_directory(BACKUP_FOLDER, secure_filename(filename), as_attachment=True)

def safe_restore_zip(zip_path):
    global LEAVE_REQUESTS, OFF_CHANGE_REQUESTS, SHIFT_CHANGE_REQUESTS, MONTHLY_ROSTERS
    allowed_roots = {'uploads'}
    allowed_files = {LEAVE_JSON_FILE, LEAVE_EXCEL_FILE, MONTHLY_ROSTER_FILE, OFF_CHANGE_JSON_FILE, SHIFT_CHANGE_JSON_FILE, 'attendance_logs.json', 'users.json', 'employee_data.json', EMPLOYEES_JSON_FILE, REGULARISATION_FILE, BIOMETRIC_PUSH_QUEUE_FILE}
    with zipfile.ZipFile(zip_path, 'r') as zf:
        names = zf.namelist()
        for name in names:
            normalized = os.path.normpath(name).replace('\\','/')
            if normalized.startswith('../') or os.path.isabs(normalized):
                raise ValueError('Unsafe backup archive path')
        for name in names:
            normalized = os.path.normpath(name).replace('\\','/')
            if normalized in allowed_files or normalized.split('/')[0] in allowed_roots:
                target = os.path.abspath(normalized)
                base = os.path.abspath(os.getcwd())
                if not target.startswith(base + os.sep):
                    raise ValueError('Unsafe restore target')
                os.makedirs(os.path.dirname(target), exist_ok=True)
                if not name.endswith('/'):
                    with zf.open(name) as src, open(target, 'wb') as dst:
                        shutil.copyfileobj(src, dst)
    LEAVE_REQUESTS = load_leave_requests()
    OFF_CHANGE_REQUESTS = load_off_change_requests()
    SHIFT_CHANGE_REQUESTS = load_shift_change_requests()
    MONTHLY_ROSTERS = load_monthly_rosters()

@app.route('/developer/backup/restore', methods=['POST'])
def developer_backup_restore():
    if not session.get('logged_in') or session.get('role') != 'developer':
        return {'status':'error','message':'Developer access required'}, 403
    uploaded = request.files.get('backup_file')
    existing = secure_filename(request.form.get('existing_backup',''))
    if not uploaded and not existing:
        return {'status':'error','message':'Backup ZIP select karein'}, 400
    # Safety backup before every restore.
    safety = create_full_backup('pre_restore')
    try:
        if uploaded:
            if not uploaded.filename.lower().endswith('.zip'):
                return {'status':'error','message':'Only ZIP backup allowed'}, 400
            temp = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
            temp.close(); uploaded.save(temp.name); restore_path=temp.name
        else:
            restore_path=os.path.join(BACKUP_FOLDER, existing)
            if not os.path.isfile(restore_path):
                return {'status':'error','message':'Backup not found'}, 404
        safe_restore_zip(restore_path)
        if uploaded: os.unlink(restore_path)
        return {'status':'success','message':'Complete data restored. Portal service restart recommended.','safety_backup':os.path.basename(safety) if safety else None}
    except Exception as e:
        return {'status':'error','message':str(e)}, 500


# ===== VERSIONED PORTAL DEPLOYMENT MANAGER =====
PORTAL_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PORTAL_VERSIONS_DIR = os.path.join(PORTAL_BASE_DIR, 'portal_versions')
PORTAL_ACTIVE_FILE = os.path.join(PORTAL_BASE_DIR, 'active_portal_version.json')
PORTAL_MANIFEST_FILE = os.path.join(PORTAL_BASE_DIR, 'portal_versions_manifest.json')
PORTAL_MAX_UPLOAD_BYTES = 8 * 1024 * 1024
os.makedirs(PORTAL_VERSIONS_DIR, exist_ok=True)

def _ensure_portal_launcher():
    launcher_path = os.path.join(PORTAL_BASE_DIR, 'portal_launcher.py')
    if os.path.isfile(launcher_path):
        return launcher_path
    launcher_source = """import argparse, json, os, runpy, sys, time, traceback
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.environ['PORTAL_BASE_DIR'] = BASE_DIR
os.environ['PORTAL_DATA_DIR'] = BASE_DIR
ACTIVE_FILE = os.path.join(BASE_DIR, 'active_portal_version.json')
VERSIONS_DIR = os.path.join(BASE_DIR, 'portal_versions')
FALLBACK = os.path.join(BASE_DIR, 'app.py')
def active_app():
    try:
        with open(ACTIVE_FILE, 'r', encoding='utf-8') as f:
            version = json.load(f).get('version','')
        candidate = os.path.abspath(os.path.join(VERSIONS_DIR, version, 'app.py'))
        if os.path.commonpath([candidate, os.path.abspath(VERSIONS_DIR)]) != os.path.abspath(VERSIONS_DIR):
            raise ValueError('Invalid active version path')
        if os.path.isfile(candidate): return candidate, version
    except Exception as exc:
        print('Active version lookup failed:', exc)
    return FALLBACK, 'root-fallback'
if __name__ == '__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--wait',type=float,default=0);args=parser.parse_args()
    if args.wait: time.sleep(max(0,min(args.wait,30)))
    target,version=active_app();os.chdir(BASE_DIR);sys.path.insert(0,os.path.dirname(target))
    print(f'Launching portal version {version}: {target}')
    try: runpy.run_path(target,run_name='__main__')
    except Exception: traceback.print_exc();raise
"""
    temp = launcher_path + '.tmp'
    with open(temp, 'w', encoding='utf-8') as f:
        f.write(launcher_source)
    os.replace(temp, launcher_path)
    return launcher_path

_ensure_portal_launcher()

def _deployment_access():
    return session.get('logged_in') and session.get('role') == 'developer' and session.get('user_id') == 'NCSA0608'

def _deployment_json(path, default):
    try:
        with open(path, 'r', encoding='utf-8') as f: return json.load(f)
    except Exception: return default

def _deployment_save(path, data):
    temp = path + '.tmp'
    with open(temp, 'w', encoding='utf-8') as f: json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(temp, path)

def _safe_version_name(value):
    value = (value or '').strip().lower().replace(' ', '-')
    value = ''.join(ch for ch in value if ch.isalnum() or ch in ('-','_','.'))
    return value[:60].strip('.-_')

def _ensure_current_version_registered():
    manifest = _deployment_json(PORTAL_MANIFEST_FILE, [])
    if not manifest:
        version = 'baseline-' + datetime.now().strftime('%Y%m%d-%H%M%S')
        version_dir = os.path.join(PORTAL_VERSIONS_DIR, version); os.makedirs(version_dir, exist_ok=True)
        source = os.path.abspath(__file__); target = os.path.join(version_dir, 'app.py')
        shutil.copy2(source, target)
        digest = hashlib.sha256(Path(target).read_bytes()).hexdigest()
        manifest.append({'version':version,'file':'app.py','sha256':digest,'size':os.path.getsize(target),'uploaded_at':datetime.now().isoformat(timespec='seconds'),'uploaded_by':'NCSA0608','notes':'Automatic baseline before version manager','status':'active'})
        _deployment_save(PORTAL_MANIFEST_FILE, manifest)
        _deployment_save(PORTAL_ACTIVE_FILE, {'version':version,'activated_at':datetime.now().isoformat(timespec='seconds'),'activated_by':'NCSA0608'})
    return manifest

PORTAL_UPDATE_HTML = r'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><script src="https://cdn.tailwindcss.com"></script><title>Portal Update Manager</title></head><body class="bg-slate-100"><main class="max-w-7xl mx-auto p-5 space-y-5"><header class="bg-slate-950 text-white rounded-3xl p-6 flex justify-between items-center"><div><p class="text-cyan-300 font-black text-xs tracking-widest">NCSA0608 RELEASE CONTROL</p><h1 class="text-3xl font-black">Portal Update & Rollback</h1><p class="text-slate-400 text-sm mt-1">Versioned deployment without overwriting the server root app.py.</p></div><a href="/" class="rounded-xl bg-white/10 px-4 py-2 font-bold">Dashboard</a></header>{% for c,m in get_flashed_messages(with_categories=true) %}<div class="rounded-xl p-4 font-bold {{'bg-emerald-50 text-emerald-700' if c=='success' else 'bg-rose-50 text-rose-700'}}">{{m}}</div>{% endfor %}<section class="grid lg:grid-cols-3 gap-5"><form method="post" action="/developer/portal-update/upload" enctype="multipart/form-data" class="lg:col-span-2 bg-white rounded-3xl p-6 shadow space-y-4"><div><p class="text-xs font-black text-indigo-600">NEW RELEASE</p><h2 class="text-2xl font-black">Upload updated app.py</h2></div><div class="grid md:grid-cols-2 gap-3"><input name="version" required placeholder="Version, e.g. v2.4.0" class="border rounded-xl p-3"><input name="notes" placeholder="Release notes" class="border rounded-xl p-3"></div><input name="app_file" type="file" accept=".py,text/x-python" required class="border rounded-xl p-4 w-full"><label class="flex gap-2 items-center text-sm font-bold"><input type="checkbox" name="activate" value="yes" checked class="w-4 h-4">Upload ke baad latest version ko default live version banayein</label><button class="bg-gradient-to-r from-cyan-500 to-indigo-600 text-white rounded-xl p-3 w-full font-black">Validate, Save and Deploy</button><p class="text-xs text-slate-500">Maximum 8 MB. File pehle Python compile validation se pass hogi. Failed validation wali file deploy nahi hogi.</p></form><div class="bg-white rounded-3xl p-6 shadow"><p class="text-xs font-black text-emerald-600">CURRENT LIVE</p><h2 class="text-2xl font-black break-all">{{active.version or 'Not selected'}}</h2><p class="text-sm text-slate-500 mt-2">Activated: {{active.activated_at or '-'}}</p><div class="mt-5 rounded-2xl bg-amber-50 p-4 text-xs text-amber-800"><b>Important:</b> Live switch ke baad restart button use karein. Launcher active version ko default run karega.</div><form method="post" action="/developer/portal-update/restart" class="mt-4"><button class="bg-slate-950 text-white rounded-xl p-3 w-full font-black">Restart Active Version</button></form></div></section><section class="bg-white rounded-3xl p-6 shadow"><h2 class="text-xl font-black mb-4">Version History</h2><div class="overflow-auto"><table class="w-full text-sm"><thead><tr class="bg-slate-900 text-white"><th class="p-3 text-left">Version</th><th>Uploaded</th><th>Size</th><th>Checksum</th><th>Notes</th><th>Action</th></tr></thead><tbody>{% for x in versions|reverse %}<tr class="border-b"><td class="p-3"><b>{{x.version}}</b>{% if x.version==active.version %}<span class="ml-2 rounded-full bg-emerald-100 text-emerald-700 px-2 py-1 text-[10px] font-black">LIVE</span>{% endif %}</td><td>{{x.uploaded_at}}</td><td>{{(x.size/1024)|round(1)}} KB</td><td class="font-mono text-xs">{{x.sha256[:12]}}</td><td>{{x.notes or '-'}}</td><td>{% if x.version!=active.version %}<form method="post" action="/developer/portal-update/activate/{{x.version}}"><button class="rounded-xl bg-indigo-50 text-indigo-700 px-3 py-2 font-bold">Make Live</button></form>{% else %}<span class="text-emerald-600 font-bold">Active</span>{% endif %}</td></tr>{% endfor %}</tbody></table></div></section></main></body></html>'''

@app.route('/developer/portal-update')
def portal_update_manager():
    if not _deployment_access(): return 'NCSA0608 developer access required', 403
    versions = _ensure_current_version_registered()
    active = _deployment_json(PORTAL_ACTIVE_FILE, {})
    _ensure_portal_launcher()
    return render_template_string(PORTAL_UPDATE_HTML, versions=versions, active=active)

@app.route('/developer/portal-update/upload', methods=['POST'])
def portal_update_upload():
    if not _deployment_access(): return 'NCSA0608 developer access required', 403
    upload = request.files.get('app_file'); version = _safe_version_name(request.form.get('version'))
    if not upload or not upload.filename.lower().endswith('.py') or not version:
        flash('Valid version name aur app.py select karein.','error'); return redirect(url_for('portal_update_manager'))
    manifest = _ensure_current_version_registered()
    if any(x.get('version') == version for x in manifest):
        flash('Ye version name already exists. Naya version name use karein.','error'); return redirect(url_for('portal_update_manager'))
    raw = upload.read(PORTAL_MAX_UPLOAD_BYTES + 1)
    if len(raw) > PORTAL_MAX_UPLOAD_BYTES:
        flash('app.py 8 MB se badi nahi ho sakti.','error'); return redirect(url_for('portal_update_manager'))
    try:
        text = raw.decode('utf-8'); compile(text, 'app.py', 'exec')
    except Exception as exc:
        flash('Python validation failed: ' + str(exc),'error'); return redirect(url_for('portal_update_manager'))
    if 'Flask' not in text or 'app =' not in text:
        flash('Uploaded file Flask app.py jaisi nahi lag rahi. Flask aur app object required hain.','error'); return redirect(url_for('portal_update_manager'))
    version_dir = os.path.join(PORTAL_VERSIONS_DIR, version); os.makedirs(version_dir, exist_ok=False)
    target = os.path.join(version_dir, 'app.py')
    with open(target, 'wb') as f: f.write(raw)
    digest = hashlib.sha256(raw).hexdigest()
    record={'version':version,'file':'app.py','sha256':digest,'size':len(raw),'uploaded_at':datetime.now().isoformat(timespec='seconds'),'uploaded_by':session.get('user_id'),'notes':request.form.get('notes','').strip(),'status':'ready'}
    manifest.append(record); _deployment_save(PORTAL_MANIFEST_FILE, manifest)
    if request.form.get('activate') == 'yes':
        _deployment_save(PORTAL_ACTIVE_FILE, {'version':version,'activated_at':datetime.now().isoformat(timespec='seconds'),'activated_by':session.get('user_id')})
        flash(f'{version} validated, deployed aur default live version set ho gaya. Restart Active Version click karein.','success')
    else: flash(f'{version} validated aur version folder me deploy ho gaya.','success')
    return redirect(url_for('portal_update_manager'))

@app.route('/developer/portal-update/activate/<version>', methods=['POST'])
def portal_update_activate(version):
    if not _deployment_access(): return 'NCSA0608 developer access required', 403
    version = _safe_version_name(version); manifest = _ensure_current_version_registered()
    if not any(x.get('version') == version for x in manifest) or not os.path.isfile(os.path.join(PORTAL_VERSIONS_DIR, version, 'app.py')):
        flash('Selected version file not found.','error'); return redirect(url_for('portal_update_manager'))
    _deployment_save(PORTAL_ACTIVE_FILE, {'version':version,'activated_at':datetime.now().isoformat(timespec='seconds'),'activated_by':session.get('user_id')})
    flash(f'{version} default live version set ho gaya. Restart Active Version click karein.','success')
    return redirect(url_for('portal_update_manager'))

@app.route('/developer/portal-update/restart', methods=['POST'])
def portal_update_restart():
    if not _deployment_access(): return 'NCSA0608 developer access required', 403
    launcher = _ensure_portal_launcher()
    if not os.path.isfile(launcher):
        flash('Launcher automatically create nahi ho saka. Server folder ki write permission check karein.','error'); return redirect(url_for('portal_update_manager'))
    kwargs={'cwd':PORTAL_BASE_DIR,'close_fds':True}
    if os.name == 'nt': kwargs['creationflags'] = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
    subprocess.Popen([sys.executable, launcher, '--wait', '2'], **kwargs)
    threading.Timer(0.8, lambda: os._exit(0)).start()
    return render_template_string('<script>setTimeout(()=>location.href="/login",5000)</script><main style="font-family:sans-serif;padding:40px"><h1>Portal restarting...</h1><p>Active version 5 seconds me start hogi.</p></main>')
# ===== END VERSIONED PORTAL DEPLOYMENT MANAGER =====

@app.route('/shutdown')
def shutdown():
    dev_pass = os.getenv('DEV_PWD', 'Shama@8577')
    if session.get('role') in ['admin', 'developer'] and request.args.get('pwd') == dev_pass:
        func = request.environ.get('werkzeug.server.shutdown')
        if func:
            func()
            return "Server successfully shutdown ho gaya hai."
        else:
            # Modern Werkzeug fallback
            sys.exit(0)
    return "Unauthorized access!", 403

def read_biometric_attendance_logs():
    zk = ZK(MACHINE_IP, port=PORT, timeout=10, password=0, force_udp=False, ommit_ping=False)
    conn = zk.connect()
    if not conn:
        raise ConnectionError('Could not connect to the biometric device')

    try:
        attendance = conn.get_attendance() or []
        return [
            {
                'user_id': str(record.user_id),
                'timestamp': record.timestamp.strftime('%Y-%m-%d %H:%M:%S')
            }
            for record in attendance
        ]
    finally:
        conn.disconnect()

def merge_synced_attendance_logs(logs):
    global LAST_DEVICE_SYNC_TIME
    existing_keys = {
        (str(item.get('store_id', DEFAULT_STORE_ID)), str(item.get('user_id')), str(item.get('timestamp')))
        for item in SYNCED_ATTENDANCE_LOGS
    }
    added_count = 0
    for log in logs:
        if not isinstance(log, dict) or 'user_id' not in log or 'timestamp' not in log:
            continue
        key = (str(log.get('store_id', DEFAULT_STORE_ID)), str(log['user_id']), str(log['timestamp']))
        if key not in existing_keys:
            SYNCED_ATTENDANCE_LOGS.append(log)
            existing_keys.add(key)
            added_count += 1

    LAST_DEVICE_SYNC_TIME = datetime.now()
    if added_count:
        temp = ATTENDANCE_LOGS_FILE + '.tmp'
        os.makedirs(os.path.dirname(os.path.abspath(ATTENDANCE_LOGS_FILE)), exist_ok=True)
        with open(temp, 'w', encoding='utf-8') as f:
            json.dump(SYNCED_ATTENDANCE_LOGS, f, indent=2, ensure_ascii=False, default=str)
        os.replace(temp, ATTENDANCE_LOGS_FILE)
    return added_count

def sync_biometric_attendance():
    with ATTENDANCE_SYNC_LOCK:
        logs = read_biometric_attendance_logs()
        added_count = merge_synced_attendance_logs(logs)
        return len(logs), added_count

def background_attendance_sync_loop():
    while True:
        try:
            total_count, added_count = sync_biometric_attendance()
            print(f"Background biometric sync: {total_count} records read ({added_count} new)")
        except Exception as e:
            print(f"Background biometric sync failed: {e}")
        threading.Event().wait(ATTENDANCE_SYNC_INTERVAL_SECONDS)

def start_background_attendance_sync():
    global BACKGROUND_SYNC_THREAD
    with BACKGROUND_SYNC_START_LOCK:
        if BACKGROUND_SYNC_THREAD and BACKGROUND_SYNC_THREAD.is_alive():
            return
        BACKGROUND_SYNC_THREAD = threading.Thread(
            target=background_attendance_sync_loop,
            name='biometric-attendance-sync',
            daemon=True
        )
        BACKGROUND_SYNC_THREAD.start()

@app.before_request
def ensure_background_attendance_sync():
    if not app.testing and not request.environ.get('werkzeug.test'):
        start_background_attendance_sync()
        start_backup_scheduler()

# BIOMETRIC LOCAL SYNC ENDPOINT
@app.route('/api/attendance/sync', methods=['POST'])
def sync_attendance():
    data = request.get_json(silent=True) or {}
    if 'logs' in data:
        logs = data['logs']
        store_id = safe_store_id(data.get('store_id'))
        store = STORES.get(store_id)
        if not isinstance(logs, list):
            return {'status': 'error', 'message': 'Logs must be a list'}, 400
        if not store or not secrets.compare_digest(str(data.get('bridge_token', '')), str(store.get('bridge_token', ''))):
            return {'status': 'error', 'message': 'Invalid store bridge credentials'}, 403
        tagged_logs = [dict(item, store_id=store_id) for item in logs if isinstance(item, dict) and item.get('user_id') and item.get('timestamp')]
        with ATTENDANCE_SYNC_LOCK:
            added_count = merge_synced_attendance_logs(tagged_logs)
        return {
            'status': 'success',
            'message': f'{len(logs)} records received ({added_count} new)'
        }, 200

    if not session.get('logged_in') or session.get('role') not in ['admin', 'developer']:
        return {'status': 'error', 'message': 'Admin login required to sync the biometric device'}, 403

    try:
        total_count, added_count = sync_biometric_attendance()
        print(f"Manual biometric sync: {total_count} records read ({added_count} new)")
        return {
            'status': 'success',
            'message': f'{total_count} biometric records synced ({added_count} new)'
        }, 200
    except Exception as e:
        print(f"Error in sync_attendance: {e}")
        return {'status': 'error', 'message': str(e)}, 502

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=os.getenv('FLASK_DEBUG', '0') == '1')
