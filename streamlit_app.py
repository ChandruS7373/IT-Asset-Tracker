import streamlit as st
import sqlite3, os, json, smtplib, configparser, base64
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

st.set_page_config(page_title="Qualesce IT Asset Tracker", page_icon="💻",
                   layout="wide", initial_sidebar_state="expanded")

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DB_PATH     = os.path.join(BASE_DIR, 'data', 'ittracker.db')
CONFIG_PATH = os.path.join(BASE_DIR, 'email_config.ini')
EXCEL_PATH  = os.path.join(BASE_DIR, 'data', 'assets.xlsx')
os.makedirs(os.path.join(BASE_DIR, 'data'), exist_ok=True)

# ─── Theme injection ────────────────────────────────────────────────────────────

def inject_theme():
    logo_file = os.path.join(BASE_DIR, 'static', 'logo.jpg')
    logo_b64  = base64.b64encode(open(logo_file,'rb').read()).decode() if os.path.exists(logo_file) else ""

    # Inline all CSS — no CDN dependency (Streamlit Cloud CSP blocks external stylesheets)
    st.markdown("""
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css"
      crossorigin="anonymous" referrerpolicy="no-referrer">
<style>
/* ════ Streamlit chrome ════ */
#MainMenu,footer,[data-testid="stToolbar"],[data-testid="stDecoration"],
[data-testid="stHeader"]{display:none !important;}
.block-container{padding:0 !important;max-width:100% !important;}

/* ════ Sidebar dark gradient ════ */
[data-testid="stSidebar"]{
    background:linear-gradient(180deg,#0f3460 0%,#1a1a2e 100%) !important;
    border-right:none !important;min-width:245px !important;max-width:260px !important;
}
[data-testid="stSidebar"] *{color:#e2e8f0 !important;}
[data-testid="stSidebar"] hr{border-color:rgba(255,255,255,.15) !important;}
[data-testid="stSidebarContent"]{padding:0 !important;}
/* sidebar nav buttons */
[data-testid="stSidebar"] .stButton button{
    width:100% !important;background:transparent !important;
    color:#cbd5e1 !important;border:none !important;
    text-align:left !important;padding:9px 16px !important;
    border-radius:8px !important;font-size:.88rem !important;
    margin-bottom:2px !important;transition:all .2s !important;
}
[data-testid="stSidebar"] .stButton button:hover{
    background:rgba(255,255,255,.12) !important;color:#fff !important;
    transform:translateX(4px) !important;
}

/* ════ Main area ════ */
section.main .block-container,.main .block-container{padding:1.5rem 2rem !important;}
/* main action buttons */
section.main .stButton button,.main .stButton button{
    background:linear-gradient(135deg,#0f3460,#533483) !important;
    color:#fff !important;border:none !important;
    border-radius:8px !important;font-weight:600 !important;
    padding:8px 20px !important;transition:all .2s !important;
}
section.main .stButton button:hover,.main .stButton button:hover{
    transform:translateY(-1px) !important;
    box-shadow:0 4px 15px rgba(83,52,131,.4) !important;
}

/* ════ Inputs ════ */
.stTextInput input,.stNumberInput input{
    border:1.5px solid #e2e8f0 !important;border-radius:8px !important;
}
.stTextInput input:focus{
    border-color:#533483 !important;box-shadow:0 0 0 3px rgba(83,52,131,.1) !important;
}
.stTextArea textarea{border:1.5px solid #e2e8f0 !important;border-radius:8px !important;}
.stSelectbox>div>div{border:1.5px solid #e2e8f0 !important;border-radius:8px !important;}

/* ════ Tabs ════ */
.stTabs [data-baseweb="tab-list"]{background:#f1f5f9;border-radius:10px;padding:4px;gap:4px;}
.stTabs [data-baseweb="tab"]{border-radius:8px;padding:6px 18px;font-weight:600;}
.stTabs [aria-selected="true"]{
    background:linear-gradient(135deg,#0f3460,#533483) !important;color:#fff !important;
}

/* ════ Streamlit widget overrides ════ */
[data-testid="stDataFrame"]{border-radius:12px !important;overflow:hidden !important;}
.stAlert{border-radius:12px !important;}
[data-testid="stMetricValue"]{font-size:2rem !important;font-weight:800 !important;}

/* ════ Panel component ════ */
.panel{background:#fff;border-radius:16px;box-shadow:0 2px 15px rgba(0,0,0,.06);
    border:1px solid #f1f5f9;margin-bottom:20px;overflow:hidden;}
.panel-header{display:flex;align-items:center;justify-content:space-between;
    padding:16px 20px;border-bottom:1px solid #f1f5f9;}
.panel-title{font-weight:700;color:#1e293b;font-size:1rem;}
.panel-body{padding:20px;}
.panel-body.p-0{padding:0;}
.topbar{padding:12px 2rem;border-bottom:1px solid #f1f5f9;margin-bottom:1.5rem;}
.topbar-title{font-size:1.1rem;font-weight:700;color:#1e293b;}

/* ════ Stat cards ════ */
.stat-card{border-radius:14px;padding:20px;text-align:center;
    border:1px solid rgba(0,0,0,.05);background:linear-gradient(135deg,#f0f4ff,#e8effe);}
.stat-card .icon{width:48px;height:48px;border-radius:12px;display:flex;
    align-items:center;justify-content:center;color:#fff;font-size:1.3rem;margin:0 auto 12px;}
.stat-card .value,.stat-card .val{font-size:2rem;font-weight:800;line-height:1;color:#1e293b;}
.stat-card .label,.stat-card .lbl{color:#64748b;font-size:.75rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.05em;margin-top:4px;}
.card-primary{background:linear-gradient(135deg,#eef2ff,#e0e7ff);}
.card-primary .value{color:#4f46e5;}
.card-icon-primary{background:linear-gradient(135deg,#4f46e5,#6366f1);}
.card-success{background:linear-gradient(135deg,#f0fdf4,#dcfce7);}
.card-success .value{color:#16a34a;}
.card-icon-success{background:linear-gradient(135deg,#16a34a,#22c55e);}
.card-info{background:linear-gradient(135deg,#eff6ff,#dbeafe);}
.card-info .value{color:#2563eb;}
.card-icon-info{background:linear-gradient(135deg,#2563eb,#3b82f6);}
.card-danger{background:linear-gradient(135deg,#fef2f2,#fee2e2);}
.card-danger .value{color:#dc2626;}
.card-icon-danger{background:linear-gradient(135deg,#dc2626,#ef4444);}
.card-warning{background:linear-gradient(135deg,#fffbeb,#fef3c7);}
.card-warning .value{color:#d97706;}
.card-icon-warning{background:linear-gradient(135deg,#d97706,#f59e0b);}
.card-purple{background:linear-gradient(135deg,#faf5ff,#f3e8ff);}
.card-purple .value{color:#7c3aed;}
.card-icon-purple{background:linear-gradient(135deg,#7c3aed,#8b5cf6);}

/* ════ Tables ════ */
.table-modern{width:100%;border-collapse:collapse;margin:0;}
.table-modern thead{background:#f8fafc;}
.table-modern thead th{padding:10px 14px;color:#64748b;font-size:.72rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid #e2e8f0;}
.table-modern tbody td{padding:11px 14px;border-bottom:1px solid #f8fafc;vertical-align:middle;}
.table-modern tbody tr:hover{background:#f8fafc;}

/* ════ Badges (Bootstrap-compatible, inline fallback) ════ */
.badge{display:inline-block;padding:.3em .65em;font-size:.72rem;font-weight:700;
    line-height:1;text-align:center;white-space:nowrap;vertical-align:baseline;border-radius:.4rem;}
.bg-warning{background-color:#ffc107 !important;color:#212529 !important;}
.bg-success{background-color:#198754 !important;color:#fff !important;}
.bg-danger{background-color:#dc3545 !important;color:#fff !important;}
.bg-primary{background-color:#0d6efd !important;color:#fff !important;}
.bg-secondary{background-color:#6c757d !important;color:#fff !important;}
.bg-info{background-color:#0dcaf0 !important;color:#212529 !important;}
.bg-warning.text-dark,.text-dark{color:#212529 !important;}
.bg-danger-subtle{background-color:#f8d7da !important;}
.bg-success-subtle{background-color:#d1e7dd !important;}
.text-danger{color:#dc3545 !important;}
.text-success{color:#198754 !important;}
.text-warning{color:#ffc107 !important;}

/* ════ Alerts ════ */
.alert{padding:.75rem 1rem;border-radius:.5rem;margin-bottom:1rem;border:1px solid transparent;}
.alert-info{background:#cff4fc;border-color:#b6effb;color:#055160;}
.alert-warning{background:#fff3cd;border-color:#ffe69c;color:#664d03;}
.alert-success{background:#d1e7dd;border-color:#badbcc;color:#0a3622;}

/* ════ Bootstrap layout utilities (inline fallback) ════ */
.d-flex{display:flex !important;}
.flex-column{flex-direction:column !important;}
.align-items-center{align-items:center !important;}
.align-items-start{align-items:flex-start !important;}
.justify-content-between{justify-content:space-between !important;}
.justify-content-end{justify-content:flex-end !important;}
.justify-content-center{justify-content:center !important;}
.gap-2{gap:.5rem !important;}
.gap-3{gap:1rem !important;}
.flex-fill{flex:1 1 auto !important;}
.flex-wrap{flex-wrap:wrap !important;}
.w-100{width:100% !important;}
/* spacing */
.mt-1{margin-top:.25rem !important;}.mt-2{margin-top:.5rem !important;}
.mt-3{margin-top:1rem !important;}.mt-4{margin-top:1.5rem !important;}
.mb-0{margin-bottom:0 !important;}.mb-1{margin-bottom:.25rem !important;}
.mb-2{margin-bottom:.5rem !important;}.mb-3{margin-bottom:1rem !important;}
.mb-4{margin-bottom:1.5rem !important;}.me-1{margin-right:.25rem !important;}
.me-2{margin-right:.5rem !important;}.ms-1{margin-left:.25rem !important;}
.ms-auto{margin-left:auto !important;}
.py-3{padding-top:1rem !important;padding-bottom:1rem !important;}
.py-4{padding-top:1.5rem !important;padding-bottom:1.5rem !important;}
.py-5{padding-top:3rem !important;padding-bottom:3rem !important;}
.p-2{padding:.5rem !important;}.p-3{padding:1rem !important;}
/* text */
.text-center{text-align:center !important;}.text-end{text-align:right !important;}
.text-muted{color:#6c757d !important;}.text-white{color:#fff !important;}
.fw-semibold{font-weight:600 !important;}.fw-bold{font-weight:700 !important;}
.small,.small *{font-size:.875rem !important;}.fs-5{font-size:1.25rem !important;}
.fs-6{font-size:1rem !important;}
/* misc */
.opacity-50{opacity:.5 !important;}.bg-light{background-color:#f8f9fa !important;}
.rounded{border-radius:.375rem !important;}.overflow-hidden{overflow:hidden !important;}
.word-break-break-word{word-break:break-word !important;}
/* grid — minimal col support */
.row{display:flex;flex-wrap:wrap;margin:0 -.5rem;}
.col-12{flex:0 0 100%;max-width:100%;padding:0 .5rem;}
/* table utils */
.table-responsive{overflow-x:auto;}
.table-sm td,.table-sm th{padding:.4rem .6rem;}
.table-bordered td,.table-bordered th{border:1px solid #dee2e6;}
.mb-0.table{margin-bottom:0;}
/* vertical align */
.vertical-middle,.align-middle{vertical-align:middle !important;}
</style>
""", unsafe_allow_html=True)
    return logo_b64

# ─── DB ────────────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH); conn.row_factory = sqlite3.Row; return conn

def init_db():
    conn = get_db(); c = conn.cursor()
    c.executescript('''
        CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'user',
            employee_name TEXT, email TEXT, department TEXT,
            notification_prefs TEXT DEFAULT '{}',
            created_at TEXT DEFAULT (datetime('now')));
        CREATE TABLE IF NOT EXISTS assets(
            id INTEGER PRIMARY KEY AUTOINCREMENT, asset_no TEXT UNIQUE,
            asset_status TEXT, login_id TEXT, serial_no TEXT,
            transfer_history TEXT, model TEXT, years TEXT, resolution TEXT,
            warranty_start TEXT, warranty_end TEXT, warranty_status TEXT,
            processor TEXT, ram TEXT, hdd TEXT, lan_ip TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now')));
        CREATE TABLE IF NOT EXISTS service_history(
            id INTEGER PRIMARY KEY AUTOINCREMENT, asset_no TEXT, service_date TEXT,
            service_type TEXT, description TEXT, technician TEXT,
            status TEXT DEFAULT 'Completed', cost TEXT, remarks TEXT, logged_by TEXT,
            created_at TEXT DEFAULT (datetime('now')));
        CREATE TABLE IF NOT EXISTS service_requests(
            id INTEGER PRIMARY KEY AUTOINCREMENT, asset_no TEXT,
            employee_id INTEGER, employee_name TEXT, service_type TEXT,
            description TEXT, remarks TEXT, status TEXT DEFAULT 'Pending',
            assigned_to_id INTEGER, assigned_to_name TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now')));
        CREATE TABLE IF NOT EXISTS chat_messages(
            id INTEGER PRIMARY KEY AUTOINCREMENT, request_id INTEGER,
            sender_id INTEGER, sender_name TEXT, sender_role TEXT, message TEXT,
            created_at TEXT DEFAULT (datetime('now')));
        CREATE TABLE IF NOT EXISTS stock_dashboard(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            status TEXT NOT NULL, model TEXT NOT NULL, count INTEGER DEFAULT 0,
            UNIQUE(status,model));
    ''')
    try: c.execute("ALTER TABLE users ADD COLUMN notification_prefs TEXT DEFAULT '{}'")
    except: pass
    c.execute("SELECT id FROM users WHERE username='admin'")
    if not c.fetchone():
        c.execute("INSERT INTO users(username,password,role,employee_name) VALUES(?,?,?,?)",
                  ('admin','admin123','admin','Administrator'))
    conn.commit(); conn.close()

# ─── Email ─────────────────────────────────────────────────────────────────────

def get_smtp(): cfg=configparser.ConfigParser();cfg.read(CONFIG_PATH);return cfg['SMTP'] if 'SMTP' in cfg else {}
def get_nc():
    cfg=configparser.ConfigParser();cfg.read(CONFIG_PATH)
    s=cfg['NOTIFICATIONS'] if 'NOTIFICATIONS' in cfg else {}
    b=lambda k:s.get(k,'true').lower()!='false'
    return {k:b(k) for k in ['notify_new_request','notify_employee_action',
        'notify_technician_assigned','notify_status_employee','notify_status_admin','notify_chat']}

def send_email(to, subj, html):
    c=get_smtp();h=c.get('smtp_host','');p=int(c.get('smtp_port',587))
    u=c.get('smtp_user','');pw=c.get('smtp_password','');fn=c.get('from_name','IT Tracker')
    if not u: return False,'SMTP not configured'
    if isinstance(to,str): to=[to]
    to=[e for e in to if e]
    if not to: return False,'No recipients'
    msg=MIMEMultipart('alternative');msg['Subject']=subj;msg['From']=f'{fn} <{u}>';msg['To']=', '.join(to)
    msg.attach(MIMEText(html,'html'))
    try:
        with smtplib.SMTP(h,p,timeout=10) as s: s.ehlo();s.starttls();s.login(u,pw);s.sendmail(u,to,msg.as_string())
        return True,'Sent'
    except Exception as e: return False,str(e)

def status_email_html(name,status,rid,stype,asset):
    c={'In Progress':'#f59e0b','Hold':'#6b7280','Completed':'#22c55e','Accepted':'#22c55e','Declined':'#ef4444'}.get(status,'#374151')
    return f"<div style='font-family:Segoe UI,sans-serif;max-width:540px;margin:auto'><div style='background:linear-gradient(135deg,#0f3460,#533483);padding:22px;text-align:center'><h2 style='color:#fff;margin:0'>Request #{rid} — {status}</h2></div><div style='padding:24px;border:1px solid #e2e8f0'><p>Hi <b>{name}</b>, your request status is now <b style='color:{c}'>{status}</b>.</p><p>Asset: {asset} | Issue: {stype}</p></div></div>"

def chat_email_html(sender,role,msg,rid,stype):
    rl={'admin':'Admin','technician':'Technician','user':'Employee'}.get(role,role)
    return f"<div style='font-family:Segoe UI,sans-serif;max-width:540px;margin:auto'><div style='background:linear-gradient(135deg,#0f3460,#533483);padding:22px;text-align:center'><h2 style='color:#fff;margin:0'>New Message — Request #{rid}</h2></div><div style='padding:24px;border:1px solid #e2e8f0'><p>From <b>{sender}</b> ({rl}) on Request #{rid} ({stype}):</p><div style='background:#f1f5f9;border-left:4px solid #533483;padding:12px;margin:14px 0'>{msg}</div></div></div>"

# ─── HTML helpers ──────────────────────────────────────────────────────────────

def panel(title, icon, body_html, actions_html=""):
    hdr_actions = f'<div class="d-flex gap-2">{actions_html}</div>' if actions_html else ""
    st.markdown(f"""
    <div class="panel">
      <div class="panel-header d-flex justify-content-between align-items-center">
        <div class="panel-title"><i class="{icon}"></i> {title}</div>
        {hdr_actions}
      </div>
      <div class="panel-body">{body_html}</div>
    </div>""", unsafe_allow_html=True)

STATUS_BADGE = {
    'Pending':'<span class="badge bg-warning text-dark">Pending</span>',
    'Accepted':'<span class="badge bg-success">Accepted</span>',
    'Declined':'<span class="badge bg-danger">Declined</span>',
    'In Progress':'<span class="badge bg-primary">In Progress</span>',
    'Hold':'<span class="badge bg-secondary">Hold</span>',
    'Completed':'<span class="badge bg-success">Completed</span>',
}
def sbadge(s): return STATUS_BADGE.get(s, f'<span class="badge bg-secondary">{s}</span>')

# ─── Auth ──────────────────────────────────────────────────────────────────────

def page_login(logo_b64):
    logo_tag = f'<img src="data:image/jpeg;base64,{logo_b64}" style="height:40px;margin-bottom:16px">' if logo_b64 else ""
    st.markdown(f"""
    <div style="min-height:100vh;background:linear-gradient(135deg,#0f3460 0%,#533483 100%);
         display:flex;align-items:center;justify-content:center;padding:40px 16px">
      <div style="background:#fff;border-radius:20px;box-shadow:0 25px 50px rgba(0,0,0,.25);
           padding:40px 36px;width:100%;max-width:420px;text-align:center">
        {logo_tag}
        <h2 style="color:#0f3460;font-weight:800;margin-bottom:4px">IT Asset Tracker</h2>
        <p style="color:#64748b;margin-bottom:28px">Qualesce Technology Solutions</p>
      </div>
    </div>""", unsafe_allow_html=True)

    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        with st.form("login_form"):
            u = st.text_input("", placeholder="Username or Email", label_visibility="collapsed")
            p = st.text_input("", placeholder="Password", type="password", label_visibility="collapsed")
            btn = st.form_submit_button("Sign In →", use_container_width=True)
        if btn:
            conn = get_db()
            user = conn.execute(
                "SELECT * FROM users WHERE (LOWER(email)=? OR LOWER(username)=?) AND password=?",
                (u.strip().lower(), u.strip().lower(), p.strip())).fetchone()
            conn.close()
            if user:
                st.session_state.update(logged_in=True, user_id=user['id'],
                    username=user['username'], role=user['role'],
                    employee_name=user['employee_name'] or user['username'],
                    user_email=user['email'] or '', page='dashboard')
                st.rerun()
            else:
                st.error("Invalid username or password.")

def do_logout():
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

# ─── Sidebar ───────────────────────────────────────────────────────────────────

def sidebar_nav(logo_b64):
    role = st.session_state.role
    name = st.session_state.employee_name

    logo_tag = f'<img src="data:image/jpeg;base64,{logo_b64}" style="height:28px;object-fit:contain;max-width:100%">' if logo_b64 else "💻"
    badge_style = {'admin':'background:#f59e0b;color:#1e293b','technician':'background:#22c55e;color:#fff','user':'background:#38bdf8;color:#fff'}.get(role,'')
    badge_label = role.upper()

    st.sidebar.markdown(f"""
    <div style="padding:14px;border-bottom:1px solid rgba(255,255,255,.1);text-align:center">
      <div style="background:#fff;border-radius:10px;padding:8px 14px">{logo_tag}</div>
    </div>
    <div style="padding:12px 14px;border-bottom:1px solid rgba(255,255,255,.1)">
      <div style="width:38px;height:38px;border-radius:50%;background:linear-gradient(135deg,#533483,#0f3460);
           display:inline-flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:1.1rem;margin-right:10px;vertical-align:middle">
        {name[0].upper()}
      </div>
      <span style="font-weight:600;font-size:.95rem;vertical-align:middle">{name}</span><br>
      <span style="display:inline-block;margin-top:4px;padding:2px 10px;border-radius:20px;font-size:.7rem;font-weight:700;{badge_style}">{badge_label}</span>
    </div>
    """, unsafe_allow_html=True)

    def nav(label, page):
        if st.sidebar.button(label, key=f'nav_{page}', use_container_width=True):
            st.session_state.page = page
            st.session_state.pop('req_id', None)
            st.rerun()

    if role == 'admin':
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">Dashboard</div>', unsafe_allow_html=True)
        nav('📊  Dashboard', 'dashboard')
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">Assets</div>', unsafe_allow_html=True)
        nav('💻  All Assets', 'assets')
        nav('👥  Employees', 'employees')
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">Service</div>', unsafe_allow_html=True)
        nav('📥  Service Requests', 'service_requests')
        nav('🔧  Service History', 'services')
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">Management</div>', unsafe_allow_html=True)
        nav('👤  Manage Users', 'users')
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">Settings</div>', unsafe_allow_html=True)
        nav('📧  Email Config', 'email_config')
    elif role == 'technician':
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">My Tasks</div>', unsafe_allow_html=True)
        nav('📊  Dashboard', 'dashboard')
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">Settings</div>', unsafe_allow_html=True)
        nav('🔔  Email Preferences', 'email_prefs')
    else:
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">My Portal</div>', unsafe_allow_html=True)
        nav('🏠  My Dashboard', 'dashboard')
        nav('🎫  My Requests', 'my_requests')
        nav('🔧  Request Service', 'request_service')
        st.sidebar.markdown('<div style="padding:6px 14px 2px;font-size:.7rem;font-weight:700;letter-spacing:.08em;color:#94a3b8;text-transform:uppercase">Settings</div>', unsafe_allow_html=True)
        nav('🔔  Email Preferences', 'email_prefs')

    st.sidebar.markdown('<hr style="margin:10px 0;border-color:rgba(255,255,255,.1)">', unsafe_allow_html=True)
    if st.sidebar.button('🚪  Logout', use_container_width=True): do_logout()

# ─── Admin Dashboard ───────────────────────────────────────────────────────────

def pg_admin_dashboard():
    conn = get_db()
    sc = {r[0]:r[1] for r in conn.execute("SELECT asset_status,COUNT(*) FROM assets GROUP BY asset_status").fetchall()}
    total   = sum(sc.values())
    stats   = {
        'total':    total,
        'assigned': sc.get('Assigned',0),
        'stock':    sc.get('IT Stock',0)+sc.get('ITStock',0),
        'dead':     sc.get('DEAD',0),
        'to_check': sc.get('To Check',0),
        'yashwanth':sc.get('Yashwanth',0),
        'users':    conn.execute("SELECT COUNT(*) FROM users WHERE role='user'").fetchone()[0],
        'services': conn.execute("SELECT COUNT(*) FROM service_history").fetchone()[0],
    }
    pending = conn.execute("SELECT COUNT(*) FROM service_requests WHERE status='Pending'").fetchone()[0]
    stock_rows = conn.execute("SELECT status,model,count FROM stock_dashboard ORDER BY status,model").fetchall()
    recent     = conn.execute("SELECT * FROM assets ORDER BY updated_at DESC LIMIT 10").fetchall()
    conn.close()

    # ── Page title
    st.markdown('<div class="topbar"><div class="topbar-title"><i class="fas fa-tachometer-alt me-2 text-muted"></i>Admin Dashboard</div></div>', unsafe_allow_html=True)

    # ── Stat cards (exact same HTML as Flask templates)
    def card(key, label, icon, cls, icls, extra_style="", val_style=""):
        v = stats.get(key, 0)
        if key == 'yashwanth' and v == 0: return ""
        if extra_style:
            return f'<div class="col-6 col-md-4 col-xl-2"><div class="stat-card" style="{extra_style}"><div class="icon" style="{icls}"><i class="{icon}"></i></div><div class="value" style="{val_style}">{v}</div><div class="label">{label}</div></div></div>'
        return f'<div class="col-6 col-md-4 col-xl-2"><div class="stat-card {cls}"><div class="icon {icls}"><i class="{icon}"></i></div><div class="value">{v}</div><div class="label">{label}</div></div></div>'

    cards = "".join([
        card('total',    'Total Assets',   'fas fa-laptop',      'card-primary', 'card-icon-primary'),
        card('assigned', 'Assigned',       'fas fa-user-check',  'card-success', 'card-icon-success'),
        card('stock',    'In Stock',       'fas fa-warehouse',   'card-info',    'card-icon-primary'),
        card('dead',     'Dead/Retired',   'fas fa-times-circle','card-danger',  'card-icon-danger'),
        card('to_check', 'To Check',       'fas fa-search',      '', '',
             'background:linear-gradient(135deg,#f0f4ff,#e8effe)',
             'linear-gradient(135deg,#f59e0b,#fbbf24)', ),
        card('yashwanth','Yashwanth',      'fas fa-user-shield', '', '',
             'background:linear-gradient(135deg,#f0fdf4,#dcfce7)',
             'linear-gradient(135deg,#8b5cf6,#a78bfa)'),
        card('users',    'Employees',      'fas fa-users',       'card-warning', 'card-icon-warning'),
        card('services', 'Service Records','fas fa-wrench',      'card-purple',  'card-icon-primary'),
    ])

    # Fix: custom cards need different format
    def card2(key, label, icon, extra_bg, icon_bg, val_color):
        v = stats.get(key,0)
        if key=='yashwanth' and v==0: return ""
        return f'<div class="col-6 col-md-4 col-xl-2"><div class="stat-card" style="background:{extra_bg}"><div class="icon" style="background:{icon_bg}"><i class="{icon}"></i></div><div class="value" style="color:{val_color}">{v}</div><div class="label">{label}</div></div></div>'

    st.markdown(f"""
    <div class="content-area" style="padding:1.5rem 2rem">
    <div class="row g-3 mb-4">
      {card('total','Total Assets','fas fa-laptop','card-primary','card-icon-primary')}
      {card('assigned','Assigned','fas fa-user-check','card-success','card-icon-success')}
      {card('stock','In Stock','fas fa-warehouse','card-info','card-icon-primary')}
      {card('dead','Dead/Retired','fas fa-times-circle','card-danger','card-icon-danger')}
      {card2('to_check','To Check','fas fa-search','linear-gradient(135deg,#f0f4ff,#e8effe)','linear-gradient(135deg,#f59e0b,#fbbf24)','#92400e')}
      {card2('yashwanth','Yashwanth','fas fa-user-shield','linear-gradient(135deg,#f0fdf4,#dcfce7)','linear-gradient(135deg,#8b5cf6,#a78bfa)','#5b21b6')}
      {card('users','Employees','fas fa-users','card-warning','card-icon-warning')}
      {card('services','Service Records','fas fa-wrench','card-purple','card-icon-primary')}
    </div>
    </div>
    """, unsafe_allow_html=True)

    if pending:
        st.warning(f"⚠️ {pending} pending service request(s) waiting for review.")
        if st.button("View Service Requests →"):
            st.session_state.page='service_requests'; st.rerun()

    # ── Stock dashboard
    st.markdown('<div style="padding:0 2rem">', unsafe_allow_html=True)
    col_l, col_r = st.columns([5, 1])
    with col_l:
        st.markdown('<div class="panel-title" style="font-size:1rem;font-weight:700;padding:8px 0"><i class="fas fa-boxes me-2"></i>Stock Dashboard</div>', unsafe_allow_html=True)
    with col_r:
        if st.button("🔄 Sync"):
            _sync_stock(); st.rerun()

    if stock_rows:
        statuses,models,data=[],[],{}
        for r in stock_rows:
            if r['status'] not in statuses: statuses.append(r['status'])
            if r['model'] not in models: models.append(r['model'])
            data.setdefault(r['status'],{})[r['model']]=r['count']
        rows=[]
        for s in statuses:
            row={'Status':s};row.update({m:data.get(s,{}).get(m,0) for m in models})
            row['Total']=sum(data.get(s,{}).get(m,0) for m in models);rows.append(row)
        gt={'Status':'Grand Total'}
        gt.update({m:sum(data.get(s,{}).get(m,0) for s in statuses) for m in models})
        gt['Total']=sum(gt[m] for m in models);rows.append(gt)
        df=pd.DataFrame(rows).set_index('Status')
        st.dataframe(df, use_container_width=True)
    else:
        st.info("No stock data. Click Sync to populate from assets.")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Recent assets
    st.markdown('<div style="padding:1rem 2rem 0">', unsafe_allow_html=True)
    if recent:
        rows_html="".join(f"""
        <tr>
          <td><span class="fw-semibold">{a['asset_no']}</span></td>
          <td>{a['login_id'] or '—'}</td>
          <td>{a['model'] or '—'}</td>
          <td>{sbadge(a['asset_status']) if a['asset_status'] else '—'}</td>
          <td><span class="badge {'bg-danger-subtle text-danger' if a['warranty_status']=='Out of warranty' else 'bg-success-subtle text-success' if a['warranty_status'] else 'bg-secondary'}">{a['warranty_status'] or '—'}</span></td>
        </tr>""" for a in recent)
        st.markdown(f"""
        <div class="panel">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-clock"></i> Recently Updated Assets</div></div>
          <div class="panel-body p-0">
            <div class="table-responsive">
              <table class="table table-modern mb-0">
                <thead><tr><th>Asset No</th><th>Employee</th><th>Model</th><th>Status</th><th>Warranty</th></tr></thead>
                <tbody>{rows_html}</tbody>
              </table>
            </div>
          </div>
        </div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

def _sync_stock():
    from collections import defaultdict
    conn=get_db();agg=defaultdict(int)
    for r in conn.execute("SELECT asset_status,model,COUNT(*) FROM assets WHERE asset_status!='' AND model!='' GROUP BY asset_status,model").fetchall():
        s=r[0].strip();s='IT Stock' if s=='ITStock' else s;m=r[1].strip()
        if m: agg[(s,m)]+=r[2]
    conn.execute("DELETE FROM stock_dashboard")
    for (s,m),cnt in agg.items(): conn.execute("INSERT INTO stock_dashboard(status,model,count) VALUES(?,?,?)",(s,m,cnt))
    conn.commit();conn.close()

# ─── Admin Service Requests ────────────────────────────────────────────────────

def pg_admin_service_requests():
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    st.markdown('<h4><i class="fas fa-inbox me-2 text-muted"></i>Service Requests</h4>', unsafe_allow_html=True)
    sf = st.selectbox("", ['All','Pending','Accepted','In Progress','Hold','Completed','Declined'], label_visibility="collapsed")
    conn=get_db()
    q="SELECT * FROM service_requests WHERE 1=1"+(" AND status=?" if sf!='All' else "")+" ORDER BY created_at DESC"
    reqs=conn.execute(q,[sf] if sf!='All' else []).fetchall();conn.close()

    if not reqs:
        st.markdown('<div class="panel"><div class="panel-body text-center text-muted py-5"><i class="fas fa-inbox fa-3x mb-3 opacity-50"></i><p>No requests found.</p></div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True); return

    rows_html=""
    for r in reqs:
        rows_html+=f"""
        <tr>
          <td class="fw-bold text-muted small">#{r['id']}</td>
          <td class="fw-semibold">{r['employee_name']}</td>
          <td>{r['asset_no'] or '—'}</td>
          <td class="small fw-semibold">{r['service_type']}</td>
          <td>{sbadge(r['status'])}</td>
          <td class="small text-muted">{r['assigned_to_name'] or '—'}</td>
          <td class="small text-muted">{r['created_at'][:10]}</td>
          <td><form method="get"><button class="btn btn-sm btn-outline-primary" name="manage" value="{r['id']}">Manage</button></form></td>
        </tr>"""

    st.markdown(f"""
    <div class="panel">
      <div class="panel-header"><div class="panel-title"><i class="fas fa-inbox"></i> Requests ({len(reqs)})</div></div>
      <div class="panel-body p-0"><div class="table-responsive">
        <table class="table table-modern mb-0">
          <thead><tr><th>#</th><th>Employee</th><th>Asset</th><th>Issue</th><th>Status</th><th>Assigned</th><th>Date</th><th></th></tr></thead>
          <tbody>{rows_html}</tbody>
        </table>
      </div></div>
    </div>""", unsafe_allow_html=True)

    # Use Streamlit buttons for interactivity
    st.markdown("---")
    st.caption("Click a request ID to manage it:")
    cols = st.columns(min(len(reqs), 6))
    for i, r in enumerate(reqs[:12]):
        with cols[i % 6]:
            if st.button(f"#{r['id']} {r['employee_name'][:10]}", key=f"req_{r['id']}"):
                st.session_state.req_id=r['id']; st.session_state.page='req_detail'; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def pg_admin_req_detail():
    req_id=st.session_state.get('req_id')
    if not req_id: st.session_state.page='service_requests';st.rerun();return
    conn=get_db()
    req=conn.execute("SELECT * FROM service_requests WHERE id=?",(req_id,)).fetchone()
    if not req: conn.close();st.error("Not found.");return
    msgs=conn.execute("SELECT * FROM chat_messages WHERE request_id=? ORDER BY created_at",(req_id,)).fetchall()
    techs=conn.execute("SELECT id,employee_name,username FROM users WHERE role='technician'").fetchall()
    conn.close()

    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    if st.button("← Back to Requests"): st.session_state.page='service_requests';st.rerun()

    st.markdown(f'<h4><i class="fas fa-inbox me-2 text-muted"></i>Request #{req_id}</h4>', unsafe_allow_html=True)

    left,right=st.columns([1,1.6])
    with left:
        st.markdown(f"""
        <div class="panel">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-info-circle"></i> Details</div></div>
          <div class="panel-body">
            <div class="text-center mb-3">{sbadge(req['status'])}</div>
            <table class="table table-sm small mb-0">
              <tr><td class="text-muted fw-semibold">Employee</td><td>{req['employee_name']}</td></tr>
              <tr><td class="text-muted fw-semibold">Asset</td><td>{req['asset_no'] or '—'}</td></tr>
              <tr><td class="text-muted fw-semibold">Issue Type</td><td>{req['service_type']}</td></tr>
              <tr><td class="text-muted fw-semibold">Submitted</td><td>{req['created_at'][:10]}</td></tr>
              <tr><td class="text-muted fw-semibold">Assigned To</td><td>{req['assigned_to_name'] or '—'}</td></tr>
            </table>
            <div class="mt-3 small bg-light rounded p-2">{req['description']}</div>
          </div>
        </div>""", unsafe_allow_html=True)

        if req['status']=='Pending':
            st.markdown('<div class="panel"><div class="panel-header"><div class="panel-title"><i class="fas fa-check-circle"></i> Review</div></div><div class="panel-body">', unsafe_allow_html=True)
            ca,cd=st.columns(2)
            if ca.button("✅ Accept",use_container_width=True,type="primary"): _set_status(req_id,req,'Accepted');st.rerun()
            if cd.button("❌ Decline",use_container_width=True): _set_status(req_id,req,'Declined');st.rerun()
            st.markdown('</div></div>', unsafe_allow_html=True)

        st.markdown('<div class="panel mt-3"><div class="panel-header"><div class="panel-title"><i class="fas fa-exchange-alt"></i> Change Status</div></div><div class="panel-body">', unsafe_allow_html=True)
        with st.form("sf"):
            opts=['Pending','Accepted','In Progress','Hold','Completed','Declined']
            ns=st.selectbox("",opts,index=opts.index(req['status']),label_visibility="collapsed")
            if st.form_submit_button("Update Status",use_container_width=True): _set_status(req_id,req,ns);st.rerun()
        st.markdown('</div></div>', unsafe_allow_html=True)

        if techs:
            st.markdown('<div class="panel mt-3"><div class="panel-header"><div class="panel-title"><i class="fas fa-user-hard-hat"></i> Assign Technician</div></div><div class="panel-body">', unsafe_allow_html=True)
            with st.form("tf"):
                to={t['id']:t['employee_name'] or t['username'] for t in techs}
                tid=st.selectbox("",list(to.keys()),format_func=lambda x:to[x],label_visibility="collapsed")
                if st.form_submit_button("Assign",use_container_width=True): _assign(req_id,req,tid,to[tid]);st.rerun()
            st.markdown('</div></div>', unsafe_allow_html=True)

    with right:
        msgs_html=""
        for m in msgs:
            is_admin=m['sender_role']=='admin'
            bg='#533483' if m['sender_role']=='admin' else '#0891b2' if m['sender_role']=='technician' else '#e2e8f0'
            tc='#fff' if m['sender_role'] in ('admin','technician') else '#1e293b'
            side='justify-content-end' if is_admin else ''
            r_badge=f'<span class="badge {"bg-warning text-dark" if m["sender_role"]=="admin" else "bg-info" if m["sender_role"]=="technician" else "bg-secondary"} ms-1" style="font-size:.65rem">{m["sender_role"].upper()}</span>'
            msgs_html+=f"""
            <div class="d-flex mb-3 {side}">
              <div style="max-width:75%">
                <div class="small text-muted mb-1 {'text-end' if is_admin else ''}">
                  {'You (Admin)' if is_admin else f'<strong>{m["sender_name"]}</strong>'} {r_badge} · {m['created_at'][11:16]}
                </div>
                <div style="padding:10px 14px;word-break:break-word;background:{bg};color:{tc};
                  border-radius:{'12px 12px 0 12px' if is_admin else '12px 12px 12px 0'}">
                  {m['message']}
                </div>
              </div>
            </div>"""

        st.markdown(f"""
        <div class="panel" style="min-height:500px;display:flex;flex-direction:column">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-comments"></i> Chat</div></div>
          <div style="flex:1;overflow-y:auto;padding:20px;background:#f8fafc;min-height:350px">
            {msgs_html if msgs_html else '<div class="text-center text-muted py-4"><i class="fas fa-comments fa-2x mb-2 opacity-50"></i><p class="small">No messages yet.</p></div>'}
          </div>
        </div>""", unsafe_allow_html=True)

        with st.form("cf",clear_on_submit=True):
            mi=st.text_area("",placeholder="Type your message...",height=80,label_visibility="collapsed")
            if st.form_submit_button("Send Message 📨",use_container_width=True):
                if mi.strip(): _chat(req_id,req,mi.strip(),'admin');st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def _set_status(req_id,req,status):
    conn=get_db()
    conn.execute("UPDATE service_requests SET status=?,updated_at=? WHERE id=?",(status,datetime.now().isoformat(),req_id))
    conn.commit()
    emp=conn.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
    conn.close()
    if emp and emp['email'] and get_nc()['notify_status_employee']:
        send_email([emp['email']],f"Request #{req_id}: {status}",status_email_html(req['employee_name'],status,req_id,req['service_type'],req['asset_no']))
    st.success(f"Status updated to {status}!")

def _assign(req_id,req,tech_id,tech_name):
    conn=get_db()
    conn.execute("UPDATE service_requests SET assigned_to_id=?,assigned_to_name=?,status='In Progress',updated_at=? WHERE id=?",(tech_id,tech_name,datetime.now().isoformat(),req_id))
    conn.commit()
    tech=conn.execute("SELECT email FROM users WHERE id=?",(tech_id,)).fetchone()
    emp=conn.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
    conn.close()
    nc=get_nc()
    if tech and tech['email'] and nc['notify_technician_assigned']:
        send_email([tech['email']],f"Request #{req_id} Assigned to You",f"<p>Request #{req_id} ({req['service_type']}) has been assigned to you.</p>")
    if emp and emp['email'] and nc['notify_status_employee']:
        send_email([emp['email']],f"Your Request #{req_id} is In Progress",status_email_html(req['employee_name'],'In Progress',req_id,req['service_type'],req['asset_no']))
    st.success(f"Assigned to {tech_name}!")

def _chat(req_id,req,message,sender_role):
    conn=get_db()
    conn.execute("INSERT INTO chat_messages(request_id,sender_id,sender_name,sender_role,message) VALUES(?,?,?,?,?)",
        (req_id,st.session_state.user_id,st.session_state.employee_name,sender_role,message))
    conn.commit()
    nc=get_nc()
    if nc['notify_chat']:
        if sender_role=='admin':
            emp=conn.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
            if emp and emp['email']: send_email([emp['email']],f"Admin replied on Request #{req_id}",chat_email_html(st.session_state.employee_name,'admin',message,req_id,req['service_type']))
        elif sender_role=='user':
            admins=conn.execute("SELECT email FROM users WHERE role='admin' AND email!=''").fetchall()
            send_email([a['email'] for a in admins if a['email']],f"Message on Request #{req_id}",chat_email_html(st.session_state.employee_name,'user',message,req_id,req['service_type']))
        elif sender_role=='technician':
            emp=conn.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
            admins=conn.execute("SELECT email FROM users WHERE role='admin' AND email!=''").fetchall()
            recip=([emp['email']] if emp and emp['email'] else [])+[a['email'] for a in admins if a['email']]
            if recip: send_email(recip,f"Tech message on Request #{req_id}",chat_email_html(st.session_state.employee_name,'technician',message,req_id,req['service_type']))
    conn.close()

# ─── Admin Assets / Employees / Services / Users ───────────────────────────────

def pg_admin_assets():
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    st.markdown('<h4><i class="fas fa-laptop me-2 text-muted"></i>All Assets</h4>', unsafe_allow_html=True)
    c1,c2=st.columns([3,1])
    search=c1.text_input("","",placeholder="🔍 Search employee, model, asset no...",label_visibility="collapsed")
    sf=c2.selectbox("",['All','Assigned','IT Stock','DEAD','To Check','Service','Yashwanth'],label_visibility="collapsed")
    conn=get_db()
    q="SELECT * FROM assets WHERE 1=1";p=[]
    if search: q+=" AND (login_id LIKE ? OR asset_no LIKE ? OR model LIKE ? OR serial_no LIKE ?)";p+=[f'%{search}%']*4
    if sf!='All': q+=" AND asset_status=?";p.append(sf)
    assets=conn.execute(q+" ORDER BY asset_no",p).fetchall();conn.close()
    if assets:
        rows="".join(f"""
        <tr><td><span class="fw-semibold">{a['asset_no']}</span></td>
        <td>{a['login_id'] or '—'}</td><td>{a['model'] or '—'}</td>
        <td>{sbadge(a['asset_status']) if a['asset_status'] else '—'}</td>
        <td>{a['serial_no'] or '—'}</td>
        <td><span class="badge {'bg-danger-subtle text-danger' if a['warranty_status']=='Out of warranty' else 'bg-success-subtle text-success' if a['warranty_status'] else 'bg-secondary'}">{a['warranty_status'] or '—'}</span></td>
        </tr>""" for a in assets)
        st.markdown(f"""
        <div class="panel">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-laptop"></i> All Assets ({len(assets)})</div></div>
          <div class="panel-body p-0"><div class="table-responsive">
            <table class="table table-modern mb-0">
              <thead><tr><th>Asset No</th><th>Employee</th><th>Model</th><th>Status</th><th>Serial No</th><th>Warranty</th></tr></thead>
              <tbody>{rows}</tbody>
            </table>
          </div></div>
        </div>""", unsafe_allow_html=True)
    else: st.info("No assets found.")
    st.markdown('</div>', unsafe_allow_html=True)

def pg_admin_employees():
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    st.markdown('<h4><i class="fas fa-users me-2 text-muted"></i>Employees</h4>', unsafe_allow_html=True)
    conn=get_db()
    emps=conn.execute("SELECT login_id,COUNT(*) as cnt,GROUP_CONCAT(model,', ') as models FROM assets WHERE login_id NOT IN ('IT Stock','DEAD','') AND login_id IS NOT NULL GROUP BY login_id ORDER BY login_id").fetchall()
    conn.close()
    search=st.text_input("","",placeholder="🔍 Search employee...",label_visibility="collapsed")
    rows="".join(f"""<tr><td class="fw-semibold">{e['login_id']}</td><td>{e['cnt']}</td><td class="small text-muted">{e['models'] or '—'}</td></tr>"""
                 for e in emps if not search or search.lower() in e['login_id'].lower())
    st.markdown(f"""
    <div class="panel">
      <div class="panel-header"><div class="panel-title"><i class="fas fa-users"></i> Employees</div></div>
      <div class="panel-body p-0"><div class="table-responsive">
        <table class="table table-modern mb-0">
          <thead><tr><th>Employee</th><th>Assets</th><th>Models</th></tr></thead>
          <tbody>{rows or '<tr><td colspan="3" class="text-center text-muted py-4">No employees found.</td></tr>'}</tbody>
        </table>
      </div></div>
    </div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

def pg_admin_services():
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    st.markdown('<h4><i class="fas fa-tools me-2 text-muted"></i>Service History</h4>', unsafe_allow_html=True)
    search=st.text_input("","",placeholder="🔍 Search...",label_visibility="collapsed")
    conn=get_db()
    q="SELECT sh.*,a.login_id FROM service_history sh LEFT JOIN assets a ON sh.asset_no=a.asset_no WHERE 1=1";p=[]
    if search: q+=" AND (sh.asset_no LIKE ? OR sh.description LIKE ?)";p+=[f'%{search}%']*2
    svcs=conn.execute(q+" ORDER BY sh.service_date DESC",p).fetchall();conn.close()
    rows="".join(f"""<tr><td>{s['service_date'] or '—'}</td><td class="fw-semibold">{s['asset_no']}</td>
    <td>{s['login_id'] or '—'}</td><td>{s['service_type'] or '—'}</td>
    <td><span class="badge {'bg-success' if s['status']=='Completed' else 'bg-warning text-dark' if s['status']=='Pending' else 'bg-secondary'}">{s['status']}</span></td>
    <td>{s['technician'] or '—'}</td></tr>""" for s in svcs)
    st.markdown(f"""
    <div class="panel">
      <div class="panel-header"><div class="panel-title"><i class="fas fa-tools"></i> Service Records ({len(svcs)})</div></div>
      <div class="panel-body p-0"><div class="table-responsive">
        <table class="table table-modern mb-0">
          <thead><tr><th>Date</th><th>Asset</th><th>Employee</th><th>Type</th><th>Status</th><th>Technician</th></tr></thead>
          <tbody>{rows or '<tr><td colspan="6" class="text-center py-4 text-muted">No records.</td></tr>'}</tbody>
        </table>
      </div></div>
    </div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

def pg_admin_users():
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    st.markdown('<h4><i class="fas fa-user-cog me-2 text-muted"></i>Manage Users</h4>', unsafe_allow_html=True)
    conn=get_db();users=conn.execute("SELECT * FROM users ORDER BY role,username").fetchall();conn.close()
    rows="".join(f"""<tr><td>{u['id']}</td><td class="fw-semibold">{u['username']}</td>
    <td>{u['employee_name'] or '—'}</td>
    <td><span class="badge {'bg-warning text-dark' if u['role']=='admin' else 'bg-success' if u['role']=='technician' else 'bg-info'}">{u['role'].upper()}</span></td>
    <td>{u['email'] or '—'}</td><td>{u['department'] or '—'}</td></tr>""" for u in users)
    st.markdown(f"""
    <div class="panel">
      <div class="panel-header"><div class="panel-title"><i class="fas fa-users"></i> All Users ({len(users)})</div></div>
      <div class="panel-body p-0"><div class="table-responsive">
        <table class="table table-modern mb-0">
          <thead><tr><th>ID</th><th>Username</th><th>Name</th><th>Role</th><th>Email</th><th>Dept</th></tr></thead>
          <tbody>{rows}</tbody>
        </table>
      </div></div>
    </div>""", unsafe_allow_html=True)
    st.markdown('<div class="panel mt-3"><div class="panel-header"><div class="panel-title"><i class="fas fa-user-plus"></i> Add New User</div></div><div class="panel-body">', unsafe_allow_html=True)
    with st.form("auf"):
        c1,c2=st.columns(2);un=c1.text_input("Username *");pw=c2.text_input("Password *",value="pass123")
        c3,c4=st.columns(2);en=c3.text_input("Employee Name");rl=c4.selectbox("Role",['user','technician','admin'])
        c5,c6=st.columns(2);em=c5.text_input("Email");dp=c6.text_input("Department")
        if st.form_submit_button("Create User",use_container_width=True):
            if un and pw:
                try:
                    conn2=get_db();conn2.execute("INSERT INTO users(username,password,role,employee_name,email,department) VALUES(?,?,?,?,?,?)",(un,pw,rl,en,em,dp));conn2.commit();conn2.close()
                    st.success(f"User '{un}' created!");st.rerun()
                except Exception as e: st.error(str(e))
            else: st.warning("Username and password required.")
    st.markdown('</div></div></div>', unsafe_allow_html=True)

# ─── Admin Email Config ────────────────────────────────────────────────────────

def pg_email_config():
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    st.markdown('<h4><i class="fas fa-envelope-cog me-2 text-muted"></i>Email Configuration</h4>', unsafe_allow_html=True)
    cfg=configparser.ConfigParser();cfg.read(CONFIG_PATH)
    smtp=dict(cfg['SMTP']) if 'SMTP' in cfg else {}
    notif=dict(cfg['NOTIFICATIONS']) if 'NOTIFICATIONS' in cfg else {}
    t1,t2,t3=st.tabs(["⚙️ SMTP Settings","🔔 Notifications","📨 Test Email"])
    with t1:
        if smtp.get('smtp_user'): st.success(f"✅ Configured: {smtp.get('smtp_user')}")
        else: st.warning("⚠️ Email not configured yet.")
        with st.form("smf"):
            c1,c2=st.columns([3,1])
            h=c1.text_input("SMTP Host",value=smtp.get('smtp_host','smtp.gmail.com'))
            p=int(c2.text_input("Port",value=smtp.get('smtp_port','587')))
            u=st.text_input("Sender Email",value=smtp.get('smtp_user',''))
            pw=st.text_input("App Password",value=smtp.get('smtp_password',''),type="password")
            fn=st.text_input("Display Name",value=smtp.get('from_name','Qualesce IT Tracker'))
            if st.form_submit_button("Save SMTP Settings",use_container_width=True):
                if 'SMTP' not in cfg: cfg['SMTP']={}
                cfg['SMTP'].update({'SMTP_HOST':h,'SMTP_PORT':str(p),'SMTP_USER':u,'SMTP_PASSWORD':pw,'FROM_NAME':fn})
                with open(CONFIG_PATH,'w') as f: cfg.write(f)
                st.success("Saved!"); st.rerun()
    with t2:
        def nb(k): return notif.get(k,'true')!='false'
        with st.form("nf"):
            n1=st.checkbox("New request → notify admin",value=nb('notify_new_request'))
            n2=st.checkbox("Accept/Decline → notify employee",value=nb('notify_employee_action'))
            n3=st.checkbox("Technician assigned → notify tech + employee",value=nb('notify_technician_assigned'))
            n4=st.checkbox("Status change → notify employee",value=nb('notify_status_employee'))
            n5=st.checkbox("Tech status change → notify admin",value=nb('notify_status_admin'))
            n6=st.checkbox("Chat messages → notify all parties",value=nb('notify_chat'))
            if st.form_submit_button("Save Preferences",use_container_width=True):
                if 'NOTIFICATIONS' not in cfg: cfg['NOTIFICATIONS']={}
                for k,v in zip(['notify_new_request','notify_employee_action','notify_technician_assigned','notify_status_employee','notify_status_admin','notify_chat'],[n1,n2,n3,n4,n5,n6]):
                    cfg['NOTIFICATIONS'][k]='true' if v else 'false'
                with open(CONFIG_PATH,'w') as f: cfg.write(f)
                st.success("Saved!")
    with t3:
        with st.form("tf"):
            to=st.text_input("Send test to",value=smtp.get('smtp_user',''))
            if st.form_submit_button("Send Test Email",use_container_width=True):
                ok,msg=send_email([to],"IT Tracker — Test","<p>✅ SMTP is working correctly!</p>")
                st.success("Test email sent!") if ok else st.error(f"Failed: {msg}")
    st.markdown('</div>', unsafe_allow_html=True)

# ─── User pages ────────────────────────────────────────────────────────────────

def pg_user_dashboard():
    name=st.session_state.employee_name
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    st.markdown(f'<div class="panel mb-4" style="background:linear-gradient(135deg,#0f3460,#533483);color:#fff;padding:24px 28px;border-radius:16px"><h4 style="margin:0;color:#fff"><i class="fas fa-hand-wave me-2"></i>Welcome back, {name}!</h4><p style="margin:6px 0 0;opacity:.8">Here\'s your IT asset overview</p></div>', unsafe_allow_html=True)
    conn=get_db()
    assets=conn.execute("SELECT * FROM assets WHERE login_id LIKE ? ORDER BY asset_no",(f'%{name}%',)).fetchall()
    reqs=conn.execute("SELECT * FROM service_requests WHERE employee_id=? ORDER BY created_at DESC LIMIT 5",(st.session_state.user_id,)).fetchall()
    conn.close()
    c1,c2=st.columns(2)
    c1.metric("My Assets",len(assets)); c2.metric("My Requests",len(reqs))
    if assets:
        rows="".join(f"""<tr><td class="fw-semibold">{a['asset_no']}</td>
        <td>{a['model'] or '—'}</td><td>{sbadge(a['asset_status']) if a['asset_status'] else '—'}</td>
        <td>{a['serial_no'] or '—'}</td></tr>""" for a in assets)
        st.markdown(f"""
        <div class="panel mt-3">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-laptop"></i> My Assets</div></div>
          <div class="panel-body p-0"><div class="table-responsive">
            <table class="table table-modern mb-0">
              <thead><tr><th>Asset No</th><th>Model</th><th>Status</th><th>Serial</th></tr></thead>
              <tbody>{rows}</tbody>
            </table>
          </div></div>
        </div>""", unsafe_allow_html=True)
    else: st.info("No assets assigned to you. Contact IT Admin.")
    st.markdown('</div>', unsafe_allow_html=True)

def pg_user_my_requests():
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    st.markdown('<h4><i class="fas fa-ticket-alt me-2 text-muted"></i>My Service Requests</h4>', unsafe_allow_html=True)
    if st.button("+ New Request"):
        st.session_state.page='request_service'; st.rerun()
    conn=get_db()
    reqs=conn.execute("SELECT * FROM service_requests WHERE employee_id=? ORDER BY created_at DESC",(st.session_state.user_id,)).fetchall()
    conn.close()
    if not reqs:
        st.markdown('<div class="panel"><div class="panel-body text-center text-muted py-5"><i class="fas fa-inbox fa-3x mb-3 opacity-50"></i><p>No requests yet.</p></div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True); return
    rows="".join(f"""<tr>
      <td class="fw-bold text-muted">#{r['id']}</td>
      <td class="fw-semibold">{r['service_type']}</td>
      <td>{r['asset_no'] or '—'}</td>
      <td>{sbadge(r['status'])}</td>
      <td class="small text-muted">{r['assigned_to_name'] or '—'}</td>
      <td class="small text-muted">{r['created_at'][:10]}</td>
    </tr>""" for r in reqs)
    st.markdown(f"""
    <div class="panel">
      <div class="panel-header"><div class="panel-title"><i class="fas fa-list"></i> My Requests ({len(reqs)})</div></div>
      <div class="panel-body p-0"><div class="table-responsive">
        <table class="table table-modern mb-0">
          <thead><tr><th>#</th><th>Issue</th><th>Asset</th><th>Status</th><th>Assigned</th><th>Date</th></tr></thead>
          <tbody>{rows}</tbody>
        </table>
      </div></div>
    </div>""", unsafe_allow_html=True)
    st.markdown("---"); st.caption("Click a request to view and chat:")
    cols=st.columns(min(len(reqs),4))
    for i,r in enumerate(reqs[:8]):
        with cols[i%4]:
            if st.button(f"#{r['id']} {sbadge(r['status'])}",key=f"ur_{r['id']}",use_container_width=True):
                st.session_state.req_id=r['id']; st.session_state.page='req_detail_user'; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def pg_user_req_detail():
    req_id=st.session_state.get('req_id')
    if not req_id: st.session_state.page='my_requests';st.rerun();return
    conn=get_db()
    req=conn.execute("SELECT * FROM service_requests WHERE id=? AND employee_id=?",(req_id,st.session_state.user_id)).fetchone()
    if not req: conn.close();st.error("Not found.");return
    msgs=conn.execute("SELECT * FROM chat_messages WHERE request_id=? ORDER BY created_at",(req_id,)).fetchall()
    conn.close()
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    if st.button("← My Requests"): st.session_state.page='my_requests';st.rerun()
    st.markdown(f'<h4><i class="fas fa-ticket-alt me-2 text-muted"></i>Request #{req_id}</h4>', unsafe_allow_html=True)
    left,right=st.columns([1,1.6])
    with left:
        st.markdown(f"""
        <div class="panel">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-info-circle"></i> Details</div></div>
          <div class="panel-body">
            <div class="text-center mb-3 fs-5">{sbadge(req['status'])}</div>
            <table class="table table-sm small">
              <tr><td class="text-muted fw-semibold">Asset</td><td>{req['asset_no'] or '—'}</td></tr>
              <tr><td class="text-muted fw-semibold">Issue</td><td>{req['service_type']}</td></tr>
              <tr><td class="text-muted fw-semibold">Submitted</td><td>{req['created_at'][:10]}</td></tr>
              <tr><td class="text-muted fw-semibold">Assigned To</td><td>{req['assigned_to_name'] or 'Not assigned yet'}</td></tr>
            </table>
            <div class="small bg-light rounded p-2 mt-2">{req['description']}</div>
          </div>
        </div>""", unsafe_allow_html=True)
    with right:
        msgs_html=""
        for m in msgs:
            is_me=m['sender_id']==st.session_state.user_id
            bg='#533483' if is_me else '#0f3460' if m['sender_role']=='admin' else '#0891b2' if m['sender_role']=='technician' else '#e2e8f0'
            tc='#fff' if m['sender_role']!='user' or is_me else '#1e293b'
            side='justify-content-end' if is_me else ''
            lbl='You' if is_me else f"{m['sender_name']} ({m['sender_role']})"
            msgs_html+=f"""
            <div class="d-flex mb-3 {side}">
              <div style="max-width:75%">
                <div class="small text-muted mb-1 {'text-end' if is_me else ''}">{lbl} · {m['created_at'][11:16]}</div>
                <div style="padding:10px 14px;word-break:break-word;background:{bg};color:{tc};
                  border-radius:{'12px 12px 0 12px' if is_me else '12px 12px 12px 0'}">{m['message']}</div>
              </div>
            </div>"""
        st.markdown(f"""
        <div class="panel" style="min-height:460px;display:flex;flex-direction:column">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-comments"></i> Chat</div></div>
          <div style="flex:1;overflow-y:auto;padding:20px;background:#f8fafc;min-height:300px">
            {msgs_html or '<div class="text-center text-muted py-4"><i class="fas fa-comments fa-2x mb-2 opacity-50"></i><p class="small">No messages yet.</p></div>'}
          </div>
        </div>""", unsafe_allow_html=True)
        if req['status'] not in ('Declined','Completed'):
            with st.form("ucf",clear_on_submit=True):
                mi=st.text_area("",placeholder="Type your message...",height=80,label_visibility="collapsed")
                if st.form_submit_button("Send Message 📨",use_container_width=True):
                    if mi.strip(): _chat(req_id,req,mi.strip(),'user');st.rerun()
        else: st.info(f"Chat closed — request is {req['status'].lower()}.")
    st.markdown('</div>', unsafe_allow_html=True)

def pg_request_service():
    name=st.session_state.employee_name
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    conn=get_db()
    assets=conn.execute("SELECT asset_no,model FROM assets WHERE login_id LIKE ? ORDER BY asset_no",(f'%{name}%',)).fetchall()
    conn.close()
    st.markdown(f"""
    <div class="panel" style="max-width:640px">
      <div class="panel-header"><div class="panel-title"><i class="fas fa-wrench"></i> Submit a Service Request</div></div>
      <div class="panel-body">""", unsafe_allow_html=True)
    if not assets:
        st.warning("No assets assigned to you. Contact IT Admin.")
        st.markdown('</div></div></div>', unsafe_allow_html=True); return
    with st.form("rsf"):
        ao={a['asset_no']:f"{a['asset_no']} — {a['model']}" for a in assets}
        an=st.selectbox("Select Asset *",list(ao.keys()),format_func=lambda x:ao[x])
        st_=st.selectbox("Issue Type *",['Hardware Problem','Software Issue','Network Issue','Battery Problem','Screen Issue','Keyboard/Touchpad Issue','Slow Performance','Virus/Malware','OS Issue','Other'])
        desc=st.text_area("Describe the Issue *",height=120,placeholder="Please describe the issue in detail...")
        rem=st.text_input("Additional Remarks",placeholder="Any extra information...")
        if st.form_submit_button("Submit Request 📤",use_container_width=True):
            if desc.strip():
                conn2=get_db()
                cur=conn2.execute("INSERT INTO service_requests(asset_no,employee_id,employee_name,service_type,description,remarks,status) VALUES(?,?,?,?,?,?,?)",
                    (an,st.session_state.user_id,name,st_,desc.strip(),rem,'Pending'))
                rid=cur.lastrowid;conn2.commit()
                admins=conn2.execute("SELECT email FROM users WHERE role='admin' AND email!=''").fetchall()
                conn2.close()
                if get_nc()['notify_new_request']:
                    send_email([a['email'] for a in admins if a['email']],f"New Request #{rid} from {name}",f"<p><b>{name}</b> submitted request #{rid}: {st_} on {an}</p><p>{desc}</p>")
                st.success("Request submitted! You'll be notified once reviewed.")
                st.session_state.page='my_requests';st.rerun()
            else: st.warning("Please describe the issue.")
    st.markdown('</div></div></div>', unsafe_allow_html=True)

def pg_email_prefs():
    uid=st.session_state.user_id;role=st.session_state.role
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    conn=get_db();row=conn.execute("SELECT notification_prefs,email FROM users WHERE id=?",(uid,)).fetchone();conn.close()
    uemail=row['email'] if row else ''
    try: prefs=json.loads(row['notification_prefs']) if row and row['notification_prefs'] else {}
    except: prefs={}
    st.markdown(f"""
    <div class="panel" style="max-width:640px">
      <div class="panel-header"><div class="panel-title"><i class="fas fa-bell"></i> Email Preferences</div></div>
      <div class="panel-body">
        {'<div class="alert alert-info small"><i class="fas fa-envelope me-2"></i>Notifications sent to <strong>'+uemail+'</strong></div>' if uemail else '<div class="alert alert-warning small"><i class="fas fa-exclamation-triangle me-2"></i>No email set. Ask your admin to add one.</div>'}
      </div>
    </div>""", unsafe_allow_html=True)
    with st.form("epf"):
        if role=='technician':
            na=st.checkbox("New task assigned to me",value=prefs.get('assigned',True))
            nc=st.checkbox("Chat messages on my tasks",value=prefs.get('chat',True))
            ns=st.checkbox("Admin status updates",value=prefs.get('status',True))
            if st.form_submit_button("Save Preferences",use_container_width=True):
                conn2=get_db();conn2.execute("UPDATE users SET notification_prefs=? WHERE id=?",(json.dumps({'assigned':na,'chat':nc,'status':ns}),uid));conn2.commit();conn2.close();st.success("Saved!")
        else:
            na=st.checkbox("Request accepted/declined",value=prefs.get('action',True))
            ns=st.checkbox("Status changes",value=prefs.get('status',True))
            nc=st.checkbox("Chat messages",value=prefs.get('chat',True))
            if st.form_submit_button("Save Preferences",use_container_width=True):
                conn2=get_db();conn2.execute("UPDATE users SET notification_prefs=? WHERE id=?",(json.dumps({'action':na,'status':ns,'chat':nc}),uid));conn2.commit();conn2.close();st.success("Saved!")
    st.markdown('</div>', unsafe_allow_html=True)

# ─── Technician pages ──────────────────────────────────────────────────────────

def pg_tech_dashboard():
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    conn=get_db()
    reqs=conn.execute("SELECT * FROM service_requests WHERE assigned_to_id=? ORDER BY updated_at DESC",(st.session_state.user_id,)).fetchall()
    conn.close()
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Total Assigned",len(reqs));c2.metric("In Progress",sum(1 for r in reqs if r['status']=='In Progress'))
    c3.metric("On Hold",sum(1 for r in reqs if r['status']=='Hold'));c4.metric("Completed",sum(1 for r in reqs if r['status']=='Completed'))
    if not reqs:
        st.markdown('<div class="panel mt-3"><div class="panel-body text-center text-muted py-5"><i class="fas fa-inbox fa-3x mb-3 opacity-50"></i><p>No requests assigned yet.</p></div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True); return
    rows="".join(f"""<tr>
      <td class="fw-bold text-muted">#{r['id']}</td>
      <td class="fw-semibold">{r['employee_name']}</td>
      <td>{r['asset_no'] or '—'}</td>
      <td class="small">{r['service_type']}</td>
      <td>{sbadge(r['status'])}</td>
      <td class="small text-muted">{r['updated_at'][:10]}</td>
    </tr>""" for r in reqs)
    st.markdown(f"""
    <div class="panel mt-3">
      <div class="panel-header"><div class="panel-title"><i class="fas fa-tasks"></i> My Assigned Requests</div></div>
      <div class="panel-body p-0"><div class="table-responsive">
        <table class="table table-modern mb-0">
          <thead><tr><th>#</th><th>Employee</th><th>Asset</th><th>Issue</th><th>Status</th><th>Updated</th></tr></thead>
          <tbody>{rows}</tbody>
        </table>
      </div></div>
    </div>""", unsafe_allow_html=True)
    st.markdown("---"); st.caption("Click to open a request:")
    cols=st.columns(min(len(reqs),4))
    for i,r in enumerate(reqs[:8]):
        with cols[i%4]:
            if st.button(f"#{r['id']} {r['employee_name'][:10]}",key=f"tr_{r['id']}",use_container_width=True):
                st.session_state.req_id=r['id']; st.session_state.page='tech_req_detail'; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def pg_tech_req_detail():
    req_id=st.session_state.get('req_id')
    if not req_id: st.session_state.page='dashboard';st.rerun();return
    conn=get_db()
    req=conn.execute("SELECT * FROM service_requests WHERE id=? AND assigned_to_id=?",(req_id,st.session_state.user_id)).fetchone()
    if not req: conn.close();st.error("Not found.");return
    msgs=conn.execute("SELECT * FROM chat_messages WHERE request_id=? ORDER BY created_at",(req_id,)).fetchall()
    conn.close()
    st.markdown('<div class="content-area" style="padding:1.5rem 2rem">', unsafe_allow_html=True)
    if st.button("← Dashboard"): st.session_state.page='dashboard';st.rerun()
    st.markdown(f'<h4><i class="fas fa-tasks me-2 text-muted"></i>Request #{req_id} · {req["employee_name"]}</h4>', unsafe_allow_html=True)
    left,right=st.columns([1,1.6])
    with left:
        st.markdown(f"""
        <div class="panel">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-info-circle"></i> Details</div></div>
          <div class="panel-body">
            <div class="text-center mb-3 fs-5">{sbadge(req['status'])}</div>
            <table class="table table-sm small">
              <tr><td class="text-muted fw-semibold">Employee</td><td>{req['employee_name']}</td></tr>
              <tr><td class="text-muted fw-semibold">Asset</td><td>{req['asset_no'] or '—'}</td></tr>
              <tr><td class="text-muted fw-semibold">Issue</td><td>{req['service_type']}</td></tr>
            </table>
            <div class="small bg-light rounded p-2">{req['description']}</div>
          </div>
        </div>""", unsafe_allow_html=True)
        if req['status']!='Completed':
            st.markdown('<div class="panel mt-3"><div class="panel-header"><div class="panel-title"><i class="fas fa-exchange-alt"></i> Update Status</div></div><div class="panel-body">', unsafe_allow_html=True)
            with st.form("tsf"):
                opts=['In Progress','Hold','Completed']
                ns=st.selectbox("",opts,index=opts.index(req['status']) if req['status'] in opts else 0,label_visibility="collapsed")
                if st.form_submit_button("Update Status",use_container_width=True):
                    conn2=get_db();conn2.execute("UPDATE service_requests SET status=?,updated_at=? WHERE id=?",(ns,datetime.now().isoformat(),req_id));conn2.commit()
                    emp=conn2.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
                    admins=conn2.execute("SELECT email FROM users WHERE role='admin' AND email!=''").fetchall()
                    conn2.close();nc=get_nc()
                    if emp and emp['email'] and nc['notify_status_employee']:
                        send_email([emp['email']],f"Request #{req_id}: {ns}",status_email_html(req['employee_name'],ns,req_id,req['service_type'],req['asset_no']))
                    if nc['notify_status_admin']:
                        send_email([a['email'] for a in admins if a['email']],f"Tech updated Request #{req_id} to {ns}",status_email_html(req['employee_name'],ns,req_id,req['service_type'],req['asset_no']))
                    st.success(f"Status → {ns}");st.rerun()
            st.markdown('</div></div>', unsafe_allow_html=True)
        else: st.success("✅ Task completed.")
    with right:
        msgs_html=""
        for m in msgs:
            is_me=m['sender_id']==st.session_state.user_id
            bg='#0891b2' if is_me else '#0f3460' if m['sender_role']=='admin' else '#e2e8f0'
            tc='#fff' if m['sender_role'] in ('admin','technician') else '#1e293b'
            side='justify-content-end' if is_me else ''
            lbl='You' if is_me else f"{m['sender_name']} ({m['sender_role']})"
            msgs_html+=f"""
            <div class="d-flex mb-3 {side}">
              <div style="max-width:75%">
                <div class="small text-muted mb-1 {'text-end' if is_me else ''}">{lbl} · {m['created_at'][11:16]}</div>
                <div style="padding:10px 14px;word-break:break-word;background:{bg};color:{tc};
                  border-radius:{'12px 12px 0 12px' if is_me else '12px 12px 12px 0'}">{m['message']}</div>
              </div>
            </div>"""
        st.markdown(f"""
        <div class="panel" style="min-height:460px;display:flex;flex-direction:column">
          <div class="panel-header"><div class="panel-title"><i class="fas fa-comments"></i> Chat</div></div>
          <div style="flex:1;overflow-y:auto;padding:20px;background:#f8fafc;min-height:300px">
            {msgs_html or '<div class="text-center text-muted py-4"><i class="fas fa-comments fa-2x mb-2 opacity-50"></i><p class="small">No messages yet.</p></div>'}
          </div>
        </div>""", unsafe_allow_html=True)
        if req['status']!='Completed':
            with st.form("tcf",clear_on_submit=True):
                mi=st.text_area("",placeholder="Type your message...",height=80,label_visibility="collapsed")
                if st.form_submit_button("Send Message 📨",use_container_width=True):
                    if mi.strip(): _chat(req_id,req,mi.strip(),'technician');st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    init_db()
    logo_b64 = inject_theme()

    if not st.session_state.get('logged_in'):
        page_login(logo_b64); return

    sidebar_nav(logo_b64)

    role  = st.session_state.role
    page  = st.session_state.get('page','dashboard')

    if role == 'admin':
        {'dashboard':pg_admin_dashboard,'service_requests':pg_admin_service_requests,
         'req_detail':pg_admin_req_detail,'assets':pg_admin_assets,'employees':pg_admin_employees,
         'services':pg_admin_services,'users':pg_admin_users,'email_config':pg_email_config
        }.get(page, pg_admin_dashboard)()
    elif role == 'technician':
        {'dashboard':pg_tech_dashboard,'tech_req_detail':pg_tech_req_detail,
         'email_prefs':pg_email_prefs}.get(page, pg_tech_dashboard)()
    else:
        {'dashboard':pg_user_dashboard,'my_requests':pg_user_my_requests,
         'req_detail_user':pg_user_req_detail,'request_service':pg_request_service,
         'email_prefs':pg_email_prefs}.get(page, pg_user_dashboard)()

main()
