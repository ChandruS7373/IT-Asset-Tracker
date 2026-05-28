import streamlit as st
import sqlite3, os, json, random, smtplib, configparser, io
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

st.set_page_config(page_title="Qualesce IT Tracker", page_icon="💻", layout="wide")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH     = os.path.join(BASE_DIR, 'data', 'ittracker.db')
CONFIG_PATH = os.path.join(BASE_DIR, 'email_config.ini')
EXCEL_PATH  = os.path.join(BASE_DIR, 'data', 'assets.xlsx')
os.makedirs(os.path.join(BASE_DIR, 'data'), exist_ok=True)

# ─── DB ────────────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db(); c = conn.cursor()
    c.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL, password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user', employee_name TEXT,
            email TEXT, department TEXT,
            notification_prefs TEXT DEFAULT '{}',
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS assets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            asset_no TEXT UNIQUE, asset_status TEXT, login_id TEXT,
            serial_no TEXT, transfer_history TEXT, model TEXT, years TEXT,
            resolution TEXT, sn TEXT, model2 TEXT,
            warranty_start TEXT, warranty_end TEXT, warranty_type TEXT, warranty_status TEXT,
            lan_mac TEXT, lan_ip TEXT, wireless_mac TEXT, wan_ip TEXT,
            admin_usb TEXT, bios_password TEXT, admin_password TEXT,
            spiceworks TEXT, windows_update TEXT, unwanted_apps TEXT,
            processor TEXT, ram TEXT, hdd TEXT, office365 TEXT,
            sharepoint TEXT, onedrive TEXT, worksoft TEXT, ia TEXT,
            sql_version TEXT, system_cleanup TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS service_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            asset_no TEXT, service_date TEXT, service_type TEXT,
            description TEXT, technician TEXT, status TEXT DEFAULT 'Completed',
            cost TEXT, remarks TEXT, logged_by TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS service_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            asset_no TEXT, employee_id INTEGER, employee_name TEXT,
            service_type TEXT, description TEXT, remarks TEXT,
            status TEXT DEFAULT 'Pending',
            assigned_to_id INTEGER, assigned_to_name TEXT,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS chat_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            request_id INTEGER, sender_id INTEGER,
            sender_name TEXT, sender_role TEXT, message TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS stock_dashboard (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            status TEXT NOT NULL, model TEXT NOT NULL,
            count INTEGER DEFAULT 0, UNIQUE(status, model)
        );
        CREATE TABLE IF NOT EXISTS password_resets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL, otp TEXT NOT NULL,
            expires_at TEXT NOT NULL, used INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now'))
        );
    ''')
    try:
        c.execute("ALTER TABLE users ADD COLUMN notification_prefs TEXT DEFAULT '{}'")
    except Exception:
        pass
    c.execute("SELECT id FROM users WHERE username='admin'")
    if not c.fetchone():
        c.execute("INSERT INTO users (username,password,role,employee_name) VALUES (?,?,?,?)",
                  ('admin','admin123','admin','Administrator'))
    conn.commit(); conn.close()

# ─── Email ─────────────────────────────────────────────────────────────────────

def get_smtp_cfg():
    cfg = configparser.ConfigParser(); cfg.read(CONFIG_PATH)
    return cfg['SMTP'] if 'SMTP' in cfg else {}

def get_notif_cfg():
    cfg = configparser.ConfigParser(); cfg.read(CONFIG_PATH)
    sec = cfg['NOTIFICATIONS'] if 'NOTIFICATIONS' in cfg else {}
    def b(k): return sec.get(k,'true').lower() != 'false'
    return {'new_request':b('notify_new_request'),'employee_action':b('notify_employee_action'),
            'technician_assigned':b('notify_technician_assigned'),'status_employee':b('notify_status_employee'),
            'status_admin':b('notify_status_admin'),'chat':b('notify_chat')}

def send_email(to_emails, subject, html):
    cfg = get_smtp_cfg()
    host=cfg.get('smtp_host',''); port=int(cfg.get('smtp_port',587))
    user=cfg.get('smtp_user',''); pwd=cfg.get('smtp_password','')
    from_name=cfg.get('from_name','Qualesce IT Tracker')
    if not user: return False,'SMTP not configured'
    if isinstance(to_emails,str): to_emails=[to_emails]
    to_emails=[e for e in to_emails if e]
    if not to_emails: return False,'No recipients'
    msg=MIMEMultipart('alternative')
    msg['Subject']=subject; msg['From']=f'{from_name} <{user}>'; msg['To']=', '.join(to_emails)
    msg.attach(MIMEText(html,'html'))
    try:
        with smtplib.SMTP(host,port,timeout=10) as s:
            s.ehlo(); s.starttls(); s.login(user,pwd); s.sendmail(user,to_emails,msg.as_string())
        return True,'Sent'
    except Exception as e: return False,str(e)

def status_email(emp_name, status, req_id, stype, asset):
    c={'In Progress':'#f59e0b','Hold':'#6b7280','Completed':'#22c55e','Accepted':'#22c55e','Declined':'#ef4444'}.get(status,'#374151')
    return f"""<div style="font-family:Segoe UI,sans-serif;max-width:540px;margin:auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">
      <div style="background:linear-gradient(135deg,#0f3460,#533483);padding:22px;text-align:center"><h2 style="color:#fff;margin:0">Request #{req_id} — {status}</h2></div>
      <div style="padding:24px"><p>Hi <strong>{emp_name}</strong>, your request status is now <strong style="color:{c}">{status}</strong>.</p>
      <p>Asset: {asset} &nbsp;|&nbsp; Issue: {stype}</p></div>
      <div style="background:#f8fafc;padding:10px;text-align:center;color:#94a3b8;font-size:0.8rem">Qualesce IT Asset Tracker</div></div>"""

def chat_email(sender, role, message, req_id, stype):
    rl={'admin':'Admin','technician':'Technician','user':'Employee'}.get(role,role)
    return f"""<div style="font-family:Segoe UI,sans-serif;max-width:540px;margin:auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">
      <div style="background:linear-gradient(135deg,#0f3460,#533483);padding:22px;text-align:center"><h2 style="color:#fff;margin:0">New Message — Request #{req_id}</h2></div>
      <div style="padding:24px"><p>Message from <strong>{sender}</strong> ({rl}) on Request #{req_id} ({stype}):</p>
      <div style="background:#f1f5f9;border-left:4px solid #533483;border-radius:4px;padding:12px;margin:14px 0">{message}</div></div>
      <div style="background:#f8fafc;padding:10px;text-align:center;color:#94a3b8;font-size:0.8rem">Qualesce IT Asset Tracker</div></div>"""

# ─── Auth ──────────────────────────────────────────────────────────────────────

def page_login():
    st.markdown("<div style='text-align:center;padding:30px 0 10px'><h1 style='color:#0f3460'>💻 Qualesce IT Asset Tracker</h1><p style='color:#64748b'>Sign in to continue</p></div>", unsafe_allow_html=True)
    _, col, _ = st.columns([1,1,1])
    with col:
        with st.form("login"):
            u = st.text_input("Username or Email")
            p = st.text_input("Password", type="password")
            if st.form_submit_button("Sign In", use_container_width=True):
                conn=get_db()
                user=conn.execute("SELECT * FROM users WHERE (LOWER(email)=? OR LOWER(username)=?) AND password=?",
                    (u.strip().lower(),u.strip().lower(),p.strip())).fetchone()
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

def nav_btn(label, page):
    if st.sidebar.button(label, use_container_width=True, key=f'nav_{page}'):
        st.session_state.page = page
        st.session_state.pop('req_id', None)
        st.rerun()

def sidebar_nav():
    role = st.session_state.role
    name = st.session_state.employee_name
    badges = {'admin':'🟡 ADMIN','technician':'🟢 TECHNICIAN','user':'🔵 USER'}
    st.sidebar.markdown(f"### {name}")
    st.sidebar.caption(badges.get(role,'USER'))
    st.sidebar.divider()
    if role == 'admin':
        nav_btn('📊 Dashboard','dashboard')
        nav_btn('📥 Service Requests','service_requests')
        nav_btn('💻 Assets','assets')
        nav_btn('👥 Employees','employees')
        nav_btn('🔧 Service History','services')
        nav_btn('👤 Users','users')
        nav_btn('📧 Email Config','email_config')
    elif role == 'technician':
        nav_btn('📊 Dashboard','dashboard')
        nav_btn('🔔 Email Preferences','email_prefs')
    else:
        nav_btn('🏠 My Dashboard','dashboard')
        nav_btn('🎫 My Requests','my_requests')
        nav_btn('🔧 Request Service','request_service')
        nav_btn('🔔 Email Preferences','email_prefs')
    st.sidebar.divider()
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        do_logout()

SB = {'Pending':'🟡','Accepted':'🟢','Declined':'🔴','In Progress':'🔵','Hold':'⚪','Completed':'✅'}

def sbadge(s): return f"{SB.get(s,'⚫')} {s}"

# ─── Admin: Dashboard ──────────────────────────────────────────────────────────

def pg_admin_dashboard():
    conn = get_db()
    sc = {r[0]:r[1] for r in conn.execute("SELECT asset_status,COUNT(*) FROM assets GROUP BY asset_status").fetchall()}
    total    = conn.execute("SELECT COUNT(*) FROM assets").fetchone()[0]
    assigned = sc.get('Assigned',0)
    stock    = sc.get('IT Stock',0)+sc.get('ITStock',0)
    dead     = sc.get('DEAD',0)
    to_check = sc.get('To Check',0)
    yash     = sc.get('Yashwanth',0)
    users    = conn.execute("SELECT COUNT(*) FROM users WHERE role='user'").fetchone()[0]
    svc      = conn.execute("SELECT COUNT(*) FROM service_history").fetchone()[0]
    pending  = conn.execute("SELECT COUNT(*) FROM service_requests WHERE status='Pending'").fetchone()[0]
    stock_rows = conn.execute("SELECT status,model,count FROM stock_dashboard ORDER BY status,model").fetchall()
    conn.close()

    st.title("📊 Admin Dashboard")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total Assets", total); c2.metric("Assigned", assigned)
    c3.metric("IT Stock", stock);     c4.metric("DEAD", dead)
    c5,c6,c7,c8 = st.columns(4)
    c5.metric("To Check", to_check)
    if yash: c6.metric("Yashwanth", yash)
    c7.metric("Employees", users); c8.metric("Service Records", svc)

    if pending:
        st.warning(f"⚠️ {pending} pending service request(s) need review.")
        if st.button("View Requests →"):
            st.session_state.page='service_requests'; st.rerun()

    st.divider()
    col_l, col_r = st.columns([3,1])
    col_l.subheader("📦 Stock Dashboard")
    with col_r:
        if st.button("🔄 Sync from Assets"):
            _sync_stock(); st.rerun()

    if stock_rows:
        statuses, models, data = [], [], {}
        for r in stock_rows:
            if r['status'] not in statuses: statuses.append(r['status'])
            if r['model'] not in models: models.append(r['model'])
            data.setdefault(r['status'],{})[r['model']] = r['count']
        rows = []
        for s in statuses:
            row = {'Status': s}
            row.update({m: data.get(s,{}).get(m,0) for m in models})
            row['Total'] = sum(data.get(s,{}).get(m,0) for m in models)
            rows.append(row)
        gt = {'Status':'Grand Total'}
        gt.update({m: sum(data.get(s,{}).get(m,0) for s in statuses) for m in models})
        gt['Total'] = sum(gt[m] for m in models)
        rows.append(gt)
        st.dataframe(pd.DataFrame(rows).set_index('Status'), use_container_width=True)
    else:
        st.info("No stock data. Click 'Sync from Assets' to populate.")

def _sync_stock():
    from collections import defaultdict
    conn=get_db(); agg=defaultdict(int)
    for r in conn.execute("SELECT asset_status,model,COUNT(*) FROM assets WHERE asset_status!='' AND model!='' GROUP BY asset_status,model").fetchall():
        s=r[0].strip(); s='IT Stock' if s=='ITStock' else s
        m=r[1].strip()
        if m: agg[(s,m)]+=r[2]
    conn.execute("DELETE FROM stock_dashboard")
    for (s,m),cnt in agg.items():
        conn.execute("INSERT INTO stock_dashboard (status,model,count) VALUES (?,?,?)",(s,m,cnt))
    conn.commit(); conn.close()

# ─── Admin: Service Requests ───────────────────────────────────────────────────

def pg_admin_service_requests():
    st.title("📥 Service Requests")
    sf = st.selectbox("Filter", ['All','Pending','Accepted','In Progress','Hold','Completed','Declined'])
    conn=get_db()
    q="SELECT * FROM service_requests WHERE 1=1"+((" AND status=?" if sf!='All' else ""))+" ORDER BY created_at DESC"
    reqs=conn.execute(q,[sf] if sf!='All' else []).fetchall()
    conn.close()
    if not reqs: st.info("No requests found."); return
    for r in reqs:
        with st.container(border=True):
            c1,c2,c3,c4,c5=st.columns([0.4,1.8,1.6,1.2,0.8])
            c1.write(f"**#{r['id']}**")
            c2.write(f"**{r['employee_name']}**  \n{r['service_type']}")
            c3.write(f"Asset: {r['asset_no'] or '—'}  \n{r['created_at'][:10]}")
            c4.write(sbadge(r['status']))
            c5.write(r['assigned_to_name'] or '—')
            if st.button("Manage →", key=f"adm_{r['id']}"):
                st.session_state.req_id=r['id']; st.session_state.page='req_detail'; st.rerun()

def pg_admin_req_detail():
    req_id=st.session_state.get('req_id')
    if not req_id: st.session_state.page='service_requests'; st.rerun(); return
    conn=get_db()
    req=conn.execute("SELECT * FROM service_requests WHERE id=?",(req_id,)).fetchone()
    if not req: conn.close(); st.error("Not found."); return
    msgs=conn.execute("SELECT * FROM chat_messages WHERE request_id=? ORDER BY created_at",(req_id,)).fetchall()
    techs=conn.execute("SELECT id,employee_name,username FROM users WHERE role='technician'").fetchall()
    conn.close()
    if st.button("← Back"): st.session_state.page='service_requests'; st.rerun()
    st.title(f"Request #{req_id} · {req['employee_name']}")
    left,right=st.columns([1,1.6])
    with left:
        st.write(f"**Status:** {sbadge(req['status'])}")
        st.write(f"**Asset:** {req['asset_no'] or '—'}")
        st.write(f"**Issue:** {req['service_type']}")
        st.write(f"**Submitted:** {req['created_at'][:10]}")
        st.write(f"**Assigned:** {req['assigned_to_name'] or '—'}")
        st.markdown(f"**Description:** {req['description']}")
        st.divider()
        if req['status']=='Pending':
            ca,cd=st.columns(2)
            if ca.button("✅ Accept",use_container_width=True,type="primary"):
                _set_status(req_id,req,'Accepted'); st.rerun()
            if cd.button("❌ Decline",use_container_width=True):
                _set_status(req_id,req,'Declined'); st.rerun()
        with st.form("sf"):
            ns=st.selectbox("Change Status",['Pending','Accepted','In Progress','Hold','Completed','Declined'],
                index=['Pending','Accepted','In Progress','Hold','Completed','Declined'].index(req['status']))
            if st.form_submit_button("Update"): _set_status(req_id,req,ns); st.rerun()
        if techs:
            with st.form("tf"):
                to={t['id']:t['employee_name'] or t['username'] for t in techs}
                tid=st.selectbox("Assign Technician",list(to.keys()),format_func=lambda x:to[x])
                if st.form_submit_button("Assign"): _assign(req_id,req,tid,to[tid]); st.rerun()
    with right:
        st.subheader("💬 Chat")
        for m in msgs:
            role=m['sender_role']
            with st.chat_message("user" if role=='admin' else "assistant"):
                st.caption(f"**{m['sender_name']}** ({role}) · {m['created_at'][11:16]}")
                st.write(m['message'])
        with st.form("cf",clear_on_submit=True):
            mi=st.text_area("",placeholder="Type message...",height=70,label_visibility="collapsed")
            if st.form_submit_button("Send 📨",use_container_width=True):
                if mi.strip(): _chat(req_id,req,mi.strip(),'admin'); st.rerun()

def _set_status(req_id,req,status):
    conn=get_db()
    conn.execute("UPDATE service_requests SET status=?,updated_at=? WHERE id=?",(status,datetime.now().isoformat(),req_id))
    conn.commit()
    emp=conn.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
    conn.close()
    if emp and emp['email'] and get_notif_cfg()['status_employee']:
        send_email([emp['email']],f"Request #{req_id}: {status}",status_email(req['employee_name'],status,req_id,req['service_type'],req['asset_no']))
    st.success(f"Status → {status}")

def _assign(req_id,req,tech_id,tech_name):
    conn=get_db()
    conn.execute("UPDATE service_requests SET assigned_to_id=?,assigned_to_name=?,status='In Progress',updated_at=? WHERE id=?",
        (tech_id,tech_name,datetime.now().isoformat(),req_id))
    conn.commit()
    tech=conn.execute("SELECT email FROM users WHERE id=?",(tech_id,)).fetchone()
    emp=conn.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
    conn.close()
    nc=get_notif_cfg()
    if tech and tech['email'] and nc['technician_assigned']:
        send_email([tech['email']],f"Request #{req_id} Assigned to You",f"<p>Request #{req_id} ({req['service_type']}) has been assigned to you by admin.</p>")
    if emp and emp['email'] and nc['status_employee']:
        send_email([emp['email']],f"Your Request #{req_id} is In Progress",status_email(req['employee_name'],'In Progress',req_id,req['service_type'],req['asset_no']))
    st.success(f"Assigned to {tech_name}!")

def _chat(req_id,req,message,sender_role):
    conn=get_db()
    conn.execute("INSERT INTO chat_messages (request_id,sender_id,sender_name,sender_role,message) VALUES (?,?,?,?,?)",
        (req_id,st.session_state.user_id,st.session_state.employee_name,sender_role,message))
    conn.commit()
    nc=get_notif_cfg()
    if nc['chat']:
        if sender_role=='admin':
            emp=conn.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
            if emp and emp['email']: send_email([emp['email']],f"Admin message on Request #{req_id}",chat_email(st.session_state.employee_name,'admin',message,req_id,req['service_type']))
        elif sender_role=='user':
            admins=conn.execute("SELECT email FROM users WHERE role='admin' AND email!=''").fetchall()
            send_email([a['email'] for a in admins],f"Employee message on Request #{req_id}",chat_email(st.session_state.employee_name,'user',message,req_id,req['service_type']))
        elif sender_role=='technician':
            emp=conn.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
            admins=conn.execute("SELECT email FROM users WHERE role='admin' AND email!=''").fetchall()
            recipients=([emp['email']] if emp and emp['email'] else [])+[a['email'] for a in admins if a['email']]
            if recipients: send_email(recipients,f"Technician message on Request #{req_id}",chat_email(st.session_state.employee_name,'technician',message,req_id,req['service_type']))
    conn.close()

# ─── Admin: Assets / Employees / Services / Users ──────────────────────────────

def pg_admin_assets():
    st.title("💻 All Assets")
    search=st.text_input("🔍 Search asset, employee, model...")
    sf=st.selectbox("Status",['All','Assigned','IT Stock','DEAD','To Check','Service','Yashwanth'])
    conn=get_db()
    q="SELECT * FROM assets WHERE 1=1"; p=[]
    if search:
        q+=" AND (login_id LIKE ? OR asset_no LIKE ? OR model LIKE ? OR serial_no LIKE ?)"; p+=[f'%{search}%']*4
    if sf!='All': q+=" AND asset_status=?"; p.append(sf)
    assets=conn.execute(q+" ORDER BY asset_no",p).fetchall(); conn.close()
    st.caption(f"{len(assets)} assets")
    if assets:
        st.dataframe(pd.DataFrame([{'Asset No':a['asset_no'],'Status':a['asset_status'] or '—','Employee':a['login_id'] or '—',
            'Model':a['model'] or '—','Serial':a['serial_no'] or '—','Warranty':a['warranty_status'] or '—'} for a in assets]),use_container_width=True,height=480)
    else: st.info("No assets found.")

def pg_admin_employees():
    st.title("👥 Employees")
    conn=get_db()
    emps=conn.execute("SELECT login_id,COUNT(*) as cnt,GROUP_CONCAT(model,', ') as models FROM assets WHERE login_id NOT IN ('IT Stock','DEAD','') AND login_id IS NOT NULL GROUP BY login_id ORDER BY login_id").fetchall()
    conn.close()
    search=st.text_input("🔍 Search employee...")
    rows=[{'Employee':e['login_id'],'Assets':e['cnt'],'Models':e['models']} for e in emps if not search or search.lower() in e['login_id'].lower()]
    if rows: st.dataframe(pd.DataFrame(rows),use_container_width=True)
    else: st.info("No employees found.")

def pg_admin_services():
    st.title("🔧 Service History")
    search=st.text_input("🔍 Search...")
    conn=get_db()
    q="SELECT sh.*,a.login_id FROM service_history sh LEFT JOIN assets a ON sh.asset_no=a.asset_no WHERE 1=1"; p=[]
    if search: q+=" AND (sh.asset_no LIKE ? OR sh.description LIKE ?)"; p+=[f'%{search}%']*2
    svcs=conn.execute(q+" ORDER BY sh.service_date DESC",p).fetchall(); conn.close()
    if svcs:
        st.dataframe(pd.DataFrame([{'Date':s['service_date'],'Asset':s['asset_no'],'Employee':s['login_id'] or '—',
            'Type':s['service_type'],'Status':s['status'],'Technician':s['technician'] or '—'} for s in svcs]),use_container_width=True)
    else: st.info("No records.")

def pg_admin_users():
    st.title("👤 Manage Users")
    conn=get_db(); users=conn.execute("SELECT * FROM users ORDER BY role,username").fetchall(); conn.close()
    st.dataframe(pd.DataFrame([{'ID':u['id'],'Username':u['username'],'Name':u['employee_name'] or '—',
        'Role':u['role'].upper(),'Email':u['email'] or '—','Dept':u['department'] or '—'} for u in users]),use_container_width=True)
    st.divider(); st.subheader("Add User")
    with st.form("auf"):
        c1,c2=st.columns(2)
        un=c1.text_input("Username *"); pw=c2.text_input("Password *",value="pass123")
        c3,c4=st.columns(2)
        en=c3.text_input("Employee Name"); rl=c4.selectbox("Role",['user','technician','admin'])
        c5,c6=st.columns(2)
        em=c5.text_input("Email"); dp=c6.text_input("Department")
        if st.form_submit_button("Create User"):
            if un and pw:
                try:
                    conn2=get_db()
                    conn2.execute("INSERT INTO users (username,password,role,employee_name,email,department) VALUES (?,?,?,?,?,?)",(un,pw,rl,en,em,dp))
                    conn2.commit(); conn2.close(); st.success(f"User '{un}' created!"); st.rerun()
                except Exception as e: st.error(str(e))
            else: st.warning("Username and password required.")

# ─── Admin: Email Config ───────────────────────────────────────────────────────

def pg_email_config():
    st.title("📧 Email Configuration")
    cfg=configparser.ConfigParser(); cfg.read(CONFIG_PATH)
    smtp=dict(cfg['SMTP']) if 'SMTP' in cfg else {}
    notif=dict(cfg['NOTIFICATIONS']) if 'NOTIFICATIONS' in cfg else {}
    t1,t2,t3=st.tabs(["SMTP","Notifications","Test Email"])
    with t1:
        with st.form("smf"):
            c1,c2=st.columns([3,1])
            h=c1.text_input("SMTP Host",value=smtp.get('smtp_host','smtp.gmail.com'))
            p=c2.number_input("Port",value=int(smtp.get('smtp_port',587)),step=1)
            u=st.text_input("Sender Email",value=smtp.get('smtp_user',''))
            pw=st.text_input("Password / App Password",value=smtp.get('smtp_password',''),type="password")
            fn=st.text_input("Display Name",value=smtp.get('from_name','Qualesce IT Tracker'))
            if st.form_submit_button("Save SMTP"):
                if 'SMTP' not in cfg: cfg['SMTP']={}
                cfg['SMTP'].update({'SMTP_HOST':h,'SMTP_PORT':str(p),'SMTP_USER':u,'SMTP_PASSWORD':pw,'FROM_NAME':fn})
                with open(CONFIG_PATH,'w') as f: cfg.write(f)
                st.success("SMTP saved!"); st.rerun()
        st.success(f"✅ Configured: {smtp.get('smtp_user')}") if smtp.get('smtp_user') else st.warning("⚠️ Not configured.")
    with t2:
        def nb(k): return notif.get(k,'true')!='false'
        with st.form("nf"):
            n1=st.checkbox("New request → admin email",value=nb('notify_new_request'))
            n2=st.checkbox("Accept/Decline → employee email",value=nb('notify_employee_action'))
            n3=st.checkbox("Technician assigned → email",value=nb('notify_technician_assigned'))
            n4=st.checkbox("Status change → employee email",value=nb('notify_status_employee'))
            n5=st.checkbox("Tech status change → admin email",value=nb('notify_status_admin'))
            n6=st.checkbox("Chat messages → email",value=nb('notify_chat'))
            if st.form_submit_button("Save Preferences"):
                if 'NOTIFICATIONS' not in cfg: cfg['NOTIFICATIONS']={}
                for k,v in zip(['notify_new_request','notify_employee_action','notify_technician_assigned',
                                 'notify_status_employee','notify_status_admin','notify_chat'],
                                [n1,n2,n3,n4,n5,n6]):
                    cfg['NOTIFICATIONS'][k]='true' if v else 'false'
                with open(CONFIG_PATH,'w') as f: cfg.write(f)
                st.success("Saved!")
    with t3:
        with st.form("tf"):
            to=st.text_input("Send test to",value=smtp.get('smtp_user',''))
            if st.form_submit_button("Send Test Email"):
                ok,msg=send_email([to],"IT Tracker — Test","<p>✅ SMTP is working!</p>")
                st.success("Sent!") if ok else st.error(f"Failed: {msg}")

# ─── User pages ────────────────────────────────────────────────────────────────

def pg_user_dashboard():
    name=st.session_state.employee_name; st.title(f"🏠 Welcome, {name}")
    conn=get_db()
    assets=conn.execute("SELECT * FROM assets WHERE login_id LIKE ? ORDER BY asset_no",(f'%{name}%',)).fetchall()
    reqs=conn.execute("SELECT COUNT(*) FROM service_requests WHERE employee_id=?",(st.session_state.user_id,)).fetchone()[0]
    conn.close()
    c1,c2=st.columns(2); c1.metric("My Assets",len(assets)); c2.metric("My Requests",reqs)
    st.divider(); st.subheader("💻 My Assets")
    if assets:
        st.dataframe(pd.DataFrame([{'Asset No':a['asset_no'],'Model':a['model'] or '—','Status':a['asset_status'] or '—','Serial':a['serial_no'] or '—'} for a in assets]),use_container_width=True)
    else: st.info("No assets assigned to you.")

def pg_user_my_requests():
    st.title("🎫 My Service Requests")
    if st.button("+ New Request"): st.session_state.page='request_service'; st.rerun()
    conn=get_db()
    reqs=conn.execute("SELECT * FROM service_requests WHERE employee_id=? ORDER BY created_at DESC",(st.session_state.user_id,)).fetchall()
    conn.close()
    if not reqs: st.info("No requests yet."); return
    for r in reqs:
        with st.container(border=True):
            c1,c2,c3,c4,c5=st.columns([0.4,1.8,1.6,1.2,0.8])
            c1.write(f"**#{r['id']}**"); c2.write(f"**{r['service_type']}**  \n{r['asset_no'] or '—'}")
            c3.write(f"{r['created_at'][:10]}  \n{r['assigned_to_name'] or '—'}"); c4.write(sbadge(r['status']))
            if c5.button("View →",key=f"ur_{r['id']}"): st.session_state.req_id=r['id']; st.session_state.page='req_detail_user'; st.rerun()

def pg_user_req_detail():
    req_id=st.session_state.get('req_id')
    if not req_id: st.session_state.page='my_requests'; st.rerun(); return
    conn=get_db()
    req=conn.execute("SELECT * FROM service_requests WHERE id=? AND employee_id=?",(req_id,st.session_state.user_id)).fetchone()
    if not req: conn.close(); st.error("Not found."); return
    msgs=conn.execute("SELECT * FROM chat_messages WHERE request_id=? ORDER BY created_at",(req_id,)).fetchall()
    conn.close()
    if st.button("← My Requests"): st.session_state.page='my_requests'; st.rerun()
    st.title(f"Request #{req_id}")
    left,right=st.columns([1,1.6])
    with left:
        st.write(f"**Status:** {sbadge(req['status'])}")
        st.write(f"**Asset:** {req['asset_no'] or '—'}")
        st.write(f"**Issue:** {req['service_type']}")
        st.write(f"**Submitted:** {req['created_at'][:10]}")
        st.write(f"**Assigned To:** {req['assigned_to_name'] or 'Not yet assigned'}")
        st.markdown(f"**Description:** {req['description']}")
    with right:
        st.subheader("💬 Chat")
        for m in msgs:
            is_me=m['sender_id']==st.session_state.user_id
            with st.chat_message("user" if is_me else "assistant"):
                lbl="You" if is_me else f"{m['sender_name']} ({m['sender_role']})"
                st.caption(f"**{lbl}** · {m['created_at'][11:16]}")
                st.write(m['message'])
        if req['status'] not in ('Declined','Completed'):
            with st.form("ucf",clear_on_submit=True):
                mi=st.text_area("",placeholder="Type message...",height=70,label_visibility="collapsed")
                if st.form_submit_button("Send 📨",use_container_width=True):
                    if mi.strip(): _chat(req_id,req,mi.strip(),'user'); st.rerun()
        else: st.info(f"Chat closed — {req['status'].lower()}.")

def pg_request_service():
    st.title("🔧 Submit Service Request")
    name=st.session_state.employee_name
    conn=get_db()
    assets=conn.execute("SELECT asset_no,model FROM assets WHERE login_id LIKE ? ORDER BY asset_no",(f'%{name}%',)).fetchall()
    conn.close()
    if not assets: st.warning("No assets assigned to you. Contact IT Admin."); return
    with st.form("rsf"):
        ao={a['asset_no']:f"{a['asset_no']} — {a['model']}" for a in assets}
        an=st.selectbox("Select Asset *",list(ao.keys()),format_func=lambda x:ao[x])
        st_=st.selectbox("Issue Type *",['Hardware Problem','Software Issue','Network Issue','Battery Problem',
            'Screen Issue','Keyboard/Touchpad Issue','Slow Performance','Virus/Malware','OS Issue','Other'])
        desc=st.text_area("Describe the Issue *",height=120,placeholder="Describe in detail...")
        rem=st.text_input("Remarks")
        if st.form_submit_button("Submit Request 📤"):
            if desc.strip():
                conn2=get_db()
                cur=conn2.execute("INSERT INTO service_requests (asset_no,employee_id,employee_name,service_type,description,remarks,status) VALUES (?,?,?,?,?,?,?)",
                    (an,st.session_state.user_id,name,st_,desc.strip(),rem,'Pending'))
                rid=cur.lastrowid; conn2.commit()
                admins=conn2.execute("SELECT email FROM users WHERE role='admin' AND email!=''").fetchall()
                conn2.close()
                if get_notif_cfg()['new_request']:
                    send_email([a['email'] for a in admins if a['email']],f"New Request #{rid} from {name}",
                        f"<p><b>{name}</b> submitted request #{rid}: {st_} on {an}</p><p>{desc}</p>")
                st.success("Request submitted!"); st.session_state.page='my_requests'; st.rerun()
            else: st.warning("Please describe the issue.")

def pg_email_prefs():
    st.title("🔔 Email Preferences")
    uid=st.session_state.user_id; role=st.session_state.role
    conn=get_db()
    row=conn.execute("SELECT notification_prefs,email FROM users WHERE id=?",(uid,)).fetchone()
    conn.close()
    uemail=row['email'] if row else ''
    try: prefs=json.loads(row['notification_prefs']) if row and row['notification_prefs'] else {}
    except: prefs={}
    st.info(f"Notifications → **{uemail}**") if uemail else st.warning("⚠️ No email set. Ask your admin to add one.")
    with st.form("epf"):
        if role=='technician':
            na=st.checkbox("New task assigned to me",value=prefs.get('assigned',True))
            nc=st.checkbox("Chat messages",value=prefs.get('chat',True))
            ns=st.checkbox("Admin status updates",value=prefs.get('status',True))
            if st.form_submit_button("Save"): np={'assigned':na,'chat':nc,'status':ns}; _save_prefs(uid,np)
        else:
            na=st.checkbox("Request accepted/declined",value=prefs.get('action',True))
            ns=st.checkbox("Status changes",value=prefs.get('status',True))
            nc=st.checkbox("Chat messages",value=prefs.get('chat',True))
            if st.form_submit_button("Save"): np={'action':na,'status':ns,'chat':nc}; _save_prefs(uid,np)

def _save_prefs(uid,prefs):
    conn=get_db(); conn.execute("UPDATE users SET notification_prefs=? WHERE id=?",(json.dumps(prefs),uid)); conn.commit(); conn.close(); st.success("Saved!")

# ─── Technician pages ──────────────────────────────────────────────────────────

def pg_tech_dashboard():
    st.title("📊 Technician Dashboard")
    conn=get_db()
    reqs=conn.execute("SELECT * FROM service_requests WHERE assigned_to_id=? ORDER BY updated_at DESC",(st.session_state.user_id,)).fetchall()
    conn.close()
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Total",len(reqs)); c2.metric("In Progress",sum(1 for r in reqs if r['status']=='In Progress'))
    c3.metric("Hold",sum(1 for r in reqs if r['status']=='Hold')); c4.metric("Completed",sum(1 for r in reqs if r['status']=='Completed'))
    st.divider()
    if not reqs: st.info("No requests assigned to you yet."); return
    for r in reqs:
        with st.container(border=True):
            c1,c2,c3,c4,c5=st.columns([0.4,1.8,1.6,1.2,0.8])
            c1.write(f"**#{r['id']}**"); c2.write(f"**{r['employee_name']}**  \n{r['service_type']}")
            c3.write(f"Asset: {r['asset_no'] or '—'}  \n{r['updated_at'][:10]}"); c4.write(sbadge(r['status']))
            if c5.button("Open →",key=f"tr_{r['id']}"): st.session_state.req_id=r['id']; st.session_state.page='tech_req_detail'; st.rerun()

def pg_tech_req_detail():
    req_id=st.session_state.get('req_id')
    if not req_id: st.session_state.page='dashboard'; st.rerun(); return
    conn=get_db()
    req=conn.execute("SELECT * FROM service_requests WHERE id=? AND assigned_to_id=?",(req_id,st.session_state.user_id)).fetchone()
    if not req: conn.close(); st.error("Not found."); return
    msgs=conn.execute("SELECT * FROM chat_messages WHERE request_id=? ORDER BY created_at",(req_id,)).fetchall()
    conn.close()
    if st.button("← Dashboard"): st.session_state.page='dashboard'; st.rerun()
    st.title(f"Request #{req_id} · {req['employee_name']}")
    left,right=st.columns([1,1.6])
    with left:
        st.write(f"**Status:** {sbadge(req['status'])}")
        st.write(f"**Asset:** {req['asset_no'] or '—'}")
        st.write(f"**Issue:** {req['service_type']}")
        st.markdown(f"**Description:** {req['description']}")
        if req['status']!='Completed':
            st.divider()
            with st.form("tsf"):
                ns=st.selectbox("Update Status",['In Progress','Hold','Completed'],
                    index=['In Progress','Hold','Completed'].index(req['status']) if req['status'] in ['In Progress','Hold','Completed'] else 0)
                if st.form_submit_button("Update"):
                    conn2=get_db()
                    conn2.execute("UPDATE service_requests SET status=?,updated_at=? WHERE id=?",(ns,datetime.now().isoformat(),req_id))
                    conn2.commit()
                    emp=conn2.execute("SELECT email FROM users WHERE id=?",(req['employee_id'],)).fetchone()
                    admins=conn2.execute("SELECT email FROM users WHERE role='admin' AND email!=''").fetchall()
                    conn2.close()
                    nc=get_notif_cfg()
                    if emp and emp['email'] and nc['status_employee']:
                        send_email([emp['email']],f"Request #{req_id}: {ns}",status_email(req['employee_name'],ns,req_id,req['service_type'],req['asset_no']))
                    if nc['status_admin']:
                        send_email([a['email'] for a in admins if a['email']],f"Tech updated Request #{req_id} to {ns}",status_email(req['employee_name'],ns,req_id,req['service_type'],req['asset_no']))
                    st.success(f"Status → {ns}"); st.rerun()
        else: st.success("✅ Task completed.")
    with right:
        st.subheader("💬 Chat")
        for m in msgs:
            is_me=m['sender_id']==st.session_state.user_id
            with st.chat_message("user" if is_me else "assistant"):
                st.caption(f"**{'You' if is_me else m['sender_name']+' ('+m['sender_role']+')'}** · {m['created_at'][11:16]}")
                st.write(m['message'])
        if req['status']!='Completed':
            with st.form("tcf",clear_on_submit=True):
                mi=st.text_area("",placeholder="Type message...",height=70,label_visibility="collapsed")
                if st.form_submit_button("Send 📨",use_container_width=True):
                    if mi.strip(): _chat(req_id,req,mi.strip(),'technician'); st.rerun()

# ─── Router ────────────────────────────────────────────────────────────────────

def main():
    init_db()
    if not st.session_state.get('logged_in'):
        page_login(); return
    sidebar_nav()
    role=st.session_state.role
    page=st.session_state.get('page','dashboard')
    if role=='admin':
        {'dashboard':pg_admin_dashboard,'service_requests':pg_admin_service_requests,
         'req_detail':pg_admin_req_detail,'assets':pg_admin_assets,'employees':pg_admin_employees,
         'services':pg_admin_services,'users':pg_admin_users,'email_config':pg_email_config
        }.get(page,pg_admin_dashboard)()
    elif role=='technician':
        {'dashboard':pg_tech_dashboard,'tech_req_detail':pg_tech_req_detail,
         'email_prefs':pg_email_prefs}.get(page,pg_tech_dashboard)()
    else:
        {'dashboard':pg_user_dashboard,'my_requests':pg_user_my_requests,
         'req_detail_user':pg_user_req_detail,'request_service':pg_request_service,
         'email_prefs':pg_email_prefs}.get(page,pg_user_dashboard)()

main()
