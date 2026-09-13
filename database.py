from __future__ import annotations
import sqlite3, os, json, hashlib
from datetime import datetime

BASE=os.path.dirname(os.path.dirname(__file__))
DB_PATH=os.path.join(BASE,'data','lease_enterprise_v5.db')
os.makedirs(os.path.dirname(DB_PATH),exist_ok=True)

SCHEMA='''
CREATE TABLE IF NOT EXISTS companies(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT UNIQUE NOT NULL,company_name TEXT NOT NULL,country TEXT,functional_currency TEXT DEFAULT 'INR',reporting_currency TEXT DEFAULT 'INR',fy_end TEXT DEFAULT '31-Mar',default_reporting_date TEXT,status TEXT DEFAULT 'ACTIVE',created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,username TEXT NOT NULL,password_hash TEXT NOT NULL,role TEXT NOT NULL,active INTEGER DEFAULT 1,created_at TEXT NOT NULL,UNIQUE(company_id,username));
CREATE TABLE IF NOT EXISTS leases(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,lease_id TEXT NOT NULL,status TEXT DEFAULT 'DRAFT',contract_date TEXT,commencement_date TEXT,expiry_date TEXT,first_payment_date TEXT,payment_frequency TEXT DEFAULT 'Monthly',payment_timing TEXT DEFAULT 'Arrears',base_payment REAL DEFAULT 0,escalation_type TEXT DEFAULT 'None',escalation_amount REAL DEFAULT 0,escalation_rate REAL DEFAULT 0,first_escalation_date TEXT,escalation_frequency TEXT DEFAULT 'Annual',currency TEXT DEFAULT 'INR',ib_rate REAL DEFAULT 0,ib_rate_type TEXT DEFAULT 'Annual Effective',prepaid REAL DEFAULT 0,idc REAL DEFAULT 0,restoration_pv REAL DEFAULT 0,lease_incentive REAL DEFAULT 0,deposit_paid REAL DEFAULT 0,deposit_refund_date TEXT,deposit_discount_rate REAL DEFAULT 0,deposit_ecl_rate REAL DEFAULT 0,deposit_classification TEXT DEFAULT 'Pending Assessment',ecl_pd REAL DEFAULT 0,ecl_lgd REAL DEFAULT 0,reporting_fx_rate REAL DEFAULT 1,counterparty TEXT,cost_center TEXT,gl_rou TEXT DEFAULT 'ROU Asset - Other',gl_liability TEXT DEFAULT 'Lease Liability',gl_interest TEXT DEFAULT 'Finance Cost',gl_depreciation TEXT DEFAULT 'Depreciation Expense',gl_accum_dep TEXT DEFAULT 'Accumulated Depreciation - ROU',gl_bank TEXT DEFAULT 'Bank',created_by TEXT,created_at TEXT NOT NULL,updated_at TEXT NOT NULL,is_deleted INTEGER DEFAULT 0,UNIQUE(company_id,lease_id));
CREATE TABLE IF NOT EXISTS events(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,event_id TEXT UNIQUE NOT NULL,lease_id TEXT NOT NULL,event_type TEXT NOT NULL,event_date TEXT NOT NULL,effective_date TEXT,old_value TEXT,new_value TEXT,reason TEXT,supporting_document TEXT,prepared_by TEXT,reviewed_by TEXT,approved_by TEXT,status TEXT DEFAULT 'DRAFT',created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS calculation_runs(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,run_id TEXT UNIQUE NOT NULL,reporting_date TEXT NOT NULL,engine_version TEXT,input_hash TEXT,status TEXT,created_by TEXT,created_at TEXT NOT NULL,lease_count INTEGER,error_count INTEGER);
CREATE TABLE IF NOT EXISTS errors(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,run_id TEXT,lease_id TEXT,module TEXT,severity TEXT,description TEXT,expected TEXT,actual TEXT,formula TEXT,status TEXT DEFAULT 'OPEN',resolver TEXT,resolution_date TEXT,created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS audit_trail(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,timestamp TEXT NOT NULL,username TEXT,role TEXT,lease_id TEXT,object_type TEXT,object_id TEXT,field TEXT,old_value TEXT,new_value TEXT,reason TEXT,action TEXT,approval_status TEXT,run_id TEXT,ip_or_session TEXT);
CREATE TABLE IF NOT EXISTS period_locks(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,reporting_date TEXT NOT NULL,locked_by TEXT,locked_at TEXT NOT NULL,UNIQUE(company_id,reporting_date));
CREATE TABLE IF NOT EXISTS import_batches(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,batch_id TEXT UNIQUE NOT NULL,filename TEXT,rows_total INTEGER,rows_valid INTEGER,rows_error INTEGER,status TEXT,created_by TEXT,created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS settings(company_id TEXT PRIMARY KEY,rounding_tolerance REAL DEFAULT 0.01,timezone TEXT DEFAULT 'Asia/Kolkata');
CREATE TABLE IF NOT EXISTS modifications(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id TEXT NOT NULL,modification_id TEXT UNIQUE NOT NULL,lease_id TEXT NOT NULL,event_id TEXT,effective_date TEXT NOT NULL,mod_type TEXT,new_base_payment REAL,new_ib_rate REAL,new_expiry_date TEXT,new_escalation_type TEXT,new_escalation_amount REAL,new_escalation_rate REAL,new_first_escalation_date TEXT,new_escalation_frequency TEXT,new_payment_frequency TEXT,new_payment_timing TEXT,new_first_payment_date TEXT,old_liability REAL,revised_liability REAL,rou_adjustment REAL,pl_gain_loss REAL,status TEXT DEFAULT 'DRAFT',created_by TEXT,created_at TEXT NOT NULL);
'''

# Columns added after the original V5/V5.1/V5.2 schema. Kept as an explicit
# migration list (rather than folded into CREATE TABLE) so that a database
# created by an earlier version of this software is upgraded in place
# instead of silently missing the new fields.
IMPORT_BATCH_COLUMN_MIGRATIONS = [("input_hash", "TEXT"), ("lease_ids_json", "TEXT")]

LEASE_COLUMN_MIGRATIONS = [
    ("extension_option_years", "REAL DEFAULT 0"),
    ("extension_reasonably_certain", "TEXT DEFAULT 'No'"),
    ("termination_option_date", "TEXT"),
    ("termination_reasonably_certain", "TEXT DEFAULT 'No'"),
    ("purchase_option_reasonably_certain", "TEXT DEFAULT 'No'"),
    ("restoration_cost_estimate", "REAL DEFAULT 0"),
    ("restoration_expected_date", "TEXT"),
    ("restoration_discount_rate", "REAL DEFAULT 0"),
    ("gl_liability_extract", "REAL"),
    ("escalation_schedule_json", "TEXT"),
    ("validation_tolerance", "REAL DEFAULT 0.01"),
    ("ecl_method", "TEXT DEFAULT 'Configured loss rate'"),
    ("ecl_horizon_years", "REAL DEFAULT 1"),
    ("ecl_discount_rate", "REAL DEFAULT 0"),
    ("initial_fx_rate", "REAL DEFAULT 0"),
    ("prior_reporting_fx_rate", "REAL DEFAULT 0"),
]

def now(): return datetime.now().isoformat(timespec='seconds')
def connect():
    c=sqlite3.connect(DB_PATH); c.row_factory=sqlite3.Row; return c

def _migrate_import_batch_columns(c):
    existing={row[1] for row in c.execute('PRAGMA table_info(import_batches)').fetchall()}
    for col,ddl in IMPORT_BATCH_COLUMN_MIGRATIONS:
        if col not in existing:
            c.execute(f'ALTER TABLE import_batches ADD COLUMN {col} {ddl}')

def _migrate_lease_columns(c):
    existing={row[1] for row in c.execute('PRAGMA table_info(leases)').fetchall()}
    for col,ddl in LEASE_COLUMN_MIGRATIONS:
        if col not in existing:
            c.execute(f'ALTER TABLE leases ADD COLUMN {col} {ddl}')

def init_db():
    with connect() as c:
        c.executescript(SCHEMA)
        _migrate_lease_columns(c)
        _migrate_import_batch_columns(c)
        c.commit()

def audit(company_id,username,role,action,object_type,object_id='',lease_id='',field='',old='',new='',reason='',approval='',run_id='',session=''):
    with connect() as c:
        c.execute('INSERT INTO audit_trail(company_id,timestamp,username,role,lease_id,object_type,object_id,field,old_value,new_value,reason,action,approval_status,run_id,ip_or_session) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(company_id,now(),username,role,lease_id,object_type,object_id,field,str(old),str(new),reason,action,approval,run_id,session)); c.commit()

def list_companies():
    with connect() as c:return c.execute('SELECT * FROM companies WHERE status="ACTIVE" ORDER BY company_name').fetchall()
def get_company(cid):
    with connect() as c:return c.execute('SELECT * FROM companies WHERE company_id=? AND status="ACTIVE"',(cid,)).fetchone()
def create_company(cid,name,country,currency,reporting,fy_end,reporting_date,username,password_hash):
    t=now()
    with connect() as c:
        c.execute('INSERT INTO companies(company_id,company_name,country,functional_currency,reporting_currency,fy_end,default_reporting_date,created_at) VALUES(?,?,?,?,?,?,?,?)',(cid,name,country,currency,reporting,fy_end,reporting_date,t))
        c.execute('INSERT INTO users(company_id,username,password_hash,role,created_at) VALUES(?,?,?,?,?)',(cid,username,password_hash,'ADMIN',t)); c.execute('INSERT INTO settings(company_id) VALUES(?)',(cid,)); c.commit()
def get_user(cid,u):
    with connect() as c:return c.execute('SELECT * FROM users WHERE company_id=? AND username=? AND active=1',(cid,u)).fetchone()
def list_users(cid):
    with connect() as c:return c.execute('SELECT id,username,role,active,created_at FROM users WHERE company_id=? ORDER BY username',(cid,)).fetchall()
def add_user(cid,u,pw,role):
    with connect() as c:c.execute('INSERT INTO users(company_id,username,password_hash,role,created_at) VALUES(?,?,?,?,?)',(cid,u,pw,role,now())); c.commit()

def lease_rows(cid,include_deleted=False):
    q='SELECT * FROM leases WHERE company_id=?'+('' if include_deleted else ' AND is_deleted=0')+' ORDER BY lease_id'
    with connect() as c:return c.execute(q,(cid,)).fetchall()
def get_lease(cid,lid):
    with connect() as c:return c.execute('SELECT * FROM leases WHERE company_id=? AND lease_id=? AND is_deleted=0',(cid,lid)).fetchone()

def upsert_lease(cid,data,username,role):
    t=now(); lid=str(data['lease_id']).strip(); fields=['status','contract_date','commencement_date','expiry_date','first_payment_date','payment_frequency','payment_timing','base_payment','escalation_type','escalation_amount','escalation_rate','first_escalation_date','escalation_frequency','currency','ib_rate','ib_rate_type','prepaid','idc','restoration_pv','lease_incentive','deposit_paid','deposit_refund_date','deposit_discount_rate','deposit_ecl_rate','deposit_classification','ecl_pd','ecl_lgd','reporting_fx_rate','counterparty','cost_center','gl_rou','gl_liability','gl_interest','gl_depreciation','gl_accum_dep','gl_bank','extension_option_years','extension_reasonably_certain','termination_option_date','termination_reasonably_certain','purchase_option_reasonably_certain','restoration_cost_estimate','restoration_expected_date','restoration_discount_rate','gl_liability_extract','escalation_schedule_json','validation_tolerance','ecl_method','ecl_horizon_years','ecl_discount_rate','initial_fx_rate','prior_reporting_fx_rate']
    with connect() as c:
        old=c.execute('SELECT * FROM leases WHERE company_id=? AND lease_id=?',(cid,lid)).fetchone(); vals=[data.get(f) for f in fields]
        if old:
            c.execute('UPDATE leases SET '+','.join(f'{f}=?' for f in fields)+',is_deleted=0,updated_at=? WHERE company_id=? AND lease_id=?',vals+[t,cid,lid]); action='UPDATE'
        else:
            c.execute('INSERT INTO leases(company_id,lease_id,'+','.join(fields)+',created_by,created_at,updated_at) VALUES('+','.join(['?']*(2+len(fields)))+',?,?,?)',[cid,lid]+vals+[username,t,t]); action='CREATE'
        c.commit()
    if old:
        for f in fields:
            if str(old[f])!=str(data.get(f)): audit(cid,username,role,'FIELD_CHANGE','LEASE',lid,lid,f,old[f],data.get(f),'Lease master change')
    else:audit(cid,username,role,'CREATE','LEASE',lid,lid,reason='Lease created')

def soft_delete_lease(cid,lid,username,role,reason):
    with connect() as c:c.execute('UPDATE leases SET is_deleted=1,status="DELETED",updated_at=? WHERE company_id=? AND lease_id=?',(now(),cid,lid)); c.commit()
    audit(cid,username,role,'SOFT_DELETE','LEASE',lid,lid,reason=reason)

def add_event(cid,data,username,role):
    t=now()
    with connect() as c:c.execute('INSERT INTO events(company_id,event_id,lease_id,event_type,event_date,effective_date,old_value,new_value,reason,supporting_document,prepared_by,status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(cid,data['event_id'],data['lease_id'],data['event_type'],data['event_date'],data.get('effective_date'),data.get('old_value'),data.get('new_value'),data.get('reason'),data.get('supporting_document'),username,'DRAFT',t)); c.commit()
    audit(cid,username,role,'CREATE','EVENT',data['event_id'],data['lease_id'],reason=data.get('reason',''))
def events(cid):
    with connect() as c:return c.execute('SELECT * FROM events WHERE company_id=? ORDER BY event_date DESC',(cid,)).fetchall()
def update_event_status(cid,eid,status,username,role):
    with connect() as c:c.execute('UPDATE events SET status=?,reviewed_by=CASE WHEN ?="REVIEWED" THEN ? ELSE reviewed_by END,approved_by=CASE WHEN ?="APPROVED" THEN ? ELSE approved_by END WHERE company_id=? AND event_id=?',(status,status,username,status,username,cid,eid)); c.commit()
    audit(cid,username,role,'STATUS_CHANGE','EVENT',eid,approval=status)

def add_modification(cid,data,username,role):
    """Create or idempotently refresh a modification import/post."""
    t=now(); mid=str(data['modification_id']).strip(); lid=str(data['lease_id']).strip()
    with connect() as c:
        existing=c.execute('SELECT company_id,lease_id FROM modifications WHERE modification_id=?',(mid,)).fetchone()
        if existing:
            if existing['company_id'] != cid or existing['lease_id'] != lid:
                raise ValueError(f'Modification_ID {mid} already exists for another company/lease; use a unique Modification_ID.')
            c.execute("""UPDATE modifications SET event_id=?,effective_date=?,mod_type=?,new_base_payment=?,new_ib_rate=?,new_expiry_date=?,
                new_escalation_type=?,new_escalation_amount=?,new_escalation_rate=?,new_first_escalation_date=?,new_escalation_frequency=?,
                new_payment_frequency=?,new_payment_timing=?,new_first_payment_date=?,old_liability=?,revised_liability=?,rou_adjustment=?,
                pl_gain_loss=?,status=?,created_by=?,created_at=? WHERE modification_id=?""",
                (data.get('event_id'),data['effective_date'],data.get('mod_type'),data.get('new_base_payment'),data.get('new_ib_rate'),data.get('new_expiry_date'),
                 data.get('new_escalation_type'),data.get('new_escalation_amount'),data.get('new_escalation_rate'),data.get('new_first_escalation_date'),
                 data.get('new_escalation_frequency'),data.get('new_payment_frequency'),data.get('new_payment_timing'),data.get('new_first_payment_date'),
                 data.get('old_liability'),data.get('revised_liability'),data.get('rou_adjustment'),data.get('pl_gain_loss'),data.get('status','APPROVED'),username,t,mid))
            action='UPDATE'
        else:
            c.execute("""INSERT INTO modifications(company_id,modification_id,lease_id,event_id,effective_date,mod_type,
                new_base_payment,new_ib_rate,new_expiry_date,new_escalation_type,new_escalation_amount,new_escalation_rate,
                new_first_escalation_date,new_escalation_frequency,new_payment_frequency,new_payment_timing,new_first_payment_date,
                old_liability,revised_liability,rou_adjustment,pl_gain_loss,status,created_by,created_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (cid,mid,lid,data.get('event_id'),data['effective_date'],data.get('mod_type'),
                 data.get('new_base_payment'),data.get('new_ib_rate'),data.get('new_expiry_date'),data.get('new_escalation_type'),
                 data.get('new_escalation_amount'),data.get('new_escalation_rate'),data.get('new_first_escalation_date'),
                 data.get('new_escalation_frequency'),data.get('new_payment_frequency'),data.get('new_payment_timing'),
                 data.get('new_first_payment_date'),data.get('old_liability'),data.get('revised_liability'),
                 data.get('rou_adjustment'),data.get('pl_gain_loss'),data.get('status','APPROVED'),username,t))
            action='CREATE'
        c.commit()
    audit(cid,username,role,action,'MODIFICATION',mid,lid,reason=data.get('mod_type',''))

def modifications_for_lease(cid,lid,approved_only=True):
    q='SELECT * FROM modifications WHERE company_id=? AND lease_id=?'+(" AND status='APPROVED'" if approved_only else '')+' ORDER BY effective_date'
    with connect() as c:return c.execute(q,(cid,lid)).fetchall()

def list_modifications(cid):
    with connect() as c:return c.execute('SELECT * FROM modifications WHERE company_id=? ORDER BY id DESC',(cid,)).fetchall()

def save_run(cid,run_id,reporting_date,input_hash,status,username,n,e):
    with connect() as c:c.execute('INSERT INTO calculation_runs(company_id,run_id,reporting_date,engine_version,input_hash,status,created_by,created_at,lease_count,error_count) VALUES(?,?,?,?,?,?,?,?,?,?)',(cid,run_id,reporting_date,'V7.3',input_hash,status,username,now(),n,e)); c.commit()
    audit(cid,username,'SYSTEM','CALCULATION_RUN','RUN',run_id,run_id=run_id)
def add_error(cid,run_id,lid,module,severity,description,expected='',actual='',formula=''):
    with connect() as c:c.execute('INSERT INTO errors(company_id,run_id,lease_id,module,severity,description,expected,actual,formula,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)',(cid,run_id,lid,module,severity,description,expected,actual,formula,now())); c.commit()
def get_errors(cid):
    with connect() as c:return c.execute('SELECT * FROM errors WHERE company_id=? ORDER BY id DESC',(cid,)).fetchall()
def get_audit(cid,limit=5000):
    with connect() as c:return c.execute('SELECT * FROM audit_trail WHERE company_id=? ORDER BY id DESC LIMIT ?',(cid,limit)).fetchall()
def lock_period(cid,rd,user):
    with connect() as c:c.execute('INSERT OR IGNORE INTO period_locks(company_id,reporting_date,locked_by,locked_at) VALUES(?,?,?,?)',(cid,rd,user,now())); c.commit()
    audit(cid,user,'APPROVER','LOCK_PERIOD','PERIOD',rd,reason='Reporting period locked')
def is_locked(cid,rd):
    with connect() as c:return bool(c.execute('SELECT 1 FROM period_locks WHERE company_id=? AND reporting_date=?',(cid,rd)).fetchone())
def create_import_batch(cid,batch,filename,total,valid,errors,status,user,input_hash='',lease_ids=None):
    with connect() as c:
        c.execute('INSERT INTO import_batches(company_id,batch_id,filename,rows_total,rows_valid,rows_error,status,created_by,created_at,input_hash,lease_ids_json) VALUES(?,?,?,?,?,?,?,?,?,?,?)',
                  (cid,batch,filename,total,valid,errors,status,user,now(),input_hash,json.dumps(list(lease_ids or []))))
        c.commit()

def latest_import_batch(cid):
    with connect() as c:
        return c.execute('SELECT * FROM import_batches WHERE company_id=? ORDER BY id DESC LIMIT 1',(cid,)).fetchone()

def update_import_batch(cid,batch_id,**kwargs):
    allowed={'status','rows_total','rows_valid','rows_error','input_hash','lease_ids_json'}
    sets=[]; vals=[]
    for k,v in kwargs.items():
        if k in allowed:
            sets.append(f'{k}=?'); vals.append(json.dumps(v) if k=='lease_ids_json' and not isinstance(v,str) else v)
    if not sets: return
    vals.extend([cid,batch_id])
    with connect() as c:
        c.execute(f'UPDATE import_batches SET {','.join(sets)} WHERE company_id=? AND batch_id=?',vals); c.commit()

def replace_active_leases_from_import(cid,rows,username,role):
    imported_ids={str(r['lease_id']).strip() for r in rows if r.get('lease_id')}
    current=[r['lease_id'] for r in lease_rows(cid,include_deleted=False)]
    for lid in current:
        if lid not in imported_ids:
            soft_delete_lease(cid,lid,username,role,'Replaced by full import batch')
    for r in rows:
        upsert_lease(cid,r,username,role)
    return imported_ids



def save_calculation_snapshot(company_id, run_id, lease_id, payload, payload_hash=None):
    if payload_hash is None:
        payload_hash=hashlib.sha256(json.dumps(payload, sort_keys=True, default=str, separators=(',',':')).encode()).hexdigest()
    with connect() as c:
        c.execute('''CREATE TABLE IF NOT EXISTS calculation_snapshots(
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_id TEXT NOT NULL, run_id TEXT NOT NULL,
            lease_id TEXT, payload_hash TEXT NOT NULL, payload_json TEXT NOT NULL, created_at TEXT NOT NULL)''')
        c.execute('INSERT INTO calculation_snapshots(company_id,run_id,lease_id,payload_hash,payload_json,created_at) VALUES(?,?,?,?,?,?)',
                  (company_id,run_id,lease_id,payload_hash,json.dumps(payload, sort_keys=True, default=str),now()))
        c.commit()
    return payload_hash


def calculation_snapshots(company_id, run_id=None):
    q='SELECT * FROM calculation_snapshots WHERE company_id=?'
    args=[company_id]
    if run_id:
        q+=' AND run_id=?'; args.append(run_id)
    q+=' ORDER BY id DESC'
    with connect() as c:return c.execute(q,args).fetchall()
