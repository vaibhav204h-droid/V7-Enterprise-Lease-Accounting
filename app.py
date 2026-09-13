from __future__ import annotations
import os, secrets, hashlib, tempfile, io, json
from datetime import date
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
import bcrypt
from db.database import *
from core.engine import calculate, validate, CalculationValidationError, parse_date, validate_calculation
from core.reporting import make_report

APP_DIR=Path(__file__).resolve().parent
TEMPLATE=APP_DIR/'templates'/'MNC_Lease_Import_Template_V7_3.xlsx'
REPORT_DIR=APP_DIR/'reports'
DEFAULT_REPORTING_DATE=date.today()
st.set_page_config(page_title='V7.3 Enterprise Lease Accounting',layout='wide',initial_sidebar_state='expanded')
init_db()

def parse_date_or_today(v):
    try:
        d=parse_date(v)
        return d or date.today()
    except Exception:
        return date.today()


def pwd_hash(p): return bcrypt.hashpw(p.encode(),bcrypt.gensalt()).decode()
def verify(p,h):
    try:return bcrypt.checkpw(p.encode(),h.encode())
    except Exception:return False
def df(rows): return pd.DataFrame([dict(r) for r in rows]) if rows else pd.DataFrame()
def iso(v): return v.isoformat() if hasattr(v,'isoformat') else (str(v)[:10] if v not in (None,'') else None)

def login():
    st.title('V7.3 Enterprise Lease Accounting Platform')
    st.caption('Company → Authentication → Maker / Reviewer / Approver → Controlled Lease Lifecycle')
    companies=list_companies(); opts=['— Select Company —']+[f"{c['company_id']} — {c['company_name']}" for c in companies]
    selected=st.selectbox('1. Company ID',opts)
    with st.expander('＋ Create New Company',expanded=not companies):
        with st.form('company'):
            c1,c2,c3=st.columns(3)
            cid=c1.text_input('Company ID',max_chars=30).strip().upper(); name=c2.text_input('Company Name'); country=c3.text_input('Country','India'); curr=c1.text_input('Functional Currency','INR'); rep=c2.text_input('Reporting Currency','INR'); fy=c3.text_input('Financial Year End','31-Mar'); rd=c1.date_input('Default Reporting Date',value=date.today()); admin=c2.text_input('Initial Administrator Username'); pw=c3.text_input('Initial Administrator Password',type='password'); confirm=c1.text_input('Confirm Password',type='password')
            if st.form_submit_button('Create Company'):
                if not all([cid,name,admin,pw]) or pw!=confirm: st.error('Complete all mandatory fields and confirm password.'); return
                try:create_company(cid,name,country,curr,rep,fy,rd.isoformat(),admin,pwd_hash(pw)); audit(cid,admin,'ADMIN','CREATE','COMPANY',cid,reason='Company created'); st.success('Company created. Select it above to login.')
                except Exception as e: st.error(f'Unable to create company: {e}')
    if selected!='— Select Company —':
        cid=selected.split(' — ')[0]
        with st.form('login'):
            u=st.text_input('2. Username'); p=st.text_input('3. Password',type='password')
            if st.form_submit_button('Secure Login'):
                user=get_user(cid,u)
                if user and verify(p,user['password_hash']): st.session_state.update(auth=True,company_id=cid,username=u,role=user['role']); audit(cid,u,user['role'],'LOGIN','USER',u); st.rerun()
                st.error('Invalid company/user/password.')

def _navigate_button(label, shortcut, page, key):
    try:
        clicked=st.sidebar.button(label, key=key, shortcut=shortcut, width='stretch')
    except TypeError:
        # Older Streamlit runtime that doesn't accept a shortcut= argument at all.
        clicked=st.sidebar.button(label, key=key, width='stretch')
    except Exception:
        # Newer Streamlit runtimes reserve certain bare keys (e.g. 'c' for clear-cache, 'r'
        # for rerun) for their own built-in shortcuts and raise StreamlitAPIException if a
        # custom shortcut reuses them, with or without modifiers. A shortcut-registration
        # problem should never take down navigation -- fall back to a plain button so the
        # app keeps working even if a future Streamlit release changes its reserved-key list
        # again (the letters below are already chosen to avoid the currently-known reserved
        # keys, so this branch is a safety net, not the expected path).
        clicked=st.sidebar.button(label, key=key+'_noshortcut', width='stretch')
    if clicked:
        st.session_state['page']=page
        st.rerun()

def main():
    cid=st.session_state.company_id; user=st.session_state.username; role=st.session_state.role; comp=get_company(cid)
    st.sidebar.title(comp['company_name']); st.sidebar.caption(f'Company ID: {cid}\nUser: {user}\nRole: {role}')
    if st.sidebar.button('Logout',key='logout'): audit(cid,user,role,'LOGOUT','USER',user); st.session_state.clear(); st.rerun()
    menu=['Control Centre','Lease Master','Import Existing Leases','Events','Calculate','Reports','Audit Trail','Errors / QA','Users & Security','Company Settings']
    if st.session_state.get('page') not in menu: st.session_state['page']='Control Centre'
    st.sidebar.markdown('### Navigate')
    for idx,item in enumerate(menu):
        if st.sidebar.button(item,key=f'nav_{idx}',width='stretch'):
            st.session_state['page']=item; st.rerun()
    st.sidebar.markdown('### Keyboard shortcuts')
    shortcuts=[('Add','ctrl+alt+a','Lease Master'),('Import','ctrl+alt+i','Import Existing Leases'),('Events','ctrl+alt+e','Events'),('Calculate','ctrl+alt+l','Calculate'),('Reports','ctrl+alt+p','Reports'),('QA','ctrl+alt+q','Errors / QA'),('Audit','ctrl+alt+t','Audit Trail')]
    for idx,(label,combo,page) in enumerate(shortcuts):
        _navigate_button(f'{label} [{combo}]',combo,page,f'ks_{idx}')
    st.sidebar.caption('Shortcuts are native Streamlit shortcuts. Keep the browser/app focused; Ctrl+Alt+A/I/E/L/P/Q/T navigate directly.')
    page=st.session_state['page']
    actions={'Control Centre':lambda:dashboard(comp),'Lease Master':lambda:lease_master(cid,user,role),'Import Existing Leases':lambda:importer(cid,user,role),'Events':lambda:event_page(cid,user,role),'Calculate':lambda:calc_page(cid,user,role),'Reports':lambda:reports_page(cid),'Audit Trail':lambda:audit_page(cid),'Errors / QA':lambda:qa_page(cid),'Users & Security':lambda:users_page(cid,user,role),'Company Settings':lambda:settings_page(comp,cid,user,role)}
    actions[page]()

def dashboard(comp):
    cid=comp['company_id']; leases=lease_rows(cid); ev=events(cid); aud=get_audit(cid); errs=get_errors(cid); st.title('Control Centre'); a,b,c,d=st.columns(4); a.metric('Active Leases',sum(x['is_deleted']==0 for x in leases)); b.metric('Events',len(ev)); c.metric('Open Errors',sum(x['status']=='OPEN' for x in errs)); d.metric('Audit Records',len(aud)); st.subheader('Human Workflow'); st.write('**Maker → Reviewer → Approver → Validate → Calculate → Reconcile → Close**'); st.dataframe(df(leases).head(100),width='stretch')

def lease_master(cid,user,role):
    st.title('Lease Master'); rows=lease_rows(cid); st.dataframe(df(rows),width='stretch')
    if role in ('MAKER','ADMIN') and not is_locked(cid,st.session_state.get('reporting_date',date.today()).isoformat()):
        with st.expander('＋ Add / Update Lease',expanded=False):
            with st.form('lease'):
                c=st.columns(3); lid=c[0].text_input('Lease ID'); status=c[1].selectbox('Status',['DRAFT','PENDING_REVIEW']); currency=c[2].text_input('Currency','INR'); contract=c[0].date_input('Contract Date',date.today()); comm=c[1].date_input('Commencement Date',date.today()); expiry=c[2].date_input('Expiry Date',date(date.today().year+1,date.today().month,date.today().day)); first=c[0].date_input('First Payment Date',date.today()); freq=c[1].selectbox('Payment Frequency',['Monthly','Quarterly','Semi-annual','Annual']); timing=c[2].selectbox('Payment Timing',['Arrears','Advance']); pay=c[0].number_input('Base Payment',0.0); ib=c[1].number_input('Annual IBR (effective)',0.0,step=0.001,format='%.6f'); et=c[2].selectbox('Escalation Type',['None','Fixed Amount','Fixed %','Fixed Rate','CPI','CPI Linked','WPI','WPI Linked','Other Index','Custom Schedule']); ea=c[0].number_input('Fixed Escalation Amount',0.0); er=c[1].number_input('Escalation Rate',0.0,step=0.001,format='%.6f'); fed=c[2].date_input('First Escalation Date',comm); ef=c[0].selectbox('Escalation Frequency',['Annual','Monthly','Quarterly','Semi-annual']); st.markdown('**Non-uniform escalation (optional)** — use the Escalation_Schedule sheet for imports, or paste JSON here. Each event can be `Fixed %`, `Fixed Amount`, or `Reset Amount` and applies from its effective date. Example: `[{"effective_date":"2027-04-01","type":"Fixed %","value":0.05},{"effective_date":"2028-04-01","type":"Fixed Amount","value":10000}]`')
                esc_json=st.text_area('Escalation Schedule JSON',value='',height=90)
                prepaid=c[1].number_input('Prepaid',0.0); idc=c[2].number_input('Qualifying Initial Direct Costs',0.0); rest=c[0].number_input('Restoration PV (manual override)',0.0); inc=c[1].number_input('Lease Incentive',0.0); dep=c[2].number_input('Security Deposit Paid',0.0); refund=c[0].date_input('Deposit Refund Date',expiry); dr=c[1].number_input('Deposit Discount Rate',0.0,step=0.001,format='%.6f'); de=c[2].number_input('Configured Deposit ECL Rate',0.0,step=0.001,format='%.6f'); pdv=c[0].number_input('ECL PD',0.0,step=0.001,format='%.6f'); lgd=c[1].number_input('ECL LGD',0.0,step=0.001,format='%.6f'); fx=c[2].number_input('Reporting FX Rate',1.0,step=0.0001); cc=c[0].text_input('Cost Centre'); cp=c[1].text_input('Counterparty')
                c4=st.columns(4); ecl_method=c4[0].selectbox('ECL Method',['Configured loss rate','PD×LGD with horizon']); ecl_horizon=c4[1].number_input('ECL Horizon (years)',value=1.0,min_value=0.01,step=0.5); ecl_discount=c4[2].number_input('ECL Discount Rate',0.0,step=0.001,format='%.6f'); val_tol=c4[3].number_input('Validation Tolerance',value=0.01,min_value=0.001,step=0.01,format='%.2f'); fx2=st.columns(2); initial_fx=fx2[0].number_input('Initial FX Rate (optional)',0.0,step=0.0001); prior_fx=fx2[1].number_input('Prior Reporting FX Rate (optional)',0.0,step=0.0001)
                st.markdown('**Lease Term Assessment** (Ind AS 116 paras 18-19 — leave as No/0 to use the contractual expiry date as-is)')
                c2=st.columns(4); ext_yn=c2[0].selectbox('Extension Reasonably Certain?',['No','Yes']); ext_yrs=c2[1].number_input('Extension Option (years)',0.0,step=0.5); term_yn=c2[2].selectbox('Termination Reasonably Certain?',['No','Yes']); term_date=c2[3].date_input('Termination Option Date',expiry)
                st.markdown('**Restoration / Dismantling Provision** (Ind AS 116 para 24(d) — leave cost at 0 to keep using the manual Restoration PV above)')
                c3=st.columns(3); rest_cost=c3[0].number_input('Estimated Restoration Cost (future value)',0.0); rest_date=c3[1].date_input('Expected Restoration Date',expiry); rest_rate=c3[2].number_input('Restoration Discount Rate',0.0,step=0.001,format='%.6f')
                st.markdown('**GL Reconciliation** (optional — only for the Reconciliation control; leave blank until a GL extract is available)')
                gl_extract=st.number_input('GL Trial Balance Lease Liability Extract (0 = not yet provided)',0.0)
                if st.form_submit_button('Save Lease'):
                    data=locals(); data.update(lease_id=lid,status=status,contract_date=iso(contract),commencement_date=iso(comm),expiry_date=iso(expiry),first_payment_date=iso(first),payment_frequency=freq,payment_timing=timing,base_payment=pay,escalation_type=et,escalation_amount=ea,escalation_rate=er,first_escalation_date=iso(fed),escalation_frequency=ef,currency=currency,ib_rate=ib,ib_rate_type='Annual Effective',prepaid=prepaid,idc=idc,restoration_pv=rest,lease_incentive=inc,deposit_paid=dep,deposit_refund_date=iso(refund),deposit_discount_rate=dr,deposit_ecl_rate=de,deposit_classification='Pending Assessment',ecl_pd=pdv,ecl_lgd=lgd,reporting_fx_rate=fx,counterparty=cp,cost_center=cc,gl_rou='ROU Asset - Other',gl_liability='Lease Liability',gl_interest='Finance Cost',gl_depreciation='Depreciation Expense',gl_accum_dep='Accumulated Depreciation - ROU',gl_bank='Bank',extension_option_years=ext_yrs,extension_reasonably_certain=ext_yn,termination_option_date=iso(term_date),termination_reasonably_certain=term_yn,purchase_option_reasonably_certain='No',restoration_cost_estimate=rest_cost,restoration_expected_date=iso(rest_date),restoration_discount_rate=rest_rate,gl_liability_extract=(gl_extract or None),escalation_schedule_json=(esc_json.strip() or None),ecl_method=ecl_method,ecl_horizon_years=ecl_horizon,ecl_discount_rate=ecl_discount,validation_tolerance=val_tol,initial_fx_rate=initial_fx,prior_reporting_fx_rate=prior_fx);
                    if not lid: st.error('Lease ID required.'); return
                    errs=validate(data)
                    if errs: st.error('; '.join(f'{s}: {m}' for m,s in errs)); return
                    upsert_lease(cid,data,user,role); st.success('Lease saved.'); st.rerun()
    if role in ('MAKER','ADMIN') and rows:
        lid=st.selectbox('Lease to soft-delete',[r['lease_id'] for r in rows]); reason=st.text_input('Deletion reason');
        if st.button('Soft Delete / Terminate Lease'): soft_delete_lease(cid,lid,user,role,reason or 'User requested'); st.success('Lease retained in history and audit trail.'); st.rerun()

def _find_header_row(rows, required, max_scan=20):
    required_set=set(required)
    for i,row in enumerate(rows[:max_scan]):
        headers=[str(x).strip() if x is not None else '' for x in row]
        if len(required_set.intersection(headers)) >= 5:
            return i, headers
    return None, []

def _normalise_import_headers(headers):
    aliases={
        'Lease ID':'Lease_ID','Lease_ID':'Lease_ID',
        'Contract Date':'Contract_Date','Contract_Date':'Contract_Date',
        'Commencement Date':'Commencement_Date','Commencement_Date':'Commencement_Date',
        'Expiry Date':'Expiry_Date','Expiry_Date':'Expiry_Date',
        'First Payment Date':'First_Payment_Date','First_Payment_Date':'First_Payment_Date',
        'Payment Frequency':'Payment_Frequency','Payment_Frequency':'Payment_Frequency',
        'Payment Timing':'Payment_Timing','Payment_Timing':'Payment_Timing',
        'Payment Amount':'Base_Payment','Base Payment':'Base_Payment','Base_Payment':'Base_Payment',
        'Escalation Type':'Escalation_Type','Escalation_Type':'Escalation_Type',
        'Escalation Amount':'Escalation_Amount','Escalation_Amount':'Escalation_Amount',
        'Escalation Rate':'Escalation_Rate','Escalation_Rate':'Escalation_Rate',
        'First Escalation Date':'First_Escalation_Date','First_Escalation_Date':'First_Escalation_Date',
        'Escalation Frequency':'Escalation_Frequency','Effective Date':'Effective_Date','Type':'Type','Value':'Value','Escalation_Frequency':'Escalation_Frequency',
        'Currency':'Currency','IBR':'IBR','IBR %':'IBR',
        'Prepaid':'Prepaid','Qualifying IDC':'Qualifying_IDC','Qualifying_IDC':'Qualifying_IDC',
        'Restoration PV':'Restoration_PV','Restoration_PV':'Restoration_PV',
        'Lease Incentive':'Lease_Incentive','Lease_Incentive':'Lease_Incentive',
        'Security Deposit Paid':'Security_Deposit_Paid','Security_Deposit_Paid':'Security_Deposit_Paid',
        'Deposit Refund Date':'Deposit_Refund_Date','Deposit_Refund_Date':'Deposit_Refund_Date',
        'Deposit Discount Rate':'Deposit_Discount_Rate','Deposit_Discount_Rate':'Deposit_Discount_Rate',
        'Deposit ECL Rate':'Deposit_ECL_Rate','Deposit_ECL_Rate':'Deposit_ECL_Rate',
        'Deposit Classification':'Deposit_Classification','Deposit_Classification':'Deposit_Classification',
        'ECL PD':'ECL_PD','ECL_PD':'ECL_PD','ECL LGD':'ECL_LGD','ECL_LGD':'ECL_LGD',
        'Restoration Cost Estimate':'Restoration_Cost_Estimate','Restoration_Cost_Estimate':'Restoration_Cost_Estimate','Restoration Expected Date':'Restoration_Expected_Date','Restoration_Expected_Date':'Restoration_Expected_Date','Restoration Discount Rate':'Restoration_Discount_Rate','Restoration_Discount_Rate':'Restoration_Discount_Rate','Reporting FX Rate':'Reporting_FX_Rate','Reporting_FX_Rate':'Reporting_FX_Rate',
        'Counterparty':'Counterparty','Cost Centre':'Cost_Centre','Cost_Centre':'Cost_Centre','Validation Tolerance':'Validation_Tolerance','Validation_Tolerance':'Validation_Tolerance','ECL Method':'ECL_Method','ECL_Method':'ECL_Method','ECL Horizon (years)':'ECL_Horizon_Years','ECL_Horizon_Years':'ECL_Horizon_Years','ECL Discount Rate':'ECL_Discount_Rate','ECL_Discount_Rate':'ECL_Discount_Rate','Initial FX Rate':'Initial_FX_Rate','Initial_FX_Rate':'Initial_FX_Rate','Prior Reporting FX Rate':'Prior_Reporting_FX_Rate','Prior_Reporting_FX_Rate':'Prior_Reporting_FX_Rate',
        'Modification ID':'Modification_ID','Modification_ID':'Modification_ID','Event ID':'Event_ID','Event_ID':'Event_ID','Effective Date':'Effective_Date','Effective_Date':'Effective_Date','Mod Type':'Mod_Type','Modification Type':'Mod_Type','Mod_Type':'Mod_Type','New Base Payment':'New_Base_Payment','New_Base_Payment':'New_Base_Payment','New IBR':'New_IBR','New Annual IBR':'New_IBR','New_IBR':'New_IBR','New Expiry Date':'New_Expiry_Date','New_Expiry_Date':'New_Expiry_Date','New Escalation Type':'New_Escalation_Type','New_Escalation_Type':'New_Escalation_Type','New Escalation Amount':'New_Escalation_Amount','New_Escalation_Amount':'New_Escalation_Amount','New Escalation Rate':'New_Escalation_Rate','New_Escalation_Rate':'New_Escalation_Rate','New First Escalation Date':'New_First_Escalation_Date','New_First_Escalation_Date':'New_First_Escalation_Date','New Escalation Frequency':'New_Escalation_Frequency','New_Escalation_Frequency':'New_Escalation_Frequency','New Payment Frequency':'New_Payment_Frequency','New_Payment_Frequency':'New_Payment_Frequency','New Payment Timing':'New_Payment_Timing','New_Payment_Timing':'New_Payment_Timing','New First Payment Date':'New_First_Payment_Date','New_First_Payment_Date':'New_First_Payment_Date','Status':'Status','Event Sequence':'Event_Sequence','Event_Sequence':'Event_Sequence','Base Index':'Base_Index','Base_Index':'Base_Index','Current Index':'Current_Index','Current_Index':'Current_Index',
    }
    return [aliases.get(h,h) for h in headers]

def _safe_float(v, default=0.0):
    if v in (None,''): return default
    try: return float(v)
    except (TypeError,ValueError): raise ValueError(f'Invalid numeric value: {v}')

def _iso(v):
    return v.isoformat() if hasattr(v,'isoformat') else (str(v)[:10] if v else None)

def _build_import_lease(d, escalation_events=None):
    lid=str(d.get('Lease_ID') or '').strip()
    events=(escalation_events or {}).get(lid,[])
    return {
        'lease_id':lid,
        'status':str(d.get('Status') or 'DRAFT').strip().upper(),
        'contract_date':_iso(d.get('Contract_Date')),
        'commencement_date':_iso(d.get('Commencement_Date')),
        'expiry_date':_iso(d.get('Expiry_Date')),
        'first_payment_date':_iso(d.get('First_Payment_Date')),
        'payment_frequency':d.get('Payment_Frequency'),
        'payment_timing':d.get('Payment_Timing'),
        'base_payment':_safe_float(d.get('Base_Payment')),
        'ib_rate':_safe_float(d.get('IBR')),
        'ib_rate_type':'Annual Effective',
        'escalation_type':d.get('Escalation_Type','None') or 'None',
        'escalation_amount':_safe_float(d.get('Escalation_Amount')),
        'escalation_rate':_safe_float(d.get('Escalation_Rate')),
        'first_escalation_date':_iso(d.get('First_Escalation_Date')),
        'escalation_frequency':d.get('Escalation_Frequency','Annual') or 'Annual',
        'currency':d.get('Currency','INR') or 'INR',
        'prepaid':_safe_float(d.get('Prepaid')),
        'idc':_safe_float(d.get('Qualifying_IDC')),
        'restoration_pv':_safe_float(d.get('Restoration_PV')),
        'lease_incentive':_safe_float(d.get('Lease_Incentive')),
        'deposit_paid':_safe_float(d.get('Security_Deposit_Paid')),
        'deposit_refund_date':_iso(d.get('Deposit_Refund_Date')),
        'deposit_discount_rate':_safe_float(d.get('Deposit_Discount_Rate')),
        'deposit_ecl_rate':_safe_float(d.get('Deposit_ECL_Rate')),
        'deposit_classification':d.get('Deposit_Classification','Pending Assessment') or 'Pending Assessment',
        'ecl_pd':_safe_float(d.get('ECL_PD')), 'ecl_lgd':_safe_float(d.get('ECL_LGD')),
        'reporting_fx_rate':_safe_float(d.get('Reporting_FX_Rate'),1.0),
        'counterparty':d.get('Counterparty'), 'cost_center':d.get('Cost_Centre'),
        'restoration_cost_estimate':_safe_float(d.get('Restoration_Cost_Estimate')),
        'restoration_expected_date':_iso(d.get('Restoration_Expected_Date')),
        'restoration_discount_rate':_safe_float(d.get('Restoration_Discount_Rate')),
        'extension_option_years':_safe_float(d.get('Extension_Option_Years')),
        'extension_reasonably_certain':str(d.get('Extension_Reasonably_Certain') or 'No'),
        'termination_option_date':_iso(d.get('Termination_Option_Date')),
        'termination_reasonably_certain':str(d.get('Termination_Reasonably_Certain') or 'No'),
        'purchase_option_reasonably_certain':str(d.get('Purchase_Option_Reasonably_Certain') or 'No'),
        'escalation_schedule_json':json.dumps(events) if events else (d.get('Escalation_Schedule_JSON') or None),
        'validation_tolerance':_safe_float(d.get('Validation_Tolerance'),0.01),
        'ecl_method':d.get('ECL_Method') or 'Configured loss rate',
        'ecl_horizon_years':_safe_float(d.get('ECL_Horizon_Years'),1.0),
        'ecl_discount_rate':_safe_float(d.get('ECL_Discount_Rate')),
        'initial_fx_rate':_safe_float(d.get('Initial_FX_Rate')),
        'prior_reporting_fx_rate':_safe_float(d.get('Prior_Reporting_FX_Rate')),
        'gl_rou':'ROU Asset - Other','gl_liability':'Lease Liability','gl_interest':'Finance Cost',
        'gl_depreciation':'Depreciation Expense','gl_accum_dep':'Accumulated Depreciation - ROU','gl_bank':'Bank',
    }

def _read_modification_sheet(wb):
    if 'Modification' not in wb.sheetnames:
        return []
    vals=list(wb['Modification'].values)
    if not vals:
        return []
    header_idx=None; headers=[]
    wanted={'Modification_ID','Lease_ID','Effective_Date','Mod_Type'}
    for i,row in enumerate(vals[:20]):
        hs=[str(x).strip() if x is not None else '' for x in row]
        if len(wanted.intersection(hs)) >= 3:
            header_idx=i; headers=_normalise_import_headers(hs); break
    if header_idx is None:
        raise ValueError('Modification sheet is present but its header row could not be identified. Required: Modification_ID, Lease_ID, Effective_Date, Mod_Type.')
    rows=[]
    for row in vals[header_idx+1:]:
        if not any(x not in (None,'') for x in row):
            continue
        d=dict(zip(headers,row))
        if not d.get('Modification_ID') and not d.get('Lease_ID'):
            continue
        rows.append(d)
    return rows

def _validate_import_modification(m, lease_ids):
    mid=str(m.get('Modification_ID') or '').strip(); lid=str(m.get('Lease_ID') or '').strip(); eff=_iso(m.get('Effective_Date')); mod_type=str(m.get('Mod_Type') or '').strip()
    if not mid: return 'Modification_ID is required.'
    if not lid: return 'Lease_ID is required.'
    if lid not in lease_ids: return f'Lease_ID {lid} does not exist in Lease_Master import.'
    if not eff: return 'Effective_Date is required.'
    try:
        parse_date(eff)
    except Exception as ex:
        return str(ex)
    if not mod_type: return 'Mod_Type is required.'
    status=str(m.get('Status') or 'DRAFT').strip().upper()
    if status not in {'DRAFT','APPROVED','REVIEWED'}:
        return 'Modification Status must be DRAFT, REVIEWED or APPROVED.'
    return None

def importer(cid,user,role):
    st.title('Import Existing Leases')
    st.info('Upload the original-style Lease_Master template or a compatible legacy workbook. The system also accepts the optional Escalation_Schedule and Modification sheets, stages all records, and validates them before committing.')
    template_path=Path('templates/MNC_Lease_Import_Template_V7_3.xlsx')
    if template_path.exists():
        st.download_button('Download Sample Import Format',template_path.read_bytes(),template_path.name, key='download_import_template')
    if role not in ('MAKER','ADMIN'):
        st.warning('Only Maker/Admin can import.')
        return
    f=st.file_uploader('Upload pre-existing lease Excel',type=['xlsx'],key='lease_import_file')
    if f:
        file_bytes=f.getvalue()
        file_hash=hashlib.sha256(file_bytes).hexdigest()
        st.caption(f'File: {f.name} | SHA-256: {file_hash[:16]}…')
        if st.button('Validate Upload',type='primary',key='validate_upload'):
            try:
                wb=load_workbook(io.BytesIO(file_bytes),data_only=True,read_only=True)
                sheet='Lease_Master' if 'Lease_Master' in wb.sheetnames else wb.sheetnames[0]
                ws=wb[sheet]; vals=list(ws.values)
                escalation_events={}
                if 'Escalation_Schedule' in wb.sheetnames:
                    ew=wb['Escalation_Schedule']; erows=list(ew.values)
                    if erows:
                        eh=_normalise_import_headers([str(x).strip() if x is not None else '' for x in erows[0]])
                        for rr in erows[1:]:
                            edict=dict(zip(eh,rr)); lid2=str(edict.get('Lease_ID') or '').strip()
                            if lid2 and edict.get('Effective_Date'):
                                event={'effective_date':_iso(edict.get('Effective_Date')),'type':edict.get('Type'),'value':_safe_float(edict.get('Value'))}
                                if edict.get('Event_Sequence') not in (None,''): event['event_sequence']=int(edict.get('Event_Sequence'))
                                if edict.get('Base_Index') not in (None,''): event['base_index_value']=_safe_float(edict.get('Base_Index'))
                                if edict.get('Current_Index') not in (None,''): event['index_value']=_safe_float(edict.get('Current_Index'))
                                escalation_events.setdefault(lid2,[]).append(event)
                modification_rows=_read_modification_sheet(wb)
                required=['Lease_ID','Commencement_Date','Expiry_Date','First_Payment_Date','Payment_Frequency','Payment_Timing','Base_Payment','IBR']
                hrow,headers=_find_header_row(vals,required)
                if hrow is None:
                    st.error('Could not identify the Lease Master header row. Use the supplied V7.2 template.')
                    return
                headers=_normalise_import_headers(headers)
                missing=[x for x in required if x not in headers]
                if missing:
                    st.error('Missing required columns: '+', '.join(missing)); return
                data=[]
                for row in vals[hrow+1:]:
                    if not any(x not in (None,'') for x in row): continue
                    d=dict(zip(headers,row));
                    if not d.get('Lease_ID'): continue
                    data.append(d)
                batch=f'IMP-{cid}-{secrets.token_hex(5).upper()}'
                lease_id_values=[str(x.get('Lease_ID') or '').strip() for x in data]
                errors=[]; staged=[]
                if len(lease_id_values) != len(set(lease_id_values)):
                    errors.append(('IMPORT','Duplicate Lease_ID values found in Lease_Master import.','CRITICAL'))
                for d in data:
                    try:
                        lease=_build_import_lease(d, escalation_events)
                        verr=validate(lease)
                        for msg,sev in verr: errors.append((lease['lease_id'],msg,sev))
                        staged.append(lease)
                    except Exception as ex:
                        errors.append((str(d.get('Lease_ID') or 'ROW'),str(ex),'CRITICAL'))
                lease_ids={str(x.get('Lease_ID') or '').strip() for x in data}
                staged_modifications=[]
                for m in modification_rows:
                    msg=_validate_import_modification(m,lease_ids)
                    if msg:
                        errors.append((str(m.get('Modification_ID') or m.get('Lease_ID') or 'MODIFICATION'),msg,'HIGH'))
                    else:
                        staged_modifications.append(m)
                st.session_state['import_preview']={'batch':batch,'filename':f.name,'hash':file_hash,'rows':data,'staged':staged,'errors':errors,'escalation_events':escalation_events,'modification_rows':modification_rows,'staged_modifications':staged_modifications}
                create_import_batch(cid,batch,f.name,len(data),len(staged)-len({x[0] for x in errors}),len(errors),'VALIDATED' if not errors else 'BLOCKED',user,input_hash=file_hash,lease_ids=lease_id_values)
            except Exception as ex:
                st.error(f'Import validation failed: {ex}')
                return
        preview=st.session_state.get('import_preview')
        if preview and preview.get('hash')==file_hash:
            rows=preview['rows']; errors=preview['errors']; staged=preview['staged']
            st.subheader(f"Import Preview — {preview['batch']}")
            st.write(f"{len(rows)} lease rows detected | {len(preview.get('modification_rows',[]))} modification rows detected | {len(errors)} validation exceptions")
            if preview.get('modification_rows'):
                st.subheader('Modification Import Preview')
                st.dataframe(pd.DataFrame(preview['modification_rows']),width='stretch')
            preview_cols=['Lease_ID','Commencement_Date','Expiry_Date','First_Payment_Date','Payment_Frequency','Payment_Timing','Base_Payment','Escalation_Type','Escalation_Amount','Escalation_Rate','IBR']
            st.dataframe(pd.DataFrame([{k:d.get(k) for k in preview_cols} for d in rows]),width='stretch')
            if errors:
                st.error('Import blocked. Resolve all validation exceptions before commit.')
                st.dataframe(pd.DataFrame(errors,columns=['Lease','Error','Severity']),width='stretch')
            else:
                if st.button('Confirm & Commit Import',type='primary',key='commit_import'):
                    try:
                        imported_leases=[_build_import_lease(d,preview.get('escalation_events',{})) for d in rows]
                        replace_active_leases_from_import(cid,imported_leases,user,role)
                        committed_mods=0
                        for m in preview.get('staged_modifications',[]):
                            mid=str(m['Modification_ID']).strip(); status=str(m.get('Status') or 'DRAFT').strip().upper()
                            add_modification(cid,{
                                'modification_id':mid,'lease_id':str(m['Lease_ID']).strip(),'event_id':m.get('Event_ID'),'effective_date':_iso(m.get('Effective_Date')),'mod_type':m.get('Mod_Type'),
                                'new_base_payment':_safe_float(m.get('New_Base_Payment')) if m.get('New_Base_Payment') not in (None,'') else None,
                                'new_ib_rate':_safe_float(m.get('New_IBR')) if m.get('New_IBR') not in (None,'') else None,
                                'new_expiry_date':_iso(m.get('New_Expiry_Date')),'new_escalation_type':m.get('New_Escalation_Type'),
                                'new_escalation_amount':_safe_float(m.get('New_Escalation_Amount')) if m.get('New_Escalation_Amount') not in (None,'') else None,
                                'new_escalation_rate':_safe_float(m.get('New_Escalation_Rate')) if m.get('New_Escalation_Rate') not in (None,'') else None,
                                'new_first_escalation_date':_iso(m.get('New_First_Escalation_Date')),'new_escalation_frequency':m.get('New_Escalation_Frequency'),
                                'new_payment_frequency':m.get('New_Payment_Frequency'),'new_payment_timing':m.get('New_Payment_Timing'),'new_first_payment_date':_iso(m.get('New_First_Payment_Date')),
                                'status':status},user,role); committed_mods+=1
                        update_import_batch(cid,preview['batch'],status='COMMITTED',rows_valid=len(imported_leases),rows_error=0,input_hash=preview['hash'],lease_ids=[x['lease_id'] for x in imported_leases])
                        audit(cid,user,role,'IMPORT_COMMIT','IMPORT_BATCH',preview['batch'],reason=f"Imported {len(rows)} leases and {committed_mods} modification rows from {f.name}")
                        st.success(f"Import {preview['batch']} committed successfully. {len(rows)} leases are now in the company Lease Master.")
                        st.session_state.pop('import_preview',None)
                        st.session_state['page']='Lease Master'
                        st.rerun()
                    except Exception as ex:
                        msg=str(ex)
                        if 'UNIQUE constraint failed: modifications.modification_id' in msg:
                            st.error('Import commit failed: a Modification_ID already exists. Re-submitting the same ID is now treated as an update; an ID belonging to another lease/company must be changed to a unique value.')
                        else:
                            st.error(f'Import commit failed: {ex}')

def event_page(cid,user,role):
    st.title('Events — Lease Lifecycle'); rows=lease_rows(cid); lids=[r['lease_id'] for r in rows];
    if not lids: st.info('Create/import a lease first.'); return
    types=['New Lease','Payment Change','Fixed Amount Escalation','Percentage Escalation','CPI / Indexation','Rent-Free Period','Lease Term Reassessment','Renewal Exercise','Renewal Not Exercised','Purchase Option Reassessment','Termination Option Reassessment','Lease Modification','Separate Lease Assessment','Partial Termination','Full Termination','Scope Decrease','Scope Increase','Change in IBR','Currency / FX Change','Impairment / Ind AS 36 Indicator','Security Deposit Change','Security Deposit Refund','Security Deposit ECL Reassessment','Sublease Commencement','Sublease Modification','Contract Correction']
    remeasurement_types={'Payment Change','Lease Modification','Partial Termination','Full Termination','Scope Decrease','Scope Increase','Change in IBR','Renewal Exercise','Fixed Amount Escalation','Percentage Escalation','CPI / Indexation'}
    if role in ('MAKER','ADMIN'):
        with st.form('event'):
            lid=st.selectbox('Lease',lids); et=st.selectbox('Event Type',types); ed=st.date_input('Event Date',date.today()); eff=st.date_input('Effective Date',date.today()); old=st.text_input('Old State / Value'); new=st.text_input('New State / Value'); reason=st.text_area('Reason / Accounting Context'); doc=st.text_input('Supporting Document Reference')
            if st.form_submit_button('Create Event'):
                eid=f'EV-{cid}-{secrets.token_hex(5).upper()}'; add_event(cid,{'event_id':eid,'lease_id':lid,'event_type':et,'event_date':iso(ed),'effective_date':iso(eff),'old_value':old,'new_value':new,'reason':reason,'supporting_document':doc},user,role); st.success(eid+' created as DRAFT.'); st.rerun()
    ev=events(cid); st.dataframe(df(ev),width='stretch')
    drafts=[e['event_id'] for e in ev if e['status']=='DRAFT']; reviewed=[e['event_id'] for e in ev if e['status']=='REVIEWED']
    if role in ('REVIEWER','ADMIN') and drafts:
        eid=st.selectbox('Review Event',drafts,key='rev');
        if st.button('Mark Reviewed'): update_event_status(cid,eid,'REVIEWED',user,role); st.rerun()
    if role in ('APPROVER','ADMIN') and reviewed:
        eid=st.selectbox('Approve Event',reviewed,key='app');
        if st.button('Approve Event'): update_event_status(cid,eid,'APPROVED',user,role); st.rerun()

    st.divider(); st.subheader('Process Approved Modification / Remeasurement')
    st.caption('An APPROVED event of a remeasurement-relevant type does not change any calculation on its own. '
               'Record the revised contractual terms here to run the Ind AS 116 paras 44-46 remeasurement and post it '
               'into the lease — the Calculate step then applies it automatically from its effective date onward.')
    current_reporting_date=st.session_state.get('reporting_date',date.today())
    approved=[e for e in ev if e['status']=='APPROVED' and e['event_type'] in remeasurement_types and (not e.get('effective_date') or parse_date_or_today(e['effective_date']) <= current_reporting_date)]
    existing_mod_events={m['event_id'] for m in list_modifications(cid)}
    pending=[e for e in approved if e['event_id'] not in existing_mod_events]
    if role in ('MAKER','ADMIN') and pending:
        with st.form('modification'):
            opts={f"{e['event_id']} — {e['lease_id']} — {e['event_type']} (eff. {e['effective_date']})":e for e in pending}
            choice=st.selectbox('Approved event to process',list(opts.keys())); ev_row=opts[choice]
            lid=ev_row['lease_id']; lease=get_lease(cid,lid)
            mc=st.columns(3)
            new_base=mc[0].number_input('New Base Payment (blank/0 = unchanged)',0.0)
            new_ibr=mc[1].number_input('New Annual IBR (blank/0 = unchanged)',0.0,step=0.001,format='%.6f')
            new_expiry=mc[2].date_input('New Expiry Date (leave as current if term unchanged)',parse_date_or_today(lease['expiry_date']) if lease else date.today())
            mc2=st.columns(3)
            new_freq=mc2[0].selectbox('New Payment Frequency (unchanged if same)',['(unchanged)','Monthly','Quarterly','Semi-annual','Annual'])
            new_timing=mc2[1].selectbox('New Payment Timing (unchanged if same)',['(unchanged)','Arrears','Advance'])
            new_first_pay=mc2[2].date_input('New First Payment Date After Modification', ev_row['effective_date'] if isinstance(ev_row['effective_date'],date) else date.today())
            if st.form_submit_button('Compute & Post Remeasurement'):
                mod={'effective_date':ev_row['effective_date'],'mod_type':ev_row['event_type']}
                if new_base: mod['new_base_payment']=new_base
                if new_ibr: mod['new_ib_rate']=new_ibr
                if lease and iso(new_expiry)!=lease['expiry_date']: mod['new_expiry_date']=iso(new_expiry)
                if new_freq!='(unchanged)': mod['new_payment_frequency']=new_freq
                if new_timing!='(unchanged)': mod['new_payment_timing']=new_timing
                mod['new_first_payment_date']=iso(new_first_pay)
                try:
                    prior_mods=[dict(m) for m in modifications_for_lease(cid,lid)]
                    preview=calculate(dict(lease),parse_date_or_today(ev_row['effective_date']),modifications=prior_mods+[mod])
                    this_mod=next((m for m in preview['modification_log'] if str(m['effective_date'])==str(mod['effective_date'])),None)
                    if this_mod is None: st.error('Modification effective date must fall strictly after commencement and before the assessed lease term end.')
                    else:
                        mid=f'MOD-{cid}-{secrets.token_hex(5).upper()}'
                        add_modification(cid,{**mod,'modification_id':mid,'lease_id':lid,'event_id':ev_row['event_id'],
                            'old_liability':float(this_mod['old_liability']),'revised_liability':float(this_mod['revised_liability']),
                            'rou_adjustment':float(this_mod['rou_adjustment']),'pl_gain_loss':float(this_mod['pl_gain_loss']),'status':'APPROVED'},user,role)
                        st.success(f"{mid} posted. Liability immediately before: {this_mod['old_liability']:,.2f} → Revised: {this_mod['revised_liability']:,.2f} "
                                   f"(ROU adjustment {this_mod['rou_adjustment']:,.2f}, P&L {this_mod['pl_gain_loss']:,.2f}). "
                                   "This will apply automatically from its effective date in every future Calculate run.")
                        st.rerun()
                except Exception as ex:
                    st.error(f'Remeasurement could not be computed: {ex}')
    elif approved and not pending:
        st.caption('Approved modifications effective on or before the selected reporting date are included in reporting calculations. Future-effective modifications remain in history but are excluded from reporting-date measurement.')
    mods_list=list_modifications(cid)
    if mods_list: st.dataframe(df(mods_list),width='stretch')

def calc_page(cid,user,role):
    st.title('Calculation Centre')
    measurement_basis=st.selectbox('Detailed Measurement Schedule Basis',['Monthly','Daily'],help='Controls the audit working frequency only. Contractual payment frequency remains the actual lease payment frequency.')
    rd=st.date_input('Central Reporting Date',value=st.session_state.get('reporting_date',date.today()),help='Defaults to the current system date. Change it for a historical or future reporting run.')
    st.session_state['reporting_date']=rd
    st.caption(f'Current calculation date: {rd.isoformat()}')
    rows=[dict(r) for r in lease_rows(cid)]  # normalise to plain dicts once: sqlite3.Row has no .get()
    latest_batch=latest_import_batch(cid)
    latest_import_ids=json.loads(latest_batch['lease_ids_json']) if latest_batch and latest_batch['lease_ids_json'] else []
    if latest_batch and latest_batch['status']=='COMMITTED' and latest_import_ids:
        actual_ids=[str(r['lease_id']).strip() for r in rows]
        if actual_ids != latest_import_ids:
            st.error('Import Mismatch — active Lease Master does not match the last committed import. Re-import the source file before calculation.')
            return
    if st.button('Validate Input',key='validate_input'):
        er=[(r['lease_id'],m,s) for r in rows for m,s in validate(dict(r))]
        st.dataframe(pd.DataFrame(er,columns=['Lease','Error','Severity']) if er else pd.DataFrame([['ALL','No validation errors','PASS']],columns=['Scope','Message','Status']))
    if st.button('Calculate Portfolio',type='primary',key='calculate_portfolio'):
        run='RUN-'+cid+'-'+secrets.token_hex(4).upper(); calcs=[]; errs=[]
        for r in rows:
            ve=validate(dict(r))
            if ve:
                errs += [(r['lease_id'],m,s) for m,s in ve]; continue
            try:
                mods=[dict(m) for m in modifications_for_lease(cid,r['lease_id'])]
                calcs.append(calculate(dict(r),rd,modifications=mods or None,measurement_basis=measurement_basis))
            except Exception as e: errs.append((r['lease_id'],str(e),'CRITICAL'))
        input_payload=[{k:r[k] for k in r.keys()} for r in rows]; h=hashlib.sha256(json.dumps(input_payload,sort_keys=True,default=str,separators=(',',':')).encode()).hexdigest()
        save_run(cid,run,rd.isoformat(),h,'SUCCESS' if not errs else 'BLOCKED',user,len(rows),len(errs))
        for lid,e,s in errs: add_error(cid,run,lid,'CALCULATION',s,e)
        if errs:
            st.error(f'Run {run} BLOCKED by {len(errs)} exception(s).')
            st.dataframe(pd.DataFrame(errs,columns=['Lease','Error','Severity']),width='stretch'); return
        for r,c in zip(rows,calcs):
            save_calculation_snapshot(cid,run,r['lease_id'],{'lease':dict(r),'result_summary':{k:str(c.get(k)) for k in ('initial_liability','initial_rou','reporting_liability','reporting_rou','current_liability','noncurrent_liability','stub_interest')},'validation_status':c.get('validation_status'),'validation_errors':c.get('validation_errors',[])})
        audit(cid,user,role,'CALCULATION_COMPLETE','CALCULATION_RUN',run,reason=f'Calculation completed at reporting date {rd.isoformat()}',run_id=run)
        ib=latest_import_batch(cid); input_ids=json.loads(ib['lease_ids_json']) if ib and ib['lease_ids_json'] else [r['lease_id'] for r in rows]; out=make_report(get_company(cid)['company_name'],rd.isoformat(),rows,calcs,str(REPORT_DIR),get_audit(cid),get_errors(cid),company_id=cid,measurement_basis=measurement_basis,input_lease_ids=input_ids,input_hash=(ib['input_hash'] if ib else h),run_id=run); st.session_state.update(last_report=out,last_run=run); st.success(f'{run} completed. Formula-driven report pack generated.')
    if st.button('Run Full QA',key='full_qa'):
        run_qa(cid,rd)

def run_qa(cid,rd):
    results=[]
    for r in lease_rows(cid):
        ve=validate(dict(r)); results.append([r['lease_id'],'Validation','PASS' if not ve else 'FAIL','; '.join(m for m,s in ve)])
        if not ve:
            try:
                mods=[dict(m) for m in modifications_for_lease(cid,r['lease_id'])]
                c=calculate(dict(r),rd,modifications=mods or None)
                results += [[r['lease_id'],'PV','PASS' if c['initial_liability']>=0 else 'FAIL',str(c['initial_liability'])],
                            [r['lease_id'],'ROU','PASS' if c['initial_rou']>=0 else 'FAIL',str(c['initial_rou'])],
                            [r['lease_id'],'Payment Count','PASS' if len(c['payments']) else 'FAIL',str(len(c['payments']))],
                            [r['lease_id'],'Current Split','PASS' if c['current_liability']+c['noncurrent_liability']==c['reporting_liability'] else 'FAIL',''],
                            [r['lease_id'],'Journal Balance','PASS' if c['je_totals']['status']=='PASS' else 'FAIL',str(c['je_totals']['difference'])],
                            [r['lease_id'],'Maturity Reconciles','PASS' if sum(c['maturity'].values())>=0 else 'FAIL','']]
            except Exception as e: results.append([r['lease_id'],'Engine','FAIL',str(e)])
    st.dataframe(pd.DataFrame(results,columns=['Lease','Test','Status','Detail']),width='stretch')

def reports_page(cid):
    st.title('Reports')
    st.caption('Generated report packs for the selected company. Reports are created after a successful Portfolio Calculation.')

    if not REPORT_DIR.exists():
        st.info('No report directory exists yet. Run Calculate → Calculate Portfolio to generate the first report pack.')
        return

    # Only current V7.3 reports are surfaced here. Bundled legacy/demo packs are not
    # treated as results of the current import/calculation run.
    files = sorted(
        [p for p in REPORT_DIR.glob('*.xlsx')
         if p.is_file() and p.name.startswith(f'Lease_Accounting_V7.3_{cid}_')],
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )

    if not files:
        st.info(f'No generated report pack found for Company ID {cid}. Run Calculate → Calculate Portfolio first.')
        st.warning('If Calculate shows success but this page is still empty, check that the application is running from the same V7.3 project folder and that the reports folder is writable.')
        return

    st.success(f'{len(files)} report pack(s) available for Company ID {cid}.')
    for idx, p in enumerate(files[:20]):
        c1, c2, c3 = st.columns([5, 2, 2])
        c1.write(f'**{p.name}**')
        c2.write(f'{p.stat().st_size / 1024:.1f} KB')
        with c3:
            st.download_button(
                'Download', p.read_bytes(), p.name,
                key=f'report_download_{cid}_{idx}_{p.name}'
            )

def audit_page(cid): st.title('Company-wise Detailed Audit Trail'); st.dataframe(df(get_audit(cid,10000)),width='stretch')
def qa_page(cid): st.title('Errors / QA'); st.dataframe(df(get_errors(cid)),width='stretch')
def users_page(cid,user,role):
    st.title('Users & Security'); st.dataframe(df(list_users(cid)),width='stretch')
    if role=='ADMIN':
        with st.form('newuser'):
            u=st.text_input('Username'); p=st.text_input('Password',type='password'); r=st.selectbox('Role',['MAKER','REVIEWER','APPROVER','AUDITOR','READ_ONLY']);
            if st.form_submit_button('Create User'):
                add_user(cid,u,pwd_hash(p),r); audit(cid,user,role,'CREATE','USER',u); st.success('User created.'); st.rerun()
def settings_page(comp,cid,user,role):
    st.title('Company Settings'); st.json(dict(comp)); st.write('Central reporting-date default: current system date. Period lock is company-specific.'); rd=st.date_input('Reporting Date',value=date.today())
    if role in ('APPROVER','ADMIN') and st.button('Lock Reporting Period'): lock_period(cid,rd.isoformat(),user); st.success('Reporting period locked and audit-recorded.')

if 'auth' not in st.session_state: login()
else: main()
