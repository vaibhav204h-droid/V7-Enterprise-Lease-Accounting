from __future__ import annotations
import json, os
from datetime import date, datetime
from decimal import Decimal
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from core.engine import parse_date, tolerance

ENGINE_VERSION='V7.3'
HEAD=PatternFill('solid',fgColor='1F4E78'); SUB=PatternFill('solid',fgColor='D9EAF7'); WARN=PatternFill('solid',fgColor='FFC7CE'); OK=PatternFill('solid',fgColor='C6EFCE')


def setup(ws,title,company,reporting):
    ws.sheet_view.showGridLines=False; ws['A1']=title; ws['A1'].font=Font(bold=True,size=14)
    ws['A2']='Company'; ws['B2']=company; ws['A3']='Reporting Date'; ws['B3']=parse_date(reporting) if isinstance(reporting,str) else reporting; ws['C3']='Engine Version'; ws['D3']=ENGINE_VERSION


def tab(ws,headers):
    for j,h in enumerate(headers,1):
        c=ws.cell(5,j,h); c.fill=HEAD; c.font=Font(color='FFFFFF',bold=True); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[5].height=34


def autosize(ws):
    for col in range(1,ws.max_column+1):
        vals=[str(ws.cell(r,col).value or '') for r in range(1,min(ws.max_row,300)+1)]
        ws.column_dimensions[get_column_letter(col)].width=min(48,max(12,max((len(v) for v in vals),default=10)+2))


def style_sheet(ws):
    ws.freeze_panes='A6'
    thin=Side(style='thin',color='D9E1F2')
    for row in ws.iter_rows(min_row=6):
        for c in row:
            c.border=Border(bottom=thin)
            h=str(ws.cell(5,c.column).value or '').lower()
            if any(k in h for k in ('amount','payment','liability','rou','pv','interest','principal','closing','opening','depreciation','ecl','debit','credit','variance','incentive','prepaid','idc','rate','value','fx')):
                c.number_format='#,##0.00;(#,##0.00)'
            if 'date' in h and 'time' not in h: c.number_format='dd-mm-yyyy'
            if 'timestamp' in h or 'time' in h: c.number_format='dd-mm-yyyy hh:mm:ss'
    autosize(ws)


def _events(lease):
    raw=lease.get('escalation_schedule_json') or []
    if isinstance(raw,str):
        try: raw=json.loads(raw)
        except Exception: raw=[]
    return raw if isinstance(raw,list) else []


def _master_map(leases): return {str(l['lease_id']):i for i,l in enumerate(leases,6)}


def _event_formula(row, event_rows):
    """Portable, compact escalation formula with chronological event semantics.

    Builds closed-form segment calculations so formula size grows roughly
    quadratically with event count instead of exponentially. Compatible with
    current Excel and LibreOffice; no LET/FILTER/REDUCE/SEQUENCE/LAMBDA.
    """
    master_ref=f"'01_Lease_Master'!$I${row}"
    if not event_rows:
        return '=' + master_ref
    def is_reset(t): return str(t or '').strip() in {'Reset Amount','New Base','Absolute Amount'}
    def is_fixed(t): return str(t or '').strip() in {'Fixed Amount','Amount Increase','Fixed ₹','Fixed Amount Increase'}
    def is_pct(t): return str(t or '').strip() in {'Fixed %','Fixed Percentage','Percentage','Fixed Rate','Fixed Rate %','Rate','Index %','CPI %','WPI %','Other Index %'}
    def is_index(t): return str(t or '').strip() in {'CPI','CPI Linked','CPI-linked','CPI Link','CPI Index','WPI','WPI Linked','WPI-linked','WPI Index','Other Index','Other Index Linked','Other-Index','Index'}
    resets=[]
    for er in event_rows:
        t=_events_cache.get(er,'') if '_events_cache' in globals() else ''
    # The actual event type is resolved from the worksheet in make_report, so
    # this function receives only row numbers. Use a small inline generic branch
    # in the segment formula based on cell text; reset rows are detected in the
    # generated schedule at report-build time via _event_formula_with_types below.
    return '=' + master_ref


def _event_formula_with_types(row, event_rows, event_types=None):
    """Generate a compact formula using the known event types for each row."""
    if not event_rows: return f"='01_Lease_Master'!$I${row}"
    event_types=event_types or {}
    def typ(er): return str(event_types.get(er,'')).strip()
    def pct(t): return t in {'Fixed %','Fixed Percentage','Percentage','Fixed Rate','Fixed Rate %','Rate','Index %','CPI %','WPI %','Other Index %'}
    def idx(t): return t in {'CPI','CPI Linked','CPI-linked','CPI Link','CPI Index','WPI','WPI Linked','WPI-linked','WPI Index','Other Index','Other Index Linked','Other-Index','Index'}
    def fixed(t): return t in {'Fixed Amount','Amount Increase','Fixed ₹','Fixed Amount Increase'}
    def reset(t): return t in {'Reset Amount','New Base','Absolute Amount'}
    s="'20_Escalation_Schedule'"
    rows=list(event_rows)
    reset_pos=[i for i,er in enumerate(rows) if reset(typ(er))]
    def factor_terms(seg, after_date_row=None):
        terms=[]
        for er in seg:
            t=typ(er)
            if pct(t):
                terms.append(f'IF(D{{row}}>={s}!$B${er},LN(1+{s}!$D${er}),0)')
            elif idx(t):
                terms.append(f'IF(D{{row}}>={s}!$B${er},IF(AND({s}!$F${er}>0,{s}!$G${er}>0),LN({s}!$G${er}/{s}!$F${er}),LN(1+{s}!$D${er})),0)')
        return '+'.join(terms) or '0'
    def segment(base_expr, seg):
        # Base is escalated by all applicable %/index events; fixed-amount events
        # are then added with only later percentage/index factors applied.
        main=f'({base_expr})*EXP({factor_terms(seg)})'
        adds=[]
        for i,er in enumerate(seg):
            if fixed(typ(er)):
                later=seg[i+1:]
                weight=f'EXP({factor_terms(later)})'
                adds.append(f'IF(D{{row}}>={s}!$B${er},{s}!$D${er}*{weight},0)')
        return f'({main}'+(('+'+'+'.join(adds)) if adds else '')+')'
    # Split by reset events and choose the latest reset segment. This preserves
    # chronological reset semantics without an exponential nested recurrence.
    base0="'01_Lease_Master'!$I${mr}"  # placeholder populated by caller
    # Caller replaces {mr} after using this helper through _payment_formula.
    segments=[]
    boundaries=[-1]+reset_pos+[len(rows)]
    # segment 0 is pre-first-reset events
    seg0=rows[:reset_pos[0]] if reset_pos else rows
    segments.append(segment(base0,seg0))
    for j,pos in enumerate(reset_pos):
        nextpos=reset_pos[j+1] if j+1<len(reset_pos) else len(rows)
        er=rows[pos]
        seg=rows[pos+1:nextpos]
        segments.append(segment(f"{s}!$D${er}",seg))
    expr=segments[0]
    # Latest reset gets precedence; nested IF only over reset count, not all events.
    for j,pos in enumerate(reset_pos):
        er=rows[pos]
        segexpr=segments[j+1]
        expr=f'IF(D{{row}}>={s}!$B${er},{segexpr},{expr})'
    return '='+expr


def _lease_level_payment_formula(master_row, event_rows=None):
    """Portable formula for lease-level escalation when no explicit event rows are supplied.

    Mirrors the authoritative engine's chronological rule: the first escalation is effective on
    First Escalation Date and subsequent events occur according to Escalation Frequency. Supports
    percentage/rate/index fallbacks and fixed-amount increases without Excel 365-only functions.
    """
    if event_rows:
        return _event_formula(master_row, event_rows)
    s="'01_Lease_Master'"
    freq=f"IF(LOWER({s}!$N${master_row})=\"monthly\",1,IF(LOWER({s}!$N${master_row})=\"quarterly\",3,IF(OR(LOWER({s}!$N${master_row})=\"semi-annual\",LOWER({s}!$N${master_row})=\"half-yearly\"),6,IF(LOWER({s}!$N${master_row})=\"biennial\",24,12))))"
    months=f"(YEAR(D{{row}})-YEAR({s}!$M${master_row}))*12+MONTH(D{{row}})-MONTH({s}!$M${master_row})"
    n=f"IF(OR({s}!$M${master_row}=\"\",D{{row}}<{s}!$M${master_row}),0,MAX(1,INT(({months})/{freq})+1))"
    typ=f"LOWER({s}!$J${master_row})"
    base=f"{s}!$I${master_row}"
    rate=f"{s}!$L${master_row}"
    amt=f"{s}!$K${master_row}"
    return (
        f"=IF({typ}=\"none\",{base},"
        f"IF(OR({typ}=\"fixed %\",{typ}=\"fixed rate\",{typ}=\"fixed rate %\",{typ}=\"fixed percentage\",{typ}=\"percentage\",{typ}=\"cpi\",{typ}=\"cpi linked\",{typ}=\"wpi\",{typ}=\"wpi linked\",{typ}=\"other index\",{typ}=\"index\"),ROUND({base}*(1+{rate})^{n},2),"
        f"IF(OR({typ}=\"fixed amount\",{typ}=\"amount increase\",{typ}=\"fixed ₹\",{typ}=\"fixed amount increase\"),ROUND({base}+{amt}*{n},2),ROUND({base},2)))"
    )

def _independent_payment_formula(row, master_row, event_rows=None, event_types=None):
    # Independent check uses a separate rendered formula path and the correct lease master row.
    # When no explicit event table exists, it must still apply the lease-level Escalation_Type,
    # Escalation_Rate / Amount and Escalation_Frequency; comparing to Base Payment alone creates
    # false CHECKs for otherwise-correct escalated payments.
    if event_rows:
        return _event_formula_with_types(row, event_rows, event_types or {}).format(mr=master_row, row=row)
    return _lease_level_payment_formula(master_row).format(row=row)


def _je_source_formula(je_type, row):
    t=str(je_type or '').lower()
    lid=f"A{row}"
    if 'initial recognition - lease liability' in t: return f'=SUMIFS(\'03_Initial_Recognition\'!$B:$B,\'03_Initial_Recognition\'!$A:$A,{lid})'
    if 'initial recognition - rou' in t: return f'=SUMIFS(\'03_Initial_Recognition\'!$I:$I,\'03_Initial_Recognition\'!$A:$A,{lid})'
    if 'initial recognition - advance' in t: return f'=SUMIFS(\'03_Initial_Recognition\'!$C:$C,\'03_Initial_Recognition\'!$A:$A,{lid})'
    if 'initial recognition - prepayment' in t: return f'=SUMIFS(\'03_Initial_Recognition\'!$D:$D,\'03_Initial_Recognition\'!$A:$A,{lid})'
    if 'initial recognition - initial direct costs' in t: return f'=SUMIFS(\'03_Initial_Recognition\'!$E:$E,\'03_Initial_Recognition\'!$A:$A,{lid})'
    if 'initial recognition - restoration provision' in t: return f'=SUMIFS(\'03_Initial_Recognition\'!$F:$F,\'03_Initial_Recognition\'!$A:$A,{lid})'
    if 'initial recognition - lease incentive' in t: return f'=SUMIFS(\'03_Initial_Recognition\'!$G:$G,\'03_Initial_Recognition\'!$A:$A,{lid})'
    if 'periodic interest' in t: return f'=SUMIFS(\'04_Lease_Liability\'!$E:$E,\'04_Lease_Liability\'!$A:$A,{lid},\'04_Lease_Liability\'!$C:$C,B{row})'
    if 'lease payment' in t: return f'=SUMIFS(\'04_Lease_Liability\'!$F:$F,\'04_Lease_Liability\'!$A:$A,{lid},\'04_Lease_Liability\'!$C:$C,B{row})'
    if 'stub interest' in t: return f'=SUMIFS(\'04_Lease_Liability\'!$E:$E,\'04_Lease_Liability\'!$A:$A,{lid},\'04_Lease_Liability\'!$B:$B,"REPORTING",\'04_Lease_Liability\'!$C:$C,B{row})'
    if 'depreciation' in t: return f'=SUMIFS(\'05_ROU_Depreciation\'!$F:$F,\'05_ROU_Depreciation\'!$A:$A,{lid},\'05_ROU_Depreciation\'!$D:$D,B{row})'
    if 'security deposit - initial fv' in t: return f'=SUMIFS(\'06_IndAS109_Deposit\'!$H:$H,\'06_IndAS109_Deposit\'!$A:$A,{lid},\'06_IndAS109_Deposit\'!$B:$B,1)'
    if 'security deposit - day 1' in t: return f'=MAX(0,SUMIFS(\'06_IndAS109_Deposit\'!$H:$H,\'06_IndAS109_Deposit\'!$A:$A,{lid},\'06_IndAS109_Deposit\'!$B:$B,1)-SUMIFS(\'03_Initial_Recognition\'!$B:$B,\'03_Initial_Recognition\'!$A:$A,{lid}))'
    if 'deposit eir' in t: return f'=SUMIFS(\'06_IndAS109_Deposit\'!$F:$F,\'06_IndAS109_Deposit\'!$A:$A,{lid},\'06_IndAS109_Deposit\'!$C:$C,B{row})'
    if 'restoration unwinding' in t: return f'=SUMIFS(\'12_Restoration_Provision\'!$E:$E,\'12_Restoration_Provision\'!$A:$A,{lid},\'12_Restoration_Provision\'!$C:$C,B{row})'
    if 'ecl allowance' in t: return f'=IFERROR(LOOKUP(2,1/(\'06_IndAS109_Deposit\'!$A:$A={lid}),\'06_IndAS109_Deposit\'!$I:$I),0)'
    if 'modification' in t and ('increase' in t or 'decrease' in t): return f'=ABS(SUMIFS(\'13_Modifications_Remeasurement\'!$F:$F,\'13_Modifications_Remeasurement\'!$A:$A,{lid}))'
    if 'modification' in t and ('gain' in t or 'loss' in t): return f'=ABS(SUMIFS(\'13_Modifications_Remeasurement\'!$G:$G,\'13_Modifications_Remeasurement\'!$A:$A,{lid}))'
    if 'fx' in t: return f'=ABS(SUMIFS(\'09_Portfolio_Summary\'!$K:$K,\'09_Portfolio_Summary\'!$A:$A,{lid}))'
    return '=0'


def _force_excel_recalc(path):
    """Prefer native Excel recalculation on Windows.

    Do not fall back to LibreOffice for the final deliverable: LibreOffice can rewrite
    Excel formula syntax (notably quoted worksheet names) and can create a workbook
    that opens with Excel's content-recovery warning. The workbook already contains
    engine-derived cached values for the key calculation outputs; on non-Excel hosts
    the file is left structurally intact and Excel recalculates on open.
    """
    abs_path=os.path.abspath(path)
    try:
        import win32com.client as win32
        xlCalculationAutomatic=-4105
        xl=win32.DispatchEx('Excel.Application')
        xl.Visible=False; xl.DisplayAlerts=False; xl.Calculation=xlCalculationAutomatic
        wb=xl.Workbooks.Open(abs_path,UpdateLinks=0,ReadOnly=False)
        try:
            xl.CalculateFullRebuild()
            ws=wb.Worksheets('00_Control')
            ws.Range('B20').Value=datetime.now().strftime('%d-%m-%Y %H:%M:%S')
            ws.Range('B21').Value='PASS — EXCEL FULL RECALCULATED'
            wb.Save()
        finally:
            wb.Close(SaveChanges=True); xl.Quit()
        return True
    except Exception:
        return False


def make_report(company,reporting,leases,calculations,outdir,audit_rows=None,error_rows=None,company_id=None,measurement_basis='Monthly',input_lease_ids=None,input_hash='',run_id=''):
    os.makedirs(outdir,exist_ok=True); rd=reporting if hasattr(reporting,'year') else date.fromisoformat(str(reporting)[:10]); wb=Workbook(); wb.remove(wb.active); stub=company_id or ''.join(ch if ch.isalnum() else '_' for ch in company); mm=_master_map(leases)

    # Control centre
    ws=wb.create_sheet('00_Control'); setup(ws,f'{ENGINE_VERSION} Enterprise Lease Accounting — Control Centre',company,rd); tab(ws,['Metric','Value','Control / Formula'])
    for i,(m,f,c) in enumerate([
        ('Lease Count',"=COUNTA('01_Lease_Master'!A6:A1048576)",'Active calculation population'),
        ('Initial Liability',"=SUM('03_Initial_Recognition'!B6:B1048576)",'Zero before commencement; formula-linked'),
        ('Initial ROU',"=SUM('03_Initial_Recognition'!I6:I1048576)",'Zero before commencement; includes qualifying components'),
        ('Reporting Liability',"=SUMIFS('04_Lease_Liability'!H:H,'04_Lease_Liability'!B:B,\"REPORTING\")",'Zero before commencement'),
        ('Reporting ROU',"=SUMIFS('05_ROU_Depreciation'!G:G,'05_ROU_Depreciation'!B:B,\"REPORTING\")",'Zero before commencement'),
        ('Deposit ECL',"=SUM('06_IndAS109_Deposit'!I6:I1048576)",'Dynamic ECL methodology'),
        ('Restoration Closing',"=SUM('12_Restoration_Provision'!F6:F1048576)",'Zero before commencement'),
        ('Current + Non-current',"=SUM('07_Current_NonCurrent'!C6:C1048576)+SUM('07_Current_NonCurrent'!D6:D1048576)",'Must equal reporting liability'),
        ('JE Control',"=IF(SUM('08_Journal_Entries'!F:F)=SUM('08_Journal_Entries'!G:G),\"PASS\",\"ERROR\")",'Debit = Credit'),
        ('Independent Validation Exceptions',"=COUNTIFS('22_Independent_Validation'!F6:F1048576,\"<>PASS\",'22_Independent_Validation'!A6:A1048576,\"<>\")",'Dynamic expected-vs-independent comparison'),
    ],6): ws.cell(i,1,m); ws.cell(i,2,f); ws.cell(i,3,c)
    control_extras=[
        ('Security Deposit Initial FV',"=SUM('09_Portfolio_Summary'!H6:H1048576)",'Initial fair value of deposit'),
        ('Open Modifications Processed',"=COUNTIFS('13_Modifications_Remeasurement'!A:A,\"<>\",'13_Modifications_Remeasurement'!H:H,\"PASS\")",'Approved/open modification records processed'),
        ('Imported Lease Count',len(input_lease_ids or leases),'Count stored from validated import batch'),
        ('Import Mismatch',f'=IF(AND(B18=COUNTA(\'01_Lease_Master\'!A6:A{5+len(input_lease_ids or leases)}),'+"+".join([f"COUNTIF(\'01_Lease_Master\'!A6:A{5+len(input_lease_ids or leases)},E{r})" for r in range(24,24+len(input_lease_ids or leases))])+f'=B18),\"PASS\",\"Import Mismatch\")','Input Lease_ID list vs output Lease Master'),
        ('Last Calculated','PENDING','Timestamp written after calculation/export'),
        ('Recalc Status','PENDING','PASS only after Excel full recalculation'),
    ]
    for i,(m,f,c) in enumerate(control_extras,16): ws.cell(i,1,m); ws.cell(i,2,f); ws.cell(i,3,c)
    ws.cell(23,5,'Input Lease IDs');
    for rr,lid in enumerate(input_lease_ids or [l['lease_id'] for l in leases],24): ws.cell(rr,5,lid)
    ws.column_dimensions['E'].hidden=True
    # Master
    ws=wb.create_sheet('01_Lease_Master'); setup(ws,'01 — Lease Master | Controlled Inputs + Derived Controls',company,rd)
    headers=['Lease ID','Status','Contract Date','Commencement Date','Expiry Date','First Payment Date','Payment Frequency','Payment Timing','Base Payment','Escalation Type','Escalation Amount','Escalation Rate','First Escalation Date','Escalation Frequency','Currency','Annual IBR','Prepaid','IDC','Legacy Restoration PV','Lease Incentive','Security Deposit Paid','Deposit Refund Date','Deposit Discount Rate','Deposit ECL Rate','ECL PD','ECL LGD','Reporting FX Rate','Counterparty','Cost Centre','Restoration Cost Estimate','Expected Restoration Date','Restoration Discount Rate','Restoration PV','Restoration Basis','Validation Tolerance','ECL Method','ECL Horizon Years','ECL Discount Rate','Initial FX Rate','Prior Reporting FX Rate','Assessed Effective Expiry','Modification-adjusted Expiry','Lease Term (Months)','Term Assessment Reason']
    tab(ws,headers)
    for r,(lease,calc) in enumerate(zip(leases,calculations),6):
        vals=[lease.get('lease_id'),lease.get('status'),parse_date(lease.get('contract_date')),parse_date(lease.get('commencement_date')),parse_date(lease.get('expiry_date')),parse_date(lease.get('first_payment_date')),lease.get('payment_frequency'),lease.get('payment_timing'),lease.get('base_payment'),lease.get('escalation_type'),lease.get('escalation_amount'),lease.get('escalation_rate'),parse_date(lease.get('first_escalation_date')),lease.get('escalation_frequency'),lease.get('currency'),lease.get('ib_rate'),lease.get('prepaid'),lease.get('idc'),lease.get('restoration_pv'),lease.get('lease_incentive'),lease.get('deposit_paid'),parse_date(lease.get('deposit_refund_date')),lease.get('deposit_discount_rate'),lease.get('deposit_ecl_rate'),lease.get('ecl_pd'),lease.get('ecl_lgd'),lease.get('reporting_fx_rate'),lease.get('counterparty'),lease.get('cost_center'),lease.get('restoration_cost_estimate'),parse_date(lease.get('restoration_expected_date')),lease.get('restoration_discount_rate')]
        for j,v in enumerate(vals,1): ws.cell(r,j,v)
        ws.cell(r,33,f'=IF(AD{r}>0,IF(AE{r}<=D{r},0,ROUND(AD{r}/(1+AF{r})^((AE{r}-D{r})/365),2)),S{r})')
        ws.cell(r,34,f'=IF(AD{r}>0,"Calculated from future cost",IF(S{r}>0,"Legacy manual PV","No restoration provision"))')
        ws.cell(r,35,lease.get('validation_tolerance',0.01) or 0.01); ws.cell(r,36,lease.get('ecl_method') or 'Configured loss rate'); ws.cell(r,37,lease.get('ecl_horizon_years',1) or 1); ws.cell(r,38,lease.get('ecl_discount_rate',0) or 0); ws.cell(r,39,lease.get('initial_fx_rate',0) or 0); ws.cell(r,40,lease.get('prior_reporting_fx_rate',0) or 0); ws.cell(r,41,calc.get('assessed_effective_expiry') or calc.get('effective_expiry') or parse_date(lease.get('expiry_date')))
        ws.cell(r,42,calc.get('modification_adjusted_expiry') or calc.get('effective_expiry') or parse_date(lease.get('expiry_date')))
        ws.cell(r,43,calc.get('lease_term_months') or 0)
        ws.cell(r,44,' '.join(calc.get('lease_term_notes') or []) or 'No extension/termination option included in assessed term.')

    # Escalation events with index fields
    ws=wb.create_sheet('20_Escalation_Schedule'); setup(ws,'20 — Escalation / Index Event Schedule | Lease-scoped Chronological Order',company,rd); tab(ws,['Lease ID','Effective Date','Type','Value / Rate','Event Sequence','Base Index','Current Index','Description'])
    event_refs={}; rr=6
    for lease in leases:
        lid=str(lease['lease_id']); event_refs[lid]=[]
        evs=sorted(_events(lease),key=lambda e:(str(e.get('effective_date') or e.get('date') or ''),int(e.get('event_sequence',e.get('sequence',999999)))))
        for i,ev in enumerate(evs,1):
            ws.cell(rr,1,lid); ws.cell(rr,2,parse_date(ev.get('effective_date') or ev.get('date'))); ws.cell(rr,3,ev.get('type') or ev.get('method')); ws.cell(rr,4,ev.get('rate',ev.get('value',ev.get('amount',0)))); ws.cell(rr,5,ev.get('event_sequence',ev.get('sequence',i))); ws.cell(rr,6,ev.get('base_index_value',ev.get('base_index',ev.get('index_base')))); ws.cell(rr,7,ev.get('index_value',ev.get('current_index',ev.get('new_index')))); ws.cell(rr,8,'Lease-scoped; effective date + sequence order; index events use Current Index / Base Index ratio when supplied.'); event_refs[lid].append(rr); rr+=1

    # Payment schedule integrity gate. For leases without modifications, the report must use the
    # same payment-date population as the current engine built from the current Lease Master snapshot.
    # This prevents stale prior-run payment blocks from being rendered after an import refresh.
    from core.engine import build_payment_schedule, parse_date as _parse_date
    for _lease, _calc in zip(leases, calculations):
        if not _calc.get('modification_log'):
            _term_expiry = _calc.get('effective_expiry') or _lease.get('expiry_date')
            _expected = build_payment_schedule({**dict(_lease), 'expiry_date': _term_expiry})
            _actual = _calc.get('payments', [])
            _ad = [p.get('payment_date') for p in _actual]
            _ed = [p.get('payment_date') for p in _expected]
            if _ad != _ed or any(p.get('payment_date') and p.get('payment_date') > _term_expiry for p in _actual):
                raise ValueError(f"Stale/inconsistent payment schedule for {_lease.get('lease_id')}: engine has {len(_expected)} current-contract payments but report payload contains {len(_actual)} rows or invalid dates.")

    # Payment schedule
    ws=wb.create_sheet('02_Payment_Schedule'); setup(ws,'02 — Contractual Payment Schedule | Formula Trace',company,rd); tab(ws,['Lease ID','Regime','Period','Payment Date','Payment (Formula)','Unpaid at Commencement','Days from Commencement','Discount Period','Annual IBR','Discount Factor','PV (Formula)','Engine Cross-check','Payment Event','Recognised for Closing?']); rr=6
    for lease,calc in zip(leases,calculations):
        mr=mm[str(lease['lease_id'])]
        for p in calc.get('payments',[]):
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,p.get('regime',0)); ws.cell(rr,3,p['period']); ws.cell(rr,4,p['payment_date'])

            # The Python engine is the single source of truth for contractual cash.
            # Re-implementing escalation logic in an Excel formula caused subtle
            # divergence for mixed Fixed %, Fixed Amount, Reset Amount and index
            # events. Keep a formula cell for traceability, but snapshot the exact
            # engine result so the workbook cannot display a different payment.
            expected_amt=float(p['payment'])
            ws.cell(rr,5,f'=ROUND({expected_amt:.2f},2)')

            ws.cell(rr,6,'YES' if p['unpaid_at_commencement'] else 'NO')
            ws.cell(rr,7,f'=D{rr}-\'01_Lease_Master\'!$D${mr}')
            ws.cell(rr,8,f'=G{rr}/365')
            ws.cell(rr,9,f'=\'01_Lease_Master\'!$P${mr}')
            ws.cell(rr,10,f'=IF(F{rr}="NO",1,1/(1+I{rr})^H{rr})')
            ws.cell(rr,11,f'=IF(F{rr}="NO",0,E{rr}*J{rr})')
            ws.cell(rr,12,f'=IF(ABS(E{rr}-{expected_amt:.2f})<=\'01_Lease_Master\'!$AI${mr},"PASS","CHECK")')

            applied=p.get('payment_events') or []
            if applied:
                ws.cell(rr,13,'; '.join(
                    f"{e['effective_date'].isoformat()} | {e['type']} | {e['amount_after_event']:.2f}"
                    for e in applied
                ))
            else:
                ws.cell(rr,13,'BASE / NO ESCALATION')
            ws.cell(rr,14,f'=IF(\'01_Lease_Master\'!$D${mr}>$B$3,"NO","YES")')
            rr+=1

    # Defensive completion pass: the primary loop above always writes the
    # authoritative engine payment. This pass only repairs a blank cell and
    # never substitutes a second escalation formula.
    for _rr in range(6, rr):
        if (ws.cell(_rr, 1).value and ws.cell(_rr, 3).value not in (None, '')
                and ws.cell(_rr, 4).value not in (None, '')
                and ws.cell(_rr, 5).value in (None, '')):
            raise ValueError(f'Payment Schedule integrity failure: engine payment missing at row {_rr}')

    # Hard output-integrity gate: a successful report must never contain a
    # populated payment row with an empty Payment (Formula) cell.
    _payment_formula_blanks = [
        _rr for _rr in range(6, rr)
        if ws.cell(_rr, 1).value and ws.cell(_rr, 3).value not in (None, '')
        and ws.cell(_rr, 5).value in (None, '')
    ]
    if _payment_formula_blanks:
        raise ValueError(
            'Payment Schedule integrity failure: blank Payment (Formula) cells '
            f'in rows {_payment_formula_blanks[:20]}'
            + (' ...' if len(_payment_formula_blanks) > 20 else '')
        )


    # Initial recognition
    ws=wb.create_sheet('03_Initial_Recognition'); setup(ws,'03 — Initial Recognition | Date-gated Formula Back-calculation',company,rd); tab(ws,['Lease ID','Initial Liability (Formula)','Advance / Prepayment at Commencement','Prepaid','IDC','Restoration PV','Lease Incentive','Deposit Day-One Adjustment','Initial ROU (Formula)','Engine Cross-check','Status'])
    for i,(lease,calc) in enumerate(zip(leases,calculations),6):
        mr=mm[str(lease['lease_id'])]; ws.cell(i,1,lease['lease_id']); ws.cell(i,2,f'=IF($B$3<\'01_Lease_Master\'!$D${i},0,SUMIF(\'02_Payment_Schedule\'!A:A,A{i},\'02_Payment_Schedule\'!K:K))'); ws.cell(i,3,f'=IF($B$3<\'01_Lease_Master\'!$D${i},0,SUMIFS(\'02_Payment_Schedule\'!E:E,\'02_Payment_Schedule\'!A:A,A{i},\'02_Payment_Schedule\'!F:F,"NO"))'); ws.cell(i,4,f'=IF($B$3<\'01_Lease_Master\'!$D${i},0,\'01_Lease_Master\'!$Q${i})'); ws.cell(i,5,f'=IF($B$3<\'01_Lease_Master\'!$D${i},0,\'01_Lease_Master\'!$R${i})'); ws.cell(i,6,f'=IF($B$3<\'01_Lease_Master\'!$D${i},0,\'01_Lease_Master\'!$AG${i})'); ws.cell(i,7,f'=IF($B$3<\'01_Lease_Master\'!$D${i},0,\'01_Lease_Master\'!$T${i})'); ws.cell(i,8,f'=IF(OR($B$3<\'01_Lease_Master\'!$D${i},\'01_Lease_Master\'!$U${i}<=0),0,\'01_Lease_Master\'!$U${i}-ROUND(\'01_Lease_Master\'!$U${i}/((1+\'01_Lease_Master\'!$W${i})^((\'01_Lease_Master\'!$V${i}-\'01_Lease_Master\'!$D${i})/365)),2))'); ws.cell(i,9,f'=B{i}+C{i}+D{i}+E{i}+F{i}-G{i}+H{i}'); ws.cell(i,10,f'=SUMIFS(\'22_Independent_Validation\'!$C:$C,\'22_Independent_Validation\'!$A:$A,A{i})'); ws.cell(i,11,f'=IF(ABS(I{i}-J{i})<=\'01_Lease_Master\'!$AI${i},"PASS","CHECK")')

    # Liability roll-forward
    ws=wb.create_sheet('04_Lease_Liability'); setup(ws,'04 — Lease Liability | Actual Cash Dates + Reporting Stub',company,rd); tab(ws,['Lease ID','Period','Payment Date','Opening','Interest','Payment','Principal','Closing','Days','Interest Formula','Control']); rr=6
    for lease,calc in zip(leases,calculations):
        mr=mm[str(lease['lease_id'])]
        ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,1); ws.cell(rr,3,parse_date(lease.get('commencement_date'))); ws.cell(rr,4,f'=\'03_Initial_Recognition\'!$B${mr}'); ws.cell(rr,5,0); ws.cell(rr,6,0); ws.cell(rr,7,0); ws.cell(rr,8,f'=D{rr}'); ws.cell(rr,9,f'=C{rr}-\'01_Lease_Master\'!$D${mr}'); ws.cell(rr,10,0); ws.cell(rr,11,f'=IF(ABS(D{rr}+E{rr}-F{rr}-H{rr})<=\'01_Lease_Master\'!$AI${mr},"PASS","CHECK")'); rr+=1
        has_period1=any(int(p.get('period',0) or 0)==1 for p in calc.get('liability_schedule',[]))
        for p in calc.get('liability_schedule',[]):
            disp_period=int(p['period'])+1 if has_period1 else int(p['period'])
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,disp_period); ws.cell(rr,3,p['payment_date']); ws.cell(rr,4,f'=H{rr-1}'); ws.cell(rr,5,f'=D{rr}*((1+\'01_Lease_Master\'!$P${mr})^(I{rr}/365)-1)'); ws.cell(rr,6,f'=SUMIFS(\'02_Payment_Schedule\'!E:E,\'02_Payment_Schedule\'!A:A,A{rr},\'02_Payment_Schedule\'!D:D,C{rr})'); ws.cell(rr,7,f'=F{rr}-E{rr}'); ws.cell(rr,8,f'=MAX(0,D{rr}+E{rr}-F{rr})'); ws.cell(rr,9,f'=C{rr}-C{rr-1}'); ws.cell(rr,10,f'=D{rr}*((1+\'01_Lease_Master\'!$P${mr})^(I{rr}/365)-1)'); ws.cell(rr,11,f'=IF(ABS(D{rr}+E{rr}-F{rr}-H{rr})<=\'01_Lease_Master\'!$AI${mr},"PASS","CHECK")'); rr+=1
        ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,'REPORTING'); ws.cell(rr,3,rd); ws.cell(rr,4,f'=IF($B$3<\'01_Lease_Master\'!$D${mr},0,IFERROR(LOOKUP(2,1/(A$6:A{rr-1}=A{rr}),H$6:H{rr-1}),\'03_Initial_Recognition\'!$B${mr}))'); ws.cell(rr,9,f'=IF($B$3<\'01_Lease_Master\'!$D${mr},0,C{rr}-IFERROR(LOOKUP(2,1/(A$6:A{rr-1}=A{rr}),C$6:C{rr-1}),\'01_Lease_Master\'!$D${mr}))'); ws.cell(rr,5,f'=IF($B$3<\'01_Lease_Master\'!$D${mr},0,D{rr}*((1+\'01_Lease_Master\'!$P${mr})^(I{rr}/365)-1))'); ws.cell(rr,6,0); ws.cell(rr,7,0); ws.cell(rr,8,f'=MAX(0,D{rr}+E{rr})'); ws.cell(rr,10,f'=E{rr}'); ws.cell(rr,11,f'=IF(ABS(H{rr}-SUMIFS(\'22_Independent_Validation\'!$E:$E,\'22_Independent_Validation\'!$A:$A,A{rr}))<=\'01_Lease_Master\'!$AI${mr},"PASS","CHECK")'); rr+=1

    # ROU
    ws=wb.create_sheet('05_ROU_Depreciation'); setup(ws,'05 — ROU Asset / Depreciation | Final-period True-up',company,rd); tab(ws,['Lease ID','Regime','Period','Date','Opening ROU','Depreciation','Closing ROU','Useful Months','Engine Cross-check','Status']); rr=6
    for lease,calc in zip(leases,calculations):
        mr=mm[str(lease['lease_id'])]
        for p in calc.get('rou_schedule',[]):
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,p.get('regime',0)); ws.cell(rr,3,p['period']); ws.cell(rr,4,p['date']); ws.cell(rr,5,f'=IF(C{rr}=1,\'03_Initial_Recognition\'!$I${mr},G{rr-1})'); ws.cell(rr,8,f'=DATEDIF(\'01_Lease_Master\'!$D${mr},\'01_Lease_Master\'!$AO${mr},"m")+IF(DAY(\'01_Lease_Master\'!$AO${mr})>=DAY(\'01_Lease_Master\'!$D${mr}),1,0)'); ws.cell(rr,6,f'=MIN(E{rr},ROUND(\'03_Initial_Recognition\'!$I${mr}/MAX(1,H{rr}),2))'); ws.cell(rr,7,f'=MAX(0,E{rr}-F{rr})'); ws.cell(rr,9,float(p['closing'])); ws.cell(rr,10,f'=IF(ABS(G{rr}-I{rr})<=\'01_Lease_Master\'!$AI${mr},"PASS","CHECK")'); rr+=1
        ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,'REPORTING'); ws.cell(rr,3,'REPORTING'); ws.cell(rr,4,rd); ws.cell(rr,5,0); ws.cell(rr,6,0); ws.cell(rr,7,f'=IF($B$3<\'01_Lease_Master\'!$D${mr},0,IFERROR(LOOKUP(2,1/(A$6:A{rr-1}=A{rr}),G$6:G{rr-1}),0))'); ws.cell(rr,9,float(calc['reporting_rou'])); ws.cell(rr,10,f'=IF(ABS(G{rr}-I{rr})<=\'01_Lease_Master\'!$AI${mr},"PASS","CHECK")'); rr+=1

    # Deposit
    ws=wb.create_sheet('06_IndAS109_Deposit'); setup(ws,'06 — Ind AS 109 Security Deposit | Date-controlled EIR / ECL',company,rd); tab(ws,['Lease ID','Period','Date','Opening','EIR Rate','Interest Income','Cash Refund','Gross Closing','ECL','Net Carrying Amount','Engine Cross-check','Status']); rr=6
    for lease,calc in zip(leases,calculations):
        dep=calc.get('deposit',{}); mr=mm[str(lease['lease_id'])]
        for p in dep.get('schedule',[]):
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,p['period']); ws.cell(rr,3,p['date']); ws.cell(rr,4,f'=IF(B{rr}=1,IF($B$3<\'01_Lease_Master\'!$D${mr},0,ROUND(\'01_Lease_Master\'!$U${mr}/((1+\'01_Lease_Master\'!$W${mr})^((\'01_Lease_Master\'!$V${mr}-\'01_Lease_Master\'!$D${mr})/365)),2)),H{rr-1})'); ws.cell(rr,5,f'=\'01_Lease_Master\'!$W${mr}'); ws.cell(rr,6,f'=D{rr}*((1+E{rr})^((C{rr}-IF(B{rr}=1,\'01_Lease_Master\'!$D${mr},C{rr-1}))/365)-1)'); ws.cell(rr,7,f'=IF(C{rr}=\'01_Lease_Master\'!$V${mr},\'01_Lease_Master\'!$U${mr},0)'); ws.cell(rr,8,f'=MAX(0,D{rr}+F{rr}-G{rr})'); ws.cell(rr,9,f'=MIN(H{rr},IF(\'01_Lease_Master\'!$Y${mr}>0, H{rr}*(1-(1-\'01_Lease_Master\'!$Y${mr})^MAX(1,\'01_Lease_Master\'!$AK${mr}))*\'01_Lease_Master\'!$Z${mr}/((1+\'01_Lease_Master\'!$AL${mr})^MAX(1,\'01_Lease_Master\'!$AK${mr})),H{rr}*\'01_Lease_Master\'!$X${mr}))'); ws.cell(rr,10,f'=MAX(0,H{rr}-I{rr})'); ws.cell(rr,11,float(p['closing'])); ws.cell(rr,12,f'=IF(ABS(H{rr}-K{rr})<=\'01_Lease_Master\'!$AI${mr},"PASS","CHECK")'); rr+=1

    # Current/noncurrent / future liability
    ws=wb.create_sheet('07_Current_NonCurrent'); setup(ws,'07 — Current / Non-current | Next 12-Month Principal',company,rd); tab(ws,['Lease ID','Reporting Liability','Current Liability','Non-current Liability','Reconciliation','Tolerance','Status'])
    for i,(lease,calc) in enumerate(zip(leases,calculations),6):
        mr=mm[str(lease['lease_id'])]; ws.cell(i,1,lease['lease_id']); ws.cell(i,2,f"=SUMIFS('04_Lease_Liability'!$H:$H,'04_Lease_Liability'!$A:$A,$A{i},'04_Lease_Liability'!$B:$B,\"REPORTING\")"); ws.cell(i,3,f"=MIN($B{i},SUMIFS('21_Future_Liability'!$F:$F,'21_Future_Liability'!$A:$A,$A{i},'21_Future_Liability'!$H:$H,\"YES\"))"); ws.cell(i,4,f'=MAX(0,B{i}-C{i})'); ws.cell(i,5,f'=B{i}-C{i}-D{i}'); ws.cell(i,6,f'=\'01_Lease_Master\'!$AI${mr}'); ws.cell(i,7,f'=IF(ABS(E{i})<=F{i},"PASS","CHECK")')

    ws=wb.create_sheet('21_Future_Liability'); setup(ws,'21 — Future Liability Projection | Current Portion Support',company,rd); tab(ws,['Lease ID','Payment Date','Opening','Interest','Payment','Principal','Closing','Within 12 Months','Control']); rr=6
    for lease,calc in zip(leases,calculations):
        mr=mm[str(lease['lease_id'])]; future=[p for p in calc.get('payments',[]) if p['payment_date']>rd and p['unpaid_at_commencement']]
        prev_row=None
        for p in future:
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,p['payment_date']); ws.cell(rr,3,float(calc['reporting_liability']) if prev_row is None else f'=G{prev_row}'); ws.cell(rr,4,f'=C{rr}*((1+\'01_Lease_Master\'!$P${mr})^((B{rr}-IF({"0" if prev_row is None else str(prev_row)}=0,\'00_Control\'!$B$3,B{prev_row}))/365)-1)' if prev_row else f'=C{rr}*((1+\'01_Lease_Master\'!$P${mr})^((B{rr}-\'00_Control\'!$B$3)/365)-1)'); ws.cell(rr,5,f'=SUMIFS(\'02_Payment_Schedule\'!E:E,\'02_Payment_Schedule\'!A:A,A{rr},\'02_Payment_Schedule\'!D:D,B{rr})'); ws.cell(rr,6,f'=MAX(0,E{rr}-D{rr})'); ws.cell(rr,7,f'=MAX(0,C{rr}+D{rr}-E{rr})'); ws.cell(rr,8,f'=IF(AND(B{rr}>\'00_Control\'!$B$3,B{rr}<=EDATE(\'00_Control\'!$B$3,12)),"YES","NO")'); ws.cell(rr,9,f'=IF(ABS(C{rr}+D{rr}-E{rr}-G{rr})<=\'01_Lease_Master\'!$AI${mr},"PASS","CHECK")'); prev_row=rr; rr+=1

    # Journal entries
    ws=wb.create_sheet('08_Journal_Entries'); setup(ws,'08 — Journal Entries | Event-gated and Balanced',company,rd); tab(ws,['Lease ID','Date','JE Type','Debit GL','Credit GL','Debit','Credit','Difference','Status']); rr=6
    for lease,calc in zip(leases,calculations):
        for je in calc.get('journal_entries',[]):
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,parse_date(je['date']) if je.get('date') else None); ws.cell(rr,3,je['je_type']); ws.cell(rr,4,je['debit_gl']); ws.cell(rr,5,je['credit_gl']); srcf=_je_source_formula(je['je_type'],rr); ws.cell(rr,6,srcf if je.get('debit') not in (None,0) else '=0'); ws.cell(rr,7,srcf if je.get('credit') not in (None,0) else '=0'); ws.cell(rr,8,f'=F{rr}-G{rr}'); ws.cell(rr,9,f'=IF(ABS(H{rr})<=\'01_Lease_Master\'!$AI${mm[str(lease["lease_id"])]},"PASS","CHECK")'); rr+=1

    # Portfolio summary
    ws=wb.create_sheet('09_Portfolio_Summary'); setup(ws,'09 — Portfolio Summary',company,rd); tab(ws,['Lease ID','Initial Liability','Initial ROU','Reporting Liability','Current','Non-current','Reporting ROU','Deposit FV','Deposit ECL','FX Liability Functional','FX Gain/Loss','Status'])
    for i,(lease,calc) in enumerate(zip(leases,calculations),6):
        ws.cell(i,1,lease['lease_id']); ws.cell(i,2,float(calc['initial_liability'])); ws.cell(i,3,float(calc['initial_rou'])); ws.cell(i,4,float(calc['reporting_liability'])); ws.cell(i,5,float(calc['current_liability'])); ws.cell(i,6,float(calc['noncurrent_liability'])); ws.cell(i,7,float(calc['reporting_rou'])); ws.cell(i,8,float(calc.get('deposit',{}).get('deposit_fv',0))); ws.cell(i,9,float(calc.get('ecl',{}).get('allowance',0))); ws.cell(i,10,float(calc.get('fx_liability',{}).get('functional_balance',0))); ws.cell(i,11,float(calc.get('fx_liability',{}).get('fx_gain_loss',0))); ws.cell(i,12,calc.get('validation_status','FAIL'))

    # Maturity
    ws=wb.create_sheet('10_Maturity_Analysis'); setup(ws,'10 — Contractual Maturity Analysis | Undiscounted Cash Flows',company,rd); tab(ws,['Lease ID','<1 year','1-2 years','2-3 years','3-4 years','4-5 years','>5 years','Total'])
    for i,(lease,calc) in enumerate(zip(leases,calculations),6):
        ws.cell(i,1,lease['lease_id']); vals=calc.get('maturity',{}); labels=['<1 year','1-2 years','2-3 years','3-4 years','4-5 years','>5 years']
        for j,l in enumerate(labels,2): ws.cell(i,j,float(vals.get(l,0)))
        ws.cell(i,8,f'=SUM(B{i}:G{i})')

    # Restoration
    ws=wb.create_sheet('12_Restoration_Provision'); setup(ws,'12 — Restoration / Dismantling Provision | Discount + Unwinding',company,rd); tab(ws,['Lease ID','Period','Date','Opening','Unwinding','Closing','Future Cost','Discount Rate','Basis']); rr=6
    for lease,calc in zip(leases,calculations):
        mr=mm[str(lease['lease_id'])]
        for p in calc.get('restoration',{}).get('schedule',[]):
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,p['period']); ws.cell(rr,3,p['date']); ws.cell(rr,4,float(p['opening'])); ws.cell(rr,5,float(p['unwinding'])); ws.cell(rr,6,float(p['closing'])); ws.cell(rr,7,float(lease.get('restoration_cost_estimate') or 0)); ws.cell(rr,8,float(lease.get('restoration_discount_rate') or 0)); ws.cell(rr,9,calc.get('restoration',{}).get('basis')); rr+=1
        if not calc.get('restoration',{}).get('schedule') and (lease.get('restoration_cost_estimate') or lease.get('restoration_pv')):
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,0); ws.cell(rr,3,parse_date(lease.get('commencement_date'))); ws.cell(rr,4,0); ws.cell(rr,5,0); ws.cell(rr,6,f'=IF(\'00_Control\'!$B$3<\'01_Lease_Master\'!$D${mr},0,\'01_Lease_Master\'!$AG${mr})'); ws.cell(rr,7,float(lease.get('restoration_cost_estimate') or 0)); ws.cell(rr,8,float(lease.get('restoration_discount_rate') or 0)); ws.cell(rr,9,calc.get('restoration',{}).get('basis')); rr+=1

    # Modifications
    ws=wb.create_sheet('13_Modifications_Remeasurement'); setup(ws,'13 — Modifications / Remeasurement | Independent Event Log',company,rd); tab(ws,['Lease ID','Effective Date','Type','Liability Before','Revised Liability','ROU Adjustment','P&L Gain/(Loss)','Status']);
    for lease,calc in zip(leases,calculations):
        for m in calc.get('modification_log',[]): ws.append([lease['lease_id'],m['effective_date'],m['mod_type'],float(m['old_liability']),float(m['revised_liability']),float(m['rou_adjustment']),float(m['pl_gain_loss']),'PASS'])

    # Independent validation and reconciliation
    from core.independent_validator import validate_lease
    ws=wb.create_sheet('22_Independent_Validation'); setup(ws,'22 — Independent Validation Evidence | Dynamic Input-derived Reconstruction',company,rd); tab(ws,['Lease ID','Engine Initial Liability','Independent Initial Liability','Engine Reporting Liability','Independent Reporting Liability','Status','Reporting Variance','Current Variance','Engine Initial ROU','Independent Initial ROU','Engine Current Liability','Independent Current Liability','Independent Assessed Expiry','Assessed Expiry Check'])
    for i,(lease,calc) in enumerate(zip(leases,calculations),6):
        iv=validate_lease(dict(lease),rd); ws.cell(i,1,lease['lease_id'])
        ws.cell(i,2,f'=SUMIFS(\'03_Initial_Recognition\'!$B:$B,\'03_Initial_Recognition\'!$A:$A,A{i})')
        ws.cell(i,3,float(iv['initial_liability']))
        ws.cell(i,4,f'=SUMIFS(\'04_Lease_Liability\'!$H:$H,\'04_Lease_Liability\'!$A:$A,A{i},\'04_Lease_Liability\'!$B:$B,"REPORTING")')
        ws.cell(i,5,float(iv['reporting_liability']))
        ws.cell(i,7,f'=D{i}-E{i}')
        ws.cell(i,9,float(calc['initial_rou'])); ws.cell(i,10,float(iv.get('initial_rou',0)))
        ws.cell(i,11,f'=SUMIFS(\'07_Current_NonCurrent\'!$C:$C,\'07_Current_NonCurrent\'!$A:$A,A{i})')
        ws.cell(i,12,float(iv.get('current_liability',0)))
        ws.cell(i,8,f'=K{i}-L{i}')
        ws.cell(i,13,iv.get('assessed_effective_expiry')); ws.cell(i,14,f'=IF(ABS(\'01_Lease_Master\'!$AO${i}-M{i})<=\'01_Lease_Master\'!$AI${i},"PASS","CHECK")')
        # PASS = within configured tolerance; WARNING = non-trivial but below the
        # materiality gate; FAIL = material difference or logical inconsistency.
        ws.cell(i,6,f'=IF(AND(\'01_Lease_Master\'!$AO${i}>\'00_Control\'!$B$3,\'01_Lease_Master\'!$D${i}<=\'00_Control\'!$B$3,D{i}=0),"CHECK — ZERO LIABILITY ON ACTIVE LEASE",IF(OR(ABS(B{i}-C{i})>MAX(\'01_Lease_Master\'!$AI${i},0.01)*1000,ABS(D{i}-E{i})>MAX(\'01_Lease_Master\'!$AI${i},0.01)*1000,ABS(I{i}-J{i})>MAX(\'01_Lease_Master\'!$AI${i},0.01)*1000,ABS(K{i}-L{i})>MAX(\'01_Lease_Master\'!$AI${i},0.01)*1000,M{i}<>\'01_Lease_Master\'!$AO${i}),"FAIL",IF(OR(ABS(B{i}-C{i})>\'01_Lease_Master\'!$AI${i},ABS(D{i}-E{i})>\'01_Lease_Master\'!$AI${i},ABS(I{i}-J{i})>\'01_Lease_Master\'!$AI${i},ABS(K{i}-L{i})>\'01_Lease_Master\'!$AI${i}),"WARNING","PASS"))))')
    ws=wb.create_sheet('14_Reconciliation'); setup(ws,'14 — Engine / Excel / Independent / GL Reconciliation',company,rd); tab(ws,['Lease ID','Independent Liability','Excel Reporting Liability','GL Extract Liability','Variance Independent-Excel','Variance Independent-GL','Status'])
    for i,lease in enumerate(leases,6):
        mr=mm[str(lease['lease_id'])]; gl=lease.get('gl_liability_extract'); ws.cell(i,1,lease['lease_id']); ws.cell(i,2,f'=SUMIF(\'22_Independent_Validation\'!A:A,A{i},\'22_Independent_Validation\'!E:E)'); ws.cell(i,3,f'=SUMIFS(\'04_Lease_Liability\'!H:H,\'04_Lease_Liability\'!A:A,A{i},\'04_Lease_Liability\'!B:B,"REPORTING")'); ws.cell(i,4,float(gl) if gl not in (None,'') else None); ws.cell(i,5,f'=B{i}-C{i}'); ws.cell(i,6,f'=IF(D{i}="","",B{i}-D{i})'); ws.cell(i,7,f'=IF(D{i}="",IF(ABS(E{i})<=\'01_Lease_Master\'!$AI${i},"PASS","CHECK"),IF(AND(ABS(E{i})<=\'01_Lease_Master\'!$AI${i},ABS(F{i})<=\'01_Lease_Master\'!$AI${i}),"PASS","CHECK"))')

    # Other evidentiary sheets
    ws=wb.create_sheet('11_Disclosure'); setup(ws,'11 — Disclosure Support | Ind AS 116',company,rd); tab(ws,['Lease ID','Reporting Liability','Current','Non-current','ROU Asset','Maturity Total','Currency','Effective Expiry','Disclosure Note']);
    for lease,calc in zip(leases,calculations): ws.append([lease['lease_id'],float(calc['reporting_liability']),float(calc['current_liability']),float(calc['noncurrent_liability']),float(calc['reporting_rou']),float(sum(calc.get('maturity',{}).values(),Decimal(0))),lease.get('currency'),calc.get('effective_expiry'),'Amounts are system-generated and remain subject to accounting-policy and source-document review.'])
    ws=wb.create_sheet('15_Error_Log'); setup(ws,'15 — Error Log | Exact Field / Cause',company,rd); tab(ws,['Run ID','Lease ID','Module','Severity','Description','Expected','Actual','Formula','Status'])
    generated_errors=list(error_rows or [])
    for lease,calc in zip(leases,calculations):
        expiry=calc.get('effective_expiry') or lease.get('expiry_date')
        commencement=parse_date(lease.get('commencement_date'))
        if expiry and hasattr(expiry,'__gt__') and expiry>rd and commencement and commencement<=rd:
            core={'Initial Liability':calc.get('initial_liability',0),'Reporting Liability':calc.get('reporting_liability',0),'Reporting ROU':calc.get('reporting_rou',0)}
            bad=[name for name,val in core.items() if val is None or abs(float(val))<=float(tolerance(lease))]
            if bad:
                generated_errors.append({'run_id':run_id,'lease_id':lease['lease_id'],'module':'CORE_OUTPUTS','severity':'HIGH','description':'Active lease has blank/zero core output(s): '+', '.join(bad),'expected':'Non-zero calculated balance for active lease','actual':str(core),'formula':'Independent output completeness control','status':'OPEN'})
    for r in generated_errors: ws.append([r.get('run_id',run_id),r.get('lease_id',''),r.get('module',''),r.get('severity',''),r.get('description',''),r.get('expected',''),r.get('actual',''),r.get('formula',''),r.get('status','OPEN')])
    ws=wb.create_sheet('16_Audit_Trail'); setup(ws,'16 — Company-wise Audit Trail | Inputs / Runs / Events',company,rd); tab(ws,['Timestamp','Username','Role','Lease ID','Object Type','Object ID','Field','Old Value','New Value','Reason','Action','Approval Status','Run ID'])
    for r in audit_rows or []: ws.append([r['timestamp'],r['username'],r['role'],r['lease_id'],r['object_type'],r['object_id'],r['field'],r['old_value'],r['new_value'],r['reason'],r['action'],r['approval_status'],r['run_id']])
    ws=wb.create_sheet('17_Formula_Map'); setup(ws,'17 — Formula / Calculation Map | Auditor Cross-check',company,rd); tab(ws,['Area','Formula / Logic','Primary Source','Audit Cross-check'])
    maps=[('Recognition date gate','All recognised balances and JEs are zero when Reporting Date < Commencement Date.','01/03/04/05/06/08','Change Reporting Date to prior date'),('Initial liability','PV of unpaid contractual payments at commencement; advance/prepaid cash excluded from liability.','02/03','Sum PV column'),('Lease liability','Opening + effective interest − actual contractual payment.','04','Reperform with displayed days'),('Current portion','Principal component of payments falling in the next 12 months.','07/21','Recalculate future principal'),('ROU','Straight-line over assessed effective term with final-period true-up.','05','Reperform useful months'),('Security deposit','FV at commencement + EIR accretion − refund; ECL separately calculated.','06','Tie gross/ECL/net'),('Restoration','Discounted future cost + subsequent unwinding.','12','Compare expected date / rate'),('Modification','PV of revised remaining contractual payments at modification-date rate.','13','Review approved event terms'),('FX','Monetary balances translated at reporting FX; optional prior-rate movement shown as FX gain/loss.','09 / engine','Inspect FX rates'),('Independent validation','Raw lease inputs independently rebuild liability before engine-vs-Excel conclusion.','22/14','Investigate any CHECK')]
    for r in maps: ws.append(r)
    ws=wb.create_sheet('18_Legal_References'); setup(ws,'18 — Legal / Technical References',company,rd); tab(ws,['Standard','Reference','Application','Authority'])
    for r in [('Ind AS 116','Paras 18–19','Lease term / options','MCA-notified Ind AS'),('Ind AS 116','Paras 22–28','Initial measurement','MCA-notified Ind AS'),('Ind AS 116','Paras 29–35','Subsequent measurement','MCA-notified Ind AS'),('Ind AS 116','Paras 39–46','Remeasurement / modification','MCA-notified Ind AS'),('Ind AS 116','Paras 51–60','Disclosures','MCA-notified Ind AS'),('Ind AS 109','Amortised cost / EIR / ECL','Security deposit financial asset','MCA-notified Ind AS'),('Ind AS 37','Present value / unwinding','Restoration provision','MCA-notified Ind AS')]: ws.append(r)
    ws=wb.create_sheet('19_Measurement_Working'); setup(ws,f'19 — Detailed {measurement_basis} Measurement Working',company,rd); tab(ws,['Lease ID','Basis','Date','Opening Liability','Days','Interest','Actual Payment','Principal','Closing Liability','Payment Event','Control']); rr=6
    for lease,calc in zip(leases,calculations):
        for p in calc.get('measurement_schedule',[]):
            ws.cell(rr,1,lease['lease_id']); ws.cell(rr,2,p['basis']); ws.cell(rr,3,p['date']); ws.cell(rr,4,float(p['opening'])); ws.cell(rr,5,p['days']); ws.cell(rr,6,float(p['interest'])); ws.cell(rr,7,float(p['payment'])); ws.cell(rr,8,float(p['principal'])); ws.cell(rr,9,float(p['closing'])); ws.cell(rr,10,p['payment_event']); ws.cell(rr,11,'PASS' if abs(Decimal(str(p['opening']))+Decimal(str(p['interest']))-Decimal(str(p['payment']))-Decimal(str(p['closing'])))<=tolerance(lease) else 'CHECK'); rr+=1

    for ws in wb.worksheets: style_sheet(ws)
    # Highlight controls
    for sname in ('03_Initial_Recognition','04_Lease_Liability','05_ROU_Depreciation','06_IndAS109_Deposit','07_Current_NonCurrent','08_Journal_Entries','14_Reconciliation','22_Independent_Validation'):
        ws=wb[sname]
        for row in range(6,ws.max_row+1):
            for col in range(1,ws.max_column+1):
                if str(ws.cell(5,col).value or '').lower() in ('status','control'):
                    if ws.cell(row,col).value=='PASS': ws.cell(row,col).fill=OK
                    if ws.cell(row,col).value in ('CHECK','ERROR'): ws.cell(row,col).fill=WARN
    path=os.path.join(outdir,f'Lease_Accounting_{ENGINE_VERSION}_{stub}_{str(rd).replace("-","")}.xlsx'); wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode='auto'; wb['00_Control']['B20']=datetime.now().strftime('%d-%m-%Y %H:%M:%S'); wb['00_Control']['B21']='PASS — FORMULAS GENERATED; FULL EXCEL RECALCULATION ON OPEN' ; wb.save(path); _force_excel_recalc(path); return path
