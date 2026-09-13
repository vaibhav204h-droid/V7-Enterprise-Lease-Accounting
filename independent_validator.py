"""Independent reconstruction used for validation evidence.

The validator deliberately does not call ``core.engine.calculate``. It rebuilds payment dates,
escalation, initial PV, liability roll-forward and current/non-current classification from raw
lease inputs so the validation control can detect a defect in the production calculation.
"""
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from dateutil.relativedelta import relativedelta
import json
from .engine import D, money, parse_date, effective_rate_for_days, canonical_frequency, add_period, tolerance, escalation_number


def _events(lease):
    raw = lease.get('escalation_schedule_json') or lease.get('escalation_schedule') or []
    if isinstance(raw, str):
        raw = json.loads(raw)
    out=[]
    for i,e in enumerate(raw or [],1):
        if not isinstance(e,dict):
            raise ValueError('Escalation event must be an object.')
        eff=parse_date(e.get('effective_date') or e.get('date'))
        typ=str(e.get('type') or e.get('method') or '').strip().lower()
        seq=int(e.get('event_sequence',e.get('sequence',i)))
        out.append((eff,typ,D(e.get('rate',e.get('value',e.get('amount',0)))),seq,e))
    return sorted(out,key=lambda x:(x[0],x[3]))


def _event_amount(amount, typ, val, raw):
    if typ in {'cpi','cpi linked','cpi-linked','cpi link','cpi index','cpi indexation','wpi','wpi linked','wpi-linked','wpi index','wpi indexation','other index','other index linked','other-index','index','index %','cpi %','wpi %','other index %'}:
        idx=raw.get('index_value',raw.get('current_index',raw.get('new_index'))); base=raw.get('base_index_value',raw.get('base_index',raw.get('index_base')))
        if idx not in (None,'') and base not in (None,''):
            idx,base=D(idx),D(base)
            if idx<=0 or base<=0: raise ValueError('Index values must be positive.')
            return amount*(idx/base)
        return amount*(1+val)
    if typ in {'fixed %','fixed percentage','percentage','fixed rate','fixed-rate','fixed rate %','rate'}: return amount*(1+val)
    if typ in {'fixed amount','amount increase','fixed ₹','fixed amount increase'}: return amount+val
    if typ in {'reset amount','new base','absolute amount'}: return val
    raise ValueError(f'Unsupported escalation event type: {typ}')


def independent_payment(lease,pay_date):
    amount=D(lease.get('base_payment')); evs=_events(lease)
    if evs:
        commencement=parse_date(lease.get('commencement_date'))
        # Match the independent reconstruction's stated accounting rule:
        # Base Payment is already the commencement amount, so pre-commencement
        # escalation history must not be re-applied.
        for eff,typ,val,_,raw in evs:
            if commencement and eff < commencement:
                continue
            if pay_date < eff: break
            amount=_event_amount(amount,typ,val,raw)
        return money(amount)
    typ=str(lease.get('escalation_type') or 'None').strip().lower(); first=parse_date(lease.get('first_escalation_date'))
    if typ in ('none','') or not first or pay_date<first: return money(amount)
    n=escalation_number(pay_date, first, lease.get('escalation_frequency') or 'Annual')
    if typ in ('fixed amount','fixed ₹','fixed amount increase'): amount += D(lease.get('escalation_amount'))*n
    elif typ in ('fixed %','fixed percentage','percentage','fixed rate','fixed-rate','fixed rate %','rate','cpi','cpi linked','cpi-linked','cpi link','cpi index','cpi indexation','wpi','wpi linked','wpi-linked','wpi index','wpi indexation','other index','other index linked','other-index','index'): amount *= (1+D(lease.get('escalation_rate'))) ** n
    elif typ in {'custom schedule','custom-schedule','custom schedule / event schedule','schedule'}: raise ValueError('Custom Schedule requires escalation event rows.')
    else: raise ValueError(f'Unsupported escalation type: {typ}')
    return money(amount)


def _independent_assessed_expiry(lease):
    c=parse_date(lease.get('commencement_date')); contractual=parse_date(lease.get('expiry_date'))
    if not contractual:
        raise ValueError('Contractual expiry is required.')
    end=contractual
    yes=lambda v: str(v or '').strip().lower() in {'yes','y','true','1'}
    if yes(lease.get('extension_reasonably_certain')) and D(lease.get('extension_option_years') or 0)>0:
        months=int((D(lease.get('extension_option_years'))*12).to_integral_value(rounding=ROUND_HALF_UP))
        end=contractual+relativedelta(months=months)
    if yes(lease.get('termination_reasonably_certain')) and lease.get('termination_option_date'):
        td=parse_date(lease.get('termination_option_date'))
        if td and c and c<td<end:
            end=td
    if c and end<=c:
        raise ValueError('Assessed expiry must be after commencement.')
    return end


def payment_dates(lease):
    c=parse_date(lease['commencement_date']); e=_independent_assessed_expiry(lease)
    f=parse_date(lease['first_payment_date']); freq=canonical_frequency(lease.get('payment_frequency','Monthly')); timing=str(lease.get('payment_timing') or 'Arrears').lower()
    if timing=='arrears' and f<c: raise ValueError('Arrears first payment cannot precede commencement.')
    months={'Monthly':1,'Quarterly':3,'Semi-annual':6,'Annual':12,'Biennial':24}.get(freq)
    if not months: raise ValueError(f'Unsupported frequency: {freq}')
    out=[]; n=0
    while True:
        d=f+relativedelta(months=months*n)
        if d>e: break
        if timing=='advance' or d>=c: out.append(d)
        n+=1
    return out


def independent_restoration_pv(lease):
    legacy=D(lease.get('restoration_pv') or 0)
    cost=D(lease.get('restoration_cost_estimate') or 0)
    exp=parse_date(lease.get('restoration_expected_date'))
    c=parse_date(lease.get('commencement_date'))
    rate=D(lease.get('restoration_discount_rate') or 0)
    if cost>0 and exp and c:
        if exp<=c: return Decimal(0)
        return money(cost/((1+rate)**(Decimal((exp-c).days)/Decimal(365))))
    return money(legacy)

def independent_initial_rou(lease, initial_liability):
    c=parse_date(lease.get('commencement_date'))
    if not c: return Decimal(0)
    prepaid=D(lease.get('prepaid') or 0); idc=D(lease.get('idc') or 0)
    incentive=D(lease.get('lease_incentive') or 0)
    deposit=D(lease.get('deposit_paid') or 0); dr=D(lease.get('deposit_discount_rate') or 0); refund=parse_date(lease.get('deposit_refund_date'))
    deposit_fv=Decimal(0)
    if deposit>0 and refund and refund>c:
        deposit_fv=money(deposit/((1+dr)**(Decimal((refund-c).days)/Decimal(365))))
    deposit_adjustment=deposit-deposit_fv if deposit>0 else Decimal(0)
    return money(initial_liability+prepaid+idc+independent_restoration_pv(lease)-incentive+deposit_adjustment)

def validate_lease(lease, reporting_date):
    c=parse_date(lease['commencement_date']); e=parse_date(lease['expiry_date']); annual=D(lease['ib_rate']); timing=str(lease.get('payment_timing') or 'Arrears').lower()
    pdates=payment_dates(lease); rows=[]; opening=Decimal(0)
    for pd in pdates:
        amt=independent_payment(lease,pd); unpaid = pd>c or (pd==c and timing!='advance')
        if unpaid:
            t=Decimal((pd-c).days)/Decimal(365); pv=amt/((1+annual)**t); opening += pv
        else: pv=Decimal(0)
        rows.append((pd,amt,unpaid,pv))
    if reporting_date<c:
        return {'initial_liability':Decimal(0),'initial_rou':Decimal(0),'reporting_liability':Decimal(0),'current_liability':Decimal(0),'noncurrent_liability':Decimal(0),'payments':rows,'assessed_effective_expiry':_independent_assessed_expiry(lease)}
    initial=money(opening); bal=opening; prev=c
    for pd,amt,unpaid,_ in rows:
        if not unpaid or pd>reporting_date: continue
        days=(pd-prev).days; interest=bal*effective_rate_for_days(annual,days); bal=bal+interest-amt; prev=pd
        if abs(bal)<=Decimal('0.10'): bal=Decimal(0)
    if reporting_date>prev and bal>0:
        bal += bal*effective_rate_for_days(annual,(reporting_date-prev).days)
    reporting=money(max(Decimal(0),bal)); horizon=reporting_date+relativedelta(years=1); future=bal; future_prev=reporting_date; current=Decimal(0)
    for pd,amt,unpaid,_ in rows:
        if not unpaid or not(reporting_date<pd<=horizon): continue
        days=(pd-future_prev).days; interest=future*effective_rate_for_days(annual,days); principal=max(Decimal(0),amt-interest); future=max(Decimal(0),future+interest-amt); current += principal; future_prev=pd
    current=money(min(current,bal)); noncurrent=money(max(Decimal(0),reporting-current))
    return {'initial_liability':initial,'initial_rou':independent_initial_rou(lease,initial),'reporting_liability':reporting,'current_liability':current,'noncurrent_liability':noncurrent,'payments':rows,'assessed_effective_expiry':_independent_assessed_expiry(lease)}
