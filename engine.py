from __future__ import annotations

import json
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation, getcontext
from dateutil.relativedelta import relativedelta
from typing import Any, Dict, List, Tuple

getcontext().prec = 34
MONEY = Decimal('0.01')
DEFAULT_TOLERANCE = Decimal('0.01')
TOL = DEFAULT_TOLERANCE
DEFAULT_REPORTING_DATE = date.today()
FREQ_MONTHS = {'Monthly': 1, 'Quarterly': 3, 'Semi-annual': 6, 'Half-Yearly': 6, 'Annual': 12, 'Biennial': 24}
FREQ_ALIASES = {
    'monthly':'Monthly','month':'Monthly','1 month':'Monthly','1-month':'Monthly',
    'quarterly':'Quarterly','quarter':'Quarterly','3 monthly':'Quarterly','3-monthly':'Quarterly',
    'semi-annual':'Semi-annual','semi annual':'Semi-annual','semiannual':'Semi-annual',
    'half-yearly':'Semi-annual','half yearly':'Semi-annual','half-year':'Semi-annual',
    'half year':'Semi-annual','6 monthly':'Semi-annual','6-monthly':'Semi-annual',
    'annual':'Annual','yearly':'Annual','year':'Annual','12 monthly':'Annual',
    'biennial':'Biennial','bi-annual':'Biennial','biannual':'Biennial','2 yearly':'Biennial','24 monthly':'Biennial',
}

INDEX_TYPES = {
    'cpi', 'cpi linked', 'cpi-linked', 'cpi link', 'cpi index', 'cpi indexation',
    'wpi', 'wpi linked', 'wpi-linked', 'wpi index', 'wpi indexation',
    'other index', 'other index linked', 'other-index', 'index', 'index %',
    'cpi %', 'wpi %', 'other index %',
}
PERCENT_ESCALATIONS = {
    'fixed %', 'fixed percentage', 'percentage', 'fixed rate', 'fixed-rate',
    'fixed rate %', 'rate', 'index %', 'cpi %', 'wpi %', 'other index %',
    'cpi', 'wpi', 'other index', 'index',
}
FIXED_AMOUNT_ESCALATIONS = {'fixed amount','amount increase','fixed ₹','fixed amount increase'}
RESET_ESCALATIONS = {'reset amount','new base','absolute amount'}
CUSTOM_SCHEDULE_TYPES = {'custom schedule','custom-schedule','custom schedule / event schedule','schedule'}


class CalculationValidationError(ValueError):
    pass


def canonical_frequency(v: Any) -> str:
    s = str(v or 'Monthly').strip().lower().replace('_', ' ').replace('/', ' ')
    return FREQ_ALIASES.get(s, str(v or 'Monthly').strip())


def D(v: Any) -> Decimal:
    if v is None or v == '':
        return Decimal('0')
    try:
        return Decimal(str(v))
    except (InvalidOperation, ValueError) as e:
        raise CalculationValidationError(f'Invalid numeric value: {v}') from e


def money(v: Decimal) -> Decimal:
    return D(v).quantize(MONEY, rounding=ROUND_HALF_UP)


def tolerance(lease: Dict[str, Any] | None = None) -> Decimal:
    lease = lease or {}
    raw = lease.get('validation_tolerance', lease.get('rounding_tolerance', DEFAULT_TOLERANCE))
    t = D(raw)
    if t <= 0:
        raise CalculationValidationError('Validation tolerance must be greater than zero.')
    return t.quantize(MONEY, rounding=ROUND_HALF_UP)


def parse_date(v: Any) -> date | None:
    if v in (None, ''):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if 'T' in s:
        s = s.split('T', 1)[0]
    s = s[:10]
    for fmt in ('%Y-%m-%d','%d-%m-%Y','%d/%m/%Y','%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise CalculationValidationError(f'Invalid date: {v}')


def annual_to_periodic(annual_rate: Any, frequency: str) -> Decimal:
    r = D(annual_rate)
    if r < 0:
        raise CalculationValidationError('IBR cannot be negative.')
    freq=canonical_frequency(frequency)
    if freq == 'Biennial':
        n = Decimal('0.5')  # one payment every two years
    else:
        n = Decimal(str({'Monthly':12,'Quarterly':4,'Semi-annual':2,'Half-Yearly':2,'Annual':1}.get(freq, 12)))
    return (Decimal(1) + r) ** (Decimal(1) / n) - Decimal(1)


def effective_rate_for_days(annual_rate: Any, days: int) -> Decimal:
    r = D(annual_rate)
    if days < 0:
        raise CalculationValidationError('Elapsed days cannot be negative.')
    if days == 0:
        return Decimal(0)
    if r <= Decimal('-1'):
        raise CalculationValidationError('Annual effective rate cannot be less than or equal to -100%.')
    return (Decimal(1) + r) ** (Decimal(days) / Decimal(365)) - Decimal(1)


def add_period(d: date, frequency: str) -> date:
    return d + relativedelta(months=FREQ_MONTHS.get(canonical_frequency(frequency), 1))


def generate_payment_dates(commencement: date, expiry: date, first_payment: date, frequency: str, payment_timing: str = 'Arrears') -> List[date]:
    """Generate contractual dates from the original payment-date anchor.

    IMPORTANT: do not repeatedly add a period to the previously generated date.
    Repeated relativedelta(months=1) causes 31-Jan -> 28-Feb -> 28-Mar drift.
    Each date is calculated from the original first-payment anchor:
        anchor + n * frequency
    with month-end clamping performed by relativedelta.
    """
    if not (commencement and expiry and first_payment):
        return []
    if expiry <= commencement:
        raise CalculationValidationError('Expiry date must be after commencement date.')
    freq = canonical_frequency(frequency)
    timing = str(payment_timing or 'Arrears').strip().lower()
    if timing == 'arrears' and first_payment < commencement:
        raise CalculationValidationError('Arrears first payment cannot precede commencement date.')
    if timing == 'advance' and first_payment < commencement:
        pass

    months = FREQ_MONTHS.get(freq)
    if not months:
        raise CalculationValidationError(f'Unsupported payment frequency: {frequency}')

    out: List[date] = []
    n = 0
    for _ in range(10000):
        d = first_payment + relativedelta(months=months * n)
        if d > expiry:
            break
        if timing == 'advance' or d >= commencement:
            out.append(d)
        n += 1
    if not out:
        raise CalculationValidationError('No payment dates generated; check first payment date/frequency.')
    return out


def calculate_fixed_amount_escalation(base_payment: Any, escalation_amount: Any, escalation_number: int) -> Decimal:
    return money(D(base_payment) + D(escalation_amount) * Decimal(escalation_number))


def calculate_percentage_escalation(base_payment: Any, escalation_percentage: Any, escalation_number: int) -> Decimal:
    r = D(escalation_percentage)
    if r <= Decimal('-1'):
        raise CalculationValidationError('Percentage escalation cannot be -100% or lower.')
    return money(D(base_payment) * ((Decimal(1) + r) ** int(escalation_number)))


def escalation_number(pay_date: date, first_escalation_date: date | None, frequency: str) -> int:
    if not first_escalation_date or pay_date < first_escalation_date:
        return 0
    months = FREQ_MONTHS.get(canonical_frequency(frequency), 12)
    months_diff = (pay_date.year-first_escalation_date.year)*12 + pay_date.month-first_escalation_date.month
    return max(1, months_diff // months + 1)


def _event_value_as_rate(ev: Dict[str, Any]) -> Decimal:
    # Accept the three common input spellings used by the import template.
    # `amount` is deliberately supported in addition to `value`/`rate`.
    if ev.get('rate') not in (None, ''):
        return D(ev['rate'])
    if ev.get('value') not in (None, ''):
        return D(ev['value'])
    if ev.get('amount') not in (None, ''):
        return D(ev['amount'])
    return Decimal(0)


def _event_index_multiplier(ev: Dict[str, Any]) -> Decimal | None:
    idx = ev.get('index_value', ev.get('current_index', ev.get('new_index')))
    base = ev.get('base_index_value', ev.get('base_index', ev.get('index_base')))
    if idx in (None, '') or base in (None, ''):
        return None
    idx_d, base_d = D(idx), D(base)
    if idx_d <= 0 or base_d <= 0:
        raise CalculationValidationError('Index-linked escalation requires positive index and base-index values.')
    return idx_d / base_d


def _escalation_events(lease: Dict[str, Any]) -> List[Dict[str, Any]]:
    raw = lease.get('escalation_schedule_json') or lease.get('escalation_schedule')
    if raw in (None, '', []):
        return []
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except Exception as ex:
            raise CalculationValidationError(f'Invalid escalation schedule JSON: {ex}') from ex
    if not isinstance(raw, list):
        raise CalculationValidationError('Escalation schedule must be a list of events.')

    out = []
    for i, ev in enumerate(raw, 1):
        if not isinstance(ev, dict):
            raise CalculationValidationError('Each escalation event must be an object.')
        d = parse_date(ev.get('effective_date') or ev.get('date'))
        typ = str(ev.get('type') or ev.get('method') or '').strip()
        if not d or not typ:
            raise CalculationValidationError('Escalation event requires effective_date and type.')
        seq = ev.get('event_sequence', ev.get('sequence', i))
        try:
            seq = int(seq)
        except Exception:
            raise CalculationValidationError(f'Escalation event sequence must be an integer (event {i}).')
        if seq < 1:
            raise CalculationValidationError(f'Escalation event sequence must be >= 1 (event {i}).')
        out.append({
            **ev,
            'effective_date': d,
            'type': typ,
            'value': _event_value_as_rate(ev),
            'event_sequence': seq,
        })

    out.sort(key=lambda x: (x['effective_date'], x['event_sequence']))
    seen = set()
    for ev in out:
        key = (ev['effective_date'], ev['event_sequence'])
        if key in seen:
            raise CalculationValidationError(
                f'Duplicate escalation Event Sequence {ev["event_sequence"]} on {ev["effective_date"]}.'
            )
        seen.add(key)
    return out


def calculate_contractual_payment(lease: Dict[str, Any], pay_date: date) -> Tuple[Decimal, List[Dict[str, Any]]]:
    """
    AUTHORITATIVE PAYMENT RULE.

    Base Payment is the contractual payment at commencement.  An explicit
    Escalation Schedule is therefore applied only from events whose effective
    date is on/after commencement.  Each event is applied exactly once, in
    effective-date + Event Sequence order, to the amount produced by the
    preceding event:

      Fixed %       -> current amount × (1 + rate)
      Fixed Amount  -> current amount + amount
      Reset Amount  -> current amount = reset amount
      Index linked  -> current amount × Current Index / Base Index
                       (or × (1 + value) when indices are not supplied)

    The event is effective for the first contractual payment on or after its
    effective date.  No annual/frequency interpolation is applied to an
    explicit event schedule.
    """
    base = D(lease.get('base_payment'))
    if base < 0:
        raise CalculationValidationError('Base payment cannot be negative.')

    commencement = parse_date(lease.get('commencement_date'))
    events = _escalation_events(lease)

    if events:
        if commencement:
            # Historical events are not re-applied to a Base Payment that is
            # already stated at commencement. This prevents double escalation
            # on leases whose schedule contains pre-commencement history.
            events = [ev for ev in events if ev['effective_date'] >= commencement]

        amount = base
        applied = []
        for ev in events:
            if ev['effective_date'] > pay_date:
                break

            typ = ev['type'].strip().lower()
            val = ev['value']

            if typ in INDEX_TYPES:
                multiplier = _event_index_multiplier(ev)
                if multiplier is not None:
                    amount *= multiplier
                    method = f"Index ratio {ev.get('index_value', ev.get('current_index', ev.get('new_index')))} / {ev.get('base_index_value', ev.get('base_index', ev.get('index_base')))}"
                else:
                    if val <= Decimal('-1'):
                        raise CalculationValidationError(
                            f'Index-linked escalation cannot be -100% or lower on {ev["effective_date"]}.'
                        )
                    amount *= (Decimal(1) + val)
                    method = f'Index fallback rate {val}'
            elif typ in PERCENT_ESCALATIONS:
                if val <= Decimal('-1'):
                    raise CalculationValidationError(
                        f'Percentage escalation cannot be -100% or lower on {ev["effective_date"]}.'
                    )
                amount *= (Decimal(1) + val)
                method = f'Percentage {val}'
            elif typ in FIXED_AMOUNT_ESCALATIONS:
                amount += val
                method = f'Fixed amount {val}'
            elif typ in RESET_ESCALATIONS:
                if val < 0:
                    raise CalculationValidationError(
                        f'Reset amount cannot be negative on {ev["effective_date"]}.'
                    )
                amount = val
                method = f'Reset to {val}'
            elif typ in CUSTOM_SCHEDULE_TYPES:
                raise CalculationValidationError(
                    f'Custom Schedule must contain executable events; event {ev["type"]} is not itself a pricing rule.'
                )
            else:
                raise CalculationValidationError(f'Unsupported escalation event type: {ev["type"]}')

            if amount < 0:
                raise CalculationValidationError(f'Payment amount becomes negative on {ev["effective_date"]}.')

            applied.append({
                'effective_date': ev['effective_date'],
                'event_sequence': ev['event_sequence'],
                'type': ev['type'],
                'method': method,
                'amount_after_event': money(amount),
            })

        return money(amount), applied

    typ = str(lease.get('escalation_type') or 'None').strip().lower()
    n = escalation_number(
        pay_date,
        parse_date(lease.get('first_escalation_date')),
        lease.get('escalation_frequency') or 'Annual'
    )
    if typ in ('none', '') or n == 0:
        return money(base), []

    if typ in CUSTOM_SCHEDULE_TYPES:
        raise CalculationValidationError(
            'Escalation Type "Custom Schedule" requires entries in the Escalation_Schedule sheet or Escalation Schedule JSON.'
        )
    if typ in FIXED_AMOUNT_ESCALATIONS:
        amt = D(lease.get('escalation_amount'))
        if amt < 0:
            raise CalculationValidationError('Fixed escalation amount cannot be negative.')
        amount = calculate_fixed_amount_escalation(base, amt, n)
        return amount, [{'effective_date': parse_date(lease.get('first_escalation_date')), 'event_sequence': n,
                         'type': lease.get('escalation_type'), 'method': f'Legacy fixed amount step {n}',
                         'amount_after_event': amount}]
    if typ in INDEX_TYPES or typ in PERCENT_ESCALATIONS:
        rate = D(lease.get('escalation_rate', 0))
        amount = calculate_percentage_escalation(base, rate, n)
        return amount, [{'effective_date': parse_date(lease.get('first_escalation_date')), 'event_sequence': n,
                         'type': lease.get('escalation_type'), 'method': f'Legacy percentage/index step {n}',
                         'amount_after_event': amount}]
    raise CalculationValidationError(f'Unsupported escalation type: {lease.get("escalation_type")}')


def payment_amount(lease: Dict[str, Any], pay_date: date) -> Decimal:
    return calculate_contractual_payment(lease, pay_date)[0]


def build_payment_schedule(lease: Dict[str, Any]) -> List[Dict[str, Any]]:
    c = parse_date(lease.get('commencement_date')); e = parse_date(lease.get('expiry_date')); f = parse_date(lease.get('first_payment_date'))
    freq = canonical_frequency(lease.get('payment_frequency', 'Monthly')); timing = str(lease.get('payment_timing') or 'Arrears')
    dates = generate_payment_dates(c, e, f, freq, timing)
    out = []
    ib_rate = D(lease.get('ib_rate'))
    for i, p in enumerate(dates, 1):
        amount, applied_events = calculate_contractual_payment(lease, p)
        unpaid_at_commencement = p > c or (p == c and timing.lower() != 'advance')
        is_prepaid = p < c
        if unpaid_at_commencement:
            t = Decimal((p-c).days) / Decimal(365)
            pv_exact = amount / ((Decimal(1) + ib_rate) ** t)
        else:
            t = Decimal(0); pv_exact = Decimal(0)
        out.append({
            'period': i, 'payment_date': p, 'payment': amount,
            'payment_events': applied_events,
            'unpaid_at_commencement': unpaid_at_commencement,
            'prepaid_before_commencement': is_prepaid,
            'discount_period_years': t, 'pv': money(pv_exact), '_pv_exact': pv_exact,
        })
    return out


def calculate_initial_measurement(lease: Dict[str, Any], payments: List[Dict[str, Any]], restoration_pv: Decimal | None = None) -> Dict[str, Decimal]:
    ll_exact = sum((p.get('_pv_exact', D(p['pv'])) for p in payments if p['unpaid_at_commencement']), Decimal(0))
    ll = money(ll_exact)
    advance = sum((p['payment'] for p in payments if not p['unpaid_at_commencement']), Decimal(0))
    rest = D(restoration_pv) if restoration_pv is not None else D(lease.get('restoration_pv'))
    rou = money(ll + advance + D(lease.get('prepaid')) + D(lease.get('idc')) + rest - D(lease.get('lease_incentive')))
    return {'initial_liability': ll, 'advance_at_commencement': money(advance), 'initial_rou': rou}


def _zero_liability_result(c: date, reporting_date: date) -> Dict[str, Any]:
    return {'liability_schedule': [], 'last_payment_date': c, 'stub_days': 0, 'stub_interest': Decimal(0),
            'reporting_liability': Decimal(0), 'current_liability': Decimal(0), 'noncurrent_liability': Decimal(0),
            'reporting_exact': Decimal(0)}


def calculate_lease_liability(lease: Dict[str, Any], payments: List[Dict[str, Any]], reporting_date: date) -> Dict[str, Any]:
    c = parse_date(lease.get('commencement_date')); annual = D(lease.get('ib_rate'))
    if reporting_date < c:
        return _zero_liability_result(c, reporting_date)
    opening_exact = sum((p.get('_pv_exact', D(p['pv'])) for p in payments if p['unpaid_at_commencement']), Decimal(0))
    prev = c; rows = []
    for p in payments:
        if not p['unpaid_at_commencement']:
            continue
        pd = p['payment_date']
        if pd > reporting_date:
            break
        days = (pd-prev).days
        interest_exact = opening_exact * effective_rate_for_days(annual, days)
        principal_exact = D(p['payment']) - interest_exact
        closing_exact = opening_exact + interest_exact - D(p['payment'])
        if closing_exact < -Decimal('0.10'):
            raise CalculationValidationError(f'Lease liability became materially negative at period {p["period"]}. Check payment timing, payment amounts and IBR.')
        if abs(closing_exact) <= Decimal('0.10'):
            closing_exact = Decimal(0)
        rows.append({**p, 'opening': money(opening_exact), 'interest': money(interest_exact), 'principal': money(principal_exact), 'closing': money(max(Decimal(0), closing_exact)), 'days': days})
        opening_exact = max(Decimal(0), closing_exact); prev = pd
    last_payment_date = prev
    if reporting_date > last_payment_date and opening_exact > 0:
        stub_days = (reporting_date-last_payment_date).days
        stub_exact = opening_exact * effective_rate_for_days(annual, stub_days)
        stub = money(stub_exact); reporting_exact = opening_exact + stub_exact
    else:
        stub_days = 0; stub = Decimal(0); reporting_exact = opening_exact
    reporting_liability = money(max(Decimal(0), reporting_exact))
    horizon = reporting_date + relativedelta(years=1)
    future_balance = reporting_exact; future_prev = reporting_date; current_exact = Decimal(0)
    for p in payments:
        if not p['unpaid_at_commencement'] or not (reporting_date < p['payment_date'] <= horizon):
            continue
        pd = p['payment_date']; days = (pd-future_prev).days
        interest_exact = future_balance * effective_rate_for_days(annual, days)
        principal_exact = max(Decimal(0), D(p['payment']) - interest_exact)
        closing_exact = max(Decimal(0), future_balance + interest_exact - D(p['payment']))
        current_exact += principal_exact; future_balance = closing_exact; future_prev = pd
    current = money(min(current_exact, reporting_exact)); noncurrent = money(max(Decimal(0), reporting_liability-current))
    return {'liability_schedule': rows, 'last_payment_date': last_payment_date, 'stub_days': stub_days, 'stub_interest': stub,
            'reporting_liability': reporting_liability, 'current_liability': current, 'noncurrent_liability': noncurrent,
            'reporting_exact': reporting_exact}


def calculate_rou_asset(lease: Dict[str, Any], initial_rou: Decimal, reporting_date: date) -> Dict[str, Any]:
    c = parse_date(lease.get('commencement_date')); e = parse_date(lease.get('effective_expiry_date') or lease.get('expiry_date'))
    if reporting_date < c or initial_rou <= 0:
        return {'rou_schedule': [], 'reporting_rou': Decimal(0), 'monthly_depreciation': Decimal(0), 'useful_months': 0}
    useful_months = lease_term_months(c, e)
    monthly_exact = initial_rou / Decimal(useful_months)
    rows = []; op = initial_rou; d = c; period = 1
    end = min(e, reporting_date)
    while d < end and op > 0:
        nd = min(e, d + relativedelta(months=1))
        dep = money(min(op, monthly_exact))
        if nd >= e:
            dep = money(op)  # final-period true-up
        cl = money(max(Decimal(0), op-dep))
        rows.append({'period': period, 'date': d, 'opening': money(op), 'depreciation': dep, 'closing': cl})
        op = cl; d = nd; period += 1
        if period > 10000:
            raise CalculationValidationError('ROU depreciation schedule exceeded safety limit.')
    if reporting_date >= e:
        if rows and rows[-1]['closing'] != 0:
            rows[-1]['depreciation'] = money(rows[-1]['depreciation'] + rows[-1]['closing'])
            rows[-1]['closing'] = Decimal(0)
        return {'rou_schedule': rows, 'reporting_rou': Decimal(0), 'monthly_depreciation': money(monthly_exact), 'useful_months': useful_months}
    return {'rou_schedule': rows, 'reporting_rou': money(op), 'monthly_depreciation': money(monthly_exact), 'useful_months': useful_months}


def calculate_security_deposit(lease: Dict[str, Any], reporting_date: date) -> Dict[str, Any]:
    paid = D(lease.get('deposit_paid')); c = parse_date(lease.get('commencement_date')); refund = parse_date(lease.get('deposit_refund_date'))
    if paid <= 0:
        return {'deposit_fv': Decimal(0), 'schedule': [], 'gross_carrying_amount': Decimal(0), 'ecl_allowance': Decimal(0), 'net_carrying_amount': Decimal(0), 'day_one_adjustment': Decimal(0), 'status': 'NO_DEPOSIT'}
    if not refund:
        raise CalculationValidationError('Security deposit refund date is required when deposit is paid.')
    if refund <= c:
        raise CalculationValidationError('Security deposit refund date must be after commencement date.')
    if reporting_date < c:
        return {'deposit_fv': Decimal(0), 'schedule': [], 'gross_carrying_amount': Decimal(0), 'ecl_allowance': Decimal(0), 'net_carrying_amount': Decimal(0), 'day_one_adjustment': Decimal(0), 'status': 'NOT_RECOGNISED_YET'}
    rate = D(lease.get('deposit_discount_rate'))
    if rate < 0:
        raise CalculationValidationError('Deposit discount rate cannot be negative.')
    years = Decimal((refund-c).days) / Decimal(365)
    fv = money(paid / ((Decimal(1)+rate) ** years) if rate else paid)
    end = min(refund, reporting_date); op = fv; d = c; rows = []
    while d < end:
        nd = min(refund, end, d + relativedelta(months=1)); days = (nd-d).days
        interest = money(op * effective_rate_for_days(rate, days))
        cash = paid if nd == refund and refund <= reporting_date else Decimal(0)
        cl = money(max(Decimal(0), op + interest - cash))
        rows.append({'period': len(rows)+1, 'date': nd, 'opening': money(op), 'eir_rate': rate, 'interest_income': interest, 'cash_refund': cash, 'closing': cl})
        op = cl; d = nd
    closing = rows[-1]['closing'] if rows else fv
    ecl = calculate_ecl(lease, closing)
    return {'deposit_fv': fv, 'schedule': [dict(r, ecl=ecl['allowance'], net=money(max(0, r['closing']-ecl['allowance']))) for r in rows],
            'gross_carrying_amount': money(closing), 'ecl_allowance': ecl['allowance'], 'net_carrying_amount': money(max(Decimal(0), closing-ecl['allowance'])),
            'day_one_adjustment': money(paid-fv), 'status': 'ACTIVE' if closing > 0 else 'REFUNDED'}


def calculate_ecl(lease: Dict[str, Any], gross: Decimal) -> Dict[str, Decimal | str]:
    ead = max(Decimal(0), D(gross)); pd = D(lease.get('ecl_pd')); lgd = D(lease.get('ecl_lgd'))
    if pd < 0 or pd > 1 or lgd < 0 or lgd > 1:
        raise CalculationValidationError('ECL PD and LGD must be between 0% and 100%.')
    method = str(lease.get('ecl_method') or '').strip().lower()
    if method in ('pd_lgd', 'pd×lgd', 'pd x lgd', 'pd*lgd') or pd > 0 or lgd > 0:
        horizon = D(lease.get('ecl_horizon_years'))
        if horizon <= 0:
            horizon = Decimal(1)
        cumulative_pd = Decimal(1) - ((Decimal(1)-pd) ** horizon) if pd else Decimal(0)
        allowance = ead * cumulative_pd * lgd
        discount_rate = D(lease.get('ecl_discount_rate'))
        if discount_rate < 0:
            raise CalculationValidationError('ECL discount rate cannot be negative.')
        if discount_rate:
            allowance = allowance / ((Decimal(1)+discount_rate) ** horizon)
        method_used = 'PD×LGD with horizon'
    else:
        rate = D(lease.get('deposit_ecl_rate'))
        if rate < 0 or rate > 1:
            raise CalculationValidationError('Deposit ECL rate must be between 0% and 100%.')
        allowance = ead * rate; method_used = 'Configured loss rate'
    return {'ead': money(ead), 'pd': pd, 'lgd': lgd, 'allowance': money(min(ead, max(Decimal(0), allowance))), 'method': method_used}


def calculate_fx(lease: Dict[str, Any], reporting_balance: Decimal, opening_balance: Decimal | None = None) -> Dict[str, Decimal]:
    rate = D(lease.get('reporting_fx_rate') or 1)
    if rate <= 0:
        raise CalculationValidationError('Reporting FX rate must be greater than zero.')
    foreign = money(reporting_balance)
    functional = money(foreign * rate)
    prior_rate = D(lease.get('prior_reporting_fx_rate') or lease.get('initial_fx_rate') or 0)
    fx_gain_loss = Decimal(0)
    if opening_balance is not None and prior_rate > 0:
        fx_gain_loss = money(money(opening_balance) * (rate-prior_rate))
    return {'foreign_balance': foreign, 'reporting_fx_rate': rate, 'functional_balance': functional, 'prior_fx_rate': prior_rate, 'fx_gain_loss': fx_gain_loss}


def _is_yes(v: Any) -> bool:
    return str(v or '').strip().lower() in ('y','yes','true','1')


def assess_lease_term(lease: Dict[str, Any]) -> Dict[str, Any]:
    """Central lease-term assessment used by all downstream calculations.

    Contractual expiry is never overwritten.  Assessed effective expiry is the
    accounting term after reasonably-certain extension/termination assessment.
    Modification-adjusted expiry is handled separately once an effective
    modification is actually recognised.
    """
    c = parse_date(lease.get('commencement_date'))
    contractual = parse_date(lease.get('expiry_date'))
    if not contractual:
        raise CalculationValidationError('Contractual expiry is required.')
    end = contractual
    notes = []
    if _is_yes(lease.get('extension_reasonably_certain')) and D(lease.get('extension_option_years')) > 0:
        months = int((D(lease.get('extension_option_years')) * 12).to_integral_value(rounding=ROUND_HALF_UP))
        end = contractual + relativedelta(months=months)
        notes.append(f'Extension option included: management assessed the extension as reasonably certain to be exercised; assessed expiry extended to {end.isoformat()}.')
    elif _is_yes(lease.get('extension_reasonably_certain')):
        notes.append('Extension option marked reasonably certain, but no positive extension period was supplied.')
    if _is_yes(lease.get('termination_reasonably_certain')) and lease.get('termination_option_date'):
        td = parse_date(lease.get('termination_option_date'))
        if td and c and td > c and td < end:
            end = td
            notes.append(f'Termination option included: management assessed termination as reasonably certain; assessed expiry shortened to {td.isoformat()}.')
    if c and end <= c:
        raise CalculationValidationError('Assessed lease term must be after commencement date.')
    return {
        'effective_expiry': end,
        'contractual_expiry': contractual,
        'assessed_effective_expiry': end,
        'modification_adjusted_expiry': end,
        'notes': notes,
        'useful_months': lease_term_months(c, end),
    }


def lease_term_months(commencement: Any, effective_expiry: Any) -> int:
    """Return the single controlled monthly depreciation period count.

    The accounting lease term is represented by monthly periods beginning on the
    commencement anniversary.  A lease ending one day before its anniversary
    still has the final monthly period in the depreciation schedule (e.g.
    08-Sep-2026 to 07-Sep-2039 = 156 periods).
    """
    c = parse_date(commencement)
    e = parse_date(effective_expiry)
    if not c or not e or e <= c:
        return 0
    months = (e.year - c.year) * 12 + e.month - c.month
    # Count a partial final month as one monthly depreciation period.
    # This gives 12 periods for 01-Apr to 31-Mar and 156 periods for
    # 08-Sep-2026 to 07-Sep-2039, without creating an extra period.
    if e.day >= c.day:
        months += 1
    return max(1, months)


def calculate_restoration_provision(lease: Dict[str, Any], reporting_date: date) -> Dict[str, Any]:
    cost = D(lease.get('restoration_cost_estimate')); c = parse_date(lease.get('commencement_date'))
    if reporting_date < c:
        return {'initial_pv': Decimal(0), 'schedule': [], 'closing_provision': Decimal(0), 'basis': 'NOT_RECOGNISED_YET'}
    expected = parse_date(lease.get('restoration_expected_date')) or parse_date(lease.get('expiry_date'))
    rate = D(lease.get('restoration_discount_rate'))
    if rate < 0:
        raise CalculationValidationError('Restoration discount rate cannot be negative.')
    if cost <= 0:
        manual = D(lease.get('restoration_pv'))
        if manual <= 0:
            return {'initial_pv': Decimal(0), 'schedule': [], 'closing_provision': Decimal(0), 'basis': 'NO_RESTORATION_PROVISION'}
        # For legacy manual PV, unwind only when an expected date and rate are supplied;
        # otherwise preserve the manually supplied carrying amount without inventing accretion.
        if expected and rate > 0 and expected > c:
            initial_pv = money(manual)
            end = min(reporting_date, expected); op = initial_pv; d = c; rows = []
            while d < end:
                nd = min(expected, end, d + relativedelta(months=12)); days = (nd-d).days
                unwind = money(op * effective_rate_for_days(rate, days)); cl = money(op + unwind)
                rows.append({'period': len(rows)+1, 'date': nd, 'opening': op, 'unwinding': unwind, 'closing': cl}); op = cl; d = nd
            return {'initial_pv': initial_pv, 'schedule': rows, 'closing_provision': rows[-1]['closing'] if rows else initial_pv, 'basis': 'manual_pv_unwound'}
        return {'initial_pv': money(manual), 'schedule': [], 'closing_provision': money(manual), 'basis': 'manual_pv_input'}
    if not expected or expected <= c:
        raise CalculationValidationError('Restoration expected date must be after commencement date when a restoration cost estimate is supplied.')
    years = Decimal((expected-c).days) / Decimal(365)
    initial_pv = money(cost / ((Decimal(1)+rate) ** years) if rate else cost)
    end = min(reporting_date, expected); op = initial_pv; d = c; rows = []
    while d < end:
        nd = min(expected, end, d + relativedelta(months=12)); days = (nd-d).days
        unwind = money(op * effective_rate_for_days(rate, days)); cl = money(op + unwind)
        if nd >= expected:
            cl = money(cost)
            unwind = money(cl-op)
        rows.append({'period': len(rows)+1, 'date': nd, 'opening': op, 'unwinding': unwind, 'closing': cl}); op = cl; d = nd
    return {'initial_pv': initial_pv, 'schedule': rows, 'closing_provision': rows[-1]['closing'] if rows else initial_pv, 'basis': 'discounted_cost_estimate'}


def maturity_analysis(payments: List[Dict[str, Any]], reporting_date: date) -> Dict[str, Decimal]:
    labels = ['<1 year','1-2 years','2-3 years','3-4 years','4-5 years','>5 years']; buckets = {l: Decimal(0) for l in labels}
    bounds = [reporting_date + relativedelta(years=n) for n in range(1,6)]
    for p in payments:
        pdte = p['payment_date']
        if pdte <= reporting_date:
            continue
        amt = D(p['payment'])
        for lbl, bound in zip(labels[:-1], bounds):
            if pdte <= bound:
                buckets[lbl] += amt; break
        else:
            buckets['>5 years'] += amt
    return {k: money(v) for k,v in buckets.items()}


def _guard_abnormal_escalation(lease: Dict[str, Any], payments: List[Dict[str, Any]]) -> None:
    base = D(lease.get('base_payment'))
    if base < 0:
        raise CalculationValidationError('Base payment cannot be negative.')
    if payments and max((D(p['payment']) for p in payments), default=Decimal(0)) > base * Decimal('100000'):
        raise CalculationValidationError('Abnormal escalation detected: payment exceeds the configured reasonableness ceiling.')


_REGIME_OVERRIDE_KEYS = ['base_payment','ib_rate','expiry_date','escalation_type','escalation_amount','escalation_rate',
                         'first_escalation_date','escalation_frequency','payment_frequency','payment_timing','escalation_schedule_json']


def _next_payment_date(previous_lease: Dict[str, Any], effective_date: date, effective_expiry: date) -> date | None:
    c = parse_date(previous_lease.get('commencement_date')); fp = parse_date(previous_lease.get('first_payment_date'))
    freq = canonical_frequency(previous_lease.get('payment_frequency', 'Monthly')); timing = previous_lease.get('payment_timing', 'Arrears')
    if not (c and fp):
        return None
    try:
        dates = generate_payment_dates(c, effective_expiry, fp, freq, timing)
    except CalculationValidationError:
        return None
    for d in dates:
        if d >= effective_date:
            return d
    return None


def build_regimes(lease: Dict[str, Any], modifications: List[Dict[str, Any]] | None, effective_expiry: date) -> List[Dict[str, Any]]:
    mods = sorted([m for m in (modifications or []) if m.get('effective_date')], key=lambda m: (parse_date(m['effective_date']), str(m.get('modification_id',''))))
    regimes = []; current = dict(lease); current['expiry_date'] = effective_expiry; seg_start = parse_date(lease.get('commencement_date'))
    seen = set()
    for m in mods:
        eff = parse_date(m.get('effective_date'))
        if not eff:
            continue
        if eff in seen:
            raise CalculationValidationError(f'Multiple modifications have the same effective date {eff.isoformat()}; resolve sequencing before calculation.')
        seen.add(eff)
        if eff <= seg_start or eff > current['expiry_date']:
            raise CalculationValidationError(f'Modification effective date {eff} is outside the active lease term.')
        seg = dict(current); seg['commencement_date'] = seg_start; regimes.append({'lease':seg,'start':seg_start,'boundary':eff,'modification':m})
        nxt = dict(current)
        for k in _REGIME_OVERRIDE_KEYS:
            src = 'new_'+k
            if m.get(src) not in (None, ''):
                nxt[k] = m[src]
        if not m.get('new_first_payment_date'):
            inherited_next = _next_payment_date(current, eff, current['expiry_date'])
            if inherited_next:
                nxt['first_payment_date'] = inherited_next.isoformat()
        else:
            nxt['first_payment_date'] = m['new_first_payment_date']
        nxt['expiry_date'] = parse_date(nxt.get('expiry_date')) or effective_expiry
        if nxt['expiry_date'] <= eff:
            raise CalculationValidationError(f'Modification at {eff} leaves no valid revised lease term.')
        current = nxt; seg_start = eff
    seg = dict(current); seg['commencement_date'] = seg_start; regimes.append({'lease':seg,'start':seg_start,'boundary':None,'modification':None})
    return regimes


def calculate_lease_with_modifications(lease: Dict[str, Any], modifications: List[Dict[str, Any]] | None, reporting_date: date) -> Dict[str, Any]:
    # Future-effective modifications are not known/recognized at the reporting date.
    # Keep them in the event/modification history, but exclude them from measurement.
    rd = parse_date(reporting_date)
    applicable_modifications = []
    for m in (modifications or []):
        eff = parse_date(m.get('effective_date')) if m.get('effective_date') else None
        if eff and eff <= rd:
            applicable_modifications.append(m)
    term = assess_lease_term(lease); effective_expiry = term['effective_expiry']; regimes = build_regimes(lease, applicable_modifications, effective_expiry)
    restoration = calculate_restoration_provision(lease, reporting_date)
    all_payments: List[Dict[str, Any]] = []; modification_log = []; rou_schedule_all = []
    carry_rou = None; initial_liability = None; initial_rou = None; advance0 = Decimal(0); final = None; reporting_rou = Decimal(0); monthly_dep = Decimal(0)
    for idx, r in enumerate(regimes):
        rl = r['lease']; payments = build_payment_schedule(rl); _guard_abnormal_escalation(rl, payments)
        for p in payments:
            if r['boundary'] is None or p['payment_date'] < r['boundary']:
                p2 = dict(p); p2['regime'] = idx; all_payments.append(p2)
        init_meas = calculate_initial_measurement(rl, payments, restoration_pv=restoration['initial_pv'] if idx == 0 else Decimal(0))
        if idx == 0:
            initial_liability = init_meas['initial_liability']; advance0 = init_meas['advance_at_commencement']
            initial_rou = money(init_meas['initial_liability'] + init_meas['advance_at_commencement'] + D(lease.get('prepaid')) + D(lease.get('idc')) + restoration['initial_pv'] - D(lease.get('lease_incentive')))
            dep = calculate_security_deposit(lease, reporting_date); initial_rou = money(initial_rou + dep.get('day_one_adjustment', Decimal(0))); carry_rou = initial_rou
        else:
            prev = regimes[idx-1]; prev_payments = build_payment_schedule(prev['lease']); prev_roll = calculate_lease_liability(prev['lease'], prev_payments, r['start'])
            carrying_before = prev_roll['reporting_liability']; revised = init_meas['initial_liability']; rou_adj = money(revised-carrying_before); pl_gain = Decimal(0)
            if carry_rou + rou_adj < 0:
                pl_gain = money(-(carry_rou + rou_adj)); carry_rou = Decimal(0)
            else:
                carry_rou = money(carry_rou + rou_adj)
            trigger = prev.get('modification') or {}
            modification_log.append({'effective_date':r['start'],'mod_type':trigger.get('mod_type','Modification'),'old_liability':carrying_before,'revised_liability':revised,'rou_adjustment':rou_adj,'pl_gain_loss':pl_gain})
        stop_at = r['boundary'] if r['boundary'] else reporting_date
        rou_calc = calculate_rou_asset({**rl,'commencement_date':r['start'],'effective_expiry_date':rl.get('expiry_date')}, carry_rou, min(reporting_date, stop_at))
        for row in rou_calc['rou_schedule']:
            row2 = dict(row); row2['regime'] = idx; rou_schedule_all.append(row2)
        monthly_dep = rou_calc['monthly_depreciation']
        if r['boundary'] is None or r['boundary'] > reporting_date:
            final = calculate_lease_liability(rl, payments, reporting_date); reporting_rou = rou_calc['reporting_rou']; break
        carry_rou = money(max(Decimal(0), rou_calc['reporting_rou']))
    if final is None:
        last = regimes[-1]; final = calculate_lease_liability(last['lease'], build_payment_schedule(last['lease']), reporting_date)
    # Remove duplicate contractual cash events at a modification boundary if two regimes carry the same payment date.
    dedup: Dict[tuple, Dict[str, Any]] = {}
    for p in all_payments:
        dedup[(p['payment_date'], p['regime'])] = p
    all_payments = list(dedup.values())
    all_payments.sort(key=lambda p: (p['payment_date'], p['regime']))
    return {'payments':all_payments,'initial_liability':initial_liability,'advance_at_commencement':advance0,'initial_rou':initial_rou,
            'liability_schedule':final['liability_schedule'],'last_payment_date':final['last_payment_date'],'stub_days':final['stub_days'],
            'stub_interest':final['stub_interest'],'reporting_liability':final['reporting_liability'],'current_liability':final['current_liability'],
            'noncurrent_liability':final['noncurrent_liability'],'rou_schedule':rou_schedule_all,'reporting_rou':reporting_rou,
            'monthly_depreciation':monthly_dep,'modification_log':modification_log,'effective_expiry':effective_expiry,'lease_term_notes':term['notes'],'restoration':restoration}


def build_journal_entries(lease: Dict[str, Any], calc: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    lid = lease.get('lease_id'); c = parse_date(lease.get('commencement_date')); rd = parse_date(calc.get('reporting_date'))
    if rd and rd < c:
        return [], {'debit_total': Decimal(0), 'credit_total': Decimal(0), 'difference': Decimal(0), 'status': 'PASS', 'event_status': 'NO_ACCOUNTING_EVENT_YET'}
    gl = {'rou':lease.get('gl_rou') or 'ROU Asset - Other','liability':lease.get('gl_liability') or 'Lease Liability','interest':lease.get('gl_interest') or 'Finance Cost','bank':lease.get('gl_bank') or 'Bank','dep':lease.get('gl_depreciation') or 'Depreciation Expense','accum_dep':lease.get('gl_accum_dep') or 'Accumulated Depreciation - ROU'}
    rows = []
    initial_liability = D(calc.get('initial_liability')); advance = D(calc.get('advance_at_commencement')); prepaid = D(lease.get('prepaid')); idc = D(lease.get('idc')); rest = D(calc.get('restoration', {}).get('initial_pv', 0)); incentive = D(lease.get('lease_incentive'))
    if initial_liability:
        rows.append({'lease_id':lid,'date':c,'je_type':'Initial Recognition - Lease Liability','debit_gl':gl['rou'],'credit_gl':gl['liability'],'debit':initial_liability,'credit':initial_liability})
    for amt, typ, debit_gl, credit_gl in [(advance,'Initial Recognition - Advance Payment',gl['rou'],gl['bank']),(prepaid,'Initial Recognition - Prepayment',gl['rou'],gl['bank']),(idc,'Initial Recognition - Initial Direct Costs',gl['rou'],gl['bank']),(rest,'Initial Recognition - Restoration Provision',gl['rou'],'Restoration Provision')]:
        if amt: rows.append({'lease_id':lid,'date':c,'je_type':typ,'debit_gl':debit_gl,'credit_gl':credit_gl,'debit':amt,'credit':amt})
    if incentive: rows.append({'lease_id':lid,'date':c,'je_type':'Initial Recognition - Lease Incentive','debit_gl':gl['bank'],'credit_gl':gl['rou'],'debit':incentive,'credit':incentive})
    for p in calc.get('liability_schedule',[]):
        if D(p['interest']): rows.append({'lease_id':lid,'date':p['payment_date'],'je_type':'Periodic Interest','debit_gl':gl['interest'],'credit_gl':gl['liability'],'debit':D(p['interest']),'credit':D(p['interest'])})
        if D(p['payment']): rows.append({'lease_id':lid,'date':p['payment_date'],'je_type':'Lease Payment','debit_gl':gl['liability'],'credit_gl':gl['bank'],'debit':D(p['payment']),'credit':D(p['payment'])})
    if D(calc.get('stub_interest')): rows.append({'lease_id':lid,'date':rd,'je_type':'Stub Interest to Reporting Date','debit_gl':gl['interest'],'credit_gl':gl['liability'],'debit':D(calc['stub_interest']),'credit':D(calc['stub_interest'])})
    for p in calc.get('rou_schedule',[]):
        if D(p['depreciation']): rows.append({'lease_id':lid,'date':p['date'],'je_type':'Depreciation','debit_gl':gl['dep'],'credit_gl':gl['accum_dep'],'debit':D(p['depreciation']),'credit':D(p['depreciation'])})
    dep = calc.get('deposit', {}); fv = D(dep.get('deposit_fv',0)); paid = D(lease.get('deposit_paid'))
    if fv: rows.append({'lease_id':lid,'date':c,'je_type':'Security Deposit - Initial FV','debit_gl':'Security Deposit Financial Asset','credit_gl':gl['bank'],'debit':fv,'credit':fv})
    day1 = money(paid-fv)
    if day1: rows.append({'lease_id':lid,'date':c,'je_type':'Security Deposit - Day 1 Lease Adjustment','debit_gl':gl['rou'],'credit_gl':gl['bank'],'debit':day1,'credit':day1})
    for p in dep.get('schedule',[]):
        if D(p.get('interest_income',0)): rows.append({'lease_id':lid,'date':p['date'],'je_type':'Deposit EIR','debit_gl':'Security Deposit Financial Asset','credit_gl':'Interest Income','debit':D(p['interest_income']),'credit':D(p['interest_income'])})
    for p in calc.get('restoration',{}).get('schedule',[]):
        if D(p.get('unwinding',0)): rows.append({'lease_id':lid,'date':p['date'],'je_type':'Restoration Unwinding','debit_gl':'Finance Cost - Restoration','credit_gl':'Restoration Provision','debit':D(p['unwinding']),'credit':D(p['unwinding'])})
    ecl = D(calc.get('ecl',{}).get('allowance',0))
    if ecl: rows.append({'lease_id':lid,'date':rd,'je_type':'ECL Allowance','debit_gl':'Impairment Loss','credit_gl':'ECL Allowance - Security Deposit','debit':ecl,'credit':ecl})
    for m in calc.get('modification_log',[]):
        adj = D(m['rou_adjustment']);
        if adj > 0: rows.append({'lease_id':lid,'date':m['effective_date'],'je_type':f"Modification ({m['mod_type']}) - Increase",'debit_gl':gl['rou'],'credit_gl':gl['liability'],'debit':adj,'credit':adj})
        elif adj < 0: rows.append({'lease_id':lid,'date':m['effective_date'],'je_type':f"Modification ({m['mod_type']}) - Decrease",'debit_gl':gl['liability'],'credit_gl':gl['rou'],'debit':-adj,'credit':-adj})
        pl = D(m['pl_gain_loss'])
        if pl: rows.append({'lease_id':lid,'date':m['effective_date'],'je_type':f"Modification ({m['mod_type']}) - Gain/Loss",'debit_gl':gl['liability'],'credit_gl':'Gain on Lease Modification','debit':pl,'credit':pl})
    debit_total = sum((D(r['debit']) for r in rows), Decimal(0)); credit_total = sum((D(r['credit']) for r in rows), Decimal(0)); diff = money(debit_total-credit_total); tol = tolerance(lease)
    return rows, {'debit_total':money(debit_total),'credit_total':money(credit_total),'difference':diff,'status':'PASS' if abs(diff)<=tol else 'ERROR','tolerance':tol}


def validate(lease: Dict[str, Any]) -> List[Tuple[str, str]]:
    errs: List[Tuple[str, str]] = []
    required = ('lease_id','commencement_date','expiry_date','first_payment_date','payment_frequency','payment_timing','base_payment','ib_rate')
    for f in required:
        if lease.get(f) in (None, ''):
            errs.append((f'[{f}] Missing required input.', 'HIGH'))
    try:
        c = parse_date(lease.get('commencement_date')); e = parse_date(lease.get('expiry_date')); fp = parse_date(lease.get('first_payment_date'))
        if c and e and e <= c: errs.append(('[expiry_date] Expiry must be after commencement.', 'CRITICAL'))
        if c and fp and fp < c and str(lease.get('payment_timing')).strip().lower() == 'arrears': errs.append(('[first_payment_date] Arrears first payment cannot precede commencement.', 'HIGH'))
        if c and fp and fp < c and str(lease.get('payment_timing')).strip().lower() == 'advance': pass
    except Exception as ex: errs.append((f'[DATES] {ex}', 'CRITICAL'))
    if canonical_frequency(lease.get('payment_frequency','Monthly')) not in FREQ_MONTHS: errs.append(('[payment_frequency] Unsupported frequency.', 'CRITICAL'))
    esc_raw = str(lease.get('escalation_type') or 'None').strip()
    esc_key = esc_raw.lower()
    supported_esc = {
        'none', 'fixed amount', 'fixed %', 'fixed rate', 'fixed-rate',
        'fixed rate %', 'rate', 'cpi', 'cpi linked', 'cpi-linked', 'cpi link',
        'cpi index', 'cpi indexation', 'wpi', 'wpi linked', 'wpi-linked',
        'wpi index', 'wpi indexation', 'other index', 'other index linked', 'other-index'
    }
    if esc_key not in supported_esc and esc_key not in CUSTOM_SCHEDULE_TYPES:
        errs.append((f'[escalation_type] Unsupported escalation type: {esc_raw}. Supported: None, Fixed Amount, Fixed %, Fixed Rate, CPI, CPI Linked, WPI, WPI Linked, Other Index, Custom Schedule.', 'CRITICAL'))
    elif esc_key in CUSTOM_SCHEDULE_TYPES and not _escalation_events(lease):
        errs.append(('[escalation_type] Custom Schedule requires one or more Escalation_Schedule rows or Escalation_Schedule_JSON events.', 'HIGH'))
    for fld in ('base_payment','prepaid','idc','lease_incentive','deposit_paid','extension_option_years','restoration_cost_estimate'):
        try:
            if D(lease.get(fld)) < 0: errs.append((f'[{fld}] Value cannot be negative.', 'HIGH'))
        except CalculationValidationError as ex: errs.append((f'[{fld}] {ex}', 'CRITICAL'))
    for fld in ('ib_rate','deposit_discount_rate','restoration_discount_rate','ecl_discount_rate'):
        try:
            if D(lease.get(fld)) < 0: errs.append((f'[{fld}] Rate cannot be negative.', 'HIGH'))
        except CalculationValidationError as ex: errs.append((f'[{fld}] {ex}', 'CRITICAL'))
    for fld in ('deposit_ecl_rate','ecl_pd','ecl_lgd'):
        try:
            if not (Decimal(0) <= D(lease.get(fld)) <= Decimal(1)): errs.append((f'[{fld}] Must be between 0% and 100%.', 'HIGH'))
        except CalculationValidationError as ex: errs.append((f'[{fld}] {ex}', 'CRITICAL'))
    try:
        if D(lease.get('reporting_fx_rate') if lease.get('reporting_fx_rate') not in (None,'') else 1) <= 0: errs.append(('[reporting_fx_rate] Reporting FX rate must be greater than zero.', 'CRITICAL'))
        if D(lease.get('validation_tolerance', DEFAULT_TOLERANCE)) <= 0: errs.append(('[validation_tolerance] Must be greater than zero.', 'HIGH'))
    except CalculationValidationError as ex: errs.append((f'[INPUT] {ex}', 'CRITICAL'))
    if lease.get('deposit_paid') and not lease.get('deposit_refund_date'): errs.append(('[deposit_refund_date] Required when deposit is paid.', 'HIGH'))
    try:
        if lease.get('deposit_paid') and lease.get('deposit_refund_date') and parse_date(lease.get('deposit_refund_date')) <= parse_date(lease.get('commencement_date')): errs.append(('[deposit_refund_date] Must be after commencement.', 'CRITICAL'))
    except Exception as ex: errs.append((f'[deposit_refund_date] {ex}', 'CRITICAL'))
    try:
        events = _escalation_events(lease); c = parse_date(lease.get('commencement_date')); e = parse_date(lease.get('expiry_date'))
        for ev in events:
            if c and ev['effective_date'] < c: errs.append((f'[escalation_schedule] Event {ev["effective_date"]} precedes commencement.', 'CRITICAL'))
            if e and ev['effective_date'] > e: errs.append((f'[escalation_schedule] Event {ev["effective_date"]} is after expiry.', 'CRITICAL'))
            if ev['type'].strip().lower() in FIXED_AMOUNT_ESCALATIONS and ev['value'] < 0: errs.append((f'[escalation_schedule] Fixed amount cannot be negative on {ev["effective_date"]}.', 'HIGH'))
            if ev['type'].strip().lower() in PERCENT_ESCALATIONS or ev['type'].strip().lower() in INDEX_TYPES:
                if ev['value'] <= -1 and _event_index_multiplier(ev) is None: errs.append((f'[escalation_schedule] Rate cannot be -100% or lower on {ev["effective_date"]}.', 'HIGH'))
                if _event_index_multiplier(ev) is not None:
                    _event_index_multiplier(ev)
    except Exception as ex: errs.append((f'[escalation_schedule] {ex}', 'CRITICAL'))
    try:
        assess_lease_term(lease)
    except Exception as ex: errs.append((f'[lease_term] {ex}', 'CRITICAL'))
    return errs


def validate_calculation(lease: Dict[str, Any], result: Dict[str, Any]) -> List[Tuple[str, str]]:
    errs: List[Tuple[str, str]] = []; tol = tolerance(lease); rd = parse_date(result.get('reporting_date')); c = parse_date(lease.get('commencement_date'))
    if rd < c:
        expected_zero = ('initial_liability','initial_rou','reporting_liability','reporting_rou','current_liability','noncurrent_liability')
        for k in expected_zero:
            if abs(D(result.get(k))) > tol: errs.append((f'[PRE_COMMENCEMENT] {k} must be zero before commencement.', 'CRITICAL'))
        if result.get('journal_entries'): errs.append(('[PRE_COMMENCEMENT] No journal entries are allowed before commencement.', 'CRITICAL'))
        return errs
    if D(result.get('current_liability')) + D(result.get('noncurrent_liability')) != D(result.get('reporting_liability')):
        errs.append(('[CURRENT_NONCURRENT] Current + non-current does not reconcile to reporting liability.', 'CRITICAL'))
    debit = D(result.get('je_totals',{}).get('debit_total')); credit = D(result.get('je_totals',{}).get('credit_total'))
    if abs(debit-credit) > tol or result.get('je_totals',{}).get('status') != 'PASS': errs.append(('[JOURNAL_ENTRY] Journal entries are out of balance.', 'CRITICAL'))
    if D(result.get('reporting_liability')) < -tol or D(result.get('reporting_rou')) < -tol: errs.append(('[CARRYING_VALUES] Recognised carrying values cannot be negative.', 'CRITICAL'))
    for p in result.get('liability_schedule',[]):
        if D(p.get('closing')) < -tol: errs.append((f'[LIABILITY] Negative closing balance at {p.get("payment_date")}.', 'CRITICAL'))
    return errs


def build_measurement_schedule(lease: Dict[str, Any], payments: List[Dict[str, Any]], reporting_date: date, basis: str='Monthly') -> List[Dict[str, Any]]:
    c = parse_date(lease.get('commencement_date')); e = parse_date(lease.get('expiry_date'))
    if not c or reporting_date < c:
        return []
    end = min(reporting_date, e); basis = (basis or 'Monthly').strip().title(); dates = []; d = c
    if basis == 'Daily':
        while d <= end: dates.append(d); d += relativedelta(days=1)
    else:
        dates = [c]; d = c + relativedelta(months=1)
        while d < end: dates.append(d); d += relativedelta(months=1)
        if dates[-1] != end: dates.append(end)
    pmap = {p['payment_date']:p for p in payments}; annual = D(lease.get('ib_rate')); opening = sum((p.get('_pv_exact', D(p['pv'])) for p in payments if p['unpaid_at_commencement']), Decimal(0)); prev = c; rows = []
    for d in dates:
        days = (d-prev).days; interest = opening*effective_rate_for_days(annual, days) if days else Decimal(0); cash = pmap.get(d,{}).get('payment',Decimal(0))
        principal = D(cash)-interest; closing = max(Decimal(0), opening+interest-cash)
        rows.append({'date':d,'opening':money(opening),'days':days,'interest':money(interest),'payment':money(cash),'principal':money(principal),'closing':money(closing),'basis':basis,'payment_event':'YES' if cash else 'NO'})
        opening = closing; prev = d
    return rows


def calculate(lease: Dict[str, Any], reporting_date: date | None = None, modifications: List[Dict[str, Any]] | None = None, measurement_basis: str='Monthly') -> Dict[str, Any]:
    reporting_date = reporting_date or date.today()
    validation = validate(lease)
    if validation:
        raise CalculationValidationError('; '.join(f'{s}: {m}' for m,s in validation))
    c = parse_date(lease.get('commencement_date')); term = assess_lease_term(lease); payments = build_payment_schedule({**lease,'expiry_date':term['effective_expiry']})
    if reporting_date < c:
        maturity = maturity_analysis(payments, reporting_date)
        result = {'payments':payments,'initial_liability':Decimal(0),'advance_at_commencement':Decimal(0),'initial_rou':Decimal(0),'liability_schedule':[],
                  'last_payment_date':c,'stub_days':0,'stub_interest':Decimal(0),'reporting_liability':Decimal(0),'current_liability':Decimal(0),
                  'noncurrent_liability':Decimal(0),'rou_schedule':[],'reporting_rou':Decimal(0),'monthly_depreciation':Decimal(0),'restoration':calculate_restoration_provision(lease,reporting_date),
                  'modification_log':[],'effective_expiry':term['effective_expiry'],'lease_term_notes':term['notes'],'reporting_date':reporting_date,
                  'deposit':calculate_security_deposit(lease,reporting_date),'ecl':calculate_ecl(lease,Decimal(0)),'deposit_day_one_adjustment':Decimal(0),
                  'maturity':maturity,'measurement_basis':(measurement_basis or 'Monthly').title(),'measurement_schedule':[],
                  'fx_liability':calculate_fx(lease,Decimal(0)),'fx_current_liability':calculate_fx(lease,Decimal(0)),'fx_noncurrent_liability':calculate_fx(lease,Decimal(0)),
                  'fx_deposit':calculate_fx(lease,Decimal(0))}
        result['journal_entries'], result['je_totals'] = build_journal_entries(lease, result)
        result['validation_status'] = 'PASS' if not validate_calculation(lease, result) else 'FAIL'
        return result
    use_multi = bool(modifications) or term['effective_expiry'] != term['contractual_expiry']
    if not use_multi:
        active = {**lease,'expiry_date':term['effective_expiry']}; restoration = calculate_restoration_provision(active,reporting_date); payments = build_payment_schedule(active); _guard_abnormal_escalation(active,payments)
        initial = calculate_initial_measurement(active,payments,restoration_pv=restoration['initial_pv']); dep = calculate_security_deposit(active,reporting_date); initial['initial_rou'] = money(initial['initial_rou']+dep.get('day_one_adjustment',Decimal(0)))
        ll = calculate_lease_liability(active,payments,reporting_date); rou = calculate_rou_asset({**active,'effective_expiry_date':term['effective_expiry']},initial['initial_rou'],reporting_date)
        result = {'payments':payments,**initial,**ll,**rou,'restoration':restoration,'modification_log':[],'effective_expiry':term['effective_expiry'],'lease_term_notes':term['notes']}
    else:
        result = calculate_lease_with_modifications(lease, modifications, reporting_date)
    dep = calculate_security_deposit(lease,reporting_date); ecl = calculate_ecl(lease,dep['gross_carrying_amount'])
    result.update({'deposit':dep,'ecl':ecl,'deposit_day_one_adjustment':dep.get('day_one_adjustment',Decimal(0)),'reporting_date':reporting_date,'contractual_expiry':term['contractual_expiry'],'assessed_effective_expiry':result.get('effective_expiry',term['effective_expiry']),'modification_adjusted_expiry':result.get('effective_expiry',term['effective_expiry']),'lease_term_months':term['useful_months'],'maturity':maturity_analysis(result['payments'],reporting_date),'measurement_basis':(measurement_basis or 'Monthly').title(),'measurement_schedule':build_measurement_schedule(lease,result['payments'],reporting_date,measurement_basis)})
    result['fx_liability'] = calculate_fx(lease,result['reporting_liability'],result.get('initial_liability')); result['fx_current_liability'] = calculate_fx(lease,result['current_liability']); result['fx_noncurrent_liability'] = calculate_fx(lease,result['noncurrent_liability']); result['fx_deposit'] = calculate_fx(lease,dep.get('gross_carrying_amount',Decimal(0)))
    result['journal_entries'], result['je_totals'] = build_journal_entries(lease,result)
    calc_errors = validate_calculation(lease,result); result['validation_errors'] = calc_errors; result['validation_status'] = 'PASS' if not calc_errors else 'FAIL'
    if calc_errors:
        raise CalculationValidationError('; '.join(f'{s}: {m}' for m,s in calc_errors))
    return result
