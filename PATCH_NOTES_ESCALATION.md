# V7.3 Escalation & Clean-Company Patch

## Payment schedule correction
The Python calculation engine is now the authoritative source for contractual payment amounts.

Rules:
1. Base Payment is treated as the contractual payment at lease commencement.
2. Explicit Escalation_Schedule events are processed chronologically by Effective Date and Event Sequence.
3. An event affects the first contractual payment on/after its Effective Date.
4. Fixed %: current payment × (1 + rate).
5. Fixed Amount: current payment + amount.
6. Reset Amount: current payment is replaced by the reset amount.
7. CPI/WPI/Other Index: when Base Index and Current Index are provided, current payment × Current Index / Base Index; otherwise the configured rate is used as fallback.
8. Multiple events on the same date are applied in Event Sequence order.
9. Explicit event schedules do not use Escalation Frequency; that field is only for the legacy single-rule escalation fields.
10. Events dated before commencement are ignored so historical escalations cannot be applied a second time to a Base Payment already stated at commencement.
11. The Excel Payment Schedule keeps a formula cell for traceability, but its displayed formula snapshots the exact Python-engine result. This prevents a second, divergent Excel escalation implementation from producing a different payment.
12. Payment Event now records the applied escalation path for each payment row.

## Company cleanup
All existing companies and their associated users, leases, events, modifications, calculation runs, errors, audit records, period locks, import batches, settings and calculation snapshots were removed from the bundled SQLite database. The application starts with an empty company list and no legacy company data.

## Validation
The escalation regression suite and existing extracted project tests pass: 8 tests passed.
