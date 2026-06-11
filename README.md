# tropi-service-common

Shared service plumbing for the Tropi Railway fleet. **Public repo — no business data** (same rule as tropi-storage-adapter / tropi-excel-adapter).

| Module | What | Replaces |
|---|---|---|
| `tropi_common.excel_safe` | `SafeWorkbook` OOXML-safe ZIP/XML xlsx editor — union build (keyaccounts base + warehouse-receipts grafts: `cell_str` xml:space-preserve, `shift_formula_refs`) | 8 per-service copies |
| `tropi_common.sentry` | `init_sentry()` fleet-standard Sentry block | 12 copy-pasted blocks |
| `tropi_common.cc_track` | `record_flow()` Command-Center telemetry | 7 byte-identical copies |
| `tropi_common.telegram` | `send()` — log-and-swallow contract, never raises | 3 skeletons |

## Pinning

Pin by commit SHA, exactly like the adapters:
```
tropi-service-common @ git+https://github.com/stoynovski-a11y/tropi-service-common.git@<SHA>
```
A change here deploys nowhere until a service bumps its own pin. Rollback = revert that service's pin-bump PR.

## Tests
`python -m pytest tests/` — includes the golden-file round-trip (open → save → byte-compare modulo documented save() transforms) and the regression for every grafted fix.
