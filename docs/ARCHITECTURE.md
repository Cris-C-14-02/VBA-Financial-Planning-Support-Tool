# Architecture

## Runtime boundaries

`AddItemForm` and `AddSavingGoalForm` parse user input and call `FinanceData`. `OutputForm` invokes `FinanceReports.RefreshFinancialReport`. `FinanceCharts` owns chart series and layout. All worksheet access goes through `ThisWorkbook`, directly or through `WorkbookSupport.AppSheet`.

The original standard-module names remain as compatibility adapters so existing controls keep their macro bindings. UserForm identities, controls and designer resources remain in the workbook.

## Worksheet contracts

| Sheet | Columns / cells | Contract |
| --- | --- | --- |
| `Expenses&Incomes` | A:E = Date, Item, Value, Category, Description | Header in row 1; positive amount; category determines its sign |
| `Goals` | A:D = Beginning date, End date, Opening savings, Target | Most recently populated row is active; start ≤ end, opening ≥ 0, target > 0 |
| `Output` | E:J = Date, Item, Amount, Category, Description, Running balance | Derived report, sorted by ascending date |
| `Output` | A2 / A4 / B2 | Active goal start / end / opening balance |
| `Output` | A6:B8 | Savings target, closing balance and nonnegative amount remaining |
| `Output` | L2:N5 / L6:M6 | Four expense-category amounts and shares / total spending |
| `Financial Advice` | Existing control surface | Original entry point now reports the largest expense category and amount |

The latest-goal rule replaces the previous implicit first-row rule. Older goals stay in the input sheet. Transactions on both boundary dates are included, and transactions outside that window do not affect this goal's balance.

## Calculation rules

For transactions in the active period:

```text
running balance = opening balance + cumulative income − cumulative expenses
total spending = Shopping + Entertainment + Bills + Other
category share = category spending / total spending
remaining to goal = max(target − closing balance, 0)
```

Zero spending produces blank shares and no spending chart. An empty period produces the opening balance as its closing balance and no transaction charts. Money is represented as VBA `Currency`; dates are typed Excel dates. The dollar format is a display convention, with no currency conversion.

## Refresh behavior

1. Validate the latest goal and every nonempty transaction row. A blank row does not stop the scan.
2. Replace the old output, copy eligible transactions and sort by date.
3. Calculate balances and category totals, then write summary values.
4. Create the named `SavingsBalance` and `SpendingBreakdown` charts when data exists. Charts sit below the report data and have separate horizontal positions.

The generated views are a macro-refreshed snapshot. Submitting a new transaction or goal invalidates the old report by clearing it. Direct source-cell edits require another report generation. Equal-date transactions retain no guaranteed intra-day ordering; their aggregate closing balance is unaffected.

## Migration from the original prototype

| Original concern | Implementation |
| --- | --- |
| Calculations embedded in OutputForm | `FinanceReports` and `FinanceCharts` |
| Duplicated nested nonempty checks | Typed parsing and guard clauses |
| Goal target field not actually validated | Separate positive-target validation |
| Income included in spending denominator | Expense-only total |
| Dates shown without filtering the report | Inclusive goal-period filtering |
| String balance and Integer row counters | `Currency` balances and `Long` row bounds |
| Unqualified cells and ranges | Explicit workbook and worksheet references |
| Old rows and overlapping chart objects | Bounded clearing and managed chart replacement |

`src/legacy/` retains the original loose text files byte-for-byte. The original workbook is recoverable in Git history. The current `.xlsm` and `src/manifest.json` identify the runnable revision; extracted text alone is insufficient to reconstruct form designers.

## Remaining engineering work

Future changes could introduce stable transaction IDs, an explicit goal selector and a configured currency code. These are not implemented. Keep computation rules in services, retain the button compatibility adapters, and regenerate source snapshots after saving a VBA change. See [testing](TESTING.md) for regression coverage and remaining manual checks.
