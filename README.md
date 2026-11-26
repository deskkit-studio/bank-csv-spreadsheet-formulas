# Bank CSV spreadsheet formulas

Plain spreadsheet formulas for cleaning up a bank CSV export: turn text amounts and dates into real numbers and dates, give each row a category from a keyword list, and keep transfers out of your income and spending totals.

No scripts, no macros, no LAMBDA. Every formula below was recalculated in **LibreOffice 26.8.3** against hand-written expected values (`test_formulas.py`). The functions used also exist in Excel and Google Sheets, but I haven't tested these formulas there.

Files:

- `README.md`: the formulas, explained.
- `sample-bank.csv`: 9 made-up rows (UK style: day/month/year dates, separate *Money out* / *Money in* columns, `£` amounts).
- `test_formulas.py`: builds `example.xlsx` from the sample, recalculates it in LibreOffice and checks each row's date, amount, rule score and category, the four totals, and the single cases listed in section 6.
- `example.xlsx`: the sample with all formulas filled in, as built and recalculated by `test_formulas.py` (LibreOffice). Running the test rebuilds it.

The formulas assume English-style settings: a **point** as the decimal separator and a comma between formula arguments. With `;` as the argument separator or a comma as the decimal separator they need changes, and I haven't tested that.

## 1. Paste or import as text first

If you open the CSV directly, the app may read `01/11/2026` (1 November) as January 11 using your computer's date order, and the cell value alone can't tell you afterwards which rows were swapped. Import the date and amount columns as **Text** instead. The formulas below expect text and leave already-numeric amounts alone. (Menus differ between apps; I didn't test the import step itself. In the test the values are entered as text directly.)

## 2. Amounts: `$1,234.56`, `(12.00)`, `-$5.00` → numbers

Raw amount in A2:

```
=IF(TRIM(A2)="",0,IF(ISNUMBER(A2),A2,IF(OR(ISNUMBER(SEARCH("(",A2)),ISNUMBER(SEARCH("-",A2))),-1,1)*VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"(",""),")",""),"-",""),",",""),"$",""),"£",""),"€",""))))
```

- Empty cell → `0`. Already a number → kept as is.
- `(` or `-` anywhere in the cell → negative (`(12.00)`, `-$5.00`, `$-5.00`).
- Removes `( ) - , $ £ €`, then `VALUE` makes it a number. Add a `SUBSTITUTE` layer for other symbols.
- It does **not** validate the text: `1-2` becomes -12, and leftover letters may give `#VALUE!` (not every bad input is rejected). Use it on amount columns only and compare a few rows with the original.

### Separate *Money out* (C2) and *Money in* (D2) columns → one signed amount

```
=ABS(IF(TRIM(D2)="",0,IF(ISNUMBER(D2),D2,IF(OR(ISNUMBER(SEARCH("(",D2)),ISNUMBER(SEARCH("-",D2))),-1,1)*VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(D2),"(",""),")",""),"-",""),",",""),"$",""),"£",""),"€","")))))-ABS(IF(TRIM(C2)="",0,IF(ISNUMBER(C2),C2,IF(OR(ISNUMBER(SEARCH("(",C2)),ISNUMBER(SEARCH("-",C2))),-1,1)*VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(C2),"(",""),")",""),"-",""),",",""),"$",""),"£",""),"€","")))))
```

That's `ABS(formula above on D2) - ABS(formula above on C2)`: money out becomes negative, money in positive, an empty row is 0. If a row has both columns filled you get the difference, so check those rows by hand.

## 3. Dates as text → real dates

Day/month/year (`29/10/2026`, `1/11/2026`):

```
=DATE(VALUE(RIGHT(TRIM(A2),4)),VALUE(MID(TRIM(A2),FIND("/",TRIM(A2))+1,FIND("/",TRIM(A2),FIND("/",TRIM(A2))+1)-FIND("/",TRIM(A2))-1)),VALUE(LEFT(TRIM(A2),FIND("/",TRIM(A2))-1)))
```

Month/day/year (`11/26/2026`):

```
=DATE(VALUE(RIGHT(TRIM(A2),4)),VALUE(LEFT(TRIM(A2),FIND("/",TRIM(A2))-1)),VALUE(MID(TRIM(A2),FIND("/",TRIM(A2))+1,FIND("/",TRIM(A2),FIND("/",TRIM(A2))+1)-FIND("/",TRIM(A2))-1)))
```

Year-month-day (`2026-11-01`):

```
=DATE(VALUE(LEFT(TRIM(A2),4)),VALUE(MID(TRIM(A2),6,2)),VALUE(MID(TRIM(A2),9,2)))
```

They split the text and build the date with `DATE(year, month, day)`, so your computer's date order doesn't matter. Format the result cells as dates. Input must be exactly `D/M/YYYY`, `M/D/YYYY` or `YYYY-MM-DD` with nothing after it; the formulas don't validate that, and `DATE` rolls impossible dates over (`31/02/2026` becomes 3 March). To find your bank's order, look for a row where the first or second number is above 12.

## 4. Categories from a keyword list (first match wins)

A sheet named **Rules**: keyword in column A, category in column B, header in row 1. Put more specific keywords higher (`UBER EATS` above `UBER`).

| Keyword | Category |
|---|---|
| NETFLIX | Subscriptions |
| TRANSFER TO SAV | Transfer |
| TRANSFER FROM SAV | Transfer |
| TESCO | Groceries |
| UBER EATS | Eating Out |
| UBER | Transport |
| PAYROLL | Salary |

With the description in B2, G2 = score of the topmost matching rule (0 = none):

```
=SUMPRODUCT(MAX((Rules!$A$2:$A$300<>"")*ISNUMBER(SEARCH(Rules!$A$2:$A$300,B2))*(1000-ROW(Rules!$A$2:$A$300))))
```

H2 = the category:

```
=IF(G2=0,"Uncategorized",INDEX(Rules!$B$1:$B$1000,1000-G2))
```

`SEARCH` finds the keyword anywhere in the description (not case-sensitive), `1000-ROW(...)` gives higher rules a bigger score so `MAX` picks the topmost match, and the outer `SUMPRODUCT` gets the array evaluated without Ctrl+Shift+Enter in LibreOffice. If Excel or Google Sheets shows a wrong number, enter G2 as an array formula (Ctrl+Shift+Enter in older Excel, `ARRAYFORMULA` in Sheets). As written it reads rules from rows 2 to 300 only. For more rules, say up to row 1500: change `$A$2:$A$300` to `$A$2:$A$1500` (three times in G2), both `1000`s (G2 and H2) to `1501`, and `$B$1:$B$1000` in H2 to `$B$1:$B$1500`. The number must stay bigger than the last rule row, and H2's range must reach it. Avoid `*`, `?` and `~` in keywords: `SEARCH` can treat them as wildcards.

## 5. Totals without transfers

Signed amount in column F, category in column H:

```
Income:    =SUMIFS($F:$F,$F:$F,">0",$H:$H,"<>Transfer")
Spending:  =-SUMIFS($F:$F,$F:$F,"<0",$H:$H,"<>Transfer")
Transfers: =SUMIFS($F:$F,$H:$H,"Transfer")
Not yet categorized: =COUNTIF($H:$H,"Uncategorized")
```

Put these outside columns F and H (the example uses J:K). `Spending` includes money-out rows that are still `Uncategorized`. A refund is money in, so it counts as income here, not as less spending. A credit-card payment is only a transfer **if** you also track that card's purchases as their own rows; otherwise those purchases disappear from spending. Banks word transfers differently, so add rules for your bank's wording.

## 6. Result on the sample

| Description | Money out | Money in | Date | Amount | Category |
|---|---|---|---|---|---|
| ACME LTD PAYROLL | | £2,100.00 | 2026-11-01 | 2100.00 | Salary |
| TESCO STORES 2231 | £45.20 | | 2026-11-02 | -45.20 | Groceries |
| UBER *TRIP | £12.40 | | 2026-11-03 | -12.40 | Transport |
| UBER EATS ORDER | £18.75 | | 2026-11-03 | -18.75 | Eating Out |
| NETFLIX.COM | £10.99 | | 2026-11-05 | -10.99 | Subscriptions |
| TRANSFER TO SAVINGS 1234 | £300.00 | | 2026-11-06 | -300.00 | Transfer |
| TRANSFER FROM SAVINGS 1234 | | £50.00 | 2026-11-13 | 50.00 | Transfer |
| CORNER BAKERY | £6.80 | | 2026-11-14 | -6.80 | Uncategorized |
| REFUND TESCO STORES 2231 | | £5.20 | 2026-11-29 | 5.20 | Groceries |

Income 2,105.20 · Spending 94.14 · Transfers -250.00 · 1 row left to categorize.

`test_formulas.py` also checks these single cases: `$1,234.56` → 1234.56, `(12.00)` → -12, `-$5.00` and `$-5.00` → -5, `£2,100.00` → 2100, `-1,000` → -1000, ` 7 ` → 7, `€0.99` → 0.99, empty → 0, `1-2` → -12, the number -3.5 → -3.5; money out / money in pairs `£45.20`/empty, `-45.20`/empty, `(45.20)`/empty → -45.20, empty/`£2,100.00` → 2100, empty/empty → 0, `10.00`/`3.00` → -7; `29/10/2026`, `1/11/2026`, ` 05/02/2026` (D/M/Y), `11/26/2026`, `2/3/2026` (M/D/Y), `2026-11-01`, and `31/02/2026` → 2026-03-03.

Run it yourself (Python 3, `pip install openpyxl`, LibreOffice; set `SOFFICE=/path/to/soffice` if it isn't found). It uses a fresh LibreOffice profile but not a fixed system locale; I ran it with English (point as decimal separator) settings, and with other number or date settings it may fail:

```
python3 test_formulas.py
```

## Limits

- Decimal point only (`1.234,56` is not supported). `VALUE` can depend on your spreadsheet's language setting.
- Amounts marked `CR` / `DR` instead of a sign need another rule.
- Dates with a time or two-digit years are not supported.
- Apart from the cases above I didn't test malformed input, cells with only spaces, or other language settings.
- Tested only in LibreOffice.

Longer walkthroughs of the same formulas: [categorizing with one formula](https://dev.to/deskkitstudio/categorize-bank-csv-transactions-in-a-spreadsheet-with-one-formula-first-matching-keyword-wins-4ap1) and [cleaning amounts and dates](https://dev.to/deskkitstudio/turn-messy-bank-csv-amounts-and-dates-into-real-numbers-with-spreadsheet-formulas-24c1).

## License

MIT (see `LICENSE`). The formulas are free to use.

## The ready-made version (disclosure: I sell it)

I also sell a template built around the same ideas: **BankPaste Budget** ($12, .xlsx). Its formulas are **not** the ones above; an amount it can't read becomes **0** instead of an error, so compare a few rows and the month's totals with your bank statement. You paste a bank CSV, type which columns are date / description / amount (one signed column or separate money-out / money-in), pick the date order, and it builds a monthly budget vs actual view with keyword rules you can edit. It has no starter rule for credit-card payments; you pick that category yourself. Limits: US/UK-style numbers (1,234.56), dates as M/D/Y, D/M/Y or Y-M-D, up to 1,000 rows and one account per paste. I tested it in LibreOffice only, with made-up sample CSVs, not with real exports from any bank. Check the requirements on the product page before buying: https://caremini.gumroad.com/l/bankpaste-budget
