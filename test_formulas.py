"""Builds example.xlsx from sample-bank.csv and checks every formula in the README
against hand-written expected values after a LibreOffice recalculation.

Needs: Python 3, openpyxl, LibreOffice (set SOFFICE=/path/to/soffice if it isn't found).
Run:   python3 test_formulas.py
"""
import csv, datetime, os, shutil, subprocess, sys, tempfile
import openpyxl

# --- the formulas, exactly as in README.md ({c} = the input cell) -------------
AMT = ('IF(TRIM({c})="",0,IF(ISNUMBER({c}),{c},IF(OR(ISNUMBER(SEARCH("(",{c})),ISNUMBER(SEARCH("-",{c}))),-1,1)'
       '*VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM({c}),"(",""),")",""),"-",""),",",""),"$",""),"£",""),"€",""))))')
SPLIT = 'ABS(' + AMT.replace('{c}', '{inn}') + ')-ABS(' + AMT.replace('{c}', '{out}') + ')'
DMY = ('DATE(VALUE(RIGHT(TRIM({c}),4)),VALUE(MID(TRIM({c}),FIND("/",TRIM({c}))+1,FIND("/",TRIM({c}),FIND("/",TRIM({c}))+1)-FIND("/",TRIM({c}))-1)),'
       'VALUE(LEFT(TRIM({c}),FIND("/",TRIM({c}))-1)))')
MDY = ('DATE(VALUE(RIGHT(TRIM({c}),4)),VALUE(LEFT(TRIM({c}),FIND("/",TRIM({c}))-1)),'
       'VALUE(MID(TRIM({c}),FIND("/",TRIM({c}))+1,FIND("/",TRIM({c}),FIND("/",TRIM({c}))+1)-FIND("/",TRIM({c}))-1)))')
ISO = 'DATE(VALUE(LEFT(TRIM({c}),4)),VALUE(MID(TRIM({c}),6,2)),VALUE(MID(TRIM({c}),9,2)))'
SCORE = 'SUMPRODUCT(MAX((Rules!$A$2:$A$300<>"")*ISNUMBER(SEARCH(Rules!$A$2:$A$300,{c}))*(1000-ROW(Rules!$A$2:$A$300))))'
CAT = 'IF({s}=0,"Uncategorized",INDEX(Rules!$B$1:$B$1000,1000-{s}))'
# section 7: a money-out row whose exact description and amount also appear on a row at least 25 days away
REPEAT = ('IF({amt}>=0,"",IF(SUMPRODUCT(($B$2:$B$1000={desc})*($F$2:$F$1000={amt})*(ABS($E$2:$E$1000-{date})>=25))>0,'
          '"Repeats",""))')

RULES = [("NETFLIX", "Subscriptions"), ("TRANSFER TO SAV", "Transfer"), ("TRANSFER FROM SAV", "Transfer"),
         ("TESCO", "Groceries"), ("UBER EATS", "Eating Out"), ("UBER", "Transport"), ("PAYROLL", "Salary")]

# --- hand-written expected values (not computed from the formulas) -----------
EXPECT_ROWS = [  # date, signed amount, category  (same order as sample-bank.csv)
    ((2026, 11, 1), 2100.00, "Salary"),
    ((2026, 11, 2), -45.20, "Groceries"),
    ((2026, 11, 3), -12.40, "Transport"),
    ((2026, 11, 3), -18.75, "Eating Out"),
    ((2026, 11, 5), -10.99, "Subscriptions"),
    ((2026, 11, 6), -300.00, "Transfer"),
    ((2026, 11, 13), 50.00, "Transfer"),
    ((2026, 11, 14), -6.80, "Uncategorized"),
    ((2026, 11, 29), 5.20, "Groceries"),
]
EXPECT_TOTALS = {"Income": 2105.20, "Spending": 94.14, "Transfers": -250.00, "Uncategorized rows": 1}
AMT_CASES = [("$1,234.56", 1234.56), ("(12.00)", -12), ("-$5.00", -5), ("$-5.00", -5), ("£2,100.00", 2100),
             ("45.2", 45.2), (" 7 ", 7), ("", 0), ("€0.99", 0.99), ("-1,000", -1000), ("1-2", -12)]
DATE_CASES = [(DMY, "29/10/2026", (2026, 10, 29)), (DMY, "1/11/2026", (2026, 11, 1)), (DMY, " 05/02/2026", (2026, 2, 5)),
              (MDY, "11/26/2026", (2026, 11, 26)), (MDY, "2/3/2026", (2026, 2, 3)), (ISO, "2026-11-01", (2026, 11, 1)),
              (DMY, "31/02/2026", (2026, 3, 3))]
SPLIT_CASES = [("£45.20", "", -45.20), ("-45.20", "", -45.20), ("(45.20)", "", -45.20), ("", "£2,100.00", 2100),
               ("", "", 0), ("10.00", "3.00", -7)]  # money out, money in, expected; last = both filled
REPEAT_ROWS = [  # date, description, amount, expected flag (hand-written)
    ((2026, 9, 5), "NETFLIX.COM", -10.99, "Repeats"),
    ((2026, 10, 5), "NETFLIX.COM", -10.99, "Repeats"),
    ((2026, 11, 5), "netflix.com", -10.99, "Repeats"),   # case doesn't matter
    ((2026, 11, 6), "NETFLIX.COM ", -10.99, ""),         # trailing space: a different description
    ((2026, 10, 2), "TESCO STORES 2231", -45.20, ""),    # same shop, different amount
    ((2026, 11, 2), "TESCO STORES 2231", -38.10, ""),
    ((2026, 10, 3), "UBER *TRIP", -12.40, ""),           # * is not a wildcard here
    ((2026, 11, 3), "UBER X TRIP", -12.40, ""),
    ((2026, 11, 10), "GYM", -30.00, ""),                 # same twice, only 1 day apart
    ((2026, 11, 11), "GYM", -30.00, ""),
    ((2026, 10, 1), "ACME LTD PAYROLL", 2100.00, ""),    # money in is never flagged
    ((2026, 11, 1), "ACME LTD PAYROLL", 2100.00, ""),
    ((2025, 11, 20), "DOMAIN RENEWAL", -12.00, "Repeats"),  # yearly
    ((2026, 11, 20), "DOMAIN RENEWAL", -12.00, "Repeats"),
    ((2026, 10, 15), "SPOTIFY", -9.99, "Repeats"),       # exactly 25 days apart
    ((2026, 11, 9), "SPOTIFY", -9.99, "Repeats"),
    ((2026, 10, 16), "COSTA COFFEE", -3.20, ""),         # 24 days apart
    ((2026, 11, 9), "COSTA COFFEE", -3.20, ""),
]
EXPECT_SCORES = [992, 995, 993, 994, 998, 997, 996, 0, 995]  # 1000 - sheet row of the matching rule (NETFLIX is row 2), 0 = none

here = os.path.dirname(os.path.abspath(__file__))


def build(path):
    wb = openpyxl.Workbook()
    tx = wb.active
    tx.title = "Transactions"
    ru = wb.create_sheet("Rules")
    ru.append(["Keyword", "Category"])
    for r in RULES:
        ru.append(list(r))
    tx.append(["Date (text)", "Description", "Money out", "Money in", "Date", "Amount", "Rule score", "Category", "Repeat?"])
    with open(os.path.join(here, "sample-bank.csv"), newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))[1:]
    for i, (d, desc, out, inn) in enumerate(rows, start=2):
        for col, val in zip("ABCD", (d, desc, out, inn)):
            tx[f"{col}{i}"] = val
            tx[f"{col}{i}"].data_type = "s"  # keep as text, like an import with Text columns
        tx[f"E{i}"] = "=" + DMY.format(c=f"A{i}")
        tx[f"E{i}"].number_format = "yyyy-mm-dd"
        tx[f"F{i}"] = "=" + SPLIT.format(out=f"C{i}", inn=f"D{i}")
        tx[f"G{i}"] = "=" + SCORE.format(c=f"B{i}")
        tx[f"H{i}"] = "=" + CAT.format(s=f"G{i}")
        tx[f"I{i}"] = "=" + REPEAT.format(desc=f"B{i}", date=f"E{i}", amt=f"F{i}")
    for r, (label, f) in enumerate([
            ("Income", '=SUMIFS($F:$F,$F:$F,">0",$H:$H,"<>Transfer")'),
            ("Spending", '=-SUMIFS($F:$F,$F:$F,"<0",$H:$H,"<>Transfer")'),
            ("Transfers", '=SUMIFS($F:$F,$H:$H,"Transfer")'),
            ("Uncategorized rows", '=COUNTIF($H:$H,"Uncategorized")')], start=1):
        tx[f"J{r}"], tx[f"K{r}"] = label, f
    t = wb.create_sheet("Formula tests")
    t.append(["Input", "Result", "Expected"])
    for raw, exp in AMT_CASES:
        n = t.max_row + 1
        t[f"A{n}"] = raw
        t[f"A{n}"].data_type = "s"
        t[f"B{n}"], t[f"C{n}"] = "=" + AMT.format(c=f"A{n}"), exp
    for out_, in_, exp in SPLIT_CASES:
        n = t.max_row + 1
        t[f"A{n}"], t[f"D{n}"] = out_, in_
        t[f"A{n}"].data_type = t[f"D{n}"].data_type = "s"
        t[f"B{n}"], t[f"C{n}"] = "=" + SPLIT.format(out=f"A{n}", inn=f"D{n}"), exp
    n = t.max_row + 1  # an amount that is already a number stays as it is
    t[f"A{n}"], t[f"B{n}"], t[f"C{n}"] = -3.5, "=" + AMT.format(c=f"A{n}"), -3.5
    for f, raw, exp in DATE_CASES:
        n = t.max_row + 1
        t[f"A{n}"], t[f"B{n}"], t[f"C{n}"] = raw, "=" + f.format(c=f"A{n}"), datetime.date(*exp)
        t[f"B{n}"].number_format = t[f"C{n}"].number_format = "yyyy-mm-dd"
    rp = wb.create_sheet("Repeats")  # same columns as Transactions: B description, E date, F amount
    for col, h in zip("BEFI", ("Description", "Date", "Amount", "Repeat?")):
        rp[f"{col}1"] = h
    for i, (d, desc, amt, _) in enumerate(REPEAT_ROWS, start=2):
        rp[f"B{i}"], rp[f"E{i}"], rp[f"F{i}"] = desc, datetime.date(*d), amt
        rp[f"E{i}"].number_format = "yyyy-mm-dd"
        rp[f"I{i}"] = "=" + REPEAT.format(desc=f"B{i}", date=f"E{i}", amt=f"F{i}")
    # one error value in E or F: every money-out row shows an error, money-in rows stay empty (README section 7)
    for name, data in (("Repeats error E", [((2026, 10, 5), "NETFLIX.COM", -10.99), ((2026, 11, 5), "NETFLIX.COM", -10.99),
                                            (None, "GYM", -30.00), ((2026, 11, 1), "ACME LTD PAYROLL", 2100.00)]),
                       ("Repeats error F", [((2026, 10, 5), "NETFLIX.COM", -10.99), ((2026, 11, 5), "GYM", None),
                                            ((2026, 11, 1), "ACME LTD PAYROLL", 2100.00)])):
        er = wb.create_sheet(name)
        for i, (d, desc, amt) in enumerate(data, start=2):
            er[f"B{i}"] = desc
            er[f"E{i}"] = datetime.date(*d) if d else "=1/0"
            er[f"F{i}"] = amt if amt is not None else "=1/0"
            er[f"I{i}"] = "=" + REPEAT.format(desc=f"B{i}", date=f"E{i}", amt=f"F{i}")
    wb.save(path)
    return len(rows)


def soffice():
    for p in (os.environ.get("SOFFICE"), shutil.which("soffice"), shutil.which("libreoffice"),
              os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if p and os.path.exists(p):
            return p
    sys.exit("LibreOffice not found: set SOFFICE=/path/to/soffice")


def main():
    example = os.path.join(here, "example.xlsx")
    n = build(example)
    with tempfile.TemporaryDirectory() as tmp:  # own LibreOffice profile, so local settings don't leak in
        profile = "file://" + os.path.join(tmp, "profile").replace(os.sep, "/")
        subprocess.run([soffice(), f"-env:UserInstallation={profile}", "--headless", "--convert-to", "xlsx",
                        "--outdir", os.path.join(tmp, "out"), example], check=True, capture_output=True)
        shutil.copy(os.path.join(tmp, "out", "example.xlsx"), example)  # keep the recalculated copy
    wb = openpyxl.load_workbook(example, data_only=True)
    bad = 0

    def check(got, exp, label):
        nonlocal bad
        if isinstance(exp, tuple):
            ok = isinstance(got, datetime.datetime) and (got.year, got.month, got.day) == exp
        elif isinstance(exp, str):
            ok = got == exp
        else:
            ok = isinstance(got, (int, float)) and abs(got - exp) < 1e-9
        print("OK  " if ok else "FAIL", label, "->", got)
        bad += not ok

    tx = wb["Transactions"]
    assert n == len(EXPECT_ROWS)
    for i, (d, a, c) in enumerate(EXPECT_ROWS, start=2):
        desc = tx[f"B{i}"].value
        check(tx[f"E{i}"].value, d, f"date   {desc}")
        check(tx[f"F{i}"].value, a, f"amount {desc}")
        check(tx[f"G{i}"].value, EXPECT_SCORES[i - 2], f"score  {desc}")
        check(tx[f"H{i}"].value, c, f"cat    {desc}")
    for r in range(1, 5):
        check(tx[f"K{r}"].value, EXPECT_TOTALS[tx[f"J{r}"].value], tx[f"J{r}"].value)
    t = wb["Formula tests"]
    for r in range(2, t.max_row + 1):
        exp = t[f"C{r}"].value
        exp = (exp.year, exp.month, exp.day) if isinstance(exp, datetime.datetime) else exp
        label = f"case {t[f'A{r}'].value!r}" + (f" / in {t[f'D{r}'].value!r}" if t[f"D{r}"].value is not None else "")
        check(t[f"B{r}"].value, exp, label)
    rp = wb["Repeats"]
    for i, (d, desc, amt, exp) in enumerate(REPEAT_ROWS, start=2):
        got = rp[f"I{i}"].value
        check("" if got is None else got, exp, f"repeat {d} {desc!r} {amt}")
    for name, exp_err in (("Repeats error E", [True, True, True, False]), ("Repeats error F", [True, True, False])):
        er = wb[name]
        for i, e in enumerate(exp_err, start=2):
            got = er[f"I{i}"].value
            if e:
                check(isinstance(got, str) and got.startswith("#"), True, f"{name} row {i} -> error value ({got})")
            else:
                check("" if got is None else got, "", f"{name} row {i} (money in) -> empty")
    for i in range(2, 2 + len(EXPECT_ROWS)):  # section 7 on the sample too: no row repeats in one month
        got = tx[f"I{i}"].value
        check("" if got is None else got, "", f"repeat on sample {tx[f'B{i}'].value}")
    print("ALL PASS" if not bad else f"{bad} FAILED")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
