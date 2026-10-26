"""Builds one email per ASM from the latest ASMwise summary (a preview; --send mails them).

Reads the newest Output/ASM/ASMwise_AllSources_Summary_*.xlsx and
asm-report/Mapping/ASM_Name_Mapping.xlsx (ASM, ASM Mail ID; Store Name for the tables).
Report ASMs are matched to the mapping by name with regular expressions (see match_asms), and
each match gets one email to that ASM's mail IDs - their overall totals and a dealer-wise table
per period in the body (no attachment). Report ASMs with no match are listed in index.html and the log.
DISPLAY_NAMES renames a DMS ASM in its email ("Akram Shaikh" is greeted as "Mohammed Akram Khan").

Writes Output/ASM/email_preview/<date>/:
    <ASM>.eml       the complete email (open it in Outlook to see it as received)
    <ASM>.html      the body only, for a quick look in the browser
    index.html      who gets what, and the report ASMs with no match in the mapping

Usage (from the repo root):
    python asm-report/asm_email.py [--summary PATH] [--mapping PATH] [--only ADDRESS ...]
                                       [--dealer CODE ...] [--send] [--test-to ADDRESS]
                                       [--also-to ADDRESS ...] [--asm NAME ...] [--cc ADDRESS ...]

--send mails them through Gmail (GMAIL_USER / GMAIL_APP_PASSWORD from .env), and only when the
summary is for yesterday unless --summary is given. --only limits it to those ASM addresses,
--asm to those ASM names (matched like the mapping, e.g. "Akram") and --dealer to those dealer
codes; --test-to sends every email to ADDRESS instead of the ASMs (subject marked [TEST], no Cc);
--also-to adds those addresses to every email and --cc copies them on every email.

Environment: ASM_MAIL_FROM (From address; defaults to GMAIL_USER).
"""
from __future__ import annotations

import argparse
import html
import os
import re
import smtplib
import sys
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from email.utils import formatdate
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from asmwise_report import (HERE, METRICS, OUTPUT_DIR, PERCENT_METRICS, SOURCE_LABEL, SOURCES,  # noqa: E402
                            SUMMARY_NAME, TOTAL, clean, col, column, columns, log, norm_dealer,
                            overall, periods)

MAPPING_FILE = HERE / "Mapping" / "ASM_Name_Mapping.xlsx"
SUMMARY_GLOB = f"{SUMMARY_NAME}_*.xlsx"
ENV_FILE = HERE.parent / ".env"
# DMS ASM name -> the name the email uses (greeting, subject, file name) where the DMS has it wrong.
DISPLAY_NAMES = {"Akram Shaikh": "Mohammed Akram Khan"}


def load_env(path: Path = ENV_FILE) -> None:
    """KEY=value lines from the repo .env, without overriding variables already set."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        key, sep, value = line.partition("=")
        if sep and not key.strip().startswith("#"):
            os.environ.setdefault(key.strip(), value.strip().strip('"\''))


load_env()
MAIL_FROM = os.environ.get("ASM_MAIL_FROM") or os.environ.get("GMAIL_USER", "")


def latest_summary() -> Path:
    found = sorted(OUTPUT_DIR.glob(SUMMARY_GLOB), key=lambda f: f.stat().st_mtime)
    if not found:
        raise SystemExit(f"No {SUMMARY_GLOB} in {OUTPUT_DIR}. Run asm-report/asmwise_report.py first.")
    return found[-1]


def report_date(path: Path) -> datetime:
    m = re.search(r"_(\d{2}-\d{2}-\d{4})", path.stem)
    if not m:
        raise SystemExit(f"No dd-mm-yyyy date in {path.name}")
    return datetime.strptime(m.group(1), "%d-%m-%Y")


def read_summary(path: Path) -> pd.DataFrame:
    """The Master Summary rows, without All Dealers; value columns named "<source group> <period>"."""
    raw = pd.read_excel(path, sheet_name="Master Summary", header=None, dtype=str)
    groups, period_row = raw.iloc[2].ffill(), raw.iloc[3]  # two header rows, groups merged
    df = raw.iloc[4:].copy()
    df.columns = [g if i < 3 else column(g, p) for i, (g, p) in enumerate(zip(groups, period_row))]
    df[["Dealer Code", "ASM"]] = df[["Dealer Code", "ASM"]].ffill()  # merged cells
    df = df[df["Dealer Code"] != "All Dealers"].copy()
    labels = [c for c in df.columns if c not in ("Dealer Code", "ASM", "Metric")]
    df[labels] = df[labels].apply(pd.to_numeric, errors="coerce")
    df["ASM"] = df["ASM"].map(clean)
    return df.reset_index(drop=True)


def read_mapping(path: Path) -> tuple[dict[str, list[str]], dict[str, str]]:
    """ASM name -> its distinct mail addresses, and Dealer Code -> Store Name (for the tables)."""
    df = pd.read_excel(path, dtype=str).fillna("")
    mails: dict[str, list[str]] = {}
    for name, cell_mails in zip(df[col(df, "ASM", "ASM Name")].map(clean), df[col(df, "ASM Mail ID", "ASM mail")]):
        for addr in re.split(r"[;,\s]+", cell_mails.strip()):
            if name and "@" in addr and addr not in mails.setdefault(name, []):
                mails[name].append(addr)
    stores = dict(zip(df[col(df, "Dealer Code")].map(norm_dealer), df[col(df, "Store Name", "Dealer Name")].map(clean)))
    return mails, {code: name for code, name in stores.items() if code and name}


def name_pattern(name: str) -> re.Pattern:
    # Anywhere in the name, starting at a word: "Sai Ram" matches "Vempati Sai Ram" and "Narendra
    # Raju" matches "Narendra Raju Rayana", but "Ram" doesn't match "Sriram".
    words = r"[\s_]+".join(map(re.escape, name.split()))
    return re.compile(rf"(?<![A-Za-z0-9]){words}", re.I)


def letters(s: str) -> str:
    return re.sub(r"[^a-z]", "", s.casefold())


def match_asms(report_asms: list[str], mapping: dict[str, list[str]]) -> dict[str, tuple[str, tuple[str, ...]]]:
    """Report ASM -> (mapping ASM name, its mails).

    1. The mapping name, exact or as a regex match inside the report name (the longest name wins).
       A mapping name that matches several report ASMs keeps only those it starts (so "Akram"
       goes to "Akram Shaikh", not "MOHAMMED AKRAM KHAN").
    2. Else a mail whose user part starts the report name ("sudhirdeshpande@" -> "Sudhirdeshpande",
       "danis@" -> "Danis Davis Thommana"); the mail itself is then the recipient.
    """
    patterns = sorted(((n, name_pattern(n)) for n in mapping), key=lambda x: -len(x[0]))
    hits: dict[str, tuple[str, int]] = {}
    for asm in report_asms:
        found = next(((n, m.start()) for n, p in patterns if (m := p.search(asm))), None)
        if found:
            hits[asm] = found
    for name in {n for n, _ in hits.values()}:
        mine = [a for a, (n, _) in hits.items() if n == name]
        if len(mine) > 1 and any(hits[a][1] == 0 for a in mine):
            for a in mine:
                if hits[a][1] > 0:
                    del hits[a]
    out = {a: (n, tuple(mapping[n])) for a, (n, _) in hits.items()}

    users = sorted({m for ms in mapping.values() for m in ms}, key=lambda m: -len(letters(m.split("@")[0])))
    for asm in report_asms:
        if asm not in out:
            mail = next((m for m in users if len(letters(m.split("@")[0])) >= 4
                         and letters(asm).startswith(letters(m.split("@")[0]))), None)
            if mail:
                out[asm] = (mail, (mail,))
    return out


def fmt(metric: str, v) -> str:
    if pd.isna(v):
        return "-"
    return f"{v:.1f}" if metric in PERCENT_METRICS else f"{int(v):,}"


PERIOD_NAMES = {"YTD": "Year to Date", "MTD": "Month to Date", "Yesterday": "Yesterday"}
TH = "background:#1F4E78;color:#fff;padding:6px 8px;border:1px solid #BFBFBF;text-align:center;font-weight:bold"
TD = "padding:5px 8px;border:1px solid #BFBFBF"


def cell(text: str, style: str = "", tag: str = "td", attrs: str = "") -> str:
    base = TH if tag == "th" else TD
    return f'<{tag} {attrs}style="{base};{style}">{text}</{tag}>'


def num_cell(metric: str, v, style: str = "") -> str:
    pct = "font-style:italic;background:#F2F7FC;" if metric in PERCENT_METRICS else ""
    return cell(fmt(metric, v), f"text-align:right;{pct}{style}")


def dealer_table(rows: pd.DataFrame, total: pd.DataFrame, label: str, names: dict[str, str]) -> str:
    """One period: a row per dealer, a column per metric, the ASM total at the bottom."""
    wide = rows.pivot_table(index="Dealer Code", columns="Metric", values=label, aggfunc="first", dropna=False)
    head = cell("Dealer", tag="th") + "".join(cell(html.escape(m), tag="th") for m in METRICS)
    trs = []
    for i, (code, r) in enumerate(wide.sort_index().iterrows()):
        shade = "background:#F7F7F7;" if i % 2 else ""
        name = names.get(code, "")
        dealer = f"<b>{html.escape(code)}</b>" + (f"<br><span style=\"color:#595959\">{html.escape(name)}</span>"
                                                  if name else "")
        trs.append(f'<tr style="{shade}">' + cell(dealer, "white-space:nowrap")
                   + "".join(num_cell(m, r.get(m)) for m in METRICS) + "</tr>")
    t = total.set_index("Metric")[label]
    trs.append('<tr style="background:#DDEBF7;font-weight:bold">' + cell("Total")
               + "".join(num_cell(m, t.get(m), "font-weight:bold") for m in METRICS) + "</tr>")
    return (f'<table style="border-collapse:collapse;font-size:10pt;margin-bottom:18px">'
            f"<tr>{head}</tr>{''.join(trs)}</table>")


def body_html(asm: str, total: pd.DataFrame, rows: pd.DataFrame, plist: list, report_day: datetime,
              names: dict[str, str]) -> str:
    labels = columns(plist)
    overview = "".join(
        "<tr>" + cell(html.escape(m), "font-style:italic;white-space:nowrap" if m in PERCENT_METRICS
                      else "white-space:nowrap")
        + "".join(num_cell(m, r[c]) for c in labels) + "</tr>"
        for m, r in total.set_index("Metric").iterrows())
    group_head = (cell("Metric", "vertical-align:middle", tag="th", attrs='rowspan="2" ')
                  + "".join(cell(html.escape(g), tag="th", attrs=f'colspan="{len(plist)}" ') for g, _ in SOURCES))
    period_head = "".join(cell(p.label, tag="th") for _ in SOURCES for p in plist)
    sections = "".join(
        f'<h3 style="color:#1F4E78;margin:18px 0 6px">{PERIOD_NAMES.get(p.label, p.label)}'
        f' <span style="font-weight:normal;color:#595959;font-size:10pt">'
        f"({p.start:%d-%b-%Y} to {p.end:%d-%b-%Y}, all sources)</span></h3>"
        + dealer_table(rows, total, column(TOTAL, p.label), names)
        for p in plist)
    n = rows["Dealer Code"].nunique()
    return f"""<html><body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222">
<p>Hi {html.escape(asm)},</p>
<p>Here is your daily sales funnel update &ndash;
<b>Enquiry &rarr; Test Ride &rarr; Booking &rarr; Retail</b> &ndash; as of <b>{report_day:%d %b %Y}</b>,
covering your {n} dealer{'s' if n != 1 else ''}.</p>
<p>The first table shows your overall position for Yesterday, Month to Date and Year to Date,
in total and by lead source. The tables after it break the totals down dealer by dealer, so you
can see where each store stands and where follow-up is needed.</p>
<h3 style="color:#1F4E78;margin:18px 0 6px">Your overall position</h3>
<table style="border-collapse:collapse;font-size:10pt">
<tr>{group_head}</tr>
<tr>{period_head}</tr>
{overview}
</table>
{sections}
<p style="color:#595959;font-size:9.5pt">Conversion percentages are worked out from the counts in each period;
&ldquo;-&rdquo; means there was nothing to convert from.</p>
<p>Regards,<br>River DMS Reports</p>
</body></html>"""


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--summary", type=Path, help="summary workbook (default: newest in Output)")
    ap.add_argument("--mapping", type=Path, default=MAPPING_FILE, help="ASM mapping workbook")
    ap.add_argument("--only", nargs="+", metavar="ADDRESS", help="only these ASM addresses")
    ap.add_argument("--dealer", nargs="+", metavar="CODE", help="only these dealer codes")
    ap.add_argument("--send", action="store_true", help="send the emails through Gmail")
    ap.add_argument("--test-to", metavar="ADDRESS", help="send every email to ADDRESS instead of the ASMs")
    ap.add_argument("--also-to", nargs="+", metavar="ADDRESS", default=[], help="add these addresses to every email")
    ap.add_argument("--asm", nargs="+", metavar="NAME", help="only these ASMs (by name)")
    ap.add_argument("--cc", nargs="+", metavar="ADDRESS", default=[], help="Cc these addresses on every email")
    args = ap.parse_args()

    summary_path = args.summary or latest_summary()
    yesterday = report_date(summary_path)
    log(f"Summary: {summary_path.name}")
    # A failed morning download must not resend an old summary to the ASMs.
    if args.send and not args.summary and yesterday.date() != date.today() - timedelta(days=1):
        raise SystemExit(f"Not sending: the latest summary is for {yesterday:%d-%m-%Y}, not yesterday.")
    plist = periods(yesterday.date() + timedelta(days=1))
    labels = columns(plist)

    summary = read_summary(summary_path)
    if args.dealer:
        codes = {norm_dealer(d) for d in args.dealer}
        summary = summary[summary["Dealer Code"].map(norm_dealer).isin(codes)]
        if summary.empty:
            raise SystemExit(f"No dealer {', '.join(sorted(codes))} in {summary_path.name}")
    mapping, names = read_mapping(args.mapping)
    report_asms = sorted(summary["ASM"].unique())
    matched = match_asms(report_asms, mapping)
    # One email per set of addresses: report ASMs that reach the same mails ("Vempati Sai Ram" by name,
    # "Vempati Sairam" by mail) share one email, labelled with the first mapping name.
    by_mails: dict[tuple[str, ...], tuple[str, list[str]]] = {}
    for asm, (label, mails) in sorted(matched.items()):
        by_mails.setdefault(mails, (label, []))[1].append(asm)
    recipients = {(label, mails): asms for mails, (label, asms) in by_mails.items()}
    if args.only:
        only = {a.casefold() for a in args.only}
        recipients = {(label, tuple(m for m in mails if m.casefold() in only)): asms
                      for (label, mails), asms in recipients.items() if any(m.casefold() in only for m in mails)}
        if not recipients:
            raise SystemExit(f"No report ASM is mapped to the address(es) {', '.join(args.only)}")
    if args.asm:
        wanted = [name_pattern(n) for n in args.asm]
        recipients = {k: asms for k, asms in recipients.items()
                      if any(p.search(n) for p in wanted
                             for n in (k[0], *asms, *(DISPLAY_NAMES.get(a, a) for a in asms)))}
        if not recipients:
            raise SystemExit(f"No report ASM matches the name(s) {', '.join(args.asm)}")

    smtp = None
    if args.send:
        user, password = os.environ.get("GMAIL_USER"), os.environ.get("GMAIL_APP_PASSWORD")
        if not user or not password:
            raise SystemExit("--send needs GMAIL_USER and GMAIL_APP_PASSWORD in .env")
        smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        smtp.login(user, password)

    out_dir = OUTPUT_DIR / "email_preview" / f"{yesterday:%d-%m-%Y}"
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("*.*"):  # a rerun replaces the day's preview
        old.unlink()
    index_rows = []
    for (label, mails), asms in sorted(recipients.items()):
        to = ", ".join(mails)
        # The DMS name, never a mail address ("Vempati Sai Ram" for both spellings), unless renamed
        display = DISPLAY_NAMES.get(asms[0], asms[0])
        rows = summary[summary["ASM"].isin(asms)]
        total = overall(rows, labels)
        slug = re.sub(r"\W+", "_", display).strip("_")

        body = body_html(display, total, rows, plist, yesterday, names)
        msg = EmailMessage()
        msg["From"] = MAIL_FROM or "(set ASM_MAIL_FROM)"
        msg["To"] = ", ".join([args.test_to or to, *args.also_to])
        if args.cc and not args.test_to:
            msg["Cc"] = ", ".join(args.cc)
        msg["Subject"] = (f"{'[TEST] ' if args.test_to else ''}ASM-wise {SOURCE_LABEL} Summary - "
                          f"{display} - {yesterday:%d-%b-%Y}")
        msg["Date"] = formatdate(localtime=True)
        if not smtp:
            msg["X-Unsent"] = "1"  # Outlook opens the .eml as a draft
        msg.set_content(f"Hi {display},\n\nPlease see the HTML version of this email.\n")
        msg.add_alternative(body, subtype="html")
        (out_dir / f"{slug}.eml").write_bytes(bytes(msg))
        (out_dir / f"{slug}.html").write_text(body, encoding="utf-8")
        if smtp:
            smtp.send_message(msg)
        n = rows["Dealer Code"].nunique()
        cc = f"  cc {msg['Cc']}" if msg["Cc"] else ""
        log(f"{', '.join(asms)} [{label}] <{to}> -> {msg['To']}{cc}  ({n} dealers)" + ("  SENT" if smtp else ""))
        index_rows.append((label, to, ", ".join(asms), n, msg["Subject"], f"{slug}.eml"))

    if smtp:
        smtp.quit()

    unmatched = [] if args.only else [a for a in report_asms if a not in matched]
    for a in unmatched:
        log(f"[WARN] No ASM in the mapping matches report ASM '{a}'")

    rows_html = "".join(
        "<tr>" + "".join(f"<td>{html.escape(str(v))}</td>" for v in r[:5])
        + f'<td><a href="{r[5]}">{r[5]}</a></td></tr>' for r in index_rows)
    (out_dir / "index.html").write_text(f"""<html><body style="font-family:Calibri,Arial,sans-serif">
<h3>Email preview - {html.escape(summary_path.name)}</h3>
<table border="1" cellpadding="5" style="border-collapse:collapse">
<tr><th>Mapping ASM / mail</th><th>To</th><th>Report ASM</th><th>Dealers</th><th>Subject</th><th>Email</th></tr>
{rows_html}</table>
<h4>Report ASMs with no match in the mapping ({len(unmatched)})</h4>
<p>{html.escape(', '.join(unmatched)) or 'None'}</p>
</body></html>""", encoding="utf-8")
    log(f"Preview: {out_dir / 'index.html'}" + ("" if smtp else "  (nothing was sent)"))


if __name__ == "__main__":
    main()
