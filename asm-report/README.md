# asm-report

Everything for the ASM-wise report in one place. Run the commands from the repo root.

| Path | What it is |
|---|---|
| `asmwise_report.py` | Downloads the DMS exports and builds the summary (`npm run report:asmwise`) |
| `asm_email.py` | Builds one email per ASM mail from the latest summary as a preview (`npm run report:asmwise:email`); `--send` mails them, `--only <address>` limits it to those ASMs, `--test-to <address>` sends them all to one address instead |
| `run_asmwise_report.bat` | Local run: builds the summary, then sends the emails if that worked. `--send` refuses a summary that isn't for yesterday |
| `../.github/workflows/asmwise-report.yml` | The schedule: GitHub Actions runs the same two steps daily at 07:47, 08:07, 08:22 and 08:37 IST, retrying the build up to 3 times, and the Google trigger in `resources/apps-script/triggerAsmReport.gs` starts it between 07:00 and 08:00 IST; scheduled runs mail only between 08:00 and 09:00 IST (waiting for 08:00 if ready early, never after 09:00) and only if today's email hasn't gone yet. Run it by hand with `mode` = `test` (all emails to `GMAIL_TO`), `preview`, `send` or `auto` (send unless already sent today). Needs the secrets `CREDENTIALS_JSON` (with `riverUserValidUser`), `GMAIL_USER`, `GMAIL_APP_PASSWORD`, `GMAIL_TO`; optional `ASM_MAIL_FROM` |
| `requirements.txt` | `pip install -r asm-report/requirements.txt` |
| `Mapping/` | `ASM_Name_Mapping.xlsx`: ASM, ASM Mail ID (Store Name labels the dealers). Report ASMs are matched by name with regular expressions: the mapping name anywhere in the DMS name ("Sai Ram" → "Vempati Sai Ram"), only at its start when it fits several ("Akram" → "Akram Shaikh"), else a mail ID that starts the name ("danis@" → "Danis Davis Thommana"). Report ASMs with no match get no email and are listed in the log and `index.html` |
| `../Output/Downloads/ASM/` | The six DMS exports; emptied after each run unless `--keep-downloads` (not committed) |
| `../Output/ASM/` | `ASMwise_AllSources_Summary_<yesterday>.xlsx`, `email_preview/<yesterday>/`, `logs/asmwise_report.log` (not committed) |

Login comes from `resources/credentials.json` (`riverUserValidUser`), shared with the rest of the repo.
