r"""
Downloads the GNC chat "Ticket" export from KwikEngage for a date range (default: yesterday):
  1. https://app.kwikengage.ai/login  (login from D:\MIS\GNC\Automate\chat.env)
  2. Analytics > Export Reports > e-mail + date range + type "Ticket" > Export
  3. KwikEngage mails a link (TicketBulkExport, https://media.kwikengage.ai/exports/tickets/....csv) to the e-mail
     address - found by read-only IMAP in the same Gmail login the other jobs use (email.env.txt) - and the csv is downloaded
Saves  <out>\ChatTickets_<from>_<to>.csv  and writes <out>\_result.json. Exit code 1 on any failure (page saved to
<logs>\inbound-debug\ as png + html).
    python fetch_chat.py --from 2026-10-11 --to 2026-10-11 --out "D:\MIS\GNC\Automate\Oct'26\Data\Chat\2026-10-11"
    options: --show (visible browser)   --mail-timeout 900   --link <url> (skip the export, just download this link)
"""
import argparse, email, imaplib, json, re, sys, time, traceback
from datetime import date, datetime, timedelta, timezone
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from pathlib import Path

import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

ENV_FILE = Path(r"D:\MIS\GNC\Automate\chat.env")
GMAIL_ENV = [Path("D:\\MIS\\Bella Vitta\\Oct'26\\automate") / n for n in ("gmail.env", "email.env", "email.env.txt")]
LOG_DIR = Path(r"D:\MIS\GNC\Automate\Logs")
LOGIN_URL = "https://app.kwikengage.ai/login"
LINK_RE = re.compile(r'https://[^\s"<>\']*kwikengage[^\s"<>\']*\.csv', re.I)
TXT = "translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"


def log(m):
    print(f"{datetime.now():%H:%M:%S}  {m}", flush=True)


def read_env(path, need):
    env = {}
    if path.exists():
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            if "=" in line and not line.strip().startswith("#"):
                k, v = line.split("=", 1); env[k.strip()] = v.strip().strip('"')
    for k in need:
        if not env.get(k):
            raise SystemExit(f"FAILED: {k} missing in {path}")
    return env


def make_driver(show):
    o = webdriver.ChromeOptions()
    o.add_argument("--window-size=1600,1000"); o.add_argument("--disable-notifications")
    if not show:
        o.add_argument("--headless=new")
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=o)
    except Exception:
        return webdriver.Chrome(options=o)


def js_click(drv, el):
    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    drv.execute_script("arguments[0].click();", el)


def type_into(drv, el, text, enter=False):
    el.click(); el.send_keys(Keys.CONTROL + "a"); el.send_keys(Keys.DELETE); time.sleep(0.2)
    el.send_keys(text)
    if enter:
        el.send_keys(Keys.ENTER)
    time.sleep(0.4)


def visible(drv, xp):
    return [e for e in drv.find_elements(By.XPATH, xp) if e.is_displayed()]


# ---------------------------------------------------------------- KwikEngage UI
def export_report(drv, env, d1: date, d2: date):
    w = WebDriverWait(drv, 40)
    drv.get(LOGIN_URL)
    w.until(lambda d: visible(d, "//input[@type='password']"))
    u = (visible(drv, "//input[@id='actual-email']") or visible(drv, "//input[@type='email' or @type='text'][not(@name='dummy-email')]"))[0]
    type_into(drv, u, env["KWIK_USER"])
    p = visible(drv, "//input[@type='password']")[0]
    type_into(drv, p, env["KWIK_PASSWORD"])
    for cb in drv.find_elements(By.XPATH, "//input[@type='checkbox']"):          # "I accept the Terms of service and Privacy Policy" (allowed by the account owner)
        if not cb.is_selected():
            drv.execute_script("arguments[0].click();", cb); time.sleep(0.4)
            if not cb.is_selected():                                                  # custom checkbox: click its label / box
                for lab in drv.find_elements(By.XPATH, "//label[.//input[@type='checkbox']] | //*[contains(.,'By logging in')]/preceding-sibling::*[1]"):
                    if lab.is_displayed():
                        lab.click(); break
    btn = visible(drv, f"//button[@type='submit'] | //button[contains({TXT},'LOGIN') or contains({TXT},'SIGN IN') or contains({TXT},'LOG IN')]")
    js_click(drv, btn[0]) if btn else p.send_keys(Keys.ENTER)
    w.until(lambda d: "login" not in d.current_url.lower())
    log(f"logged in ({drv.current_url})")
    time.sleep(4)

    # Analytics: in the sidebar flyout (appears when the chat icon is hovered / clicked)
    ana = "//li[@role='menuitem'][.//*[normalize-space(text())='Analytics'] or normalize-space(text())='Analytics'] | //*[normalize-space(text())='Analytics']"
    found = visible(drv, ana)
    if not found:                                         # the sidebar popup opens on mouse-over; real hover is not enough here, so send the events
        for t in drv.find_elements(By.CSS_SELECTOR, "#side-nav .ant-menu-submenu-title, #side-nav li.ant-menu-submenu"):
            for ev in ("pointerover", "pointerenter", "mouseover", "mouseenter", "mousemove"):
                drv.execute_script("arguments[0].dispatchEvent(new MouseEvent(arguments[1],{bubbles:true,cancelable:true,view:window}));", t, ev)
            time.sleep(1.5)
            found = visible(drv, ana)
            if found:
                break
    if not found:
        raise RuntimeError("Analytics menu item not found")
    js_click(drv, found[0]); log("Analytics opened")
    exp = w.until(EC.presence_of_element_located((By.XPATH, "//*[normalize-space(text())='Export Reports']")))
    time.sleep(3)
    js_click(drv, exp); log("Export Reports clicked")

    dlg = w.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".ant-modal-content")))
    w.until(lambda d: dlg.find_elements(By.CSS_SELECTOR, "input[date-range='start']") and dlg.find_element(By.CSS_SELECTOR, "input[date-range='start']").is_displayed())
    time.sleep(1.5)                                                       # modal animation
    # e-mail (ant AutoComplete: plain typing is ignored, a value set with an input event is accepted)
    em = dlg.find_element(By.CSS_SELECTOR, ".ant-select-selection-search-input")
    addr = env.get("KWIK_EXPORT_EMAIL", "tausif.ansari@teammas.in")
    drv.execute_script("""const el=arguments[0];el.focus();const set=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set;
                          set.call(el,arguments[1]);el.dispatchEvent(new Event('input',{bubbles:true}));""", em, addr)
    time.sleep(0.8); em.send_keys(Keys.ENTER); time.sleep(0.5)
    if em.get_attribute("value") != addr:
        raise RuntimeError(f"e-mail not accepted: {em.get_attribute('value')!r}")
    # dates: pick the days in the calendar panel (typing is cleared by the picker)
    st = dlg.find_element(By.CSS_SELECTOR, "input[date-range='start']"); en = dlg.find_element(By.CSS_SELECTOR, "input[date-range='end']")
    def pick(day):
        iso = day.strftime("%Y-%m-%d")
        for _ in range(8):
            cells = [c for c in drv.find_elements(By.CSS_SELECTOR, ".ant-picker-dropdown:not(.ant-picker-dropdown-hidden) td.ant-picker-cell[title]") if c.get_attribute("title") == iso]
            if cells:
                drv.execute_script("arguments[0].click();", cells[0]); time.sleep(0.8); return
            heads = drv.find_elements(By.CSS_SELECTOR, ".ant-picker-dropdown:not(.ant-picker-dropdown-hidden) .ant-picker-header-prev-btn")
            if not heads:
                break
            drv.execute_script("arguments[0].click();", heads[0]); time.sleep(0.5)       # earlier month
        raise RuntimeError(f"day {iso} not found in the calendar")
    drv.execute_script("arguments[0].click();", st); time.sleep(1.0)
    pick(d1); pick(d2)
    got = [st.get_attribute("value"), en.get_attribute("value")]
    log(f"dates: {got}")
    if got != [d1.strftime("%d-%m-%Y"), d2.strftime("%d-%m-%Y")]:
        raise RuntimeError(f"date range not accepted: {got}")
    # report type must be "Ticket"
    sel = dlg.find_elements(By.CSS_SELECTOR, ".ant-select:not(.ant-select-auto-complete) .ant-select-selection-item")
    cur = (sel[0].get_attribute("title") or sel[0].text) if sel else ""
    log(f"report type: {cur!r}")
    if cur.strip() != "Ticket":
        box = dlg.find_elements(By.CSS_SELECTOR, ".ant-select:not(.ant-select-auto-complete) .ant-select-selector")
        if not box:
            raise RuntimeError("report type dropdown not found")
        js_click(drv, box[0]); time.sleep(0.8)
        js_click(drv, w.until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class,'ant-select-item-option')][@title='Ticket' or normalize-space(.)='Ticket']"))))
    sent = datetime.now(timezone.utc)
    ex = [e for e in dlg.find_elements(By.XPATH, ".//*[normalize-space(text())='Export']") if e.is_displayed()]
    if not ex:
        raise RuntimeError("Export button not found in the dialog")
    js_click(drv, ex[0]); log("Export clicked")
    time.sleep(5)
    return sent


# ---------------------------------------------------------------- the mail with the link
def find_link(since: datetime, timeout: int):
    env = read_env(next((p for p in GMAIL_ENV if p.exists()), GMAIL_ENV[-1]), ["GMAIL_USER", "GMAIL_APP_PASSWORD"])
    end = time.time() + timeout
    while time.time() < end:
        im = imaplib.IMAP4_SSL("imap.gmail.com")
        try:
            im.login(env["GMAIL_USER"], env["GMAIL_APP_PASSWORD"].replace(" ", ""))
            im.select('"[Gmail]/All Mail"', readonly=True)
            typ, ids = im.search(None, f'(FROM "kwikengage" SINCE {(since - timedelta(days=1)):%d-%b-%Y})')
            best = None
            for mid in ids[0].split()[-15:]:
                m = email.message_from_bytes(im.fetch(mid, "(BODY.PEEK[])")[1][0][1])
                when = parsedate_to_datetime(m.get("Date"))
                if when < since - timedelta(seconds=90):
                    continue
                body = ""
                for part in m.walk():
                    if part.get_content_type() in ("text/plain", "text/html"):
                        body += part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "ignore")
                mm = LINK_RE.search(body)
                if mm and "ticket" in mm.group(0).lower() and (best is None or when > best[0]):
                    best = (when, mm.group(0))
            if best:
                log(f"mail of {best[0]:%H:%M:%S} found")
                return best[1]
        finally:
            try: im.logout()
            except Exception: pass
        log("export mail not here yet ..."); time.sleep(20)
    raise TimeoutError("no TicketBulkExport mail with a download link arrived")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="d1"); ap.add_argument("--to", dest="d2")
    ap.add_argument("--out", required=True); ap.add_argument("--show", action="store_true")
    ap.add_argument("--mail-timeout", type=int, default=900); ap.add_argument("--link")
    a = ap.parse_args()
    y = date.today() - timedelta(days=1)
    d1 = datetime.strptime(a.d1, "%Y-%m-%d").date() if a.d1 else y
    d2 = datetime.strptime(a.d2, "%Y-%m-%d").date() if a.d2 else y
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    result = {}
    drv = None
    try:
        link = a.link
        if not link:
            env = read_env(ENV_FILE, ["KWIK_USER", "KWIK_PASSWORD"])
            drv = make_driver(a.show)
            log(f"Chat tickets {d1} .. {d2}")
            sent = export_report(drv, env, d1, d2)
            drv.quit(); drv = None
            link = find_link(sent, a.mail_timeout)
        log(f"downloading {link}")
        r = requests.get(link, timeout=300); r.raise_for_status()
        f = out / f"ChatTickets_{d1:%Y-%m-%d}_{d2:%Y-%m-%d}.csv"
        f.write_bytes(r.content)
        head = r.content[:400].decode("utf-8-sig", "ignore")
        if "TicketId" not in head:
            raise RuntimeError(f"the downloaded file is not the ticket export (starts: {head[:80]!r})")
        log(f"OK  {f.name}  ({len(r.content):,} bytes)")
        result["Chat"] = {"ok": True, "file": str(f)}
    except Exception as e:
        log(f"FAILED - {type(e).__name__}: {str(e).splitlines()[0] if str(e) else ''}")
        result["Chat"] = {"ok": False, "error": f"{type(e).__name__}: {e}"}
        traceback.print_exc()
        try:
            if drv:
                dbg = LOG_DIR / "inbound-debug"; dbg.mkdir(parents=True, exist_ok=True)
                st = datetime.now().strftime("%Y%m%d-%H%M%S")
                drv.save_screenshot(str(dbg / f"Chat-{st}.png")); (dbg / f"Chat-{st}.html").write_text(drv.page_source, encoding="utf-8", errors="ignore")
                log(f"page saved to {dbg}\\Chat-{st}.png/.html")
        except Exception:
            pass
    finally:
        if drv:
            try: drv.quit()
            except Exception: pass
    (out / "_result.json").write_text(json.dumps(result, indent=1), encoding="utf-8")
    return 0 if result["Chat"]["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
