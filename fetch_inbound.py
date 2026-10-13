r"""
Downloads the three Inbound reports for a date range (default: yesterday) with Chrome / Selenium and saves them as
    <out>\Feedback_<from>_<to>.xlsx     http://192.168.11.244   CDR/CSAT Report > GNC (asterisk) > Search > Export Excel   -> "Feedback Raw"
    <out>\InboundCDR_<from>_<to>.xls    http://192.168.11.240/mis_dashboardGNC/row_in.php  > CDR Report IN > Export       -> "CDR Raw"
    <out>\DDTagging_<from>_<to>.xls     http://192.168.11.6/dialdesk  GNC > In Call Operations > In Call Details > All     -> "Inbound DD Tagging"
and writes <out>\_result.json  ({"Feedback": {"ok": true, "file": ...}, ...}). Exit code 1 if any of them failed
(the others are still downloaded). On a failure the page is saved to <logs>\inbound-debug\ (png + html).

    python fetch_inbound.py --from 2026-10-10 --to 2026-10-11 --out "D:\MIS\GNC\Automate\Oct'26\Data\Inbound\2026-10-11"
    options: --only feedback,cdr,dd   --show (visible browser)   --timeout 240

Logins are read from D:\MIS\GNC\Automate\inbound.env (FEEDBACK_USER / FEEDBACK_PASSWORD, DIALDESK_USER / DIALDESK_PASSWORD).
The 192.168.11.x sites are plain http: Chrome blocks such downloads ("Unconfirmed ... .crdownload") unless the origin is
treated as secure - that is what --unsafely-treat-insecure-origin-as-secure below is for.
"""
import argparse, json, os, shutil, sys, time, traceback
from datetime import date, datetime
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

ENV_FILE = Path(r"D:\MIS\GNC\Automate\inbound.env")
LOG_DIR = Path(r"D:\MIS\GNC\Automate\Logs")
HOSTS = ["http://192.168.11.244", "http://192.168.11.240", "http://192.168.11.6"]
UP = "translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"


def log(msg):
    print(f"{datetime.now():%H:%M:%S}  {msg}", flush=True)


def read_env():
    env = {}
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text(encoding="utf-8-sig").splitlines():
            if "=" in line and not line.strip().startswith("#"):
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    for k in ("FEEDBACK_USER", "FEEDBACK_PASSWORD", "DIALDESK_USER", "DIALDESK_PASSWORD"):
        if not env.get(k):
            raise SystemExit(f"FAILED: {k} missing in {ENV_FILE}")
    return env


def make_driver(dl_dir: Path, show: bool):
    dl_dir.mkdir(parents=True, exist_ok=True)
    o = webdriver.ChromeOptions()
    o.add_experimental_option("prefs", {
        "download.default_directory": str(dl_dir), "download.prompt_for_download": False,
        "download.directory_upgrade": True, "safebrowsing.enabled": False,
        "profile.default_content_setting_values.automatic_downloads": 1})
    o.add_argument("--allow-running-insecure-content")
    o.add_argument("--ignore-certificate-errors")
    o.add_argument("--disable-features=InsecureDownloadWarnings")
    o.add_argument("--unsafely-treat-insecure-origin-as-secure=" + ",".join(HOSTS))
    o.add_argument("--window-size=1600,1000")
    if not show:
        o.add_argument("--headless=new")
    try:                                                   # same driver handling as the morning tasks
        from webdriver_manager.chrome import ChromeDriverManager
        drv = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=o)
    except Exception:
        drv = webdriver.Chrome(options=o)                  # Selenium Manager
    drv.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": str(dl_dir)})
    return drv


def wait_download(dl_dir: Path, timeout: int) -> Path:
    end = time.time() + timeout
    last = None
    while time.time() < end:
        files = [f for f in dl_dir.iterdir() if f.is_file()]
        if any(f.suffix.lower() in (".crdownload", ".tmp") or f.name.startswith("Unconfirmed") for f in files):
            time.sleep(1); continue
        done = [f for f in files if f.stat().st_size > 0]
        if done:
            f = max(done, key=lambda p: p.stat().st_mtime)
            sig = (f.name, f.stat().st_size)
            if sig == last:
                return f
            last = sig
        time.sleep(2)
    left = [f.name for f in dl_dir.iterdir()]
    raise TimeoutError(f"no finished download in {timeout}s (folder: {left})")


def dates_of(el_value: str):
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(el_value.strip(), fmt).date()
        except ValueError:
            pass
    return None


def fill_date(drv, el, d: date):
    """Sets a date input so that the page really sees it (React / plain forms); verifies by reading it back."""
    iso, dmy = d.strftime("%Y-%m-%d"), d.strftime("%d-%m-%Y")
    js = """const el=arguments[0], v=arguments[1];
            const set=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set; set.call(el,v);
            for (const t of ['input','change','blur']) el.dispatchEvent(new Event(t,{bubbles:true}));"""
    drv.execute_script(js, el, iso)
    time.sleep(0.3)
    if dates_of(el.get_attribute("value") or "") == d:
        return
    for text in (dmy, iso):                                 # typed, like the old scripts
        el.click(); el.send_keys(Keys.CONTROL + "a"); el.send_keys(Keys.DELETE); time.sleep(0.2); el.send_keys(text)
        drv.execute_script("arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
        time.sleep(0.3)
        if dates_of(el.get_attribute("value") or "") == d:
            return
    raise RuntimeError(f"could not set the date field to {iso} (page shows '{el.get_attribute('value')}')")


def js_click(drv, el):
    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    drv.execute_script("arguments[0].click();", el)


# ------------------------------------------------------------------ 1. Feedback (244 UI) -> xlsx
def feedback(drv, env, d1, d2, timeout):
    w = WebDriverWait(drv, 40)
    drv.get("http://192.168.11.244/")
    u = w.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Enter username'] | (//input[@type='text'])[1]")))
    u.clear(); u.send_keys(env["FEEDBACK_USER"])
    p = drv.find_element(By.XPATH, "//input[@placeholder='Enter password'] | //input[@type='password']")
    p.clear(); p.send_keys(env["FEEDBACK_PASSWORD"])
    js_click(drv, drv.find_element(By.XPATH, f"//button[contains({UP},'SIGN IN')] | //button[@type='submit']"))
    menu = w.until(EC.presence_of_element_located((By.XPATH, "//a[contains(.,'CDR/CSAT Report')] | //*[self::span or self::div or self::button][contains(normalize-space(.),'CDR/CSAT Report')]")))
    js_click(drv, menu)
    sel = w.until(EC.presence_of_element_located((By.XPATH, "//select[.//option[contains(.,'GNC (asterisk)')]]")))
    Select(sel).select_by_visible_text("GNC (asterisk)")
    time.sleep(2)
    ins = w.until(lambda dr: dr.find_elements(By.XPATH, "//input[@type='date']") if len(dr.find_elements(By.XPATH, "//input[@type='date']")) >= 2 else False)
    fill_date(drv, ins[0], d1); fill_date(drv, ins[1], d2)
    log(f"Feedback: dates set {ins[0].get_attribute('value')} .. {ins[1].get_attribute('value')}")
    btn = drv.find_elements(By.XPATH, f"//button[contains({UP},'SEARCH')] | //input[@type='submit' or @type='button'][contains(translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'SEARCH')]")
    if btn:
        js_click(drv, btn[0]); log("Feedback: Search clicked"); time.sleep(5)
        fill_date(drv, drv.find_elements(By.XPATH, "//input[@type='date']")[0], d1)       # the page may reset them after a search
        fill_date(drv, drv.find_elements(By.XPATH, "//input[@type='date']")[1], d2)
    exp = w.until(EC.presence_of_element_located((By.XPATH, f"//button[contains({UP},'EXPORT')]")))
    js_click(drv, exp); log("Feedback: Export clicked")


# ------------------------------------------------------------------ 2. CDR Report IN (240) -> xls
def cdr(drv, env, d1, d2, timeout):
    w = WebDriverWait(drv, 40)
    drv.get("http://192.168.11.240/mis_dashboardGNC/row_in.php")
    w.until(EC.element_to_be_clickable((By.LINK_TEXT, "CDR Report IN"))).click()
    s = w.until(EC.presence_of_element_located((By.ID, "start_date")))
    e = drv.find_element(By.ID, "end_date")
    for el, d in ((s, d1), (e, d2)):
        try:
            fill_date(drv, el, d)
        except RuntimeError:                                # the old script: clear + yyyy-mm-dd
            el.clear(); el.send_keys(d.strftime("%Y-%m-%d"))
    log(f"CDR: dates set {s.get_attribute('value')} .. {e.get_attribute('value')}")
    js_click(drv, drv.find_element(By.NAME, "btnExport")); log("CDR: Export clicked")


# ------------------------------------------------------------------ 3. DialDesk In Call Details (All) -> xls
def dd(drv, env, d1, d2, timeout):
    w = WebDriverWait(drv, 40)
    drv.get("http://192.168.11.6/dialdesk/admins")
    w.until(EC.presence_of_element_located((By.ID, "UserName"))).send_keys(env["DIALDESK_USER"])
    drv.find_element(By.ID, "Password").send_keys(env["DIALDESK_PASSWORD"])
    drv.find_element(By.XPATH, "//input[@value='Login']").click()
    w.until(EC.presence_of_element_located((By.ID, "HomesClientID")))
    Select(drv.find_element(By.ID, "HomesClientID")).select_by_visible_text("GNC")
    time.sleep(2)
    ops, det = "//span[text()='In Call Operations']", "//span[text()='In Call Details']"
    if not drv.find_elements(By.XPATH, ops) or not drv.find_element(By.XPATH, ops).is_displayed():
        w.until(EC.presence_of_element_located((By.XPATH, "//i[text()='menu']")))      # sidebar collapsed: open it
        js_click(drv, drv.find_element(By.XPATH, "//i[text()='menu']")); time.sleep(2)
    for xp in (ops, det):                                   # presence + JS click, like the old script
        w.until(EC.presence_of_element_located((By.XPATH, xp)))
        js_click(drv, drv.find_element(By.XPATH, xp)); time.sleep(1.5)
    w.until(EC.presence_of_element_located((By.ID, "fdate")))
    drv.execute_script("document.getElementById('fdate').value = arguments[0]; document.getElementById('ldate').value = arguments[1];",
                       d1.strftime("%Y-%m-%d"), d2.strftime("%Y-%m-%d"))
    Select(drv.find_element(By.ID, "type")).select_by_visible_text("All")
    time.sleep(1)
    log(f"DD: dates set {drv.find_element(By.ID, 'fdate').get_attribute('value')} .. {drv.find_element(By.ID, 'ldate').get_attribute('value')}")
    js_click(drv, drv.find_element(By.XPATH, "//input[@value='Export']")); log("DD: Export clicked")


JOBS = {"feedback": ("Feedback", feedback, ".xlsx"), "cdr": ("InboundCDR", cdr, ".xls"), "dd": ("DDTagging", dd, ".xls")}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="d1", required=True); ap.add_argument("--to", dest="d2", required=True)
    ap.add_argument("--out", required=True); ap.add_argument("--only", default="feedback,cdr,dd")
    ap.add_argument("--show", action="store_true"); ap.add_argument("--timeout", type=int, default=240)
    a = ap.parse_args()
    d1 = datetime.strptime(a.d1, "%Y-%m-%d").date(); d2 = datetime.strptime(a.d2, "%Y-%m-%d").date()
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    env = read_env()
    result = {}
    for key in [k.strip() for k in a.only.split(",") if k.strip()]:
        name, fn, ext = JOBS[key]
        dl = out / f"_dl_{key}"
        shutil.rmtree(dl, ignore_errors=True)
        drv = None
        try:
            log(f"{name}: {d1} .. {d2}")
            drv = make_driver(dl, a.show)
            fn(drv, env, d1, d2, a.timeout)
            f = wait_download(dl, a.timeout)
            final = out / f"{name}_{d1:%Y-%m-%d}_{d2:%Y-%m-%d}{f.suffix.lower() or ext}"
            if final.exists():
                final.unlink()
            shutil.move(str(f), str(final))
            log(f"{name}: OK  {final.name}  ({final.stat().st_size:,} bytes)")
            result[name] = {"ok": True, "file": str(final)}
        except Exception as e:
            log(f"{name}: FAILED - {type(e).__name__}: {str(e).splitlines()[0] if str(e) else ''}")
            result[name] = {"ok": False, "error": f"{type(e).__name__}: {e}"}
            try:
                dbg = LOG_DIR / "inbound-debug"; dbg.mkdir(parents=True, exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                if drv:
                    drv.save_screenshot(str(dbg / f"{name}-{stamp}.png"))
                    (dbg / f"{name}-{stamp}.html").write_text(drv.page_source, encoding="utf-8", errors="ignore")
                    log(f"{name}: page saved to {dbg}\\{name}-{stamp}.png/.html")
            except Exception:
                pass
            traceback.print_exc()
        finally:
            if drv:
                try: drv.quit()
                except Exception: pass
            shutil.rmtree(dl, ignore_errors=True)
    (out / "_result.json").write_text(json.dumps(result, indent=1), encoding="utf-8")
    return 0 if all(v["ok"] for v in result.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
