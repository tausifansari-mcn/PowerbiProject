r"""
Downloads the GNC APR and the PP APR (VICIdial "Agent Time Detail" report) for one day from the dialer
http://192.168.11.240 and saves them as  <out>\GNCAPR <date>.csv  and  <out>\PPAPR <date>.csv
  GNC APR: all campaigns, all user groups   (same as Desktop\Python\gncapr.py)
  PP APR : campaigns INFLOB1 + INFLOB2       (same as Desktop\Python\PPAPR.py, with its own login)
The dialer login is read from D:\MIS\GNC\Automate\dialer.env (else from the URL in Desktop\Python\gncapr.py).
    python fetch_apr.py --date 2026-10-10 --out "D:\MIS\GNC\Automate\Data\APR"
"""
import argparse, re, sys
from pathlib import Path
import requests

HOST = "http://192.168.11.240"
import os
if os.environ.get("APR_HOST"): HOST = os.environ["APR_HOST"]
LOGIN_SRC = Path(r"C:\Users\MAS60358\Desktop\Python\gncapr.py")
REPORTS = {"GNCAPR": ["--ALL--"], "PPAPR": ["INFLOB1", "INFLOB2"]}
# each report uses the login of its own morning script (the PP campaigns are only visible to the PP login)
LOGIN_FOR = {"GNCAPR": Path(r"C:\Users\MAS60358\Desktop\Python\gncapr.py"), "PPAPR": Path(r"C:\Users\MAS60358\Desktop\Python\PPAPR.py")}


ENV_FILE = Path(r"D:\MIS\GNC\Automate\dialer.env")


def login(src_file=None):
    # 1) D:\MIS\GNC\Automate\dialer.env  (DIALER_USER=... / DIALER_PASSWORD=..., filled in by you)
    if ENV_FILE.exists():
        env = {}
        for line in ENV_FILE.read_text(encoding="utf-8-sig").splitlines():
            if "=" in line and not line.strip().startswith("#"):
                k, v = line.split("=", 1); env[k.strip()] = v.strip()
        if env.get("DIALER_USER") and env.get("DIALER_PASSWORD"):
            return env["DIALER_USER"], env["DIALER_PASSWORD"]
    # 2) USERNAME = "..." / PASSWORD = "..." in Desktop\Python\gncapr.py (the morning APR script)
    src = (src_file or LOGIN_SRC).read_text(encoding="utf-8", errors="ignore")
    u = re.search(r'^USERNAME\s*=\s*["\'](.*?)["\']', src, re.M)
    p = re.search(r'^PASSWORD\s*=\s*["\'](.*?)["\']', src, re.M)
    if not (u and p):
        sys.exit(f"FAILED: no dialer login found in {LOGIN_SRC}")
    return u.group(1), p.group(1)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    for name, groups in REPORTS.items():
        auth = login(LOGIN_FOR.get(name))
        params = [("query_date", a.date), ("end_date", a.date), ("user_group[]", "--ALL--"),
                  ("shift", "ALL"), ("file_download", "1"), ("SUBMIT", "SUBMIT")] + [("group[]", g) for g in groups]
        r = requests.get(f"{HOST}/vicidial/AST_agent_time_detail.php", params=params, auth=auth, timeout=300)
        if r.status_code == 401:
            print(f"FAILED: the dialer {HOST} refused the login - put the {HOST} user and password in {ENV_FILE}")
            return 1
        r.raise_for_status()
        body = r.content
        if f"Time range: {a.date} 00:00:00".encode() not in body[:600]:
            print(f"FAILED: {name} - the dialer did not return the {a.date} report (HTTP {r.status_code}, {len(body)} bytes)")
            return 1
        f = out / f"{name} {a.date}.csv"
        f.write_bytes(body)
        print(f"{name}: {f.name} ({body.count(b'\n')} lines)")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(f"FAILED: {e}")
        sys.exit(1)
