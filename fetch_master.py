"""OneDrive 공유 링크(환경변수 SHARE_URL)로 Master를 내려받아 master.xlsx로 저장한다.
링크·주소는 기록에 남기지 않는다 (호스트만 표시). 실패하면 종료 코드 1."""
import base64, io, os, sys, urllib.error, urllib.parse, urllib.request, zipfile

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
OUT = sys.argv[1] if len(sys.argv) > 1 else "master.xlsx"
link = (os.environ.get("SHARE_URL") or "").strip()
if not link.startswith("https://"):
    sys.exit("실패: 저장소 Secrets에 ONEDRIVE_SHARE_URL(https로 시작하는 공유 링크)이 없습니다.")


def host(u):
    return urllib.parse.urlsplit(u).netloc or "?"


def get(url, limit=80_000_000):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, r.geturl(), r.headers.get("Content-Type", ""), r.read(limit)
    except urllib.error.HTTPError as e:
        return e.code, url, "", b""


def with_param(u, k, v):
    p = urllib.parse.urlsplit(u)
    q = [(a, b) for a, b in urllib.parse.parse_qsl(p.query, keep_blank_values=True) if a != k] + [(k, v)]
    return urllib.parse.urlunsplit((p.scheme, p.netloc, p.path, urllib.parse.urlencode(q), p.fragment))


def enc_share(u):
    return "u!" + base64.b64encode(u.encode()).decode().rstrip("=").replace("/", "_").replace("+", "-")


def is_master(data):
    if data[:2] != b"PK":
        return False
    try:
        wb = zipfile.ZipFile(io.BytesIO(data)).read("xl/workbook.xml").decode("utf-8", "ignore")
        return "월별 Dry Powder" in wb or "DP_현황판" in wb
    except Exception:
        return False


st0, fin0, ct0, body0 = get(link, 3_000_000)
q0 = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(fin0).query))
tries = [("공유 링크 그대로", None), ("공유 링크 + download=1", with_param(link, "download", "1")),
         ("이동한 주소 + download=1", with_param(fin0, "download", "1")),
         ("OneDrive API", f"https://api.onedrive.com/v1.0/shares/{enc_share(link)}/root/content")]
if q0.get("redeem"):
    tries.append(("OneDrive API (redeem)", f"https://api.onedrive.com/v1.0/shares/u!{q0['redeem']}/root/content"))
resid = q0.get("resid") or q0.get("id")
if resid and q0.get("authkey"):
    tries.append(("download?resid&authkey", "https://onedrive.live.com/download?" + urllib.parse.urlencode({"resid": resid, "authkey": q0["authkey"]})))

for name, url in tries:
    try:
        st, fin, ct, data = (st0, fin0, ct0, body0) if url is None else get(url)
    except Exception as e:
        print(f"- {name}: 실패 ({type(e).__name__})"); continue
    ok = is_master(data)
    print(f"- {name}: HTTP {st} · {host(fin)} · {(ct or '-').split(';')[0]} · {len(data):,}바이트 · {'성공' if ok else '실패'}")
    if ok:
        open(OUT, "wb").write(data)
        print(f"Master 받기 성공 ({name}, {len(data)/1e6:.1f}MB)")
        sys.exit(0)
sys.exit("실패: 모든 방법으로 Master를 받지 못했습니다. 공유 링크가 '링크가 있는 모든 사람'인지, 파일에 '월별 Dry Powder' 시트가 있는지 확인해 주세요.")
