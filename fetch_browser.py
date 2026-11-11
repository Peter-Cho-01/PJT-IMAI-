"""실제 브라우저(헤드리스 크롬)로 OneDrive 공유 링크를 게스트처럼 열어 Master를 내려받는다.
단순 다운로드(fetch_master.py)가 로그인 화면으로 막힐 때 쓴다.
- 링크·주소·화면 내용은 기록에 남기지 않는다 (호스트·건수만 표시). 화면 캡처도 저장하지 않는다.
usage: SHARE_URL=... python fetch_browser.py OUT.xlsx
"""
import io, json, os, re, sys, urllib.parse, zipfile
from playwright.sync_api import sync_playwright

OUT = sys.argv[1] if len(sys.argv) > 1 else "master.xlsx"
LINK = (os.environ.get("SHARE_URL") or "").strip()
if not LINK.startswith(("https://", "http://127.0.0.1")):
    sys.exit("실패: ONEDRIVE_SHARE_URL이 없습니다.")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
host = lambda u: urllib.parse.urlsplit(u).netloc or "?"


def is_master(b):
    if not b or b[:2] != b"PK":
        return False
    try:
        wb = zipfile.ZipFile(io.BytesIO(b)).read("xl/workbook.xml").decode("utf-8", "ignore")
        return "월별 Dry Powder" in wb or "DP_현황판" in wb
    except Exception:
        return False


def save(b, how):
    open(OUT, "wb").write(b)
    print(f"Master 받기 성공 ({how}, {len(b)/1e6:.1f}MB)")
    sys.exit(0)


found_urls, direct_files, json_resps, seen_hosts = [], [], [], set()
KEYS = ("@content.downloadUrl", "@microsoft.graph.downloadUrl", "downloadUrl")


def on_response(resp):
    try:
        u = resp.url; seen_hosts.add(host(u))
        ct = (resp.headers.get("content-type") or "").lower()
        if "spreadsheetml" in ct or "octet-stream" in ct:
            direct_files.append(resp)
        elif "json" in ct:
            json_resps.append(resp)   # 본문은 페이지가 다 열린 뒤 읽음 (이벤트 안에서 읽으면 멈출 수 있음)
    except Exception:
        pass


def scan_json():
    for resp in json_resps:
        try:
            t = resp.text()
        except Exception:
            continue
        for k in KEYS:
            for m in re.finditer(re.escape(f'"{k}"') + r'\s*:\s*"([^"]+)"', t):
                v = m.group(1).replace('\\/', '/').encode().decode("unicode_escape")
                if v.startswith(("https://", "http://127.0.0.1")) and v not in found_urls:
                    found_urls.append(v)


with sync_playwright() as pw:
    br = pw.chromium.launch()
    ctx = br.new_context(user_agent=UA, locale="ko-KR", accept_downloads=True, viewport={"width": 1400, "height": 900})
    page = ctx.new_page()
    page.on("response", on_response)
    try:
        page.goto(LINK, wait_until="domcontentloaded", timeout=60000)
        try:
            page.wait_for_load_state("networkidle", timeout=45000)
        except Exception:
            pass
        page.wait_for_timeout(5000)
    except Exception as e:
        print(f"- 링크 열기 오류: {type(e).__name__}")
    final = page.url
    scan_json()
    print(f"- 페이지: {host(final)} · 로그인 화면 여부 {'예' if 'login.' in host(final) else '아니오'} · 응답 호스트 {len(seen_hosts)}곳 · 내려받기 주소 후보 {len(found_urls)}개")

    def tryget(url, how):
        try:
            r = ctx.request.get(url, timeout=90000, max_redirects=10)
            b = r.body() if r.ok else b""
            print(f"- {how}: HTTP {r.status} · {host(r.url)} · {len(b):,}바이트 · {'성공' if is_master(b) else '실패'}")
            if is_master(b):
                save(b, how)
        except Exception as e:
            print(f"- {how}: 실패 ({type(e).__name__})")

    # A. 페이지가 스스로 받아 온 파일·내려받기 주소
    for r in direct_files:
        try:
            b = r.body()
            if is_master(b):
                save(b, "A. 페이지가 받은 파일")
        except Exception:
            pass
    for i, u in enumerate(found_urls[:6]):
        tryget(u, f"A. 내려받기 주소 {i + 1}")

    # B. 최종 주소의 파일 ID로 download.aspx (게스트 쿠키 사용)
    q = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(final).query))
    fid = q.get("resid") or q.get("sourcedoc") or q.get("id") or ""
    fid = fid.strip("{}")
    mcid = re.search(r"/personal/([^/]+)/", final)
    if fid and mcid:
        sch = urllib.parse.urlsplit(final).scheme or "https"
        for h in dict.fromkeys(("onedrive.live.com", "my.microsoftpersonalcontent.com", host(final))):
            tryget(f"{sch if h == host(final) else 'https'}://{h}/personal/{mcid.group(1)}/_layouts/15/download.aspx?UniqueId={urllib.parse.quote(fid)}", f"B. download.aspx ({h})")

    # C. 공유 링크 + download=1 (게스트 쿠키 사용)
    p = urllib.parse.urlsplit(LINK)
    qs = [(a, b) for a, b in urllib.parse.parse_qsl(p.query, keep_blank_values=True) if a != "download"] + [("download", "1")]
    tryget(urllib.parse.urlunsplit((p.scheme, p.netloc, p.path, urllib.parse.urlencode(qs), "")), "C. 공유 링크 + download=1")

    # D. 화면의 '다운로드' 버튼 (OneDrive 미리보기 / Excel 웹 '사본 다운로드')
    labels = ["다운로드", "Download", "사본 다운로드", "Download a Copy", "Download a copy"]
    for fr in [page.main_frame] + page.frames:
        for lb in labels:
            try:
                loc = fr.get_by_role("button", name=lb, exact=False).first
                if loc.count() and loc.is_visible():
                    with page.expect_download(timeout=60000) as dl:
                        loc.click()
                    b = open(dl.value.path(), "rb").read()
                    print(f"- D. '{lb}' 버튼: {len(b):,}바이트 · {'성공' if is_master(b) else '실패'}")
                    if is_master(b):
                        save(b, f"D. '{lb}' 버튼")
            except Exception:
                continue
    # Excel 웹: 파일 → 다른 이름으로 저장 → 사본 다운로드
    for fr in page.frames:
        try:
            f = fr.get_by_role("tab", name=re.compile("^(파일|File)$")).first
            if not (f.count() and f.is_visible()):
                f = fr.get_by_role("button", name=re.compile("^(파일|File)$")).first
            if not (f.count() and f.is_visible()):
                continue
            f.click(); page.wait_for_timeout(1500)
            for step in [re.compile("다른 이름으로 저장|Save [Aa]s|복사본 만들기|Create a [Cc]opy"), re.compile("사본 다운로드|Download a [Cc]opy")]:
                it = fr.get_by_text(step).first
                if not it.count():
                    raise RuntimeError("메뉴 없음")
                if step.pattern.startswith("사본"):
                    with page.expect_download(timeout=60000) as dl:
                        it.click()
                    b = open(dl.value.path(), "rb").read()
                    print(f"- D. Excel 웹 사본 다운로드: {len(b):,}바이트 · {'성공' if is_master(b) else '실패'}")
                    if is_master(b):
                        save(b, "D. Excel 웹 사본 다운로드")
                else:
                    it.click(); page.wait_for_timeout(1500)
        except Exception as e:
            print(f"- D. Excel 웹 메뉴: 실패 ({type(e).__name__})")
            break
    br.close()
sys.exit("실패: 브라우저 방식으로도 Master를 받지 못했습니다. 위 기록을 캡처해 보내 주세요.")
