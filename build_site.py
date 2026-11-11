"""master.xlsx → 비밀번호로 암호화된 현황판 site/index.html (+ site/hash.txt)
- 현황판에 필요한 시트 5개의 저장값만 뽑아(수식·다른 시트 제외) AES-256-GCM으로 암호화한다.
- 키 = PBKDF2-SHA256(비밀번호, salt, 600,000회). salt는 저장소마다 고정 → '7일 기억'이 갱신 후에도 유지됨.
usage: python build_site.py master.xlsx template/viewer.html site   (환경변수 DASH_PASSWORD 필수)
"""
import base64, datetime, hashlib, html, io, json, os, re, sys, zipfile
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

MASTER, TPL, OUTDIR = sys.argv[1:4]
PW = os.environ.get("DASH_PASSWORD", "")
if len(PW) < 8:
    sys.exit("실패: 저장소 Secrets에 DASH_PASSWORD(8자 이상, 12자 이상 권장)가 없습니다.")
ITER = 600_000
NEED = ['월별 Dry Powder', 'DP_현황판', 'Dash Board', 'S3~G2 배분', '국내 딜 인벤토리', '해외 딜 인벤토리']


def slim(path):
    z = zipfile.ZipFile(path)
    R = lambda n: z.read(n).decode('utf-8')
    wbx, rels = R('xl/workbook.xml'), R('xl/_rels/workbook.xml.rels')
    relmap = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', rels):
        el = m.group(0); i = re.search(r'Id="([^"]+)"', el); t = re.search(r'Target="([^"]+)"', el)
        if i and t: relmap[i.group(1)] = t.group(1)
    keep, found = [], []
    for m in re.finditer(r'<sheet\b[^>]*/>', wbx):
        el = m.group(0); nm = html.unescape(re.search(r'name="([^"]+)"', el).group(1)); rid = re.search(r'r:id="([^"]+)"', el).group(1)
        if nm in NEED:
            t = relmap[rid]; t = t.lstrip('/') if t.startswith('/') else 'xl/' + t
            keep.append((nm, rid, t, el)); found.append(nm)
    if not any(n in found for n in ('월별 Dry Powder', 'DP_현황판')):
        sys.exit('실패: Master에 월별 Dry Powder 시트가 없습니다.')
    sst = re.findall(r'<si>.*?</si>|<si/>', R('xl/sharedStrings.xml'), flags=re.S) if 'xl/sharedStrings.xml' in z.namelist() else []
    remap, used = {}, []

    def fix(m):
        attrs, inner = m.group(1), m.group(2)
        if not re.search(r'\bt="s"', attrs): return m.group(0)
        v = re.search(r'<v>(\d+)</v>', inner)
        if not v: return m.group(0)
        i = int(v.group(1))
        if i not in remap: remap[i] = len(used); used.append(sst[i] if i < len(sst) else '<si/>')
        return f'<c{attrs}>' + inner.replace(v.group(0), f'<v>{remap[i]}</v>', 1) + '</c>'

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as o:
        def put(n, data):  # 고정 시각 → 내용이 같으면 결과도 같음 (변경 감지용)
            zi = zipfile.ZipInfo(n, (2000, 1, 1, 0, 0, 0)); zi.compress_type = zipfile.ZIP_DEFLATED; o.writestr(zi, data)
        put('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        put('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + ''.join(el for *_, el in keep) + '</sheets></workbook>')
        put('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(f'<Relationship Id="{rid}" Target="/{t}"/>' for _, rid, t, _ in keep) + '</Relationships>')
        for nm, rid, t, _ in keep:
            x = R(t)
            x = re.sub(r'<f\b[^>]*/>|<f\b[^>]*>.*?</f>', '', x, flags=re.S)
            x = re.sub(r'<(conditionalFormatting|dataValidations|extLst)\b.*?</\1>', '', x, flags=re.S)
            x = re.sub(r'<c(\s[^>]*?[^/])>(.*?)</c>', fix, x, flags=re.S)
            put(t, x)
        put('xl/sharedStrings.xml', '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + ''.join(used) + '</sst>')
    return buf.getvalue(), found


data, found = slim(MASTER)
digest = hashlib.sha256(data + PW.encode()).hexdigest()  # 데이터나 비밀번호가 바뀌면 달라짐
repo = os.environ.get("GITHUB_REPOSITORY", "local")
salt = hashlib.sha256(("ima-dp-salt:" + repo).encode()).digest()[:16]
key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(PW.encode())
iv = os.urandom(12)
ct = AESGCM(key).encrypt(iv, data, None)
built = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds')
payload = {"v": 1, "name": "IMA Master", "built": built, "iter": ITER,
           "salt": base64.b64encode(salt).decode(), "iv": base64.b64encode(iv).decode(), "ct": base64.b64encode(ct).decode()}
tpl = open(TPL, encoding='utf-8').read()
ph = '<script type="application/json" id="embedded-enc"></script>'
if ph not in tpl: sys.exit('실패: 템플릿에 암호화 데이터 자리(embedded-enc)가 없습니다.')
out = tpl.replace(ph, '<script type="application/json" id="embedded-enc">' + json.dumps(payload) + '</script>', 1)
out = out.replace('<head>', '<head>\n<meta name="robots" content="noindex,nofollow">', 1)
os.makedirs(OUTDIR, exist_ok=True)
open(os.path.join(OUTDIR, 'index.html'), 'w', encoding='utf-8').write(out)
open(os.path.join(OUTDIR, 'hash.txt'), 'w').write(digest)
open(os.path.join(OUTDIR, 'robots.txt'), 'w').write('User-agent: *\nDisallow: /\n')
open(os.environ.get('GITHUB_OUTPUT', os.devnull), 'a').write(f'hash={digest}\n')
print(f"현황판 생성: 시트 {len(found)}개, 데이터 {len(data)/1e6:.2f}MB(암호화), index.html {os.path.getsize(os.path.join(OUTDIR, 'index.html'))/1e6:.2f}MB")
