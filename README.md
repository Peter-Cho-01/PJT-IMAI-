# IMA Dry Powder 현황판 (자동 갱신)

- 15분마다 OneDrive의 Master를 공유 링크로 받아, 현황판에 필요한 값만 뽑아 **비밀번호로 암호화**해 GitHub Pages에 올립니다.
- 이 저장소에는 코드만 있고 Master 데이터는 없습니다. 웹에 올라가는 데이터는 암호화되어 있어 비밀번호 없이는 볼 수 없습니다.
- 바로 갱신: Actions 탭 → "IMA 현황판 갱신" → Run workflow

## 설정 값 (Settings → Secrets and variables → Actions)
| 이름 | 값 |
|---|---|
| ONEDRIVE_SHARE_URL | Master의 OneDrive 공유 링크 ("링크가 있는 모든 사람") |
| DASH_PASSWORD | 현황판 비밀번호 (12자 이상 권장) |

## 파일
- `.github/workflows/dashboard.yml` 자동 실행 설정
- `scripts/fetch_master.py` OneDrive에서 Master 받기 (링크는 기록에 남기지 않음)
- `scripts/build_site.py` 필요한 시트 5개 저장값만 뽑아 AES-256-GCM 암호화
- `template/viewer.html` 현황판 화면
