# GitHub Releases 자동 업데이트 마이그레이션 가이드

> 주식 대시보드 앱의 자동 업데이트를 **Supabase → GitHub Releases**로 마이그레이션
> 메모앱(MemoSync)에서 적용한 동일한 방식

---

## 📋 현재 상황

### ✅ 이미 구현된 부분
- `/api/update/check` — 최신 버전 체크
- `/api/update/history` — 버전 히스토리
- `/api/update/install` — 인스톨러 다운로드 + 자동 설치 + 재시작
- 헬퍼 배치 스크립트 (앱 종료 대기 → 설치 → 재시작)

### 🔄 변경 대상
- **데이터 소스**: Supabase `app_versions` 테이블 → GitHub Releases API
- 인스톨러 호스팅: Supabase Storage (또는 자체) → GitHub Releases Assets

---

## 🎯 마이그레이션 단계

### 1️⃣ GitHub 공개 레포 생성

```bash
# gh CLI로 자동 생성
gh repo create stock-dashboard-releases \
  --public \
  --description "주식 대시보드 업데이트 기록" \
  --add-readme
```

또는 웹에서: https://github.com/new
- Public 선택
- README 추가

### 2️⃣ 기존 인스톨러를 GitHub에 업로드

빌드된 `StockDashboard_Setup.exe`를 GitHub Release로 등록:

```bash
# 첫 릴리즈 만들기
gh release create v1.1.5 \
  --repo {USER}/stock-dashboard-releases \
  --title "v1.1.5 - 안정 버전" \
  --notes-file release-notes.md

# 인스톨러 업로드
gh release upload v1.1.5 \
  "dist\StockDashboard_Setup.exe" \
  --repo {USER}/stock-dashboard-releases
```

`release-notes.md` 예시:
```markdown
## 주요 변경사항

- 포트폴리오 자동 재계산 개선
- Supabase 동기화 안정성 향상
- UI 가독성 개선

[MANDATORY] 이 줄을 추가하면 필수 업데이트로 분류됩니다.
```

### 3️⃣ Python 코드 수정 (`app.py`)

#### 변경 1: 환경변수 추가
```python
# 기존 SUPABASE_URL/KEY 옆에 추가
GH_OWNER = os.environ.get('GH_OWNER', '') or _app_cfg.get('gh_owner', '')
GH_REPO = os.environ.get('GH_REPO', '') or _app_cfg.get('gh_repo', '')
IS_GITHUB_UPDATE = bool(GH_OWNER and GH_REPO)
```

`config.json` 추가:
```json
{
  "gh_owner": "your-github-id",
  "gh_repo": "stock-dashboard-releases"
}
```

#### 변경 2: `/api/update/check` 엔드포인트 교체

```python
@app.route('/api/update/check')
def api_update_check():
    """GitHub Releases에서 최신 버전 조회 → 현재 버전과 비교"""
    if not IS_GITHUB_UPDATE:
        return jsonify({
            'update_available': False,
            'current_version': APP_VERSION,
            'error': 'GitHub 설정이 없습니다'
        })
    try:
        r = requests.get(
            f'https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/releases/latest',
            headers={'Accept': 'application/vnd.github+json'},
            timeout=10
        )
        if r.status_code != 200:
            return jsonify({
                'update_available': False,
                'current_version': APP_VERSION,
                'error': f'GitHub API 오류: {r.status_code}'
            })
        latest = r.json()
        latest_version = latest.get('tag_name', '').lstrip('v')
        update_available = _is_newer(latest_version, APP_VERSION)

        # 인스톨러 자산 찾기
        download_url = ''
        for asset in latest.get('assets', []):
            if asset['name'].endswith('.exe'):
                download_url = asset['browser_download_url']
                break

        body = latest.get('body', '') or ''
        return jsonify({
            'update_available': update_available,
            'current_version': APP_VERSION,
            'latest_version': latest_version,
            'release_notes': body,
            'download_url': download_url,
            'is_mandatory': '[MANDATORY]' in body or '[필수]' in body,
            'published_at': latest.get('published_at', ''),
            'html_url': latest.get('html_url', ''),
        })
    except Exception as e:
        return jsonify({
            'update_available': False,
            'current_version': APP_VERSION,
            'error': f'업데이트 확인 실패: {e}'
        })
```

#### 변경 3: `/api/update/history` 엔드포인트 교체

```python
@app.route('/api/update/history')
def api_update_history():
    """모든 버전 히스토리 조회 (GitHub Releases)"""
    if not IS_GITHUB_UPDATE:
        return jsonify({'versions': []})
    try:
        r = requests.get(
            f'https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/releases?per_page=30',
            headers={'Accept': 'application/vnd.github+json'},
            timeout=10
        )
        if r.status_code != 200:
            return jsonify({'versions': [], 'error': f'status {r.status_code}'})

        releases = r.json() or []
        versions = []
        for rel in releases:
            if rel.get('draft'):
                continue
            body = rel.get('body', '') or ''
            versions.append({
                'version': rel.get('tag_name', '').lstrip('v'),
                'release_notes': body,
                'is_mandatory': '[MANDATORY]' in body or '[필수]' in body,
                'published_at': rel.get('published_at', ''),
                'html_url': rel.get('html_url', ''),
                'prerelease': rel.get('prerelease', False),
            })
        return jsonify({'versions': versions})
    except Exception as e:
        return jsonify({'versions': [], 'error': str(e)})
```

#### 변경 4: `/api/update/install` 그대로 유지

기존 인스톨러 다운로드 + 헬퍼 배치 스크립트 로직은 변경 없음.
URL이 GitHub assets URL로 바뀌었을 뿐.

### 4️⃣ Supabase `app_versions` 테이블 정리 (선택)

더 이상 사용 안 하므로 삭제 가능:
```sql
DROP TABLE IF EXISTS app_versions;
```

(기존 동기화 테이블 `sync_data`는 그대로 유지 — 데이터 동기화용)

### 5️⃣ 새 버전 배포 워크플로

```bash
# 1. 코드 수정 + 버전 올림 (app.py: APP_VERSION = "v1.2.0")

# 2. PyInstaller로 빌드
build.bat
# 또는
pyinstaller dashboard.spec

# 3. Inno Setup으로 인스톨러 빌드
"%PROGRAMFILES(X86)%\Inno Setup 6\ISCC.exe" installer.iss

# 4. GitHub에 자동 배포
gh release create v1.2.0 \
  --repo {USER}/stock-dashboard-releases \
  --title "v1.2.0 - 새 기능" \
  --notes-file CHANGELOG.md \
  "Output\StockDashboard_Setup.exe"
```

---

## 🆚 메모앱과 차이점

| 항목 | 메모앱 (Electron) | 주식앱 (Python+PyInstaller) |
|------|------------------|----------------------------|
| **델타 업데이트** | ✅ blockmap으로 변경분만 | ❌ 인스톨러 전체 다운로드 |
| **자동 백그라운드** | ✅ electron-updater | ⚠️ 동기 다운로드 (UI 블록) |
| **재시작** | ✅ 자동 | ✅ 헬퍼 배치 스크립트 (이미 구현) |
| **데이터 보존** | ✅ AppData 별도 폴더 | ✅ AppData 별도 폴더 |
| **호스팅 비용** | 무료 (GitHub) | 무료 (GitHub) |
| **무결성 검증** | ✅ SHA512 자동 | ⚠️ 크기만 (개선 필요) |

### ⚠️ 개선 권장사항

**1. SHA256 무결성 검증 추가**
```python
import hashlib

# 다운로드 후 검증
sha256 = hashlib.sha256()
with open(installer_path, 'rb') as f:
    for chunk in iter(lambda: f.read(8192), b''):
        sha256.update(chunk)
calculated = sha256.hexdigest()

# release notes에 SHA256 포함시키고 비교
# 또는 별도 .sha256 파일을 release에 첨부
```

**2. 다운로드 진행률 표시**
```python
# response.headers['content-length'] 활용
# WebSocket 또는 Server-Sent Events로 프론트에 진행률 전송
```

**3. 다운로드 비동기화**
```python
# threading 또는 asyncio로 백그라운드 다운로드
# 다운로드 중에도 사용자 작업 가능
```

---

## 📌 빠른 시작 체크리스트

- [ ] GitHub 레포 생성 (`stock-dashboard-releases`)
- [ ] 기존 v1.1.5 인스톨러 GitHub Release로 등록
- [ ] `config.json`에 `gh_owner`, `gh_repo` 추가
- [ ] `app.py`의 3개 엔드포인트 수정 (`/check`, `/history`, `/install`은 변경 없음)
- [ ] 테스트: 더미 v1.1.6 release 만들어서 알림 뜨는지 확인
- [ ] (선택) Supabase `app_versions` 테이블 삭제
- [ ] CHANGELOG.md 형식 정립

---

## 💡 추가 팁

### 자동 빌드 + 배포 스크립트 (`release.py` 개선)

```python
# release.py에 추가
import subprocess

def github_release(version: str, notes_file: str, installer_path: str, owner: str, repo: str):
    """GitHub에 릴리즈 자동 생성 + 인스톨러 업로드"""
    subprocess.run([
        'gh', 'release', 'create', f'v{version}',
        '--repo', f'{owner}/{repo}',
        '--title', f'v{version}',
        '--notes-file', notes_file,
        installer_path,
    ], check=True)
    print(f'✓ Released v{version} to GitHub')
```

### `installer.iss` 확인

Inno Setup이 이미 다음을 처리하는지 확인:
- 기존 설치 자동 감지 → 덮어쓰기
- 사용자 데이터 폴더 보존 (`AppData`)
- 시작 메뉴/바탕화면 바로가기 갱신

---

## 🔒 보안

- **publisher 검증 약함**: 코드 서명 인증서 권장 (없으면 SmartScreen 경고)
- **공개 release**: release notes는 누구나 볼 수 있음
- **Rate limit**: GitHub API 60 req/hr (인증 안 하면) → 사용자 다수 시 인증 토큰 필요할 수 있음

---

## 🚀 한 줄 요약

> Supabase REST API 호출 부분을 GitHub Releases API로 바꾸면 끝.
> 인스톨러 다운로드/설치 헬퍼는 그대로.
> 무료, 자동, 안전.
</thinking>
