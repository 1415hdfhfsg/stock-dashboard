"""
배포 자동화 스크립트
사용법: python release.py v1.2.0 "릴리즈 노트"

- 모든 파일의 버전 번호 자동 업데이트
- PyInstaller 빌드
- Inno Setup 빌드
- (선택) Supabase app_versions 업데이트
"""
import sys, os, re, subprocess, json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 환경 확인
PYTHON = r"C:\Users\User\AppData\Local\Programs\Python\Python313\python.exe"
ISCC   = r"C:\Users\User\AppData\Local\InnoSetup\ISCC.exe"

def usage():
    print("사용법: python release.py <새버전> [릴리즈노트]")
    print("예시: python release.py v1.2.0 \"새 기능 추가\"")
    sys.exit(1)

if len(sys.argv) < 2:
    usage()

NEW_VERSION = sys.argv[1].strip()
if not re.match(r'^v\d+\.\d+\.\d+', NEW_VERSION):
    print("[ERROR] 버전 형식: v1.2.0 또는 v1.2.0 Beta")
    sys.exit(1)

RELEASE_NOTES = sys.argv[2] if len(sys.argv) > 2 else f"{NEW_VERSION} 업데이트"

# 현재 버전 감지
with open(os.path.join(BASE_DIR, 'app.py'), 'r', encoding='utf-8') as f:
    match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', f.read())
    CURRENT_VERSION = match.group(1) if match else 'unknown'

print(f"[BUILD] 버전 업그레이드: {CURRENT_VERSION} -> {NEW_VERSION}")
print(f"[NOTE]  릴리즈 노트: {RELEASE_NOTES[:80]}{'...' if len(RELEASE_NOTES)>80 else ''}")
print()

# 1. app.py
def update_file(path, replacements):
    full = os.path.join(BASE_DIR, path)
    if not os.path.exists(full):
        print(f"  [SKIP] {path} (없음)")
        return
    with open(full, 'r', encoding='utf-8') as f:
        content = f.read()
    changed = False
    for old, new in replacements:
        if old in content:
            content = content.replace(old, new)
            changed = True
    if changed:
        with open(full, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"  [OK] {path}")
    else:
        print(f"  [--] {path} (대상 없음)")

print("1) 버전 번호 업데이트")
update_file('app.py', [
    (f'APP_VERSION = "{CURRENT_VERSION}"', f'APP_VERSION = "{NEW_VERSION}"'),
])
update_file('main.py', [
    (f'v{CURRENT_VERSION.lstrip("v")}' if CURRENT_VERSION.startswith('v') else CURRENT_VERSION, NEW_VERSION),
])
# installer.iss
iss_ver = NEW_VERSION.lstrip('v')
with open(os.path.join(BASE_DIR, 'installer.iss'), 'r', encoding='utf-8') as f:
    iss = f.read()
iss = re.sub(r'AppVersion=[^\r\n]+', f'AppVersion={iss_ver}', iss)
iss = re.sub(r'AppVerName=[^\r\n]+', f'AppVerName=내 주식 대시보드 {NEW_VERSION}', iss)
iss = re.sub(r'OutputBaseFilename=[^\r\n]+', f'OutputBaseFilename=StockDashboard_{iss_ver}_Setup', iss)
iss = re.sub(r'UninstallDisplayName=[^\r\n]+', f'UninstallDisplayName=내 주식 대시보드 {NEW_VERSION}', iss)
iss = re.sub(r'BeveledLabel=[^\r\n]+', f'BeveledLabel=내 주식 대시보드 {NEW_VERSION}', iss)
with open(os.path.join(BASE_DIR, 'installer.iss'), 'w', encoding='utf-8') as f:
    f.write(iss)
print("  [OK] installer.iss")

# dashboard.html 버전 배지
html_path = os.path.join(BASE_DIR, 'templates', 'dashboard.html')
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()
html = re.sub(r'>v\d+\.\d+[^<]*</span>(?=\s*<button id="update-badge")',
              f'>{NEW_VERSION}</span>', html, count=1)
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print("  [OK] templates/dashboard.html")

# 2. PyInstaller 빌드
print("\n2) PyInstaller EXE 빌드 (dashboard.spec → dist/StockDashboard/)")
result = subprocess.run(
    [PYTHON, '-m', 'PyInstaller', 'dashboard.spec', '--noconfirm', '--clean'],
    cwd=BASE_DIR, capture_output=True, text=True, encoding='utf-8', errors='replace'
)
if result.returncode == 0:
    print("  [OK] 빌드 성공")
else:
    print("  [FAIL] 빌드 실패")
    print(result.stdout[-500:] if result.stdout else '')
    print(result.stderr[-500:] if result.stderr else '')
    sys.exit(1)

# 3. Inno Setup 빌드
print("\n3) Inno Setup 설치기 빌드")
if not os.path.exists(ISCC):
    print(f"  [FAIL] ISCC 없음: {ISCC}")
    sys.exit(1)
result = subprocess.run(
    [ISCC, 'installer.iss'],
    cwd=BASE_DIR, capture_output=True, text=True, encoding='utf-8', errors='replace'
)
if 'Successful compile' in (result.stdout or ''):
    setup_path = os.path.join(BASE_DIR, 'Output', f'StockDashboard_{iss_ver}_Setup.exe')
    size_mb = os.path.getsize(setup_path) / (1024*1024) if os.path.exists(setup_path) else 0
    print(f"  [OK] 설치기 생성: {setup_path} ({size_mb:.1f} MB)")
else:
    print("  [FAIL] 설치기 빌드 실패")
    print(result.stdout[-500:] if result.stdout else '')
    sys.exit(1)

# 4. GitHub Release 자동 생성 + 자산 업로드 (Setup.exe)
print("\n4) GitHub Release 자동 배포")
try:
    with open(os.path.join(BASE_DIR, 'config.json'), 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    GH_OWNER = cfg.get('gh_owner', '')
    GH_REPO  = cfg.get('gh_repo', '')
    if not (GH_OWNER and GH_REPO):
        print("  [SKIP] GitHub 설정 없음 (config.json: gh_owner, gh_repo)")
    else:
        setup_file = os.path.join(BASE_DIR, 'Output', f'StockDashboard_{iss_ver}_Setup.exe')
        if not os.path.exists(setup_file):
            print(f"  [FAIL] 설치기 없음: {setup_file}")
        else:
            # gh release create + upload (Setup.exe)
            cmd = [
                'gh', 'release', 'create', NEW_VERSION,
                setup_file,
                '--repo', f'{GH_OWNER}/{GH_REPO}',
                '--title', NEW_VERSION,
                '--notes', RELEASE_NOTES,
                '--latest',
            ]
            res = subprocess.run(cmd, cwd=BASE_DIR, capture_output=True, text=True, encoding='utf-8', errors='replace')
            if res.returncode == 0:
                url = (res.stdout or '').strip().split('\n')[-1]
                print(f"  [OK] GitHub Release 생성됨: {url}")
                print(f"       파일: StockDashboard_{iss_ver}_Setup.exe")
            else:
                err = (res.stderr or '').strip()
                if 'already exists' in err.lower():
                    print(f"  [INFO] {NEW_VERSION} 이미 존재 → 자산만 업로드 시도")
                    res2 = subprocess.run(
                        ['gh', 'release', 'upload', NEW_VERSION, setup_file,
                         '--repo', f'{GH_OWNER}/{GH_REPO}', '--clobber'],
                        cwd=BASE_DIR, capture_output=True, text=True, encoding='utf-8', errors='replace'
                    )
                    print(f"  업로드 결과: {res2.returncode}")
                else:
                    print(f"  [FAIL] gh release create 실패: {err[:200]}")
except FileNotFoundError:
    print("  [SKIP] gh CLI 미설치 (winget install GitHub.cli)")
except Exception as e:
    print(f"  [SKIP] GitHub 배포 실패: {e}")

print("\n" + "="*60)
print(f"  배포 완료: {NEW_VERSION}")
print("="*60)
print(f"\n[OK] 사용자에게 즉시 자동 업데이트 알림 전송됨")
print(f"   다운로드 링크: https://github.com/{cfg.get('gh_owner','')}/{cfg.get('gh_repo','')}/releases/latest")
