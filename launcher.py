"""
StockDashboard Launcher - 자동 업데이트 적용 + 앱 실행 부트스트랩
런처 패턴 (Squirrel/Discord 방식)

흐름:
  1. pending_update.zip 있으면 → app/ 폴더 교체 (파일 락 없음, 앱 종료 상태)
  2. app/StockDashboard.exe 실행
"""
import os
import sys
import time
import shutil
import zipfile
import subprocess
import traceback
from pathlib import Path


def get_paths():
    """런처 자체 경로 + app 폴더 경로 결정"""
    # 런처 EXE의 위치 (PyInstaller frozen 또는 dev)
    if getattr(sys, 'frozen', False):
        launcher_dir = Path(sys.executable).parent
    else:
        launcher_dir = Path(__file__).parent.absolute()

    app_dir = launcher_dir / 'app'
    app_exe = app_dir / 'StockDashboard.exe'
    backup_dir = launcher_dir / 'app.bak'

    # 업데이트 ZIP 위치 (사용자별 데이터 폴더)
    appdata = os.environ.get('LOCALAPPDATA') or os.environ.get('APPDATA') or os.path.expanduser('~')
    update_root = Path(appdata) / 'StockDashboard'
    update_root.mkdir(parents=True, exist_ok=True)
    pending_zip = update_root / 'pending_update.zip'
    log_file = update_root / 'launcher.log'

    return {
        'launcher_dir': launcher_dir,
        'app_dir': app_dir,
        'app_exe': app_exe,
        'backup_dir': backup_dir,
        'pending_zip': pending_zip,
        'log_file': log_file,
        'update_root': update_root,
    }


def log(msg, log_file):
    """로그 파일에 기록 + stdout 출력"""
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    try:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass


def apply_update(paths):
    """pending_update.zip 적용 (있을 경우)"""
    pending = paths['pending_zip']
    if not pending.exists():
        return True  # 적용할 게 없음 = 정상

    log_file = paths['log_file']
    log(f"=== 업데이트 적용 시작: {pending} ({pending.stat().st_size/1024/1024:.1f}MB)", log_file)

    app_dir = paths['app_dir']
    backup_dir = paths['backup_dir']
    launcher_dir = paths['launcher_dir']

    try:
        # ── 1: ZIP 검증 ─
        if not zipfile.is_zipfile(str(pending)):
            log("ZIP 파일 손상 - 삭제 후 정상 실행", log_file)
            try: pending.unlink()
            except: pass
            return True

        with zipfile.ZipFile(str(pending), 'r') as zf:
            names = zf.namelist()
            if not any('StockDashboard.exe' in n for n in names):
                log("ZIP에 StockDashboard.exe 없음 - 무시", log_file)
                try: pending.unlink()
                except: pass
                return True

        # ── 2: 임시 staging 폴더에 압축 해제 ─
        staging = launcher_dir / '.staging'
        if staging.exists():
            shutil.rmtree(staging, ignore_errors=True)
        staging.mkdir(parents=True, exist_ok=True)

        log(f"압축 해제 중: {staging}", log_file)
        with zipfile.ZipFile(str(pending), 'r') as zf:
            zf.extractall(str(staging))

        # ZIP 안에 StockDashboard/ 폴더 있으면 평탄화
        nested = staging / 'StockDashboard'
        if nested.is_dir() and (nested / 'StockDashboard.exe').exists():
            staging = nested

        # ── 3: 기존 app/ 백업 ─
        if app_dir.exists():
            if backup_dir.exists():
                shutil.rmtree(backup_dir, ignore_errors=True)
            log(f"백업: {app_dir} → {backup_dir}", log_file)
            os.rename(str(app_dir), str(backup_dir))

        # ── 4: staging → app/ ─
        log(f"새 버전 적용: {staging} → {app_dir}", log_file)
        os.rename(str(staging), str(app_dir))

        # ── 5: 검증 ─
        if not paths['app_exe'].exists():
            log("ERROR: 새 app/StockDashboard.exe 없음 - 롤백", log_file)
            if app_dir.exists():
                shutil.rmtree(app_dir, ignore_errors=True)
            os.rename(str(backup_dir), str(app_dir))
            return False

        # ── 6: 정리 ─
        try: pending.unlink()
        except: pass
        if backup_dir.exists():
            shutil.rmtree(backup_dir, ignore_errors=True)

        log("✅ 업데이트 적용 완료", log_file)
        return True

    except Exception as e:
        log(f"ERROR: {e}\n{traceback.format_exc()}", log_file)
        # 롤백 시도
        try:
            if backup_dir.exists() and not paths['app_exe'].exists():
                if app_dir.exists():
                    shutil.rmtree(app_dir, ignore_errors=True)
                os.rename(str(backup_dir), str(app_dir))
                log("롤백 완료", log_file)
        except Exception as e2:
            log(f"롤백 실패: {e2}", log_file)
        return False


def launch_app(paths):
    """app/StockDashboard.exe 실행"""
    log_file = paths['log_file']
    app_exe = paths['app_exe']

    if not app_exe.exists():
        log(f"FATAL: app EXE 없음 ({app_exe})", log_file)
        # 사용자에게 안내
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                None,
                f"앱 파일을 찾을 수 없습니다.\n\n경로: {app_exe}\n\nGitHub에서 최신 버전을 다운로드해 재설치해주세요.",
                "StockDashboard 오류",
                0x10  # MB_ICONERROR
            )
        except Exception:
            pass
        return False

    log(f"앱 실행: {app_exe}", log_file)
    try:
        # 런처 종료 후에도 앱은 계속 실행
        DETACHED_PROCESS = 0x00000008
        subprocess.Popen(
            [str(app_exe)],
            cwd=str(paths['app_dir']),
            creationflags=DETACHED_PROCESS,
            close_fds=True,
        )
        return True
    except Exception as e:
        log(f"앱 실행 실패: {e}", log_file)
        return False


def main():
    paths = get_paths()
    log_file = paths['log_file']

    log("========== Launcher Start ==========", log_file)
    log(f"Launcher dir: {paths['launcher_dir']}", log_file)
    log(f"App dir:      {paths['app_dir']}", log_file)
    log(f"Pending ZIP:  {paths['pending_zip']}", log_file)

    # 1. 업데이트 적용 (있을 경우)
    apply_update(paths)

    # 2. 앱 실행
    success = launch_app(paths)

    log("========== Launcher End ==========", log_file)

    # 런처는 즉시 종료 (앱은 백그라운드 실행)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
