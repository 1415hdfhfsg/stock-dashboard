"""
내 주식 대시보드 v1.4.1
실행 진입점 — 데스크탑 앱 (pywebview) / 클라우드 배포 모두 지원
"""
import sys, os, threading, time, socket

APP_VERSION = "v1.4.1"

IS_SERVER = bool(os.environ.get('RENDER') or os.environ.get('FLY') or os.environ.get('SERVER_MODE'))

def find_free_port(start=5000, end=5099):
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('0.0.0.0', port))
                return port
            except OSError:
                continue
    return 5000

def main():
    from app import app

    port = int(os.environ.get('PORT', 0)) or find_free_port()
    host = '0.0.0.0' if IS_SERVER else '127.0.0.1'

    print("=" * 52)
    print(f"  내 주식 대시보드  {APP_VERSION}")
    print("=" * 52)

    if IS_SERVER:
        print(f"  서버 모드  →  0.0.0.0:{port}")
        print("=" * 52)
        app.run(host=host, port=port, debug=False, use_reloader=False)
    else:
        # 데스크탑 앱 모드: pywebview 창으로 실행
        print(f"  데스크탑 모드  →  localhost:{port}")
        print(f"  종료: 창을 닫으면 자동 종료")
        print("=" * 52)

        # Flask 서버를 백그라운드 스레드로
        def run_flask():
            app.run(host=host, port=port, debug=False, use_reloader=False)

        flask_thread = threading.Thread(target=run_flask, daemon=True)
        flask_thread.start()

        # Flask 서버 준비 대기
        time.sleep(1.5)

        try:
            import webview
            # 쿠키/스토리지 영속 경로: AppData\StockDashboard\webview
            appdata = os.environ.get('LOCALAPPDATA') or os.environ.get('APPDATA') or os.path.expanduser('~')
            storage_path = os.path.join(appdata, 'StockDashboard', 'webview')
            os.makedirs(storage_path, exist_ok=True)

            webview.create_window(
                '내 주식 대시보드',
                f'http://localhost:{port}',
                width=1400,
                height=900,
                min_size=(800, 600),
            )
            # private_mode=False + storage_path → 쿠키·로컬스토리지 영속화
            # debug 모드: 환경변수 STOCKDASH_DEBUG=1 설정 시에만 활성화 (기본 비활성)
            _DEBUG_MODE = bool(os.environ.get('STOCKDASH_DEBUG'))
            webview.start(private_mode=False, storage_path=storage_path, debug=_DEBUG_MODE)
        except ImportError:
            # pywebview 없으면 브라우저 폴백
            import webbrowser
            webbrowser.open(f'http://localhost:{port}')
            print("  pywebview 미설치 → 브라우저로 열림")
            flask_thread.join()

if __name__ == '__main__':
    main()
