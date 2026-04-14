"""
내 주식 대시보드 v1.0 Beta
실행 진입점 — PyInstaller EXE / 로컬 직접 실행 / Render·Fly.io 클라우드 배포 모두 지원
"""
import sys, os, threading, time, socket

APP_VERSION = "v1.0 Beta"

# 클라우드 환경 감지 (Render / Fly.io / 수동 서버모드)
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

def open_browser(port, delay=2.0):
    import webbrowser
    time.sleep(delay)
    webbrowser.open(f'http://localhost:{port}')

def main():
    from app import app

    # 포트 결정: 환경변수 PORT → 자동 탐색
    port = int(os.environ.get('PORT', 0)) or find_free_port()
    # 바인드 주소: 서버 배포는 0.0.0.0, 로컬은 127.0.0.1
    host = '0.0.0.0' if IS_SERVER else '127.0.0.1'

    print("=" * 52)
    print(f"  내 주식 대시보드  {APP_VERSION}")
    print("=" * 52)
    if IS_SERVER:
        print(f"  서버 모드  →  0.0.0.0:{port}")
    else:
        print(f"  로컬 모드  →  http://localhost:{port}")
        print(f"  종료: 이 창을 닫거나 Ctrl+C")
        threading.Thread(target=open_browser, args=(port,), daemon=True).start()
    print("=" * 52)

    app.run(host=host, port=port, debug=False, use_reloader=False)

if __name__ == '__main__':
    main()
