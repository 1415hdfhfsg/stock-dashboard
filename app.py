from flask import Flask, render_template, jsonify, request, g
import yfinance as yf
from pykrx import stock as krx
from datetime import datetime, timedelta
import json, os, sys, requests, feedparser, sqlite3, uuid
from xml.etree import ElementTree as ET
import openpyxl
import pandas as pd
import numpy as np
from flask.json.provider import DefaultJSONProvider

# ── DB 모드 결정 (PostgreSQL 우선, 없으면 SQLite) ─────────
DATABASE_URL = os.environ.get('DATABASE_URL', '')
IS_POSTGRES  = bool(DATABASE_URL and 'postgres' in DATABASE_URL)
if IS_POSTGRES:
    import psycopg2, psycopg2.extras
    PH = '%s'   # PostgreSQL 파라미터 플레이스홀더
else:
    PH = '?'    # SQLite 파라미터 플레이스홀더

# ── 버전 정보 ─────────────────────────────────────────────
APP_VERSION = "v1.0 Beta"
APP_NAME    = "내 주식 대시보드"

# ── PyInstaller 번들 환경 대응 ───────────────────────────
def _resource(rel):
    """templates 등 읽기 전용 리소스 경로 (번들 내부)"""
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, rel)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), rel)

def _get_data_dir():
    """
    데이터 저장 디렉토리 결정 순서:
      1. Render / Fly.io 클라우드: /data (환경변수 RENDER 또는 FLY 존재 시)
      2. EXE 옆 config.json 의 data_dir (PyInstaller 빌드)
      3. EXE 옆 폴더 (PyInstaller 기본값)
      4. 소스 실행 시 스크립트 폴더
    """
    # 클라우드 배포 (Render / Fly.io): 영구 디스크 /data 사용
    if os.environ.get('RENDER') or os.environ.get('FLY'):
        d = '/data'
        os.makedirs(d, exist_ok=True)
        return d

    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        config_path = os.path.join(exe_dir, 'config.json')
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as _f:
                    cfg = json.load(_f)
                data_dir = cfg.get('data_dir', '').strip()
                if data_dir:
                    os.makedirs(data_dir, exist_ok=True)
                    return data_dir
            except Exception:
                pass
        return exe_dir
    return os.path.dirname(os.path.abspath(__file__))

DATA_DIR = _get_data_dir()

def _data(rel):
    """쓰기 가능한 데이터 파일 경로"""
    return os.path.join(DATA_DIR, rel)

DATABASE = _data('dashboard.db')

class NumpyJSONProvider(DefaultJSONProvider):
    def default(self, obj):
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        return super().default(obj)

app = Flask(__name__,
    template_folder=_resource('templates'),
    static_folder=_resource('static') if os.path.isdir(_resource('static')) else None,
)
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.json_provider_class = NumpyJSONProvider
app.json = NumpyJSONProvider(app)

@app.after_request
def add_no_cache_headers(response):
    if request.path == '/' or request.path.endswith('.html'):
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
    return response

BASE_DIR = _resource('.')

# ── DB 초기화 ─────────────────────────────────────────────
_CREATE_TABLE = '''
    CREATE TABLE IF NOT EXISTS user_data (
        user_token        TEXT PRIMARY KEY,
        portfolio         TEXT NOT NULL DEFAULT '{{"holdings":[]}}',
        transactions      TEXT NOT NULL DEFAULT '[]',
        wishlist          TEXT NOT NULL DEFAULT '[]',
        hidden            TEXT NOT NULL DEFAULT '[]',
        target_prices     TEXT NOT NULL DEFAULT '{{}}',
        notes             TEXT NOT NULL DEFAULT '{{}}',
        rebalance_targets TEXT NOT NULL DEFAULT '{{}}',
        created_at        TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at        TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
'''

def init_db():
    if IS_POSTGRES:
        con = psycopg2.connect(DATABASE_URL)
        cur = con.cursor()
        cur.execute(_CREATE_TABLE.replace('{{', '{').replace('}}', '}'))
        for col, default in [
            ('target_prices',     "'{}'"),
            ('notes',             "'{}'"),
            ('rebalance_targets', "'{}'"),
        ]:
            try:
                cur.execute(f"ALTER TABLE user_data ADD COLUMN IF NOT EXISTS {col} TEXT NOT NULL DEFAULT {default}")
            except Exception:
                con.rollback()
        con.commit()
        cur.close(); con.close()
    else:
        os.makedirs(DATA_DIR, exist_ok=True)
        con = sqlite3.connect(DATABASE)
        con.execute(_CREATE_TABLE.replace('{{', '{').replace('}}', '}'))
        for col, default in [
            ('target_prices',     "'{}'"),
            ('notes',             "'{}'"),
            ('rebalance_targets', "'{}'"),
        ]:
            try:
                con.execute(f"ALTER TABLE user_data ADD COLUMN {col} TEXT NOT NULL DEFAULT {default}")
            except sqlite3.OperationalError:
                pass
        con.commit(); con.close()

init_db()

# ── 요청별 DB 연결 ────────────────────────────────────────
def get_db():
    if 'db' not in g:
        if IS_POSTGRES:
            g.db = psycopg2.connect(DATABASE_URL)
            g.db.autocommit = False
        else:
            g.db = sqlite3.connect(DATABASE)
            g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db:
        db.close()

# ── 사용자 토큰 미들웨어 ──────────────────────────────────
@app.before_request
def load_user_token():
    """쿠키에서 UUID 토큰 읽기. 없으면 새로 발급."""
    token = request.cookies.get('user_token', '')
    if len(token) != 36 or token.count('-') != 4:
        token = str(uuid.uuid4())
        g._new_token = token
    else:
        g._new_token = None
    g.user_token = token

@app.after_request
def set_user_token_cookie(response):
    """신규 사용자에게 토큰 쿠키 발급 (1년 유효)"""
    tok = g.get('_new_token')
    if tok:
        response.set_cookie(
            'user_token', tok,
            max_age=365 * 24 * 3600,
            httponly=True, samesite='Lax'
        )
    return response

def _upsert(field: str, value):
    """user_data 특정 필드 upsert (SQLite / PostgreSQL 공통)"""
    db  = get_db()
    sql = f'''
        INSERT INTO user_data (user_token, {field})
        VALUES ({PH}, {PH})
        ON CONFLICT(user_token) DO UPDATE SET
            {field} = EXCLUDED.{field},
            updated_at = CURRENT_TIMESTAMP
    '''
    if IS_POSTGRES:
        cur = db.cursor()
        cur.execute(sql, (g.user_token, json.dumps(value, ensure_ascii=False)))
        db.commit(); cur.close()
    else:
        db.execute(sql, (g.user_token, json.dumps(value, ensure_ascii=False)))
        db.commit()

def _fetch(field: str, default):
    """user_data 특정 필드 조회 (SQLite / PostgreSQL 공통)"""
    db  = get_db()
    sql = f'SELECT {field} FROM user_data WHERE user_token={PH}'
    if IS_POSTGRES:
        cur = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql, (g.user_token,))
        row = cur.fetchone(); cur.close()
        return json.loads(row[field]) if row else default
    else:
        row = db.execute(sql, (g.user_token,)).fetchone()
        return json.loads(row[field]) if row else default

def load_portfolio():
    return _fetch('portfolio', {'holdings': []})

def save_portfolio(data):
    _upsert('portfolio', data)

def load_transactions():
    return _fetch('transactions', [])

def save_transactions(txs):
    _upsert('transactions', txs)

def load_wishlist():
    return _fetch('wishlist', [])

def save_wishlist(items):
    _upsert('wishlist', items)

def load_hidden():
    return _fetch('hidden', [])

def save_hidden(items):
    _upsert('hidden', items)

def load_target_prices():
    return _fetch('target_prices', {})

def save_target_prices(d):
    _upsert('target_prices', d)

def load_notes():
    return _fetch('notes', {})

def save_notes(d):
    _upsert('notes', d)

def load_rebalance_targets():
    return _fetch('rebalance_targets', {})

def save_rebalance_targets(d):
    _upsert('rebalance_targets', d)

def rebuild_portfolio_from_transactions():
    """거래내역에서 포트폴리오 재계산 (평단가·수량 자동 업데이트)"""
    txs = load_transactions()
    # ticker 기준으로 그룹화
    positions = {}  # ticker → {name, market, qty, total_cost, realized_pnl}
    for tx in sorted(txs, key=lambda x: x['date']):
        tk = tx['ticker']
        if tk not in positions:
            positions[tk] = {
                'name': tx['name'], 'ticker': tk, 'market': tx['market'],
                'qty': 0.0, 'total_cost': 0.0, 'realized_pnl': 0.0
            }
        p = positions[tk]
        if tx['type'] == 'buy':
            new_qty   = p['qty'] + tx['qty']
            p['total_cost'] = p['total_cost'] + tx['qty'] * tx['price']
            p['qty']  = new_qty
        else:  # sell
            if p['qty'] > 0:
                sell_qty = min(tx['qty'], p['qty'])  # 보유량 초과 매도 방지
                avg = p['total_cost'] / p['qty']
                p['realized_pnl'] += (tx['price'] - avg) * sell_qty
                p['total_cost'] -= avg * sell_qty
            p['qty'] = max(0.0, p['qty'] - tx['qty'])
            if p['qty'] <= 0.000001:
                p['qty'] = 0.0
                p['total_cost'] = 0.0  # total_cost 음수 방지

    # 수량 0인 종목 제외, portfolio.json 형식으로 변환
    holdings = []
    for p in positions.values():
        if p['qty'] > 0.000001:
            holdings.append({
                'name':   p['name'],
                'ticker': p['ticker'],
                'market': p['market'],
                'qty':    round(p['qty'], 8),
                'cost':   round(p['total_cost']),
            })
    save_portfolio({'holdings': holdings})
    return holdings

def get_usd_krw():
    try:
        t = yf.Ticker("KRW=X")
        return t.info.get('regularMarketPrice') or t.fast_info.get('lastPrice', 1450)
    except:
        return 1450

@app.route('/')
def index():
    return render_template('dashboard.html', version=APP_VERSION)

@app.route('/api/version')
def api_version():
    return jsonify({'version': APP_VERSION, 'name': APP_NAME})

@app.route('/api/settings')
def api_settings():
    return jsonify({
        'data_dir':    DATA_DIR if not IS_POSTGRES else 'Supabase PostgreSQL',
        'db_path':     DATABASE if not IS_POSTGRES else DATABASE_URL.split('@')[-1],
        'user_token':  g.user_token,
        'is_frozen':   getattr(sys, 'frozen', False),
    })

@app.route('/api/settings/data-dir', methods=['POST'])
def api_change_data_dir():
    """데이터 저장 위치 변경 (config.json 업데이트 후 재시작 필요)"""
    data = request.get_json() or {}
    new_dir = data.get('data_dir', '').strip()
    if not new_dir:
        return jsonify({'error': '경로를 입력하세요'}), 400
    try:
        os.makedirs(new_dir, exist_ok=True)
        # config.json 위치 결정
        if getattr(sys, 'frozen', False):
            cfg_path = os.path.join(os.path.dirname(sys.executable), 'config.json')
        else:
            cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
        # 기존 config 불러오기
        cfg = {}
        if os.path.exists(cfg_path):
            with open(cfg_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        cfg['data_dir'] = new_dir
        with open(cfg_path, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return jsonify({'ok': True, 'new_dir': new_dir, 'config_path': cfg_path})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/my-token')
def api_my_token():
    """현재 사용자 토큰 반환 (다른 기기 접속용)"""
    return jsonify({'token': g.user_token})

@app.route('/api/set-token', methods=['POST'])
def api_set_token():
    """다른 기기에서 같은 데이터를 쓰려면 토큰을 입력"""
    data = request.get_json() or {}
    tok  = str(data.get('token', '')).strip()
    if len(tok) != 36 or tok.count('-') != 4:
        return jsonify({'ok': False, 'error': '잘못된 토큰 형식입니다'}), 400
    resp = jsonify({'ok': True})
    resp.set_cookie('user_token', tok, max_age=365*24*3600, httponly=True, samesite='Lax')
    return resp

@app.route('/api/open-data-folder')
def api_open_data_folder():
    """데이터 폴더를 탐색기로 열기 (Windows)"""
    import subprocess
    try:
        subprocess.Popen(['explorer', DATA_DIR])
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

# ── 위시리스트 (서버 DB 저장) ────────────────────────────
@app.route('/api/wishlist', methods=['GET'])
def api_get_wishlist():
    return jsonify({'wishlist': load_wishlist()})

@app.route('/api/wishlist', methods=['POST'])
def api_add_wishlist():
    item = request.get_json() or {}
    if not item.get('ticker'):
        return jsonify({'error': '티커 없음'}), 400
    items = load_wishlist()
    if not any(w['ticker'] == item['ticker'] for w in items):
        items.append({'ticker': item['ticker'], 'name': item.get('name',''), 'market': item.get('market','')})
        save_wishlist(items)
    return jsonify({'ok': True, 'wishlist': items})

@app.route('/api/wishlist/<ticker>', methods=['DELETE'])
def api_del_wishlist(ticker):
    items = [w for w in load_wishlist() if w['ticker'] != ticker]
    save_wishlist(items)
    return jsonify({'ok': True, 'wishlist': items})

# ── 숨김 종목 (서버 DB 저장) ────────────────────────────
@app.route('/api/hidden', methods=['GET'])
def api_get_hidden():
    return jsonify({'hidden': load_hidden()})

@app.route('/api/hidden', methods=['POST'])
def api_add_hidden():
    data  = request.get_json() or {}
    ticker = data.get('ticker', '')
    if not ticker:
        return jsonify({'error': '티커 없음'}), 400
    items = load_hidden()
    if ticker not in items:
        items.append(ticker)
        save_hidden(items)
    return jsonify({'ok': True})

@app.route('/api/hidden/<ticker>', methods=['DELETE'])
def api_del_hidden(ticker):
    items = [t for t in load_hidden() if t != ticker]
    save_hidden(items)
    return jsonify({'ok': True})

@app.route('/api/portfolio')
def api_portfolio():
    data = load_portfolio()
    usd_krw = get_usd_krw()
    results = []
    total_cost = 0
    total_value = 0

    for h in data['holdings']:
        try:
            if h['market'] == 'US':
                t = yf.Ticker(h['ticker'])
                price_usd = t.fast_info.get('lastPrice', 0)
                price_krw = price_usd * usd_krw
                current_value = price_krw * h['qty']
            else:
                today = datetime.now().strftime('%Y%m%d')
                yesterday = (datetime.now() - timedelta(days=7)).strftime('%Y%m%d')
                df = krx.get_market_ohlcv(yesterday, today, h['ticker'])
                if not df.empty:
                    price_krw = int(df.iloc[-1]['종가'])
                else:
                    price_krw = 0
                current_value = price_krw * h['qty']

            cost = h['cost']
            profit = current_value - cost
            profit_pct = (profit / cost * 100) if cost > 0 else 0

            results.append({
                'name': h['name'],
                'ticker': h['ticker'],
                'market': h['market'],
                'qty': h['qty'],
                'avg_price': round(cost / h['qty']) if h['qty'] > 0 else 0,
                'current_price': round(price_krw),
                'cost': cost,
                'current_value': round(current_value),
                'profit': round(profit),
                'profit_pct': round(profit_pct, 2),
            })
            total_cost += cost
            total_value += current_value
        except Exception as e:
            results.append({
                'name': h['name'], 'ticker': h['ticker'], 'market': h['market'],
                'qty': h['qty'], 'error': str(e)
            })

    total_profit = total_value - total_cost
    total_pct = (total_profit / total_cost * 100) if total_cost > 0 else 0

    return jsonify({
        'holdings': results,
        'summary': {
            'total_cost': round(total_cost),
            'total_value': round(total_value),
            'total_profit': round(total_profit),
            'total_profit_pct': round(total_pct, 2),
            'usd_krw': round(usd_krw, 2)
        }
    })

# 국내 업종 대표 종목 (섹터별 주요 종목으로 평균 등락률 계산)
KR_SECTOR_STOCKS = {
    '반도체·전자':    ['005930', '000660', '042700', '066570'],
    '바이오·제약':    ['207940', '068270', '326030', '128940'],
    '자동차·부품':    ['005380', '000270', '012330', '204320'],
    '금융·보험':      ['105560', '055550', '086790', '032830'],
    '인터넷·플랫폼':  ['035420', '035720', '251270', '293490'],
    '화학·에너지':    ['010950', '011170', '096770', '051910'],
    '통신':           ['017670', '030200', '032640'],
    '건설':           ['000720', '028050', '047040'],
    '유통·식품':      ['069960', '004170', '023530', '097950'],
    '철강·소재':      ['004020', '005490', '010780'],
}

# 분야 검색 키워드 → 대표 종목 매핑
SECTOR_SEARCH_MAP = {
    '반도체': {'kr': ['005930','000660','042700','066570','000990'], 'us': ['NVDA','AMD','INTC','SOXX']},
    '바이오': {'kr': ['207940','068270','326030','128940','091990'], 'us': ['MRNA','ABBV','AMGN','XBI']},
    '제약':   {'kr': ['207940','068270','326030','128940'],          'us': ['PFE','JNJ','ABBV','MRNA']},
    'AI':     {'kr': ['005930','042700','035420','377300'],          'us': ['NVDA','MSFT','GOOGL','META']},
    '인공지능':{'kr': ['005930','042700','035420','377300'],         'us': ['NVDA','MSFT','GOOGL','META']},
    '자동차': {'kr': ['005380','000270','012330','204320'],          'us': ['TSLA','F','GM','TM']},
    '전기차': {'kr': ['005380','000270','247540','086520'],          'us': ['TSLA','RIVN','NIO','LCID']},
    '배터리': {'kr': ['247540','006400','051910','086520'],          'us': ['ALB','LTHM','TSLA']},
    '2차전지':{'kr': ['247540','006400','051910','086520'],          'us': ['ALB','LTHM','TSLA']},
    '금융':   {'kr': ['105560','055550','086790','032830','000810'], 'us': ['JPM','GS','BAC','WFC']},
    '은행':   {'kr': ['105560','055550','086790','000810'],          'us': ['JPM','BAC','C','WFC']},
    '인터넷': {'kr': ['035420','035720','251270','293490'],          'us': ['GOOGL','META','AMZN','NFLX']},
    '플랫폼': {'kr': ['035420','035720','251270'],                   'us': ['GOOGL','META','AMZN','UBER']},
    '게임':   {'kr': ['036570','259960','263750','112040'],          'us': ['EA','TTWO','RBLX','ATVI']},
    '화학':   {'kr': ['051910','010950','011170','096770'],          'us': ['LIN','APD','DOW']},
    '에너지': {'kr': ['010950','096770','267250'],                   'us': ['XOM','CVX','COP','XLE']},
    '통신':   {'kr': ['017670','030200','032640'],                   'us': ['T','VZ','TMUS']},
    '건설':   {'kr': ['000720','028050','047040'],                   'us': ['DHI','LEN','PHM']},
    '소비재': {'kr': ['069960','004170','023530'],                   'us': ['AMZN','HD','NKE','MCD']},
    '식품':   {'kr': ['097950','004170','003230'],                   'us': ['MCD','SBUX','KO','PEP']},
    '헬스케어':{'kr': ['207940','068270','128940'],                  'us': ['JNJ','UNH','ABBV','MRK']},
    '엔터':   {'kr': ['035900','041510','122870','352820'],          'us': ['NFLX','DIS','SPOT']},
    '방산':   {'kr': ['047810','064350','000880'],                   'us': ['LMT','RTX','GD','NOC']},
    '철강':   {'kr': ['004020','005490','010780'],                   'us': ['X','NUE','STLD']},
    '조선':   {'kr': ['009540','042660','010140'],                   'us': []},
    '항공':   {'kr': ['003490','020560'],                            'us': ['DAL','UAL','AAL','BA']},
    '부동산': {'kr': ['016360','145720'],                            'us': ['XLRE','AMT','PLD','SPG']},
}

@app.route('/api/sectors')
def api_sectors():
    result = {'kr': [], 'us': []}

    # ── 국내 업종별 등락률 (대표 종목 평균) ───────────────────
    try:
        today     = datetime.now().strftime('%Y%m%d')
        fromdate  = (datetime.now() - timedelta(days=7)).strftime('%Y%m%d')

        for sector_name, tickers in KR_SECTOR_STOCKS.items():
            changes = []
            last_prices = []
            for tk in tickers:
                try:
                    df = krx.get_market_ohlcv(fromdate, today, tk)
                    if df is None or df.empty or len(df) < 2:
                        continue
                    prev = float(df['종가'].iloc[-2])
                    curr = float(df['종가'].iloc[-1])
                    if prev > 0:
                        changes.append((curr - prev) / prev * 100)
                        last_prices.append(curr)
                except:
                    continue
            if changes:
                avg_chg = sum(changes) / len(changes)
                result['kr'].append({
                    'name': sector_name,
                    'change': round(avg_chg, 2),
                    'value': round(sum(last_prices) / len(last_prices)),
                    'stocks': len(changes),
                })
    except Exception as e:
        result['kr_error'] = str(e)

    # ── 미국 섹터 ETF ───────────────────────────────────────
    us_sectors = {
        'XLK': '기술·IT',        'XLF': '금융',        'XLV': '헬스케어',
        'XLE': '에너지',          'XLI': '산업재',       'XLY': '경기소비재',
        'XLP': '필수소비재',      'XLU': '유틸리티',     'XLB': '소재·원자재',
        'XLRE': '부동산(리츠)',    'XLC': '커뮤니케이션', 'XBI': '바이오',
        'SOXX': '반도체',         'ARKK': '혁신·테크',
    }
    try:
        for sym, name in us_sectors.items():
            try:
                t    = yf.Ticker(sym)
                hist = t.history(period='5d')
                if len(hist) < 2:
                    continue
                prev = float(hist['Close'].iloc[-2])
                curr = float(hist['Close'].iloc[-1])
                chg  = (curr - prev) / prev * 100
                result['us'].append({
                    'name': name, 'ticker': sym,
                    'change': round(chg, 2), 'value': round(curr, 2)
                })
            except:
                continue
    except Exception as e:
        result['us_error'] = str(e)

    result['kr'] = sorted(result['kr'], key=lambda x: x['change'], reverse=True)
    result['us'] = sorted(result['us'], key=lambda x: x['change'], reverse=True)
    return jsonify(result)

@app.route('/api/news')
def api_news():
    # 카테고리별 RSS 피드
    categories = {
        '🔥 핫 트렌드': [
            'https://news.google.com/rss/search?q=트렌드+핫이슈&hl=ko&gl=KR&ceid=KR:ko',
            'https://news.google.com/rss/search?q=화제+인기&hl=ko&gl=KR&ceid=KR:ko',
        ],
        '🌍 세계 이슈': [
            'https://news.google.com/rss/search?q=세계+국제+이슈&hl=ko&gl=KR&ceid=KR:ko',
            'https://news.google.com/rss/search?q=미국+중국+글로벌&hl=ko&gl=KR&ceid=KR:ko',
        ],
        '🏛️ 정치·경제': [
            'https://news.google.com/rss/search?q=정치+경제+정책&hl=ko&gl=KR&ceid=KR:ko',
            'https://news.google.com/rss/search?q=금리+환율+연준&hl=ko&gl=KR&ceid=KR:ko',
        ],
        '📈 증시·투자': [
            'https://news.google.com/rss/search?q=주식+증시+코스피&hl=ko&gl=KR&ceid=KR:ko',
            'https://news.google.com/rss/search?q=나스닥+S&P500+투자&hl=ko&gl=KR&ceid=KR:ko',
        ],
    }

    # 키워드 → 종목 매핑
    keyword_stock = {
        '엔비디아': ('NVDA','US'), '테슬라': ('TSLA','US'), '마이크로소프트': ('MSFT','US'),
        'MS': ('MSFT','US'), '애플': ('AAPL','US'), '쿠팡': ('CPNG','US'),
        '한화': ('462330','KR'), '삼성': ('005930','KR'), 'SK하이닉스': ('000660','KR'),
        'AI': None, '반도체': None, '전기차': ('TSLA','US'), '빅테크': None,
        '메타': ('META','US'), '구글': ('GOOGL','US'), '아마존': ('AMZN','US'),
        'nvidia': ('NVDA','US'), 'tesla': ('TSLA','US'), 'apple': ('AAPL','US'),
    }

    result = {}
    seen_all = set()
    for cat, urls in categories.items():
        items = []
        for url in urls:
            try:
                feed = feedparser.parse(url)
                for entry in feed.entries[:8]:
                    title = entry.get('title', '').strip()
                    if not title or title in seen_all:
                        continue
                    seen_all.add(title)
                    # 관련 종목 감지
                    related = []
                    title_lower = title.lower()
                    for kw, stock in keyword_stock.items():
                        if kw.lower() in title_lower and stock:
                            related.append({'name': kw, 'ticker': stock[0], 'market': stock[1]})
                    # 상승/하락 감지 (부정 맥락 우선 처리)
                    up_kw   = ['상승','급등','최고','돌파','호재','강세','매수','회복','반등','극복','개선']
                    down_kw = ['하락','급락','최저','붕괴','악재','약세','매도','위기','우려','폭락','침체']
                    neg_ctx = ['우려','위기','불안','경고','리스크']  # 상승 키워드 무효화 맥락
                    title_has_up   = any(k in title for k in up_kw)
                    title_has_down = any(k in title for k in down_kw)
                    title_has_neg  = any(k in title for k in neg_ctx)
                    if title_has_down or (title_has_neg and not title_has_up):
                        sentiment = 'down'
                    elif title_has_up and not title_has_neg:
                        sentiment = 'up'
                    else:
                        sentiment = 'neutral'
                    items.append({
                        'title': title,
                        'link':  entry.get('link', ''),
                        'date':  entry.get('published', '')[:16] if entry.get('published') else '',
                        'source': entry.get('source', {}).get('title', '') if isinstance(entry.get('source'), dict) else '',
                        'related': related[:3],
                        'sentiment': sentiment,
                    })
            except:
                continue
        result[cat] = items[:6]

    return jsonify({'categories': result})


# ── 국내 종목명 캐시 (서버 기동 후 첫 검색 시 빌드) ────────
_kr_name_cache = {}   # { ticker: name }

def _build_kr_name_cache():
    """pykrx get_market_ohlcv_by_ticker 로 한 번에 종목명 수집"""
    global _kr_name_cache
    if _kr_name_cache:
        return
    try:
        today = datetime.now().strftime('%Y%m%d')
        # 한 번의 API 호출로 모든 종목 데이터 (종목명 포함) 수집
        for mkt in ('KOSPI', 'KOSDAQ'):
            df = krx.get_market_ohlcv(today, today, market=mkt)
            if df is not None and not df.empty:
                for tk in df.index:
                    try:
                        nm = krx.get_market_ticker_name(str(tk))
                        _kr_name_cache[str(tk)] = nm
                    except:
                        pass
    except:
        pass

def _fetch_kr_stock(tk, fromdate, today, name=None):
    """pykrx 종목 OHLCV 조회 후 결과 dict 반환"""
    try:
        df = krx.get_market_ohlcv(fromdate, today, tk)
        if df is None or df.empty: return None
        curr = float(df['종가'].iloc[-1])
        prev = float(df['종가'].iloc[-2]) if len(df) > 1 else curr
        chg  = (curr - prev) / prev * 100 if prev else 0
        nm   = name or _kr_name_cache.get(tk) or krx.get_market_ticker_name(tk)
        return {'ticker': tk, 'name': nm, 'market': 'KR',
                'price_krw': int(curr), 'change_pct': round(chg, 2), 'currency': 'KRW'}
    except:
        return None

def _fetch_us_stock(sym, usd_krw):
    """yfinance 종목 조회 후 결과 dict 반환"""
    try:
        t     = yf.Ticker(sym)
        info  = t.fast_info
        price = info.get('lastPrice') or info.get('regularMarketPrice', 0)
        prev  = info.get('previousClose', price)
        if not price or price <= 0: return None
        chg   = (price - prev) / prev * 100 if prev else 0
        name  = t.info.get('shortName') or t.info.get('longName') or sym
        return {
            'ticker': sym, 'name': name, 'market': 'US',
            'price_usd': round(price, 2),
            'price_krw': round(price * usd_krw),
            'change_pct': round(chg, 2),
            'currency': 'USD',
            'sector':   t.info.get('sector', ''),
            'industry': t.info.get('industry', ''),
            'market_cap': t.info.get('marketCap', 0),
        }
    except:
        return None

# 한국어/영어 이름 → 미국 티커 매핑 (이름으로 해외 종목 검색 지원)
KR_NAME_TO_US = {
    '엔비디아': 'NVDA',  'nvidia': 'NVDA',
    '테슬라':   'TSLA',  'tesla':  'TSLA',
    '애플':     'AAPL',  'apple':  'AAPL',
    '마이크로소프트': 'MSFT', 'microsoft': 'MSFT',
    '구글':     'GOOGL', '알파벳': 'GOOGL', 'google': 'GOOGL', 'alphabet': 'GOOGL',
    '아마존':   'AMZN',  'amazon': 'AMZN',
    '메타':     'META',  '페이스북': 'META', 'meta': 'META', 'facebook': 'META',
    '넷플릭스': 'NFLX',  'netflix': 'NFLX',
    '스타벅스': 'SBUX',  'starbucks': 'SBUX',
    '나이키':   'NKE',   'nike': 'NKE',
    '코카콜라': 'KO',    'coca cola': 'KO',
    '펩시':     'PEP',   'pepsi': 'PEP',
    '월마트':   'WMT',   'walmart': 'WMT',
    '홈디포':   'HD',    'home depot': 'HD',
    '코스트코': 'COST',  'costco': 'COST',
    '비자':     'V',     'visa': 'V',
    '마스터카드': 'MA',  'mastercard': 'MA',
    'JP모건':   'JPM',   'jpmorgan': 'JPM',
    '골드만삭스': 'GS',  'goldman': 'GS',
    '뱅크오브아메리카': 'BAC', 'bank of america': 'BAC',
    '모건스탠리': 'MS',  'morgan stanley': 'MS',
    '엑슨모빌': 'XOM',   'exxon': 'XOM',
    '쉐브론':   'CVX',   'chevron': 'CVX',
    '쿠팡':     'CPNG',  'coupang': 'CPNG',
    '팔란티어': 'PLTR',  'palantir': 'PLTR',
    '리비안':   'RIVN',  'rivian': 'RIVN',
    '루시드':   'LCID',  'lucid': 'LCID',
    '화이자':   'PFE',   'pfizer': 'PFE',
    '모더나':   'MRNA',  'moderna': 'MRNA',
    '존슨앤존슨': 'JNJ', 'johnson': 'JNJ',
    '유나이티드헬스': 'UNH', 'unitedhealth': 'UNH',
    '일라이릴리': 'LLY', 'eli lilly': 'LLY', '릴리': 'LLY',
    '노보노디스크': 'NVO', 'novo nordisk': 'NVO',
    'AMD': 'AMD', 'amd': 'AMD',
    '인텔':     'INTC',  'intel': 'INTC',
    '퀄컴':     'QCOM',  'qualcomm': 'QCOM',
    '브로드컴': 'AVGO',  'broadcom': 'AVGO',
    'TSMC':     'TSM',   'tsmc': 'TSM',
    'ASML':     'ASML',  'asml': 'ASML',
    '마이크론': 'MU',    'micron': 'MU',
    '어플라이드머티리얼즈': 'AMAT', 'applied materials': 'AMAT',
    '디즈니':   'DIS',   'disney': 'DIS',
    '스포티파이': 'SPOT', 'spotify': 'SPOT',
    '우버':     'UBER',  'uber': 'UBER',
    '에어비앤비': 'ABNB', 'airbnb': 'ABNB',
    '오라클':   'ORCL',  'oracle': 'ORCL',
    '세일즈포스': 'CRM', 'salesforce': 'CRM',
    '어도비':   'ADBE',  'adobe': 'ADBE',
    '서비스나우': 'NOW', 'servicenow': 'NOW',
    '스노우플레이크': 'SNOW', 'snowflake': 'SNOW',
    '크라우드스트라이크': 'CRWD', 'crowdstrike': 'CRWD',
    '데이터독': 'DDOG',  'datadog': 'DDOG',
    '코인베이스': 'COIN', 'coinbase': 'COIN',
    '페이팔':   'PYPL',  'paypal': 'PYPL',
    '블록':     'SQ',    'block': 'SQ', '스퀘어': 'SQ',
    '인튜이트': 'INTU',  'intuit': 'INTU',
    '버크셔':   'BRK-B', 'berkshire': 'BRK-B',
    '록히드마틴': 'LMT', 'lockheed': 'LMT',
    '레이시온':  'RTX',  'raytheon': 'RTX',
    'ARM': 'ARM', 'arm': 'ARM',
    '일본소니': 'SONY', 'sony': 'SONY',
    '삼성SDI': '006400',  # 국내 종목 → KR 티커 직접 사용
}

# 국내 인기 종목 이름 → 티커 (빠른 검색용, 이름이 key)
KR_POPULAR_STOCKS = {
    '삼성전자': '005930', 'SK하이닉스': '000660', 'LG에너지솔루션': '373220',
    '삼성바이오로직스': '207940', '현대차': '005380', '기아': '000270',
    'POSCO홀딩스': '005490', 'NAVER': '035420', '카카오': '035720',
    'LG화학': '051910', '삼성SDI': '006400', '현대모비스': '012330',
    'KB금융': '105560', '신한지주': '055550', '삼성물산': '028260',
    'SK이노베이션': '096770', '하나금융지주': '086790', '우리금융지주': '316140',
    'KT&G': '033780', 'SK텔레콤': '017670', 'KT': '030200',
    'LG전자': '066570', '한국전력': '015760', '두산에너빌리티': '034020',
    '셀트리온': '068270', '삼성생명': '032830', 'DB손해보험': '005830',
    '고려아연': '010130', '한화에어로스페이스': '012450', '한화오션': '042660',
    '한미반도체': '042700', 'SK': '034730', 'LG': '003550',
    'S-Oil': '010950', '롯데케미칼': '011170', '금호석유': '011780',
    '카카오뱅크': '323410', '카카오페이': '377300', '크래프톤': '259960',
    'NC소프트': '036570', '엔씨소프트': '036570', '넷마블': '251270',
    '포스코퓨처엠': '003670', '에코프로': '086520', '에코프로비엠': '247540',
    '엘앤에프': '066970', '일진머티리얼즈': '020150', 'SK아이이테크놀로지': '361610',
    '현대건설': '000720', '대우건설': '047040', 'GS건설': '006360',
    '롯데쇼핑': '023530', '신세계': '004170', '이마트': '139480',
    'CJ제일제당': '097950', '오리온': '271560', '농심': '004370',
    '아모레퍼시픽': '090430', 'LG생활건강': '051900', '한국콜마': '161890',
    '대한항공': '003490', '아시아나항공': '020560', '제주항공': '089590',
    'HMM': '011200', '현대글로비스': '086280', '한진칼': '180640',
    '삼성전기': '009150', '엘지이노텍': '011070', '삼성화재': '000810',
    '메리츠화재': '000060', '한국투자증권': '071050', '미래에셋증권': '006800',
    '키움증권': '039490', '삼성증권': '016360', '한국금융지주': '071050',
    '두산': '000150', '두산밥캣': '241560', '한화': '000880',
    '현대중공업': '009540', 'HD현대': '267250', 'HD현대중공업': '329180',
}

@app.route('/api/search')
def api_search():
    q      = request.args.get('q', '').strip()
    market = request.args.get('market', 'all').lower()  # all | kr | us
    if not q:
        return jsonify({'results': []})

    results      = []
    seen_tickers = set()
    today    = datetime.now().strftime('%Y%m%d')
    fromdate = (datetime.now() - timedelta(days=7)).strftime('%Y%m%d')
    usd_krw  = get_usd_krw()
    q_lower  = q.lower()

    def add(item):
        if item and item['ticker'] not in seen_tickers:
            seen_tickers.add(item['ticker'])
            results.append(item)

    # ── 0. 분야/섹터 키워드 검색 ─────────────────────────────
    matched_sector = None
    for kw, info in SECTOR_SEARCH_MAP.items():
        if kw == q or kw in q or q in kw:
            matched_sector = info
            break

    if matched_sector:
        if market in ('all', 'kr'):
            for tk in matched_sector.get('kr', []):
                add(_fetch_kr_stock(tk, fromdate, today))
                if len(results) >= 5: break
        if market in ('all', 'us'):
            for sym in matched_sector.get('us', []):
                add(_fetch_us_stock(sym, usd_krw))
                if len(results) >= 10: break
        return jsonify({'results': results, 'sector_search': True})

    # ── 1. 해외(US) 검색 ────────────────────────────────────
    if market in ('all', 'us'):
        # 직접 티커로 시도
        add(_fetch_us_stock(q.upper(), usd_krw))
        # 한국어/영어 이름 → 티커 매핑
        us_from_name = KR_NAME_TO_US.get(q) or KR_NAME_TO_US.get(q_lower)
        if us_from_name:
            add(_fetch_us_stock(us_from_name, usd_krw))
        # 부분 일치 이름 검색 (KR_NAME_TO_US 키에서 q가 포함된 것 모두)
        for name_kw, sym in KR_NAME_TO_US.items():
            if sym and q_lower in name_kw.lower() and sym not in seen_tickers:
                add(_fetch_us_stock(sym, usd_krw))
                if len([r for r in results if r['market']=='US']) >= 5: break

    # ── 2. 국내(KR) 검색 ────────────────────────────────────
    if market in ('all', 'kr'):
        try:
            if q.isdigit():
                add(_fetch_kr_stock(q.zfill(6), fromdate, today))
            else:
                # 1단계: 사전 정의 인기 종목에서 빠른 검색
                kr_found = 0
                for nm, tk in KR_POPULAR_STOCKS.items():
                    if q_lower in nm.lower():
                        add(_fetch_kr_stock(tk, fromdate, today, name=nm))
                        kr_found += 1
                        if kr_found >= 6: break

                # 2단계: 백그라운드 캐시가 있으면 추가 검색
                if kr_found < 3 and _kr_name_cache:
                    for tk, nm in _kr_name_cache.items():
                        if q_lower in nm.lower() and tk not in seen_tickers:
                            add(_fetch_kr_stock(tk, fromdate, today, name=nm))
                            kr_found += 1
                            if kr_found >= 6: break
        except:
            pass

    # 국내→해외 순 정렬: 전체 검색 시 KR 먼저
    if market == 'all':
        results.sort(key=lambda r: (0 if r['market'] == 'KR' else 1))

    return jsonify({'results': results[:12]})

@app.route('/api/upload', methods=['POST'])
def api_upload():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다'}), 400
    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '.xlsx 파일만 지원됩니다'}), 400

    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active

        # 헤더 찾기
        header_row = None
        col_map = {}
        for row in ws.iter_rows(min_row=1, max_row=10):
            cells = [str(c.value).strip().lower() if c.value else '' for c in row]
            for i, val in enumerate(cells):
                if val in ('종목명', '이름', 'name', '종목'): col_map['name'] = i
                if val in ('종목코드', '코드', 'ticker', 'code', '티커'): col_map['ticker'] = i
                if val in ('시장', 'market', '거래소'): col_map['market'] = i
                if val in ('수량', 'qty', 'quantity', '보유수량', '주수'): col_map['qty'] = i
                if val in ('매수금액', '총매수금액', 'cost', '투자금액', '매수총액', '총투자금액'): col_map['cost'] = i
            if 'name' in col_map and 'ticker' in col_map:
                header_row = row[0].row
                break

        if header_row is None:
            return jsonify({'error': '헤더를 찾을 수 없습니다. 종목명, 종목코드, 시장, 수량, 매수금액 열이 필요합니다.'}), 400

        holdings = []
        for row in ws.iter_rows(min_row=header_row + 1, values_only=False):
            vals = [c.value for c in row]
            name = vals[col_map.get('name', 0)]
            if not name:
                continue
            ticker = str(vals[col_map.get('ticker', 1)] or '')
            market = str(vals[col_map.get('market', 2)] or 'KR').upper()
            if market in ('미국', 'US', 'USA', 'NYSE', 'NASDAQ'):
                market = 'US'
            else:
                market = 'KR'
            qty = float(vals[col_map.get('qty', 3)] or 0)
            cost = float(vals[col_map.get('cost', 4)] or 0)
            if qty > 0:
                holdings.append({
                    'name': str(name),
                    'ticker': ticker,
                    'market': market,
                    'qty': qty,
                    'cost': cost
                })

        if not holdings:
            return jsonify({'error': '유효한 종목 데이터가 없습니다'}), 400

        data = {'holdings': holdings}
        save_portfolio(data)
        return jsonify({'success': True, 'count': len(holdings), 'holdings': holdings})

    except Exception as e:
        return jsonify({'error': f'파일 처리 오류: {str(e)}'}), 500

@app.route('/api/template')
def api_template():
    from io import BytesIO
    from flask import send_file
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '포트폴리오'
    headers = ['종목명', '종목코드', '시장', '수량', '매수금액']
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
        ws.cell(row=1, column=i).font = openpyxl.styles.Font(bold=True)
    ws.append(['삼성전자', '005930', 'KR', 10, 550000])
    ws.append(['엔비디아', 'NVDA', 'US', 5, 500000])
    for col in ['A','B','C','D','E']:
        ws.column_dimensions[col].width = 15
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, download_name='포트폴리오_템플릿.xlsx', as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def calc_indicators(df):
    """8개 기술적 지표 계산 + 점수화"""
    import numpy as np
    close = df['Close']
    high = df['High']
    low = df['Low']
    volume = df['Volume']

    # 1. 이동평균선
    ma5  = close.rolling(5).mean()
    ma20 = close.rolling(20).mean()
    ma60 = close.rolling(60).mean()

    # 2. RSI(14)
    delta = close.diff()
    gain  = delta.clip(lower=0)
    loss  = (-delta).clip(lower=0)
    avg_gain = gain.ewm(com=13, adjust=False).mean()
    avg_loss = loss.ewm(com=13, adjust=False).mean()
    rs  = avg_gain / avg_loss.replace(0, np.nan)
    rsi = 100 - (100 / (1 + rs))

    # 3. MACD(12/26/9)
    ema12  = close.ewm(span=12, adjust=False).mean()
    ema26  = close.ewm(span=26, adjust=False).mean()
    macd   = ema12 - ema26
    signal = macd.ewm(span=9, adjust=False).mean()
    hist   = macd - signal

    # 4. 볼린저밴드(20/2)
    bb_mid   = close.rolling(20).mean()
    bb_std   = close.rolling(20).std()
    bb_upper = bb_mid + 2 * bb_std
    bb_lower = bb_mid - 2 * bb_std

    # 5. ATR(14)
    tr = pd.concat([
        high - low,
        (high - close.shift()).abs(),
        (low  - close.shift()).abs()
    ], axis=1).max(axis=1)
    atr = tr.rolling(14).mean()

    # 6. ADX(14)
    up_move   = high.diff()
    down_move = (-low.diff())
    plus_dm   = np.where((up_move > down_move) & (up_move > 0), up_move, 0.0)
    minus_dm  = np.where((down_move > up_move) & (down_move > 0), down_move, 0.0)
    atr14     = tr.ewm(com=13, adjust=False).mean()
    plus_di   = 100 * pd.Series(plus_dm,  index=df.index).ewm(com=13, adjust=False).mean() / atr14
    minus_di  = 100 * pd.Series(minus_dm, index=df.index).ewm(com=13, adjust=False).mean() / atr14
    dx        = 100 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    adx       = dx.ewm(com=13, adjust=False).mean()

    # 최신값 추출
    c   = close.iloc[-1]
    m5  = ma5.iloc[-1]
    m20 = ma20.iloc[-1]
    m60 = ma60.iloc[-1]
    r   = rsi.iloc[-1]
    mc  = macd.iloc[-1]
    sig = signal.iloc[-1]
    h   = hist.iloc[-1]
    h_prev = hist.iloc[-2] if len(hist) > 1 else h
    bb_u = bb_upper.iloc[-1]
    bb_l = bb_lower.iloc[-1]
    bb_m = bb_mid.iloc[-1]
    atr_val = atr.iloc[-1]
    adx_val = adx.iloc[-1]
    vol5  = volume.iloc[-5:].mean()
    vol20 = volume.iloc[-20:].mean()

    # 52주 고저
    w52_high = high.iloc[-252:].max() if len(high) >= 252 else high.max()
    w52_low  = low.iloc[-252:].min()  if len(low)  >= 252 else low.min()
    w52_pos  = (c - w52_low) / (w52_high - w52_low) * 100 if w52_high != w52_low else 50

    # ── 점수화 ──────────────────────────────────────────────
    scores = {}

    # ① 이평선 배열
    if m5 > m20 and m20 > m60:   scores['ma_align'] = 2
    elif m5 > m20 or m20 > m60:  scores['ma_align'] = 1
    elif m5 < m20 and m20 < m60: scores['ma_align'] = -2
    elif m5 < m20 or m20 < m60:  scores['ma_align'] = -1
    else:                          scores['ma_align'] = 0

    # ② 골든/데드크로스 (최근 10일 내 교차 감지)
    cross = 0
    for i in range(1, min(11, len(ma5))):
        prev5 = ma5.iloc[-i-1]; prev20 = ma20.iloc[-i-1]
        cur5  = ma5.iloc[-i];   cur20  = ma20.iloc[-i]
        if prev5 <= prev20 and cur5 > cur20:
            # 최근일수록 강한 신호 (1일 전=+2, 10일 전=+1)
            cross = 2 if i <= 3 else 1; break
        elif prev5 >= prev20 and cur5 < cur20:
            cross = -2 if i <= 3 else -1; break
    if cross == 0:
        cross = 1 if m5 > m20 else (-1 if m5 < m20 else 0)
    scores['cross'] = cross

    # ③ RSI
    if   r < 30:  scores['rsi'] = 2
    elif r < 45:  scores['rsi'] = 1
    elif r < 55:  scores['rsi'] = 0
    elif r < 70:  scores['rsi'] = -1
    else:          scores['rsi'] = -2

    # ④ MACD
    if   mc > sig and h > h_prev: scores['macd'] = 2
    elif mc > sig:                 scores['macd'] = 1
    elif mc < sig and h < h_prev: scores['macd'] = -2
    elif mc < sig:                 scores['macd'] = -1
    else:                          scores['macd'] = 0

    # ⑤ 볼린저밴드
    bb_range = bb_u - bb_l if bb_u != bb_l else 1
    bb_pos   = (c - bb_l) / bb_range  # 0~1
    prev_c   = close.iloc[-2] if len(close) > 1 else c
    if   bb_pos < 0.15 and c > prev_c: scores['bb'] = 2
    elif bb_pos < 0.35:                 scores['bb'] = 1
    elif bb_pos < 0.65:                 scores['bb'] = 0
    elif bb_pos < 0.85:                 scores['bb'] = -1
    else:                               scores['bb'] = -2

    # ⑥ 52주 위치
    if   w52_pos < 20:  scores['w52'] = 2
    elif w52_pos < 40:  scores['w52'] = 1
    elif w52_pos < 60:  scores['w52'] = 0
    elif w52_pos < 80:  scores['w52'] = -1
    else:               scores['w52'] = -2

    # ⑦ 거래량 트렌드
    vr = vol5 / vol20 if vol20 > 0 else 1
    if   vr > 1.5:  scores['volume'] = 1
    elif vr > 0.9:  scores['volume'] = 0
    else:           scores['volume'] = -1

    # ⑧ RSI 다이버전스 (최근 10일)
    divergence = 0
    if len(close) >= 10:
        price_hi = close.iloc[-10:].max(); price_lo = close.iloc[-10:].min()
        rsi_hi   = rsi.iloc[-10:].max();   rsi_lo   = rsi.iloc[-10:].min()
        if close.iloc[-1] >= price_hi * 0.99 and rsi.iloc[-1] < rsi_hi * 0.97:
            divergence = -2  # 하락 다이버전스
        elif close.iloc[-1] <= price_lo * 1.01 and rsi.iloc[-1] > rsi_lo * 1.03:
            divergence = 2   # 상승 다이버전스

    # ADX 신뢰도
    if   adx_val > 25: reliability = 'HIGH'
    elif adx_val > 20: reliability = 'MEDIUM'
    else:               reliability = 'LOW'

    total = sum(scores.values()) + divergence

    # 등급
    if   total >= 9:  grade, grade_ko = 'strong_up',   '강한 상승'
    elif total >= 4:  grade, grade_ko = 'up',           '상승'
    elif total >= -3: grade, grade_ko = 'neutral',      '중립'
    elif total >= -8: grade, grade_ko = 'down',         '하락'
    else:             grade, grade_ko = 'strong_down',  '강한 하락'

    # 예상 범위
    short_high = round(c + atr_val * 1.5)
    short_low  = round(c - atr_val * 1.5)
    mid_high   = round(bb_u)
    mid_low    = round(bb_l)
    # 이평선 기울기로 중기 범위 조정
    slope = (m20 - ma20.iloc[-6]) / ma20.iloc[-6] if len(ma20) >= 6 and ma20.iloc[-6] > 0 else 0
    adj   = round(c * slope * 3)
    mid_high += adj; mid_low += adj

    # 미니차트용 데이터 (최근 60일)
    n = min(60, len(close))
    chart = {
        'dates':  [str(d.date()) for d in close.iloc[-n:].index],
        'close':  [round(v, 2) for v in close.iloc[-n:].tolist()],
        'ma5':    [round(v, 2) if not pd.isna(v) else None for v in ma5.iloc[-n:].tolist()],
        'ma20':   [round(v, 2) if not pd.isna(v) else None for v in ma20.iloc[-n:].tolist()],
        'ma60':   [round(v, 2) if not pd.isna(v) else None for v in ma60.iloc[-n:].tolist()],
    }

    return {
        'total_score': total,
        'grade': grade,
        'grade_ko': grade_ko,
        'reliability': reliability,
        'adx': round(float(adx_val), 1),
        'rsi': round(float(r), 1),
        'current_price': round(float(c)),
        'atr': round(float(atr_val)),
        'scores': scores,
        'divergence': divergence,
        'short_range': {'high': int(short_high), 'low': int(short_low)},
        'mid_range':   {'high': int(mid_high),   'low': int(mid_low)},
        'w52': {'high': round(float(w52_high)), 'low': round(float(w52_low)), 'pos': round(float(w52_pos), 1)},
        'chart': chart,
        'indicators': {
            '이평선 배열': {'score': scores['ma_align'],  'desc': f'MA5={round(m5):,} MA20={round(m20):,} MA60={round(m60):,}'},
            '골든/데드크로스': {'score': scores['cross'], 'desc': f'MA5 {">" if m5>m20 else "<"} MA20'},
            'RSI(14)':     {'score': scores['rsi'],     'desc': f'RSI={r:.1f}'},
            'MACD':        {'score': scores['macd'],    'desc': f'MACD={mc:.2f} Signal={sig:.2f}'},
            '볼린저밴드':  {'score': scores['bb'],      'desc': f'상단={round(bb_u):,} 중심={round(bb_m):,} 하단={round(bb_l):,}'},
            '52주 위치':   {'score': scores['w52'],     'desc': f'{w52_pos:.0f}% (저={round(w52_low):,} 고={round(w52_high):,})'},
            '거래량':      {'score': scores['volume'],  'desc': f'5일평균/20일평균={vr:.2f}x'},
        }
    }


@app.route('/api/analysis')
def api_analysis():
    data = load_portfolio()
    usd_krw = get_usd_krw()
    results = []

    for h in data['holdings']:
        try:
            if h['market'] == 'US':
                t = yf.Ticker(h['ticker'])
                df = t.history(period='1y')
                if df.empty or len(df) < 30:
                    raise ValueError('데이터 부족')
                df.index = df.index.tz_localize(None)
                # 원화 환산
                df['Close'] = df['Close'] * usd_krw
                df['High']  = df['High']  * usd_krw
                df['Low']   = df['Low']   * usd_krw
            else:
                end   = datetime.now().strftime('%Y%m%d')
                start = (datetime.now() - timedelta(days=400)).strftime('%Y%m%d')
                raw   = krx.get_market_ohlcv(start, end, h['ticker'])
                if raw.empty or len(raw) < 30:
                    raise ValueError('데이터 부족')
                df = raw.rename(columns={'시가':'Open','고가':'High','저가':'Low','종가':'Close','거래량':'Volume'})

            result = calc_indicators(df)
            result['name']   = h['name']
            result['ticker'] = h['ticker']
            result['market'] = h['market']
            results.append(result)

        except Exception as e:
            results.append({
                'name': h['name'], 'ticker': h['ticker'], 'market': h['market'],
                'error': str(e)
            })

    return jsonify({'analysis': results})


# ────────────────────────────────────────────────────────────
# 알고리즘 비교용 4개 알고리즘
# ────────────────────────────────────────────────────────────

def algo_comprehensive(df):
    """Algo A: 종합형 — 8지표 조합 (73~82% 실증 승률)"""
    result = calc_indicators(df)
    score = result['total_score']
    detail = {k: v['score'] for k, v in result['indicators'].items()}
    detail['다이버전스'] = result['divergence']
    return {
        'name': 'Algo A: 종합형',
        'desc': '8개 지표 + 다이버전스 + ADX 신뢰도 필터',
        'basis': '73~82% 실증 승률 (RSI+MACD+거래량+ADX 조합)',
        'score': score,
        'max_score': 15,
        'grade': result['grade'],
        'grade_ko': result['grade_ko'],
        'reliability': result['reliability'],
        'adx': result['adx'],
        'detail': detail,
        'short_range': result['short_range'],
        'mid_range': result['mid_range'],
        'strength': '다양한 시장 상황에 균형적으로 대응',
        'weakness': '지표 충돌 시 중립 신호가 많아질 수 있음',
    }


def algo_trend_follow(df):
    """Algo B: 추세 추종형 — MA 정배열 + ADX + MACD (추세장 특화)"""
    import numpy as np
    close  = df['Close']
    high   = df['High']
    low    = df['Low']

    ma5  = close.rolling(5).mean()
    ma20 = close.rolling(20).mean()
    ma60 = close.rolling(60).mean()
    ma120= close.rolling(120).mean()
    ema12= close.ewm(span=12, adjust=False).mean()
    ema26= close.ewm(span=26, adjust=False).mean()
    macd = ema12 - ema26
    sig  = macd.ewm(span=9, adjust=False).mean()
    hist = macd - sig

    tr = pd.concat([(high-low),(high-close.shift()).abs(),(low-close.shift()).abs()],axis=1).max(axis=1)
    up_m  = high.diff(); dn_m = (-low.diff())
    pdm   = np.where((up_m>dn_m)&(up_m>0), up_m, 0.0)
    ndm   = np.where((dn_m>up_m)&(dn_m>0), dn_m, 0.0)
    atr14 = tr.ewm(com=13, adjust=False).mean()
    pdi   = 100*pd.Series(pdm,index=df.index).ewm(com=13,adjust=False).mean()/atr14
    ndi   = 100*pd.Series(ndm,index=df.index).ewm(com=13,adjust=False).mean()/atr14
    dx    = 100*(pdi-ndi).abs()/(pdi+ndi).replace(0,np.nan)
    adx   = dx.ewm(com=13,adjust=False).mean()

    c = close.iloc[-1]
    m5=ma5.iloc[-1]; m20=ma20.iloc[-1]; m60=ma60.iloc[-1]; m120=ma120.iloc[-1]
    mc=macd.iloc[-1]; sg=sig.iloc[-1]; h=hist.iloc[-1]; hp=hist.iloc[-2] if len(hist)>1 else h
    adx_v=adx.iloc[-1]; atr_v=tr.rolling(14).mean().iloc[-1]

    detail = {}
    score = 0

    # ① 완전 정배열 4단계 (5>20>60>120)
    if m5>m20>m60>m120:    s=4; txt='완전 정배열 (4단계)'
    elif m5>m20>m60:       s=3; txt='정배열 (3단계)'
    elif m5>m20:           s=2; txt='부분 정배열'
    elif m5<m20<m60<m120:  s=-4; txt='완전 역배열 (4단계)'
    elif m5<m20<m60:       s=-3; txt='역배열 (3단계)'
    elif m5<m20:           s=-2; txt='부분 역배열'
    else:                  s=0; txt='혼조'
    detail['이평선 배열'] = {'score': s, 'desc': txt}
    score += s

    # ② ADX 추세 강도 (0~3점)
    if   adx_v > 40: s=3; txt=f'ADX {adx_v:.1f} - 매우 강한 추세'
    elif adx_v > 25: s=2; txt=f'ADX {adx_v:.1f} - 강한 추세'
    elif adx_v > 20: s=1; txt=f'ADX {adx_v:.1f} - 추세 형성 중'
    else:            s=0; txt=f'ADX {adx_v:.1f} - 횡보장 (신호 약함)'
    detail['ADX 추세강도'] = {'score': s, 'desc': txt}
    score += s

    # ③ MACD 방향 + 히스토그램 가속도
    if   mc>sg and h>hp and h>0: s=3; txt='MACD 상승 + 히스토그램 확대'
    elif mc>sg and h>0:          s=2; txt='MACD 상승'
    elif mc>sg:                  s=1; txt='MACD>시그널 (히스토그램 감소)'
    elif mc<sg and h<hp and h<0: s=-3; txt='MACD 하락 + 히스토그램 확대'
    elif mc<sg and h<0:          s=-2; txt='MACD 하락'
    else:                        s=-1; txt='MACD<시그널'
    detail['MACD'] = {'score': s, 'desc': txt}
    score += s

    # ④ 현재가 vs 이평선 위치
    above = sum([c>m5, c>m20, c>m60, c>m120])
    s = above - 2  # -2 ~ +2
    detail['이평선 대비 위치'] = {'score': s, 'desc': f'이평선 {above}개 위에 위치'}
    score += s

    max_score = 12
    grade, grade_ko = _score_to_grade(score, max_score)
    atr_v_safe = float(atr_v) if not pd.isna(atr_v) else c * 0.02
    return {
        'name': 'Algo B: 추세 추종형',
        'desc': '이평선 4단계 배열 + ADX + MACD 히스토그램 가속도',
        'basis': '추세장(ADX>25) 집중, 데드크로스/골든크로스 강화',
        'score': score, 'max_score': max_score,
        'grade': grade, 'grade_ko': grade_ko,
        'reliability': 'HIGH' if adx_v > 25 else ('MEDIUM' if adx_v > 20 else 'LOW'),
        'adx': round(float(adx_v), 1), 'detail': detail,
        'short_range': {'high': round(c + atr_v_safe*1.5), 'low': round(c - atr_v_safe*1.5)},
        'mid_range':   {'high': round(c + atr_v_safe*4),   'low': round(c - atr_v_safe*4)},
        'strength': '추세장에서 매우 강력, 노이즈 적음',
        'weakness': '횡보장(ADX<20)에서 신호 부정확',
    }


def algo_contrarian(df):
    """Algo C: 역발상형 — RSI 극단 + 볼린저 + 52주 저점 (반등 특화)"""
    import numpy as np
    close = df['Close']; high = df['High']; low = df['Low']

    delta = close.diff()
    gain  = delta.clip(lower=0); loss = (-delta).clip(lower=0)
    rsi   = 100 - 100/(1+close.ewm(com=13,adjust=False).mean().pipe(lambda _:
            gain.ewm(com=13,adjust=False).mean()/loss.ewm(com=13,adjust=False).mean().replace(0,np.nan)))
    # 간단 재계산
    ag = gain.ewm(com=13,adjust=False).mean()
    al = loss.ewm(com=13,adjust=False).mean()
    rs = ag / al.replace(0, np.nan)
    rsi = 100 - (100/(1+rs))

    bb_mid  = close.rolling(20).mean()
    bb_std  = close.rolling(20).std()
    bb_up   = bb_mid + 2*bb_std
    bb_lo   = bb_mid - 2*bb_std

    w52_hi = high.iloc[-252:].max() if len(high)>=252 else high.max()
    w52_lo = low.iloc[-252:].min()  if len(low)>=252  else low.min()
    tr     = pd.concat([(high-low),(high-close.shift()).abs(),(low-close.shift()).abs()],axis=1).max(axis=1)
    atr_v  = tr.rolling(14).mean().iloc[-1]

    c=close.iloc[-1]; r=rsi.iloc[-1]; r_prev=rsi.iloc[-5] if len(rsi)>5 else r
    bu=bb_up.iloc[-1]; bl=bb_lo.iloc[-1]; bm=bb_mid.iloc[-1]
    w52_pos = (c-w52_lo)/(w52_hi-w52_lo)*100 if w52_hi!=w52_lo else 50

    detail = {}; score = 0

    # ① RSI 극단값 (과매도/과매수)
    if   r < 20:  s=5; txt=f'RSI {r:.1f} - 심각한 과매도 (강한 반등 기대)'
    elif r < 30:  s=3; txt=f'RSI {r:.1f} - 과매도'
    elif r < 40:  s=1; txt=f'RSI {r:.1f} - 약한 과매도'
    elif r > 80:  s=-5; txt=f'RSI {r:.1f} - 심각한 과매수 (조정 경고)'
    elif r > 70:  s=-3; txt=f'RSI {r:.1f} - 과매수'
    elif r > 60:  s=-1; txt=f'RSI {r:.1f} - 약한 과매수'
    else:         s=0; txt=f'RSI {r:.1f} - 중립'
    detail['RSI 극단값'] = {'score': s, 'desc': txt}
    score += s

    # ② RSI 방향 전환 (상승 전환 중인지)
    rsi_trend = r - r_prev
    if rsi_trend > 5 and r < 50: s=2; txt=f'과매도 구간서 RSI 상승 전환 (+{rsi_trend:.1f})'
    elif rsi_trend < -5 and r > 50: s=-2; txt=f'과매수 구간서 RSI 하락 전환 ({rsi_trend:.1f})'
    else: s=0; txt='RSI 방향 전환 없음'
    detail['RSI 방향 전환'] = {'score': s, 'desc': txt}
    score += s

    # ③ 볼린저밴드 위치 (역발상: 하단 = 매수 기회)
    bb_range = bu - bl if bu != bl else 1
    bb_pos   = (c - bl) / bb_range
    if   bb_pos < 0.05: s=4; txt='볼린저 하단 이탈 (강한 반등 신호)'
    elif bb_pos < 0.20: s=3; txt='볼린저 하단 근접'
    elif bb_pos < 0.35: s=2; txt='볼린저 하단~중심 구간'
    elif bb_pos > 0.95: s=-4; txt='볼린저 상단 이탈 (과열)'
    elif bb_pos > 0.80: s=-3; txt='볼린저 상단 근접'
    elif bb_pos > 0.65: s=-2; txt='볼린저 중심~상단 구간'
    else: s=0; txt='볼린저 중심 근처'
    detail['볼린저밴드'] = {'score': s, 'desc': txt}
    score += s

    # ④ 52주 저가 근접도 (역발상: 저점 = 매수 기회)
    if   w52_pos < 10: s=4; txt=f'52주 최저점 근접 ({w52_pos:.0f}%)'
    elif w52_pos < 25: s=2; txt=f'52주 저점권 ({w52_pos:.0f}%)'
    elif w52_pos > 90: s=-4; txt=f'52주 최고점 근접 ({w52_pos:.0f}%)'
    elif w52_pos > 75: s=-2; txt=f'52주 고점권 ({w52_pos:.0f}%)'
    else: s=0; txt=f'52주 중립 구간 ({w52_pos:.0f}%)'
    detail['52주 위치'] = {'score': s, 'desc': txt}
    score += s

    max_score = 15
    grade, grade_ko = _score_to_grade(score, max_score)
    atr_safe = float(atr_v) if not pd.isna(atr_v) else c*0.02
    return {
        'name': 'Algo C: 역발상형',
        'desc': 'RSI 극단값 + 볼린저밴드 하단 + 52주 저점 집중',
        'basis': '횡보장/반등 국면 특화, 과매도 구간 매수 시점 포착',
        'score': score, 'max_score': max_score,
        'grade': grade, 'grade_ko': grade_ko,
        'reliability': 'HIGH' if abs(r-50) > 25 else ('MEDIUM' if abs(r-50) > 15 else 'LOW'),
        'adx': 0, 'detail': detail,
        'short_range': {'high': round(bm), 'low': round(bl)},
        'mid_range':   {'high': round(bu), 'low': round(bl)},
        'strength': '과매도 반등 포착, 횡보장 유효',
        'weakness': '강한 추세장에서 조기 역추세 진입 위험',
    }


def algo_momentum(df):
    """Algo D: 모멘텀형 — 단기 가격 모멘텀 + 거래량 급증 + MACD 히스토그램"""
    import numpy as np
    close  = df['Close']; high = df['High']; low = df['Low']; volume = df['Volume']

    # 모멘텀 (ROC: Rate of Change)
    roc5  = (close / close.shift(5)  - 1) * 100
    roc20 = (close / close.shift(20) - 1) * 100
    roc60 = (close / close.shift(60) - 1) * 100

    ema12 = close.ewm(span=12,adjust=False).mean()
    ema26 = close.ewm(span=26,adjust=False).mean()
    macd  = ema12 - ema26
    sig   = macd.ewm(span=9,adjust=False).mean()
    hist  = macd - sig

    vol5  = volume.iloc[-5:].mean()
    vol20 = volume.iloc[-20:].mean()
    vol_ratio = vol5 / vol20 if vol20 > 0 else 1

    tr    = pd.concat([(high-low),(high-close.shift()).abs(),(low-close.shift()).abs()],axis=1).max(axis=1)
    atr_v = tr.rolling(14).mean().iloc[-1]

    c=close.iloc[-1]
    r5=roc5.iloc[-1]; r20=roc20.iloc[-1]; r60=roc60.iloc[-1]
    h=hist.iloc[-1]; hp=hist.iloc[-2] if len(hist)>1 else h; hp2=hist.iloc[-3] if len(hist)>2 else hp

    detail = {}; score = 0

    # ① 단기 ROC (5일 모멘텀)
    if   r5 > 5:   s=3; txt=f'5일 모멘텀 +{r5:.1f}% (강한 상승)'
    elif r5 > 2:   s=2; txt=f'5일 모멘텀 +{r5:.1f}%'
    elif r5 > 0:   s=1; txt=f'5일 모멘텀 +{r5:.1f}% (약한 상승)'
    elif r5 > -2:  s=-1; txt=f'5일 모멘텀 {r5:.1f}% (약한 하락)'
    elif r5 > -5:  s=-2; txt=f'5일 모멘텀 {r5:.1f}%'
    else:           s=-3; txt=f'5일 모멘텀 {r5:.1f}% (강한 하락)'
    detail['단기 모멘텀(5일)'] = {'score': s, 'desc': txt}
    score += s

    # ② 중기 ROC (20일 모멘텀)
    if   r20 > 10: s=2; txt=f'20일 모멘텀 +{r20:.1f}% (강한 추세)'
    elif r20 > 3:  s=1; txt=f'20일 모멘텀 +{r20:.1f}%'
    elif r20 < -10:s=-2; txt=f'20일 모멘텀 {r20:.1f}% (강한 하락)'
    elif r20 < -3: s=-1; txt=f'20일 모멘텀 {r20:.1f}%'
    else:           s=0; txt=f'20일 모멘텀 {r20:.1f}% (중립)'
    detail['중기 모멘텀(20일)'] = {'score': s, 'desc': txt}
    score += s

    # ③ 거래량 급증 (모멘텀의 확신도)
    if   vol_ratio > 2.0:  s=3; txt=f'거래량 {vol_ratio:.1f}x 폭증 (강한 확신)'
    elif vol_ratio > 1.5:  s=2; txt=f'거래량 {vol_ratio:.1f}x 급증'
    elif vol_ratio > 1.2:  s=1; txt=f'거래량 {vol_ratio:.1f}x 증가'
    elif vol_ratio < 0.5:  s=-2; txt=f'거래량 {vol_ratio:.1f}x 급감 (모멘텀 약화)'
    elif vol_ratio < 0.8:  s=-1; txt=f'거래량 {vol_ratio:.1f}x 감소'
    else:                   s=0; txt=f'거래량 {vol_ratio:.1f}x 보통'
    detail['거래량 강도'] = {'score': s, 'desc': txt}
    score += s

    # ④ MACD 히스토그램 가속도 (연속 방향 확인)
    h_accel = h - hp  # 1차 가속
    h_jerk  = (h-hp) - (hp-hp2)  # 2차 변화
    if   h > 0 and h_accel > 0 and h_jerk > 0: s=3; txt='히스토그램 가속 상승 (모멘텀 가속)'
    elif h > 0 and h_accel > 0:                 s=2; txt='히스토그램 상승 중'
    elif h > 0:                                 s=1; txt='히스토그램 양수 (감속)'
    elif h < 0 and h_accel < 0 and h_jerk < 0: s=-3; txt='히스토그램 가속 하락 (모멘텀 악화)'
    elif h < 0 and h_accel < 0:                 s=-2; txt='히스토그램 하락 중'
    else:                                        s=-1; txt='히스토그램 음수 (감속)'
    detail['MACD 히스토그램'] = {'score': s, 'desc': txt}
    score += s

    # ⑤ 장기 모멘텀 (60일, 추세 방향 최종 확인)
    if   r60 > 15: s=2; txt=f'60일 장기 상승 모멘텀 +{r60:.1f}%'
    elif r60 > 5:  s=1; txt=f'60일 모멘텀 +{r60:.1f}%'
    elif r60 < -15:s=-2; txt=f'60일 장기 하락 모멘텀 {r60:.1f}%'
    elif r60 < -5: s=-1; txt=f'60일 모멘텀 {r60:.1f}%'
    else:           s=0; txt=f'60일 모멘텀 {r60:.1f}% (중립)'
    detail['장기 모멘텀(60일)'] = {'score': s, 'desc': txt}
    score += s

    max_score = 13
    grade, grade_ko = _score_to_grade(score, max_score)
    atr_safe = float(atr_v) if not pd.isna(atr_v) else c*0.02
    return {
        'name': 'Algo D: 모멘텀형',
        'desc': '단기/중기/장기 ROC + 거래량 강도 + MACD 히스토그램 가속도',
        'basis': 'Jegadeesh & Titman 모멘텀 팩터 (3~12개월 지속성 학술 검증)',
        'score': score, 'max_score': max_score,
        'grade': grade, 'grade_ko': grade_ko,
        'reliability': 'HIGH' if vol_ratio > 1.5 and abs(r20) > 5 else ('MEDIUM' if abs(r20) > 2 else 'LOW'),
        'adx': 0, 'detail': detail,
        'short_range': {'high': round(c + atr_safe*2), 'low': round(c - atr_safe*1.5)},
        'mid_range':   {'high': round(c*(1+r20/100*0.5)), 'low': round(c*(1-abs(r20)/100*0.3))},
        'strength': '추세 초기 진입 강점, 거래량 동반 신호 강력',
        'weakness': '과매수 구간 추격 매수 위험, 횡보장 오신호 가능',
    }


def _score_to_grade(score, max_score):
    ratio = score / max_score
    if   ratio >= 0.55:  return 'strong_up',   '강한 상승'
    elif ratio >= 0.25:  return 'up',           '상승'
    elif ratio >= -0.25: return 'neutral',      '중립'
    elif ratio >= -0.55: return 'down',         '하락'
    else:                return 'strong_down',  '강한 하락'


def _fetch_df(h, usd_krw):
    if h['market'] == 'US':
        t  = yf.Ticker(h['ticker'])
        df = t.history(period='1y')
        if df.empty or len(df) < 30:
            raise ValueError('데이터 부족')
        df.index = df.index.tz_localize(None)
        df['Close'] = df['Close'] * usd_krw
        df['High']  = df['High']  * usd_krw
        df['Low']   = df['Low']   * usd_krw
    else:
        end   = datetime.now().strftime('%Y%m%d')
        start = (datetime.now() - timedelta(days=400)).strftime('%Y%m%d')
        raw   = krx.get_market_ohlcv(start, end, h['ticker'])
        if raw.empty or len(raw) < 30:
            raise ValueError('데이터 부족')
        df = raw.rename(columns={'시가':'Open','고가':'High','저가':'Low','종가':'Close','거래량':'Volume'})
    return df


@app.route('/api/compare')
def api_compare():
    data    = load_portfolio()
    usd_krw = get_usd_krw()
    results = []

    for h in data['holdings']:
        item = {'name': h['name'], 'ticker': h['ticker'], 'market': h['market'], 'algos': []}
        try:
            df = _fetch_df(h, usd_krw)
            # 미니차트 데이터 (60일)
            close  = df['Close']
            ma5    = close.rolling(5).mean()
            ma20   = close.rolling(20).mean()
            ma60   = close.rolling(60).mean()
            n      = min(60, len(close))
            item['chart'] = {
                'dates': [str(d.date()) for d in close.iloc[-n:].index],
                'close': [round(v,2) for v in close.iloc[-n:].tolist()],
                'ma5':   [round(v,2) if not pd.isna(v) else None for v in ma5.iloc[-n:].tolist()],
                'ma20':  [round(v,2) if not pd.isna(v) else None for v in ma20.iloc[-n:].tolist()],
                'ma60':  [round(v,2) if not pd.isna(v) else None for v in ma60.iloc[-n:].tolist()],
            }
            item['current_price'] = round(float(close.iloc[-1]))

            for fn in [algo_comprehensive, algo_trend_follow, algo_contrarian, algo_momentum]:
                try:
                    item['algos'].append(fn(df))
                except Exception as e:
                    item['algos'].append({'name': fn.__doc__.split('—')[0].strip(), 'error': str(e)})
        except Exception as e:
            item['error'] = str(e)
        results.append(item)

    return jsonify({'compare': results})


@app.route('/api/period-perf')
def api_period_perf():
    """각 종목의 기간별 수익률 반환 (1일/5일/1개월/3개월/1년)"""
    portfolio = load_portfolio()
    usd_krw = get_usd_krw()

    periods = [
        ('1일',  1),
        ('5일',  5),
        ('1개월', 21),
        ('3개월', 63),
        ('1년',  252),
    ]

    results = []
    for h in portfolio['holdings']:
        try:
            if h['market'] == 'US':
                ticker_sym = h['ticker']
                t = yf.Ticker(ticker_sym)
                hist = t.history(period='1y')
                fx = usd_krw
            else:
                ticker_sym = h['ticker']
                end = datetime.now()
                start = end - timedelta(days=400)
                df = krx.get_market_ohlcv_by_date(
                    start.strftime('%Y%m%d'), end.strftime('%Y%m%d'), ticker_sym
                )
                hist = df[['종가']].rename(columns={'종가': 'Close'})
                fx = 1

            if hist.empty or len(hist) < 2:
                continue

            current = float(hist['Close'].iloc[-1]) * fx * h['qty']
            perf = []
            for label, days in periods:
                if len(hist) > days:
                    past_price = float(hist['Close'].iloc[-days-1]) * fx * h['qty']
                    chg = current - past_price
                    pct = (chg / past_price * 100) if past_price else 0
                    perf.append({'label': label, 'chg': round(chg), 'pct': round(pct, 2)})
                else:
                    perf.append({'label': label, 'chg': None, 'pct': None})

            results.append({
                'name': h['name'],
                'ticker': h['ticker'],
                'market': h['market'],
                'qty': h['qty'],
                'current_value': round(current),
                'perf': perf,
            })
        except Exception as e:
            results.append({'name': h['name'], 'ticker': h['ticker'], 'market': h['market'],
                            'qty': h['qty'], 'current_value': 0, 'perf': [], 'error': str(e)})

    return jsonify({'items': results})


@app.route('/api/transactions', methods=['GET'])
def api_get_transactions():
    txs = load_transactions()
    return jsonify({'transactions': txs})


@app.route('/api/transactions', methods=['POST'])
def api_add_transaction():
    try:
        tx = request.get_json()
        required = ['date', 'name', 'ticker', 'market', 'type', 'qty', 'price']
        for f in required:
            if f not in tx:
                return jsonify({'error': f'필드 누락: {f}'}), 400

        tx['qty']   = float(tx['qty'])
        tx['price'] = float(tx['price'])
        if tx['qty'] <= 0:
            return jsonify({'error': '수량은 0보다 커야 합니다'}), 400
        if tx['price'] <= 0:
            return jsonify({'error': '가격은 0보다 커야 합니다'}), 400
        tx['id']    = datetime.now().strftime('%Y%m%d%H%M%S%f')
        tx['total'] = round(tx['qty'] * tx['price'])

        txs = load_transactions()
        txs.append(tx)
        save_transactions(txs)
        holdings = rebuild_portfolio_from_transactions()
        return jsonify({'success': True, 'id': tx['id'], 'holdings_count': len(holdings)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/transactions/<tx_id>', methods=['DELETE'])
def api_delete_transaction(tx_id):
    try:
        txs = load_transactions()
        txs = [t for t in txs if t['id'] != tx_id]
        save_transactions(txs)
        holdings = rebuild_portfolio_from_transactions()
        return jsonify({'success': True, 'holdings_count': len(holdings)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/transactions/summary', methods=['GET'])
def api_tx_summary():
    """종목별 실현손익 요약"""
    txs = load_transactions()
    positions = {}
    for tx in sorted(txs, key=lambda x: x['date']):
        tk = tx['ticker']
        if tk not in positions:
            positions[tk] = {'name': tx['name'], 'ticker': tk, 'market': tx['market'],
                             'qty': 0.0, 'total_cost': 0.0, 'realized_pnl': 0.0,
                             'buy_count': 0, 'sell_count': 0}
        p = positions[tk]
        if tx['type'] == 'buy':
            p['total_cost'] += tx['qty'] * tx['price']
            p['qty'] += tx['qty']
            p['buy_count'] += 1
        else:
            if p['qty'] > 0:
                avg = p['total_cost'] / p['qty']
                p['realized_pnl'] += (tx['price'] - avg) * tx['qty']
                p['total_cost'] -= avg * tx['qty']
            p['qty'] = max(0.0, p['qty'] - tx['qty'])
            p['sell_count'] += 1
    result = []
    for p in positions.values():
        avg_cost = round(p['total_cost'] / p['qty']) if p['qty'] > 0 else 0
        result.append({
            'name': p['name'], 'ticker': p['ticker'], 'market': p['market'],
            'qty': round(p['qty'], 8), 'avg_cost': avg_cost,
            'realized_pnl': round(p['realized_pnl']),
            'buy_count': p['buy_count'], 'sell_count': p['sell_count'],
        })
    return jsonify({'summary': result})


@app.route('/api/reset-portfolio', methods=['POST'])
def api_reset_portfolio():
    try:
        save_portfolio({'holdings': []})
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/upload-json', methods=['POST'])
def api_upload_json():
    try:
        data = request.get_json()
        if not data or 'holdings' not in data:
            return jsonify({'error': '잘못된 데이터입니다'}), 400
        holdings = data['holdings']
        if not holdings:
            return jsonify({'error': '종목 데이터가 없습니다'}), 400
        save_portfolio({'holdings': holdings})
        return jsonify({'success': True, 'count': len(holdings)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# 티커 매핑 테이블
TICKER_MAP = {
    '엔비디아': ('NVDA', 'US'), '테슬라': ('TSLA', 'US'),
    '마이크로소프트': ('MSFT', 'US'), '애플': ('AAPL', 'US'),
    '쿠팡': ('CPNG', 'US'), '알파벳': ('GOOGL', 'US'),
    '아마존': ('AMZN', 'US'), '메타': ('META', 'US'),
    '한화그룹': ('462330', 'KR'), '삼성전자': ('005930', 'KR'),
    'PLUS 한화그룹주': ('462330', 'KR'),
}

_ocr_reader = None

def get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        _ocr_reader = easyocr.Reader(['ko', 'en'], gpu=False)
    return _ocr_reader


def prepare_mobile_img(img_bytes):
    """모바일 스크린샷 공통 전처리 (정확도 최대화)"""
    from PIL import Image, ImageOps
    import io
    img = Image.open(io.BytesIO(img_bytes)).convert('RGB')
    w, h = img.size

    # ① 다크 배경 → 반전 (토스 다크테마) — 먼저 체크
    arr = np.array(img)
    if arr.mean() < 130:
        img = ImageOps.invert(img)

    # ② 왼쪽 아이콘 영역 제거 (종목 로고가 OCR 방해)
    crop_left = int(w * 0.17)
    img = img.crop((crop_left, 0, w, h))

    # ③ 해상도 업스케일 (EasyOCR 정확도는 큰 이미지에서 훨씬 높음)
    w2, h2 = img.size
    target_w = 1200
    if w2 < target_w:
        ratio = target_w / w2
        img = img.resize((target_w, int(h2 * ratio)), Image.LANCZOS)
    elif w2 > 1600:
        ratio = 1600 / w2
        img = img.resize((1600, int(h2 * ratio)), Image.LANCZOS)

    return img


def ocr_image(img_bytes, preprocess_fn=None):
    """이미지 바이트 → OCR 텍스트 라인 리스트 (정확도 최대화)"""
    from PIL import Image
    img = prepare_mobile_img(img_bytes)
    if preprocess_fn:
        img = preprocess_fn(img)
    reader = get_ocr_reader()
    results = reader.readtext(
        np.array(img),
        detail=0,
        paragraph=False,
        width_ths=0.7,        # 더 좁은 텍스트 블록도 인식
        contrast_ths=0.05,
        adjust_contrast=0.8,
        text_threshold=0.35,  # 낮춰서 흐린 텍스트도 인식
        low_text=0.2,
        mag_ratio=1.0,        # 이미 업스케일했으므로 1.0
        slope_ths=0.2,        # 약간 기울어진 텍스트 허용
        ycenter_ths=0.7,
        link_threshold=0.4,
    )
    return results


def parse_toss_lines(lines):
    """OCR 라인 → 종목 리스트 파싱 (모바일 토스 레이아웃 특화)"""
    import re
    holdings = []
    current_section = 'US'

    def clean_num(s):
        """OCR 오인식 보정 후 숫자 추출 (강화판)"""
        s = str(s)
        # 단위 제거
        for unit in ['원','주','%','달러','$','₩',',']:
            s = s.replace(unit, '')
        # % 뒤 잡음 제거 (61.0%0a 등)
        s = re.sub(r'(\d+\.?\d*)[a-zA-Z]+\d*$', r'\1', s)
        # 흔한 OCR 문자 오인식 보정
        for bad, good in [
            ('O','0'),('o','0'),('D','0'),('Q','0'),
            ('l','1'),('I','1'),('|','1'),('i','1'),
            ('S','5'),('s','5'),('G','6'),('B','8'),
            ('Z','2'),('z','2'),(' ',''),
        ]:
            s = s.replace(bad, good)
        s = re.sub(r'[^\d.\-+]', '', s)
        return s

    # 전체 텍스트를 하나로 합쳐서 패턴 매칭 (줄 단위 파싱보다 유연)
    full_text = '\n'.join(lines)

    # ── 섹션별 분리 ──────────────────────────────────────────
    # 해외주식 / 국내주식 섹션 분리
    kr_idx = next((i for i, l in enumerate(lines) if '국내주식' in l), len(lines))
    us_lines = lines[:kr_idx]
    kr_lines = lines[kr_idx:]

    def extract_from_lines(src_lines, section):
        result = []
        i = 0
        while i < len(src_lines):
            line = src_lines[i].strip()

            # 수량 패턴: "28.016662주" 또는 "2주"
            qty_m = re.search(r'(\d[\d.]*)\s*주\b', line)
            if qty_m:
                qty_str = qty_m.group(1)
                try:
                    qty = float(qty_str)
                except:
                    i += 1; continue

                # 종목명: 수량 줄 바로 전 줄 (또는 같은 줄에 포함된 경우)
                # OCR 순서: 종목명 → 평가금(원) → 수량(주) → 손익
                def is_non_name(s):
                    s = s.strip()
                    if not s or len(s) < 2: return True
                    skip_words = {'보기', '해외주식', '국내주식', '주문내역', '배당금', '관심', '발견', '피드', '증권', '현재가', '평가금'}
                    if s in skip_words: return True
                    return bool(re.match(r'^[\d\s,+\-원%().\[\]$~]+$', s))

                name = ''
                # 같은 줄에 종목명 포함 여부 확인
                before_qty = line[:qty_m.start()].strip()
                if len(before_qty) >= 2 and not is_non_name(before_qty):
                    name = before_qty
                else:
                    # 최대 3줄 위까지 거슬러 올라가며 종목명 찾기
                    for back in range(1, 4):
                        if i - back < 0:
                            break
                        candidate = src_lines[i - back].strip()
                        if not is_non_name(candidate):
                            name = candidate
                            break

                if not name:
                    i += 1; continue

                # 평가금 + 손익: 수량 앞 2줄 + 이후 6줄 범위에서 찾기
                current_value, profit, profit_pct = None, None, None
                search_start = max(0, i - 2)
                search_text = ' '.join(src_lines[search_start:min(i+6, len(src_lines))])

                # 평가금: 숫자,숫자원 or 숫자원
                for val_m in re.finditer(r'(\d[\d,]+)\s*원', search_text):
                    v = float(clean_num(val_m.group(1)))
                    if v >= 100:  # 최소 100원 이상
                        current_value = v
                        break

                # 손익: +숫자 (숫자%) 또는 -숫자 (숫자%)
                # % 오인식 대응: %자리에 영문/숫자 혼합 허용
                pnl_m = re.search(
                    r'([+\-])\s*(\d[\d,]*)\s*[\(\[]\s*(\d+\.?\d*)[%\w]*\s*[\)\]]',
                    search_text
                )
                if pnl_m:
                    sign = 1 if pnl_m.group(1) == '+' else -1
                    profit = sign * float(clean_num(pnl_m.group(2)))
                    profit_pct = sign * float(clean_num(pnl_m.group(3)))

                if current_value is None:
                    i += 1; continue

                # 티커 매핑
                ticker, market = None, section
                for k, (t, m) in TICKER_MAP.items():
                    if k in name or name in k:
                        ticker, market = t, m
                        break
                if not ticker:
                    ticker = re.sub(r'[^A-Z0-9]', '', name.upper())[:6] or name[:4]

                cost = round(current_value - (profit or 0))
                result.append({
                    'name': name, 'ticker': ticker, 'market': market,
                    'qty': qty, 'current_value': round(current_value),
                    'profit': round(profit) if profit else 0,
                    'profit_pct': round(profit_pct, 1) if profit_pct else 0,
                    'cost': cost,
                })
            i += 1
        return result

    holdings += extract_from_lines(us_lines, 'US')
    holdings += extract_from_lines(kr_lines, 'KR')

    # 중복 제거 (같은 ticker)
    seen = set()
    unique = []
    for h in holdings:
        if h['ticker'] not in seen:
            seen.add(h['ticker'])
            unique.append(h)

    if unique:
        return unique

    # fallback: 옛 방식 라인 파싱
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if '국내주식' in line: current_section = 'KR'
        if '해외주식' in line: current_section = 'US'

        qty_m = re.search(r'(\d[\d.]*)\s*주\b', line)
        if qty_m and i > 0:
            name_line = lines[i - 1].strip()
            if len(name_line) >= 2 and not re.match(r'^[\d\s,+\-원%().]+$', name_line):
                qty = float(qty_m.group(1))
                current_value, profit, profit_pct = None, None, None
                for j in range(i + 1, min(i + 6, len(lines))):
                    l = lines[j].strip()
                    val_m = re.search(r'(\d[\d,]+)\s*원', l)
                    if val_m and current_value is None:
                        v = float(clean_num(val_m.group(1)))
                        if v >= 100: current_value = v
                    pnl_m = re.search(r'([+\-])\s*(\d[\d,]*)\s*[\(\[](\d+\.?\d*)[%\w]*[\)\]]', l)
                    if pnl_m and profit is None:
                        sign = 1 if pnl_m.group(1)=='+' else -1
                        profit = sign * float(clean_num(pnl_m.group(2)))
                        profit_pct = sign * float(pnl_m.group(3))
                if current_value is not None:
                    name = name_line
                    ticker, market = None, current_section
                    for k, (t, m) in TICKER_MAP.items():
                        if k in name or name in k: ticker, market = t, m; break
                    if not ticker: ticker = name[:4].upper()
                    cost = round(current_value - (profit or 0))
                    holdings.append({
                        'name': name, 'ticker': ticker, 'market': market,
                        'qty': qty, 'current_value': round(current_value),
                        'profit': round(profit) if profit else 0,
                        'profit_pct': round(profit_pct, 1) if profit_pct else 0,
                        'cost': cost,
                    })
        i += 1

    return holdings


@app.route('/api/screenshot', methods=['POST'])
def api_screenshot():
    from PIL import Image, ImageEnhance, ImageFilter, ImageOps
    from collections import Counter
    import io, re

    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다'}), 400

    f = request.files['file']
    if not f.filename.lower().rsplit('.', 1)[-1] in ('png', 'jpg', 'jpeg', 'webp'):
        return jsonify({'error': 'PNG / JPG / WEBP 이미지만 지원합니다'}), 400

    img_bytes = f.read()

    # ── 모바일 특화 5가지 전처리 ─────────────────────────────
    # prepare_mobile_img 에서 이미 다크→라이트 반전됨
    # 여기서는 추가 화질 개선에 집중

    def pp_clean(img):
        """① 기본 정리: 대비 + 선명화"""
        img = ImageEnhance.Contrast(img).enhance(2.0)
        img = ImageEnhance.Sharpness(img).enhance(2.0)
        img = ImageEnhance.Brightness(img).enhance(1.1)
        return img

    def pp_high_contrast(img):
        """② 고대비 흑백: 텍스트/배경 분리 극대화"""
        img = ImageEnhance.Contrast(img).enhance(3.5)
        img = ImageEnhance.Color(img).enhance(0.0)
        return img

    def pp_unsharp(img):
        """③ 언샤프 마스크: 경계선 강화"""
        img = ImageEnhance.Contrast(img).enhance(2.5)
        img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=250, threshold=1))
        return img

    def pp_binarize(img):
        """④ 적응형 이진화: 음영 영역별 개별 처리"""
        from PIL import Image as PILImage
        gray = np.array(img.convert('L'))
        # 블록별 임계값 (Otsu-like)
        from PIL import Image as PILImage
        import io as _io
        binary = np.where(gray > gray.mean() * 0.85, 255, 0).astype(np.uint8)
        return PILImage.fromarray(binary).convert('RGB')

    def pp_dilate(img):
        """⑤ 텍스트 굵게: 얇은 획 보완"""
        img = ImageEnhance.Contrast(img).enhance(2.5)
        img = img.filter(ImageFilter.MaxFilter(size=3))
        return img

    def pp_sharpen_gray(img):
        """⑥ 그레이스케일 + 극한 선명화"""
        img = img.convert('L').convert('RGB')
        img = ImageEnhance.Contrast(img).enhance(3.0)
        img = img.filter(ImageFilter.SHARPEN)
        img = img.filter(ImageFilter.SHARPEN)
        return img

    preprocessors = [pp_clean, pp_high_contrast, pp_unsharp, pp_binarize, pp_dilate, pp_sharpen_gray]

    all_batches = []
    debug_passes = []   # 각 전처리별 raw OCR 결과 저장

    pp_names = ['기본정리', '고대비', '언샤프', '이진화', '텍스트굵게', '그레이극한']
    for idx, pp in enumerate(preprocessors):
        try:
            lines = ocr_image(img_bytes, pp)
            parsed = parse_toss_lines(lines)
            debug_passes.append({
                'name': pp_names[idx],
                'lines': lines,
                'parsed_count': len(parsed),
            })
            if parsed:
                all_batches.append(parsed)
        except Exception as e:
            debug_passes.append({'name': pp_names[idx], 'lines': [], 'error': str(e), 'parsed_count': 0})

    if not all_batches:
        all_lines = [l for p in debug_passes for l in p.get('lines', [])]
        return jsonify({
            'error': '종목을 인식하지 못했습니다.',
            'debug': debug_passes,
            'all_lines': all_lines,
        }), 400

    # 교차 검증: 종목명 기준 집계
    bucket = {}
    for batch in all_batches:
        for h in batch:
            key = h.get('ticker') or h.get('name', '')
            bucket.setdefault(key, []).append(h)

    def majority(entries, field, cast=float):
        vals = []
        for e in entries:
            v = e.get(field)
            if v is not None:
                try: vals.append(round(cast(v), 4) if cast == float else cast(v))
                except: pass
        if not vals: return None
        return Counter(vals).most_common(1)[0][0]

    total_checks = len(all_batches)
    consensus = []
    for key, entries in bucket.items():
        name          = majority(entries, 'name',          str)
        ticker        = majority(entries, 'ticker',        str) or key
        market        = majority(entries, 'market',        str) or 'US'
        qty           = majority(entries, 'qty',           float)
        current_value = majority(entries, 'current_value', float)
        profit        = majority(entries, 'profit',        float)
        profit_pct    = majority(entries, 'profit_pct',    float)
        agreement     = len(entries)
        confidence    = round(agreement / total_checks * 100)

        if not name or qty is None or current_value is None:
            continue

        cost = round(current_value - (profit or 0))
        consensus.append({
            'name': name, 'ticker': ticker, 'market': market.upper()[:2],
            'qty': qty,
            'current_value': round(current_value),
            'profit':     round(profit) if profit is not None else 0,
            'profit_pct': round(profit_pct, 1) if profit_pct is not None else 0,
            'cost':       cost,
            'confidence': confidence,
            'agreement':  agreement,
            'total_checks': total_checks,
        })

    consensus.sort(key=lambda x: -x['current_value'])
    return jsonify({
        'success': True,
        'consensus': consensus,
        'checks_done': total_checks,
        'debug': debug_passes,   # 전처리별 OCR 원문
    })


# ════════════════════════════════════════════════════════════
# 목표가 & 메모
# ════════════════════════════════════════════════════════════

@app.route('/api/prefs', methods=['GET'])
def api_get_prefs():
    return jsonify({
        'target_prices': load_target_prices(),
        'notes':         load_notes(),
    })

@app.route('/api/prefs/target', methods=['POST'])
def api_set_target():
    d = request.get_json() or {}
    ticker = d.get('ticker','').strip()
    price  = d.get('price')
    if not ticker:
        return jsonify({'error':'티커 없음'}), 400
    tp = load_target_prices()
    if price is None:
        tp.pop(ticker, None)
    else:
        tp[ticker] = float(price)
    save_target_prices(tp)
    return jsonify({'ok': True, 'target_prices': tp})

@app.route('/api/prefs/note', methods=['POST'])
def api_set_note():
    d = request.get_json() or {}
    ticker = d.get('ticker','').strip()
    note   = d.get('note','').strip()
    if not ticker:
        return jsonify({'error':'티커 없음'}), 400
    notes = load_notes()
    if note:
        notes[ticker] = note
    else:
        notes.pop(ticker, None)
    save_notes(notes)
    return jsonify({'ok': True})


# ════════════════════════════════════════════════════════════
# 벤치마크 비교
# ════════════════════════════════════════════════════════════

@app.route('/api/benchmark')
def api_benchmark():
    """내 포트폴리오 수익률 vs 시장 지수 비교"""
    portfolio = load_portfolio()
    usd_krw   = get_usd_krw()

    periods = [('1일',1),('1주',5),('1달',21),('3달',63),('1년',252)]
    indices  = {
        'S&P500':  '^GSPC',
        '나스닥':  '^IXIC',
        '코스피':  '^KS11',
    }

    # 지수 수익률
    bench = {}
    for name, sym in indices.items():
        bench[name] = {}
        try:
            df = yf.Ticker(sym).history(period='1y')
            if df.empty: continue
            closes = df['Close']
            for label, days in periods:
                if len(closes) > days:
                    bench[name][label] = round((closes.iloc[-1]/closes.iloc[-days-1]-1)*100, 2)
        except:
            pass

    # 포트폴리오 수익률 (현재가 vs N일 전 가격)
    port_ret = {label: [] for label, _ in periods}
    for h in portfolio['holdings']:
        try:
            if h['market'] == 'US':
                df = yf.Ticker(h['ticker']).history(period='1y')
                if df.empty: continue
                closes = df['Close'] * usd_krw
            else:
                end   = datetime.now().strftime('%Y%m%d')
                start = (datetime.now()-timedelta(days=400)).strftime('%Y%m%d')
                raw   = krx.get_market_ohlcv(start, end, h['ticker'])
                if raw.empty: continue
                closes = raw['종가']
            cur_val = closes.iloc[-1] * h['qty']
            for label, days in periods:
                if len(closes) > days:
                    past_val = closes.iloc[-days-1] * h['qty']
                    port_ret[label].append((cur_val - past_val) / past_val * 100 if past_val else 0)
        except:
            pass

    port = {label: round(sum(v)/len(v), 2) if v else None for label, v in port_ret.items()}
    return jsonify({'portfolio': port, 'benchmarks': bench, 'periods': [l for l,_ in periods]})


# ════════════════════════════════════════════════════════════
# 종목 비교
# ════════════════════════════════════════════════════════════

@app.route('/api/compare-stocks')
def api_compare_stocks():
    a = request.args.get('a','').upper().strip()
    b = request.args.get('b','').upper().strip()
    if not a or not b:
        return jsonify({'error':'두 티커를 모두 입력하세요'}), 400

    usd_krw = get_usd_krw()
    results = {}

    for sym in [a, b]:
        try:
            # KR 종목은 숫자 6자리
            if sym.isdigit():
                end   = datetime.now().strftime('%Y%m%d')
                start = (datetime.now()-timedelta(days=400)).strftime('%Y%m%d')
                raw   = krx.get_market_ohlcv(start, end, sym)
                if raw.empty: raise ValueError('데이터 없음')
                df = raw.rename(columns={'시가':'Open','고가':'High','저가':'Low','종가':'Close','거래량':'Volume'})
                name   = krx.get_market_ticker_name(sym) or sym
                market = 'KR'
            else:
                t  = yf.Ticker(sym)
                df = t.history(period='1y')
                if df.empty: raise ValueError('데이터 없음')
                df.index = df.index.tz_localize(None)
                df['Close'] *= usd_krw; df['High'] *= usd_krw; df['Low'] *= usd_krw
                info   = t.info
                name   = info.get('shortName', sym)
                market = 'US'

            ind = calc_indicators(df)
            n   = min(60, len(df))
            closes = df['Close']
            results[sym] = {
                'name':    name,
                'market':  market,
                'price':   round(float(closes.iloc[-1])),
                'change1d': round((closes.iloc[-1]/closes.iloc[-2]-1)*100, 2) if len(closes)>1 else 0,
                'change1m': round((closes.iloc[-1]/closes.iloc[-22]-1)*100, 2) if len(closes)>22 else 0,
                'change3m': round((closes.iloc[-1]/closes.iloc[-63]-1)*100, 2) if len(closes)>63 else 0,
                'grade':       ind['grade'],
                'grade_ko':    ind['grade_ko'],
                'total_score': ind['total_score'],
                'rsi':         ind['rsi'],
                'adx':         ind['adx'],
                'reliability': ind['reliability'],
                'indicators':  ind['indicators'],
                'w52':         ind['w52'],
                'short_range': ind['short_range'],
                'mid_range':   ind['mid_range'],
                'chart': {
                    'dates': ind['chart']['dates'],
                    'close': ind['chart']['close'],
                    'ma5':   ind['chart']['ma5'],
                    'ma20':  ind['chart']['ma20'],
                },
            }
        except Exception as e:
            results[sym] = {'error': str(e), 'name': sym}

    return jsonify({'a': results.get(a,{}), 'b': results.get(b,{}), 'tickers': [a,b]})


# ════════════════════════════════════════════════════════════
# 리밸런싱 제안
# ════════════════════════════════════════════════════════════

@app.route('/api/rebalance-targets', methods=['GET'])
def api_get_rebalance_targets():
    return jsonify({'targets': load_rebalance_targets()})

@app.route('/api/rebalance-targets', methods=['POST'])
def api_set_rebalance_targets():
    d = request.get_json() or {}
    targets = d.get('targets', {})
    save_rebalance_targets(targets)
    return jsonify({'ok': True})

@app.route('/api/rebalance')
def api_rebalance():
    """현재 포트폴리오 vs 목표 비중 비교 → 매수/매도 제안"""
    portfolio = load_portfolio()
    targets   = load_rebalance_targets()  # {ticker: target_pct}
    usd_krw   = get_usd_krw()

    if not portfolio['holdings']:
        return jsonify({'error': '포트폴리오가 비어있습니다'}), 400

    # 현재 평가금액 계산
    values = {}
    for h in portfolio['holdings']:
        try:
            if h['market'] == 'US':
                price = yf.Ticker(h['ticker']).fast_info.get('lastPrice', 0) * usd_krw
            else:
                today = datetime.now().strftime('%Y%m%d')
                yest  = (datetime.now()-timedelta(days=7)).strftime('%Y%m%d')
                df    = krx.get_market_ohlcv(yest, today, h['ticker'])
                price = int(df.iloc[-1]['종가']) if not df.empty else 0
            values[h['ticker']] = {
                'name':    h['name'],
                'market':  h['market'],
                'qty':     h['qty'],
                'price':   price,
                'value':   round(price * h['qty']),
            }
        except:
            values[h['ticker']] = {
                'name': h['name'], 'market': h['market'],
                'qty': h['qty'], 'price': 0, 'value': 0
            }

    total = sum(v['value'] for v in values.values()) or 1

    # 현재 비중
    for tk, v in values.items():
        v['current_pct'] = round(v['value'] / total * 100, 1)
        v['target_pct']  = float(targets.get(tk, 0))
        diff_pct = v['target_pct'] - v['current_pct']
        diff_val = round(total * diff_pct / 100)
        v['diff_pct'] = round(diff_pct, 1)
        v['action']   = '매수' if diff_val > 0 else ('매도' if diff_val < 0 else '유지')
        v['diff_val'] = abs(diff_val)
        v['diff_qty'] = round(abs(diff_val) / v['price'], 6) if v['price'] > 0 else 0

    items = sorted(values.values(), key=lambda x: abs(x['diff_pct']), reverse=True)
    target_total = sum(float(targets.get(tk, 0)) for tk in values)

    return jsonify({
        'items':        items,
        'total_value':  total,
        'target_total': round(target_total, 1),
        'usd_krw':      usd_krw,
    })


if __name__ == '__main__':
    # 백그라운드에서 KR 종목명 캐시 미리 빌드
    import threading
    def _prebuild_cache():
        try:
            today = datetime.now().strftime('%Y%m%d')
            tickers = krx.get_market_ticker_list(today, market='ALL')
            for tk in tickers:
                try:
                    _kr_name_cache[tk] = krx.get_market_ticker_name(tk)
                except:
                    pass
            print(f"[Cache] KR 종목명 {len(_kr_name_cache)}개 로드 완료")
        except Exception as e:
            print(f"[Cache] KR 종목명 캐시 빌드 실패: {e}")
    threading.Thread(target=_prebuild_cache, daemon=True).start()

    print("Dashboard: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
