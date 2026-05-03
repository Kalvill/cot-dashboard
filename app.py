#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
COT Dashboard — Flask сервер
Логін / реєстрація / збереження bias
"""

from flask import Flask, request, jsonify, session, send_from_directory, redirect, url_for
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3, os, json
from pathlib import Path
from datetime import timedelta
from functools import wraps

# ================================================================
# ⚙️  КОНФІГУРАЦІЯ
# ================================================================
BASE_DIR    = Path(__file__).parent
DB_PATH     = BASE_DIR / "users.db"
OUTPUT_DIR  = BASE_DIR / "output"

app = Flask(__name__, static_folder=str(OUTPUT_DIR))

# Секретний ключ для сесій — зміни на щось своє!
app.secret_key = os.environ.get("SECRET_KEY", "local_dev_key")
app.permanent_session_lifetime = timedelta(days=30)

# ================================================================
# 🗄️  БАЗА ДАНИХ
# ================================================================
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Створює таблиці якщо не існують"""
    with get_db() as db:
        db.executescript("""
            CREATE TABLE IF NOT EXISTS users (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                email    TEXT    UNIQUE NOT NULL,
                password TEXT    NOT NULL,
                role     TEXT    DEFAULT 'user',
                created  DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS bias (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id    INTEGER NOT NULL,
                instrument TEXT    NOT NULL,
                value      TEXT    NOT NULL,
                note       TEXT    DEFAULT '',
                updated    DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(user_id, instrument),
                FOREIGN KEY(user_id) REFERENCES users(id)
            );
        """)
    print(f"✅  База даних: {DB_PATH}")

# ================================================================
# 🔐  ДЕКОРАТОР — перевірка логіну
# ================================================================
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'error': 'Unauthorized'}), 401
        return f(*args, **kwargs)
    return decorated

# ================================================================
# 📄  ГОЛОВНА СТОРІНКА — роздає дашборд
# ================================================================
@app.route('/')
def index():
    """Якщо не залогінений — редірект на логін"""
    if 'user_id' not in session:
        return redirect(url_for('login_page'))
    return send_from_directory(str(OUTPUT_DIR), 'cot_dashboard.html')

@app.route('/login')
def login_page():
    """Сторінка логіну"""
    return LOGIN_HTML

# ================================================================
# 🔑  API — АВТОРИЗАЦІЯ
# ================================================================
@app.route('/api/register', methods=['POST'])
def register():
    data = request.get_json()
    email    = (data.get('email') or '').strip().lower()
    password = (data.get('password') or '').strip()

    if not email or not password:
        return jsonify({'error': 'Email та пароль обовʼязкові'}), 400
    if len(password) < 6:
        return jsonify({'error': 'Пароль мінімум 6 символів'}), 400

    hashed = generate_password_hash(password)
    try:
        with get_db() as db:
            db.execute("INSERT INTO users (email, password) VALUES (?,?)", (email, hashed))
        return jsonify({'ok': True})
    except sqlite3.IntegrityError:
        return jsonify({'error': 'Email вже зареєстровано'}), 409

@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    email    = (data.get('email') or '').strip().lower()
    password = (data.get('password') or '').strip()

    with get_db() as db:
        user = db.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()

    if not user or not check_password_hash(user['password'], password):
        return jsonify({'error': 'Невірний email або пароль'}), 401

    session.permanent = True
    session['user_id'] = user['id']
    session['email']   = user['email']
    return jsonify({'ok': True, 'email': user['email']})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'ok': True})

@app.route('/api/me')
def me():
    """Перевірка поточного користувача"""
    if 'user_id' not in session:
        return jsonify({'logged_in': False})
    return jsonify({'logged_in': True, 'email': session.get('email')})

# ================================================================
# 📊  API — BIAS
# ================================================================
@app.route('/api/bias', methods=['GET'])
@login_required
def get_all_bias():
    """Повертає всі bias поточного користувача"""
    with get_db() as db:
        rows = db.execute(
            "SELECT instrument, value, note FROM bias WHERE user_id=?",
            (session['user_id'],)
        ).fetchall()
    result = {r['instrument']: {'value': r['value'], 'note': r['note']} for r in rows}
    return jsonify(result)

@app.route('/api/bias/<instrument>', methods=['POST'])
@login_required
def set_bias(instrument):
    """Зберігає bias для інструменту"""
    data  = request.get_json()
    value = data.get('value', 'neutral')   # long / short / neutral
    note  = data.get('note', '')[:200]     # нотатка до 200 символів

    if value not in ('long', 'short', 'neutral'):
        return jsonify({'error': 'Невірне значення'}), 400

    with get_db() as db:
        db.execute("""
            INSERT INTO bias (user_id, instrument, value, note, updated)
            VALUES (?,?,?,?, CURRENT_TIMESTAMP)
            ON CONFLICT(user_id, instrument)
            DO UPDATE SET value=excluded.value, note=excluded.note,
                          updated=CURRENT_TIMESTAMP
        """, (session['user_id'], instrument, value, note))
    return jsonify({'ok': True})

@app.route('/api/bias/<instrument>', methods=['DELETE'])
@login_required
def delete_bias(instrument):
    """Скидає bias для інструменту"""
    with get_db() as db:
        db.execute("DELETE FROM bias WHERE user_id=? AND instrument=?",
                   (session['user_id'], instrument))
    return jsonify({'ok': True})

# ================================================================
# 🖼️  HTML СТОРІНКА ЛОГІНУ
# ================================================================
LOGIN_HTML = """<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>COT Dashboard — Вхід</title>
<style>
:root{--bg:#0b0d12;--bg2:#1a1e2d;--bd:#343d5a;--g:#20d483;--r:#f0515a;
  --t:#dde2ee;--d:#8090b0;--f:'Courier New',monospace;}
*{margin:0;padding:0;box-sizing:border-box;}
body{background:var(--bg);color:var(--t);font-family:var(--f);
  min-height:100vh;display:flex;align-items:center;justify-content:center;}
.box{background:var(--bg2);border:1px solid var(--bd);border-radius:8px;
  padding:40px;width:360px;}
.logo{text-align:center;margin-bottom:32px;}
.logo-t{font-size:22px;font-weight:bold;letter-spacing:3px;}
.logo-t em{color:var(--g);font-style:normal;}
.logo-s{font-size:9px;color:var(--d);letter-spacing:2px;margin-top:4px;}
.tabs{display:flex;margin-bottom:24px;border-bottom:1px solid var(--bd);}
.tab{flex:1;padding:8px;text-align:center;cursor:pointer;font-size:11px;
  color:var(--d);letter-spacing:1px;border-bottom:2px solid transparent;
  margin-bottom:-1px;transition:all .15s;}
.tab.active{color:#fff;border-bottom-color:var(--g);}
.field{margin-bottom:14px;}
.field label{display:block;font-size:9px;color:var(--d);
  letter-spacing:.8px;margin-bottom:5px;}
.field input{width:100%;background:var(--bg);border:1px solid var(--bd);
  border-radius:4px;padding:10px 12px;color:var(--t);font-family:var(--f);
  font-size:12px;outline:none;transition:border .15s;}
.field input:focus{border-color:var(--g);}
.btn{width:100%;padding:11px;background:var(--g);color:#000;
  border:none;border-radius:4px;cursor:pointer;font-family:var(--f);
  font-size:12px;font-weight:bold;letter-spacing:1px;margin-top:6px;
  transition:opacity .15s;}
.btn:hover{opacity:.85;}
.msg{font-size:11px;padding:8px 12px;border-radius:4px;margin-top:12px;
  text-align:center;display:none;}
.msg.err{background:rgba(240,81,90,.15);border:1px solid var(--r);color:var(--r);}
.msg.ok{background:rgba(32,212,131,.15);border:1px solid var(--g);color:var(--g);}
</style>
</head>
<body>
<div class="box">
  <div class="logo">
    <div class="logo-t">COT <em>DASHBOARD</em></div>
    <div class="logo-s">COMMITMENTS OF TRADERS</div>
  </div>
  <div class="tabs">
    <div class="tab active" onclick="showTab('login')">ВХІД</div>
    <div class="tab" onclick="showTab('reg')">РЕЄСТРАЦІЯ</div>
  </div>

  <!-- Форма входу -->
  <div id="t-login">
    <div class="field">
      <label>EMAIL</label>
      <input type="email" id="l-email" placeholder="your@email.com">
    </div>
    <div class="field">
      <label>ПАРОЛЬ</label>
      <input type="password" id="l-pass" placeholder="••••••••"
             onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <button class="btn" onclick="doLogin()">УВІЙТИ</button>
    <div class="msg" id="l-msg"></div>
  </div>

  <!-- Форма реєстрації -->
  <div id="t-reg" style="display:none">
    <div class="field">
      <label>EMAIL</label>
      <input type="email" id="r-email" placeholder="your@email.com">
    </div>
    <div class="field">
      <label>ПАРОЛЬ</label>
      <input type="password" id="r-pass" placeholder="мінімум 6 символів">
    </div>
    <div class="field">
      <label>ПАРОЛЬ ЩЕ РАЗ</label>
      <input type="password" id="r-pass2" placeholder="••••••••"
             onkeydown="if(event.key==='Enter')doRegister()">
    </div>
    <button class="btn" onclick="doRegister()">ЗАРЕЄСТРУВАТИСЬ</button>
    <div class="msg" id="r-msg"></div>
  </div>
</div>

<script>
function showTab(t){
  document.querySelectorAll('.tab').forEach((el,i)=>
    el.classList.toggle('active', (i===0&&t==='login')||(i===1&&t==='reg')));
  document.getElementById('t-login').style.display = t==='login'?'':'none';
  document.getElementById('t-reg').style.display   = t==='reg' ?'':'none';
}

function showMsg(id, text, isErr){
  const el = document.getElementById(id);
  el.textContent = text;
  el.className = 'msg ' + (isErr ? 'err' : 'ok');
  el.style.display = 'block';
}

async function doLogin(){
  const email = document.getElementById('l-email').value.trim();
  const pass  = document.getElementById('l-pass').value;
  const res   = await fetch('/api/login',{
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({email, password: pass})
  });
  const data = await res.json();
  if(data.ok){ window.location.href = '/'; }
  else { showMsg('l-msg', data.error || 'Помилка', true); }
}

async function doRegister(){
  const email = document.getElementById('r-email').value.trim();
  const pass  = document.getElementById('r-pass').value;
  const pass2 = document.getElementById('r-pass2').value;
  if(pass !== pass2){ showMsg('r-msg','Паролі не збігаються',true); return; }
  const res  = await fetch('/api/register',{
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({email, password: pass})
  });
  const data = await res.json();
  if(data.ok){
    showMsg('r-msg','Успішно! Тепер увійдіть.',false);
    setTimeout(()=>showTab('login'),1500);
  } else {
    showMsg('r-msg', data.error || 'Помилка', true);
  }
}
</script>
</body>
</html>
"""

# ================================================================
# 🚀  ЗАПУСК
# ================================================================
if __name__ == '__main__':
    init_db()
    print("\n" + "="*50)
    print("   COT Dashboard Flask Server")
    print("="*50)
    print(f"   URL:  http://localhost:5000")
    print(f"   DB:   {DB_PATH}")
    print("="*50 + "\n")
    app.run(debug=True, host='0.0.0.0', port=5000)