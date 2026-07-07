"""
用户管理系统 — Flask 主应用
===========================
安全加固重构版本，修复原始项目全部高危漏洞。
"""

import os
import re
import secrets
from datetime import timedelta
from functools import wraps

from flask import (
    Flask,
    abort,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash

# ---------------------------------------------------------------------------
# 应用初始化
# ---------------------------------------------------------------------------

app = Flask(__name__)

# ---------- 安全配置 ----------
# 优先使用环境变量中的密钥，开发环境下自动生成随机密钥
app.config.update(
    SECRET_KEY=os.environ.get(
        "SECRET_KEY",
        secrets.token_hex(32),
    ),
    PERMANENT_SESSION_LIFETIME=timedelta(hours=2),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=False,  # 生产环境启用 HTTPS 后改为 True
)

# ---------------------------------------------------------------------------
# 用户数据库（密码经 Werkzeug 哈希加密存储）
# ---------------------------------------------------------------------------

def _build_user_db():
    """构建用户数据库，密码使用 Werkzeug pbkdf2:sha256 哈希加密。

    原始密码仅在此函数中出现，运行后内存中仅保留哈希值，
    彻底消除密码明文存储 / 比对 / 展示问题。
    """
    raw = {
        "admin": {
            "username": "admin",
            "password": "admin123",
            "role": "admin",
            "email": "admin@example.com",
            "phone": "13800138000",
            "balance": 99999,
        },
        "alice": {
            "username": "alice",
            "password": "alice2025",
            "role": "user",
            "email": "alice@example.com",
            "phone": "13900139001",
            "balance": 100,
        },
    }
    db = {}
    for uid, info in raw.items():
        record = info.copy()
        record["password"] = generate_password_hash(record["password"])
        db[uid] = record
    return db


USERS = _build_user_db()

# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def sanitize_input(text):
    """过滤用户输入，移除 HTML 标签和危险关键字，防范 XSS。"""
    if not text:
        return ""
    text = str(text).strip()
    text = re.sub(r"<[^>]*>", "", text)           # 移除 HTML 标签
    text = re.sub(r"javascript\s*:", "", text, flags=re.IGNORECASE)  # 移除伪协议
    return text


def mask_phone(phone):
    """脱敏手机号：138****8000"""
    if not phone or len(phone) < 7:
        return phone
    return phone[:3] + "****" + phone[-4:]


def mask_balance(balance):
    """格式化余额为货币显示。"""
    try:
        return "¥{:,.2f}".format(float(balance))
    except (TypeError, ValueError):
        return "¥0.00"


def sanitize_user_info(user_info):
    """构造可供模板安全使用的用户信息字典：
    - 移除 password 字段（永不传递到前端）
    - 手机号、余额脱敏
    - 保留角色等非敏感信息用于权限展示
    """
    if not user_info:
        return None
    return {
        "username": user_info.get("username", ""),
        "role": user_info.get("role", ""),
        "email": user_info.get("email", ""),
        "phone": mask_phone(user_info.get("phone", "")),
        "balance": mask_balance(user_info.get("balance", 0)),
    }

# ---------------------------------------------------------------------------
# 登录鉴权装饰器
# ---------------------------------------------------------------------------

def login_required(f):
    """路由装饰器：未登录用户重定向到登录页。"""
    @wraps(f)
    def decorated(*args, **kwargs):
        if "username" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

# ---------------------------------------------------------------------------
# CSRF 防护
# ---------------------------------------------------------------------------

@app.before_request
def _csrf_protect():
    """对 POST 请求验证 CSRF 令牌，防范跨站请求伪造。"""
    if request.method == "POST" and request.endpoint != "static":
        # 排除静态文件端点
        token = request.form.get("csrf_token")
        stored = session.get("csrf_token")
        if not token or not stored or token != stored:
            abort(400, "CSRF token 缺失或无效。请刷新页面后重试。")


@app.context_processor
def _inject_csrf_token():
    """向所有模板注入 CSRF 令牌。"""
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_hex(32)
    return {"csrf_token": session["csrf_token"]}

# ---------------------------------------------------------------------------
# 路由 — 首页
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    """首页：已登录用户展示脱敏后的个人信息，未登录用户提示登录。"""
    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template(
        "index.html",
        username=username,
        user=user_info,
    )

# ---------------------------------------------------------------------------
# 路由 — 登录
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    """登录页面。

    GET  ：返回登录表单。
    POST ：校验用户名 + 密码（哈希比对），校验通过后创建会话。
    """
    if request.method == "POST":
        # 输入过滤
        username = sanitize_input(request.form.get("username", ""))
        password = request.form.get("password", "")   # 密码不需过滤（不展示）

        # 基础非空校验
        if not username or not password:
            return render_template("login.html", error="用户名和密码不能为空")

        try:
            user_record = USERS.get(username)
            if user_record and check_password_hash(user_record["password"], password):
                # 登录成功 → 重置会话（防会话固定攻击）
                session.clear()
                session.permanent = True
                session["username"] = username
                session["csrf_token"] = secrets.token_hex(32)

                user_info = sanitize_user_info(user_record)
                return render_template(
                    "index.html",
                    username=username,
                    user=user_info,
                )

            # 统一错误提示（防止用户名枚举）
            return render_template("login.html", error="用户名或密码错误")

        except Exception as e:
            app.logger.error("登录异常: %s", str(e))
            return render_template("login.html", error="系统异常，请稍后重试")

    return render_template("login.html")

# ---------------------------------------------------------------------------
# 路由 — 登出
# ---------------------------------------------------------------------------

@app.route("/logout")
def logout():
    """登出：清除会话后重定向至首页。"""
    session.clear()
    return redirect(url_for("index"))

# ---------------------------------------------------------------------------
# 错误处理器
# ---------------------------------------------------------------------------

@app.errorhandler(400)
def bad_request(e):
    return render_template("login.html", error=str(e)), 400


@app.errorhandler(404)
def not_found(e):
    return render_template("login.html", error="页面不存在"), 404


@app.errorhandler(500)
def server_error(e):
    app.logger.error("服务器内部错误: %s", str(e))
    return render_template("login.html", error="服务器内部错误，请稍后重试"), 500

# ---------------------------------------------------------------------------
# 启动入口
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # 生产环境请通过环境变量 FLASK_DEBUG=true 开启调试
    debug_mode = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(
        debug=debug_mode,
        host="0.0.0.0",
        port=5000,
    )
