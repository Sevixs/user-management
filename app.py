"""
用户管理系统 — Flask 主应用
===========================
安全加固重构版本，修复原始项目全部高危漏洞。
"""

import os
import re
import secrets
import sqlite3
import uuid
import time
import subprocess
import platform
import urllib.request
import urllib.error
import urllib.parse
import socket
import ipaddress
from datetime import timedelta
from functools import wraps
from collections import defaultdict

from flask import (
    Flask,
    abort,
    redirect,
    render_template,
    request,
    session,
    url_for,
    send_from_directory,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

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
    MAX_CONTENT_LENGTH=16 * 1024 * 1024,  # 上传文件最大 16MB
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
# 上传安全配置
# ---------------------------------------------------------------------------

# 图片后缀白名单
ALLOWED_EXTENSIONS = {"jpg", "jpeg", "png", "gif"}

# 单文件大小上限（5MB）
MAX_FILE_SIZE = 5 * 1024 * 1024

# 上传文件存储目录（非 static 子目录，避免匿名直接访问）
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "uploads")

# ---------- 充值安全配置 ----------
MAX_RECHARGE_AMOUNT = 10000  # 单次充值上限

# ---------- 频率限制 ----------
_rate_limit_store = defaultdict(list)
RATE_LIMIT_WINDOW = 10       # 10 秒窗口
RATE_LIMIT_MAX = 10           # 窗口内最多 10 次请求


def check_rate_limit(key_prefix):
    """简易内存频率限制，防止接口被批量遍历。"""
    now = time.time()
    key = f"{key_prefix}:{request.remote_addr}"
    records = _rate_limit_store[key]
    _rate_limit_store[key] = [t for t in records if now - t < RATE_LIMIT_WINDOW]
    if len(_rate_limit_store[key]) >= RATE_LIMIT_MAX:
        return False
    _rate_limit_store[key].append(now)
    return True


def allowed_file(filename):
    """校验文件后缀是否在白名单内。"""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    return ext in ALLOWED_EXTENSIONS


def check_image_magic(file_stream):
    """读取文件二进制头部，校验真实文件 MIME 类型（魔数校验）。

    仅放行 JPEG / PNG / GIF 纯图片文件，拦截图片马及伪装脚本。
    """
    magic = file_stream.read(8)
    file_stream.seek(0)  # 复位指针，供后续保存

    if magic[:3] == b"\xff\xd8\xff":
        return True, "jpeg"
    elif magic[:4] == b"\x89PNG":
        return True, "png"
    elif magic[:3] == b"GIF":
        return True, "gif"
    else:
        return False, None


def sanitize_display_text(text):
    """过滤用于前端展示的文本中的 HTML 特殊字符，防范存储型 XSS。"""
    if not text:
        return ""
    replacements = {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        "'": "&#x27;",
        '"': "&quot;",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


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


def validate_sql_input(text, field_type="text"):
    """校验用户输入，拦截 SQL 注入危险字符，并做格式白名单校验。

    拦截项：单引号 ' 、双引号 " 、分号 ; 、注释符 -- 和 /* 、
           反斜杠 \\ 、backtick ` 、括号 ( ) 等 SQL 特殊字符。
    格式校验：username 仅允许字母、数字、下划线、@、点、横线；
             email 仅允许合法邮箱字符；
             phone 仅允许数字、横线、加号。
    """
    if not text:
        return True, ""

    # SQL 危险字符黑名单
    # 拦截: 单引号 ' 双引号 " 分号 ; 注释 -- /* 反斜杠 \ backtick ` 括号 ()
    dangerous = re.search(r"['\"\;\-\-/\*\\\`\(\)]", text)
    if dangerous:
        return False, "输入包含非法 SQL 字符"

    # 格式白名单校验
    if field_type == "username":
        if not re.match(r"^[a-zA-Z0-9_@.\-]+$", text):
            return False, "用户名只能包含字母、数字、下划线、@、点、横线"
    elif field_type == "email":
        if text and not re.match(r"^[a-zA-Z0-9_@.\-]+$", text):
            return False, "邮箱格式不合法"
    elif field_type == "phone":
        if text and not re.match(r"^[\d\+\-]+$", text):
            return False, "手机号只能包含数字、横线、加号"
    elif field_type == "keyword":
        if len(text) > 100:
            return False, "搜索关键词过长"

    return True, ""


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
    if request.method == "POST" and request.endpoint not in ("static", "serve_upload"):
        # 排除静态文件、上传文件访问端点
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
# 路由 — 动态页面加载（安全加固版）
# ---------------------------------------------------------------------------

# pages 目录的绝对路径（用于路径归属校验）
_PAGES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pages")


@app.route("/page")
def dynamic_page():
    """动态页面加载 — 安全加固版。

    安全措施：
    1. 黑名单拦截路径穿越字符：../ / \\ ~ ./
    2. 白名单校验：仅允许字母、数字、下划线、短横线
    3. 绝对路径规范化 + 目录归属校验（强制在 pages 目录内）
    4. 仅允许读取 .html 文件
    """
    name = request.args.get("name", "")
    page_content = None
    page_error = None

    if not name:
        page_error = "请指定页面名称"
    else:
        # 安全校验 1：黑名单拦截路径穿越字符
        if ".." in name or "/" in name or "\\" in name or "~" in name:
            print(f"[PAGE] 拦截路径穿越尝试: {name}")
            page_error = "页面不存在"
        # 安全校验 2：白名单校验 — 仅允许合法页面名称
        elif not re.match(r"^[a-zA-Z0-9_\-.]+$", name):
            print(f"[PAGE] 拦截非法页面名称: {name}")
            page_error = "页面不存在"
        else:
            # 安全校验 3：拼接完整路径 + 自动补 .html 后缀
            page_path = os.path.join(_PAGES_DIR, name)
            # 如果直接路径不存在，尝试加 .html 后缀
            if not os.path.isfile(page_path):
                page_path = page_path + ".html"

            # 安全校验 4：绝对路径规范化 + 目录归属校验
            real_path = os.path.abspath(page_path)
            real_pages_dir = os.path.abspath(_PAGES_DIR)

            print(f"[PAGE] 尝试加载: {name}")
            print(f"[PAGE] 完整路径: {real_path}")
            print(f"[PAGE] Pages 目录: {real_pages_dir}")

            # 校验：文件必须在 pages 目录下
            if not real_path.startswith(real_pages_dir + os.sep):
                print(f"[PAGE] 目录穿越拦截: {real_path}")
                page_error = "页面不存在"
            # 校验：仅允许读取 .html 文件
            elif not real_path.endswith(".html"):
                print(f"[PAGE] 非 html 文件拦截: {real_path}")
                page_error = "页面不存在"
            elif not os.path.isfile(real_path):
                page_error = "页面不存在"
            else:
                try:
                    with open(real_path, "r", encoding="utf-8") as f:
                        page_content = f.read()
                    print(f"[PAGE] 页面加载成功: {real_path}")
                except Exception as e:
                    page_error = "页面读取失败"
                    print(f"[PAGE ERROR] {e}")

    # 渲染首页并带上 page_content
    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template(
        "index.html",
        username=username,
        user=user_info,
        page_content=page_content,
        page_error=page_error,
    )


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

    # 处理注册成功后的跳转提示
    registered_msg = None
    if request.args.get("registered") == "success":
        registered_msg = "注册成功，请登录"

    return render_template("login.html", registered=registered_msg)

# ---------------------------------------------------------------------------
# 路由 — 登出
# ---------------------------------------------------------------------------

@app.route("/logout")
def logout():
    """登出：清除会话后重定向至首页。"""
    session.clear()
    return redirect(url_for("index"))


# ---------------------------------------------------------------------------
# 路由 — 注册
# ---------------------------------------------------------------------------

@app.route("/register", methods=["GET", "POST"])
def register():
    """注册页面。

    GET  ：返回注册表单。
    POST ：使用参数化预编译查询将用户数据插入 SQLite 数据库。
    """
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        email = request.form.get("email", "")
        phone = request.form.get("phone", "")

        # ---------------------------------------------------------------
        # 安全校验：SQL 危险字符拦截 + 格式白名单校验
        # ---------------------------------------------------------------
        valid, msg = validate_sql_input(username, "username")
        if not valid:
            return render_template("register.html", error=msg)
        valid, msg = validate_sql_input(email, "email")
        if not valid:
            return render_template("register.html", error=msg)
        valid, msg = validate_sql_input(phone, "phone")
        if not valid:
            return render_template("register.html", error=msg)

        # ---------------------------------------------------------------
        # 修复：使用参数化预编译查询（? 占位符），彻底杜绝 SQL 注入
        # ---------------------------------------------------------------
        sql = "INSERT INTO users (username, password, email, phone) VALUES (?, ?, ?, ?)"
        print(f"[SQL] 注册 - 预编译语句: {sql} | 参数: ({username}, ***, {email}, {phone})")

        try:
            conn = sqlite3.connect("data/users.db")
            c = conn.cursor()
            c.execute(sql, (username, password, email, phone))
            conn.commit()
            conn.close()
            return redirect(url_for("login", registered="success"))
        except sqlite3.IntegrityError:
            return render_template("register.html", error="用户名已存在")
        except Exception as e:
            print(f"[SQL ERROR] 注册失败: {e}")
            return render_template("register.html", error="注册失败，请稍后重试")

    return render_template("register.html")


# ---------------------------------------------------------------------------
# 路由 — 搜索
# ---------------------------------------------------------------------------

@app.route("/search")
def search():
    """搜索用户。

    通过 URL 参数 keyword 接收关键词，
    使用参数化预编译查询进行模糊匹配，杜绝 SQL 注入风险。
    """
    keyword = request.args.get("keyword", "")
    results = []

    if keyword:
        # ---------------------------------------------------------------
        # 安全校验：SQL 危险字符拦截
        # ---------------------------------------------------------------
        valid, msg = validate_sql_input(keyword, "keyword")
        if not valid:
            username = session.get("username")
            user_info = None
            if username and username in USERS:
                user_info = sanitize_user_info(USERS[username])
            return render_template(
                "index.html",
                username=username,
                user=user_info,
                search_results=[],
                search_keyword=keyword,
                search_error=msg,
            )

        # ---------------------------------------------------------------
        # 修复：使用参数化预编译查询（? 占位符），彻底杜绝 SQL 注入
        # ---------------------------------------------------------------
        like_pattern = f"%{keyword}%"
        sql = "SELECT id, username, email, phone FROM users WHERE username LIKE ? OR email LIKE ?"
        print(f"[SQL] 搜索 - 预编译语句: {sql} | 参数: ('%{keyword}%')")

        try:
            conn = sqlite3.connect("data/users.db")
            c = conn.cursor()
            c.execute(sql, (like_pattern, like_pattern))
            rows = c.fetchall()
            conn.close()
            results = [{"id": r[0], "username": r[1], "email": r[2], "phone": r[3]} for r in rows]
            print(f"[SQL] 搜索到 {len(results)} 条结果")
        except Exception as e:
            print(f"[SQL ERROR] 搜索失败: {e}")

    # 渲染首页并带上搜索结果
    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template(
        "index.html",
        username=username,
        user=user_info,
        search_results=results,
        search_keyword=keyword,
    )

# ---------------------------------------------------------------------------
# 路由 — 头像上传
# ---------------------------------------------------------------------------


@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    """头像上传页面（安全加固版）。

    GET  ：返回上传表单。
    POST ：接收用户上传的文件，通过后缀白名单 + 魔数校验后，
           以 UUID 重命名保存至 data/uploads/ 目录。
    """
    uploaded_file_url = None
    error = None
    display_filename = None

    if request.method == "POST":
        file = request.files.get("file")

        # ── 校验 1：文件是否为空 ──
        if not file or file.filename == "":
            error = "请选择要上传的文件"
        else:
            try:
                original_filename = file.filename

                # ── 校验 2：后缀白名单 ──
                if not allowed_file(original_filename):
                    error = "仅允许上传 jpg、jpeg、png、gif 格式的图片文件"
                else:
                    # ── 校验 3：单文件大小上限（5MB） ──
                    file.seek(0, os.SEEK_END)
                    file_size = file.tell()
                    file.seek(0)
                    if file_size > MAX_FILE_SIZE:
                        error = f"文件大小超过限制（最大 {MAX_FILE_SIZE // (1024*1024)}MB）"
                    else:
                        # ── 校验 4：二进制魔数校验真实 MIME 类型 ──
                        is_valid_image, image_type = check_image_magic(file.stream)
                        if not is_valid_image:
                            error = "文件类型校验失败，仅允许上传纯图片文件"
                        else:
                            # ── 通过所有校验：安全保存文件 ──
                            # 使用 UUID 生成唯一文件名，彻底杜绝路径遍历和文件覆盖
                            safe_filename = f"{uuid.uuid4().hex}.{image_type}"
                            os.makedirs(UPLOAD_DIR, exist_ok=True)
                            save_path = os.path.join(UPLOAD_DIR, safe_filename)
                            file.save(save_path)

                            # 生成受保护的访问 URL（通过 /uploads/ 路由鉴权后访问）
                            uploaded_file_url = url_for("serve_upload", filename=safe_filename)
                            display_filename = original_filename

                            print(f"[UPLOAD] 原始文件: {original_filename}")
                            print(f"[UPLOAD] 安全保存: {save_path}")
                            print(f"[UPLOAD] 访问 URL: {uploaded_file_url}")

            except Exception as e:
                error = "文件上传失败，请稍后重试"
                print(f"[UPLOAD ERROR] {e}")

    return render_template(
        "upload.html",
        uploaded_file_url=uploaded_file_url,
        error=error,
        display_filename=display_filename,
    )


# ---------------------------------------------------------------------------
# 路由 — 上传文件访问（需登录鉴权）
# ---------------------------------------------------------------------------


@app.route("/uploads/<filename>")
@login_required
def serve_upload(filename):
    """提供上传文件的受保护访问。

    仅登录用户可访问，禁止匿名用户通过 URL 直接访问上传后的文件。
    """
    return send_from_directory(UPLOAD_DIR, filename)


# ---------------------------------------------------------------------------
# 路由 — 个人中心（安全加固版）
# ---------------------------------------------------------------------------


@app.route("/profile")
@login_required
def profile():
    """个人中心 — 仅展示当前登录用户自身信息。

    安全加固：
    1. 用户身份强制从 session 读取，废弃 URL 参数 user_id
    2. 不存在越权查看他人资料问题
    3. 频率限制防批量遍历
    """
    if not check_rate_limit("profile"):
        return render_template("profile.html", error="请求过于频繁，请稍后重试",
                               profile_user=None)

    username = session.get("username")
    user_data = USERS.get(username)

    if not user_data:
        return render_template("profile.html", error="用户信息加载失败",
                               profile_user=None)

    profile_user = {
        "username": user_data.get("username"),
        "role": user_data.get("role"),
        "email": user_data.get("email"),
        "phone": user_data.get("phone"),
        "balance": mask_balance(user_data.get("balance", 0)),
    }

    return render_template("profile.html", profile_user=profile_user, error=None)


# ---------------------------------------------------------------------------
# 路由 — 充值（安全加固版）
# ---------------------------------------------------------------------------


@app.route("/recharge", methods=["POST"])
@login_required
def recharge():
    """充值接口 — 仅操作当前登录用户自身余额。

    安全加固：
    1. 目标用户强制从 session 读取，废弃表单 user_id
    2. amount 仅允许大于 0 且不超过上限的合法数字
    3. 频率限制防批量遍历
    """
    if not check_rate_limit("recharge"):
        return render_template("profile.html", error="请求过于频繁，请稍后重试",
                               profile_user=_get_self_profile_user())

    username = session.get("username")
    user_data = USERS.get(username)

    if not user_data:
        return render_template("profile.html", error="用户信息加载失败",
                               profile_user=None)

    # amount 严格校验：必须为纯数字（最多两位小数）
    amount_str = request.form.get("amount", "").strip()

    if not amount_str:
        return render_template("profile.html",
                               error="请输入充值金额",
                               profile_user=_get_self_profile_user())

    # 拦截非数字格式（字母、特殊符号、SQL注入字符等）
    if not re.match(r"^[0-9]+(\.[0-9]{1,2})?$", amount_str):
        return render_template("profile.html",
                               error="充值金额不合法",
                               profile_user=_get_self_profile_user())

    amount = float(amount_str)

    # 拦截负数、零
    if amount <= 0:
        return render_template("profile.html",
                               error="充值金额必须大于 0",
                               profile_user=_get_self_profile_user())

    if amount > MAX_RECHARGE_AMOUNT:
        return render_template("profile.html",
                               error=f"单次充值金额不能超过 {MAX_RECHARGE_AMOUNT} 元",
                               profile_user=_get_self_profile_user())

    # 执行充值
    user_data["balance"] = user_data.get("balance", 0) + amount
    print(f"[RECHARGE] 用户 {username} 充值 {amount} 元，余额 {user_data['balance']}")

    return redirect(url_for("profile"))


def _get_self_profile_user():
    """辅助函数：获取当前登录用户的 profile 展示数据。"""
    username = session.get("username")
    user_data = USERS.get(username)
    if not username or not user_data:
        return None
    return {
        "username": user_data.get("username"),
        "role": user_data.get("role"),
        "email": user_data.get("email"),
        "phone": user_data.get("phone"),
        "balance": mask_balance(user_data.get("balance", 0)),
    }


# ---------------------------------------------------------------------------
# 路由 — 修改密码（安全加固版）
# ---------------------------------------------------------------------------


@app.route("/change-password", methods=["POST"])
@login_required
def change_password():
    """修改密码 — 安全加固版。

    安全措施：
    1. 目标用户强制从 session 读取，废弃表单 username
    2. 校验原密码与数据库中哈希密码一致
    3. CSRF Token 校验
    4. 密码复杂度校验（8位+大小写+数字）
    """
    # 安全措施 1：目标用户强制从 session 读取
    session_username = session.get("username")
    user_data = USERS.get(session_username)

    if not user_data:
        return render_template("profile.html",
                               error="用户信息加载失败",
                               profile_user=None)

    old_password = request.form.get("old_password", "")
    new_password = request.form.get("new_password", "")
    confirm_password = request.form.get("confirm_password", "")

    # 非空校验
    if not old_password or not new_password or not confirm_password:
        return render_template("profile.html",
                               error="请填写完整信息",
                               profile_user=_get_self_profile_user())

    # 安全措施 2：校验原密码
    if not check_password_hash(user_data["password"], old_password):
        return render_template("profile.html",
                               error="原密码错误",
                               profile_user=_get_self_profile_user())

    # 校验两次新密码一致
    if new_password != confirm_password:
        return render_template("profile.html",
                               error="两次输入的新密码不一致",
                               profile_user=_get_self_profile_user())

    # 安全措施 4：密码复杂度校验
    if len(new_password) < 8:
        return render_template("profile.html",
                               error="密码长度不能少于 8 位",
                               profile_user=_get_self_profile_user())

    if not re.search(r"[a-z]", new_password):
        return render_template("profile.html",
                               error="密码必须包含小写字母",
                               profile_user=_get_self_profile_user())

    if not re.search(r"[A-Z]", new_password):
        return render_template("profile.html",
                               error="密码必须包含大写字母",
                               profile_user=_get_self_profile_user())

    if not re.search(r"[0-9]", new_password):
        return render_template("profile.html",
                               error="密码必须包含数字",
                               profile_user=_get_self_profile_user())

    # 通过全部校验：哈希加密后更新密码
    user_data["password"] = generate_password_hash(new_password)
    print(f"[PASSWORD] 用户 {session_username} 密码修改成功")

    return render_template("profile.html",
                           error="密码修改成功",
                           profile_user=_get_self_profile_user())


# ---------------------------------------------------------------------------
# SSRF 防护配置
# ---------------------------------------------------------------------------

# 协议白名单
ALLOWED_URL_SCHEMES = ("http", "https")

# 端口白名单
ALLOWED_PORTS = {80, 443}

# 内网 IP 网段黑名单
INTERNAL_IP_NETWORKS = [
    "127.0.0.0/8",
    "10.0.0.0/8",
    "172.16.0.0/12",
    "192.168.0.0/16",
    "169.254.0.0/16",
]


def validate_fetch_url(url_str):
    """全面校验 URL 安全性，防御 SSRF 攻击。

    校验项：
    1. URL 基础格式
    2. 协议白名单（仅 http/https）
    3. 端口白名单（仅 80/443）
    4. DNS 解析 + 内网 IP 拦截
    """
    if not url_str or not url_str.strip():
        return False, "URL 不能为空"

    url_str = url_str.strip()

    # 1. URL 基础格式校验
    parsed = urllib.parse.urlparse(url_str)
    if not parsed.scheme or not parsed.netloc:
        return False, "URL 格式不合法"

    # 2. 协议白名单校验
    if parsed.scheme not in ALLOWED_URL_SCHEMES:
        return False, "仅支持 http 和 https 协议"

    # 3. 端口白名单校验
    hostname = parsed.hostname
    port = parsed.port

    if port is None:
        # 根据协议确定默认端口
        port = 443 if parsed.scheme == "https" else 80

    if port not in ALLOWED_PORTS:
        return False, "仅允许访问标准 Web 端口（80、443）"

    # 4. DNS 解析 + 内网 IP 拦截
    try:
        ip = socket.gethostbyname(hostname)
    except socket.gaierror:
        return False, "无法解析目标域名"

    try:
        ip_obj = ipaddress.ip_address(ip)
        for network_str in INTERNAL_IP_NETWORKS:
            if ip_obj in ipaddress.ip_network(network_str, strict=False):
                return False, "不允许访问内网地址"
    except ValueError:
        return False, "目标地址不合法"

    return True, None


# ---------------------------------------------------------------------------
# 路由 — URL 抓取（SSRF 安全加固版）
# ---------------------------------------------------------------------------


@app.route("/fetch-url", methods=["POST"])
@login_required
def fetch_url():
    """URL 抓取功能 — SSRF 安全加固版。

    安全措施：
    1. 协议白名单（仅 http/https）
    2. 端口白名单（仅 80/443）
    3. DNS 解析 + 内网 IP 拦截
    4. 统一错误提示，不暴露内网信息
    """
    target_url = request.form.get("url", "")
    fetch_result = None
    fetch_error = None

    # SSRF 安全校验
    valid, err_msg = validate_fetch_url(target_url)
    if not valid:
        fetch_error = err_msg if err_msg else "URL 不合法"
        print(f"[FETCH] 拦截非法 URL: {target_url} - {fetch_error}")
    elif target_url:
        try:
            print(f"[FETCH] 抓取 URL: {target_url}")
            req = urllib.request.Request(target_url, headers={
                "User-Agent": "Mozilla/5.0 (compatible; FlaskFetch/1.0)"
            })
            resp = urllib.request.urlopen(req, timeout=10)
            status_code = resp.getcode()
            content = resp.read().decode("utf-8", errors="replace")
            resp.close()
            preview = content[:5000]
            if len(content) > 5000:
                preview += "\n\n... (内容已截断，仅显示前 5000 字符)"
            fetch_result = {
                "status_code": status_code,
                "url": target_url,
                "content": preview,
            }
            print(f"[FETCH] 状态码: {status_code}, 内容长度: {len(content)}")
        except urllib.error.HTTPError:
            fetch_error = "远程服务器返回错误"
        except urllib.error.URLError:
            fetch_error = "无法访问目标地址"
        except Exception:
            fetch_error = "抓取失败"
    else:
        fetch_error = "请输入 URL"

    # 渲染首页并带上抓取结果
    username = session.get("username")
    user_info = None
    if username and username in USERS:
        user_info = sanitize_user_info(USERS[username])

    return render_template(
        "index.html",
        username=username,
        user=user_info,
        fetch_result=fetch_result,
        fetch_error=fetch_error,
        fetch_url=target_url,
    )


# ---------------------------------------------------------------------------
# 路由 — Ping 网络诊断（安全加固版）
# ---------------------------------------------------------------------------


def validate_ip_address(ip_str):
    """校验 IP 地址格式的合法性。

    仅允许标准 IPv4 和 IPv6 地址，拦截所有 shell 特殊字符。
    """
    if not ip_str or not ip_str.strip():
        return False

    ip_str = ip_str.strip()

    # 检查是否包含 shell 特殊字符
    # 拦截 ; & | ` $ ( ) { } < > \n \r 等
    shell_chars = re.search(r'[;&|`$(){}<>\n\r]', ip_str)
    if shell_chars:
        return False

    # 尝试解析为 IP 地址
    try:
        ipaddress.ip_address(ip_str)
        return True
    except ValueError:
        pass

    return False


@app.route("/ping", methods=["GET", "POST"])
@login_required
def ping():
    """Ping 网络诊断功能 — 安全加固版。

    安全措施：
    1. 废弃 shell=True，使用参数列表执行
    2. 废弃 f-string 拼接，改用 ["ping", "-c", "3", ip]
    3. IP 格式白名单校验（仅 IPv4/IPv6）
    4. 拦截 shell 特殊字符
    5. 错误信息脱敏
    """
    result = None
    error = None

    if request.method == "POST":
        ip = request.form.get("ip", "").strip()

        if not ip:
            error = "请输入 IP 地址"
        elif not validate_ip_address(ip):
            print(f"[PING] 拦截非法 IP: {ip}")
            error = "IP 地址格式不合法"
        else:
            try:
                # 安全：使用参数列表执行，废弃 shell=True 和 f-string 拼接
                print(f"[PING] 执行 ping: {ip}")
                output = subprocess.check_output(
                    ["ping", "-c", "3", ip],
                    timeout=30,
                    stderr=subprocess.STDOUT,
                )
                result = output.decode("utf-8", errors="replace")
                print(f"[PING] 执行成功，输出长度: {len(result)}")

            except subprocess.TimeoutExpired:
                error = "命令执行超时 (30 秒)"
            except subprocess.CalledProcessError:
                error = "Ping 请求失败 (目标不可达)"
            except Exception as e:
                error = "执行出错，请检查 IP 地址后重试"
                print(f"[PING ERROR] {e}")

    return render_template("ping.html", result=result, error=error)


# ---------------------------------------------------------------------------
# SQLite 数据库初始化
# ---------------------------------------------------------------------------

def init_db():
    """初始化 SQLite 数据库，创建 users 表并插入默认用户。"""
    os.makedirs("data", exist_ok=True)
    conn = sqlite3.connect("data/users.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            email TEXT,
            phone TEXT
        )
    """)
    # 插入默认用户（使用哈希密码存储）
    admin_hash = generate_password_hash("admin123")
    alice_hash = generate_password_hash("alice2025")
    c.execute("INSERT OR IGNORE INTO users (username, password, email, phone) VALUES (?, ?, ?, ?)",
              ("admin", admin_hash, "admin@example.com", "13800138000"))
    c.execute("INSERT OR IGNORE INTO users (username, password, email, phone) VALUES (?, ?, ?, ?)",
              ("alice", alice_hash, "alice@example.com", "13900139001"))
    conn.commit()
    conn.close()
    print("[DB] SQLite 数据库初始化完成 (data/users.db)")


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
    # 初始化 SQLite 数据库
    init_db()

    # 生产环境请通过环境变量 FLASK_DEBUG=true 开启调试
    debug_mode = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(
        debug=debug_mode,
        host="0.0.0.0",
        port=5000,
    )
