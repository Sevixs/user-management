<div align="center">

# 🛡️ 用户管理系统 — Flask 安全重构版

<p>
  <img src="https://img.shields.io/badge/Python-3.10%2B-blue" alt="Python 3.10+">
  <img src="https://img.shields.io/badge/Flask-3.1-green" alt="Flask 3.1">
  <img src="https://img.shields.io/badge/license-MIT-yellow" alt="MIT License">
</p>

**从「裸奔 Demo」到「生产安全意识示范」的完整重构案例**

</div>

---

## 📋 目录

- [项目简介](#-项目简介)
- [原始漏洞清单](#-原始漏洞清单)
- [漏洞修复详解](#-漏洞修复详解)
- [技术栈](#-技术栈)
- [环境依赖](#-环境依赖)
- [本地部署](#-本地部署)
- [项目文件结构](#-项目文件结构)
- [功能展示](#-功能展示)
- [安全优化总结](#-安全优化总结)
- [后续安全升级建议](#-后续安全升级建议)
- [开源声明](#-开源声明)

---

## 📖 项目简介

本项目是一个基于 **Python Flask** 框架的简易用户信息管理平台，实现了**用户登录、信息展示、登出**等基础功能，并区分 **admin** 与 **user** 两种角色。

**关键定位：** 该项目最初以「全漏洞」形式存在，用于演示常见 Web 安全风险，现已完成**全量安全加固**，可作为安全编码教学范例或轻量级 Web 应用脚手架使用。

---

## 🚨 原始漏洞清单

原始项目存在以下 **7 大类高危安全漏洞**：

| 编号 | 漏洞类型 | 严重程度 | 描述 |
|------|----------|----------|------|
| ① | **密码明文存储 & 明文比对** | 🔴 高危 | 密码以明文形式存储在 `USERS` 字典中，登录时直接用 `==` 比对字符串 |
| ② | **HTML 注释泄露默认账号** | 🔴 高危 | 登录页 HTML 源码中硬编码了 `<!-- 调试信息 - 默认管理员账号 用户名: admin 密码: admin123 -->` |
| ③ | **XSS 跨站脚本攻击** | 🔴 高危 | 用户输入未做任何过滤，`{{ user.phone }}` 直接输出可能导致反射型 / 存储型 XSS |
| ④ | **无 CSRF 防护** | 🔴 高危 | 登录请求无 CSRF Token 校验，可被跨站请求伪造攻击 |
| ⑤ | **前端泄露完整敏感信息** | 🟠 中危 | 密码、手机号、余额等敏感信息直接回显在前端页面，浏览器缓存 / 历史记录可被窃取 |
| ⑥ | **无输入校验 & 无异常处理** | 🟠 中危 | 无输入过滤、无 try/except 包裹、无错误处理、无 404/500 处理器 |
| ⑦ | **开发密钥 & Debug 模式常开** | 🟠 中危 | `secret_key = "dev-key-2025"` 硬编码暴露；`debug=True` 在生产环境会泄露调用栈 |

---

## 🔧 漏洞修复详解

### ① 密码安全重构 — 明文 → Werkzeug 哈希加密

**问题：**
```python
# ❌ 原代码：明文存储 + == 比对
USERS = {"admin": {"password": "admin123"}}
if USERS[username]["password"] == password:
```

**修复方案：**
- 使用 `werkzeug.security.generate_password_hash()` 生成 **pbkdf2:sha256** 哈希
- 使用 `werkzeug.security.check_password_hash()` 验密
- 数据库中仅存储哈希值，内存中也无明文

```python
# ✅ 修复后：哈希存储 + 安全验密
from werkzeug.security import generate_password_hash, check_password_hash

# 构建数据库时立即哈希
record["password"] = generate_password_hash(record["password"])

# 登录验证时哈希比对
if user_record and check_password_hash(user_record["password"], password):
```

**修复位置：** `app.py` 中的 `_build_user_db()` 函数与 `/login` 路由

---

### ② 信息泄露修复 — 删除注释 & 敏感信息脱敏

**问题：**
```html
<!-- ❌ HTML 注释泄露 admin/admin123 -->
<!-- 调试信息 - 默认管理员账号 用户名: admin 密码: admin123 -->
{{ user.password }}  <!-- ❌ 密码直接在前端渲染 -->
13800138000         <!-- ❌ 手机号完整展示 -->
```

**修复方案：**
- 删除登录页全部调试注释
- 前端不再传递、展示 `password` 字段
- 手机号脱敏为 `138****8000` 格式
- 余额脱敏为 `¥99,999.00` 货币格式

```python
# ✅ app.py — sanitize_user_info() 脱敏函数
def sanitize_user_info(user_info):
    return {
        "phone": user_info["phone"][:3] + "****" + user_info["phone"][-4:],
        "balance": "¥{:,.2f}".format(float(user_info["balance"])),
        # 不包含 password 键
    }
```

**修复位置：** `app.py` → `sanitize_user_info()` / `mask_phone()` / `mask_balance()`  
**修复位置：** `templates/index.html` — 移除了 `{{ user.password }}`

---

### ③ XSS 防护 — 输入过滤 + Jinja2 自动转义

**问题：**
```python
# ❌ 直接取用用户输入
username = request.form.get("username")
```

**修复方案：**
- 自定义 `sanitize_input()` 函数移除 HTML 标签及危险协议
- 表单输入增加 `maxlength` 属性
- Flask 的 Jinja2 模板引擎默认启用 `autoescape`，`{{ ... }}` 会自动转义 HTML

```python
# ✅ app.py — sanitize_input() 过滤函数
def sanitize_input(text):
    text = re.sub(r"<[^>]*>", "", text)            # 移除 HTML 标签
    text = re.sub(r"javascript\s*:", "", text, flags=re.IGNORECASE)
    return text.strip()

# 使用时
username = sanitize_input(request.form.get("username", ""))
```

**修复位置：** `app.py` → `sanitize_input()`

---

### ④ CSRF 防护 — 令牌验证

**问题：**
```html
<!-- ❌ 无 CSRF 令牌 -->
<form method="post" action="/login">
```

**修复方案：**
- 使用 `secrets.token_hex(32)` 生成令牌
- 通过 `@app.before_request` 全局拦截 POST 请求校验
- 通过 `@app.context_processor` 将令牌注入所有模板

```python
# ✅ app.py — CSRF 校验
@app.before_request
def _csrf_protect():
    if request.method == "POST":
        token = request.form.get("csrf_token")
        stored = session.get("csrf_token")
        if not token or not stored or token != stored:
            abort(400, "CSRF token 缺失或无效")

# ✅ templates/login.html — 表单内添加隐藏字段
<input type="hidden" name="csrf_token" value="{{ csrf_token }}">
```

**修复位置：** `app.py` → `_csrf_protect()` / `_inject_csrf_token()`  
**修复位置：** `templates/login.html` → hidden input

---

### ⑤ 会话保护 — 超时 & 安全标志 & 登录鉴权

**问题：**
```python
# ❌ 无会话超时、无鉴权拦截
app.secret_key = "dev-key-2025"
```

**修复方案：**
- 设置 `PERMANENT_SESSION_LIFETIME = 2 小时`
- `SESSION_COOKIE_HTTPONLY = True`（禁止 JavaScript 读取）
- `SESSION_COOKIE_SAMESITE = "Lax"`（禁止跨站请求携带 Cookie）
- 登录成功后 `session.clear()` 重置会话，防会话固定攻击
- 新增 `@login_required` 装饰器保护需鉴权的路由

```python
# ✅ app.py — 安全配置
app.config.update(
    SECRET_KEY=os.environ.get("SECRET_KEY", secrets.token_hex(32)),
    PERMANENT_SESSION_LIFETIME=timedelta(hours=2),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
)

# ✅ app.py — 鉴权装饰器
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "username" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated
```

**修复位置：** `app.py` → 应用配置 / `login_required()` / 登录路由中 `session.clear()`

---

### ⑥ 异常处理 & 输入校验

**问题：**
```python
# ❌ 无校验、无异常捕获
if username in USERS and USERS[username]["password"] == password:
```

**修复方案：**
- 空值校验：用户名密码为空时提前返回错误
- 登录逻辑包裹 `try/except` 捕获异常并记录日志
- 全局 400/404/500 错误处理器

```python
# ✅ app.py — 异常处理
try:
    user_record = USERS.get(username)
    if user_record and check_password_hash(user_record["password"], password):
        ...
except Exception as e:
    app.logger.error("登录异常: %s", str(e))
    return render_template("login.html", error="系统异常，请稍后重试")

# 错误处理器
@app.errorhandler(500)
def server_error(e):
    return render_template("login.html", error="服务器内部错误"), 500
```

**修复位置：** `app.py` → 登录路由 try/except / 全局错误处理器

---

### ⑦ 配置安全 — 移除硬编码密钥 & Debug 模式

**问题：**
```python
# ❌ 硬编码开发密钥 & 强制 debug=True
app.secret_key = "dev-key-2025"
app.run(debug=True)
```

**修复方案：**
- 生产密钥通过环境变量 `SECRET_KEY` 注入
- 未设置时自动生成 32 字节随机密钥
- Debug 模式通过环境变量 `FLASK_DEBUG` 控制，默认关闭

```python
# ✅ app.py
app.config.update(
    SECRET_KEY=os.environ.get("SECRET_KEY", secrets.token_hex(32)),
)

if __name__ == "__main__":
    debug_mode = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(debug=debug_mode, host="0.0.0.0", port=5000)
```

**修复位置：** `app.py` → 应用配置 / 启动入口

---

## 🧰 技术栈

| 组件 | 技术 | 用途 |
|------|------|------|
| 后端框架 | **Flask 3.1** | Web 框架 |
| 密码加密 | **Werkzeug 3.1** | `generate_password_hash` / `check_password_hash` |
| 模板引擎 | **Jinja2**（Flask 内置） | HTML 模板渲染 |
| 前端样式 | **纯 CSS** | Flex 布局 / 渐变 / 卡片设计 |
| CSRF 防护 | **secrets 模块** | Flask 无插件 CSRF 自实现 |

---

## 📦 环境依赖

### Python 版本

- Python **3.10** 或更高版本

### pip 依赖

```txt
Flask==3.1.*
Werkzeug==3.1.*
```

> 依赖文件位于项目根目录 `requirements.txt`。

---

## 🚀 本地部署

### 从零开始运行

```bash
# 1. 克隆仓库
git clone https://github.com/<你的用户名>/user-management.git
cd user-management

# 2. 创建虚拟环境（推荐）
python -m venv venv
source venv/bin/activate      # Linux / macOS
# 或
venv\Scripts\activate         # Windows

# 3. 安装依赖
pip install -r requirements.txt

# 4. （可选）设置安全密钥
export SECRET_KEY="your-strong-secret-key-here"   # Linux / macOS
# 或
set SECRET_KEY="your-strong-secret-key-here"      # Windows
# 注：不设置则自动生成随机密钥

# 5. 启动应用
python app.py

# 6. 打开浏览器访问
# ➜  http://localhost:5000
```

### 调试模式

```bash
# 开启 Flask 调试模式（显示详细错误信息）
export FLASK_DEBUG=true
python app.py
```

### 测试账号

| 用户名 | 密码 | 角色 |
|--------|------|------|
| `admin` | `admin123` | 管理员 |
| `alice` | `alice2025` | 普通用户 |

> **安全提示:** 本项目为安全教学演示用途。生产环境请使用强随机密码、HTTPS、数据库后端存储。

---

## 📁 项目文件结构

```
user-management/
│
├── app.py                     # Flask 主应用（安全加固版）
├── requirements.txt           # Python 依赖清单
├── .gitignore                 # Git 忽略规则
├── README.md                  # 项目文档（就是本文件）
│
├── templates/                 # HTML 模板
│   ├── base.html              # 基础模板（导航栏 + 容器）
│   ├── login.html             # 登录页（卡片表单 + CSRF）
│   └── index.html             # 首页（用户信息展示 / 未登录提示）
│
└── static/
    └── css/
        └── style.css          # 全局样式（渐变导航栏 + 卡片 + 表单）
```

---

## ✨ 功能展示

### 🔐 登录功能

- 访问 `/login` 进入登录页
- 输入用户名和密码后提交
- 校验通过后自动跳转至首页并展示用户信息
- 校验失败显示"用户名或密码错误"（**统一错误提示，不暴露用户名是否存在**）

### 🏠 首页 / 用户信息

- 已登录：显示"欢迎回来，<用户名>！"
- 展示用户信息：**用户名、角色、邮箱、脱敏手机号、脱敏余额**
- 包含"退出登录"按钮

### 👥 角色区分

- **admin** — 管理员角色
- **user** — 普通用户角色
- （框架已预留，可按需扩展角色权限逻辑）

### 🚪 登出功能

- 点击"退出"或"退出登录"即可清除会话并跳转到首页
- 登出后页面自动显示"请先登录"提示

---

## 🛡️ 安全优化总结

| 安全维度 | 修复前 | 修复后 |
|----------|--------|--------|
| 密码存储 | 明文 | Werkzeug pbkdf2:sha256 哈希 |
| 密码比对 | `==` 字符串比对 | `check_password_hash()` |
| 前端密码展示 | 页面回显明文密码 | 不传递 password 字段 |
| 敏感信息 | 手机号、余额完整展示 | 手机号 `138****8000`，余额 `¥99,999.00` |
| 默认账号泄露 | HTML 注释硬编码 | 已删除 |
| XSS 防护 | 无 | 输入 HTML 过滤 + Jinja2 autoescape |
| CSRF 防护 | 无 | Token 校验 |
| 会话超时 | 无（永久有效） | 2 小时自动过期 |
| Cookie 安全 | 默认 | HttpOnly + SameSite=Lax |
| 会话固定攻击 | 可攻击 | 登录后 `session.clear()` |
| 登录鉴权 | 无装饰器 | `@login_required` 装饰器 |
| 异常处理 | 无 | try/except + 全局错误处理器 |
| Secret Key | 硬编码 `dev-key-2025` | 环境变量或随机生成 |
| Debug 模式 | 强制开启 | 环境变量控制，默认关闭 |

---

## 📈 后续安全升级建议

虽然本项目已完成基础安全加固，但在**生产环境**中还应考虑以下升级：

1. **数据库存储** — 将 `USERS` 字典替换为 SQLite / PostgreSQL + ORM（如 SQLAlchemy）
2. **HTTPS 强制** — 配置 Nginx 反向代理 + Let's Encrypt 证书，同时将 `SESSION_COOKIE_SECURE=True`
3. **Flask-Login** — 使用官方扩展代替手写 `login_required` 装饰器
4. **Flask-WTF / WTForms** — 使用成熟的 CSRF + 表单验证库
5. **速率限制** — 使用 Flask-Limiter 限制登录接口请求频率（防暴力破解）
6. **日志审计** — 记录登录成功/失败日志，对接集中式日志系统
7. **内容安全策略 (CSP)** — 设置 `Content-Security-Policy` 响应头
8. **双因素认证 (2FA)** — 增加 TOTP / 短信验证码二次验证
9. **密码策略** — 增加密码复杂度校验、定期更换提醒

---

## 📄 开源声明

```
MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

<div align="center">

**⚠️ 使用须知**

本项目为**安全编码教学示例**，展示常见 Web 漏洞及其修复方案。所有修复方案均遵循安全最佳实践，但不应被视为生产级安全解决方案的直接替代。**请勿将默认测试账号部署到公开生产环境。**

</div>
