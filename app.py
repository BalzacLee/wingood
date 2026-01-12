from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import requests
import logging
from datetime import datetime
import os
import time
from collections import defaultdict

# -------------------------- 1. 基础配置 --------------------------
app = Flask(__name__)
CORS(app)  # 解决跨域问题（前端访问后端）

# 复用原日志配置
def setup_logger():
    log_filename = f"开闸工具日志_{datetime.now().strftime('%Y%m%d')}.txt"
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), log_filename)
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()]
    )
    return logging.getLogger("ParkingGateTool")

logger = setup_logger()

# 复用原闸门配置
BASE_URL = "https://city.wingoodcloud.com"
GATE_CONFIG = [
    {"name": "百乐园-大门出口", "id": "101", "account": "liufu01", "password": "111111", "sentryNo": "blya01"},
    # 省略其他闸门配置（直接复制原代码的GATE_CONFIG）
    {"name": "元下田B-出入", "id": "153", "account": "xiaorui", "password": "xr$RFV5tgb", "sentryNo": "yxtdj2"},
]

# 复用原登录函数
def login(account, password):
    logger.info(f"开始登录账号：{account}")
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36 Edg/142.0.0.0",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": f"{BASE_URL}/login.html",
        "Origin": BASE_URL,
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Accept": "application/json, text/javascript, */*; q=0.01",
    }
    try:
        precheck_resp = session.post(f"{BASE_URL}/LoginUserName", data={"userName": account}, headers=headers, timeout=10)
        precheck_resp.raise_for_status()
        login_resp = session.post(f"{BASE_URL}/Login", data={"userName": account, "password": password}, headers=headers, timeout=10)
        login_resp.raise_for_status()
        login_result = login_resp.json()
        if login_result.get("flag"):
            timestamp = str(int(time.time()))
            session.cookies.set("Hm_lpvt_b393d153aeb26b46e9431fabaf0f6190", timestamp)
            logger.info(f"账号{account}登录成功！")
            return session
        else:
            err_msg = f"账号{account}登录失败：{login_result.get('msg', '未知错误')}"
            logger.error(err_msg)
            return None
    except Exception as e:
        err_msg = f"登录异常：{str(e)}"
        logger.error(err_msg, exc_info=True)
        return None

# 复用原开闸函数
def open_gate_logic(selected_gate):
    session = login(selected_gate["account"], selected_gate["password"])
    if not session:
        return {"flag": False, "msg": "登录失败"}
    logger.info(f"开始执行开闸操作：{selected_gate['name']}（ID：{selected_gate['id']}）")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36 Edg/143.0.0.0",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": f"{BASE_URL}/html/system/park-line-list.html?sentryNo={selected_gate['sentryNo']}",
        "Origin": BASE_URL,
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Accept": "application/json, text/plain, */*",
    }
    try:
        resp = session.post(f"{BASE_URL}/ajax/ajaxOpenParkLane", data={"id": selected_gate["id"]}, headers=headers, timeout=10)
        resp.raise_for_status()
        open_result = resp.json()
        return open_result
    except Exception as e:
        err_msg = f"开闸异常：{str(e)}"
        logger.error(err_msg, exc_info=True)
        return {"flag": False, "msg": err_msg}
    finally:
        session.close()

# -------------------------- 2. Flask接口 --------------------------
# 接口1：获取所有闸门配置（供前端渲染按钮）
@app.route("/api/gates", methods=["GET"])
def get_gates():
    # 按账号分组返回（方便前端分组展示）
    grouped_gates = defaultdict(list)
    for gate in GATE_CONFIG:
        grouped_gates[gate["account"]].append(gate)
    return jsonify(dict(grouped_gates))

# 接口2：执行开闸操作
@app.route("/api/open-gate", methods=["POST"])
def open_gate():
    gate_id = request.json.get("id")
    # 找到对应闸门配置
    selected_gate = next((g for g in GATE_CONFIG if g["id"] == gate_id), None)
    if not selected_gate:
        return jsonify({"flag": False, "msg": "未找到该闸门"})
    # 执行开闸逻辑
    result = open_gate_logic(selected_gate)
    return jsonify(result)

# -------------------------- 3. 前端页面（移动端适配） --------------------------
@app.route("/")
def index():
    return render_template("index.html")

# -------------------------- 程序入口 --------------------------
if __name__ == "__main__":
    # 运行Flask服务，host=0.0.0.0允许局域网访问
    app.run(host="0.0.0.0", port=5000, debug=True)
