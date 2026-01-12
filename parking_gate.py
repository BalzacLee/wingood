import tkinter as tk
from tkinter import ttk, messagebox
import requests
import logging
from datetime import datetime
import os
import time
from collections import defaultdict

# -------------------------- 1. 日志配置 --------------------------
def setup_logger():
    log_filename = f"开闸工具日志_{datetime.now().strftime('%Y%m%d')}.txt"
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), log_filename)
    
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("ParkingGateTool")

logger = setup_logger()

# -------------------------- 2. 全局配置 --------------------------
BASE_URL = "https://city.wingoodcloud.com"
GATE_CONFIG = [
    {"name": "百乐园-大门出口", "id": "101", "account": "liufu01", "password": "111111", "sentryNo": "blya01"},
    {"name": "百乐园-大门入口", "id": "100", "account": "liufu01", "password": "111111", "sentryNo": "blya01"},
    {"name": "百乐园-和平路出口", "id": "103", "account": "liufu01", "password": "111111", "sentryNo": "blya01"},
    {"name": "百乐园-和平路入口", "id": "105", "account": "liufu01", "password": "111111", "sentryNo": "blya01"},
    {"name": "百乐园-A2B1出入", "id": "104", "account": "liufu01", "password": "111111", "sentryNo": "blya01"},
    {"name": "百乐园-C夜市出入", "id": "77", "account": "liufu01", "password": "111111", "sentryNo": "blyc01"},
    {"name": "百乐园-C花鸟出入", "id": "89", "account": "liufu01", "password": "111111", "sentryNo": "blyc01"},
    {"name": "百乐园-D大门出口", "id": "161", "account": "liufu01", "password": "111111", "sentryNo": "blyc01"},
    {"name": "百乐园-D侧门出口", "id": "162", "account": "liufu01", "password": "111111", "sentryNo": "blyc01"},
    {"name": "中西医-保安亭", "id": "142", "account": "liufu", "password": "111111", "sentryNo": "zxyy01"},
    {"name": "中西医-菜市", "id": "140", "account": "liufu", "password": "111111", "sentryNo": "zxyy01"},
    {"name": "大涌-出入", "id": "132", "account": "liufu", "password": "111111", "sentryNo": "dyc01"},
    {"name": "元下田9-出入", "id": "130", "account": "liufu", "password": "111111", "sentryNo": "yxt01"},
    {"name": "北海-入A", "id": "122", "account": "liufu", "password": "111111", "sentryNo": "bh01"},
    {"name": "北海-入B", "id": "112", "account": "liufu", "password": "111111", "sentryNo": "bh01"},
    {"name": "北海-出", "id": "111", "account": "liufu", "password": "111111", "sentryNo": "bh01"},
    {"name": "霄边-出入", "id": "109", "account": "xiaorui", "password": "xr$RFV5tgb", "sentryNo": "dg03"},
    {"name": "元下田A-出入", "id": "151", "account": "xiaorui", "password": "xr$RFV5tgb", "sentryNo": "yxtdj"},
    {"name": "元下田B-出入", "id": "153", "account": "xiaorui", "password": "xr$RFV5tgb", "sentryNo": "yxtdj2"},
]

# -------------------------- 3. 核心功能函数 --------------------------
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
    logger.debug(f"登录请求头：{headers}")

    try:
        precheck_resp = session.post(
            f"{BASE_URL}/LoginUserName",
            data={"userName": account},
            headers=headers,
            timeout=10
        )
        precheck_resp.raise_for_status()
        logger.debug(f"预校验响应状态码：{precheck_resp.status_code}")
        logger.debug(f"预校验响应内容：{precheck_resp.text}")

        login_resp = session.post(
            f"{BASE_URL}/Login",
            data={"userName": account, "password": password},
            headers=headers,
            timeout=10
        )
        login_resp.raise_for_status()
        logger.debug(f"登录响应状态码：{login_resp.status_code}")
        logger.debug(f"登录响应内容：{login_resp.text}")
        logger.debug(f"登录后Cookie：{dict(session.cookies)}")

        login_result = login_resp.json()
        if login_result.get("flag"):
            timestamp = str(int(time.time()))
            session.cookies.set("Hm_lpvt_b393d153aeb26b46e9431fabaf0f6190", timestamp)
            logger.info(f"账号{account}登录成功！")
            return session
        else:
            err_msg = f"账号{account}登录失败：{login_result.get('msg', '未知错误')}"
            logger.error(err_msg)
            messagebox.showerror("登录失败", err_msg)
            return None

    except requests.exceptions.HTTPError as e:
        err_msg = f"请求失败：HTTP状态错误 {str(e)}"
        logger.error(err_msg, exc_info=True)
        messagebox.showerror("网络异常", f"登录请求被拒绝：{str(e)}\n详细日志见本地文件")
        return None
    except requests.exceptions.Timeout:
        err_msg = "登录请求超时，请检查网络连接"
        logger.error(err_msg)
        messagebox.showerror("网络超时", err_msg)
        return None
    except Exception as e:
        err_msg = f"登录异常：{str(e)}"
        logger.error(err_msg, exc_info=True)
        messagebox.showerror("系统异常", f"登录时出错：{str(e)}\n详细日志见本地文件")
        return None

def open_gate(selected_gate):
    if not selected_gate:
        messagebox.showwarning("提示", "请选择一个闸门！")
        return

    session = login(selected_gate["account"], selected_gate["password"])
    if not session:
        return

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
        open_gate_data = {"id": selected_gate["id"]}
        resp = session.post(
            f"{BASE_URL}/ajax/ajaxOpenParkLane",
            data=open_gate_data,
            headers=headers,
            timeout=10
        )
        resp.raise_for_status()
        logger.debug(f"开闸响应状态码：{resp.status_code}")
        logger.debug(f"开闸响应内容：{resp.text}")

        open_result = resp.json()
        if open_result.get("flag"):
            logger.info(f"{selected_gate['name']}开闸成功！{open_result.get('msg')}")
            messagebox.showinfo("成功", f"{selected_gate['name']}开闸成功！\n{open_result.get('msg')}")
        else:
            err_msg = f"{selected_gate['name']}开闸失败：{open_result.get('msg', '未知错误')}"
            logger.error(err_msg)
            messagebox.showerror("失败", err_msg)

    except requests.exceptions.HTTPError as e:
        err_msg = f"{selected_gate['name']}开闸请求失败：{str(e)}"
        logger.error(err_msg, exc_info=True)
        messagebox.showerror("网络异常", err_msg)
    except requests.exceptions.Timeout:
        err_msg = f"{selected_gate['name']}开闸请求超时，请检查网络连接"
        logger.error(err_msg)
        messagebox.showerror("网络超时", err_msg)
    except Exception as e:
        err_msg = f"{selected_gate['name']}开闸异常：{str(e)}"
        logger.error(err_msg, exc_info=True)
        messagebox.showerror("系统异常", err_msg)
    finally:
        session.close()

# -------------------------- 4. 优化后的GUI界面（修复空白问题）--------------------------
class DraggableFrame(tk.Frame):
    """可拖动的分组框架（改用tk.Frame，修复渲染问题）"""
    def __init__(self, parent, title, **kwargs):
        # 改用tk.Frame以便自定义样式，解决ttk.LabelFrame渲染问题
        super().__init__(parent, bd=2, relief=tk.GROOVE,** kwargs)
        self.parent = parent
        self.start_x = 0
        self.start_y = 0
        self.is_dragging = False

        # 分组标题栏（用于拖动）
        self.title_bar = tk.Frame(self, bg="#2c3e50")
        self.title_bar.pack(fill=tk.X, side=tk.TOP)
        self.title_label = tk.Label(self.title_bar, text=title, bg="#2c3e50", fg="white", padx=10, pady=5)
        self.title_label.pack(side=tk.LEFT)

        # 绑定拖动事件（仅标题栏可拖动）
        self.title_bar.bind("<Button-1>", self.on_drag_start)
        self.title_bar.bind("<B1-Motion>", self.on_drag_motion)
        self.title_bar.bind("<ButtonRelease-1>", self.on_drag_end)

        # 按钮容器，用于自动排列
        self.button_frame = tk.Frame(self, padx=10, pady=10)
        self.button_frame.pack(fill=tk.BOTH, expand=True, side=tk.BOTTOM)

        # 记录按钮位置
        self.button_count = 0
        self.columns = 3  # 每行显示3个按钮

    def on_drag_start(self, event):
        """记录拖动开始位置"""
        self.is_dragging = True
        # 获取组件相对于父容器的位置
        self.start_x = event.x
        self.start_y = event.y
        self.lift()  # 拖动时置于顶层

    def on_drag_motion(self, event):
        """处理拖动逻辑"""
        if not self.is_dragging:
            return
        # 计算新位置
        x = self.winfo_x() + (event.x - self.start_x)
        y = self.winfo_y() + (event.y - self.start_y)
        # 限制不超出父容器边界
        x = max(0, x)
        y = max(0, y)
        self.place(x=x, y=y)

    def on_drag_end(self, event):
        """结束拖动"""
        self.is_dragging = False

    def add_button(self, text, command):
        """添加按钮并自动排列"""
        row = self.button_count // self.columns
        col = self.button_count % self.columns
        
        btn = ttk.Button(
            self.button_frame, 
            text=text, 
            command=command,
            width=12
        )
        btn.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
        
        # 让网格自适应拉伸
        self.button_frame.grid_rowconfigure(row, weight=1)
        self.button_frame.grid_columnconfigure(col, weight=1)
        
        self.button_count += 1


def create_gui():
    """创建修复后的图形化操作界面"""
    root = tk.Tk()
    root.title("停车场闸门控制工具")
    root.geometry("800x600")
    root.minsize(600, 400)

    # 设置窗口背景
    root.configure(bg="#ecf0f1")

    # 添加窗口标题
    title_label = tk.Label(root, text="停车场闸门控制工具", font=("微软雅黑", 16, "bold"), bg="#ecf0f1", fg="#2c3e50")
    title_label.pack(pady=10)

    # 创建主容器（用于放置可拖动分组）
    main_container = tk.Frame(root, bg="#ecf0f1")
    main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 按账号分组处理闸门数据
    grouped_gates = defaultdict(list)
    for gate in GATE_CONFIG:
        grouped_gates[gate["account"]].append(gate)

    # 按账号字母顺序排序分组（自动排序）
    sorted_accounts = sorted(grouped_gates.keys())

    # 初始化分组的初始位置
    init_x = 20
    init_y = 20
    gap_x = 40  # 分组之间的水平间距
    gap_y = 20  # 分组之间的垂直间距
    current_x = init_x
    current_y = init_y
    max_group_width = 350  # 每个分组的宽度
    max_group_height = 0   # 记录最大分组高度

    # 创建分组框架并添加按钮
    for account in sorted_accounts:
        gates = grouped_gates[account]
        
        # 创建可拖动分组
        group_frame = DraggableFrame(main_container, f"账号: {account}", bg="white")
        # 先临时pack计算尺寸，再用place布局
        group_frame.pack_propagate(False)
        group_frame.config(width=max_group_width, height=200)  # 初始尺寸
        group_frame.place(x=current_x, y=current_y)

        # 为每个闸门添加按钮
        for gate in gates:
            group_frame.add_button(
                gate["name"],
                lambda g=gate: open_gate(g)
            )

        # 更新分组的实际高度（根据按钮数量调整）
        group_height = 60 + ( (group_frame.button_count // group_frame.columns + 1) * 40 )
        group_frame.config(height=group_height)
        max_group_height = max(max_group_height, group_height)

        # 调整下一个分组的位置（自动换行）
        current_x += max_group_width + gap_x
        # 如果超出父容器宽度，换行
        if current_x + max_group_width > root.winfo_width() - 40:
            current_x = init_x
            current_y += max_group_height + gap_y

    # 监听窗口大小变化，支持自适应
    def on_window_resize(event):
        nonlocal current_x, current_y
        current_x = init_x
        current_y = init_y
        # 重新排列所有分组
        for child in main_container.winfo_children():
            if isinstance(child, DraggableFrame):
                child.place(x=current_x, y=current_y)
                current_x += max_group_width + gap_x
                if current_x + max_group_width > event.width - 40:
                    current_x = init_x
                    current_y += max_group_height + gap_y

    root.bind("<Configure>", on_window_resize)

    root.mainloop()

# -------------------------- 程序入口 --------------------------
if __name__ == "__main__":
    create_gui()