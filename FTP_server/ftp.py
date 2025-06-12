import psutil
import socket
import os
import time
from pyftpdlib.authorizers import DummyAuthorizer
from pyftpdlib.handlers import FTPHandler
from pyftpdlib.servers import FTPServer
import argparse
import multiprocessing
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import sys
import logging
import threading

# 用于存储运行的进程和对应的IP地址
running_processes = {}
# 标记 FTP 服务器状态
ftp_running = False
# 日志文件路径，使用绝对路径
LOG_FILE = 'ftp_server.log'
# 上一次检测到的IP地址列表
last_ip_list = []
# IP监控线程
ip_monitor_thread = None
# 监控线程运行标志
monitor_running = False
# 服务器配置
server_config = {}


def get_valid_ips():
    valid_ips = []
    try:
        net_if_addrs = psutil.net_if_addrs()
        for interface, addrs in net_if_addrs.items():
            for addr in addrs:
                if addr.family == socket.AF_INET and not addr.address.startswith('169.254.'):
                    valid_ips.append(f"{interface}: {addr.address}")
    except Exception as e:
        print(f"获取网络接口信息时出错: {e}")
    return valid_ips


def get_ip_addresses():
    """获取当前所有有效IP地址列表，只返回IP地址部分"""
    return [ip_info.split(": ")[1] for ip_info in get_valid_ips()]


def start_ftp_server(ip, config):
    """为指定IP地址启动FTP服务器进程"""
    try:
        # 使用传入的配置参数
        user = config['user']
        password = config['password']
        port = config['port']
        shared_dir = config['shared_dir']
        allow_anonymous = config['allow_anonymous']
        anonymous_perm = config['anonymous_perm']
        passive_ports = config['passive_ports']

        try:
            os.chmod(shared_dir, 0o777)
        except Exception as e:
            logging.error(f"Failed to set directory permissions: {e}")
            return

        authorizer = DummyAuthorizer()
        authorizer.add_user(user, password, shared_dir, perm="elradfmw")

        if allow_anonymous:
            authorizer.add_anonymous(shared_dir, perm=anonymous_perm)

        handler = FTPHandler
        handler.authorizer = authorizer
        handler.passive_ports = range(passive_ports[0], passive_ports[1] + 1)

        address = (ip, port)
        server = FTPServer(address, handler)

        # 配置日志记录到文件和命令行
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler = logging.FileHandler(LOG_FILE)
        file_handler.setFormatter(formatter)
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(formatter)

        logger = logging.getLogger()
        logger.addHandler(file_handler)
        logger.addHandler(stream_handler)
        logger.setLevel(logging.INFO)

        logger.info(f"Starting FTP server on {ip}:{port} sharing directory {shared_dir}")
        try:
            server.serve_forever()
        except Exception as e:
            logger.error(f"FTP server error on {ip}: {e}")
    except Exception as e:
        logging.error(f"Error starting FTP server on {ip}: {e}")


def start_ftp_for_ip(ip):
    """为指定IP启动FTP服务器进程"""
    global running_processes

    if ip in running_processes and running_processes[ip].is_alive():
        return  # 该IP的服务器已在运行

    process = multiprocessing.Process(target=start_ftp_server, args=(ip, server_config))
    running_processes[ip] = process
    process.start()
    logging.info(f"Started FTP server on {ip}")


def stop_ftp_for_ip(ip):
    """停止指定IP的FTP服务器进程"""
    global running_processes

    if ip in running_processes:
        process = running_processes[ip]
        if process.is_alive():
            process.terminate()
            process.join()
        del running_processes[ip]
        logging.info(f"Stopped FTP server on {ip}")


def monitor_ip_changes():
    """监控IP地址变化的线程函数"""
    global last_ip_list, running_processes, monitor_running

    while monitor_running:
        current_ips = get_ip_addresses()

        # 检测新增的IP地址
        for ip in current_ips:
            if ip not in last_ip_list:
                logging.info(f"New IP detected: {ip}")
                if ftp_running:
                    start_ftp_for_ip(ip)

        # 检测移除的IP地址
        for ip in last_ip_list:
            if ip not in current_ips:
                logging.info(f"IP removed: {ip}")
                stop_ftp_for_ip(ip)

        last_ip_list = current_ips
        time.sleep(1)  # 每5秒检查一次


def start_all_ftp_servers(config, ui=True):
    """启动所有IP地址的FTP服务器"""
    global running_processes, ftp_running, last_ip_list, server_config, ip_monitor_thread, monitor_running

    # 保存配置供后续使用
    server_config = config

    # 获取当前IP列表
    current_ips = get_ip_addresses()
    last_ip_list = current_ips

    # 为每个IP启动FTP服务器
    for ip in current_ips:
        start_ftp_for_ip(ip)

    ftp_running = True

    # 启动IP监控线程
    if ip_monitor_thread is None or not ip_monitor_thread.is_alive():
        monitor_running = True
        ip_monitor_thread = threading.Thread(target=monitor_ip_changes)
        ip_monitor_thread.daemon = True
        ip_monitor_thread.start()

    if ui:
        status_label.config(bg="green", text="FTP 服务器已启动")
        # 更新IP地址显示
        ip_info = "\n".join(get_valid_ips())
        ip_text.delete(1.0, tk.END)
        ip_text.insert(tk.END, ip_info)


def stop_all_ftp_servers(ui=True):
    """停止所有FTP服务器"""
    global running_processes, ftp_running, monitor_running

    # 停止IP监控线程
    monitor_running = False
    if ip_monitor_thread and ip_monitor_thread.is_alive():
        ip_monitor_thread.join(timeout=1.0)

    # 停止所有FTP服务器进程
    for ip in list(running_processes.keys()):
        stop_ftp_for_ip(ip)

    running_processes = {}
    ftp_running = False

    if ui:
        status_label.config(bg="red", text="FTP 服务器已停止")
        # 更新IP地址显示
        ip_info = "\n".join(get_valid_ips())
        ip_text.delete(1.0, tk.END)
        ip_text.insert(tk.END, ip_info)


def toggle_ftp_server(button):
    """切换FTP服务器状态"""
    global ftp_running

    if ftp_running:
        stop_all_ftp_servers(True)
        button.config(text="启动 FTP 服务器")
    else:
        # 收集配置
        config = {
            'user': user_entry.get(),
            'password': password_entry.get(),
            'port': int(port_entry.get()),
            'shared_dir': shared_dir_entry.get(),
            'allow_anonymous': anonymous_var.get(),
            'anonymous_perm': 'r' if anonymous_perm_var.get() == "只读" else 'elradfmw',
            'passive_ports': (60000, 65535)
        }

        # 验证配置
        try:
            config['port'] = int(port_entry.get())
        except ValueError:
            messagebox.showerror("错误", "端口号必须是整数。")
            return

        start_all_ftp_servers(config, True)
        button.config(text="停止 FTP 服务器")


def select_shared_directory():
    """选择共享目录"""
    shared_dir = filedialog.askdirectory()
    if shared_dir:
        shared_dir_entry.delete(0, tk.END)
        shared_dir_entry.insert(0, shared_dir)


def update_log():
    """更新日志显示"""
    try:
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            log_content = f.read()
            log_text.delete(1.0, tk.END)
            log_text.insert(tk.END, log_content)
            # 将光标移动到文本末尾，实现自动翻到最新
            log_text.see(tk.END)
    except FileNotFoundError:
        pass
    root.after(1000, update_log)


def refresh_ip_info():
    """刷新IP信息显示"""
    ip_info = "\n".join(get_valid_ips())
    ip_text.delete(1.0, tk.END)
    ip_text.insert(tk.END, ip_info)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    parser = argparse.ArgumentParser(description='Start an FTP server with custom settings.')
    parser.add_argument('-u', default='xusokong', help='Username for FTP access')
    parser.add_argument('-pw', default='12345678', help='Password for FTP access')
    parser.add_argument('-p', type=int, default=21, help='Port number for the FTP server')
    parser.add_argument('-dir', default='D:/', help='Directory to be shared via FTP')
    parser.add_argument('-any', action='store_true', help='Allow anonymous access')
    parser.add_argument('-anyrw', default='r', choices=['r', 'elradfmw'],
                        help='Permissions for anonymous users, "r" for read-only, "elradfmw" for read-write')
    parser.add_argument('-pp', nargs=2, type=int, default=[60000, 65535],
                        help='Passive port range for the FTP server, e.g., 60000 65535')
    parser.add_argument('-cmd', action='store_true', help='Run in command-line mode without UI')

    args = parser.parse_args()

    if args.cmd:
        # 命令行模式，不启动 UI，直接启动 FTP 服务器
        config = {
            'user': args.u,
            'password': args.pw,
            'port': args.p,
            'shared_dir': args.dir,
            'allow_anonymous': args.any,
            'anonymous_perm': args.anyrw,
            'passive_ports': tuple(args.pp)
        }

        start_all_ftp_servers(config, False)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            stop_all_ftp_servers(False)
    else:
        # 默认模式，启动 UI
        # 创建主窗口
        root = tk.Tk()
        root.title("FTPserver")

        # 状态栏
        status_label = tk.Label(root, text="FTP 服务器未启动", bg="red", anchor=tk.W)
        status_label.pack(fill=tk.X)

        # 左侧配置面板
        left_frame = tk.Frame(root)
        left_frame.pack(side=tk.LEFT, padx=10, pady=10)

        # 显示网卡和 IP 地址
        ip_info = "\n".join(get_valid_ips())
        ip_label = tk.Label(left_frame, text="IP地址:")
        ip_label.pack()

        # 创建IP信息显示区域和刷新按钮
        ip_frame = tk.Frame(left_frame)
        ip_frame.pack(fill=tk.X)

        ip_text = tk.Text(ip_frame, height=5, width=30)
        ip_text.insert(tk.END, ip_info)
        ip_text.pack(side=tk.LEFT)

        refresh_button = tk.Button(ip_frame, text="刷新", command=refresh_ip_info)
        refresh_button.pack(side=tk.RIGHT, padx=5)

        # FTP 用户输入框
        user_label = tk.Label(left_frame, text="FTP 用户:")
        user_label.pack()
        user_entry = tk.Entry(left_frame)
        user_entry.insert(0, args.u)
        user_entry.pack()

        # FTP 密码输入框
        password_label = tk.Label(left_frame, text="FTP 密码:")
        password_label.pack()
        password_entry = tk.Entry(left_frame, show="*")
        password_entry.insert(0, args.pw)
        password_entry.pack()

        # FTP 端口输入框
        port_label = tk.Label(left_frame, text="FTP 端口:")
        port_label.pack()
        port_entry = tk.Entry(left_frame)
        port_entry.insert(0, str(args.p))
        port_entry.pack()

        # 共享地址输入框和选择按钮
        shared_dir_label = tk.Label(left_frame, text="共享地址:")
        shared_dir_label.pack()
        shared_dir_entry = tk.Entry(left_frame)
        shared_dir_entry.insert(0, args.dir)
        shared_dir_entry.pack()
        select_button = tk.Button(left_frame, text="选择共享文件夹", command=select_shared_directory)
        select_button.pack()

        # 是否允许匿名用户复选框
        anonymous_var = tk.BooleanVar()
        anonymous_var.set(args.any)
        anonymous_checkbox = tk.Checkbutton(left_frame, text="允许匿名用户", variable=anonymous_var)
        anonymous_checkbox.pack()

        # 匿名用户权限选择框
        anonymous_perm_var = tk.StringVar()
        anonymous_perm_var.set("只读" if args.anyrw == 'r' else "读写")
        anonymous_perm_label = tk.Label(left_frame, text="匿名用户权限:")
        anonymous_perm_label.pack()
        anonymous_perm_menu = tk.OptionMenu(left_frame, anonymous_perm_var, "只读", "读写")
        anonymous_perm_menu.pack()

        # 创建切换按钮
        toggle_button = tk.Button(left_frame, text="启动 FTP 服务器",
                                  command=lambda: toggle_ftp_server(toggle_button))
        toggle_button.pack(pady=20)

        # 右侧日志面板
        right_frame = tk.Frame(root)
        right_frame.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        log_label = tk.Label(right_frame, text="FTP 日志:")
        log_label.pack()
        log_text = tk.Text(right_frame, height=20, width=40)
        log_text.pack(fill=tk.BOTH, expand=True)

        # 定期更新日志
        root.after(1000, update_log)

        # 运行主循环
        root.mainloop()
