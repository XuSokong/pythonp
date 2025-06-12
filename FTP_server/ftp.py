import psutil
import socket
import os
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

# 用于存储运行的进程
running_processes = []
# 标记 FTP 服务器状态
ftp_running = False
# 日志文件路径，使用绝对路径
LOG_FILE = 'ftp_server.log'
# 默认文件编码
DEFAULT_ENCODING = 'utf-8'


def start_ftp_server(user, password, port, shared_dir, ip='0.0.0.0', allow_anonymous=False, anonymous_perm='r',
                     passive_ports=(60000, 65535), encoding=DEFAULT_ENCODING):
    try:
        os.chmod(shared_dir, 0o777)
    except Exception as e:
        logging.error(f"Failed to set directory permissions: {e}")
        messagebox.showerror("错误", f"设置共享目录权限失败: {e}")
        return

    authorizer = DummyAuthorizer()
    authorizer.add_user(user, password, shared_dir, perm="elradfmw")

    if allow_anonymous:
        authorizer.add_anonymous(shared_dir, perm=anonymous_perm)

    handler = FTPHandler
    handler.authorizer = authorizer
    handler.passive_ports = range(passive_ports[0], passive_ports[1] + 1)

    # 设置文件编码
    handler.encoding = encoding

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
    logger.info(f"Using file encoding: {encoding}")
    try:
        server.serve_forever()
    except Exception as e:
        logger.error(f"FTP server error: {e}")
        messagebox.showerror("错误", f"FTP 服务器出错: {e}")


def start_all_ftp_servers(user, password, port, shared_dir, allow_anonymous, anonymous_perm, passive_ports, encoding,
                          ui=True):
    global running_processes, ftp_running
    # 只在0.0.0.0地址启动一个FTP服务器
    process = multiprocessing.Process(target=start_ftp_server, args=(
        user, password, port, shared_dir, '0.0.0.0', allow_anonymous, anonymous_perm, passive_ports, encoding))
    running_processes.append(process)
    process.start()
    ftp_running = True
    if ui:
        status_label.config(bg="green", text="FTP 服务器已启动")


def stop_all_ftp_servers(ui=True):
    global running_processes, ftp_running
    for process in running_processes:
        if process.is_alive():
            process.terminate()
            process.join()
    running_processes = []
    ftp_running = False
    if ui:
        status_label.config(bg="red", text="FTP 服务器已停止")


def toggle_ftp_server(button):
    global ftp_running
    if ftp_running:
        stop_all_ftp_servers(True)
        button.config(text="启动 FTP 服务器")
    else:
        user = user_entry.get()
        password = password_entry.get()
        try:
            port = int(port_entry.get())
        except ValueError:
            messagebox.showerror("错误", "端口号必须是整数。")
            return
        shared_dir = shared_dir_entry.get()
        allow_anonymous = anonymous_var.get()
        # 转换匿名用户权限选项
        if anonymous_perm_var.get() == "只读":
            anonymous_perm = 'r'
        else:
            anonymous_perm = 'elradfmw'
        passive_ports = (60000, 65535)
        # 获取选择的文件编码
        encoding = encoding_var.get()
        start_all_ftp_servers(user, password, port, shared_dir, allow_anonymous, anonymous_perm, passive_ports,
                              encoding, True)
        button.config(text="停止 FTP 服务器")


def select_shared_directory():
    shared_dir = filedialog.askdirectory()
    if shared_dir:
        shared_dir_entry.delete(0, tk.END)
        shared_dir_entry.insert(0, shared_dir)


def update_log():
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
    parser.add_argument('-enc', default=DEFAULT_ENCODING, help='File encoding for FTP operations')

    args = parser.parse_args()

    if args.cmd:
        # 命令行模式，不启动 UI，直接启动 FTP 服务器
        start_all_ftp_servers(args.u, args.pw, args.p, args.dir, args.any, args.anyrw, tuple(args.pp), args.enc, False)
        try:
            while True:
                pass
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

        # 文件编码选择框
        encoding_var = tk.StringVar()
        encoding_var.set(args.enc)
        encoding_label = tk.Label(left_frame, text="文件编码:")
        encoding_label.pack()
        encoding_menu = tk.OptionMenu(left_frame, encoding_var, "utf-8", "gbk", "gb2312", "ascii", "latin-1")
        encoding_menu.pack()

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
