import socket
import psutil
import tkinter as tk
from tkinter import messagebox
from pyftpdlib.authorizers import DummyAuthorizer
from pyftpdlib.handlers import FTPHandler
from pyftpdlib.servers import FTPServer
import logging


# 自定义日志处理程序，将日志输出到 tkinter 的文本框
class TkinterLogger(logging.Handler):
    def __init__(self, text_widget):
        logging.Handler.__init__(self)
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)


def get_all_network_ips():
    network_info = []
    for interface, addrs in psutil.net_if_addrs().items():
        for addr in addrs:
            if addr.family == socket.AF_INET:
                network_info.append((interface, addr.address))
    return network_info


server = None
is_server_running = False


def toggle_ftp_server():
    global server, is_server_running
    if is_server_running:
        try:
            if server:
                server.close_all()
                server = None
                start_button.config(text="启动 FTP 服务器")
                is_server_running = False
                # messagebox.showinfo("信息", "FTP 服务器已关闭")
        except Exception as e:
            messagebox.showerror("错误", f"关闭 FTP 服务器时出错: {e}")
    else:
        try:
            server_address = address_entry.get()
            port = int(port_entry.get())
            username = user_entry.get()
            password = password_entry.get()
            share_path = share_path_entry.get()

            authorizer = DummyAuthorizer()
            authorizer.add_user(username, password, share_path, perm='elradfmwMT')
            authorizer.add_anonymous(share_path)

            handler = FTPHandler
            handler.authorizer = authorizer

            address = (server_address, port)
            server = FTPServer(address, handler)

            server.max_cons = 256
            server.max_cons_per_ip = 5

            # 配置日志记录
            logger = logging.getLogger()
            logger.setLevel(logging.INFO)
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            tkinter_logger = TkinterLogger(log_text)
            tkinter_logger.setFormatter(formatter)
            logger.addHandler(tkinter_logger)

            # 这里使用线程来启动服务器，避免阻塞主线程
            import threading
            server_thread = threading.Thread(target=server.serve_forever)
            server_thread.daemon = True
            server_thread.start()

            start_button.config(text="关闭 FTP 服务器")
            is_server_running = True
            # messagebox.showinfo("信息", "FTP 服务器已启动")
        except Exception as e:
            messagebox.showerror("错误", f"启动 FTP 服务器时出错: {e}")


root = tk.Tk()
root.title("网卡信息及 FTP 服务器")

# 创建一个框架用于放置左边的配置信息
left_frame = tk.Frame(root)
left_frame.pack(side=tk.LEFT, padx=10, pady=10)

network_info = get_all_network_ips()

info_text = tk.Text(left_frame, height=10, width=40, font=("Arial", 10))
info_text.insert(tk.END, "当前主机的网卡及其对应的 IP 地址如下：\n")
for interface, ip in network_info:
    info_text.insert(tk.END, f"网卡名称: {interface}, IP 地址: {ip}\n")
info_text.pack(pady=10)
info_text.config(state=tk.DISABLED)

tk.Label(left_frame, text="服务器地址:", font=("Arial", 10)).pack(pady=5)
address_entry = tk.Entry(left_frame, font=("Arial", 10))
address_entry.insert(0, '')
address_entry.pack(pady=5)

tk.Label(left_frame, text="端口:", font=("Arial", 10)).pack(pady=5)
port_entry = tk.Entry(left_frame, font=("Arial", 10))
port_entry.insert(0, '21')
port_entry.pack(pady=5)

tk.Label(left_frame, text="用户:", font=("Arial", 10)).pack(pady=5)
user_entry = tk.Entry(left_frame, font=("Arial", 10))
user_entry.insert(0, 'xusokong')
user_entry.pack(pady=5)

tk.Label(left_frame, text="密码:", font=("Arial", 10)).pack(pady=5)
password_entry = tk.Entry(left_frame, show='*', font=("Arial", 10))
password_entry.insert(0, '123456')
password_entry.pack(pady=5)

tk.Label(left_frame, text="共享地址:", font=("Arial", 10)).pack(pady=5)
share_path_entry = tk.Entry(left_frame, font=("Arial", 10))
share_path_entry.insert(0, 'C:/Users/Lenovo/Desktop/IRRad-20250410')
share_path_entry.pack(pady=5)

start_button = tk.Button(left_frame, text="启动 FTP 服务器", command=toggle_ftp_server, font=("Arial", 12))
start_button.pack(pady=20)

# 创建一个框架用于放置右边的日志文本框
right_frame = tk.Frame(root)
right_frame.pack(side=tk.RIGHT, padx=10, pady=10)

log_text = tk.Text(right_frame, height=30, width=60, font=("Arial", 10))
log_text.pack()

root.mainloop()
