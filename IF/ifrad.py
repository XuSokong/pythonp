import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import serial
from array import array
import serial.tools.list_ports
import threading
import time
import binascii
import logging
from datetime import datetime
import os
import json
import csv  # 新增：导入csv模块

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("rs485_tool.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class RS485Tool:
    def __init__(self, root):
        self.root = root
        self.root.title("RS-485 通信工具")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)
        
        # 设置中文字体支持
        self.font_config()
        
        # 串口状态
        self.ser = None
        self.is_connected = False
        self.receive_thread = None
        self.running = False
        self.lock = threading.Lock()  # 线程锁，确保数据操作安全
        
        # 自动工作相关变量
        self.auto_working = False
        self.auto_work_thread = None
        self.auto_work_count = 0
        self.auto_work_total = 0
        self.selected_workflow = tk.StringVar(value="流程1")
        self.workstatusflag = 0
        
        # 命令集
        self.rhcontrol = "00 00 00 00 00 00 02 00"  # 热敏电阻
        self.ptcontrol = "00 00 00 00 00 00 02 01"  # 铂电阻
        self.racontrol = "00 00 00 00 00 00 02 02"  # 热辐射
        self.alcontrol = "00 00 00 00 00 00 02 03"  # 流程1
        
        # 工作流程映射
        self.workflow_commands = {
            "热敏电阻": self.rhcontrol,
            "铂电阻": self.ptcontrol,
            "热辐射": self.racontrol,
            "流程1": self.alcontrol
        }
        
        # 数据包解析相关
        self.packet_header = [0x50, 0x52, 0x44, 0x54, 0x49, 0x52, 0x30, 0x31]  # "PRDTIR01"
        self.packet_footer = [0x24, 0x24, 0x24, 0x24]  # "$$$$"
        self.buffer = []  # 用于缓存接收的数据
        self.packet_count = 0  # 数据包计数器
        
        # 数据存储 - 新增：用于累积数据
        self.all_vodata = []
        self.all_dndata = []
        
        # 串口参数
        self.port_var = tk.StringVar()
        self.baudrate_var = tk.StringVar(value="9600")
        self.databits_var = tk.StringVar(value="8")
        self.stopbits_var = tk.StringVar(value="1")
        self.parity_var = tk.StringVar(value="N")
        
        # 日志设置
        self.save_logs_var = tk.BooleanVar(value=True)
        self.log_timestamp = datetime.now().strftime("%Y%m%d")
        self.log_path_var = tk.StringVar(value=os.getcwd())
        
        # 保存的命令集
        self.commands = self.load_commands()
        
        # 创建界面
        self.create_widgets()
        
        # 加载保存的设置
        self.load_settings()
        
        # 刷新端口列表
        self.refresh_ports()
        
        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 启动端口自动刷新定时器
        self.port_refresh_timer()
    
    def font_config(self):
        """配置字体以支持中文显示"""
        default_font = ('SimHei', 10)
        self.root.option_add("*Font", default_font)
    
    def create_widgets(self):
        """创建界面组件"""
        # 顶部按钮区域
        top_button_frame = ttk.Frame(self.root, padding="5")
        top_button_frame.pack(fill=tk.X)
        
        # 串口设置按钮
        self.serial_settings_btn = ttk.Button(top_button_frame, text="串口设置", command=self.open_serial_settings)
        self.serial_settings_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 连接/断开按钮
        self.connect_btn = ttk.Button(top_button_frame, text="连接", command=self.toggle_connection)
        self.connect_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 功能按钮
        self.thermistor_btn = ttk.Button(top_button_frame, text="热敏电阻", command=lambda: self.send_control_command(self.rhcontrol))
        self.thermistor_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        self.pt_btn = ttk.Button(top_button_frame, text="铂电阻", command=lambda: self.send_control_command(self.ptcontrol))
        self.pt_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        self.ifrad_btn = ttk.Button(top_button_frame, text="热辐射", command=lambda: self.send_control_command(self.racontrol))
        self.ifrad_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        self.collect1_btn = ttk.Button(top_button_frame, text="流程1", command=lambda: self.send_control_command(self.alcontrol))
        self.collect1_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 自动工作区域
        auto_work_frame = ttk.Frame(top_button_frame)
        auto_work_frame.pack(side=tk.LEFT, padx=10, pady=5)
        
        workflow_combo = ttk.Combobox(auto_work_frame, textvariable=self.selected_workflow, width=8, state="readonly")
        workflow_combo['values'] = list(self.workflow_commands.keys())
        workflow_combo.pack(side=tk.LEFT, padx=5)
        
        self.auto_work_count_var = tk.StringVar(value="0")
        ttk.Entry(auto_work_frame, textvariable=self.auto_work_count_var, width=5).pack(side=tk.LEFT, padx=5)
        
        self.auto_work_btn = ttk.Button(auto_work_frame, text="自动工作", command=self.toggle_auto_work)
        self.auto_work_btn.pack(side=tk.LEFT, padx=5)
        
        self.auto_work_status_var = tk.StringVar(value="未运行")
        ttk.Label(auto_work_frame, textvariable=self.auto_work_status_var).pack(side=tk.LEFT, padx=5)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建上下分栏
        upper_frame = ttk.Frame(main_frame)
        lower_frame = ttk.Frame(main_frame)
        upper_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=False, pady=(0, 10))
        lower_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        
        # 数据显示区域
        rawdata_frame = ttk.LabelFrame(upper_frame, text="数据显示", padding="10")
        rawdata_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格
        columns = [
            'index', 'time', 
            # CH1 数据
            'CH1T1', 'CH1T2', 'CH1T3', 'CH1T4', 'CH1T5', 'CH1PT', 'CH1R',
            'CH1B1', 'CH1B2', 'CH1B3', 'CH1B4', 'CH1B5',
            # CH2 数据
            'CH2T1', 'CH2T2', 'CH2T3', 'CH2T4', 'CH2T5', 'CH2PT', 'CH2R',
            'CH2B1', 'CH2B2', 'CH2B3', 'CH2B4', 'CH2B5',
            # CH3 数据
            'CH3T1', 'CH3T2', 'CH3T3', 'CH3T4', 'CH3T5', 'CH3PT', 'CH3R',
            'CH3B1', 'CH3B2', 'CH3B3', 'CH3B4', 'CH3B5',
            # CH4 数据
            'CH4T1', 'CH4T2', 'CH4T3', 'CH4T4', 'CH4T5', 'CH4PT', 'CH4R',
            'CH4B1', 'CH4B2', 'CH4B3', 'CH4B4', 'CH4B5',
            # 环境数据
            'RT', 'MT', 'RH', 'MH'
        ]
        
        self.result_data_display = ttk.Treeview(rawdata_frame, columns=columns, show='headings')
        
        # 设置列标题和宽度
        for col in columns:
            self.result_data_display.heading(col, text=col)
            width = 60 if col == 'index' else 84
            self.result_data_display.column(col, width=width, anchor=tk.CENTER)
        
        # 添加垂直滚动条
        rawdata_vscrollbar = ttk.Scrollbar(rawdata_frame, orient="vertical", command=self.result_data_display.yview)
        self.result_data_display.configure(yscrollcommand=rawdata_vscrollbar.set)
        
        # 添加水平滚动条 - 新增代码
        rawdata_hscrollbar = ttk.Scrollbar(rawdata_frame, orient="horizontal", command=self.result_data_display.xview)
        self.result_data_display.configure(xscrollcommand=rawdata_hscrollbar.set)
        
        # 布局调整
        rawdata_vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        rawdata_hscrollbar.pack(side=tk.BOTTOM, fill=tk.X)  # 水平滚动条放在底部
        self.result_data_display.pack(fill=tk.BOTH, expand=True)
        
        # 数据显示控制区
        rawdata_ctrl_frame = ttk.Frame(rawdata_frame)
        rawdata_ctrl_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.auto_clear_rawdata_var = tk.BooleanVar(value=False)
        self.auto_clear_rawdata_check = ttk.Checkbutton(rawdata_ctrl_frame, text="接收新数据前清空", variable=self.auto_clear_rawdata_var)
        self.auto_clear_rawdata_check.pack(side=tk.LEFT, padx=5)
        
        self.clear_rawdata_btn = ttk.Button(rawdata_ctrl_frame, text="清空数据", command=self.clear_data_dispaly)
        self.clear_rawdata_btn.pack(side=tk.RIGHT, padx=5)
        
        # 创建左右分栏
        left_frame = ttk.Frame(lower_frame)
        right_frame = ttk.Frame(lower_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, ipadx=5)
        
        # 左侧区域 - 发送和接收区
        # 发送区域
        send_frame = ttk.LabelFrame(left_frame, text="发送区", padding="10")
        send_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 10))
        
        self.send_text = scrolledtext.ScrolledText(send_frame, wrap=tk.WORD, height=4)
        self.send_text.pack(fill=tk.BOTH, expand=True)
        
        # 发送控制
        send_ctrl_frame = ttk.Frame(send_frame)
        send_ctrl_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.hex_send_var = tk.BooleanVar(value=True)
        self.hex_send_check = ttk.Checkbutton(send_ctrl_frame, text="十六进制发送", variable=self.hex_send_var,
                                             command=self.on_hex_send_toggle)
        self.hex_send_check.pack(side=tk.LEFT, padx=5)
        
        self.append_newline_var = tk.BooleanVar(value=False)
        self.append_newline_check = ttk.Checkbutton(send_ctrl_frame, text="自动换行", variable=self.append_newline_var)
        self.append_newline_check.pack(side=tk.LEFT, padx=5)
        self.append_newline_check.config(state=tk.DISABLED)  # 默认禁用
        
        self.send_btn = ttk.Button(send_ctrl_frame, text="发送", command=self.send_data)
        self.send_btn.pack(side=tk.RIGHT, padx=5)
        
        self.clear_send_btn = ttk.Button(send_ctrl_frame, text="清空", command=self.clear_send)
        self.clear_send_btn.pack(side=tk.RIGHT, padx=5)
        
        # 接收区域
        receive_frame = ttk.LabelFrame(left_frame, text="接收区", padding="10")
        receive_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.receive_text = scrolledtext.ScrolledText(receive_frame, wrap=tk.WORD, height=8)
        self.receive_text.pack(fill=tk.BOTH, expand=True)
        self.receive_text.config(state=tk.DISABLED)
        
        # 接收控制
        receive_ctrl_frame = ttk.Frame(receive_frame)
        receive_ctrl_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.hex_receive_var = tk.BooleanVar(value=True)
        self.hex_receive_check = ttk.Checkbutton(receive_ctrl_frame, text="十六进制显示", variable=self.hex_receive_var)
        self.hex_receive_check.pack(side=tk.LEFT, padx=5)
        
        self.timestamp_receive_var = tk.BooleanVar(value=True)
        self.timestamp_receive_check = ttk.Checkbutton(receive_ctrl_frame, text="显示时间戳", variable=self.timestamp_receive_var)
        self.timestamp_receive_check.pack(side=tk.LEFT, padx=5)
        
        self.auto_clear_var = tk.BooleanVar(value=False)
        self.auto_clear_check = ttk.Checkbutton(receive_ctrl_frame, text="接收新数据前清空", variable=self.auto_clear_var)
        self.auto_clear_check.pack(side=tk.LEFT, padx=5)
        
        self.clear_receive_btn = ttk.Button(receive_ctrl_frame, text="清空", command=self.clear_receive)
        self.clear_receive_btn.pack(side=tk.RIGHT, padx=5)
        
        # 右侧区域
        # 状态输出区域
        status_frame = ttk.LabelFrame(right_frame, text="状态输出", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 10))
        
        self.status_text = scrolledtext.ScrolledText(status_frame, wrap=tk.WORD, height=6)
        self.status_text.pack(fill=tk.BOTH, expand=True)
        self.status_text.config(state=tk.DISABLED, bg="#f0f0f0")
        
        # 状态控制
        status_ctrl_frame = ttk.Frame(status_frame)
        status_ctrl_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.auto_clear_status_var = tk.BooleanVar(value=False)
        self.auto_clear_status_check = ttk.Checkbutton(status_ctrl_frame, text="自动清空旧状态", variable=self.auto_clear_status_var)
        self.auto_clear_status_check.pack(side=tk.LEFT, padx=5)
        
        self.clear_status_btn = ttk.Button(status_ctrl_frame, text="清空状态", command=self.clear_status)
        self.clear_status_btn.pack(side=tk.RIGHT, padx=5)
        
        # 数据分析区域
        parse_frame = ttk.LabelFrame(right_frame, text="数据解析", padding="10")
        parse_frame.pack(fill=tk.BOTH, expand=True)
        
        # 解析结果显示
        self.parse_tree = ttk.Treeview(parse_frame, columns=('field', 'value', 'description'), show='headings')
        self.parse_tree.heading('field', text='字段')
        self.parse_tree.heading('value', text='值')
        self.parse_tree.heading('description', text='描述')
        
        self.parse_tree.column('field', width=100, anchor=tk.CENTER)
        self.parse_tree.column('value', width=150, anchor=tk.CENTER)
        self.parse_tree.column('description', width=200, anchor=tk.W)
        
        # 添加滚动条
        parse_scrollbar = ttk.Scrollbar(parse_frame, orient="vertical", command=self.parse_tree.yview)
        self.parse_tree.configure(yscrollcommand=parse_scrollbar.set)
        
        parse_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.parse_tree.pack(fill=tk.BOTH, expand=True)
        
        # 解析控制区
        parse_ctrl_frame = ttk.Frame(parse_frame)
        parse_ctrl_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.auto_scroll_var = tk.BooleanVar(value=True)
        self.auto_scroll_check = ttk.Checkbutton(parse_ctrl_frame, text="自动滚动到最新", variable=self.auto_scroll_var)
        self.auto_scroll_check.pack(side=tk.LEFT, padx=5)
        
        self.clear_parse_btn = ttk.Button(parse_ctrl_frame, text="清空解析结果", command=self.clear_parse_results)
        self.clear_parse_btn.pack(side=tk.RIGHT, padx=5)
        
        # 显示数据包信息
        self.packet_info_var = tk.StringVar(value="等待接收数据包...")
        ttk.Label(parse_ctrl_frame, textvariable=self.packet_info_var).pack(side=tk.LEFT, padx=5)
    
    def toggle_auto_work(self):
        """切换自动工作状态（开始/停止）"""
        if not self.is_connected:
            messagebox.showwarning("警告", "请先连接设备再开始自动工作")
            return
            
        try:
            # 获取并验证工作次数
            count_str = self.auto_work_count_var.get()
            if not count_str.isdigit():
                messagebox.showerror("输入错误", "工作次数必须是数字")
                return
                
            work_count = int(count_str)
            if work_count < 0:
                messagebox.showerror("输入错误", "工作次数不能为负数")
                return
                
        except Exception as e:
            messagebox.showerror("错误", f"参数解析错误: {str(e)}")
            return
            
        if not self.auto_working:
            # 开始自动工作
            self.auto_working = True
            self.auto_work_btn.config(text="停止")
            self.auto_work_total = work_count
            self.auto_work_count = 0
            self.auto_work_status_var.set(f" 0/{'无限' if work_count == 0 else work_count}")
            
            # 启动自动工作线程
            self.auto_work_thread = threading.Thread(target=self.auto_work_loop, daemon=True)
            self.auto_work_thread.start()
        else:
            # 停止自动工作
            self.auto_working = False
            self.workstatusflag = 0
            self.auto_work_btn.config(text="自动工作")
            self.auto_work_status_var.set(f" {self.auto_work_count}次")
    
    def auto_work_loop(self):
        """自动工作循环"""
        while self.auto_working and self.is_connected:
            try:
                # 获取当前选择的工作流程命令
                workflow = self.selected_workflow.get()
                command = self.workflow_commands.get(workflow)
                
                if self.workstatusflag == 0:
                    if command:
                        # 发送命令
                        self.send_control_command(command)
                        self.auto_work_count += 1
                        
                        # 更新状态
                        status_text = f"已完成: {self.auto_work_count}/{'无限' if self.auto_work_total == 0 else self.auto_work_total}"
                        self.root.after(0, lambda: self.auto_work_status_var.set(status_text))
                        self.log_message(f"自动工作: 发送 {workflow} 命令，第 {self.auto_work_count} 次")
                        
                        # 检查是否达到工作次数（0表示无限循环）
                        if self.auto_work_total > 0 and self.auto_work_count >= self.auto_work_total:
                            self.root.after(0, self.toggle_auto_work)
                            break
                # 等待一段时间再进行下一次操作（可根据需要调整）
                time.sleep(2)
                
            except Exception as e:
                error_msg = f"自动工作错误: {str(e)}"
                self.update_status(error_msg)
                logging.error(error_msg)
                # 发生错误时停止自动工作
                self.root.after(0, self.toggle_auto_work)
                self.workstatusflag = 0
                break
    
    def on_hex_send_toggle(self):
        """当十六进制发送选项变化时更新自动换行选项状态"""
        if self.hex_send_var.get():
            self.append_newline_var.set(False)
            self.append_newline_check.config(state=tk.DISABLED)
        else:
            self.append_newline_check.config(state=tk.NORMAL)
    
    def open_serial_settings(self):
        """打开串口设置对话框"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("串口详细设置")
        settings_window.geometry("500x400")
        settings_window.resizable(False, False)
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        # 创建设置框架
        settings_frame = ttk.Frame(settings_window, padding="20")
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # 端口选择
        ttk.Label(settings_frame, text="端口:").grid(row=0, column=0, padx=5, pady=10, sticky=tk.W)
        port_var = tk.StringVar(value=self.port_var.get())
        port_combo = ttk.Combobox(settings_frame, textvariable=port_var, width=20)
        port_combo['values'] = self.get_port_list()
        port_combo.grid(row=0, column=1, padx=5, pady=10)
        
        # 波特率选择
        ttk.Label(settings_frame, text="波特率:").grid(row=1, column=0, padx=5, pady=10, sticky=tk.W)
        baudrate_var = tk.StringVar(value=self.baudrate_var.get())
        baudrate_combo = ttk.Combobox(settings_frame, textvariable=baudrate_var, width=20)
        baudrate_combo['values'] = ['1200', '2400', '4800', '9600', '19200', '38400', '57600', '115200']
        baudrate_combo.grid(row=1, column=1, padx=5, pady=10)
        
        # 数据位选择
        ttk.Label(settings_frame, text="数据位:").grid(row=2, column=0, padx=5, pady=10, sticky=tk.W)
        databits_var = tk.StringVar(value=self.databits_var.get())
        databits_combo = ttk.Combobox(settings_frame, textvariable=databits_var, width=20)
        databits_combo['values'] = ['5', '6', '7', '8']
        databits_combo.grid(row=2, column=1, padx=5, pady=10)
        
        # 停止位选择
        ttk.Label(settings_frame, text="停止位:").grid(row=3, column=0, padx=5, pady=10, sticky=tk.W)
        stopbits_var = tk.StringVar(value=self.stopbits_var.get())
        stopbits_combo = ttk.Combobox(settings_frame, textvariable=stopbits_var, width=20)
        stopbits_combo['values'] = ['1', '1.5', '2']
        stopbits_combo.grid(row=3, column=1, padx=5, pady=10)
        
        # 校验位选择
        ttk.Label(settings_frame, text="校验位:").grid(row=4, column=0, padx=5, pady=10, sticky=tk.W)
        parity_var = tk.StringVar(value=self.parity_var.get())
        parity_combo = ttk.Combobox(settings_frame, textvariable=parity_var, width=20)
        parity_combo['values'] = ['N', 'O', 'E', 'M', 'S']
        parity_combo.grid(row=4, column=1, padx=5, pady=10)
        
        # 本地路径设置
        ttk.Label(settings_frame, text="本地路径:").grid(row=5, column=0, padx=5, pady=10, sticky=tk.W)
        path_entry = ttk.Entry(settings_frame, textvariable=self.log_path_var, width=30)
        path_entry.grid(row=5, column=1, padx=5, pady=10, sticky=tk.W)
        
        # 浏览按钮
        browse_btn = ttk.Button(settings_frame, text="...", 
                              command=lambda: self.browse_log_path(path_entry))
        browse_btn.grid(row=5, column=2, padx=5, pady=10)
        
        # 日志保存选项
        ttk.Label(settings_frame, text="日志设置:").grid(row=6, column=0, padx=5, pady=10, sticky=tk.W)
        log_check = ttk.Checkbutton(settings_frame, text="自动保存日志", variable=self.save_logs_var)
        log_check.grid(row=6, column=1, padx=5, pady=10, sticky=tk.W)
        
        # 刷新按钮
        refresh_btn = ttk.Button(settings_frame, text="刷新端口", 
                                command=lambda: self.refresh_settings_ports(port_combo))
        refresh_btn.grid(row=0, column=2, padx=5, pady=10)
        
        # 按钮区域
        btn_frame = ttk.Frame(settings_window, padding="10")
        btn_frame.pack(fill=tk.X)
        
        # 确定按钮
        def apply_settings():
            # 更新主窗口的串口设置
            self.port_var.set(port_var.get())
            self.baudrate_var.set(baudrate_var.get())
            self.databits_var.set(databits_var.get())
            self.stopbits_var.set(stopbits_var.get())
            self.parity_var.set(parity_var.get())
            
            # 确保data文件夹存在
            self.ensure_data_folder_exists()
            
            # 如果启用了日志且是首次设置，更新时间戳
            if self.save_logs_var.get():
                self.log_timestamp = datetime.now().strftime("%Y%m%d")
                self.update_status(f"日志保存已启用，文件前缀: {self.log_timestamp}")
                self.update_status(f"日志保存路径: {os.path.join(self.log_path_var.get(), 'data')}")
            
            # 保存设置
            self.save_settings()
            
            # 如果已连接，提示需要重新连接
            if self.is_connected:
                messagebox.showinfo("提示", "串口设置已更新，需要重新连接才能生效")
            
            settings_window.destroy()
        
        ok_btn = ttk.Button(btn_frame, text="确定", command=apply_settings)
        ok_btn.pack(side=tk.RIGHT, padx=10)
        
        # 取消按钮
        cancel_btn = ttk.Button(btn_frame, text="取消", command=settings_window.destroy)
        cancel_btn.pack(side=tk.RIGHT)
    
    def browse_log_path(self, entry_widget):
        """浏览选择日志保存路径"""
        current_path = self.log_path_var.get()
        if not current_path or not os.path.exists(current_path):
            current_path = os.getcwd()
            
        selected_path = filedialog.askdirectory(title="选择日志保存路径", initialdir=current_path)
        if selected_path:
            self.log_path_var.set(selected_path)
    
    def ensure_data_folder_exists(self):
        """确保data文件夹存在，如果不存在则创建"""
        try:
            data_path = os.path.join(self.log_path_var.get(), 'data')
            if not os.path.exists(data_path):
                os.makedirs(data_path)
                self.update_status(f"已创建data文件夹: {data_path}")
            return data_path
        except Exception as e:
            error_msg = f"创建data文件夹失败: {str(e)}"
            logging.error(error_msg)
            self.update_status(error_msg)
            messagebox.showerror("错误", error_msg)
            return None
    
    def save_to_log(self, log_type, content):
        """保存内容到指定类型的日志文件，保存在data文件夹下"""
        if not self.save_logs_var.get():
            return
            
        try:
            # 确保data文件夹存在
            data_path = self.ensure_data_folder_exists()
            if not data_path:
                return
                
            # 根据日志类型确定文件名
            if log_type == "receive":
                filename = os.path.join(data_path, f"RS_{self.log_timestamp}_receive.log")
            elif log_type == "status":
                filename = os.path.join(data_path, f"RS_{self.log_timestamp}_status.log")
            elif log_type == "analysis":
                filename = os.path.join(data_path, f"RS_{self.log_timestamp}_analysis.log")
            else:
                return
                
            # 添加时间戳
            timestamp = datetime.now().strftime("%Y-%m-%d")
            log_entry = f"[{timestamp}] {content}\n"
            
            # 写入文件
            with open(filename, 'a', encoding='utf-8') as f:
                f.write(log_entry)
                
        except Exception as e:
            error_msg = f"日志保存失败: {str(e)}"
            logging.error(error_msg)
    
    def get_port_list(self):
        """获取端口列表"""
        try:
            ports = serial.tools.list_ports.comports()
            return [port.device for port in ports]
        except Exception as e:
            logging.error(f"获取端口列表失败: {str(e)}")
            return []
    
    def refresh_settings_ports(self, port_combo):
        """刷新设置窗口中的端口列表"""
        port_list = self.get_port_list()
        port_combo['values'] = port_list
        if port_list and not port_combo.get():
            port_combo.current(0)
    
    def port_refresh_timer(self):
        """定时刷新端口列表"""
        if not self.is_connected:  # 仅在未连接状态下刷新
            self.refresh_ports()
        self.root.after(5000, self.port_refresh_timer)  # 每5秒刷新一次
    
    def refresh_ports(self):
        """刷新可用串口列表"""
        port_list = self.get_port_list()
        if port_list and not self.port_var.get():
            self.port_var.set(port_list[0])
    
    def toggle_connection(self):
        """切换连接状态（连接/断开）"""
        if not self.is_connected:
            self.connect()
        else:
            self.disconnect()
    
    def connect(self):
        """连接到RS-485设备"""
        try:
            # 获取串口参数
            port = self.port_var.get()
            if not port:
                messagebox.showwarning("警告", "请先在串口设置中选择端口")
                return
                
            baudrate = int(self.baudrate_var.get())
            databits = int(self.databits_var.get())
            stopbits = float(self.stopbits_var.get())
            parity = self.parity_var.get()
            
            # 转换校验位格式
            parity_map = {'N': serial.PARITY_NONE, 'O': serial.PARITY_ODD, 
                         'E': serial.PARITY_EVEN, 'M': serial.PARITY_MARK, 
                         'S': serial.PARITY_SPACE}
            parity = parity_map[parity]
            
            # 转换停止位格式
            stopbits_map = {1: serial.STOPBITS_ONE, 1.5: serial.STOPBITS_ONE_POINT_FIVE, 2: serial.STOPBITS_TWO}
            stopbits = stopbits_map[stopbits]
            
            # 打开串口
            self.ser = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=databits,
                parity=parity,
                stopbits=stopbits,
                timeout=0.1
            )
            
            if self.ser.is_open:
                self.is_connected = True
                self.connect_btn.config(text="断开")
                connect_msg = f"已连接到 {port}，波特率 {baudrate}"
                self.log_message(connect_msg)
                logging.info(connect_msg)
                
                # 如果启用了日志，记录连接信息
                self.save_to_log("status", connect_msg)
                
                # 启动接收线程
                self.running = True
                self.receive_thread = threading.Thread(target=self.receive_data, daemon=True)
                self.receive_thread.start()
            else:
                messagebox.showerror("错误", "无法打开串口")
                
        except Exception as e:
            error_msg = f"连接失败: {str(e)}"
            messagebox.showerror("连接错误", error_msg)
            logging.error(error_msg)
    
    def disconnect(self):
        """断开与RS-485设备的连接"""
        if self.is_connected and self.ser:
            # 停止自动工作
            if self.auto_working:
                self.toggle_auto_work()
                
            self.running = False
            time.sleep(0.2)  # 等待接收线程结束
            self.ser.close()
            self.is_connected = False
            self.connect_btn.config(text="连接")
            disconnect_msg = "已断开连接"
            self.log_message(disconnect_msg)
            logging.info(disconnect_msg)
            
            # 如果启用了日志，记录断开信息
            self.save_to_log("status", disconnect_msg)
    
    def receive_data(self):
        """接收数据的线程函数"""
        while self.running and self.ser and self.ser.is_open:
            try:
                if self.ser.in_waiting:
                    data = self.ser.read(self.ser.in_waiting)
                    self.display_received_data(data)
                    
                    # 将数据转换为字节列表，用于解析
                    with self.lock:  # 使用线程锁确保数据安全
                        byte_data = list(data)
                        self.buffer.extend(byte_data)
                        self.parse_packets()  # 尝试解析数据包
                time.sleep(0.01)
            except Exception as e:
                error_msg = f"接收错误: {str(e)}"
                self.log_message(error_msg)
                self.update_status(error_msg)
                self.save_to_log("status", error_msg)
                logging.error(error_msg)
                break
    
    def parse_packets(self):
        """解析缓冲区中的数据包，按照协议查找包头和包尾"""
        header_len = len(self.packet_header)
        footer_len = len(self.packet_footer)
        
        while len(self.buffer) >= header_len + footer_len + 4:  # 额外4字节用于校验和
            # 查找包头
            header_index = -1
            for i in range(len(self.buffer) - header_len + 1):
                if self.buffer[i:i+header_len] == self.packet_header:
                    header_index = i
                    break
            
            if header_index == -1:
                # 未找到包头，清除部分缓冲区
                if len(self.buffer) > header_len:
                    self.buffer = self.buffer[-(header_len-1):]
                break
            
            # 从包头之后查找包尾
            remaining_buffer = self.buffer[header_index+header_len:]
            footer_index = -1
            for i in range(len(remaining_buffer) - footer_len + 1):
                if remaining_buffer[i:i+footer_len] == self.packet_footer:
                    footer_index = i
                    break
            
            if footer_index == -1:
                # 未找到包尾，保留当前包头位置开始的缓冲区
                self.buffer = self.buffer[header_index:]
                break
            
            # 提取完整数据包（包含包头、数据内容、包尾和校验和）
            # 包尾后还有4字节校验和，所以需要额外+4
            packet_end = header_index + header_len + footer_index + footer_len + 4
            if packet_end > len(self.buffer):
                # 数据包不完整，等待更多数据
                self.buffer = self.buffer[header_index:]
                break
                
            packet = self.buffer[header_index:packet_end]
            # 移除已解析的数据包
            self.buffer = self.buffer[packet_end:]
            
            # 解析数据包内容
            self.parse_packet_content(packet)
    
    def calculate_checksum(self, data):
        """计算数据的校验和（简单求和取低4字节）"""
        checksum = sum(data) & 0xFFFFFFFF  # 取低4字节
        # 转换为4字节列表（大端模式）
        return [(checksum >> 24) & 0xFF,
                (checksum >> 16) & 0xFF,
                (checksum >> 8) & 0xFF,
                checksum & 0xFF]
    
    def parse_packet_content(self, packet):
        """解析数据包内容并显示，按照协议解析"""
        # 数据包计数加1
        self.packet_count += 1
        
        # 提取包头
        header_len = len(self.packet_header)
        header = packet[:header_len]
        
        # 提取包尾
        footer_len = len(self.packet_footer)
        footer_start = -footer_len - 4  # 包尾前4字节是校验和
        footer = packet[footer_start:-4]
        
        # 提取校验和（包尾后4字节）
        checksum = packet[-4:]
        
        # 提取数据内容（包头后到包尾前）
        content = packet[header_len:footer_start]
        
        # 数据包总长度
        packet_length = len(packet)
        
        # 更新数据包信息
        timestamp = time.strftime("%H:%M:%S")
        self.packet_info_var.set(f"最后解析时间: {timestamp} | 总接收包数: {self.packet_count} | 最后包长度: {packet_length} 字节")
        
        # 添加数据包分隔线和标识
        self.parse_tree.insert('', tk.END, values=('-'*20, '-'*20, '-'*20))
        self.parse_tree.insert('', tk.END, values=(f'数据包 #{self.packet_count}', f'时间: {timestamp}', '完整数据包解析'))
        self.parse_tree.insert('', tk.END, values=('-'*20, '-'*20, '-'*20))
        
        # 添加基本信息
        self.parse_tree.insert('', tk.END, values=('包头', ' '.join(f'{b:02X}' for b in header), '数据包起始标识 "PRDTIR01"'))
        self.parse_tree.insert('', tk.END, values=('包尾', ' '.join(f'{b:02X}' for b in footer), '数据包结束标识 "$$$$"'))
        self.parse_tree.insert('', tk.END, values=('包长度', f'{packet_length} 字节', '包括包头、数据、包尾和校验和'))
        
        # 解析数据内容
        self.parse_specific_content(content, checksum)
        
        # 自动滚动到最后一行
        if self.auto_scroll_var.get() and self.parse_tree.get_children():
            last_item = self.parse_tree.get_children()[-1]
            self.parse_tree.see(last_item)
            
        # 如果启用了日志，保存解析信息
        if self.save_logs_var.get():
            analysis_log = f"数据包 #{self.packet_count} - 长度: {packet_length} 字节 - 时间: {timestamp}"
            self.save_to_log("analysis", analysis_log)
    
    def parse_specific_content(self, content, received_checksum):
        """解析具体数据内容"""
        self.parse_tree.insert('', tk.END, values=('', '', ''))  # 空行分隔
        self.parse_tree.insert('', tk.END, values=('解析数据', '', '根据协议解析的具体字段'))
        
        offset = 0
        content_len = len(content)
        
        # 1. 解析有用数据长度 (数据头后1-2字节)
        if content_len >= offset + 2:
            data_length = (content[offset] << 8) | content[offset + 1]
            self.parse_tree.insert('', tk.END, values=(f'有用数据长度 (0x{offset:02X}-0x{offset+1:02X})', 
                                                      f'{data_length} 字节', 
                                                      '数据头后两位表示的有用数据长度'))
            offset += 2
        else:
            self.parse_tree.insert('', tk.END, values=('数据长度', '解析错误', '数据包长度不足'))
            return
        
        # 2. 解析设备号 (数据头后3字节)
        if content_len >= offset + 1:  # 设备号是1字节
            device_id = content[offset]
            self.parse_tree.insert('', tk.END, values=(f'设备号 (0x{offset:02X})', 
                                                      f'0x{device_id:02X} ({device_id})', 
                                                      '数据头后第三位表示的设备标识'))
            offset += 1
        else:
            self.parse_tree.insert('', tk.END, values=('设备号', '解析错误', '数据包长度不足'))
            return
        
        # 3. 解析数据解析方式 (数据头后4字节)
        if content_len >= offset + 1:  # 解析方式是1字节
            parse_mode = content[offset]
            self.parse_tree.insert('', tk.END, values=(f'解析方式 (0x{offset:02X})', 
                                                      f'0x{parse_mode:02X} ({parse_mode})', 
                                                      '数据头后第四位表示的解析方式'))
            offset += 1
        else:
            self.parse_tree.insert('', tk.END, values=('解析方式', '解析错误', '数据包长度不足'))
            return
        
        # 4. 提取有用数据 (从数据头后第5位开始)
        self.parse_tree.insert('', tk.END, values=('', '', ''))  # 空行分隔
        self.parse_tree.insert('', tk.END, values=('有用数据', '', ''))
        
        # 检查是否有足够的有用数据
        if content_len < offset + data_length:
            self.parse_tree.insert('', tk.END, values=('数据完整性', '不完整', f'实际长度: {content_len - offset} 字节, 预期: {data_length} 字节'))
            useful_data = content[offset:]
        else:
            useful_data = content[offset:offset + data_length]
        
        # 显示有用数据
        useful_data_str = ' '.join(f'{b:02X}' for b in useful_data)
        self.parse_tree.insert('', tk.END, values=('原始数据', useful_data_str, f'共 {len(useful_data)} 字节'))
        
        # 如果启用了日志，保存原始数据
        if self.save_logs_var.get():
            self.save_to_log("analysis", f"原始数据: {useful_data_str}")
        
        # 5. 校验和验证
        self.parse_tree.insert('', tk.END, values=('', '', ''))  # 空行分隔
        self.parse_tree.insert('', tk.END, values=('校验和验证', '', ''))
        
        # 计算有用数据的校验和
        calculated_checksum = self.calculate_checksum(useful_data)
        
        # 显示接收的校验和和计算的校验和
        received_checksum_str = ' '.join(f'{b:02X}' for b in received_checksum)
        calculated_checksum_str = ' '.join(f'{b:02X}' for b in calculated_checksum)
        
        self.parse_tree.insert('', tk.END, values=('接收校验和', received_checksum_str, '数据尾后的4字节校验和'))
        self.parse_tree.insert('', tk.END, values=('计算校验和', calculated_checksum_str, '根据有用数据计算的校验和'))
        
        # 验证校验和
        checksum_valid = (received_checksum == calculated_checksum)
        self.parse_tree.insert('', tk.END, values=('校验结果', '有效' if checksum_valid else '无效', '校验和匹配则数据有效'))
        
        # 如果启用了日志，保存校验和信息
        if self.save_logs_var.get():
            self.save_to_log("analysis", f"校验和 - 接收: {received_checksum_str}, 计算: {calculated_checksum_str}, 结果: {'有效' if checksum_valid else '无效'}")
        
        # 6. 根据解析方式解析数据（仅当校验和有效时）
        self.parse_tree.insert('', tk.END, values=('', '', ''))  # 空行分隔
        self.parse_tree.insert('', tk.END, values=('数据解析', '', '根据解析方式解析的具体含义'))
        
        if not checksum_valid:
            self.parse_tree.insert('', tk.END, values=('解析提示', '', '校验和无效，不进行数据解析'))
            self.update_status("校验和无效，不进行数据解析")
            return
        
        # 准备状态信息列表
        status_messages = []
        
        # 根据解析方式解析数据
        if parse_mode == 0x00:
            # 解析方式0x00: 各种设备的命令状态
            hex_str = ' '.join(f'{b:02X}' for b in useful_data)
            
            # 匹配命令
            status_message = ""
            if hex_str == "00 02 00 00 00 00 02 00":
                status_message = "热敏电阻开始采集"
            elif hex_str == "00 02 01 00 00 00 02 00":
                status_message = "热敏电阻采集完成"
                self.workstatusflag = 0
            elif hex_str == "00 02 00 00 00 00 02 01":
                status_message = "铂电阻开始采集"
            elif hex_str == "00 02 01 00 00 00 02 01":
                status_message = "铂电阻采集完成"
                self.workstatusflag = 0
            elif hex_str == "00 02 00 00 00 00 02 02":
                status_message = "热辐射开始采集"
            elif hex_str == "00 02 01 00 00 00 02 02":
                status_message = "热辐射采集完成"
                self.workstatusflag = 0
            elif hex_str == "00 02 00 00 00 00 02 03":
                status_message = "流程1开始采集"
            elif hex_str == "00 02 01 00 00 00 02 03":
                status_message = "流程1采集完成"
                self.workstatusflag = 0
            else:
                status_message = f"未知的命令: {hex_str}"
                
            # 在解析树中显示
            self.parse_tree.insert('', tk.END, values=('命令解析', hex_str, status_message))
            status_messages.append(status_message)
            
            # 如果启用了日志，保存解析结果
            if self.save_logs_var.get():
                self.save_to_log("analysis", f"命令解析: {hex_str} - {status_message}")
        
        # 解析方式01和02的处理
        elif parse_mode in (0x01, 0x02):
            self.parse_tree.insert('', tk.END, values=('解析方式说明', '', 'LTC2413'))
            
            # 如果勾选了自动清空，则先清空数据区域
            if self.auto_clear_rawdata_var.get():
                self.clear_data_dispaly()
            
            # 检查数据长度是否为4的倍数
            if len(useful_data) % 4 != 0:
                self.parse_tree.insert('', tk.END, values=('解析警告', '', f'数据长度为{len(useful_data)}字节，不是4的倍数，无法完全解析'))
                status_messages.append(f"解析警告: 数据长度为{len(useful_data)}字节，不是4的倍数，无法完全解析")
                
                # 如果启用了日志，保存警告
                if self.save_logs_var.get():
                    self.save_to_log("analysis", f"解析警告: 数据长度为{len(useful_data)}字节，不是4的倍数，无法完全解析")
            
            # 按4个字节一组进行处理
            for i in range(0, len(useful_data), 4):
                # 检查是否有足够的字节
                if i + 3 >= len(useful_data):
                    # 处理不完整的组
                    byte_str = ' '.join([f'0x{b:02X}' for b in useful_data[i:]])
                    self.parse_tree.insert('', tk.END, values=(f'数据组 #{i//4 + 1}', byte_str, '字节不足，无法拼接为32位数据'))
                    # 在专用区域显示
                    status_messages.append(f"数据组 #{i//4 + 1}: 字节不足，无法解析")
                    continue
                
                combined, extracted_value, calculated_value, result = self.data_analy2(useful_data[i], useful_data[i + 1], useful_data[i + 2], useful_data[i + 3])
                
                # 显示详细解析过程
                byte_str = ' '.join([f'0x{b:02X}' for b in useful_data[i:i+4]])
                self.parse_tree.insert('', tk.END, values=(f'数据组 #{i//4 + 1}', byte_str, f'32位整数: {combined}'))
                
                # 只在不是特殊值的情况下显示位提取信息
                if combined not in (0x30000000, 0x20000000):
                    self.parse_tree.insert('', tk.END, values=('', str(extracted_value), '提取位的十进制值'))
                    self.parse_tree.insert('', tk.END, values=('', f'{calculated_value:.6f}', f'计算结果'))
                    status_messages.append(f"{hex(combined)[2:].upper()} {extracted_value} {calculated_value:.6f} {result}")
                else:
                    self.parse_tree.insert('', tk.END, values=('', '', calculated_value))
                    status_messages.append(f"{hex(combined)[2:].upper()} {extracted_value} {calculated_value:.6f} {result}")
                
                
        # 解析方式03的处理 - 优化版本
        elif parse_mode == 0x03:
            # 初始化数据数组，使用float类型
            dndata = [0.0 for _ in range(54)]  # 46个数据点
            vodata = [0.0 for _ in range(54)]  # 46个数据点
            self.parse_tree.insert('', tk.END, values=('解析方式说明', '', '测试流程1'))
            
            # 如果勾选了自动清空，则先清空数据区域
            if self.auto_clear_rawdata_var.get():
                self.clear_data_dispaly()
            group_index=0
            datapar=0
            ratimes=1;pttimes=1;rhtime=1
            
            for i in range(0,4):
                for j in range(0,5*ratimes):  # CHT1~CHT5 总共5次4位数据
                    combined, dndata[group_index], vodata[group_index], result = self.data_analy2(
                        useful_data[datapar], useful_data[datapar + 1], 
                        useful_data[datapar + 2], useful_data[datapar + 3]
                    )
                    group_index += 1
                    datapar += 4
                    
                for j in range(0,pttimes):  # CHPT 总共1次4位数据
                    combined, dndata[group_index], vodata[group_index], result = self.data_analy2(
                        useful_data[datapar], useful_data[datapar + 1], 
                        useful_data[datapar + 2], useful_data[datapar + 3]
                    )
                    group_index += 1
                    datapar += 4

                for j in range(0,rhtime):  # CHRH 总共1次2位数据
                    combined, dndata[group_index], vodata[group_index], result = self.data_analy1(
                        useful_data[datapar], 
                        useful_data[datapar+ 1]
                    )
                    group_index += 1
                    datapar += 2

                
                # CHB1~CHB5 总共5次4位数据
                for j in range(0,5*ratimes):  
                    combined, dndata[group_index], vodata[group_index], result = self.data_analy2(
                        useful_data[datapar], useful_data[datapar + 1], 
                        useful_data[datapar + 2], useful_data[datapar + 3]
                    )
                    group_index += 1
                    datapar += 4
            
            print(group_index,datapar) 
            vodata[group_index+1] = int(useful_data[datapar],16)
            vodata[group_index+2] = int(useful_data[datapar+1],16)
            
            # 在专用区域显示vodata数据
            timestamp = time.strftime("%H:%M:%S")
            row_data = [f'{self.packet_count}', timestamp]
            
            # 映射vodata到表格列（根据实际需求调整映射关系）
            # 这里简单映射前44个数据到表格的各个字段
            column_mapping = [
                # CH1 数据 (索引0-11)
                0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11,
                # CH2 数据 (索引12-23)
                12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23,
                # CH3 数据 (索引24-35)
                24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35,
                # CH4 数据 (索引36-47)
                36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47,
                # T H
                48, 49,
            ]
            
            # 填充表格数据
            for i, data_idx in enumerate(column_mapping):
                if data_idx < len(vodata):
                    row_data.append(f'{vodata[data_idx]:.6f}' if isinstance(vodata[data_idx], float) else str(vodata[data_idx]))
                else:
                    row_data.append('N/A')
            
            self.result_data_display.insert('', tk.END, values=tuple(row_data))
            
            # 自动滚动到最新数据
            if self.result_data_display.get_children():
                last_item = self.result_data_display.get_children()[-1]
                self.result_data_display.see(last_item)
            
            # 在parse_specific_content方法中调用时修改为:
            # 准备要保存的数据，包含数据包索引和时间戳
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            save_vodata = [self.packet_count, timestamp] + vodata
            save_dndata = [self.packet_count, timestamp] + dndata
            
            self.save_data_csv(save_vodata, 'vodata')
            self.save_data_csv(save_dndata, 'dndata')
            status_messages.append(f"共 {len(useful_data)} 字节")
                
        # 解析方式04, 05, 07的处理
        elif parse_mode in (0x07, 0x04, 0x05):
            self.parse_tree.insert('', tk.END, values=('解析方式说明', '', '12bit_ADC计算'))
            
            # 如果勾选了自动清空，则先清空数据区域
            if self.auto_clear_rawdata_var.get():
                self.clear_data_dispaly()
            
            # 检查数据长度是否为偶数
            if len(useful_data) % 2 != 0:
                self.parse_tree.insert('', tk.END, values=('解析警告', '', f'数据长度为{len(useful_data)}字节，不是偶数，无法完全解析'))
            
            # 按两个字节一组进行处理
            for i in range(0, len(useful_data), 2):
                if i + 1 >= len(useful_data):
                    # 处理最后一个单独的字节
                    byte_str = f'0x{useful_data[i]:02X}'
                    self.parse_tree.insert('', tk.END, values=(f'数据组 #{i//2 + 1}', byte_str, '只有一个字节，无法拼接'))
                    # 在专用区域显示
                    status_messages.append(f"数据组 #{i//2 + 1}: 字节不足，无法解析")
                    continue

                # 使用数据解析方法计算结果
                combined1, combined2, result1, result2 = self.data_analy1(useful_data[i], useful_data[i + 1])
                
                # 显示详细解析过程
                byte_str = f'0x{useful_data[i]:02X} 0x{useful_data[i+1]:02X}'
                self.parse_tree.insert('', tk.END, values=(f'数据组 #{i//2 + 1}', byte_str, f'拼接值: 0x{combined1:04X} ({combined1})'))
                self.parse_tree.insert('', tk.END, values=('', f'{result1:.6f}', f'计算结果: {combined1} / 4096 × 3.258 = {result1:.6f}'))
                status_messages.append(f"{hex(combined1)[2:].upper()} {combined2} {result1:.6f} {result2:.6f}")          
                
        else:
            # 未知解析方式
            self.parse_tree.insert('', tk.END, values=('解析方式说明', '', f'未知解析方式: 0x{parse_mode:02X}'))
            self.parse_tree.insert('', tk.END, values=('原始数据', ' '.join(f'{b:02X}' for b in useful_data), ''))
            status_messages.append(f"收到未知解析方式: 0x{parse_mode:02X}")
        
        # 将解析结果添加到状态栏
        for msg in status_messages:
            self.update_status(msg)
    
    def data_analy1(self, high_byte, low_byte):
        """计算12位ADC数据，公式：(高字节<<8 | 低字节) / 4096 * 3.258"""
        combined = (high_byte << 8) | low_byte
        calculated_value = 3.258 * combined / 4096
        return combined, combined, calculated_value, calculated_value
    
    def rh_data_analy(self, high_byte, low_byte):
        """计算温湿度数据"""
        combined = (high_byte << 8) | low_byte
        calculated_value = 3.258 * combined / 4096
        return combined, combined, calculated_value, calculated_value
    
    def data_analy2(self, useful_data1, useful_data2, useful_data3, useful_data4):
        """计算LTC2413数据"""
        combined = (useful_data1 << 24) | (useful_data2 << 16) | (useful_data3 << 8) | useful_data4       
        if combined == 0x30000000:
            calculated_value = "超出上限"
            extracted_value = "N/A"
            result = "超出测量范围"
        elif combined == 0x20000000:
            calculated_value = "超出下限"
            extracted_value = "N/A"
            result = "超出测量范围"
        else:
            # 转换为32位二进制字符串
            binary_str = f"{combined:032b}"
            
            # 提取第4到27位
            extracted_bits = binary_str[3:27]
            extracted_value = int(extracted_bits, 2)
            
            # 计算结果
            if binary_str[2] == '1':
                calculated_value = extracted_value * 5 / 16777216
            else:
                calculated_value = 5 - extracted_value * 5 / 16777216
            result = f"{calculated_value:.6f}"
        return combined, extracted_value, calculated_value, result
    
    def display_received_data(self, data):
        """显示接收到的数据"""
        self.receive_text.config(state=tk.NORMAL)
        
        # 如果勾选了自动清空，则先清空
        if self.auto_clear_var.get():
            self.receive_text.delete(1.0, tk.END)
        
        # 添加时间戳
        if self.timestamp_receive_var.get():
            timestamp = time.strftime("%H:%M:%S")
            self.receive_text.insert(tk.END, f"[{timestamp}] ")
        
        # 以十六进制或ASCII显示
        log_content = ""
        if self.hex_receive_var.get():
            hex_str = binascii.hexlify(data).decode('utf-8')
            # 每两个字符添加一个空格，方便阅读
            hex_str = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
            self.receive_text.insert(tk.END, hex_str + "\n")
            log_content = hex_str
        else:
            try:
                # 尝试以ASCII解码
                text = data.decode('utf-8', errors='replace')
                self.receive_text.insert(tk.END, text)
                log_content = text
            except:
                # 解码失败则显示十六进制
                hex_str = binascii.hexlify(data).decode('utf-8')
                hex_str = ' '.join([hex_str[i:i+2] for i in range(0, len(hex_str), 2)])
                self.receive_text.insert(tk.END, f"(二进制数据: {hex_str})\n")
                log_content = f"(二进制数据: {hex_str})"
        
        # 如果启用了日志，保存接收数据
        if self.save_logs_var.get():
            self.save_to_log("receive", log_content)
        
        self.receive_text.see(tk.END)
        self.receive_text.config(state=tk.DISABLED)
    
    def update_status(self, message):
        """更新状态输出区域的内容"""
        self.status_text.config(state=tk.NORMAL)
        
        # 如果勾选了自动清空，则先清空
        if self.auto_clear_status_var.get():
            self.status_text.delete(1.0, tk.END)
        
        # 添加时间戳
        timestamp = time.strftime("%H:%M:%S")
        status_line = f"[{timestamp}] {message}\n"
        self.status_text.insert(tk.END, status_line)
        
        # 如果启用了日志，保存状态信息
        if self.save_logs_var.get():
            self.save_to_log("status", message)
        
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
    
    def clear_status(self):
        """清空状态输出区域"""
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state=tk.DISABLED)
    
    def clear_data_dispaly(self):
        """清空数据显示区域"""
        for item in self.result_data_display.get_children():
            self.result_data_display.delete(item)
        self.update_status("数据显示已清空")
    
    def send_data(self):
        """发送数据到RS-485设备"""
        if not self.is_connected or not self.ser or not self.ser.is_open:
            messagebox.showwarning("警告", "请先连接设备")
            return
        
        try:
            data = self.send_text.get(1.0, tk.END).rstrip('\n')
            
            if not data:
                return
            
            # 处理十六进制发送
            if self.hex_send_var.get():
                # 移除所有空格
                data = data.replace(' ', '')
                # 检查十六进制格式是否有效
                if not all(c in '0123456789abcdefABCDEF' for c in data):
                    messagebox.showerror("格式错误", "十六进制数据包含无效字符")
                    return
                if len(data) % 2 != 0:
                    messagebox.showerror("格式错误", "十六进制数据长度必须为偶数")
                    return
                # 转换为字节
                send_bytes = binascii.unhexlify(data)
                data_str = ' '.join([f'{b:02X}' for b in send_bytes])
            else:
                # 处理自动换行
                if self.append_newline_var.get():
                    data += '\n'
                send_bytes = data.encode('utf-8')
                data_str = data
            
            # 发送数据
            self.ser.write(send_bytes)
            send_msg = f"已发送 {len(send_bytes)} 字节: {data_str}"
            self.log_message(send_msg)
            self.update_status(send_msg)
            logging.info(send_msg)
            
        except Exception as e:
            error_msg = f"发送失败: {str(e)}"
            messagebox.showerror("发送错误", error_msg)
            self.update_status(error_msg)
            logging.error(error_msg)
    
    def send_control_command(self, hex_data):
        self.workstatusflag = 1
        """发送控制指令"""
        if not self.is_connected or not self.ser or not self.ser.is_open:
            messagebox.showwarning("警告", "请先连接设备")
            return
        
        try:
            # 设置为十六进制发送模式
            original_hex_state = self.hex_send_var.get()
            self.hex_send_var.set(True)
            
            # 移除所有空格并转换为字节
            data = hex_data.replace(' ', '')
            send_bytes = binascii.unhexlify(data)
            
            # 发送数据
            self.ser.write(send_bytes)
            send_msg = f"已发送指令: {hex_data} ({len(send_bytes)} 字节)"
            self.log_message(send_msg)
            logging.info(send_msg)
            
            # 恢复原始的十六进制发送状态
            self.hex_send_var.set(original_hex_state)
            
        except Exception as e:
            error_msg = f"发送指令失败: {str(e)}"
            messagebox.showerror("发送错误", error_msg)
            self.update_status(error_msg)
            logging.error(error_msg)
    
    def load_commands(self):
        """加载保存的命令"""
        try:
            if os.path.exists("commands.json"):
                with open("commands.json", 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            logging.error(f"加载命令失败: {str(e)}")
            return {}
    
    def save_commands(self):
        """保存命令到文件"""
        try:
            with open("commands.json", 'w', encoding='utf-8') as f:
                json.dump(self.commands, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"保存命令失败: {str(e)}")
            messagebox.showerror("错误", f"保存命令失败: {str(e)}")
    
    def save_settings(self):
        """保存串口设置到文件"""
        try:
            settings = {
                "port": self.port_var.get(),
                "baudrate": self.baudrate_var.get(),
                "databits": self.databits_var.get(),
                "stopbits": self.stopbits_var.get(),
                "parity": self.parity_var.get(),
                "save_logs": self.save_logs_var.get(),
                "log_path": self.log_path_var.get()
            }
            with open("settings.json", 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"保存设置失败: {str(e)}")
    
    def load_settings(self):
        """从文件加载串口设置"""
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.port_var.set(settings.get("port", ""))
                    self.baudrate_var.set(settings.get("baudrate", "9600"))
                    self.databits_var.set(settings.get("databits", "8"))
                    self.stopbits_var.set(settings.get("stopbits", "1"))
                    self.parity_var.set(settings.get("parity", "N"))
                    # 加载日志设置
                    self.save_logs_var.set(settings.get("save_logs", True))
                    # 加载日志路径设置
                    saved_path = settings.get("log_path")
                    if saved_path and os.path.exists(saved_path):
                        self.log_path_var.set(saved_path)
        except Exception as e:
            logging.error(f"加载设置失败: {str(e)}")
    
    def clear_receive(self):
        """清空接收区"""
        self.receive_text.config(state=tk.NORMAL)
        self.receive_text.delete(1.0, tk.END)
        self.receive_text.config(state=tk.DISABLED)
    
    def clear_send(self):
        """清空发送区"""
        self.send_text.delete(1.0, tk.END)
    
    def clear_parse_results(self):
        """清空解析结果"""
        for item in self.parse_tree.get_children():
            self.parse_tree.delete(item)
        self.packet_count = 0  # 重置数据包计数器
        self.packet_info_var.set("解析结果已清空")
        self.update_status("解析结果已清空")
    
    def log_message(self, message):
        """在接收区添加日志消息"""
        self.receive_text.config(state=tk.NORMAL)
        timestamp = time.strftime("%H:%M:%S")
        log_line = f"[{timestamp}] [系统] {message}\n"
        self.receive_text.insert(tk.END, log_line)
        
        # 如果启用了日志，保存系统消息
        if self.save_logs_var.get():
            self.save_to_log("receive", f"[系统] {message}")
            
        self.receive_text.see(tk.END)
        self.receive_text.config(state=tk.DISABLED)
    
    def save_data_csv(self, data, filename):
        """
        根据提供的数据和文件名保存CSV数据
        
        参数:
            data: 要保存的数据，格式应为列表或列表的列表
            filename: 要保存的文件名（不包含路径）
        """
        # 数据验证
        if not data:
            self.update_status("没有数据可保存到CSV")
            return
            
        try:
            # 确保data文件夹存在
            data_path = self.ensure_data_folder_exists()
            if not data_path:
                return
            # 完整文件路径
            full_path = os.path.join(data_path, f"{self.log_timestamp}_{filename}.csv")
            file_exists = os.path.exists(full_path)
            
            # 数据格式标准化 - 确保是二维列表
            if isinstance(data, (list, tuple)):
                if not isinstance(data[0], (list, tuple)):
                    data = [data]  # 转换为二维列表
            else:
                self.log_message(f"保存CSV失败：数据格式不正确，应为列表类型")
                return
                
            # 验证数据行长度一致性
            row_length = len(data[0])
            for i, row in enumerate(data):
                if len(row) != row_length:
                    self.log_message(f"保存CSV警告：第{i+1}行数据长度与表头不一致，已跳过此行")
                    data[i] = []  # 标记为无效行
            
            # 过滤无效行
            valid_data = [row for row in data if row]
            if not valid_data:
                self.log_message("保存CSV失败：所有数据行均无效")
                return
                
            # 写入CSV文件
            with open(full_path, 'a' if file_exists else 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # 如果文件不存在，写入表头
                if not file_exists:
                    # 为vodata和dndata创建适当的表头
                    if filename in ['vodata', 'dndata']:
                        headers = ['PacketIndex', 'Timestamp']
                        # 添加CH1-CH4的各个数据项
                        for ch in range(1, 5):
                            # T1-T5
                            for t in range(1, 6):
                                headers.append(f'CH{ch}T{t}')
                            # PT和R
                            headers.append(f'CH{ch}PT')
                            headers.append(f'CH{ch}R')
                            # B1-B5
                            for b in range(1, 6):
                                headers.append(f'CH{ch}B{b}')
                        # 添加环境数据
                        headers.extend(['RT', 'MT', 'RH', 'MH'])
                        
                        # 验证表头长度与数据长度是否匹配
                        if len(headers) != len(valid_data[0]):
                            self.log_message(
                                f"CSV表头与数据长度不匹配：表头{len(headers)}列，数据{len(valid_data[0])}列"
                            )
                            
                        writer.writerow(headers)
                
                # 写入数据行
                row_count = len(valid_data)
                writer.writerows(valid_data)
            
            status_msg = f"数据已{'追加到' if file_exists else '保存到'}: {full_path}，共{row_count}行"
            self.log_message(status_msg)
            logging.info(status_msg)
            
        except PermissionError:
            error_msg = f"保存CSV数据失败：没有文件写入权限 - {full_path}"
            self.update_status(error_msg)
            logging.error(error_msg)
            messagebox.showerror("权限错误", f"无法写入文件：\n{full_path}\n请检查文件权限")
        except Exception as e:
            error_msg = f"保存CSV数据失败: {str(e)}"
            self.update_status(error_msg)
            logging.error(error_msg)


    def on_close(self):
        """关闭窗口时的处理"""
        # 如果启用了日志，记录程序关闭信息
        if self.save_logs_var.get():
            self.save_to_log("status", "程序已关闭")
            
        self.disconnect()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = RS485Tool(root)
    root.mainloop()
