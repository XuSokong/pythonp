import tkinter as tk
from tkinter import messagebox, ttk
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from scipy.interpolate import interp1d
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# 设置中文字体支持
# 设置中文字体支持（移除不存在的字体）
plt.rcParams["font.family"] = ["SimHei", "Microsoft YaHei", "SimSun"]


class TemperatureResistanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("温度-电阻插值计算器")
        self.root.geometry("800x600")

        # 数据初始化
        self.x_datas = None
        self.y_datas = None
        self.interp_function = None
        self.x_cha = None
        self.y_cha = None

        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧控制面板
        control_frame = ttk.LabelFrame(main_frame, text="控制面板", padding="10")
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)

        # 文件路径输入
        ttk.Label(control_frame, text="Excel文件路径:").pack(anchor=tk.W, pady=5)
        self.file_path_var = tk.StringVar(value=r"D:\xusokong\Justintime\V-T.xlsx")
        file_entry = ttk.Entry(control_frame, textvariable=self.file_path_var, width=30)
        file_entry.pack(fill=tk.X, pady=5)

        # 加载数据按钮
        load_button = ttk.Button(control_frame, text="加载数据", command=self.load_data)
        load_button.pack(fill=tk.X, pady=10)

        # 电压输入
        ttk.Label(control_frame, text="电压值 (V):").pack(anchor=tk.W, pady=5)
        self.voltage_var = tk.DoubleVar(value=0.482257)
        voltage_entry = ttk.Entry(control_frame, textvariable=self.voltage_var, width=15)
        voltage_entry.pack(fill=tk.X, pady=5)

        # 计算按钮
        calculate_button = ttk.Button(control_frame, text="计算温度", command=self.calculate_temperature)
        calculate_button.pack(fill=tk.X, pady=10)

        # 结果显示 - 使用文本框代替标签
        ttk.Label(control_frame, text="计算结果:").pack(anchor=tk.W, pady=5)
        self.result_text = tk.Text(control_frame, height=6, width=30, wrap=tk.WORD)
        self.result_text.pack(fill=tk.X, pady=5)
        self.result_text.insert(tk.END, "等待计算...")
        self.result_text.config(state=tk.DISABLED)  # 设置为只读

        # 右侧图表区域
        self.figure, self.ax = plt.subplots(figsize=(5, 4), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=main_frame)
        self.canvas.get_tk_widget().pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 初始化图表
        self.ax.set_title("温度-电阻曲线")
        self.ax.set_xlabel("电阻值 (Ohms)")
        self.ax.set_ylabel("温度 (°C)")
        self.ax.grid(True)
        self.canvas.draw()

    def load_data(self):
        """加载Excel数据并进行插值计算"""
        try:
            file_path = self.file_path_var.get()
            df = pd.read_excel(file_path)

            self.x_datas = df['电阻值（Ohms）']
            self.y_datas = df['温度（°C）']

            # 执行插值
            self.interp_function = interp1d(
                self.x_datas, self.y_datas,
                kind='linear', bounds_error=False, fill_value=-273.15
            )

            # 生成插值后的点
            self.x_cha = np.linspace(min(self.x_datas), max(self.x_datas), 1000)
            self.y_cha = self.interp_function(self.x_cha)

            # 更新图表
            self.update_plot()

            self.status_var.set(f"数据加载成功，点数: {len(df)}")
            self.update_result_text("数据已加载，请输入电压值并计算")

        except Exception as e:
            messagebox.showerror("错误", f"加载数据时出错: {str(e)}")
            self.status_var.set("数据加载失败")

    def update_plot(self):
        """更新图表显示"""
        self.ax.clear()
        self.ax.set_title("温度-电阻曲线")
        self.ax.set_xlabel("电阻值 (Ohms)")
        self.ax.set_ylabel("温度 (°C)")
        self.ax.grid(True)

        # 绘制原始数据点
        self.ax.plot(self.x_datas, self.y_datas, 'ro', markersize=3, label='原始数据')

        # 绘制插值曲线
        self.ax.plot(self.x_cha, self.y_cha, 'b-', linewidth=1, label='插值曲线')

        self.ax.legend()
        self.canvas.draw()

    def update_result_text(self, text):
        """更新结果文本框的内容"""
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, text)
        self.result_text.config(state=tk.DISABLED)

    def calculate_temperature(self):
        """根据电压值计算温度"""
        if self.interp_function is None:
            messagebox.showwarning("警告", "请先加载数据")
            return

        try:
            voltage = self.voltage_var.get()
            resistance = voltage * 150000 / (3.29597 - voltage)
            temperature = self.interp_function(resistance)

            # 在图表上标记计算点
            self.update_plot()
            self.ax.plot(resistance, temperature, 'g*', markersize=10, label='计算点')
            self.ax.annotate(f'({resistance:.6f} Ω, {temperature:.6f}°C)',
                             xy=(resistance, temperature),
                             xytext=(10, 10),
                             textcoords='offset points',
                             arrowprops=dict(arrowstyle='->'))
            self.ax.legend()
            self.canvas.draw()

            # 更新结果文本框
            result_text = f"电压值: {voltage:.6f} V\n"
            result_text += f"电阻值: {resistance:.6f} Ω\n"
            result_text += f"温度: {temperature:.6f} °C"
            self.update_result_text(result_text)

            self.status_var.set("计算完成")

        except Exception as e:
            messagebox.showerror("错误", f"计算时出错: {str(e)}")
            self.status_var.set("计算失败")


def main():
    root = tk.Tk()
    app = TemperatureResistanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
