import csv
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
import tkinter as tk
from tkinter import filedialog


def extract_numeric_data(file_path, start_row, column_index):
    numeric_data = []
    try:
        with open(file_path, 'r', newline='', encoding='utf-16') as file:
            reader = csv.reader(file)
            for idx, row in enumerate(reader):
                if idx < start_row:
                    continue
                try:
                    value = float(row[column_index - 1].replace(' (VDC)', ''))
                    numeric_data.append(value)
                except (ValueError, IndexError):
                    continue
        return numeric_data
    except FileNotFoundError:
        print("错误: 文件未找到!")
    except Exception as e:
        print(f"错误: {str(e)}")
    return None


def plot_data(data, file_path):
    if not data:
        print("警告: 没有有效数值数据可供绘制")
        return

    fig, ax = plt.subplots(figsize=(12, 6), facecolor='#f9fafb')
    ax.plot(data, color='#2563eb', linewidth=1.8, alpha=0.9, label='Voltage (V)')

    if len(data) > 100:
        ax.set_xticks(range(0, len(data), max(1, len(data) // 20)))
    ax.grid(axis='y', linestyle='--', alpha=0.7, color='#64748b')

    ax.set_xlabel('times', fontsize=14, color='#334155')
    ax.set_ylabel('Voltage (V)', fontsize=14, color='#334155')
    fig.autofmt_xdate(rotation=45)

    stats = f'start: {data[0]:.6f}\nfinish: {data[-1]:.6f}\nnian: {min(data):.6f} ~ {max(data):.6f}'
    ax.text(0.98, 0.02, stats, transform=ax.transAxes,
            bbox=dict(facecolor='white', alpha=0.9, edgecolor='#e5e7eb'),
            ha='right', va='bottom', fontsize=12, color='#4b5563')

    button_ax = fig.add_axes([0.87, 0.01, 0.1, 0.05])
    save_button = Button(button_ax, 'save', color='#2563eb', hovercolor='#1d4ed8')

    def save_callback(event):
        plt.savefig('voltage_plot.png', dpi=300, bbox_inches='tight')
        print("图片已保存为 voltage_plot.png")

    save_button.on_clicked(save_callback)

    for text in save_button.label.get_children():
        text.set_color('white')
        text.set_fontsize(12)

    plt.tight_layout()
    plt.show()


def select_file():
    root = tk.Tk()
    root.title("CSV 文件数据绘图")

    # 创建起始行和列索引的标签和文本框
    tk.Label(root, text="起始行:").grid(row=0, column=0)
    start_row_entry = tk.Entry(root)
    start_row_entry.insert(0, "12")
    start_row_entry.grid(row=0, column=1)

    tk.Label(root, text="显示列:").grid(row=1, column=0)
    column_index_entry = tk.Entry(root)
    column_index_entry.insert(0, "3")
    column_index_entry.grid(row=1, column=1)

    def open_file_dialog():
        try:
            start_row = int(start_row_entry.get())
            column_index = int(column_index_entry.get())
        except ValueError:
            print("请输入有效的整数作为起始行和显示列。")
            return

        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            data = extract_numeric_data(file_path, start_row, column_index)
            if data:
                plot_data(data, file_path)

    # 创建选择文件的按钮
    tk.Button(root, text="选择文件", command=open_file_dialog).grid(row=2, column=0, columnspan=2)

    root.mainloop()


if __name__ == "__main__":
    select_file()
