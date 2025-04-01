import csv
import re
import os
import tkinter as tk
from tkinter import messagebox
from itertools import zip_longest
import matplotlib.pyplot as plt


def is_valid_hex_string(hex_str):
    hex_str_clean = re.sub(r'\s+', '', hex_str.strip().upper())
    return len(hex_str_clean) % 2 == 0 and all(c in '0123456789ABCDEF' for c in hex_str_clean)


def is_valid_filename(filename):
    return bool(re.match(r'^[a-zA-Z0-9_.-]+$', filename)) and not any(
        sep in filename for sep in os.path.sep + (os.path.altsep or '') if sep)


def hexstringtodecstring(hex_string):
    if not is_valid_hex_string(hex_string):
        raise ValueError("输入的十六进制字符串格式不正确")

    hex_pairs = re.findall(r'([0-9a-fA-F]{2})\s*', hex_string.strip())
    if len(hex_pairs) % 2 != 0:
        raise ValueError("十六进制数的数量必须为偶数")

    return [int(hex_pairs[i] + hex_pairs[i + 1], 16) for i in range(0, len(hex_pairs), 2)]


def saveresultascsv(decimal_result, filename):
    if not is_valid_filename(filename):
        raise ValueError("文件名包含非法字符")

    csv_filename = f"{os.path.splitext(filename)[0]}.csv"

    try:
        length = len(decimal_result)
        one_third = length // 3
        with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['times1', 'times2', 'times3'])
            for i in range(one_third):
                row = [decimal_result[i]]
                if i + one_third < length:
                    row.append(decimal_result[i + one_third])
                if i + 2 * one_third < length:
                    row.append(decimal_result[i + 2 * one_third])
                writer.writerow(row)
        print(f"数据已成功保存到 {csv_filename}")
    except PermissionError:
        print(f"权限错误：无法写入 {csv_filename}")
    except Exception as e:
        print(f"保存 CSV 失败：{str(e)}")


def save_hex_to_txt(hex_string, filename):
    if not is_valid_filename(filename):
        raise ValueError("文件名包含非法字符")

    txt_filename = f"{os.path.splitext(filename)[0]}.txt"

    try:
        with open(txt_filename, 'w', encoding='utf-8') as txtfile:
            txtfile.write(hex_string.strip() + '\n')
        print(f"十六进制已保存到 {txt_filename}")
        messagebox.showinfo("成功", f"十六进制已保存到 {txt_filename}")
    except PermissionError:
        print(f"权限错误：无法写入 {txt_filename}")
        messagebox.showerror("错误", f"权限错误：无法写入 {txt_filename}")
    except Exception as e:
        print(f"保存 TXT 失败：{str(e)}")
        messagebox.showerror("错误", f"保存 TXT 失败：{str(e)}")


def plot_decimal_result(decimal_result, csvfilename):
    length = len(decimal_result)
    one_third = length // 3
    tlength = one_third - 8
    x = range(tlength)

    times1 = []
    times2 = []
    times3 = []
    zeroleve = [32768]
    fig, ax = plt.subplots()

    for i in range(tlength):
        times1.append(decimal_result[i])
    for i in range(one_third, one_third + tlength):
        times2.append(decimal_result[i])
    for i in range(2 * one_third, 2 * one_third + tlength):
        times3.append(decimal_result[i])
    plt.plot(x, times1, label='times1')
    if len(times2) > 0:
        plt.plot(range(len(times2)), times2, label='times2')
    if len(times3) > 0:
        plt.plot(range(len(times3)), times3, label='times3')
    # 绘制 zeroleve 线
    plt.axhline(y=zeroleve[0], color='g', linestyle='--', label='zero')
    if one_third > 8:
        avg_voltage = (decimal_result[one_third - 8] + decimal_result[2 * one_third - 8] + decimal_result[
            3 * one_third - 8]) / 3
        plt.axhline(y=avg_voltage, color='b', linestyle='--', label=f'avg:{avg_voltage:.2f}')

    plt.xlabel('Index')
    plt.ylabel('Value')
    plt.title('Decimal Result Plot')
    plt.grid(True)

    # 设置图例放在图形外
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()

    # 保存图片
    img_filename = f"{os.path.splitext(csvfilename)[0]}.png"
    try:
        plt.savefig(img_filename, bbox_inches='tight')
        print(f"图片已成功保存到 {img_filename}")
    except PermissionError:
        print(f"权限错误：无法保存图片到 {img_filename}")
    except Exception as e:
        print(f"保存图片失败：{str(e)}")

    plt.show()


def run_program():
    hex_string = hex_input.get("1.0", tk.END)
    # 滤除指定行
    lines = hex_string.splitlines()
    filtered_lines = [line for line in lines if "17 62 63 64 65 66 67 00" not in line]
    hex_string = '\n'.join(filtered_lines)

    csvfilename = csvfile_input.get()

    try:
        decimal_result = hexstringtodecstring(hex_string)
        print(f"decimal_result 中元素的个数为: {len(decimal_result)}")
        saveresultascsv(decimal_result, csvfilename)
        save_hex_to_txt(hex_string, csvfilename)
        plot_decimal_result(decimal_result, csvfilename)
    except ValueError as e:
        print(f"输入错误: {e}")
        messagebox.showerror("输入错误", f"输入错误: {e}")


# 创建主窗口
root = tk.Tk()
root.title("十六进制转 CSV")

# 创建十六进制字符串输入框
hex_label = tk.Label(root, text="请输入十六进制字符串:")
hex_label.pack()
hex_input = tk.Text(root, height=10, width=50)
hex_input.pack()

# 创建 CSV 文件名输入框
csvfile_label = tk.Label(root, text="请输入 CSV 文件名:")
csvfile_label.pack()
csvfile_input = tk.Entry(root, width=50)
csvfile_input.pack()

# 创建运行按钮
run_button = tk.Button(root, text="运行程序", command=run_program)
run_button.pack()

# 运行主循环
root.mainloop()
