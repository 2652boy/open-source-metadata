import os
import re
import html
import subprocess
import sys
import threading
from datetime import datetime
from PIL import Image
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font

# -*- coding: utf-8 -*-
"""
文件名：档案元数据写入工具
作者：Community Contributors
版权声明：Copyright 2025 Community Contributors.
许可证：本程序基于 MIT 许可证发布。
说明：此程序是一个基于Tkinter的图像元数据读取与写入工具，集成ExifTool，支持XML和Excel导出。
"""

from write_metadata import write_metadata

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

user_fields = []

if sys.platform == "win32":
    CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW
else:
    CREATE_NO_WINDOW = 0
ExifVersion = '1298'


# ======= 元数据提取函数 =======
def get_file_metadata(file_path):
    try:
        stats = os.stat(file_path)
        return {
            "创建时间": datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
            "修改时间": datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        }
    except Exception as e:
        return {"错误": f"无法获取文件元数据: {e}"}


def get_image_properties(image_path):
    try:
        image = Image.open(image_path)
        dpi = image.info.get('dpi', (72, 72))
        width, height = image.size
        megapixels = (width * height) / 1_000_000
        return {
            "图片格式": image.format,
            "图片尺寸": image.size,
            "颜色模式": image.mode,
            "水平分辨率": f"{dpi[0]:.0f}",
            "垂直分辨率": f"{dpi[1]:.0f}",
            "兆像素": f"{megapixels:.2f} MP"
        }
    except Exception as e:
        return {"错误": f"无法读取图片属性: {e}"}


def get_detailed_metadata(image_path):
    try:
        exiftool_path = os.path.join(PROJECT_DIR, "exiftool", "exiftool.exe")
        result = subprocess.run(
            [exiftool_path, image_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
            check=False,
            creationflags=CREATE_NO_WINDOW
        )
        output = result.stdout.decode('utf-8', errors='ignore')
        metadata = {}
        for line in output.splitlines():
            if ": " in line:
                key, value = line.split(": ", 1)
                metadata[key.strip()] = value.strip()
        return metadata
    except Exception as e:
        return {"错误": f"无法获取元数据: {e}"}


def translate_metadata_key(key):
    translations = {
        "File Name": "文件名称",
        "Directory": "存储路径",
        "File Size": "文件大小",
        "File Type": "文件类型",
        "X Resolution": "水平分辨率",
        "Y Resolution": "垂直分辨率",
        "Image Size": "图像尺寸",
        "Megapixels": "兆像素",
    }
    return translations.get(key, key)


EXCLUDE_FIELDS = ["缩略图", "缩略图长度", "缩略图偏移", "Exif工具版本", "文件访问日期/时间"]


def filter_unwanted_fields(info):
    return {k: v for k, v in info.items() if k not in EXCLUDE_FIELDS or k == "警告"}


def get_image_info(path):
    metadata = get_file_metadata(path)
    properties = get_image_properties(path)
    detailed = get_detailed_metadata(path)

    width, height = properties.get('图片尺寸', (0, 0))
    x_dpi = properties.get("水平分辨率", "")
    y_dpi = properties.get("垂直分辨率", "")

    try:
        file_size_kb = os.path.getsize(path) / 1024
    except:
        file_size_kb = ""

    try:
        abs_path = os.path.abspath(path).replace("/", "\\")
        relative_path = re.sub(r"^[A-Za-z]:", "", abs_path)
        offline_path = abs_path
    except:
        relative_path = ""
        offline_path = path

    info = {
        "文件名称": os.path.basename(path),
        "创建时间": metadata.get("创建时间", ""),
        "修改时间": metadata.get("修改时间", ""),
        "拍摄时间": metadata.get("修改时间", ""),
        "图片格式": properties.get("图片格式", ""),
        "图片尺寸": f"{width}x{height}",
        "颜色模式": properties.get("颜色模式", ""),
        "水平分辨率": x_dpi,
        "垂直分辨率": y_dpi,
        "兆像素": properties.get("兆像素", ""),
        "文件大小": f"{file_size_kb:.1f} KB" if file_size_kb else "",
        "相对地址": relative_path,
        "离线地址": offline_path
    }

    if isinstance(detailed, dict):
        info.update(detailed)
    else:
        info["错误"] = detailed

    for field in specific_fields:
        name = field["label"].get()
        value = field["value"].get()
        if name:
            info[name] = value

    for field in user_fields:
        name = field["name"].get()
        value = field["value"].get()
        if name:
            info[name] = value
    translated = {translate_metadata_key(k): v for k, v in info.items()}
    return filter_unwanted_fields(translated)


def export_to_xml(image_infos, output_file):
    root = ET.Element("信息")
    root.set("版本", "1.0")

    def clean_field_name(name):
        return re.sub(r"[^a-zA-Z0-9_\u4e00-\u9fff]", "_", name)

    for info in image_infos:
        elem = ET.SubElement(root, "图片信息")
        for k, v in info.items():
            clean_k = clean_field_name(k)
            sub = ET.SubElement(elem, clean_k)
            sub.text = html.escape(str(v))

    tree = ET.ElementTree(root)
    tree.write(output_file, encoding="utf-8", xml_declaration=True)


def export_to_xlsx(image_infos, output_file):
    df = pd.DataFrame([filter_unwanted_fields(i) for i in image_infos])
    df.columns = [translate_metadata_key(c) for c in df.columns]
    df.to_excel(output_file, index=False, engine='openpyxl')


def get_image_files_in_folder(folder_path):
    result = {}
    for root, _, files in os.walk(folder_path):
        imgs = [os.path.join(root, f) for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
        if imgs:
            result[root] = imgs
    return result


def process_images_threaded():
    threading.Thread(target=process_images_wrapper).start()


def process_images_wrapper():
    try:
        process_images()
    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出现异常：{e}")


def process_images():
    path = source_path_var.get()
    output_dir = output_path_var.get()

    if not path or not os.path.isdir(path):
        messagebox.showwarning("警告", "请选择有效的源文件夹")
        return

    if not output_dir or not os.path.isdir(output_dir):
        messagebox.showwarning("警告", "请选择有效的输出文件夹")
        return

    folders = get_image_files_in_folder(path)
    if not folders:
        messagebox.showinfo("提示", "所选文件夹中没有找到图片文件")
        return

    total_images = sum(len(imgs) for imgs in folders.values())
    count = 0
    progress_label.config(text="开始处理...")
    progress_bar["maximum"] = total_images
    progress_bar["value"] = 0
    root.update_idletasks()

    for folder, images in folders.items():
        infos = []
        # 获取文件夹中所有图片的创建时间
        creation_times = []
        for img in images:
            try:
                stats = os.stat(img)
                creation_time = datetime.fromtimestamp(stats.st_ctime)
                creation_times.append(creation_time)
            except:
                pass

        # 计算数字化时间
        digitization_time = ""
        if len(creation_times) >= 1:
            first_time = min(creation_times).strftime('%Y-%m-%d %H:%M:%S')
            last_time = max(creation_times).strftime('%Y-%m-%d %H:%M:%S')
            digitization_time = f"{first_time}-{last_time}"

        for img in images:
            count += 1
            progress_label.config(text=f"正在处理第 {count}/{total_images} 张图片")
            progress_bar["value"] = count
            root.update_idletasks()

            info = get_image_info(img)
            img_name = os.path.basename(img)
            info['ExifVersion'] = ExifVersion
            info['标题'] = img_name
            # 添加数字化时间字段
            info['数字化时间'] = digitization_time
            write_metadata(img, info)
            infos.append(info)

        folder_name = os.path.basename(folder)
        xml_file = os.path.join(output_dir, f"{folder_name}.xml")
        xlsx_file = os.path.join(output_dir, f"{folder_name}.xlsx")
        export_to_xml(infos, xml_file)
        export_to_xlsx(infos, xlsx_file)

    progress_label.config(text="处理完成！")
    messagebox.showinfo("完成", f"所有图片信息已导出到:\n{output_dir}")


# ======= UI界面优化 =======
# 颜色和字体定义
BG_COLOR = "#f0f0f0"
FRAME_BG = "#ffffff"
BUTTON_COLOR = "#4a90e2"
BUTTON_HOVER = "#357abd"
TEXT_COLOR = "#333333"
HEADER_COLOR = "#2c3e50"
FONT_NAME = "微软雅黑"
FONT_SIZE = 10
TITLE_FONT_SIZE = 12

root = tk.Tk()
root.title("图片元数据处理工具")
root.geometry("980x1020")
root.resizable(True, True)
root.configure(bg=BG_COLOR)

# 设置样式
style = ttk.Style()
style.theme_use('clam')

# 配置字体
default_font = Font(family=FONT_NAME, size=FONT_SIZE)
title_font = Font(family=FONT_NAME, size=TITLE_FONT_SIZE, weight="bold")
root.option_add("*Font", default_font)

# 配置样式
style.configure("TFrame", background=BG_COLOR)
style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
style.configure("TButton", background=BUTTON_COLOR, foreground="white",
                borderwidth=1, focusthickness=3, focuscolor='none')
style.map("TButton",
          background=[('active', BUTTON_HOVER), ('pressed', BUTTON_HOVER)],
          foreground=[('active', 'white'), ('pressed', 'white')])
style.configure("TLabelFrame", background=BG_COLOR, foreground=HEADER_COLOR)
style.configure("TLabelFrame.Label", font=title_font)
style.configure("TEntry", fieldbackground="white")
style.configure("TProgressbar", thickness=20)

source_path_var = tk.StringVar()
output_path_var = tk.StringVar()

# 主容器
main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill=tk.BOTH, expand=True)

# 顶部控制区域
control_frame = ttk.Frame(main_frame)
control_frame.pack(fill=tk.X, pady=(0, 10))

# 源文件夹选择
frame_source = ttk.LabelFrame(control_frame, text="图片文件夹选择", padding=10)
frame_source.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

path_entry = ttk.Entry(frame_source, textvariable=source_path_var, width=40)
path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

browse_btn = ttk.Button(frame_source, text="浏览...",
                        command=lambda: source_path_var.set(filedialog.askdirectory()))
browse_btn.pack(side=tk.LEFT, padx=5)

# 输出文件夹选择
frame_output = ttk.LabelFrame(control_frame, text="输出文件夹选择", padding=10)
frame_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

output_entry = ttk.Entry(frame_output, textvariable=output_path_var, width=40)
output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

output_btn = ttk.Button(frame_output, text="浏览...",
                        command=lambda: output_path_var.set(filedialog.askdirectory()))
output_btn.pack(side=tk.LEFT, padx=5)

# 操作按钮区域
button_frame = ttk.Frame(main_frame)
button_frame.pack(fill=tk.X, pady=(0, 10))

process_btn = ttk.Button(button_frame, text="开始处理",
                         command=process_images_threaded, style="TButton")
process_btn.pack(side=tk.LEFT, padx=5)

add_field_btn = ttk.Button(button_frame, text="添加自定义字段",
                           command=lambda: add_custom_field())
add_field_btn.pack(side=tk.LEFT, padx=5)

# 进度显示区域
status_frame = ttk.LabelFrame(main_frame, text="处理进度", padding=10)
status_frame.pack(fill=tk.X, pady=(0, 10))

progress_label = ttk.Label(status_frame, text="等待处理...",
                           font=(FONT_NAME, FONT_SIZE, "bold"))
progress_label.pack(side=tk.LEFT, padx=5)

progress_bar = ttk.Progressbar(status_frame, orient="horizontal",
                               length=400, mode="determinate")
progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

# 元数据字段区域 - 使用Notebook实现标签页
notebook = ttk.Notebook(main_frame)
notebook.pack(fill=tk.BOTH, expand=True)

# 预定义字段标签页
predefined_frame = ttk.Frame(notebook)
notebook.add(predefined_frame, text="预定义字段")

# 自定义字段标签页
custom_frame = ttk.Frame(notebook)
notebook.add(custom_frame, text="自定义字段")

# 预定义字段内容
predefined_canvas = tk.Canvas(predefined_frame, bg=FRAME_BG)
predefined_scrollbar = ttk.Scrollbar(predefined_frame, orient="vertical",
                                     command=predefined_canvas.yview)
predefined_scrollable_frame = ttk.Frame(predefined_canvas)

predefined_scrollable_frame.bind(
    "<Configure>",
    lambda e: predefined_canvas.configure(
        scrollregion=predefined_canvas.bbox("all")
    )
)

predefined_canvas.create_window((0, 0), window=predefined_scrollable_frame, anchor="nw")
predefined_canvas.configure(yscrollcommand=predefined_scrollbar.set)

predefined_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
predefined_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# 预定义字段
specific_fields = [
    {"label": tk.StringVar(value=field), "value": tk.StringVar()}
    for field in [
        "标题", "主题", "标记", "备注",
        "数字化对象描述", "数字化授权描述", "格式名称", "格式版本", "色彩空间", "压缩方案", "压缩率",
        "设备类型", "设备制造商", "设备型号", "设备序列号", "设备感光器",
        "数字化软件名称", "数字化软件版本", "数字化软件生产商",
        "版权所有者", "版权ID", "版权期限", "阅读所需软硬件条件", "数字化成果移交接收信息"
    ]
]

for i, field in enumerate(specific_fields):
    frame = ttk.Frame(predefined_scrollable_frame)
    frame.pack(fill=tk.X, padx=5, pady=3)

    label_text = field["label"].get()
    ttk.Label(frame, text=label_text, width=20, anchor="w").pack(side=tk.LEFT, padx=5)
    ttk.Entry(frame, textvariable=field["value"], width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

# 自定义字段内容
custom_canvas = tk.Canvas(custom_frame, bg=FRAME_BG)
custom_scrollbar = ttk.Scrollbar(custom_frame, orient="vertical",
                                 command=custom_canvas.yview)
custom_scrollable_frame = ttk.Frame(custom_canvas)

custom_scrollable_frame.bind(
    "<Configure>",
    lambda e: custom_canvas.configure(
        scrollregion=custom_canvas.bbox("all")
    )
)

custom_canvas.create_window((0, 0), window=custom_scrollable_frame, anchor="nw")
custom_canvas.configure(yscrollcommand=custom_scrollbar.set)

custom_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
custom_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)


def add_custom_field():
    frame = ttk.Frame(custom_scrollable_frame)
    frame.pack(fill=tk.X, padx=5, pady=3)

    name_var = tk.StringVar()
    val_var = tk.StringVar()

    name_entry = ttk.Entry(frame, textvariable=name_var, width=20)
    name_entry.pack(side=tk.LEFT, padx=5)

    val_entry = ttk.Entry(frame, textvariable=val_var, width=60)
    val_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

    del_button = ttk.Button(frame, text="删除",
                            command=lambda f=frame, idx=len(user_fields): delete_custom_field(f, idx))
    del_button.pack(side=tk.LEFT, padx=5)

    user_fields.append({
        "name": name_var,
        "value": val_var,
        "frame": frame
    })


def delete_custom_field(frame, index):
    frame.pack_forget()
    frame.destroy()
    if index < len(user_fields):
        user_fields.pop(index)
    # 重新索引删除按钮的命令
    for i, field in enumerate(user_fields):
        for widget in field["frame"].winfo_children():
            if isinstance(widget, ttk.Button) and widget["text"] == "删除":
                widget["command"] = lambda f=field["frame"], idx=i: delete_custom_field(f, idx)


# 添加一些初始自定义字段
for _ in range(3):
    add_custom_field()

# 设置窗口最小大小
root.update()
root.minsize(root.winfo_width(), 600)
# 窗口底部的版权信息
copyright_label = ttk.Label(
    root,
    text="Copyright 2025 Community Contributors. Licensed under MIT.",
    foreground="gray",
    anchor="center"
)
copyright_label.pack(side=tk.BOTTOM, fill=tk.X)

root.mainloop()
