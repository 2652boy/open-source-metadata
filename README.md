# 档案元数据写入工具

一个使用 Python、Tkinter 和 ExifTool 编写的 Windows 桌面工具。程序可以递归读取图片目录，采集图片和文件属性，把已映射的元数据直接写入图片，并按每个包含图片的文件夹导出 XML 与 XLSX。

## 主要功能

- 递归扫描源目录及其子目录。
- 处理 `.jpg`、`.jpeg`、`.png` 图片。
- 自动读取文件时间、图片格式、尺寸、颜色模式、分辨率、文件大小和路径。
- 通过 ExifTool 把支持的字段直接写入源图片。
- 支持填写固定字段和添加自定义字段。
- 计算每个图片文件夹内最早到最晚的文件创建时间，作为“数字化时间”。
- 每个包含图片的文件夹分别生成一个 XML 和一个 XLSX 文件。
- 使用 UTF-8 参数文件调用 ExifTool，支持中文文件名和中文目录。
- Windows 下后台调用 ExifTool，不弹出 CMD 窗口。

## 重要说明

程序会直接修改源图片的元数据，不会自动创建图片备份。正式处理前应先备份源图片，或在复制出的测试目录中确认写入结果。

程序目前不处理 TIF/TIFF、PDF、OFD、RAW 或视频文件。界面中显示的字段不代表全部字段都能写进图片，只有下表中已映射的字段会传给 ExifTool；其他读取或自定义字段仍可出现在 XML/XLSX 中。

## 写入图片的字段

| 界面/内部字段 | ExifTool 标签 | 说明 |
|---|---|---|
| EXIF版本 | `ExifVersion` | 程序自动填入版本值 |
| 程序名称 | `Software` | 软件名称 |
| 数字化软件版本 | `Software` | 与程序名称共用标签，后写入值会覆盖前值 |
| 照相机制造商 | `Make` | 设备制造商别名 |
| 设备制造商 | `Make` | 设备制造商 |
| 照相机型号 | `Model` | 设备型号别名 |
| 设备型号 | `Model` | 设备型号 |
| 作者 | `Artist` | 作者信息 |
| 版权所有者 | `Artist` | 与作者共用标签，后写入值会覆盖前值 |
| 版权 | `Copyright` | 版权信息 |
| 版权ID | `Copyright` | 与版权共用标签，后写入值会覆盖前值 |
| 备注 | `XPComment` | Windows 图片备注 |
| 主题 | `XPSubject` | Windows 图片主题 |
| 标题 | `Title` | 主程序默认使用图片文件名覆盖该值 |
| 文件名称 | `Title` | 标题别名 |
| 标记 | `Subject` | 标记/主题标签 |
| 颜色表示 | `ColorSpace` | 色彩空间别名 |
| 色彩空间 | `ColorSpace` | 色彩空间 |
| 拍摄时间 | `DateTimeOriginal` | 主程序当前使用文件修改时间生成 |
| 修改时间 | `DateTimeOriginal` | 与拍摄时间共用标签 |

### 当前不会写入图片的预定义字段

以下字段虽然会显示在界面中，但当前 `write_metadata.py` 没有对应的 ExifTool 映射，因此主要用于 XML/XLSX 导出：

- 数字化对象描述
- 数字化授权描述
- 格式名称、格式版本
- 压缩方案、压缩率
- 设备类型、设备序列号、设备感光器
- 数字化软件名称、数字化软件生产商
- 版权期限
- 阅读所需软硬件条件
- 数字化成果移交接收信息
- 用户添加但未使用有效 ExifTool 标签名的自定义字段

## 自动读取和导出的字段

程序会读取或生成下列信息并写入 XML/XLSX：

- 文件名称
- 创建时间、修改时间、拍摄时间
- 图片格式、图片尺寸、颜色模式
- 水平分辨率、垂直分辨率
- 兆像素、文件大小
- 相对地址、离线地址
- ExifTool 返回的其他可解析元数据
- 数字化时间
- 固定字段和自定义字段

导出文件可能包含本机目录和业务元数据。不要把实际业务导出的 XML/XLSX、测试图片或扫描件直接提交到公开仓库。

## 输出结构

假设源目录如下：

```text
source/
|-- 档号-001/
|   |-- 001.jpg
|   `-- 002.jpg
`-- 档号-002/
    `-- 001.png
```

程序会在用户选择的输出目录中生成：

```text
output/
|-- 档号-001.xml
|-- 档号-001.xlsx
|-- 档号-002.xml
`-- 档号-002.xlsx
```

XML 和 XLSX 保存相同批次的图片信息。图片元数据则直接写入源目录中的原图片。

## 环境要求

- Windows 10/11
- Python 3.10 或更高版本
- ExifTool Windows 版本
- Tkinter（常规 Windows Python 安装通常自带）

安装 Python 依赖：

```powershell
python -m pip install -r requirements.txt
```

## 安装 ExifTool

从 ExifTool 官方网站下载 Windows 版本：

https://exiftool.org/

将 `exiftool.exe` 和相邻的 `exiftool_files` 目录放到项目的 `exiftool/` 目录：

```text
open-source-metadata/
|-- exiftool/
|   |-- exiftool.exe
|   `-- exiftool_files/
|-- r.py
|-- write_metadata.py
|-- requirements.txt
`-- README.md
```

ExifTool 不包含在本仓库中。若自行分发 ExifTool，请同时遵守其官方许可证。

## 运行

```powershell
python r.py
```

使用步骤：

1. 点击“图片文件夹选择”旁的“浏览”，选择源图片根目录。
2. 点击“输出文件夹选择”旁的“浏览”，选择 XML/XLSX 输出目录。
3. 在“预定义字段”中填写需要的业务字段。
4. 必要时在“自定义字段”中添加额外字段。
5. 点击“开始处理”。
6. 完成后检查源图片元数据以及输出目录中的 XML/XLSX。

## 处理流程

```text
扫描源目录
  -> 按包含图片的文件夹分组
  -> 读取文件和图片属性
  -> 合并预定义字段与自定义字段
  -> 写入已映射的图片元数据
  -> 生成该文件夹的 XML 和 XLSX
```

## 使用 Python 调用写入模块

也可以绕过图形界面直接调用：

```python
from write_metadata import write_metadata

write_metadata(
    r"D:\images\001.jpg",
    {
        "标题": "示例标题",
        "主题": "档案数字化",
        "备注": "测试写入",
        "设备制造商": "Example Vendor",
        "设备型号": "Example Model",
        "版权所有者": "Example Organization",
    },
)
```

传入未映射的中文字段时，程序会在控制台输出 `not map:`，该字段不会写入图片。

## 已知限制

- 处理线程会直接更新 Tkinter 控件；大批量任务可能出现界面响应问题。
- 写入前没有自动备份，也没有写入后的逐字段回读校验。
- 多个中文字段映射到同一个 ExifTool 标签时，后处理的字段值会覆盖前一个值。
- PNG 能被扫描和传给 ExifTool，但不同软件对 PNG 元数据的显示和兼容程度不同。
- XML 中的字段名会清理特殊字符，自定义字段名称可能发生变化。
- 路径和读取到的详细元数据会进入导出文件，导出内容不适合未经检查直接公开。

## 许可证

源代码使用 MIT License，详见 `LICENSE`。

ExifTool 是独立的第三方程序，其许可证以官方发行包为准。
