import subprocess
import sys
from pathlib import Path
import tempfile

maps = {
    # 'EXIF版本': 'ExifVersion',
    'EXIF版本': 'ExifVersion',
    '程序名称': 'Software',
    '数字化软件版本': 'Software',
    # '大小': 'FileSize',
    # '修改时间': 'FileModificationDateTime',
    # '修改日期': 'FileModificationDateTime',
    # '创建日期': 'FileCreationDateTime',
    # '创建时间': 'FileCreationDateTime',
    '照相机制造商': 'Make',
    '设备制造商': 'Make',
    '照相机型号': 'Model',
    '设备型号': 'Model',
    '作者': 'Artist',
    '版权所有者': 'Artist',
    '版权': 'Copyright',
    '版权ID': 'Copyright',
    # 'ISO速度': 'ISO',
    # '拍摄日期': 'CreateDate',
    # '测光模式': 'MeteringMode',
    # '光源': 'LightSource',
    # '闪光灯模式': 'Flash',
    # '白平衡': 'WhiteBalance',
    # '35mm焦距': 'FocalLengthIn35mmFormat',
    # '对比度': 'Contrast',
    # '饱和度': 'Saturation',
    # '清晰度': 'Sharpness',
    # '照相机序列号': 'SerialNumber',
    # '镜头制造商': 'LensMake',
    # '镜头型号': 'LensModel',
    '备注': 'XPComment',
    '主题': 'XPSubject',
    '标题': 'Title',
    '文件名称': 'Title',
    '标记': 'Subject',
    '颜色表示': 'ColorSpace',
    '色彩空间': 'ColorSpace',
    '拍摄时间': 'DateTimeOriginal',
    '修改时间': 'DateTimeOriginal',
    # '获取日期': 'DateAcquired',
    # '闪光灯型号': 'FlashModel',
    # '闪光灯制造商': 'FlashManufacturer',
    # '分级': 'Rating',
    # '宽度': 'ImageWidth',
    # '高度': 'ImageHeight',
    # '分辨率': 'ImageSize',
    # '图片格式': 'FileType',
}
maps_value = [v for k, v in maps.items()]


def write_metadata(file_path, metadata):
    """
    终极解决ExifTool中文乱码问题
    参数:
        file_path: 图片路径 (如 '001.jpg')
        metadata: 字典格式的元数据 (如 {'Title': '测试标题', 'Artist': '作者'})
    """
    metadata_format = {}
    for key, value in metadata.items():
        if key in maps:
            metadata_format[maps[key]] = value
        else:
            print(f'not map:{key}')
            if key in maps_value:
                metadata_format[key] = value
    project_dir = Path(__file__).resolve().parent
    executable = project_dir / "exiftool" / "exiftool.exe"
    if not executable.is_file():
        raise FileNotFoundError(
            f"ExifTool not found: {executable}. "
            "Place exiftool.exe and its exiftool_files directory under exiftool/."
        )

    target = Path(file_path).resolve()
    lines = ["-overwrite_original"]
    lines.extend(f"-{tag}={value}" for tag, value in metadata_format.items())
    lines.append(target.name)

    handle = tempfile.NamedTemporaryFile(
        prefix="metadata-", suffix=".args", dir=target.parent, delete=False
    )
    args_path = Path(handle.name)
    handle.close()
    try:
        args_path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
        result = subprocess.run(
            [
                str(executable),
                "-charset", "filename=UTF-8",
                "-charset", "iptc=UTF-8",
                "-charset", "exif=UTF-8",
                "-@", args_path.name,
            ],
            cwd=str(target.parent),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            shell=False,
            check=False,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
    finally:
        args_path.unlink(missing_ok=True)

    # 3. 打印执行信息（调试用）
    print(f"元数据已写入: {Path(file_path).name}")
    print("ExifTool输出:", result.stdout.strip())
    if result.stderr:
        print("错误信息:", result.stderr, file=sys.stderr)

    if result.returncode != 0:
        raise RuntimeError(f"ExifTool failed with exit code {result.returncode}: {result.stderr.strip()}")


if __name__ == '__main__':
    # 使用示例
    write_metadata('001.jpg', {
        # '标题': '测试标题.jpg',
        # '主题': '美图123',
        # '版权': '版权所有 © 20113',
        # '作者': '作者11',
        # '标记': '标记222',
        # '备注': '一个中文备注',
        # '镜头制造商': '制造商abc',
        # '饱和度': '高饱和度',
        'EXIF版本': f'11.21',
        # '颜色表示': 'sRGB',
    })
