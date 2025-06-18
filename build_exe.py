import os
import sys
import subprocess
import shutil
from pathlib import Path

def install_requirements():
    """安装必需的依赖"""
    print("正在安装必需的依赖...")
    
    requirements = [
        "pyinstaller",
        "tkinter",  # 通常预装
    ]
    
    # 安装markitdown及其依赖
    markitdown_path = Path("packages/markitdown")
    if markitdown_path.exists():
        print("安装markitdown...")
        subprocess.run([sys.executable, "-m", "pip", "install", "-e", str(markitdown_path / "[all]")], check=True)
    
    # 安装PyInstaller
    for req in requirements:
        if req != "tkinter":  # tkinter通常预装
            print(f"安装 {req}...")
            subprocess.run([sys.executable, "-m", "pip", "install", req], check=True)

def create_spec_file():
    """创建PyInstaller规范文件"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from pathlib import Path
import site

# 获取markitdown包的路径
markitdown_src = os.path.join(os.path.dirname(SPECPATH), 'packages', 'markitdown', 'src')

# 获取magika模型文件路径
try:
    import magika
    magika_package_path = os.path.dirname(magika.__file__)
    magika_models = os.path.join(magika_package_path, 'models')
    magika_config = os.path.join(magika_package_path, 'config')
except ImportError:
    magika_models = None
    magika_config = None

# 准备magika数据文件
magika_datas = []
if magika_models and os.path.exists(magika_models):
    magika_datas.append((magika_models, 'magika/models'))
if magika_config and os.path.exists(magika_config):
    magika_datas.append((magika_config, 'magika/config'))

block_cipher = None

a = Analysis(
    ['markitdown_gui.py'],
    pathex=['.', markitdown_src],
    binaries=[],
    datas=magika_datas,
    hiddenimports=[
        'markitdown',
        'markitdown._markitdown',
        'markitdown.converters',
        'markitdown.converters._pdf_converter',
        'markitdown.converters._docx_converter',
        'markitdown.converters._xlsx_converter',
        'markitdown.converters._pptx_converter',
        'markitdown.converters._html_converter',
        'markitdown.converters._image_converter',
        'markitdown.converters._audio_converter',
        'markitdown.converters._plain_text_converter',
        'markitdown.converters._csv_converter',
        'markitdown.converters._epub_converter',
        'markitdown.converters._zip_converter',
        'markitdown.converters._ipynb_converter',
        'markitdown.converters._rss_converter',
        'markitdown.converters._wikipedia_converter',
        'markitdown.converters._youtube_converter',
        'markitdown.converters._bing_serp_converter',
        'markitdown.converters._outlook_msg_converter',
        'markitdown.converters._doc_intel_converter',
        'markitdown.converter_utils',
        'markitdown.converter_utils.docx',
        'markitdown.converter_utils.docx.math',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'threading',
        'requests',
        'magika',
        'magika.magika',
        'magika.types',
        'magika.config',
        'onnxruntime',
        'onnxruntime.capi.onnxruntime_pybind11_state',
        'numpy',
        'charset_normalizer',
        'codecs',
        'mimetypes',
        'xml.etree.ElementTree',
        'zipfile',
        'json',
        'csv',
        'base64',
        'urllib.parse',
        'urllib.request',
        'html.parser',
        'bs4',
        'pandas',
        'openpyxl',
        'xlrd',
        'python-pptx',
        'pypdf',
        'PIL',
        'eyed3',
        'pydub',
        'ebooklib',
        'python-docx',
        'mammoth',
        'lxml',
        'beautifulsoup4',
        'feedparser',
        'youtube-dl',
        'yt-dlp',
        'extract-msg',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MarkItDown-GUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
'''
    
    with open('markitdown_gui.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("已创建PyInstaller规范文件: markitdown_gui.spec")

def create_version_info():
    """创建版本信息文件"""
    version_info = '''# UTF-8
#
# For more details about fixed file info 'ffi' see:
# http://msdn.microsoft.com/en-us/library/ms646997.aspx
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1,0,0,0),
    prodvers=(1,0,0,0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'MarkItDown GUI'),
        StringStruct(u'FileDescription', u'MarkItDown文件转换工具'),
        StringStruct(u'FileVersion', u'1.0.0.0'),
        StringStruct(u'InternalName', u'MarkItDown-GUI'),
        StringStruct(u'LegalCopyright', u'Based on Microsoft MarkItDown'),
        StringStruct(u'OriginalFilename', u'MarkItDown-GUI.exe'),
        StringStruct(u'ProductName', u'MarkItDown GUI'),
        StringStruct(u'ProductVersion', u'1.0.0.0')])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
'''
    
    with open('version_info.txt', 'w', encoding='utf-8') as f:
        f.write(version_info)
    
    print("已创建版本信息文件: version_info.txt")

def create_icon():
    """创建或复制图标文件"""
    # 如果没有图标文件，创建一个简单的说明
    if not os.path.exists('icon.ico'):
        print("提示: 如果您有图标文件(.ico格式)，请将其命名为'icon.ico'并放在当前目录中")
        print("      这样生成的exe文件就会有自定义图标")

def build_exe():
    """构建exe文件"""
    print("开始构建exe文件...")
    
    # 使用PyInstaller构建
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "markitdown_gui.spec",
        "--clean",
        "--noconfirm"
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("构建成功！")
        print("exe文件位置: dist/MarkItDown-GUI.exe")
        
        # 检查文件大小
        exe_path = Path("dist/MarkItDown-GUI.exe")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"文件大小: {size_mb:.1f} MB")
        
        return True
    else:
        print("构建失败！")
        print("错误输出:")
        print(result.stderr)
        return False

def cleanup():
    """清理临时文件"""
    print("清理临时文件...")
    
    dirs_to_remove = ['build', '__pycache__']
    files_to_remove = ['markitdown_gui.spec', 'version_info.txt']
    
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"已删除目录: {dir_name}")
    
    for file_name in files_to_remove:
        if os.path.exists(file_name):
            os.remove(file_name)
            print(f"已删除文件: {file_name}")

def main():
    """主函数"""
    print("=" * 50)
    print("MarkItDown GUI - exe构建脚本")
    print("=" * 50)
    
    try:
        # 步骤1: 安装依赖
        print("\n步骤1: 安装依赖")
        install_requirements()
        
        # 步骤2: 创建配置文件
        print("\n步骤2: 创建配置文件")
        create_spec_file()
        create_version_info()
        create_icon()
        
        # 步骤3: 构建exe
        print("\n步骤3: 构建exe文件")
        success = build_exe()
        
        if success:
            print("\n" + "=" * 50)
            print("构建完成！")
            print("exe文件位置: dist/MarkItDown-GUI.exe")
            print("您可以运行该exe文件来使用MarkItDown GUI")
            print("=" * 50)
            
            # 询问是否清理临时文件
            response = input("\n是否清理临时文件？(y/n): ").lower().strip()
            if response in ['y', 'yes', '是']:
                cleanup()
        else:
            print("\n构建失败，请检查错误信息")
            
    except Exception as e:
        print(f"\n构建过程中发生错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 