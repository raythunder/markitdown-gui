#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MarkItDown GUI应用程序
将各种文件格式转换为Markdown的图形界面工具
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import traceback
from pathlib import Path
import webbrowser
import tempfile
import shutil
import zipfile
import re
import base64

try:
    from markitdown import MarkItDown
except ImportError as e:
    print(f"导入 markitdown 失败: {e}")
    print("请安装markitdown包: pip install markitdown")
    print("或者安装完整版本: pip install 'markitdown[all]'")
    
    # 如果在GUI环境中，显示错误对话框
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        messagebox.showerror("导入错误", 
                            "未找到 markitdown 包！\n\n"
                            "请先安装 markitdown:\n"
                            "pip install markitdown\n\n"
                            "或者安装完整版本:\n"
                            "pip install 'markitdown[all]'")
        root.destroy()
    except:
        pass
    
    sys.exit(1)

# 设置magika模型路径（用于打包后的exe文件）
def setup_magika_paths():
    """设置magika的路径和环境变量"""
    if getattr(sys, 'frozen', False):
        # 运行在PyInstaller打包的exe中
        application_path = sys._MEIPASS
        
        # 检查magika文件是否存在
        magika_config_path = os.path.join(application_path, 'magika', 'config')
        magika_models_path = os.path.join(application_path, 'magika', 'models')
        
        if os.path.exists(magika_config_path) and os.path.exists(magika_models_path):
            # 设置magika环境变量，让其找到配置文件
            os.environ['MAGIKA_CONFIG_PATH'] = magika_config_path
            os.environ['MAGIKA_MODELS_PATH'] = magika_models_path
            return True
    return False

def get_magika_model_path():
    """获取magika模型路径"""
    if getattr(sys, 'frozen', False):
        # 运行在PyInstaller打包的exe中
        application_path = sys._MEIPASS
        magika_models_path = os.path.join(application_path, 'magika', 'models', 'standard_v3_3')
        if os.path.exists(magika_models_path):
            return Path(magika_models_path)
    return None


class MarkItDownGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MarkItDown - 文件转Markdown工具")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 设置应用图标（如果有的话）
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # 初始化 MarkItDown
        self.markitdown = self._create_markitdown_instance()
        
        # 变量
        self.extracted_images = []  # 存储提取的图片信息
        self.source_file_path = ""  # 存储源文件路径
        
        # 创建界面
        self.create_widgets()
        
        # 绑定拖拽事件
        self.setup_drag_drop()
    
    def create_widgets(self):
        """创建GUI组件"""
        # 主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(4, weight=1)  # 结果显示区域可伸缩
        
        # 标题
        title_label = ttk.Label(self.main_frame, text="MarkItDown - 文件转Markdown工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.main_frame, text="文件选择", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="输入文件:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=60)
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=(0, 5))
        
        self.browse_button = ttk.Button(file_frame, text="浏览...", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, pady=(0, 5))
        
        # 拖拽提示
        drag_label = ttk.Label(file_frame, text="或者直接拖拽文件到此窗口", 
                              foreground="gray", font=("Arial", 9))
        drag_label.grid(row=1, column=0, columnspan=3, pady=(5, 0))
        
        # 选项设置区域
        options_frame = ttk.LabelFrame(self.main_frame, text="转换选项", padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 提取图片选项（针对epub文件）
        self.extract_images_var = tk.BooleanVar(value=True)
        self.extract_images_check = ttk.Checkbutton(
            options_frame, 
            text="从EPUB提取图片（保存到输出目录）",
            variable=self.extract_images_var
        )
        self.extract_images_check.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        # 图片处理方式选项
        self.image_mode_var = tk.StringVar(value="relative")
        ttk.Label(options_frame, text="图片引用方式:").grid(row=1, column=0, sticky=tk.W, pady=2)
        
        image_mode_frame = ttk.Frame(options_frame)
        image_mode_frame.grid(row=2, column=0, sticky=tk.W, pady=2)
        
        ttk.Radiobutton(
            image_mode_frame, 
            text="相对路径",
            variable=self.image_mode_var,
            value="relative"
        ).grid(row=0, column=0, padx=(0, 10))
        
        ttk.Radiobutton(
            image_mode_frame, 
            text="绝对路径",
            variable=self.image_mode_var,
            value="absolute"
        ).grid(row=0, column=1, padx=(0, 10))
        
        ttk.Radiobutton(
            image_mode_frame, 
            text="Base64编码",
            variable=self.image_mode_var,
            value="base64"
        ).grid(row=0, column=2)
        
        # 保留数据URI选项
        self.keep_data_uris_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="保留数据URI（如base64编码的图片）", 
                       variable=self.keep_data_uris_var).grid(row=3, column=0, sticky=tk.W)
        
        # 使用插件选项
        self.use_plugins_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="启用第三方插件", 
                       variable=self.use_plugins_var).grid(row=4, column=0, sticky=tk.W)
        
        # 操作按钮区域
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        self.convert_button = ttk.Button(button_frame, text="转换为Markdown", 
                                        command=self.convert_file)
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_button = ttk.Button(button_frame, text="保存结果", 
                                     command=self.save_result, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_button = ttk.Button(button_frame, text="清空", command=self.clear_all)
        self.clear_button.pack(side=tk.LEFT)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(self.main_frame, text="转换结果", padding="10")
        result_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 创建文本框和滚动条
        self.result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, 
                                                    height=20, width=80)
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏和进度条
        self.status_var = tk.StringVar(value="准备就绪")
        status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, 
                               relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 关于按钮
        about_button = ttk.Button(self.main_frame, text="关于", command=self.show_about)
        about_button.grid(row=5, column=2, sticky=tk.E, pady=(10, 0))
        
        # 进度条（初始隐藏）
        self.progress = ttk.Progressbar(self.main_frame, mode='indeterminate')
        # 不立即网格化，在需要时再显示
        
    def _create_markitdown_instance(self):
        """创建MarkItDown实例，处理打包后的magika模型路径"""
        try:
            # 在打包环境中，尝试特殊处理
            if getattr(sys, 'frozen', False):
                # 尝试不使用magika创建实例
                try:
                    markitdown_instance = MarkItDown()
                    # 如果magika出现问题，将其设为None来禁用文件类型检测
                    if hasattr(markitdown_instance, '_magika'):
                        try:
                            # 测试magika是否工作
                            markitdown_instance._magika.identify_bytes(b"test")
                        except:
                            print("magika出现问题，禁用文件类型检测")
                            markitdown_instance._magika = None
                    return markitdown_instance
                except Exception as e:
                    print(f"创建MarkItDown实例时出错: {e}")
                    # 最后的降级方案
                    return self._create_minimal_markitdown()
            else:
                # 非打包环境，使用默认设置
                return MarkItDown()
                
        except Exception as e:
            print(f"创建MarkItDown实例时出错: {e}")
            # 降级到最小实例
            return self._create_minimal_markitdown()
    
    def _create_minimal_markitdown(self):
        """创建最小化的MarkItDown实例"""
        try:
            # 导入MarkItDown内部类
            from markitdown._markitdown import MarkItDown as _MarkItDown
            
            # 创建一个不使用magika的实例
            instance = _MarkItDown()
            instance._magika = None  # 禁用magika
            
            return instance
        except:
            # 如果这也失败了，返回一个虚拟实例
            return None
    
    def _create_markitdown_instance_with_options(self):
        """创建MarkItDown实例，包含用户选项"""
        try:
            # 获取自定义magika模型路径
            custom_magika_path = get_magika_model_path()
            
            if custom_magika_path:
                # 创建自定义magika实例
                from magika import Magika
                custom_magika = Magika(model_dir=custom_magika_path)
                
                # 创建MarkItDown实例（包含插件设置）
                markitdown_instance = MarkItDown(enable_plugins=self.use_plugins_var.get())
                
                # 替换magika实例
                markitdown_instance._magika = custom_magika
                
                return markitdown_instance
            else:
                # 使用默认设置
                return MarkItDown(enable_plugins=self.use_plugins_var.get())
                
        except Exception as e:
            print(f"创建MarkItDown实例时出错: {e}")
            # 降级到默认实例
            return MarkItDown(enable_plugins=self.use_plugins_var.get())
    
    def setup_drag_drop(self):
        """设置拖拽支持"""
        # 简单的拖拽支持，绑定文件拖拽事件
        def handle_drop(event):
            files = self.root.tk.splitlist(event.data)
            if files:
                file_path = files[0]
                self.file_path_var.set(file_path)
                self.status_var.set(f"已选择文件: {os.path.basename(file_path)}")
        
        # Windows拖拽支持
        try:
            self.root.drop_target_register('DND_Files')
            self.root.dnd_bind('<<Drop>>', handle_drop)
        except:
            # 如果拖拽不支持，忽略错误
            pass
    
    def browse_file(self):
        """浏览文件对话框"""
        filetypes = [
            ("所有支持的文件", "*.pdf;*.docx;*.pptx;*.xlsx;*.xls;*.html;*.htm;*.txt;*.csv;*.json;*.xml;*.zip;*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.mp3;*.wav;*.epub;"),
            ("PDF文件", "*.pdf"),
            ("Word文档", "*.docx;*.doc"),
            ("PowerPoint演示文稿", "*.pptx;*.ppt"),
            ("Excel工作表", "*.xlsx;*.xls"),
            ("网页文件", "*.html;*.htm"),
            ("文本文件", "*.txt;*.csv;*.json;*.xml"),
            ("图像文件", "*.jpg;*.jpeg;*.png;*.gif;*.bmp"),
            ("音频文件", "*.mp3;*.wav;*.m4a"),
            ("电子书", "*.epub"),
            ("压缩文件", "*.zip"),
            ("所有文件", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="选择要转换的文件",
            filetypes=filetypes
        )
        
        if filename:
            self.file_path_var.set(filename)
            self.status_var.set(f"已选择文件: {os.path.basename(filename)}")
    
    def extract_epub_images(self, epub_path, output_dir):
        """从EPUB文件中提取图片"""
        images_extracted = []
        
        try:
            # 创建图片输出目录
            images_dir = os.path.join(output_dir, "images")
            os.makedirs(images_dir, exist_ok=True)
            
            # 打开EPUB文件（实际上是ZIP文件）
            with zipfile.ZipFile(epub_path, 'r') as epub_zip:
                # 获取所有文件列表
                file_list = epub_zip.namelist()
                
                # 查找图片文件
                image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp'}
                image_files = [f for f in file_list 
                             if os.path.splitext(f.lower())[1] in image_extensions]
                
                # 提取图片文件
                for image_file in image_files:
                    try:
                        # 获取文件名
                        filename = os.path.basename(image_file)
                        if not filename:  # 跳过目录
                            continue
                            
                        # 确保文件名唯一
                        counter = 1
                        base_name, ext = os.path.splitext(filename)
                        final_filename = filename
                        
                        while os.path.exists(os.path.join(images_dir, final_filename)):
                            final_filename = f"{base_name}_{counter}{ext}"
                            counter += 1
                        
                        # 提取文件
                        output_path = os.path.join(images_dir, final_filename)
                        with epub_zip.open(image_file) as source:
                            with open(output_path, 'wb') as target:
                                target.write(source.read())
                        
                        # 记录原始路径和新路径的映射
                        images_extracted.append({
                            'original_path': image_file,
                            'extracted_path': output_path,
                            'filename': final_filename,
                            'relative_path': f"images/{final_filename}"
                        })
                        
                    except Exception as e:
                        print(f"提取图片失败 {image_file}: {e}")
                        continue
                        
        except Exception as e:
            print(f"打开EPUB文件失败: {e}")
            return []
        
        return images_extracted
    
    def process_markdown_images(self, markdown_content, extracted_images, output_dir):
        """处理Markdown中的图片引用"""
        if not extracted_images:
            return markdown_content
        
        # 创建路径映射字典
        path_mapping = {}
        for img in extracted_images:
            # 尝试匹配不同的路径格式
            original_name = os.path.basename(img['original_path'])
            path_mapping[original_name] = img
            path_mapping[img['original_path']] = img
            # 去掉路径前缀的变体
            if '/' in img['original_path']:
                path_mapping[img['original_path'].split('/')[-1]] = img
        
        # 图片引用的正则表达式
        img_pattern = r'!\[([^\]]*)\]\(([^)]+)\)'
        
        def replace_image_ref(match):
            alt_text = match.group(1)
            img_src = match.group(2)
            
            # 查找匹配的提取图片
            found_img = None
            img_basename = os.path.basename(img_src)
            
            # 尝试多种匹配方式
            for key in [img_src, img_basename, img_src.replace('../', ''), img_src.replace('./', '')]:
                if key in path_mapping:
                    found_img = path_mapping[key]
                    break
            
            if found_img:
                # 根据选择的模式生成新的图片引用
                if self.image_mode_var.get() == "relative":
                    new_src = found_img['relative_path']
                elif self.image_mode_var.get() == "absolute":
                    new_src = os.path.abspath(found_img['extracted_path'])
                else:  # base64
                    try:
                        with open(found_img['extracted_path'], 'rb') as img_file:
                            img_data = img_file.read()
                            img_ext = os.path.splitext(found_img['filename'])[1].lower()
                            mime_type = {
                                '.jpg': 'image/jpeg',
                                '.jpeg': 'image/jpeg',
                                '.png': 'image/png',
                                '.gif': 'image/gif',
                                '.bmp': 'image/bmp',
                                '.svg': 'image/svg+xml'
                            }.get(img_ext, 'image/jpeg')
                            
                            base64_data = base64.b64encode(img_data).decode()
                            new_src = f"data:{mime_type};base64,{base64_data}"
                    except Exception as e:
                        print(f"Base64编码图片失败: {e}")
                        new_src = found_img['relative_path']
                
                return f"![{alt_text}]({new_src})"
            else:
                # 未找到匹配的图片，保持原样
                return match.group(0)
        
        # 替换所有图片引用
        processed_content = re.sub(img_pattern, replace_image_ref, markdown_content)
        
        return processed_content

    def convert_file(self):
        """转换文件"""
        if not self.file_path_var.get():
            messagebox.showerror("错误", "请先选择要转换的文件")
            return
        
        # 禁用转换按钮
        self.convert_button.config(state=tk.DISABLED)
        self.progress.start()
        
        # 在新线程中进行转换
        thread = threading.Thread(target=self._convert_worker)
        thread.daemon = True
        thread.start()
    
    def _convert_worker(self):
        """转换工作线程"""
        try:
            md = self.markitdown
            input_path = self.file_path_var.get()
            
            # 更新状态
            self.root.after(0, lambda: self.status_var.set("正在转换文件..."))
            
            # 转换文件
            result = md.convert(input_path)
            markdown_content = result.text_content
            
            # 检查是否是EPUB文件且需要提取图片
            is_epub = input_path.lower().endswith('.epub')
            extract_images = self.extract_images_var.get()
            
            if is_epub and extract_images:
                self.root.after(0, lambda: self.status_var.set("正在提取图片..."))
                
                # 确定输出目录
                output_dir = os.path.dirname(input_path)
                
                # 提取图片
                extracted_images = self.extract_epub_images(input_path, output_dir)
                
                # 保存提取的图片信息和源文件路径
                self.extracted_images = extracted_images
                self.source_file_path = input_path
                
                if extracted_images:
                    self.root.after(0, lambda: self.status_var.set("正在处理图片引用..."))
                    
                    # 处理Markdown中的图片引用
                    markdown_content = self.process_markdown_images(markdown_content, extracted_images, output_dir)
                    
                    # 更新状态信息
                    img_count = len(extracted_images)
                    status_msg = f"转换完成！提取了 {img_count} 张图片"
                else:
                    status_msg = "转换完成！未找到图片文件"
            else:
                # 清空之前的图片信息
                self.extracted_images = []
                self.source_file_path = input_path
                status_msg = "转换完成！"
            
            # 更新结果显示
            def update_result():
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(1.0, markdown_content)
                self.status_var.set(status_msg)
                self.progress.stop()
                self.convert_button.config(state=tk.NORMAL)
                self.save_button.config(state=tk.NORMAL)  # 启用保存按钮
            
            self.root.after(0, update_result)
            
        except Exception as e:
            error_msg = f"转换失败: {str(e)}"
            print(f"转换错误: {traceback.format_exc()}")
            
            def show_error():
                self.status_var.set(error_msg)
                self.progress.stop()
                self.convert_button.config(state=tk.NORMAL)
                messagebox.showerror("转换错误", error_msg)
            
            self.root.after(0, show_error)
    
    def save_result(self):
        """保存转换结果"""
        content = self.result_text.get(1.0, tk.END).strip()
        
        if not content:
            messagebox.showwarning("警告", "没有内容可保存！")
            return
        
        # 获取建议的文件名
        input_file = self.file_path_var.get()
        if input_file:
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            suggested_name = f"{base_name}.md"
        else:
            suggested_name = "converted.md"
        
        # 保存文件对话框
        filename = filedialog.asksaveasfilename(
            title="保存Markdown文件",
            defaultextension=".md",
            initialfile=suggested_name,
            filetypes=[("Markdown文件", "*.md"), ("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if filename:
            try:
                # 保存Markdown文件
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                saved_files = [filename]
                
                # 如果有提取的图片，复制图片到新的输出目录
                if self.extracted_images:
                    output_dir = os.path.dirname(filename)
                    images_target_dir = os.path.join(output_dir, "images")
                    
                    # 创建图片目录
                    os.makedirs(images_target_dir, exist_ok=True)
                    
                    # 复制图片文件
                    copied_images = []
                    for img_info in self.extracted_images:
                        src_path = img_info['extracted_path']
                        filename_only = os.path.basename(src_path)
                        target_path = os.path.join(images_target_dir, filename_only)
                        
                        # 如果源文件和目标文件不是同一个，才复制
                        if os.path.abspath(src_path) != os.path.abspath(target_path):
                            try:
                                import shutil
                                shutil.copy2(src_path, target_path)
                                copied_images.append(filename_only)
                                saved_files.append(target_path)
                            except Exception as copy_error:
                                print(f"复制图片失败 {src_path}: {copy_error}")
                    
                    if copied_images:
                        status_msg = f"已保存: {os.path.basename(filename)} 和 {len(copied_images)} 张图片"
                        info_msg = f"文件已保存到:\n{filename}\n\n图片已保存到:\n{images_target_dir}\n\n复制的图片: {', '.join(copied_images)}"
                    else:
                        status_msg = f"已保存: {os.path.basename(filename)} (图片已在目标目录)"
                        info_msg = f"文件已保存到:\n{filename}\n\n图片目录: {images_target_dir}"
                else:
                    status_msg = f"已保存: {os.path.basename(filename)}"
                    info_msg = f"Markdown文件已保存到:\n{filename}"
                
                self.status_var.set(status_msg)
                messagebox.showinfo("保存成功", info_msg)
                
                # 询问是否打开文件
                if messagebox.askyesno("打开文件", "是否要打开保存的文件？"):
                    self.open_file(filename)
                    
            except Exception as e:
                error_msg = f"保存文件失败: {str(e)}"
                self.status_var.set(error_msg)
                messagebox.showerror("保存错误", error_msg)
    
    def open_file(self, filename):
        """打开文件"""
        try:
            if sys.platform == 'win32':
                os.startfile(filename)
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{filename}"')
            else:  # Linux
                os.system(f'xdg-open "{filename}"')
        except Exception as e:
            messagebox.showerror("打开文件失败", f"无法打开文件: {str(e)}")
    
    def clear_all(self):
        """清空所有内容"""
        self.file_path_var.set("")
        self.result_text.delete(1.0, tk.END)
        self.save_button.config(state=tk.DISABLED)
        self.status_var.set("准备就绪")
        
        # 清空图片信息
        self.extracted_images = []
        self.source_file_path = ""
    
    def show_about(self):
        """显示关于对话框"""
        about_text = """MarkItDown GUI v1.0

这是一个基于Microsoft MarkItDown的图形界面工具，
用于将各种文件格式转换为Markdown格式。

支持的文件格式:
• PDF文档
• Word文档 (.docx, .doc)
• PowerPoint演示文稿 (.pptx, .ppt)
• Excel工作表 (.xlsx, .xls)
• 网页文件 (.html, .htm)
• 文本文件 (.txt, .csv, .json, .xml)
• 图像文件 (.jpg, .png, .gif, .bmp)
• 音频文件 (.mp3, .wav, .m4a)
• 电子书 (.epub)
• 压缩文件 (.zip)
• YouTube链接

基于: Microsoft MarkItDown
项目地址: https://github.com/microsoft/markitdown

开发者: AI Assistant
版本: 1.0.0"""
        
        # 创建关于对话框
        about_window = tk.Toplevel(self.root)
        about_window.title("关于 MarkItDown GUI")
        about_window.geometry("500x400")
        about_window.resizable(False, False)
        
        # 居中显示
        about_window.transient(self.root)
        about_window.grab_set()
        
        # 内容框架
        frame = ttk.Frame(about_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(frame, text="MarkItDown GUI", 
                               font=("Arial", 18, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 内容
        text_widget = tk.Text(frame, wrap=tk.WORD, height=15, width=50)
        text_widget.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        text_widget.insert(1.0, about_text)
        text_widget.config(state=tk.DISABLED)
        
        # 按钮框架
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        # 访问项目按钮
        def open_github():
            webbrowser.open("https://github.com/microsoft/markitdown")
        
        github_button = ttk.Button(button_frame, text="访问GitHub项目", command=open_github)
        github_button.pack(side=tk.LEFT)
        
        # 关闭按钮
        close_button = ttk.Button(button_frame, text="关闭", command=about_window.destroy)
        close_button.pack(side=tk.RIGHT)


def main():
    """主函数"""
    try:
        # 设置magika路径（对于打包后的exe）
        setup_magika_paths()
        
        # 创建主窗口
        root = tk.Tk()
        
        # 创建应用程序
        app = MarkItDownGUI(root)
        
        # 运行应用程序
        root.mainloop()
        
    except Exception as e:
        # 如果GUI初始化失败，显示错误消息
        print(f"应用程序启动失败: {e}")
        traceback.print_exc()
        
        # 尝试显示错误对话框
        try:
            root = tk.Tk()
            root.withdraw()  # 隐藏主窗口
            messagebox.showerror("启动错误", f"应用程序启动失败:\n{str(e)}\n\n请检查是否正确安装了所有依赖。")
        except:
            pass


if __name__ == "__main__":
    main() 