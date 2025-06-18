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
        
        # 创建界面
        self.create_widgets()
        
        # 绑定拖拽事件
        self.setup_drag_drop()
    
    def create_widgets(self):
        """创建GUI组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="MarkItDown - 文件转Markdown工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
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
        options_frame = ttk.LabelFrame(main_frame, text="转换选项", padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 保留数据URI选项
        self.keep_data_uris_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="保留数据URI（如base64编码的图片）", 
                       variable=self.keep_data_uris_var).grid(row=0, column=0, sticky=tk.W)
        
        # 使用插件选项
        self.use_plugins_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="启用第三方插件", 
                       variable=self.use_plugins_var).grid(row=1, column=0, sticky=tk.W)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
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
        result_frame = ttk.LabelFrame(main_frame, text="转换结果", padding="10")
        result_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 创建文本框和滚动条
        self.result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, 
                                                    height=20, width=80)
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 关于按钮
        about_button = ttk.Button(main_frame, text="关于", command=self.show_about)
        about_button.grid(row=5, column=2, sticky=tk.E, pady=(10, 0))
        
        # 进度条（初始隐藏）
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        
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
    
    def convert_file(self):
        """转换文件为Markdown"""
        file_path = self.file_path_var.get().strip()
        
        if not file_path:
            messagebox.showerror("错误", "请先选择要转换的文件！")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("错误", "选择的文件不存在！")
            return
        
        # 禁用转换按钮，显示进度条
        self.convert_button.config(state=tk.DISABLED)
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        self.progress.start()
        
        # 在线程中执行转换
        threading.Thread(target=self._convert_thread, args=(file_path,), daemon=True).start()
    
    def _convert_thread(self, file_path):
        """在线程中执行文件转换"""
        try:
            self.root.after(0, lambda: self.status_var.set("正在转换文件..."))
            
            # 重新初始化 MarkItDown（考虑插件设置）
            self.markitdown = self._create_markitdown_instance_with_options()
            
            # 执行转换
            result = self.markitdown.convert(file_path)
            
            # 在主线程中更新UI
            self.root.after(0, lambda: self._conversion_complete(result, None))
            
        except Exception as e:
            # 在主线程中显示错误
            self.root.after(0, lambda: self._conversion_complete(None, e))
    
    def _conversion_complete(self, result, error):
        """转换完成后的处理"""
        # 停止进度条，隐藏进度条
        self.progress.stop()
        self.progress.grid_remove()
        
        # 重新启用转换按钮
        self.convert_button.config(state=tk.NORMAL)
        
        if error:
            # 显示错误
            error_msg = f"转换失败: {str(error)}"
            self.status_var.set(error_msg)
            messagebox.showerror("转换错误", f"{error_msg}\n\n详细信息:\n{traceback.format_exc()}")
            return
        
        if result and result.text_content:
            # 显示转换结果
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, result.text_content)
            
            # 启用保存按钮
            self.save_button.config(state=tk.NORMAL)
            
            # 更新状态
            file_name = os.path.basename(self.file_path_var.get())
            self.status_var.set(f"转换完成: {file_name} -> Markdown")
            
            # 显示成功消息
            messagebox.showinfo("转换成功", f"文件已成功转换为Markdown格式！\n\n字符数: {len(result.text_content)}")
        else:
            self.status_var.set("转换完成，但没有生成内容")
            messagebox.showwarning("转换完成", "文件转换完成，但没有生成任何内容。")
    
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
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                self.status_var.set(f"结果已保存到: {os.path.basename(filename)}")
                messagebox.showinfo("保存成功", f"Markdown文件已保存到:\n{filename}")
                
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
        self.status_var.set("就绪")
    
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
• Jupyter笔记本 (.ipynb)
• Outlook邮件 (.msg)
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