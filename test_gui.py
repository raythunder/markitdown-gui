#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MarkItDown GUI 测试脚本
用于测试GUI程序是否能正常启动和运行
"""

import sys
import os
import tempfile
from pathlib import Path

def test_imports():
    """测试所需模块的导入"""
    print("测试模块导入...")
    
    try:
        import tkinter
        print("✓ tkinter导入成功")
    except ImportError as e:
        print(f"✗ tkinter导入失败: {e}")
        return False
    
    try:
        from markitdown import MarkItDown
        print("✓ markitdown导入成功")
    except ImportError as e:
        print(f"✗ markitdown导入失败: {e}")
        print("请运行: pip install markitdown")
        print("或安装完整版本: pip install 'markitdown[all]'")
        return False
    
    return True

def test_markitdown_basic():
    """测试MarkItDown基本功能"""
    print("\n测试MarkItDown基本功能...")
    
    try:
        from markitdown import MarkItDown
        
        # 创建MarkItDown实例
        md = MarkItDown()
        print("✓ MarkItDown实例创建成功")
        
        # 测试转换简单文本
        with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as f:
            f.write("这是一个测试文件\n包含中文内容")
            temp_file = f.name
        
        try:
            result = md.convert(temp_file)
            if result and result.text_content:
                print("✓ 文件转换成功")
                print(f"  转换结果长度: {len(result.text_content)} 字符")
            else:
                print("✗ 转换结果为空")
                return False
        finally:
            os.unlink(temp_file)
        
        return True
        
    except Exception as e:
        print(f"✗ MarkItDown测试失败: {e}")
        return False

def test_gui_import():
    """测试GUI模块导入"""
    print("\n测试GUI模块导入...")
    
    try:
        # 导入GUI脚本中的类
        sys.path.insert(0, os.path.dirname(__file__))
        
        # 这里我们只测试导入，不实际创建GUI
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, scrolledtext
        import threading
        
        print("✓ 所有GUI相关模块导入成功")
        return True
        
    except Exception as e:
        print(f"✗ GUI模块导入失败: {e}")
        return False

def test_file_structure():
    """测试文件结构"""
    print("\n检查文件结构...")
    
    required_files = [
        'markitdown_gui.py',
        'build_exe.py',
        'build.bat'
    ]
    
    all_good = True
    for file_path in required_files:
        if os.path.exists(file_path):
            print(f"✓ {file_path}")
        else:
            print(f"✗ {file_path} 不存在")
            all_good = False
    
    return all_good

def main():
    """主测试函数"""
    print("=" * 50)
    print("MarkItDown GUI 测试脚本")
    print("=" * 50)
    
    tests = [
        ("文件结构检查", test_file_structure),
        ("模块导入测试", test_imports),
        ("MarkItDown功能测试", test_markitdown_basic),
        ("GUI模块测试", test_gui_import),
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"\n{'='*20} {test_name} {'='*20}")
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"✗ {test_name}发生异常: {e}")
            results.append((test_name, False))
    
    # 总结
    print("\n" + "="*50)
    print("测试结果总结:")
    print("="*50)
    
    passed = 0
    total = len(results)
    
    for test_name, result in results:
        status = "✓ 通过" if result else "✗ 失败"
        print(f"{test_name}: {status}")
        if result:
            passed += 1
    
    print(f"\n总计: {passed}/{total} 项测试通过")
    
    if passed == total:
        print("\n🎉 所有测试通过！GUI程序应该可以正常运行。")
        print("您可以运行以下命令启动GUI:")
        print("  python markitdown_gui.py")
        print("\n或者构建exe文件:")
        print("  python build_exe.py")
        print("  (或双击 build.bat)")
    else:
        print(f"\n⚠️  有 {total-passed} 项测试失败，请解决这些问题后再运行GUI程序。")
    
    return passed == total

if __name__ == "__main__":
    success = main()
    
    if not success:
        print("\n按任意键退出...")
        try:
            input()
        except:
            pass
        sys.exit(1) 