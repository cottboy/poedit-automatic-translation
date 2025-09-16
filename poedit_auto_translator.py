#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Poedit自动翻译工具
通过模拟鼠标点击和键盘操作实现自动翻译
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyautogui
import pyperclip
import time
import threading
import json
import os
import tempfile
import gc
from typing import Dict, Optional, Tuple
import psutil
import keyboard


class ClipboardMonitor:
    """剪贴板监控类，优化内存使用"""
    
    def __init__(self):
        self.last_content = ""
        self.temp_file = None
        self._setup_temp_file()
    
    def _setup_temp_file(self):
        """设置临时文件用于存储剪贴板历史"""
        temp_dir = tempfile.gettempdir()
        self.temp_file = os.path.join(temp_dir, "poedit_clipboard_history.json")
    
    def get_clipboard_content(self) -> str:
        """获取剪贴板内容，去除首尾空格"""
        try:
            content = pyperclip.paste().strip()
            return content
        except Exception:
            return ""
    
    def save_content_to_temp(self, content: str, content_type: str):
        """保存内容到临时文件"""
        try:
            data = {"type": content_type, "content": content, "timestamp": time.time()}
            with open(self.temp_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception as e:
            print(f"保存临时文件失败: {e}")
    
    def load_content_from_temp(self) -> Optional[Dict]:
        """从临时文件加载内容"""
        try:
            if os.path.exists(self.temp_file):
                with open(self.temp_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"读取临时文件失败: {e}")
        return None
    
    def cleanup_temp_file(self):
        """清理临时文件"""
        try:
            if os.path.exists(self.temp_file):
                os.remove(self.temp_file)
        except Exception as e:
            print(f"清理临时文件失败: {e}")
    
    def is_content_changed(self, new_content: str) -> bool:
        """检查内容是否发生变化"""
        if not new_content or new_content.isspace():
            return False
        return new_content != self.last_content
    
    def update_last_content(self, content: str):
        """更新最后的内容"""
        self.last_content = content
        # 强制垃圾回收
        gc.collect()


class CoordinateSelector:
    """坐标选择器"""
    
    def __init__(self, parent, callback):
        self.parent = parent
        self.callback = callback
        self.selecting = False
    
    def start_selection(self):
        """开始坐标选择"""
        self.selecting = True
        self.parent.withdraw()  # 隐藏主窗口
        
        # 创建全屏透明窗口
        self.overlay = tk.Toplevel()
        self.overlay.attributes('-fullscreen', True)
        self.overlay.attributes('-alpha', 0.3)
        self.overlay.configure(bg='red')
        self.overlay.bind('<Button-1>', self.on_click)
        
        # 显示提示
        label = tk.Label(self.overlay, text="点击目标位置选择坐标", 
                        font=('Arial', 20), bg='red', fg='white')
        label.pack(expand=True)
    
    def on_click(self, event):
        """处理点击事件"""
        x, y = event.x_root, event.y_root
        self.overlay.destroy()
        self.parent.deiconify()  # 显示主窗口
        self.callback(x, y)
        self.selecting = False


class PoeditAutoTranslator:
    """Poedit自动翻译主类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Poedit自动翻译")
        self.root.geometry("600x700")
        
        # 配置文件路径
        self.config_dir = self.get_config_directory()
        self.config_file = os.path.join(self.config_dir, "poedit-automatic-translation_config.json")
        
        # 坐标配置
        self.coordinates = {
            'poedit_source': None,      # Poedit原文框
            'poedit_target': None,      # Poedit译文框
            'service_source': None,     # 翻译服务原文框
            'service_copy_button': None, # 翻译服务复制按钮
            'service_result_box': None,  # 翻译服务译文框（用于全选复制方案）
            'scroll_gesture_position': None  # 执行鼠标手势的位置
        }
        
        # 设置选项
        self.skip_translated = tk.BooleanVar(value=True)  # 跳过已翻译
        self.translation_wait_time = tk.IntVar(value=3000)  # 翻译结果等待时间(毫秒)
        
        # 新增的翻译检测配置项
        self.check_translation_consistency = tk.BooleanVar(value=True)  # 是否检测翻译原文与复制的译文是否一致
        self.check_interval = tk.IntVar(value=500)  # 检测时间间隔(毫秒)
        self.check_timeout_count = tk.IntVar(value=20)  # 检测超时次数
        
        # 新增：翻译服务译文复制方式（0=点击复制按钮，1=全选复制，2=双击复制，3=三击复制）
        self.copy_method = tk.IntVar(value=0)
        
        # 新增：鼠标手势滚动到底部相关设置
        self.use_scroll_gesture = tk.BooleanVar(value=False)  # 复制翻译结果前使用鼠标手势滚动到底部
        self.scroll_gesture_wait_time = tk.IntVar(value=500)  # 执行滚动手势后等待时间(毫秒)
        
        # 新增：换行转换设置
        self.convert_newlines = tk.BooleanVar(value=False)  # 将换行转换为__NL_114514__无意义字符
        
        # 快捷键设置
        self.start_hotkey_combination = tk.StringVar(value="F9")  # 开始快捷键，默认F9
        self.stop_hotkey_combination = tk.StringVar(value="F10")  # 停止快捷键，默认F10
        self.is_binding_key = False  # 是否正在绑定按键
        self.binding_key_type = None  # 正在绑定的快捷键类型（'start' 或 'stop'）
        
        # 运行状态
        self.is_running = False
        self.translation_thread = None
        self.same_source_count = 0  # 连续相同原文计数器
        self.last_source_text = ""  # 上次的原文
        self.last_translated_text = ""  # 上次的译文
        
        # 剪贴板监控器
        self.clipboard_monitor = ClipboardMonitor()
        
        # 内存监控
        self.process = psutil.Process()
        
        # 设置快捷键监听（在界面创建完成后）
        self.root.after(100, self.setup_hotkey_listener)
        
        self.setup_ui()
        self.load_config()
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 设置选项区域（合并坐标设置和其他选项）
        settings_frame = ttk.LabelFrame(main_frame, text="设置选项", padding="10")
        settings_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 坐标设置部分
        coord_labels = [
            ("Poedit原文框:", 'poedit_source'),
            ("Poedit译文框:", 'poedit_target'),
            ("翻译服务原文框:", 'service_source'),
            ("翻译服务复制按钮:", 'service_copy_button'),
            ("翻译服务译文框:", 'service_result_box'),
            ("执行鼠标手势的位置:", 'scroll_gesture_position')
        ]
        
        self.coord_vars = {}
        for i, (label_text, key) in enumerate(coord_labels):
            ttk.Label(settings_frame, text=label_text).grid(row=i, column=0, sticky=tk.W, pady=2)
            
            var = tk.StringVar(value="未设置")
            self.coord_vars[key] = var
            ttk.Label(settings_frame, textvariable=var, width=20).grid(row=i, column=1, padx=(10, 5), pady=2)
            
            ttk.Button(settings_frame, text="选择", 
                      command=lambda k=key: self.select_coordinate(k)).grid(row=i, column=2, pady=2)
        
        # 计算后续控件的起始行
        base_row = len(coord_labels)
        
        # 快捷键绑定设置（放在翻译服务复制按钮设置项下方）
        ttk.Label(settings_frame, text="快捷键设置:").grid(row=base_row, column=0, sticky=tk.W, pady=2)
        
        # 快捷键设置框架
        hotkey_frame = ttk.Frame(settings_frame)
        hotkey_frame.grid(row=base_row, column=1, columnspan=2, sticky=tk.W, padx=(10, 0), pady=2)
        
        # 开始快捷键（左边）
        ttk.Label(hotkey_frame, text="开始:").pack(side=tk.LEFT)
        self.start_hotkey_display = ttk.Label(hotkey_frame, text=self.start_hotkey_combination.get(), 
                                             relief="sunken", width=11, cursor="hand2")
        self.start_hotkey_display.pack(side=tk.LEFT, padx=(5, 5))
        self.start_hotkey_display.bind("<Button-1>", lambda e: self.start_key_binding('start'))
        
        # 停止快捷键（右边）
        ttk.Label(hotkey_frame, text="停止:").pack(side=tk.LEFT)
        self.stop_hotkey_display = ttk.Label(hotkey_frame, text=self.stop_hotkey_combination.get(), 
                                            relief="sunken", width=11, cursor="hand2")
        self.stop_hotkey_display.pack(side=tk.LEFT, padx=(5, 0))
        self.stop_hotkey_display.bind("<Button-1>", lambda e: self.start_key_binding('stop'))
        
        # 新增：翻译服务译文复制方式单选（四种方式：复制按钮、全选复制、双击复制、三击复制）
        copy_method_frame = ttk.Frame(settings_frame)
        copy_method_label = ttk.Label(copy_method_frame, text="翻译服务译文复制方式:")
        copy_method_label.pack(side=tk.LEFT)
        ttk.Radiobutton(copy_method_frame, text="复制按钮", variable=self.copy_method, value=0).pack(side=tk.LEFT, padx=(110, 5))
        ttk.Radiobutton(copy_method_frame, text="全选复制", variable=self.copy_method, value=1).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Radiobutton(copy_method_frame, text="双击复制", variable=self.copy_method, value=2).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Radiobutton(copy_method_frame, text="三击复制", variable=self.copy_method, value=3).pack(side=tk.LEFT)
        copy_method_frame.grid(row=base_row+1, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        # 新增：复制翻译结果前使用鼠标手势滚动到底部
        ttk.Checkbutton(settings_frame, text="复制翻译结果前使用鼠标手势滚动到底部", 
                       variable=self.use_scroll_gesture).grid(row=base_row+2, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # 新增：执行滚动到底部鼠标手势后等待时间
        ttk.Label(settings_frame, text="执行滚动到底部鼠标手势后等待时间(毫秒):").grid(row=base_row+3, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(settings_frame, from_=100, to=3000, textvariable=self.scroll_gesture_wait_time, 
                   width=10).grid(row=base_row+3, column=1, padx=(10, 0), pady=2)
        
        # 其他选项设置
        ttk.Checkbutton(settings_frame, text="跳过已翻译的字段", 
                       variable=self.skip_translated).grid(row=base_row+4, column=0, sticky=tk.W, pady=2)
        
        ttk.Label(settings_frame, text="翻译结果等待时间(毫秒):").grid(row=base_row+5, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(settings_frame, from_=100, to=5000, textvariable=self.translation_wait_time, 
                   width=10).grid(row=base_row+5, column=1, padx=(10, 0), pady=2)
        
        # 新增的翻译检测设置项
        ttk.Checkbutton(settings_frame, text="检测翻译原文与复制的译文是否一致", 
                       variable=self.check_translation_consistency).grid(row=base_row+6, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        ttk.Label(settings_frame, text="检测翻译原文与复制的译文是否一致时间间隔(毫秒):").grid(row=base_row+7, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(settings_frame, from_=100, to=10000, textvariable=self.check_interval, 
                   width=10).grid(row=base_row+7, column=1, padx=(10, 0), pady=2)
        
        ttk.Label(settings_frame, text="检测翻译原文与复制的译文是否一致的超时次数:").grid(row=base_row+8, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(settings_frame, from_=1, to=100, textvariable=self.check_timeout_count, 
                   width=10).grid(row=base_row+8, column=1, padx=(10, 0), pady=2)
        
        # 新增：换行转换选项
        ttk.Checkbutton(settings_frame, text="将换行转换为__NL_114514__无意义字符", 
                       variable=self.convert_newlines).grid(row=base_row+9, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        
        # 控制按钮区域
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
        self.start_button = ttk.Button(control_frame, text="开始翻译", 
                                      command=self.start_translation)
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.stop_button = ttk.Button(control_frame, text="停止翻译", 
                                     command=self.stop_translation, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="保存配置", 
                  command=self.save_config).pack(side=tk.LEFT, padx=5)
        
        # 状态显示区域
        status_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="10")
        status_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.status_text = tk.Text(status_frame, height=15, width=60, state='disabled')
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 开始内存监控（后台运行，不显示）
        self.update_memory_usage()
    
    def select_coordinate(self, coord_type: str):
        """选择坐标"""
        def on_coordinate_selected(x, y):
            self.coordinates[coord_type] = (x, y)
            self.coord_vars[coord_type].set(f"({x}, {y})")
            self.log_status(f"已设置{coord_type}坐标: ({x}, {y})")
        
        selector = CoordinateSelector(self.root, on_coordinate_selected)
        selector.start_selection()
    
    def log_status(self, message: str):
        """记录状态信息"""
        timestamp = time.strftime("%H:%M:%S")
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')
        self.root.update_idletasks()
    
    def setup_hotkey_listener(self):
        """设置快捷键监听"""
        try:
            import keyboard
            # 清理之前的全局快捷键钩子
            if hasattr(self, 'hotkey_hook') and self.hotkey_hook:
                try:
                    keyboard.unhook(self.hotkey_hook)
                except Exception:
                    pass
                self.hotkey_hook = None

            # 解析开始快捷键组合
            start_combo = self.start_hotkey_combination.get().strip().lower()
            start_parts = [p.strip() for p in start_combo.split('+') if p.strip()]
            start_modifiers = {p for p in start_parts if p in ('ctrl', 'alt', 'shift')}
            start_main_key = next((p for p in reversed(start_parts) if p not in ('ctrl', 'alt', 'shift')), None)
            
            # 解析停止快捷键组合
            stop_combo = self.stop_hotkey_combination.get().strip().lower()
            stop_parts = [p.strip() for p in stop_combo.split('+') if p.strip()]
            stop_modifiers = {p for p in stop_parts if p in ('ctrl', 'alt', 'shift')}
            stop_main_key = next((p for p in reversed(stop_parts) if p not in ('ctrl', 'alt', 'shift')), None)

            def handler(event):
                try:
                    # 绑定按键期间不触发快捷键
                    if getattr(self, 'is_binding_key', False):
                        return
                    # 仅在按下事件时判断
                    if getattr(event, 'event_type', '') != 'down':
                        return
                    
                    event_key = getattr(event, 'name', '')
                    
                    # 检查开始快捷键
                    if start_main_key and event_key == start_main_key:
                        # 检查修饰键状态
                        modifiers_match = True
                        for m in start_modifiers:
                            if not keyboard.is_pressed(m):
                                modifiers_match = False
                                break
                        # 检查是否有多余的修饰键被按下
                        for m in ('ctrl', 'alt', 'shift'):
                            if m not in start_modifiers and keyboard.is_pressed(m):
                                modifiers_match = False
                                break
                        if modifiers_match:
                            self.trigger_start_translation()
                            return
                    
                    # 检查停止快捷键
                    if stop_main_key and event_key == stop_main_key:
                        # 检查修饰键状态
                        modifiers_match = True
                        for m in stop_modifiers:
                            if not keyboard.is_pressed(m):
                                modifiers_match = False
                                break
                        # 检查是否有多余的修饰键被按下
                        for m in ('ctrl', 'alt', 'shift'):
                            if m not in stop_modifiers and keyboard.is_pressed(m):
                                modifiers_match = False
                                break
                        if modifiers_match:
                            self.emergency_stop()
                            return
                            
                except Exception:
                    # 忽略钩子中的异常，避免影响主循环
                    pass

            # 使用键盘事件钩子替代 add_hotkey，避免旧版本库内部属性异常
            self.hotkey_hook = keyboard.hook(handler)
            
        except Exception as e:
            self.log_status(f"快捷键设置失败: {e}")

    def listen_for_key(self):
        """监听按键输入"""
        def on_key_event(event):
            if getattr(event, 'event_type', '') == 'down' and self.is_binding_key:
                key_name = getattr(event, 'name', '')
                
                # 跳过单独的修饰键，等待实际按键
                if key_name in ['ctrl', 'alt', 'shift', 'left ctrl', 'right ctrl', 
                               'left alt', 'right alt', 'left shift', 'right shift']:
                    return True  # 继续监听
                
                # 获取当前按下的修饰键
                modifiers = []
                import keyboard
                if keyboard.is_pressed('ctrl'):
                    modifiers.append('ctrl')
                if keyboard.is_pressed('alt'):
                    modifiers.append('alt')
                if keyboard.is_pressed('shift'):
                    modifiers.append('shift')
                
                # 构建快捷键字符串
                if modifiers:
                    hotkey_str = '+'.join(modifiers + [key_name])
                else:
                    hotkey_str = key_name
                
                # 根据绑定类型更新对应的快捷键
                if self.binding_key_type == 'start':
                    self.start_hotkey_combination.set(hotkey_str)
                    self.start_hotkey_display.config(text=hotkey_str, cursor="hand2")
                else:
                    self.stop_hotkey_combination.set(hotkey_str)
                    self.stop_hotkey_display.config(text=hotkey_str, cursor="hand2")
                
                # 结束绑定
                self.is_binding_key = False
                self.binding_key_type = None
                
                # 重新设置快捷键监听
                import keyboard as _kb
                if hasattr(self, '_binding_hook') and self._binding_hook:
                    try:
                        _kb.unhook(self._binding_hook)
                    except Exception:
                        pass
                    self._binding_hook = None
                self.root.after(100, self.setup_hotkey_listener)
                
                return False  # 停止监听
        
        if self.is_binding_key:
            import keyboard
            self._binding_hook = keyboard.hook(on_key_event)
    
    def emergency_stop(self):
        """紧急停止功能"""
        if self.is_running:
            self.is_running = False
            self.root.after(0, self.stop_translation)
            self.root.after(0, lambda: self.log_status(f"*** 紧急停止 - 通过{self.stop_hotkey_combination.get()}快捷键触发 ***"))
    
    def trigger_start_translation(self):
        """通过快捷键触发开始翻译"""
        if not self.is_running:
            self.root.after(0, self.start_translation)
            self.root.after(0, lambda: self.log_status(f"*** 开始翻译 - 通过{self.start_hotkey_combination.get()}快捷键触发 ***"))
    
    def start_key_binding(self, key_type):
        """开始按键绑定
        
        Args:
            key_type: 'start' 或 'stop'，表示绑定开始快捷键还是停止快捷键
        """
        if self.is_binding_key:
            return
            
        self.is_binding_key = True
        self.binding_key_type = key_type
        
        # 更新对应的显示
        if key_type == 'start':
            self.start_hotkey_display.config(text="等待按键...", cursor="")
        else:
            self.stop_hotkey_display.config(text="等待按键...", cursor="")
        
        # 在绑定期间暂停现有快捷键监听，避免误触
        try:
            import keyboard as _kb
            if hasattr(self, 'hotkey_hook') and self.hotkey_hook:
                try:
                    _kb.unhook(self.hotkey_hook)
                except Exception:
                    pass
                self.hotkey_hook = None
        except Exception:
            pass
        
        # 开始监听按键
        self.root.after(100, self.listen_for_key)
    
    
    def update_memory_usage(self):
        """更新内存使用（后台监控，不显示）"""
        try:
            # 静默监控内存，进行垃圾回收
            memory_mb = self.process.memory_info().rss / 1024 / 1024
            if memory_mb > 100:  # 内存超过100MB时进行垃圾回收
                gc.collect()
        except Exception:
            pass
        
        # 每5秒更新一次
        self.root.after(5000, self.update_memory_usage)
    
    def start_translation(self):
        """开始翻译"""
        # 检查坐标是否都已设置（根据复制方案选择对应必需项）
        required_keys = ['poedit_source', 'poedit_target', 'service_source']
        if self.copy_method.get() in [1, 2, 3]:  # 全选复制、双击复制、三击复制都需要译文框坐标
            required_keys.append('service_result_box')
        else:  # 复制按钮方式
            required_keys.append('service_copy_button')
        
        # 如果启用了鼠标手势滚动到底部，也需要检查手势位置坐标
        if self.use_scroll_gesture.get():
            required_keys.append('scroll_gesture_position')
        
        for coord_type in required_keys:
            if self.coordinates.get(coord_type) is None:
                messagebox.showerror("错误", f"请先设置{coord_type}的坐标")
                return
        
        self.is_running = True
        self.start_button.config(state='disabled')
        self.stop_button.config(state='normal')
        
        # 重置相同原文计数器
        self.same_source_count = 0
        self.last_source_text = ""
        
        self.log_status("开始自动翻译...")
        
        # 在新线程中运行翻译
        self.translation_thread = threading.Thread(target=self.translation_loop, daemon=True)
        self.translation_thread.start()
    
    def stop_translation(self):
        """停止翻译"""
        self.is_running = False
        self.start_button.config(state='normal')
        self.stop_button.config(state='disabled')
        
        self.log_status("已停止翻译")
        
        # 清理资源
        gc.collect()
    
    def translation_loop(self):
        """翻译循环主逻辑"""
        try:
            while self.is_running:
                # 1. 先获取原文
                source_text = self.get_poedit_source_text()
                if not source_text or not self.is_running:
                    if not self.is_running:
                        break
                    self.log_status("无法获取原文，可能已完成所有翻译")
                    break
                
                # 检查是否与上次相同（判断是否到底部）
                if source_text == self.last_source_text:
                    self.same_source_count += 1
                    self.log_status(f"检测到相同原文 ({self.same_source_count}/3)")
                    if self.same_source_count >= 3:
                        self.log_status("连续3次检测到相同原文，翻译已完成")
                        break
                else:
                    self.same_source_count = 0  # 重置计数器
                
                # 更新上次原文
                self.last_source_text = source_text
                
                # 检查停止状态
                if not self.is_running:
                    break
                
                # 2. 检查译文框是否为空或只有空格
                target_text = self.get_poedit_target_text()
                if self.skip_translated.get() and target_text and target_text.strip():
                    self.log_status("跳过已翻译字段")
                    self.next_translation_item()
                    continue
                
                # 检查停止状态
                if not self.is_running:
                    break
                
                # 保存当前原文
                self.clipboard_monitor.save_content_to_temp(source_text, "source")
                self.clipboard_monitor.update_last_content(source_text)
                
                self.log_status(f"正在翻译: {source_text[:50]}...")
                
                # 3. 将原文粘贴到翻译服务
                self.paste_to_translation_service(source_text)
                
                # 检查停止状态
                if not self.is_running:
                    break
                
                # 4. 等待翻译结果
                translated_text = self.wait_for_translation_result(source_text)
                if not translated_text or not self.is_running:
                    if not self.is_running:
                        break
                    self.log_status("翻译超时或失败，跳到下一项")
                    # 翻译失败时，先点击Poedit译文框确保焦点正确
                    try:
                        x, y = self.coordinates['poedit_target']
                        pyautogui.click(x, y)
                        time.sleep(0.05)  # 等待焦点切换
                    except Exception as e:
                        self.log_status(f"点击Poedit译文框失败: {str(e)}")
                    self.next_translation_item()
                    continue
                
                # 检查停止状态
                if not self.is_running:
                    break
                
                # 5. 将翻译结果粘贴到Poedit
                self.paste_to_poedit_target(translated_text)
                self.last_translated_text = translated_text  # 记录本次译文，防止下一条误用
                
                self.log_status(f"翻译完成: {translated_text[:50]}...")
                
                # 6. 跳到下一项
                self.next_translation_item()
                
                # 短暂延迟，但检查停止状态
                for _ in range(5):  # 分成5次检查，每次10ms
                    if not self.is_running:
                        break
                    time.sleep(0.01)
                
        except Exception as e:
            self.log_status(f"翻译过程出错: {str(e)}")
        finally:
            self.root.after(0, self.stop_translation)
    
    def get_poedit_source_text(self) -> str:
        """获取Poedit原文"""
        try:
            x, y = self.coordinates['poedit_source']
            pyautogui.click(x, y)
            time.sleep(0.05)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.05)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.05)
            pyautogui.click(x, y)  # 取消选中
            
            return self.clipboard_monitor.get_clipboard_content()
        except Exception as e:
            self.log_status(f"获取原文失败: {str(e)}")
            return ""
    
    def get_poedit_target_text(self) -> str:
        """获取Poedit译文"""
        try:
            x, y = self.coordinates['poedit_target']
            pyautogui.click(x, y)
            time.sleep(0.05)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.05)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.05)
            pyautogui.click(x, y)  # 取消选中
            
            return self.clipboard_monitor.get_clipboard_content()
        except Exception as e:
            self.log_status(f"获取译文失败: {str(e)}")
            return ""
    
    def paste_to_translation_service(self, text: str):
        """将文本粘贴到翻译服务"""
        try:
            # 如果启用了换行转换，将换行替换为特殊字符
            if self.convert_newlines.get():
                # 统一换行为LF，再用占位符替换，避免遗留的\r导致仍出现换行
                text = text.replace('\r\n', '\n').replace('\r', '\n')
                text = text.replace('\n', '__NL_114514__')
                self.log_status(f"已将换行转换为特殊字符")
            
            pyperclip.copy(text)
            x, y = self.coordinates['service_source']
            pyautogui.click(x, y)
            time.sleep(0.05)
            # 先清空原文框内容
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.05)
            pyautogui.press('backspace')
            time.sleep(0.05)
            # 粘贴新内容
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.05)
            
            # 如果启用了鼠标手势滚动到底部，执行滚动手势
            if self.use_scroll_gesture.get():
                self.perform_scroll_to_bottom_gesture()
                
        except Exception as e:
            self.log_status(f"粘贴到翻译服务失败: {str(e)}")
    
    def perform_scroll_to_bottom_gesture(self):
        """执行鼠标手势滚动到底部"""
        try:
            if self.coordinates['scroll_gesture_position'] is None:
                self.log_status("鼠标手势位置未设置，跳过滚动手势")
                return
                
            x, y = self.coordinates['scroll_gesture_position']
            self.log_status("执行鼠标手势滚动到底部...")
            
            # 移动到指定位置
            pyautogui.moveTo(x, y)
            time.sleep(0.05)
            
            # 按下右键开始拖拽
            pyautogui.mouseDown(button='right')
            time.sleep(0.05)
            
            # 先向上滑动50像素
            pyautogui.move(0, -50, duration=0.15)
            time.sleep(0.05)
            
            # 再向下滑动50像素（相对于当前位置）
            pyautogui.move(0, 50, duration=0.15)
            
            # 释放右键
            pyautogui.mouseUp(button='right')
            
            # 等待用户设置的时间
            wait_time = self.scroll_gesture_wait_time.get() / 1000.0
            self.log_status(f"滚动手势完成，等待 {self.scroll_gesture_wait_time.get()}ms...")
            time.sleep(wait_time)
            
        except Exception as e:
            self.log_status(f"执行鼠标手势失败: {str(e)}")
    
    def perform_scroll_to_top_gesture(self):
        """执行鼠标手势滚动到顶部"""
        try:
            if self.coordinates['scroll_gesture_position'] is None:
                self.log_status("鼠标手势位置未设置，跳过滚动手势")
                return
                
            x, y = self.coordinates['scroll_gesture_position']
            self.log_status("执行鼠标手势滚动到顶部...")
            
            # 移动到指定位置
            pyautogui.moveTo(x, y)
            time.sleep(0.05)
            
            # 按下右键开始拖拽
            pyautogui.mouseDown(button='right')
            time.sleep(0.05)
            
            # 先向下滑动50像素
            pyautogui.move(0, 50, duration=0.15)
            time.sleep(0.05)
            
            # 再向上滑动50像素（相对于当前位置）
            pyautogui.move(0, -50, duration=0.15)
            
            # 释放右键
            pyautogui.mouseUp(button='right')
            time.sleep(0.05)
            
        except Exception as e:
            self.log_status(f"执行滚动到顶部手势失败: {str(e)}")
    
    def wait_for_translation_result(self, original_text: str) -> str:
        """等待翻译结果"""
        # 等待用户设置的翻译结果等待时间
        wait_time = self.translation_wait_time.get() / 1000.0
        self.log_status(f"等待翻译结果生成中... ({self.translation_wait_time.get()}ms)")
        time.sleep(wait_time)
        
        try:
            # 在开始复制前，先把剪贴板重置为当前原文，确保基线
            pyperclip.copy(original_text)
            
            # 复制译文（根据设置：复制按钮、全选复制、双击复制、三击复制）
            if self.copy_method.get() == 1:  # 全选复制
                rx, ry = self.coordinates['service_result_box']
                pyautogui.click(rx, ry)
                time.sleep(0.05)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.05)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.05)
                pyautogui.click(rx, ry)  # 取消选中
                time.sleep(0.05)
            elif self.copy_method.get() == 2:  # 双击复制
                rx, ry = self.coordinates['service_result_box']
                pyautogui.doubleClick(rx, ry)  # 双击选中
                time.sleep(0.05)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.05)
                pyautogui.click(rx, ry)  # 取消选中
                time.sleep(0.05)
            elif self.copy_method.get() == 3:  # 三击复制
                rx, ry = self.coordinates['service_result_box']
                pyautogui.click(rx, ry)  # 第一次点击
                time.sleep(0.05)
                pyautogui.click(rx, ry)  # 第二次点击
                time.sleep(0.05)
                pyautogui.click(rx, ry)  # 第三次点击
                time.sleep(0.05)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.05)
                pyautogui.click(rx, ry)  # 取消选中
                time.sleep(0.05)
            else:  # 复制按钮方式
                x, y = self.coordinates['service_copy_button']
                pyautogui.click(x, y)
                time.sleep(0.05)  # 等待复制操作完成
            
            translated_text = self.clipboard_monitor.get_clipboard_content()
            
            # 如果启用了翻译一致性检测
            if self.check_translation_consistency.get():
                check_count = 0
                max_checks = self.check_timeout_count.get()
                check_interval = self.check_interval.get() / 1000.0
                
                self.log_status(f"开始检测翻译结果是否与原文一致...")
                
                while check_count < max_checks and self.is_running:
                    # 如果启用了换行转换，需要将剪贴板文本中的占位符转换为换行符再进行比较
                    comparison_text = translated_text
                    if self.convert_newlines.get() and '__NL_114514__' in translated_text:
                        comparison_text = translated_text.replace('__NL_114514__', '\n')
                    
                    # 剪贴板仍为原文：说明没有复制到译文
                    if comparison_text == original_text.strip():
                        self.log_status(f"剪贴板仍为原文，翻译服务处理中... ({check_count}/{max_checks})")
                    # 非空白且不等于原文：有可能是译文或上一条旧译文
                    elif comparison_text and not comparison_text.isspace() and comparison_text.strip() != original_text.strip():
                        # 防止误用上一条旧译文：若与上一条译文一致，则继续等待刷新
                        if getattr(self, 'last_translated_text', "") and comparison_text.strip() == self.last_translated_text.strip():
                            self.log_status(f"检测到与上一条译文相同，可能是旧内容，继续等待... ({check_count}/{max_checks})")
                        else:
                            self.log_status("检测到新译文，翻译完成")
                            break  # 跳出循环，让后面统一处理上划手势
                    else:
                        # 空内容或全空格，继续等待
                        self.log_status(f"剪贴板为空或全空格，继续等待... ({check_count}/{max_checks})")
                    
                    # 等待后重试
                    check_count += 1
                    time.sleep(check_interval)
                    
                    # 重试前重新复制译文
                    if self.copy_method.get() == 1:  # 全选复制
                        if self.coordinates['service_result_box'] is None:
                            self.log_status("翻译服务译文框坐标未设置，无法重试复制")
                            break
                        rx, ry = self.coordinates['service_result_box']
                        pyautogui.click(rx, ry)
                        time.sleep(0.05)
                        pyautogui.hotkey('ctrl', 'a')
                        time.sleep(0.05)
                        pyautogui.hotkey('ctrl', 'c')
                        time.sleep(0.05)
                        pyautogui.click(rx, ry)
                        time.sleep(0.05)
                    elif self.copy_method.get() == 2:  # 双击复制
                        if self.coordinates['service_result_box'] is None:
                            self.log_status("翻译服务译文框坐标未设置，无法重试复制")
                            break
                        rx, ry = self.coordinates['service_result_box']
                        pyautogui.doubleClick(rx, ry)  # 双击选中
                        time.sleep(0.05)
                        pyautogui.hotkey('ctrl', 'c')
                        time.sleep(0.05)
                        pyautogui.click(rx, ry)  # 取消选中
                        time.sleep(0.05)
                    elif self.copy_method.get() == 3:  # 三击复制
                        if self.coordinates['service_result_box'] is None:
                            self.log_status("翻译服务译文框坐标未设置，无法重试复制")
                            break
                        rx, ry = self.coordinates['service_result_box']
                        pyautogui.click(rx, ry)  # 第一次点击
                        time.sleep(0.05)
                        pyautogui.click(rx, ry)  # 第二次点击
                        time.sleep(0.05)
                        pyautogui.click(rx, ry)  # 第三次点击
                        time.sleep(0.05)
                        pyautogui.hotkey('ctrl', 'c')
                        time.sleep(0.05)
                        pyautogui.click(rx, ry)  # 取消选中
                        time.sleep(0.05)
                    else:  # 复制按钮方式
                        if self.coordinates['service_copy_button'] is None:
                            self.log_status("翻译服务复制按钮坐标未设置，无法重试复制")
                            break
                        x, y = self.coordinates['service_copy_button']
                        pyautogui.click(x, y)
                        time.sleep(0.05)
                    translated_text = self.clipboard_monitor.get_clipboard_content()
                
                # 超时后直接返回当前剪贴板内容（若有效）
                if check_count >= max_checks:
                    self.log_status(f"检测超时，使用当前剪贴板内容")
                    # 注意：这里不直接return，让后面统一处理上划手势
            
            # 如果启用了鼠标手势滚动到底部，在翻译一致性检测完成后滚动到顶部
            if self.use_scroll_gesture.get():
                if self.coordinates['scroll_gesture_position'] is None:
                    self.log_status("鼠标手势位置未设置，跳过滚动手势")
                else:
                    self.perform_scroll_to_top_gesture()
            
            # 如果启用了换行转换，将特殊字符转换回换行
            if self.convert_newlines.get() and '__NL_114514__' in translated_text:
                translated_text = translated_text.replace('__NL_114514__', '\n')
                # 统一换行，避免出现多余空行（比如遗留的\r 与我们插入的\n 叠加）
                translated_text = translated_text.replace('\r\n', '\n').replace('\r', '\n')
                self.log_status(f"已将特殊字符转换回换行")
            
            # 如果未启用检测或检测通过，直接返回翻译结果
            if translated_text and not translated_text.isspace():
                return translated_text
            
        except Exception as e:
            self.log_status(f"获取翻译结果时出错: {str(e)}")
        
        return ""
    
    def paste_to_poedit_target(self, text: str):
        """将翻译结果粘贴到Poedit译文框"""
        try:
            pyperclip.copy(text)
            x, y = self.coordinates['poedit_target']
            pyautogui.click(x, y)
            time.sleep(0.05)
            # 先清空译文框内容
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.05)
            pyautogui.press('backspace')
            time.sleep(0.05)
            # 粘贴新内容
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.05)
        except Exception as e:
            self.log_status(f"粘贴译文失败: {str(e)}")
    
    def next_translation_item(self):
        """跳到下一个翻译项"""
        try:
            pyautogui.hotkey('ctrl', 'down')
            time.sleep(0.05)
        except Exception as e:
            self.log_status(f"跳转下一项失败: {str(e)}")
    
    def get_config_directory(self):
        """获取配置文件目录路径"""
        user_home = os.path.expanduser("~")
        config_dir = os.path.join(user_home, "AppData", "Local", "poedit-automatic-translation")
        return config_dir
    
    def ensure_config_directory(self):
        """确保配置目录存在"""
        if not os.path.exists(self.config_dir):
            try:
                os.makedirs(self.config_dir)

            except Exception as e:
                print(f"创建配置目录失败: {e}")
                # 如果创建失败，回退到当前目录
                self.config_dir = os.getcwd()
                self.config_file = os.path.join(self.config_dir, "poedit-automatic-translation_config.json")
    
    def save_config(self):
        """保存配置"""
        # 首次保存时创建配置目录
        self.ensure_config_directory()
        
        config = {
            'coordinates': self.coordinates,
            'skip_translated': self.skip_translated.get(),
            'translation_wait_time': self.translation_wait_time.get(),
            'start_hotkey_combination': self.start_hotkey_combination.get(),
            'stop_hotkey_combination': self.stop_hotkey_combination.get(),
            'check_translation_consistency': self.check_translation_consistency.get(),
            'check_interval': self.check_interval.get(),
            'check_timeout_count': self.check_timeout_count.get(),
            'copy_method': self.copy_method.get(),
            'use_scroll_gesture': self.use_scroll_gesture.get(),
            'scroll_gesture_wait_time': self.scroll_gesture_wait_time.get(),
            'convert_newlines': self.convert_newlines.get()
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.log_status("配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def load_config(self):
        """加载配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                if 'coordinates' in config:
                    self.coordinates = config['coordinates']
                if 'skip_translated' in config:
                    self.skip_translated.set(config['skip_translated'])
                if 'translation_wait_time' in config:
                    self.translation_wait_time.set(config['translation_wait_time'])
                
                # 加载快捷键设置
                if 'start_hotkey_combination' in config:
                    self.start_hotkey_combination.set(config['start_hotkey_combination'])
                    if hasattr(self, 'start_hotkey_display'):
                        self.start_hotkey_display.config(text=config['start_hotkey_combination'])
                        
                if 'stop_hotkey_combination' in config:
                    self.stop_hotkey_combination.set(config['stop_hotkey_combination'])
                    if hasattr(self, 'stop_hotkey_display'):
                        self.stop_hotkey_display.config(text=config['stop_hotkey_combination'])
                        
                # 重新设置快捷键监听
                self.root.after(100, self.setup_hotkey_listener)
                
                # 加载新的翻译检测配置项
                if 'check_translation_consistency' in config:
                    self.check_translation_consistency.set(config['check_translation_consistency'])
                if 'check_interval' in config:
                    self.check_interval.set(config['check_interval'])
                if 'check_timeout_count' in config:
                    self.check_timeout_count.set(config['check_timeout_count'])
                
                # 加载复制方案设置
                if 'copy_method' in config:
                    self.copy_method.set(config['copy_method'])
                
                # 加载鼠标手势设置
                if 'use_scroll_gesture' in config:
                    self.use_scroll_gesture.set(config['use_scroll_gesture'])
                if 'scroll_gesture_wait_time' in config:
                    self.scroll_gesture_wait_time.set(config['scroll_gesture_wait_time'])
                
                # 加载换行转换设置
                if 'convert_newlines' in config:
                    self.convert_newlines.set(config['convert_newlines'])
                
                # 更新UI显示
                for key, coord in self.coordinates.items():
                    if coord:
                        self.coord_vars[key].set(f"({coord[0]}, {coord[1]})")
                
                # 配置加载完成，不显示日志
            else:
                pass # 首次打开软件没有配置文件时，不显示任何日志
                
        except Exception as e:
            self.log_status(f"加载配置失败: {str(e)}")


    
    def on_closing(self):
        """程序关闭时的清理工作"""
        self.is_running = False
        
        # 清理快捷键监听（仅解除我们注册的钩子）
        try:
            import keyboard as _kb
            if hasattr(self, 'hotkey_hook') and self.hotkey_hook:
                try:
                    _kb.unhook(self.hotkey_hook)
                except Exception:
                    pass
                self.hotkey_hook = None
            if hasattr(self, '_binding_hook') and self._binding_hook:
                try:
                    _kb.unhook(self._binding_hook)
                except Exception:
                    pass
                self._binding_hook = None
        except Exception:
            pass
        
        # 清理临时文件
        self.clipboard_monitor.cleanup_temp_file()
        
        # 强制垃圾回收
        gc.collect()
        
        self.root.destroy()
    
    def run(self):
        """运行程序"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()


if __name__ == "__main__":
    # 设置pyautogui安全设置
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.05
    
    app = PoeditAutoTranslator()
    app.run()