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
        self.root.geometry("600x500")
        
        # 配置文件路径
        self.config_dir = self.get_config_directory()
        self.config_file = os.path.join(self.config_dir, "poedit-automatic-translation_config.json")
        
        # 坐标配置
        self.coordinates = {
            'poedit_source': None,      # Poedit原文框
            'poedit_target': None,      # Poedit译文框
            'service_source': None,     # 翻译服务原文框
            'service_copy_button': None # 翻译服务复制按钮
        }
        
        # 设置选项
        self.skip_translated = tk.BooleanVar(value=True)  # 跳过已翻译
        self.translation_wait_time = tk.IntVar(value=3000)  # 翻译结果等待时间(毫秒)
        
        # 快捷键设置
        self.hotkey_combination = tk.StringVar(value="F9")  # 默认快捷键
        self.is_binding_key = False  # 是否正在绑定按键
        
        # 运行状态
        self.is_running = False
        self.translation_thread = None
        self.same_source_count = 0  # 连续相同原文计数器
        self.last_source_text = ""  # 上次的原文
        
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
            ("翻译服务复制按钮:", 'service_copy_button')
        ]
        
        self.coord_vars = {}
        for i, (label_text, key) in enumerate(coord_labels):
            ttk.Label(settings_frame, text=label_text).grid(row=i, column=0, sticky=tk.W, pady=2)
            
            var = tk.StringVar(value="未设置")
            self.coord_vars[key] = var
            ttk.Label(settings_frame, textvariable=var, width=20).grid(row=i, column=1, padx=(10, 5), pady=2)
            
            ttk.Button(settings_frame, text="选择", 
                      command=lambda k=key: self.select_coordinate(k)).grid(row=i, column=2, pady=2)
        
        # 快捷键绑定设置（放在翻译服务复制按钮下方）
        ttk.Label(settings_frame, text="停止快捷键:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.hotkey_display = ttk.Label(settings_frame, text=self.hotkey_combination.get(), 
                                       relief="sunken", width=20)
        self.hotkey_display.grid(row=4, column=1, padx=(10, 5), pady=2)
        self.bind_button = ttk.Button(settings_frame, text="绑定按键", 
                                     command=self.start_key_binding)
        self.bind_button.grid(row=4, column=2, pady=2)
        
        # 其他选项设置
        ttk.Checkbutton(settings_frame, text="跳过已翻译的字段", 
                       variable=self.skip_translated).grid(row=5, column=0, sticky=tk.W, pady=2)
        
        ttk.Label(settings_frame, text="翻译结果等待时间(毫秒):").grid(row=6, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(settings_frame, from_=100, to=5000, textvariable=self.translation_wait_time, 
                   width=10).grid(row=6, column=1, padx=(10, 0), pady=2)
        
        
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
        status_frame = ttk.LabelFrame(main_frame, text="运行状态", padding="10")
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

            # 解析快捷键组合
            combo = self.hotkey_combination.get().strip().lower()
            parts = [p.strip() for p in combo.split('+') if p.strip()]
            modifiers = {p for p in parts if p in ('ctrl', 'alt', 'shift')}
            main_key = next((p for p in reversed(parts) if p not in ('ctrl', 'alt', 'shift')), None)

            def handler(event):
                try:
                    # 绑定按键期间不触发紧急停止
                    if getattr(self, 'is_binding_key', False):
                        return
                    # 仅在按下事件时判断
                    if getattr(event, 'event_type', '') != 'down':
                        return
                    # 检查修饰键状态
                    for m in modifiers:
                        if not keyboard.is_pressed(m):
                            return
                    # 检查主键（如果存在）
                    if main_key and getattr(event, 'name', '') != main_key:
                        return
                    # 触发紧急停止
                    self.emergency_stop()
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
                
                # 更新快捷键
                self.hotkey_combination.set(hotkey_str)
                self.hotkey_display.config(text=hotkey_str)
                
                # 结束绑定
                self.is_binding_key = False
                self.bind_button.config(text="绑定按键", state="normal")
                
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
            self.root.after(0, lambda: self.log_status(f"*** 紧急停止 - 通过{self.hotkey_combination.get()}快捷键触发 ***"))
    
    def start_key_binding(self):
        """开始按键绑定"""
        if self.is_binding_key:
            return
            
        self.is_binding_key = True
        self.bind_button.config(text="按下按键...", state="disabled")
        self.hotkey_display.config(text="等待按键...")
        
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
        # 检查坐标是否都已设置
        for coord_type, coord in self.coordinates.items():
            if coord is None:
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
                if not source_text:
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
                
                # 2. 检查译文框是否为空
                target_text = self.get_poedit_target_text()
                if self.skip_translated.get() and target_text:
                    self.log_status("跳过已翻译字段")
                    self.next_translation_item()
                    continue
                
                # 保存当前原文
                self.clipboard_monitor.save_content_to_temp(source_text, "source")
                self.clipboard_monitor.update_last_content(source_text)
                
                self.log_status(f"正在翻译: {source_text[:50]}...")
                
                # 3. 将原文粘贴到翻译服务
                self.paste_to_translation_service(source_text)
                
                # 4. 等待翻译结果
                translated_text = self.wait_for_translation_result(source_text)
                if not translated_text:
                    self.log_status("翻译超时或失败，跳到下一项")
                    self.next_translation_item()
                    continue
                
                # 5. 将翻译结果粘贴到Poedit
                self.paste_to_poedit_target(translated_text)
                
                self.log_status(f"翻译完成: {translated_text[:50]}...")
                
                # 6. 跳到下一项
                self.next_translation_item()
                
                # 短暂延迟
                time.sleep(0.05)
                
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
        except Exception as e:
            self.log_status(f"粘贴到翻译服务失败: {str(e)}")
    
    def wait_for_translation_result(self, original_text: str) -> str:
        """等待翻译结果"""
        # 等待用户设置的翻译结果等待时间
        wait_time = self.translation_wait_time.get() / 1000.0
        self.log_status(f"等待翻译结果生成中... ({self.translation_wait_time.get()}ms)")
        time.sleep(wait_time)
        
        try:
            # 点击翻译服务复制按钮
            x, y = self.coordinates['service_copy_button']
            pyautogui.click(x, y)
            time.sleep(0.05)  # 等待复制操作完成
            
            translated_text = self.clipboard_monitor.get_clipboard_content()
            
            # 返回翻译结果（不检查是否与原文相同）
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
                print(f"创建配置目录: {self.config_dir}")
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
            'hotkey_combination': self.hotkey_combination.get()
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
                
                self.coordinates = config['coordinates']
                self.skip_translated.set(config['skip_translated'])
                self.translation_wait_time.set(config['translation_wait_time'])
                
                # 加载快捷键设置
                if 'hotkey_combination' in config:
                    self.hotkey_combination.set(config['hotkey_combination'])
                    if hasattr(self, 'hotkey_display'):
                        self.hotkey_display.config(text=config['hotkey_combination'])
                    # 重新设置快捷键监听
                    self.root.after(100, self.setup_hotkey_listener)
                
                # 更新UI显示
                for key, coord in self.coordinates.items():
                    if coord:
                        self.coord_vars[key].set(f"({coord[0]}, {coord[1]})")
                
                self.log_status("配置加载成功")
            else:
                self.log_status("请设置坐标")
                
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