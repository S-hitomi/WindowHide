import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import win32gui
import win32con
import win32process
import threading
import mouse
import keyboard
import os
import time
import json
from pystray import MenuItem, Icon
from PIL import Image, ImageDraw
import queue

# 常量定义
CONFIG_FILE = "config.json"
GESTURE_MIN_POINTS = 5
MY_PID = os.getpid()


def is_self_window(hwnd):
    """检查给定的窗口句柄是否属于当前Python进程。"""
    if not hwnd: return False
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid == MY_PID
    except Exception:
        return False


class GestureHandler:
    """一个处理手势逻辑的普通类，通过独立的线程主动记录路径。"""

    def __init__(self, app_instance, trigger_button, pattern, callback):
        self.app = app_instance
        self.trigger_button = trigger_button
        self.pattern = pattern
        self.callback = callback
        self.path = []
        self.is_recording = False
        self.screen_width = app_instance.root.winfo_screenwidth()
        self.screen_height = app_instance.root.winfo_screenheight()
        self.recording_thread = None
        self.lock = threading.Lock()

    def handle_event(self, event):
        """处理由全局分发器转发的鼠标事件，仅用于启动和停止手势。"""
        if isinstance(event, mouse.ButtonEvent) and event.button == self.trigger_button:
            if event.event_type == mouse.DOWN:
                self._start_recording()
            elif event.event_type == mouse.UP:
                self._stop_recording()

    def _start_recording(self):
        with self.lock:
            if self.is_recording: return
            self.is_recording = True
            self.path = [mouse.get_position()]
            self.recording_thread = threading.Thread(target=self._record_path_worker, daemon=True)
            self.recording_thread.start()

    def _record_path_worker(self):
        """在后台线程中主动轮询并记录鼠标位置。"""
        while self.is_recording:
            current_pos = mouse.get_position()
            with self.lock:
                if not self.path or self.path[-1] != current_pos:
                    self.path.append(current_pos)
            time.sleep(0.01)

    def _stop_recording(self):
        with self.lock:
            if not self.is_recording: return
            self.is_recording = False

        if self.recording_thread:
            self.recording_thread.join(timeout=0.5)

        with self.lock:
            if len(self.path) > GESTURE_MIN_POINTS:
                self.analyze_gesture()
            self.path = []

    def analyze_gesture(self):
        if not self.path: return
        start_x, start_y = self.path[0]
        end_x, end_y = self.path[-1]
        delta_x = end_x - start_x
        delta_y = end_y - start_y

        h_threshold = self.screen_width / 5
        v_threshold = self.screen_height / 5
        h_tolerance = self.screen_height / 10
        v_tolerance = self.screen_width / 10

        if self.pattern == 'swipe_right' and delta_x > h_threshold and abs(delta_y) < h_tolerance:
            self.app.root.after(0, self.callback)
        elif self.pattern == 'swipe_left' and delta_x < -h_threshold and abs(delta_y) < h_tolerance:
            self.app.root.after(0, self.callback)
        elif self.pattern == 'swipe_up' and delta_y < -v_threshold and abs(delta_x) < v_tolerance:
            self.app.root.after(0, self.callback)
        elif self.pattern == 'swipe_down' and delta_y > v_threshold and abs(delta_x) < v_tolerance:
            self.app.root.after(0, self.callback)


class WindowMonitor:
    """监控指定窗口，根据鼠标是否悬停来调整其透明度，并可选择隐藏其任务栏图标。"""

    def __init__(self, hwnd, root, always_on_top=False, away_transparency=50, hover_opacity=100, hide_taskbar=False):
        self.hwnd = hwnd
        self.root = root
        self.always_on_top = always_on_top
        self.hide_taskbar = hide_taskbar
        self.transparent_level_byte = int(away_transparency / 100 * 255)
        self.opaque_level_byte = int(hover_opacity / 100 * 255)
        self.running = False
        self.lock = threading.Lock()
        self.original_ex_style = win32gui.GetWindowLong(self.hwnd, win32con.GWL_EXSTYLE)

        try:
            new_style = self.original_ex_style | win32con.WS_EX_LAYERED
            if self.hide_taskbar: new_style |= win32con.WS_EX_TOOLWINDOW
            win32gui.SetWindowLong(self.hwnd, win32con.GWL_EXSTYLE, new_style)
        except Exception as e:
            app = root.app_instance
            if hasattr(e, 'winerror') and e.winerror == 5:
                raise RuntimeError(app._('error_style_permission'))
            raise RuntimeError(app._('error_style_unknown').format(e=e))

    def set_always_on_top(self):
        if win32gui.IsWindow(self.hwnd): win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                                               win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def remove_always_on_top(self):
        if win32gui.IsWindow(self.hwnd): win32gui.SetWindowPos(self.hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                                               win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def make_transparent(self):
        with self.lock:
            if win32gui.IsWindow(self.hwnd): win32gui.SetLayeredWindowAttributes(self.hwnd, 0,
                                                                                 self.transparent_level_byte,
                                                                                 win32con.LWA_ALPHA)

    def make_opaque(self):
        with self.lock:
            if win32gui.IsWindow(self.hwnd): win32gui.SetLayeredWindowAttributes(self.hwnd, 0, self.opaque_level_byte,
                                                                                 win32con.LWA_ALPHA)

    def check_mouse_position(self):
        if not self.running: return
        if not win32gui.IsWindow(self.hwnd):
            self.stop_monitoring()
            self.root.after(0, self.root.app_instance.handle_window_closed)
            return
        try:
            x, y = mouse.get_position()
            rect = win32gui.GetWindowRect(self.hwnd)
            if rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]:
                self.make_opaque()
            else:
                self.make_transparent()
        except Exception:
            pass
        if self.running: self.root.after(100, self.check_mouse_position)

    def start_monitoring(self):
        if not self.running:
            self.running = True
            if self.always_on_top: self.set_always_on_top()
            self.check_mouse_position()

    def stop_monitoring(self):
        if self.running:
            self.running = False
            with self.lock:
                if win32gui.IsWindow(self.hwnd):
                    if self.always_on_top: self.remove_always_on_top()
                    win32gui.SetLayeredWindowAttributes(self.hwnd, 0, 255, win32con.LWA_ALPHA)
                    win32gui.SetWindowLong(self.hwnd, win32con.GWL_EXSTYLE, self.original_ex_style)
                    win32gui.SetWindowPos(self.hwnd, 0, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED)


class App:
    """应用程序主界面和逻辑"""

    def __init__(self, root):
        self.root = root
        # 在这里修改程序的默认启动尺寸
        self.root.geometry("480x930")
        self.root.minsize(420, 500)
        self.root.app_instance = self

        # 初始化状态变量
        self.is_fully_initialized = False
        self.is_closing = False
        self.monitor = None
        self.windows_map = {}
        self.selected_hwnd_by_mouse = None
        self.triggers = {}
        self.mouse_button_callbacks = {}
        self.gesture_handlers = {}
        self.is_recording_hotkey = False
        self.tray_icon = None
        self.tray_thread = None
        self.is_capturing_click = False  # 鼠标监听标签
        self.temp_icon_path = None # 临时图标路径

        # 初始化国际化(i18n)系统
        self.language_var = tk.StringVar(value='zh')
        self.i18n = self.get_language_definitions()

        # 定义Combobox的内部值和映射
        self.mb_values = ['middle_click', 'wheel_up', 'wheel_down']
        self.mg_trigger_values = ['middle', 'right']
        self.mg_pattern_values = ['swipe_right', 'swipe_left', 'swipe_up', 'swipe_down']

        # 初始化其他设置
        self.tray_icon_path_var = tk.StringVar(value='icon.ico')

        # 用于鼠标事件的线程安全队列
        self.mouse_event_queue = queue.Queue()

        self.setup_ui()
        self.load_settings()

        # 标题栏图标更新
        self.tray_icon_path_var.trace_add('write', self.update_window_icon)
        self.update_window_icon()

        # 绑定语言变化事件到UI更新函数
        self.language_var.trace_add('write', self.on_language_change)
        self.update_ui_text()  # 应用加载的或默认的语言

        self.selected_label.config(text=self._('status_refreshing_list'))
        self.root.update_idletasks()

        self.refresh_windows()

        self.is_fully_initialized = True
        self.setup_all_triggers()
        self.update_ui_states()

        # --- 启动永久的鼠标监听器和队列处理器 ---
        mouse.hook(self.mouse_event_queue.put)
        self.process_mouse_queue()

        if hasattr(self, 'last_monitored_title') and self.last_monitored_title:
            self.preselect_last_window(self.last_monitored_title)
        else:
            self.selected_label.config(text=self._('status_no_window_selected'))

    def _(self, key, **kwargs):
        """获取当前语言的文本。"""
        lang = self.language_var.get()
        return self.i18n.get(lang, {}).get(key, f"_{key}_").format(**kwargs)

    def on_language_change(self, *args):
        """语言变化时的回调函数。"""
        self.update_ui_text()

    def update_window_icon(self, *args):
        """根据设置更新主窗口的标题栏图标。"""
        # 清理上一个临时图标（如果存在）
        if self.temp_icon_path and os.path.exists(self.temp_icon_path):
            try:
                os.remove(self.temp_icon_path)
                self.temp_icon_path = None
            except OSError as e:
                print(f"Error removing old temp icon: {e}")

        image = None
        custom_icon_path = self.tray_icon_path_var.get()

        if custom_icon_path and os.path.exists(custom_icon_path):
            try:
                image = Image.open(custom_icon_path)
            except Exception as e:
                print(f"Failed to load custom icon for window icon, using default: {e}")
                image = self._create_default_icon_image()
        else:
            image = self._create_default_icon_image()

        try:
            self.temp_icon_path = 'icon.ico'
            image.save('icon.ico', format="ICO", sizes=[(64, 64)])
            self.root.iconbitmap(self.temp_icon_path)

        except Exception as e:
            print(f"Could not set window icon: {e}")

    def setup_ui(self):
        # UI元素的字典，用于语言切换时更新文本
        self.ui_elements = {}
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(container, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, padding=10)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        def _configure_interior(event):
            # 更新内部Frame的宽度以匹配Canvas的宽度
            self.canvas.itemconfig(self.canvas_window, width=event.width)

        def _configure_canvas(event):
            # 更新滚动区域
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            # 根据内容是否超出Canvas高度，来决定是否显示滚动条
            if self.scrollable_frame.winfo_reqheight() > self.canvas.winfo_height():
                if not self.scrollbar.winfo_viewable():
                    self.scrollbar.grid(row=0, column=1, sticky='ns')
            else:
                if self.scrollbar.winfo_viewable():
                    self.scrollbar.grid_remove()

        self.scrollable_frame.bind("<Configure>", _configure_canvas)
        self.canvas.bind("<Configure>", _configure_interior)

        def _on_main_mousewheel(event):
            # 仅当滚动条可见时才滚动
            if self.scrollbar.winfo_viewable():
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _on_list_mousewheel(event):
            # 防止在Listbox上滚动时影响主滚动条
            pass

        self.canvas.bind("<MouseWheel>", _on_main_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", _on_main_mousewheel)

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        # 默认不显示滚动条，由_configure_canvas决定
        # self.scrollbar.grid(row=0, column=1, sticky='ns')

        main_frame = self.scrollable_frame

        # --- 窗口列表 ---
        list_container = ttk.Frame(main_frame)
        list_container.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        self.window_list = tk.Listbox(list_container, width=50, height=10)
        list_scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.window_list.yview)
        self.window_list.config(yscrollcommand=list_scrollbar.set)
        list_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.window_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.window_list.bind("<<ListboxSelect>>", self.on_list_select)
        self.window_list.bind("<MouseWheel>", _on_list_mousewheel)  # 阻止事件传播

        self.selected_label = tk.Label(main_frame, fg="blue", wraplength=380)
        self.selected_label.pack(pady=(5, 5), fill=tk.X)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=5)
        self.ui_elements['select_mouse_button'] = ttk.Button(button_frame, command=self.select_window_with_mouse)
        self.ui_elements['select_mouse_button'].pack(side=tk.LEFT, padx=5)
        self.ui_elements['refresh_button'] = ttk.Button(button_frame, command=self.refresh_windows)
        self.ui_elements['refresh_button'].pack(side=tk.LEFT, padx=5)

        # --- 设置 Notebook ---
        self.settings_notebook = ttk.Notebook(main_frame)
        self.settings_notebook.pack(fill=tk.X, pady=10)

        # --- 通用设置标签页 ---
        general_tab = ttk.Frame(self.settings_notebook, padding=10)
        self.ui_elements['general_tab'] = general_tab
        self.settings_notebook.add(general_tab, text=self._('tab_general'))

        lang_frame = ttk.LabelFrame(general_tab, padding=(10, 5))
        self.ui_elements['lang_frame'] = lang_frame
        lang_frame.pack(pady=5, fill=tk.X)
        self.ui_elements['lang_label'] = ttk.Label(lang_frame)
        self.ui_elements['lang_label'].pack(side=tk.LEFT, padx=(0, 10))
        lang_combo = ttk.Combobox(lang_frame, textvariable=self.language_var, values=['中文', 'English'],
                                  state='readonly', width=10)
        lang_combo.pack(side=tk.LEFT)
        lang_combo.bind("<<ComboboxSelected>>",
                        lambda e: self.language_var.set('zh' if lang_combo.get() == '中文' else 'en'))

        icon_frame = ttk.LabelFrame(general_tab, padding=(10, 5))
        self.ui_elements['icon_frame'] = icon_frame
        icon_frame.pack(pady=5, fill=tk.X)
        self.ui_elements['icon_button'] = ttk.Button(icon_frame, command=self.select_tray_icon)
        self.ui_elements['icon_button'].pack(side=tk.LEFT, padx=(0, 10))
        self.ui_elements['icon_button_clear'] = ttk.Button(icon_frame, command=self.clear_tray_icon)
        self.ui_elements['icon_button_clear'].pack(side=tk.LEFT, padx=(5, 10))
        ttk.Label(icon_frame, textvariable=self.tray_icon_path_var, wraplength=280, foreground="gray").pack(
            side=tk.LEFT, fill=tk.X, expand=True)

        instru_frame = ttk.LabelFrame(general_tab, padding=(10, 5))
        self.ui_elements['instru_frame'] = instru_frame
        instru_frame.pack(pady=5, fill=tk.X)
        self.ui_elements['instru_label'] = tk.Label(instru_frame, fg="red", wraplength=380)
        self.ui_elements['instru_label'].pack(fill=tk.X, padx=5, pady=5)

        # --- 透明度标签页 ---
        transparency_tab = ttk.Frame(self.settings_notebook, padding=10)
        self.ui_elements['transparency_tab'] = transparency_tab
        self.settings_notebook.add(transparency_tab, text=self._('tab_transparency'))

        settings_frame = ttk.LabelFrame(transparency_tab, padding=(10, 5))
        self.ui_elements['transparency_settings_frame'] = settings_frame
        settings_frame.pack(pady=5, fill=tk.X)
        hover_frame = ttk.Frame(settings_frame)
        hover_frame.pack(fill=tk.X, expand=True, pady=2)
        self.ui_elements['hover_opacity_label'] = ttk.Label(hover_frame)
        self.ui_elements['hover_opacity_label'].pack(side=tk.LEFT, padx=(0, 10))
        self.hover_opacity_var = tk.IntVar(value=100)
        ttk.Label(hover_frame, textvariable=self.hover_opacity_var, width=4).pack(side=tk.RIGHT)
        self.ui_elements['hover_opacity_scale'] = ttk.Scale(hover_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.hover_opacity_var,
                  command=lambda v: self.hover_opacity_var.set(int(float(v))))
        self.ui_elements['hover_opacity_scale'].pack(side=tk.RIGHT, fill=tk.X, expand=True)

        away_frame = ttk.Frame(settings_frame)
        away_frame.pack(fill=tk.X, expand=True, pady=2)
        self.ui_elements['away_opacity_label'] = ttk.Label(away_frame)
        self.ui_elements['away_opacity_label'].pack(side=tk.LEFT)
        self.away_transparency_var = tk.IntVar(value=50)
        ttk.Label(away_frame, textvariable=self.away_transparency_var, width=4).pack(side=tk.RIGHT)
        self.ui_elements['away_transparency_scale'] = ttk.Scale(away_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.away_transparency_var,
                  command=lambda v: self.away_transparency_var.set(int(float(v))))
        self.ui_elements['away_transparency_scale'].pack(side=tk.RIGHT, fill=tk.X, expand=True)

        options_frame = ttk.LabelFrame(transparency_tab, padding=(10, 5))
        self.ui_elements['monitor_options_frame'] = options_frame
        options_frame.pack(pady=10, fill=tk.X)
        self.always_on_top_var = tk.BooleanVar()
        self.always_on_top_check = ttk.Checkbutton(options_frame, variable=self.always_on_top_var)
        self.ui_elements['always_on_top_check'] = self.always_on_top_check
        self.always_on_top_check.pack(anchor=tk.W)
        self.hide_taskbar_var = tk.BooleanVar()
        self.hide_taskbar_check = ttk.Checkbutton(options_frame, variable=self.hide_taskbar_var)
        self.ui_elements['hide_taskbar_check'] = self.hide_taskbar_check
        self.hide_taskbar_check.pack(anchor=tk.W)

        # --- 触发器标签页 ---
        self.hotkey_tab = ttk.Frame(self.settings_notebook, padding=10)
        self.ui_elements['hotkey_tab'] = self.hotkey_tab
        self.settings_notebook.add(self.hotkey_tab, text=self._('tab_triggers'))
        self.trigger_actions = ['minimize_monitored_window', 'close_window', 'hide_tray', 'show_tray', 'exit_app']
        for action_name in self.trigger_actions:
            self.create_trigger_ui(self.hotkey_tab, action_name)

        # --- 控制按钮 ---
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(10, 0))
        self.ui_elements['start_button'] = ttk.Button(control_frame, command=self.start_monitoring)
        self.ui_elements['start_button'].pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        self.ui_elements['stop_button'] = ttk.Button(control_frame, command=self.stop_monitoring_ui)
        self.ui_elements['stop_button'].pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))
        self.ui_elements['tray_button'] = ttk.Button(main_frame, command=self.minimize_to_tray)
        self.ui_elements['tray_button'].pack(fill=tk.X, pady=(5, 0))

    def update_ui_text(self):
        """根据当前语言更新所有UI文本。"""
        self.root.title(self._('window_title'))

        # 更新 Notebook 标签页文本
        self.settings_notebook.tab(self.ui_elements['general_tab'], text=self._('tab_general'))
        self.settings_notebook.tab(self.ui_elements['transparency_tab'], text=self._('tab_transparency'))
        self.settings_notebook.tab(self.ui_elements['hotkey_tab'], text=self._('tab_triggers'))

        # 更新通用设置页
        self.ui_elements['lang_frame'].config(text=self._('frame_language'))
        self.ui_elements['lang_label'].config(text=self._('label_language'))
        self.ui_elements['icon_frame'].config(text=self._('frame_tray_icon'))
        self.ui_elements['icon_button'].config(text=self._('button_select_icon'))
        self.ui_elements['icon_button_clear'].config(text=self._('button_clear_icon'))
        # 使用说明
        self.ui_elements['instru_frame'].config(text=self._('instructions'))
        self.ui_elements['instru_label'].config(text=self._('instructions_label'))

        # 更新透明度页
        self.ui_elements['transparency_settings_frame'].config(text=self._('frame_transparency'))
        self.ui_elements['hover_opacity_label'].config(text=self._('label_hover_opacity'))
        self.ui_elements['away_opacity_label'].config(text=self._('label_away_opacity'))
        self.ui_elements['monitor_options_frame'].config(text=self._('frame_monitor_options'))
        self.ui_elements['always_on_top_check'].config(text=self._('check_always_on_top'))
        self.ui_elements['hide_taskbar_check'].config(text=self._('check_hide_taskbar'))

        # 更新触发器页
        for action_name in self.trigger_actions:
            ui_map = getattr(self, f"trigger_ui_{action_name}")
            ui_map['frame'].config(text=self._(f'frame_trigger_{action_name}'))
            ui_map['type_label'].config(text=self._('label_trigger_type'))
            ui_map['kb_radio'].config(text=self._('radio_keyboard'))
            ui_map['mb_radio'].config(text=self._('radio_mouse_button'))
            ui_map['mg_radio'].config(text=self._('radio_mouse_gesture'))
            ui_map['kb_label'].config(text=self._('label_hotkey'))
            ui_map['kb_button'].config(text=self._('button_set_hotkey'))
            ui_map['mb_label'].config(text=self._('label_mouse_button'))
            ui_map['mg_trigger_label'].config(text=self._('label_gesture_trigger_key'))
            ui_map['mg_pattern_label'].config(text=self._('label_gesture_pattern'))

            # 更新Combobox的显示值
            self._update_combobox_display(ui_map, 'mb_combo', self.mb_values, 'mb_var')
            self._update_combobox_display(ui_map, 'mg_trigger_combo', self.mg_trigger_values, 'mg_trigger_var')
            self._update_combobox_display(ui_map, 'mg_pattern_combo', self.mg_pattern_values, 'mg_pattern_var')

        # 更新主按钮
        self.ui_elements['select_mouse_button'].config(text=self._('button_select_with_mouse'))
        self.ui_elements['refresh_button'].config(text=self._('button_refresh_list'))
        self.ui_elements['start_button'].config(text=self._('button_start_monitoring'))
        self.ui_elements['stop_button'].config(text=self._('button_stop_monitoring'))
        self.ui_elements['tray_button'].config(text=self._('button_minimize_to_tray'))

        # 更新状态标签，如果它已经有内容
        current_text = self.selected_label.cget("text")
        # 定义语言无关的前缀
        monitoring_prefixes = (
            self.i18n['en']['status_monitoring'].split(':')[0], self.i18n['zh']['status_monitoring'].split(':')[0])
        selected_prefixes = (
            self.i18n['en']['status_selected'].split(':')[0], self.i18n['zh']['status_selected'].split(':')[0])

        if any(current_text.startswith(prefix) for prefix in monitoring_prefixes):
            title = current_text.split(":")[-1].strip()
            self.selected_label.config(text=self._('status_monitoring', title=title))
        elif any(current_text.startswith(prefix) for prefix in selected_prefixes):
            title = current_text.split(":")[-1].strip()
            self.selected_label.config(text=self._('status_selected', title=title))

    def _update_combobox_display(self, ui_map, combo_key, internal_values, var_key):
        """一个帮助函数，用于更新单个Combobox的显示值和列表。"""
        if combo_key in ui_map:
            combo = ui_map[combo_key]
            # 创建翻译后的显示列表
            display_list = [self._(f'combo_{v}') for v in internal_values]
            combo['values'] = display_list

            # 获取当前存储的内部值
            current_internal_value = ui_map[var_key].get()
            # 找到对应的显示值
            if current_internal_value in internal_values:
                current_display_value = self._(f'combo_{current_internal_value}')
                combo.set(current_display_value)
            else:  # 如果值无效，则选择第一个
                combo.current(0)
                internal_val = ui_map['reverse_map'][combo.get()]
                ui_map[var_key].set(internal_val)

    def create_trigger_ui(self, parent, action_name):
        ui_map = {}
        main_frame = ttk.LabelFrame(parent, padding=(10, 5))
        ui_map['frame'] = main_frame
        main_frame.pack(pady=5, fill=tk.X)
        setattr(self, f"trigger_ui_{action_name}", ui_map)

        type_frame = ttk.Frame(main_frame)
        type_frame.pack(fill=tk.X, pady=2)
        ui_map['type_label'] = ttk.Label(type_frame)
        ui_map['type_label'].pack(side=tk.LEFT)
        ui_map['type_var'] = tk.StringVar(value='keyboard')

        options_container = ttk.Frame(main_frame)
        options_container.pack(fill=tk.X, pady=(5, 0))
        ui_map['options_frames'] = {}

        default_hotkeys = {
            'minimize_monitored_window': 'ctrl+alt+m',
            'close_window': 'ctrl+alt+c',
            'hide_tray': 'ctrl+alt+h',
            'show_tray': 'ctrl+alt+s',
            'exit_app': 'ctrl+alt+x'
        }

        def on_type_change(*args):
            selected_type = ui_map['type_var'].get()
            for f_type, frame in ui_map['options_frames'].items():
                frame.pack_forget()
            ui_map['options_frames'][selected_type].pack(fill=tk.X)
            if self.is_fully_initialized: self.setup_all_triggers()

        ui_map['type_var'].trace_add('write', on_type_change)

        ui_map['kb_radio'] = ttk.Radiobutton(type_frame, variable=ui_map['type_var'], value="keyboard")
        ui_map['kb_radio'].pack(side=tk.LEFT, padx=5)
        ui_map['mb_radio'] = ttk.Radiobutton(type_frame, variable=ui_map['type_var'], value="mouse_button")
        ui_map['mb_radio'].pack(side=tk.LEFT, padx=5)
        ui_map['mg_radio'] = ttk.Radiobutton(type_frame, variable=ui_map['type_var'], value="mouse_gesture")
        ui_map['mg_radio'].pack(side=tk.LEFT, padx=5)

        # 键盘快捷键
        kb_frame = ttk.Frame(options_container)
        ui_map['options_frames']['keyboard'] = kb_frame
        ui_map['kb_label'] = ttk.Label(kb_frame)
        ui_map['kb_label'].pack(side=tk.LEFT)
        default_hotkey = default_hotkeys.get(action_name, '')
        ui_map['kb_var'] = tk.StringVar(value=default_hotkey)
        ttk.Label(kb_frame, textvariable=ui_map['kb_var'], font=("Segoe UI", 9, "bold"), foreground="#0078D7").pack(
            side=tk.LEFT, padx=5)
        ui_map['kb_button'] = ttk.Button(kb_frame, command=lambda: self.start_hotkey_recording(action_name))
        ui_map['kb_button'].pack(side=tk.RIGHT)

        # 鼠标按键
        mb_frame = ttk.Frame(options_container)
        ui_map['options_frames']['mouse_button'] = mb_frame
        ui_map['mb_label'] = ttk.Label(mb_frame)
        ui_map['mb_label'].pack(side=tk.LEFT)
        ui_map['mb_var'] = tk.StringVar(value='middle_click')
        ui_map['mb_reverse_map'] = {}
        mb_combo = ttk.Combobox(mb_frame, state='readonly', width=15)
        ui_map['mb_combo'] = mb_combo
        mb_combo.pack(side=tk.LEFT, padx=5)
        mb_combo.bind("<<ComboboxSelected>>",
                      lambda e, u=ui_map: self.on_combo_select(u, 'mb_var', 'mb_reverse_map', e.widget.get()))

        # 鼠标手势
        mg_frame = ttk.Frame(options_container)
        ui_map['options_frames']['mouse_gesture'] = mg_frame

        ui_map['mg_trigger_label'] = ttk.Label(mg_frame)
        ui_map['mg_trigger_label'].pack(side=tk.LEFT)
        ui_map['mg_trigger_var'] = tk.StringVar(value='right')
        ui_map['mg_trigger_reverse_map'] = {}
        mg_trigger_combo = ttk.Combobox(mg_frame, state='readonly', width=8)
        ui_map['mg_trigger_combo'] = mg_trigger_combo
        mg_trigger_combo.pack(side=tk.LEFT, padx=5)
        mg_trigger_combo.bind("<<ComboboxSelected>>",
                              lambda e, u=ui_map: self.on_combo_select(u, 'mg_trigger_var', 'mg_trigger_reverse_map',
                                                                       e.widget.get()))

        ui_map['mg_pattern_label'] = ttk.Label(mg_frame)
        ui_map['mg_pattern_label'].pack(side=tk.LEFT, padx=(10, 0))
        ui_map['mg_pattern_var'] = tk.StringVar(value='swipe_right')
        ui_map['mg_pattern_reverse_map'] = {}
        mg_pattern_combo = ttk.Combobox(mg_frame, state='readonly', width=12)
        ui_map['mg_pattern_combo'] = mg_pattern_combo
        mg_pattern_combo.pack(side=tk.LEFT, padx=5)
        mg_pattern_combo.bind("<<ComboboxSelected>>",
                              lambda e, u=ui_map: self.on_combo_select(u, 'mg_pattern_var', 'mg_pattern_reverse_map',
                                                                       e.widget.get()))

        on_type_change()

    def on_combo_select(self, ui_map, var_key, reverse_map_key, selected_display_value):
        """当Combobox被选择时，更新内部的StringVar。"""
        internal_value = ui_map[reverse_map_key].get(selected_display_value)
        if internal_value:
            ui_map[var_key].set(internal_value)
        self.setup_all_triggers()

    def on_list_select(self, event):
        selected_indices = self.window_list.curselection()
        if selected_indices:
            self.selected_hwnd_by_mouse = None
            title = self.window_list.get(selected_indices[0])
            self.selected_label.config(text=self._('status_selected', title=title))

    def select_window_with_mouse(self):
        if self.is_capturing_click:
            return
        self.is_capturing_click = True
        self.selected_label.config(text=self._('status_clicking_to_select'))
        self.root.iconify()
        mouse.on_click(self._capture_click)

    def _capture_click(self):
        if not self.is_capturing_click:
            return
        self.is_capturing_click = False

        try:
            mouse.unhook(self._capture_click)
        except (KeyError, ValueError):
            pass  # 如果已解钩 忽略

        # 在当前事件处理完成后再运行核心逻辑
        self.root.after(10, self._capture_click_logic)

    def _capture_click_logic(self):
        # 仅在鼠标点击选取结束并解钩后执行
        self.root.deiconify()  # 先恢复窗口

        try:
            pos = win32gui.GetCursorPos()
            hwnd = win32gui.WindowFromPoint(pos)
            top_level_hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)

            if is_self_window(top_level_hwnd):
                # 安全弹窗
                self.root.after(50, self._handle_self_selection)
            else:
                title = win32gui.GetWindowText(top_level_hwnd) or self._('untitled_window')
                self.update_selection_by_mouse(top_level_hwnd, title)
        except Exception as e:
            print(f"Error during window capture logic: {e}")

    def update_selection_by_mouse(self, hwnd, title):
        self.selected_hwnd_by_mouse = hwnd
        self.selected_label.config(text=self._('status_selected', title=title))
        self.window_list.selection_clear(0, tk.END)

    def _handle_self_selection(self):
        messagebox.showwarning(self._('title_invalid_op'), self._('error_cannot_select_self'))
        self.selected_label.config(text=self._('status_no_window_selected'))
        self.selected_hwnd_by_mouse = None

    def refresh_windows(self):
        self.window_list.delete(0, tk.END)
        self.windows_map.clear()

        def enum_windows(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and not is_self_window(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:
                    self.windows_map[title] = hwnd
                    self.window_list.insert(tk.END, title)

        win32gui.EnumWindows(enum_windows, None)

    def check_for_duplicate_triggers(self):
        trigger_map = {}
        for name in self.trigger_actions:
            if hasattr(self, f"trigger_ui_{name}"):
                ui_map = getattr(self, f"trigger_ui_{name}")
                trigger_type = ui_map['type_var'].get()
                trigger_config = None
                action_label = self._(f'frame_trigger_{name}')

                if trigger_type == 'keyboard':
                    value = ui_map['kb_var'].get();
                    trigger_config = (
                        self._('radio_keyboard'), value) if value else None
                elif trigger_type == 'mouse_button':
                    value = ui_map['mb_var'].get();
                    trigger_config = (
                        self._('radio_mouse_button'), value) if value else None
                elif trigger_type == 'mouse_gesture':
                    trigger_btn, pattern = ui_map['mg_trigger_var'].get(), ui_map[
                        'mg_pattern_var'].get();
                    trigger_config = (
                        self._('radio_mouse_gesture'), f"{trigger_btn} + {pattern}")

                if trigger_config:
                    if trigger_config not in trigger_map: trigger_map[trigger_config] = []
                    trigger_map[trigger_config].append(action_label)

        conflicts = [f"- {cfg[0]} '{cfg[1]}' " + self._('conflict_used_for') + " " + "、".join(f"“{a}”" for a in acts)
                     for cfg, acts in trigger_map.items() if len(acts) > 1]

        if conflicts:
            return self._('error_conflict_header') + "\n\n" + "\n".join(conflicts)
        return None

    def start_monitoring(self):
        conflict_message = self.check_for_duplicate_triggers()
        if conflict_message:
            messagebox.showerror(self._('title_conflict'), conflict_message)
            return

        hwnd_to_monitor = None
        if self.selected_hwnd_by_mouse:
            if win32gui.IsWindow(self.selected_hwnd_by_mouse):
                hwnd_to_monitor = self.selected_hwnd_by_mouse
            else:
                messagebox.showwarning(self._('title_warning'), self._('error_window_closed'))
                self.selected_hwnd_by_mouse = None
                self.selected_label.config(text=self._('status_no_window_selected'))
                return
        else:
            selected_indices = self.window_list.curselection()
            if selected_indices:
                hwnd_to_monitor = self.windows_map.get(self.window_list.get(selected_indices[0]))
            else:
                messagebox.showwarning(self._('title_warning'), self._('error_select_window_first'))
                return

        if not hwnd_to_monitor or not win32gui.IsWindow(hwnd_to_monitor):
            messagebox.showerror(self._('title_error'), self._('error_invalid_handle'))
            self.refresh_windows()
            return

        if is_self_window(hwnd_to_monitor):
            messagebox.showerror(self._('title_invalid_op'), self._('error_cannot_monitor_self'))
            return

        try:
            self.monitor = WindowMonitor(hwnd_to_monitor, self.root, self.always_on_top_var.get(),
                                         self.away_transparency_var.get(), self.hover_opacity_var.get(),
                                         self.hide_taskbar_var.get())
            self.monitor.start_monitoring()
            self.setup_all_triggers()
            self.update_ui_states()
            title = win32gui.GetWindowText(hwnd_to_monitor)
            self.selected_label.config(text=self._('status_monitoring', title=title))
        except Exception as e:
            messagebox.showerror(self._('title_start_failed'), self._('error_start_failed').format(e=e))
            if self.monitor: self.monitor.stop_monitoring(); self.monitor = None
            self.update_ui_states()

    def stop_monitoring_ui(self):
        if self.monitor: self.monitor.stop_monitoring(); self.monitor = None
        self.setup_all_triggers()
        self.update_ui_states()
        self.selected_label.config(text=self._('status_stopped'))

    def handle_window_closed(self):
        messagebox.showinfo(self._('title_info'), self._('info_window_closed'))
        self.stop_monitoring_ui()

    def update_ui_states(self):
        is_monitoring = self.monitor is not None and self.monitor.running
        is_recording = self.is_recording_hotkey
        general_state = tk.DISABLED if is_monitoring or is_recording else tk.NORMAL

        self.ui_elements['start_button'].config(state=general_state)
        self.ui_elements['stop_button'].config(state=tk.NORMAL if is_monitoring and not is_recording else tk.DISABLED)

        self.ui_elements['tray_button'].config(state=tk.DISABLED if is_recording else tk.NORMAL)
        for widget_key in ['refresh_button', 'select_mouse_button', 'always_on_top_check', 'hide_taskbar_check']:
            self.ui_elements[widget_key].config(state=general_state)
        self.window_list.config(state=general_state)

        # transparency_settings_frame
        opacity_controls_state = tk.DISABLED if is_recording or is_monitoring else tk.NORMAL
        self.ui_elements['hover_opacity_label'].config(state=opacity_controls_state)
        self.ui_elements['away_opacity_label'].config(state=opacity_controls_state)
        self.ui_elements['hover_opacity_scale'].config(state=opacity_controls_state)
        self.ui_elements['away_transparency_scale'].config(state=opacity_controls_state)

        for action_name in self.trigger_actions:
            ui_map = getattr(self, f"trigger_ui_{action_name}", {})
            trigger_controls_state = tk.DISABLED if is_recording or is_monitoring else tk.NORMAL
            for frame in ui_map.get('options_frames', {}).values():
                for child in frame.winfo_children(): child.config(state=trigger_controls_state)
            for radio in [ui_map['kb_radio'], ui_map['mb_radio'], ui_map['mg_radio']]: radio.config(
                state=trigger_controls_state)
            if is_recording and self.recording_key_name == action_name:
                ui_map['kb_button'].config(state=tk.NORMAL)

    def start_hotkey_recording(self, action_name):
        if self.is_recording_hotkey: return
        self.is_recording_hotkey = True
        self.recording_key_name = action_name
        ui_map = getattr(self, f"trigger_ui_{action_name}")
        ui_map['kb_button'].config(text=self._('button_press_hotkey'))
        self.update_ui_states()
        threading.Thread(target=self.record_hotkey_worker, daemon=True).start()

    def record_hotkey_worker(self):
        try:
            new_hotkey = keyboard.read_hotkey(suppress=False)
            self.root.after(0, self.stop_hotkey_recording, new_hotkey)
        except Exception as e:
            print(f"Error reading hotkey: {e}")
            self.root.after(0, self.stop_hotkey_recording, None)

    def stop_hotkey_recording(self, new_hotkey):
        action_name = self.recording_key_name
        ui_map = getattr(self, f"trigger_ui_{action_name}")
        if new_hotkey: ui_map['kb_var'].set(new_hotkey)
        ui_map['kb_button'].config(text=self._('button_set_hotkey'))
        self.is_recording_hotkey = False
        self.setup_all_triggers()
        self.update_ui_states()

    def process_mouse_queue(self):
        try:
            while not self.mouse_event_queue.empty(): self._global_mouse_dispatcher(self.mouse_event_queue.get_nowait())
        finally:
            if not self.is_closing: self.root.after(20, self.process_mouse_queue)

    def _global_mouse_dispatcher(self, event):
        for handler in self.gesture_handlers.values(): handler.handle_event(event)
        if isinstance(event, mouse.WheelEvent):
            cb_key = 'wheel_up' if event.delta > 0 else 'wheel_down'
            if cb_key in self.mouse_button_callbacks: self.mouse_button_callbacks[cb_key]()
        elif isinstance(event, mouse.ButtonEvent) and event.event_type == mouse.UP and event.button == mouse.MIDDLE:
            if 'middle_click' in self.mouse_button_callbacks: self.mouse_button_callbacks['middle_click']()

    def setup_all_triggers(self):
        self.remove_all_triggers()
        actions = {'minimize_monitored_window': self.trigger_minimize_monitored_window,
                   'close_window': self.trigger_force_close, 'hide_tray': self.hide_tray_icon,
                   'show_tray': self.show_tray_icon_from_hotkey, 'exit_app': self.on_closing}
        self.mouse_button_callbacks.clear()
        self.gesture_handlers.clear()
        for name, callback in actions.items():
            if name in ['close_window', 'minimize_monitored_window'] and (
                    self.monitor is None or not self.monitor.running): continue
            ui_map = getattr(self, f"trigger_ui_{name}")
            trigger_type = ui_map['type_var'].get()
            try:
                if trigger_type == 'keyboard':
                    hotkey = ui_map['kb_var'].get()
                    if hotkey: self.triggers[f"{name}_kb"] = keyboard.add_hotkey(hotkey, callback, suppress=True)
                elif trigger_type == 'mouse_button':
                    self.mouse_button_callbacks[ui_map['mb_var'].get()] = callback
                elif trigger_type == 'mouse_gesture':
                    self.gesture_handlers[name] = GestureHandler(self, ui_map['mg_trigger_var'].get(),
                                                                 ui_map['mg_pattern_var'].get(), callback)
            except Exception as e:
                messagebox.showerror(self._('title_trigger_error'), self._('error_set_trigger').format(name=name, e=e))

    def remove_all_triggers(self):
        for trigger in self.triggers.values():
            try:
                keyboard.remove_hotkey(trigger)
            except (KeyError, ValueError):
                pass
        self.triggers.clear()

    def trigger_minimize_monitored_window(self):
        if self.monitor and self.monitor.running:
            try:
                hwnd = self.monitor.hwnd
                if win32gui.IsWindow(hwnd):
                    cmd = win32con.SC_RESTORE if win32gui.IsIconic(hwnd) else win32con.SC_MINIMIZE
                    win32gui.PostMessage(hwnd, win32con.WM_SYSCOMMAND, cmd, 0)
            except Exception as e:
                print(f"Error toggling window minimization: {e}")

    def trigger_force_close(self):
        if self.monitor and self.monitor.running: self.root.after(0, self.execute_force_close)

    def execute_force_close(self):
        if not (self.monitor and self.monitor.running): return
        hwnd = self.monitor.hwnd
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            pid = None
        self.stop_monitoring_ui()
        time.sleep(0.1)
        if win32gui.IsWindow(hwnd):
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            threading.Thread(target=self.check_and_kill, args=(hwnd, pid), daemon=True).start()

    def check_and_kill(self, hwnd, pid):
        time.sleep(1.5)
        if pid and win32gui.IsWindow(hwnd): os.system(f"taskkill /PID {pid} /F /T > nul")

    def select_tray_icon(self):
        path = filedialog.askopenfilename(
            title=self._('dialog_select_icon'),
            filetypes=[(self._('dialog_image_files'), '*.png *.ico'), (self._('dialog_all_files'), '*.*')]
        )
        if path:
            self.tray_icon_path_var.set(path)

    def clear_tray_icon(self):
        """清空自定义托盘图标路径，恢复默认图标"""
        self.tray_icon_path_var.set('')

    def _create_default_icon_image(self):
        """用于绘制默认图标的PIL.Image对象"""
        width, height = 64, 64
        image = Image.new('RGB', (width, height), 'white')
        dc = ImageDraw.Draw(image)
        dc.ellipse((16, 24, 48, 40), fill='black', outline='black')
        dc.ellipse((28, 30, 36, 38), fill='white', outline='white')
        return image

    def create_tray_image(self):
        path = self.tray_icon_path_var.get()
        if path and os.path.exists(path):
            try:
                return Image.open(path)
            except Exception as e:
                print(f"Failed to load custom tray icon '{path}': {e}")
        # 调用默认图标
        width, height = 64, 64
        image = Image.new('RGB', (width, height), 'white')
        dc = ImageDraw.Draw(image)
        dc.ellipse((16, 24, 48, 40), fill='black', outline='black')
        dc.ellipse((28, 30, 36, 38), fill='white', outline='white')
        return image

    def minimize_to_tray(self):
        self.root.withdraw()
        self.show_tray_icon()

    def show_tray_icon(self):
        if self.tray_icon and self.tray_icon.visible: return
        image = self.create_tray_image()
        menu = (MenuItem(self._('tray_show_window'), self.show_window_from_tray, default=True),
                MenuItem(self._('tray_exit'), self.exit_app_from_tray))
        self.tray_icon = Icon("WindowMonitor", image, self._('window_title'), menu)
        self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        self.tray_thread.start()

    def show_tray_icon_from_hotkey(self):
        if not (self.tray_icon and self.tray_icon.visible): self.root.after(10, self.show_tray_icon)

    def hide_tray_icon(self):
        if self.tray_icon and self.tray_icon.visible:
            self.tray_icon.stop()
            self.root.after(100, self._cleanup_tray_references)

    def _cleanup_tray_references(self):
        self.tray_icon, self.tray_thread = None, None

    def show_window_from_tray(self):
        self.hide_tray_icon()
        self.root.after(0, self.root.deiconify)

    def exit_app_from_tray(self):
        self.hide_tray_icon()
        self.root.after(0, self.on_closing)

    def on_closing(self):
        if self.is_closing: return
        self.is_closing = True
        if self.monitor and self.monitor.running: self.monitor.stop_monitoring()
        self.root.withdraw()
        threading.Thread(target=self._perform_cleanup_and_exit, daemon=True).start()

    def _perform_cleanup_and_exit(self):
        self.save_settings()
        mouse.unhook(self.mouse_event_queue.put)
        if self.tray_icon and self.tray_icon.visible: self.tray_icon.stop()
        self.remove_all_triggers()
        self.root.after(0, self.root.destroy)

    def save_settings(self):
        settings = {'triggers': {}, 'general': {}}
        for action in self.trigger_actions:
            if hasattr(self, f"trigger_ui_{action}"):
                ui_map = getattr(self, f"trigger_ui_{action}")
                settings['triggers'][action] = {
                    'type': ui_map['type_var'].get(), 'keyboard': ui_map['kb_var'].get(),
                    'mouse_button': ui_map['mb_var'].get(), 'gesture_trigger': ui_map['mg_trigger_var'].get(),
                    'gesture_pattern': ui_map['mg_pattern_var'].get()
                }
        settings['options'] = {'always_on_top': self.always_on_top_var.get(),
                               'hide_taskbar': self.hide_taskbar_var.get()}
        settings['transparency'] = {'hover': self.hover_opacity_var.get(), 'away': self.away_transparency_var.get()}
        settings['general'] = {'language': self.language_var.get(), 'tray_icon_path': self.tray_icon_path_var.get()}

        last_title = None
        if self.monitor and self.monitor.running:
            try:
                if win32gui.IsWindow(self.monitor.hwnd): last_title = win32gui.GetWindowText(self.monitor.hwnd)
            except Exception:
                pass
        settings['last_window_title'] = last_title

        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Failed to save settings: {e}")

    def load_settings(self):
        # 定义默认热键，以便在加载设置失败或文件不存在时使用
        default_hotkeys = {
            'minimize_monitored_window': 'ctrl+alt+m',
            'close_window': 'ctrl+alt+c',
            'hide_tray': 'ctrl+alt+h',
            'show_tray': 'ctrl+alt+s',
            'exit_app': 'ctrl+alt+x'
        }

        if not os.path.exists(CONFIG_FILE):
            # 如果没有配置文件，应用默认热键
            for action, hotkey in default_hotkeys.items():
                if hasattr(self, f"trigger_ui_{action}"):
                    getattr(self, f"trigger_ui_{action}")['kb_var'].set(hotkey)
            return

        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)

            # 加载基础设置
            general = settings.get('general', {})
            self.language_var.set(general.get('language', 'zh'))
            self.tray_icon_path_var.set(general.get('tray_icon_path', ''))

            # 加载其他设置
            triggers = settings.get('triggers', {})
            for action, config in triggers.items():
                if hasattr(self, f"trigger_ui_{action}"):
                    ui_map = getattr(self, f"trigger_ui_{action}")
                    ui_map['type_var'].set(config.get('type', 'keyboard'))
                    # 加载已保存的快捷键，如果不存在，则使用该动作的默认值
                    ui_map['kb_var'].set(config.get('keyboard', default_hotkeys.get(action, '')))
                    ui_map['mb_var'].set(config.get('mouse_button', 'middle_click'))
                    ui_map['mg_trigger_var'].set(config.get('gesture_trigger', 'right'))
                    ui_map['mg_pattern_var'].set(config.get('gesture_pattern', 'swipe_right'))

            options = settings.get('options', {})
            self.always_on_top_var.set(options.get('always_on_top', False))
            self.hide_taskbar_var.set(options.get('hide_taskbar', False))

            transparency = settings.get('transparency', {})
            self.hover_opacity_var.set(transparency.get('hover', 100))
            self.away_transparency_var.set(transparency.get('away', 50))

            self.last_monitored_title = settings.get('last_window_title')
        except Exception as e:
            print(f"Error loading settings ({e}), using defaults.")

    def preselect_last_window(self, title_to_select):
        try:
            items = self.window_list.get(0, tk.END)
            if title_to_select in items:
                index = items.index(title_to_select)
                self.window_list.selection_set(index)
                self.window_list.activate(index)
                self.window_list.see(index)
                self.on_list_select(None)
        except Exception as e:
            print(f"Error preselecting window: {e}")

    def get_language_definitions(self):
        return {
            'zh': {
                'window_title': "窗口监控器",
                'untitled_window': "无标题的窗口",
                'status_refreshing_list': "正在刷新窗口列表...",
                'status_no_window_selected': "尚未选取窗口",
                'status_selected': "已选取: {title}",
                'status_monitoring': "正在监控: {title}",
                'status_stopped': "监控已停止。",
                'status_clicking_to_select': "请点击目标窗口以完成选取...",
                'button_select_with_mouse': "用鼠标选取窗口",
                'button_refresh_list': "刷新窗口列表",
                'button_start_monitoring': "开始监控",
                'button_stop_monitoring': "停止监控",
                'button_minimize_to_tray': "最小化到系统托盘",
                'button_set_hotkey': "点此设置",
                'button_press_hotkey': "请按下组合键...",
                'tab_general': "通用设置",
                'tab_transparency': "透明度 & 选项",
                'tab_triggers': "触发器设置",
                'frame_language': "语言 (Language)",
                'label_language': "界面语言:",
                'frame_tray_icon': "系统托盘图标",
                'button_select_icon': "选择图标文件...",
                'button_clear_icon': "清除图标",
                'dialog_select_icon': "选择一个图标文件",
                'dialog_image_files': "图片文件",
                'dialog_all_files': "所有文件",
                'frame_transparency': "透明度设置 (%)   0表示完全透明",
                'label_hover_opacity': "鼠标悬停不透明度:",
                'label_away_opacity': "鼠标移开不透明度:  ",
                'frame_monitor_options': "监控选项",
                'check_always_on_top': "被监控窗口始终置顶",
                'check_hide_taskbar': "被监控窗口隐藏任务栏图标 (及Alt+Tab)",
                'frame_trigger_minimize_monitored_window': "最小化/复原被监控窗口",
                'frame_trigger_close_window': "关闭被监控窗口",
                'frame_trigger_hide_tray': "隐藏托盘图标",
                'frame_trigger_show_tray': "显示托盘图标",
                'frame_trigger_exit_app': "关闭本程序",
                'label_trigger_type': "触发类型:",
                'radio_keyboard': "键盘",
                'radio_mouse_button': "鼠标按键",
                'radio_mouse_gesture': "鼠标手势",
                'label_hotkey': "快捷键:",
                'label_mouse_button': "按键:",
                'label_gesture_trigger_key': "触发键:",
                'label_gesture_pattern': "手势:",
                'combo_middle_click': "中键单击",
                'combo_wheel_up': "滚轮向上",
                'combo_wheel_down': "滚轮向下",
                'combo_middle': "鼠标中键",
                'combo_right': "鼠标右键",
                'combo_swipe_right': "向右滑动",
                'combo_swipe_left': "向左滑动",
                'combo_swipe_up': "向上滑动",
                'combo_swipe_down': "向下滑动",
                'title_invalid_op': "操作无效",
                'title_warning': "警告",
                'title_error': "错误",
                'title_info': "信息",
                'title_conflict': "设置冲突",
                'title_start_failed': "启动失败",
                'title_trigger_error': "触发器错误",
                'error_cannot_select_self': "不能选取应用程序本身。",
                'error_window_closed': "先前由鼠标选取的窗口已关闭。",
                'error_select_window_first': "请先选取一个窗口。",
                'error_invalid_handle': "无法获取有效窗口句柄。",
                'error_cannot_monitor_self': "不能选择本程序窗口进行监控。",
                'error_start_failed': "启动监控失败: \n{e}",
                'error_style_permission': "无法修改窗口样式，权限不足。\n请尝试以“管理员身份”运行本程序。",
                'error_style_unknown': "无法修改窗口样式，可能是系统窗口或发生未知错误: {e}",
                'error_set_trigger': "无法设置“{name}”的触发器: \n{e}",
                'error_conflict_header': "发现重复的触发器设置，请修改后重试：",
                'conflict_used_for': "同时用于",
                'info_window_closed': "被监控的窗口已关闭，监控自动停止。",
                'tray_show_window': "显示主窗口",
                'tray_exit': "结束程序",
                'instructions': "使用说明",
                'instructions_label': "选择窗口后点击开始监控 \n 通过蓝色字确认所选择窗口 \n\n  请勿使用任务管理器强制退出此程序！ \n如需强制退出请设置触发器快捷键",
            },
            'en': {
                'window_title': "Window Monitor",
                'untitled_window': "Untitled Window",
                'status_refreshing_list': "Refreshing window list...",
                'status_no_window_selected': "No window selected",
                'status_selected': "Selected: {title}",
                'status_monitoring': "Monitoring: {title}",
                'status_stopped': "Monitoring stopped.",
                'status_clicking_to_select': "Please click on the target window to select it...",
                'button_select_with_mouse': "Select with Mouse",
                'button_refresh_list': "Refresh List",
                'button_start_monitoring': "Start Monitoring",
                'button_stop_monitoring': "Stop Monitoring",
                'button_minimize_to_tray': "Minimize to Tray",
                'button_set_hotkey': "Set Hotkey",
                'button_press_hotkey': "Press any key combo...",
                'tab_general': "General",
                'tab_transparency': "Transparency & Options",
                'tab_triggers': "Triggers",
                'frame_language': "Language",
                'label_language': "UI Language:",
                'frame_tray_icon': "System Tray Icon",
                'button_select_icon': "Select Icon File...",
                'button_clear_icon': "Clear Icon",
                'dialog_select_icon': "Select an icon file",
                'dialog_image_files': "Image Files",
                'dialog_all_files': "All Files",
                'frame_transparency': "Transparency Settings (%)",
                'label_hover_opacity': "Hover Opacity:",
                'label_away_opacity': "Away Opacity:  ",
                'frame_monitor_options': "Monitoring Options",
                'check_always_on_top': "Always on Top",
                'check_hide_taskbar': "Hide Taskbar Icon (and Alt+Tab)",
                'frame_trigger_minimize_monitored_window': "Minimize/Restore Monitored Window",
                'frame_trigger_close_window': "Close Monitored Window",
                'frame_trigger_hide_tray': "Hide Tray Icon",
                'frame_trigger_show_tray': "Show Tray Icon",
                'frame_trigger_exit_app': "Exit Application",
                'label_trigger_type': "Trigger Type:",
                'radio_keyboard': "Keyboard",
                'radio_mouse_button': "Mouse Button",
                'radio_mouse_gesture': "Mouse Gesture",
                'label_hotkey': "Hotkey:",
                'label_mouse_button': "Button:",
                'label_gesture_trigger_key': "Trigger Key:",
                'label_gesture_pattern': "Gesture:",
                'combo_middle_click': "Middle Click",
                'combo_wheel_up': "Wheel Up",
                'combo_wheel_down': "Wheel Down",
                'combo_middle': "Middle Button",
                'combo_right': "Right Button",
                'combo_swipe_right': "Swipe Right",
                'combo_swipe_left': "Swipe Left",
                'combo_swipe_up': "Swipe Up",
                'combo_swipe_down': "Swipe Down",
                'title_invalid_op': "Invalid Operation",
                'title_warning': "Warning",
                'title_error': "Error",
                'title_info': "Information",
                'title_conflict': "Settings Conflict",
                'title_start_failed': "Start Failed",
                'title_trigger_error': "Trigger Error",
                'error_cannot_select_self': "Cannot select the application itself.",
                'error_window_closed': "The previously selected window has been closed.",
                'error_select_window_first': "Please select a window first.",
                'error_invalid_handle': "Could not get a valid window handle.",
                'error_cannot_monitor_self': "Cannot select the main application window for monitoring.",
                'error_start_failed': "Failed to start monitoring: \n{e}",
                'error_style_permission': "Could not modify window style, permission denied.\nPlease try running as an Administrator.",
                'error_style_unknown': "Could not modify window style, it might be a system window or an unknown error occurred: {e}",
                'error_set_trigger': "Failed to set trigger for '{name}': \n{e}",
                'error_conflict_header': "Found duplicate trigger settings. Please resolve the conflicts and try again:",
                'conflict_used_for': "is used for",
                'info_window_closed': "The monitored window has been closed. Monitoring stopped automatically.",
                'tray_show_window': "Show Main Window",
                'tray_exit': "Exit",
                'instructions': "Instructions for Use",
                'instructions_label': "After selecting the window, click \"Start Monitoring\". \nDo not use the Task Manager to force exit this program! \nIf forced exit is required, please set the trigger shortcut key",
            }
        }


if __name__ == "__main__":
    def handle_exception(exc_type, exc_value, exc_traceback):
        if "main thread is not in main loop" in str(exc_value): return
        messagebox.showerror("Critical Error", f"An unexpected error occurred:\n{exc_value}")


    tk.Tk.report_callback_exception = handle_exception
    root = tk.Tk()
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    style = ttk.Style(root)
    try:
        style.theme_use('vista')
    except tk.TclError:
        print("Could not find 'vista' theme.")
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

