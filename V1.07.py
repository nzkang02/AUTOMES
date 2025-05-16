from tkinter import ttk, messagebox, scrolledtext, simpledialog, filedialog
import tkinter as tk
import time
import threading
import keyboard
import pyperclip
import json
import os
from datetime import datetime, timedelta
import ttkthemes
import platform
import pandas as pd
import sys

try:
    import pyautogui
    USE_PYAUTOGUI = True
except ImportError:
    USE_PYAUTOGUI = False

try:
    from tooltip import Hovertip
except ImportError:
    class Hovertip:
        def __init__(self, widget, text):
            self.widget = widget
            self.text = text
            self.tip_window = None
            self.widget.bind("<Enter>", self.show_tip)
            self.widget.bind("<Leave>", self.hide_tip)
        def show_tip(self, event=None):
            if self.tip_window or not self.text: return
            x, y, cx, cy = self.widget.bbox("insert")
            x = x + self.widget.winfo_rootx() + 25
            y = y + cy + self.widget.winfo_rooty() + 25
            self.tip_window = tk.Toplevel(self.widget)
            self.tip_window.wm_overrideredirect(True)
            self.tip_window.wm_geometry(f"+{x}+{y}")
            label = ttk.Label(self.tip_window, text=self.text, justify=tk.LEFT,
                              background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                              font=("tahoma", "8", "normal"))
            label.pack(ipadx=1)
        def hide_tip(self, event=None):
            if self.tip_window: self.tip_window.destroy()
            self.tip_window = None
        def bind(self, *args): pass
        def unbind(self, *args): pass

LOG_FILE = "app_log_history.txt"
CONFIG_FILE = "config.json"

TEXTS = {
    "vi": {
        "app_title_login": "TOOL MES - LOGIN",
        "app_title_main": "TOOL MES - KANG V1.07",
        "login_title": "Đăng nhập",
        "username_label": "Tên đăng nhập:",
        "password_label": "Mật khẩu:",
        "login_button": "Đăng nhập",
        "login_success": "Đăng nhập thành công!",
        "login_failed": "Đăng nhập thất bại. Vui lòng kiểm tra tên đăng nhập và mật khẩu.",
        "language_menu": "Ngôn ngữ",
        "language_vietnamese": "Tiếng Việt",
        "language_chinese": "Tiếng Trung",
        "exit_button": "Thoát",
        "log_history_title": "Lịch sử Hoạt động",
        "log_history_window_title": "Lịch sử Hoạt động",
        "log_clear_button": "Xóa Lịch sử",
        "log_cleared_success": "Lịch sử hoạt động đã được xóa.",
        "confirm_clear_log": "Bạn có chắc chắn muốn xóa toàn bộ lịch sử hoạt động không?",
        "file_selection_title": "Chọn File Excel",
        "select_total_file_button": "Chọn File Excel Tổng",
        "select_missing_file_button": "Chọn File Excel Thiếu",
        "total_file_label": "File Tổng:",
        "missing_file_label": "File Thiếu:",
        "filter_button": "Lọc Mã",
        "station_selection_label": "Chọn Trạm:",
        "station_a04": "Trạm A04",
        "station_a07": "Trạm A07",
        "start_paste_button": "Bắt đầu Dán (F8)",
        "pause_paste_button": "Tạm dừng (F8)",
        "stop_paste_button": "Dừng Dán (F9)",
        "paste_speed_label": "Tốc độ dán:",
        "speed_fast": "Nhanh",
        "speed_medium": "Trung bình",
        "speed_slow": "Chậm",
        "speed_ultrafast": "Siêu tốc (Không Delay)",
        "paste_progress_label": "Tiến trình:",
        "estimated_time_label": "Ước tính:",
        "time_remaining_format": "{minutes} phút {seconds} giây",
        "time_calculating": "Đang tính...",
        "time_estimated": "Ước tính: ",
        "codes_found_label": "Tìm thấy {count} mã cho Trạm {station}.",
        "no_codes_found": "Không tìm thấy mã nào cho Trạm {station}.",
        "filter_success": "Lọc thành công!",
        "filter_failed": "Lọc thất bại: {error}",
        "paste_complete": "Hoàn thành dán {count} mã cho Trạm {station}.",
        "paste_stopped": "Quá trình dán đã dừng.",
        "paste_paused": "Quá trình dán đã tạm dừng.",
        "paste_resumed": "Quá trình dán tiếp tục.",
        "please_select_files": "Vui lòng chọn cả hai file Excel trước khi lọc.",
        "please_filter_first": "Vui lòng lọc mã trước khi bắt đầu dán.",
        "reading_excel_error": "Lỗi khi đọc file Excel '{file}': {error}",
        "filter_logic_error": "Lỗi trong logic lọc: {error}",
        "paste_error": "Lỗi trong quá trình dán: {error}",
        "invalid_excel_format": "File Excel '{file}' không có định dạng cột mong đợi.",
        "save_results_button": "Lưu Kết Quả Lọc",
        "save_file_dialog_title": "Lưu File Kết Quả Lọc",
        "save_file_success": "Đã lưu {count} mã lọc vào file '{file}'.",
        "save_file_failed": "Lỗi khi lưu file kết quả: {error}",
        "select_station_to_save": "Vui lòng chọn Trạm (A04 hoặc A07) trước khi lưu kết quả.",
        "no_codes_to_save": "Không có mã nào để lưu cho Trạm đã chọn.",
        "paste_speed_unit": "giây/mã",
        "elapsed_time_label": "Thời gian đã chạy:",
        "elapsed_time_format": "{minutes} phút {seconds} giây",
        "calculating": "Đang tính...",
        "estimated_completion_time": "Hoàn thành dự kiến: {time}",
        "waiting_for_paste": "Đang chờ dán...",
        "log_read_total": "Đã đọc thành công file Tổng: {}",
        "log_read_missing": "Đã đọc thành công file Thiếu: {}",
        "log_total_codes_true": "File Tổng: Tìm thấy {} mã có trạng thái TRUE.",
        "log_processing_missing_codes": "Đang xử lý các mã trong File Thiếu...",
        "log_filter_result_a04": "Kết quả lọc A04: {} mã.",
        "log_filter_result_a07": "Kết quả lọc A07: {} mã.",
        "log_missing_file_empty": "File Thiếu trống hoặc chỉ có header sau khi đọc.",
        "view_filtered_button": "Xem Kết Quả Lọc",
        "view_filtered_window_title": "Kết Quả Lọc - Trạm {}",
        "status_bar_ready": "Sẵn sàng",
        "status_bar_loading": "Đang đọc file {}...",
        "status_bar_filtering": "Đang lọc mã...",
        "status_bar_pasting": "Đang dán mã...",
        "status_bar_paused": "Tạm dừng",
        "status_bar_stopped": "Dừng",
        "status_bar_complete": "Hoàn thành",
        "status_bar_error": "Lỗi: {}",
        "missing_file_empty": "File Thiếu không chứa dữ liệu hoặc chỉ có header. Không có mã nào để lọc.",
        "current_code_label": "Mã hiện tại:",
        "tooltip_total_file": "Đường dẫn đầy đủ đến File Tổng",
        "tooltip_missing_file": "Đường dẫn đầy đủ đến File Thiếu",
        "no_missing_codes_for_station": "Không tìm thấy mã {station} nào có trạm {station} trong file Thiếu.",
        "paste_speed_unit_display": "{speed} giây/mã",
        "paste_speed_ultrafast_display": "Không Delay",
        "elapsed_time_display_paused": "-",
        "auto_switch_station": "Trạm {old_station} không có mã sau lọc, tự động chuyển sang {new_station}.",
        "no_stations_with_codes": "Không có trạm nào có mã sau lọc. Vui lòng kiểm tra file Thiếu và file Tổng.",
        "confirm_exit_pasting": "Quá trình dán đang chạy. Bạn có chắc chắn muốn thoát?",
        "confirm_exit_pasting_title": "Xác nhận thoát"
    },
    "zh": {
        "app_title_login": "MES 代码扫描工具 - 登录",
        "app_title_main": "MES 代码扫描工具",
        "login_title": "登录",
        "username_label": "用户名:",
        "password_label": "密码:",
        "login_button": "登录",
        "login_success": "登录成功！",
        "login_failed": "登录失败。请检查用户名和密码。",
        "language_menu": "语言",
        "language_vietnamese": "越南语",
        "language_chinese": "中文",
        "exit_button": "退出",
        "log_history_title": "操作历史",
        "log_history_window_title": "操作历史",
        "log_clear_button": "清除历史",
        "log_cleared_success": "操作历史已清除。",
        "confirm_clear_log": "您确定要清除所有操作历史吗？",
        "file_selection_title": "选择 Excel 文件",
        "select_total_file_button": "选择总表文件",
        "select_missing_file_button": "选择缺失文件",
        "total_file_label": "总表文件:",
        "missing_file_label": "缺失文件:",
        "filter_button": "过滤代码",
        "station_selection_label": "选择站点:",
        "station_a04": "A04 站点",
        "station_a07": "A07 站点",
        "start_paste_button": "开始粘贴 (F8)",
        "pause_paste_button": "暂停 (F8)",
        "stop_paste_button": "停止粘贴 (F9)",
        "paste_speed_label": "粘贴速度:",
        "speed_fast": "快",
        "speed_medium": "中",
        "speed_slow": "慢",
        "speed_ultrafast": "超快 (无延迟)",
        "paste_progress_label": "进度:",
        "estimated_time_label": "预计时间:",
        "time_remaining_format": "{minutes} 分钟 {seconds} 秒",
        "time_calculating": "计算中...",
        "time_estimated": "预计: ",
        "codes_found_label": "为 {station} 站点找到 {count} 个代码。",
        "no_codes_found": "为 {station} 站点未找到任何代码。",
        "filter_success": "过滤成功！",
        "filter_failed": "过滤失败: {error}",
        "paste_complete": "为 {station} 站点完成粘贴 {count} 个代码。",
        "paste_stopped": "粘贴过程已停止。",
        "paste_paused": "粘贴过程已暂停。",
        "paste_resumed": "粘贴过程继续。",
        "please_select_files": "请在过滤前选择两个 Excel 文件。",
        "please_filter_first": "请在开始粘贴前先过滤代码。",
        "reading_excel_error": "读取 Excel 文件 '{file}' 时出错: {error}",
        "filter_logic_error": "过滤逻辑错误: {error}",
        "paste_error": "粘贴过程中出错: {error}",
        "invalid_excel_format": "Excel 文件 '{file}' 没有预期的列格式。",
        "save_results_button": "保存过滤结果",
        "save_file_dialog_title": "保存过滤结果文件",
        "save_file_success": "已将 {count} 个过滤后的代码保存到文件 '{file}'。",
        "save_file_failed": "保存结果文件时出错: {error}",
        "select_station_to_save": "请在保存结果前选择站点 (A04 或 A07)。",
        "no_codes_to_save": "所选站点没有要保存的代码。",
        "paste_speed_unit": "秒/代码",
        "elapsed_time_label": "已用时间:",
        "elapsed_time_format": "{minutes} 分钟 {seconds} 秒",
        "calculating": "计算中...",
        "estimated_completion_time": "预计完成时间: {time}",
        "waiting_for_paste": "等待粘贴中...",
        "log_read_total": "成功读取总表文件: {}",
        "log_read_missing": "成功读取缺失文件: {}",
        "log_total_codes_true": "总表文件: 找到 {} 个状态为 TRUE 的代码。",
        "log_processing_missing_codes": "正在处理缺失文件中的代码...",
        "log_filter_result_a04": "A04 过滤结果: {} 个代码。",
        "log_filter_result_a07": "A07 过滤结果: {} 个代码。",
        "log_missing_file_empty": "读取后缺失文件为空或只有 header。",
        "view_filtered_button": "查看过滤结果",
        "view_filtered_window_title": "过滤结果 - 站点 {}",
        "status_bar_ready": "准备就绪",
        "status_bar_loading": "正在加载文件 {}...",
        "status_bar_filtering": "正在过滤代码...",
        "status_bar_pasting": "正在粘贴代码...",
        "status_bar_paused": "已暂停",
        "status_bar_stopped": "已停止",
        "status_bar_complete": "完成",
        "status_bar_error": "错误: {}",
        "missing_file_empty": "缺失文件不包含数据或只有header。没有代码可过滤。",
        "current_code_label": "当前代码:",
        "tooltip_total_file": "总表文件的完整路径",
        "tooltip_missing_file": "缺失文件的完整路径",
        "no_missing_codes_for_station": "在缺失文件中未找到任何包含站点 {station} 的代码。",
        "paste_speed_unit_display": "{speed} 秒/代码",
        "paste_speed_ultrafast_display": "无延迟",
        "elapsed_time_display_paused": "-",
        "auto_switch_station": "站点 {old_station} 过滤后无代码，自动切换到 {new_station}。",
        "no_stations_with_codes": "过滤后没有站点包含代码。请检查缺失文件和总表文件。",
        "confirm_exit_pasting": "粘贴过程正在运行。您确定要退出吗？",
        "confirm_exit_pasting_title": "确认退出"
    }
}

PASTE_SPEEDS = {
    "ultrafast": 0.01,
    "fast": 0.05,
    "medium": 0.3,
    "slow": 0.5
}

POST_COPY_SHORT_DELAY = 0.002
POST_PASTE_SHORT_DELAY = 0.002

def get_current_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_action(action):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{get_current_time()}] {action}\n")
    except Exception as e:
        print(f"Lỗi khi ghi log: {e}")

def show_message(title, message, type="info"):
    current_root = tk._get_default_root()
    if current_root:
         if type == "info": current_root.after(0, messagebox.showinfo, title, message)
         elif type == "warning": current_root.after(0, messagebox.showwarning, title, message)
         elif type == "error": current_root.after(0, messagebox.showerror, title, message)
    else:
         if type == "info": messagebox.showinfo(title, message)
         elif type == "warning": messagebox.showwarning(title, message)
         elif type == "error": messagebox.showerror(title, message)

def copy_to_clipboard(text):
    try:
        pyperclip.copy(str(text))
        return True
    except Exception as e:
        log_action(f"Lỗi khi sao chép vào clipboard: {e}")
        return False

def simulate_paste_and_enter():
    if USE_PYAUTOGUI:
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
    else:
        keyboard.press_and_release('ctrl+v')
        keyboard.press_and_release('enter')

def hide_file(filepath):
    if platform.system() == "Windows":
        try:
            import ctypes
            ctypes.windll.kernel32.SetFileAttributesW(filepath, 0x02)
        except Exception as e: pass

def show_file(filepath):
    if platform.system() == "Windows":
        try:
            import ctypes
            ctypes.windll.kernel32.SetFileAttributesW(filepath, 0x80)
        except Exception as e: pass

def save_config(config_data):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
        hide_file(CONFIG_FILE)
    except Exception as e:
        log_action(f"Lỗi khi lưu cấu hình: {e}")

def load_config():
    config_data = {}
    if os.path.exists(CONFIG_FILE):
        try:
            show_file(CONFIG_FILE)
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
        except Exception as e:
            log_action(f"Lỗi khi tải cấu hình: {e}")
            try:
                 os.remove(CONFIG_FILE)
                 log_action("Đã xóa file config bị lỗi.")
            except Exception as e_del:
                 log_action(f"Lỗi khi xóa file config bị lỗi: {e_del}")
        finally:
             hide_file(CONFIG_FILE)
    return config_data

class AutoPasteTool:
    def __init__(self, root):
        self.root = root
        self.current_language = tk.StringVar(value="vi")
        self.config = load_config()

        if 'language' in self.config and self.config['language'] in TEXTS:
             self.current_language.set(self.config['language'])
        else:
            self.current_language.set("vi")

        self.root.resizable(False, False)
        self.username = ""
        self.password = ""
        self.username_entry = None
        self.password_entry = None
        self.login_frame = None
        self.main_app = None

        self._create_login_widgets()
        self._apply_language()

        if os.path.exists(LOG_FILE): hide_file(LOG_FILE)

        self._start_hotkey_listener()

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _on_closing(self):
        lang = self.current_language.get()
        texts = TEXTS[lang]

        if self.main_app and self.main_app.is_pasting:
             if not messagebox.askyesno(texts["confirm_exit_pasting_title"], texts["confirm_exit_pasting"]):
                  return

        if self.main_app:
             self.config['total_file_path'] = self.main_app._full_total_file_path
             self.config['missing_file_path'] = self.main_app._full_missing_file_path
             self.config['paste_speed'] = self.main_app.paste_speed_var.get()
             if self.main_app.is_pasting: self.main_app._stop_paste()

        self.config['language'] = self.current_language.get()
        save_config(self.config)
        log_action("Ứng dụng đóng.")
        self.root.destroy()
        sys.exit(0)

    def _apply_language(self, lang=None):
        if lang: self.current_language.set(lang)
        lang = self.current_language.get()
        texts = TEXTS[lang]

        if self.main_app:
             self.root.title(texts["app_title_main"])
             self.main_app._apply_language(lang)
        else:
            self.root.title(texts["app_title_login"])
            self._create_login_widgets()

        menubar = tk.Menu(self.root)
        menubar.add_cascade(label=texts["language_menu"], menu=self._create_language_menu(menubar))
        history_menu = tk.Menu(menubar, tearoff=0)
        history_menu.add_command(label=texts["log_history_title"], command=self._show_log_history)
        menubar.add_cascade(label=texts["log_history_title"], menu=history_menu)
        self.root.config(menu=menubar)

    def _create_language_menu(self, menubar):
        language_menu = tk.Menu(menubar, tearoff=0)
        language_menu.add_radiobutton(label=TEXTS["vi"]["language_vietnamese"], variable=self.current_language, value="vi", command=lambda: self._apply_language("vi"))
        language_menu.add_radiobutton(label=TEXTS["zh"]["language_chinese"], variable=self.current_language, value="zh", command=lambda: self._apply_language("zh"))
        return language_menu

    def _create_login_widgets(self):
        lang = self.current_language.get()
        texts = TEXTS[lang]

        style = ttkthemes.ThemedStyle(self.root)
        style.set_theme("arc")

        if self.login_frame: self.login_frame.destroy()

        self.login_frame = ttk.Frame(self.root, padding="30")
        self.login_frame.grid(row=0, column=0, padx=30, pady=30, sticky="nsew")

        ttk.Label(self.login_frame, text=texts["login_title"], font=("Arial", 20, "bold")).grid(row=0, column=0, columnspan=2, pady=20)
        ttk.Label(self.login_frame, text=texts["username_label"], font=("Arial", 11)).grid(row=1, column=0, padx=10, pady=8, sticky="w")
        self.username_entry = ttk.Entry(self.login_frame, width=35, font=("Arial", 11))
        self.username_entry.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ttk.Label(self.login_frame, text=texts["password_label"], font=("Arial", 11)).grid(row=2, column=0, padx=10, pady=8, sticky="w")
        self.password_entry = ttk.Entry(self.login_frame, show='*', width=35, font=("Arial", 11))
        self.password_entry.grid(row=2, column=1, padx=10, pady=8, sticky="ew")
        ttk.Button(self.login_frame, text=texts["login_button"], command=self._login, style='TButton').grid(row=3, columnspan=2, padx=10, pady=20, sticky="ew")

        self.login_frame.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

    def _login(self):
        lang = self.current_language.get()
        texts = TEXTS[lang]
        username = self.username_entry.get()
        password = self.password_entry.get()

        if username == "AOI" and password == "6969":
            self.username = username
            self.password = password
            log_action(f"Người dùng '{username}' đăng nhập thành công.")
            show_message(texts["login_success"], texts["login_success"])
            self._show_main_app()
        else:
            log_action(f"Người dùng '{username}' đăng nhập thất bại.")
            show_message(texts["login_failed"], texts["login_failed"], type="error")
            self.password_entry.delete(0, tk.END)

    def _show_main_app(self):
        if self.login_frame: self.login_frame.destroy()
        self.main_app = MainApp(self.root, self.username, self.password, self, self.config)
        self.main_app._apply_language(self.current_language.get())

    def _show_log_history(self):
        lang = self.current_language.get()
        texts = TEXTS[lang]
        history_window = tk.Toplevel(self.root)
        history_window.title(texts["log_history_window_title"])
        history_window.geometry("600x400")
        history_window.transient(self.root)

        style = ttkthemes.ThemedStyle(history_window)
        style.set_theme("arc")

        log_text = scrolledtext.ScrolledText(history_window, wrap=tk.WORD, font=("Arial", 10))
        log_text.pack(padx=10, pady=10, fill="both", expand=True)

        try:
            show_file(LOG_FILE)
            if os.path.exists(LOG_FILE):
                with open(LOG_FILE, "r", encoding="utf-8") as f:
                    log_text.insert(tk.END, f.read())
            else:
                log_text.insert(tk.END, "Chưa có lịch sử hoạt động." if lang == "vi" else "尚无操作历史记录。")
        except Exception as e:
            log_text.insert(tk.END, f"Lỗi khi đọc lịch sử: {e}" if lang == "vi" else f"读取历史记录时出错: {e}")
        finally:
            hide_file(LOG_FILE)

        log_text.config(state=tk.DISABLED)
        clear_button = ttk.Button(history_window, text=texts["log_clear_button"], command=self._clear_log_history)
        clear_button.pack(pady=5)

    def _clear_log_history(self):
        lang = self.current_language.get()
        texts = TEXTS[lang]
        if messagebox.askyesno(texts["log_clear_button"], texts["confirm_clear_log"]):
            try:
                if os.path.exists(LOG_FILE):
                    show_file(LOG_FILE)
                    os.remove(LOG_FILE)
                show_message(texts["log_clear_button"], texts["log_cleared_success"])
                log_action("Đã xóa lịch sử hoạt động.")
                for widget in self.root.winfo_children():
                     if isinstance(widget, tk.Toplevel) and widget.title() == texts["log_history_window_title"]:
                          widget.destroy()
                with open(LOG_FILE, "w", encoding="utf-8") as f: pass
                hide_file(LOG_FILE)
            except Exception as e:
                show_message("Lỗi", f"Không thể xóa file log: {e}", type="error")

    def _start_hotkey_listener(self):
        threading.Thread(target=self._listen_for_hotkeys, daemon=True).start()

    def _listen_for_hotkeys(self):
        try:
            keyboard.add_hotkey('f8', self._on_f8_pressed)
            keyboard.add_hotkey('f9', self._on_f9_pressed)
            keyboard.wait()
        except Exception as e:
            print(f"Lỗi khi lắng nghe hotkey: {e}")
            log_action(f"Lỗi khi lắng nghe hotkey: {e}")

    def _on_f8_pressed(self):
        if self.main_app: self.root.after(0, self.main_app._toggle_paste)

    def _on_f9_pressed(self):
        if self.main_app: self.root.after(0, self.main_app._stop_paste)

class MainApp:
    def __init__(self, root, username, password, parent, config):
        self.root = root
        self.username = username
        self.password = password
        self.parent = parent
        self.config = config

        self.total_file_path = tk.StringVar()
        self.missing_file_path = tk.StringVar()
        self._full_total_file_path = ""
        self._full_missing_file_path = ""

        self.total_df = None
        self.missing_df = None

        self.filtered_codes_a04 = []
        self.filtered_codes_a07 = []

        self.selected_station = tk.StringVar(value="A04")

        self.is_pasting = False
        self.is_paused = False

        self._paste_thread = None
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()

        self.current_code_index = 0
        self.total_codes_to_paste = 0

        self.start_time = None
        self.paused_time = None
        self._total_elapsed_time_at_pause = 0

        self.paste_speed_var = tk.StringVar(value="medium")
        if 'paste_speed' in self.config and self.config['paste_speed'] in PASTE_SPEEDS:
             self.paste_speed_var.set(self.config['paste_speed'])
        self.current_speed_value = PASTE_SPEEDS[self.paste_speed_var.get()]

        self.current_code_display = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="")
        self._is_reading_file = False

        self._create_main_widgets()
        self._load_saved_paths()
        self._update_time_display_periodically()

    def _apply_language(self, lang=None):
        if lang: self.parent.current_language.set(lang)
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]

        self.root.title(texts["app_title_main"])

        self.file_frame.config(text=texts["file_selection_title"])
        self.select_total_button.config(text=texts["select_total_file_button"])
        self.select_missing_button.config(text=texts["select_missing_file_button"])
        self.total_file_label_widget.config(text=texts["total_file_label"])
        self.missing_file_label_widget.config(text=texts["missing_file_label"])
        self.filter_button.config(text=texts["filter_button"])

        self.station_label.config(text=texts["station_selection_label"])
        self.station_a04_radio.config(text=texts["station_a04"])
        self.station_a07_radio.config(text=texts["station_a07"])

        if self.is_pasting and not self.is_paused: self.start_paste_button.config(text=texts["pause_paste_button"])
        else: self.start_paste_button.config(text=texts["start_paste_button"])

        self.stop_paste_button.config(text=texts["stop_paste_button"])
        self.paste_speed_label.config(text=texts["paste_speed_label"])
        self.speed_ultrafast_radio.config(text=texts["speed_ultrafast"])
        self.speed_fast_radio.config(text=texts["speed_fast"])
        self.speed_medium_radio.config(text=texts["speed_medium"])
        self.speed_slow_radio.config(text=texts["speed_slow"])
        self.view_filtered_button.config(text=texts["view_filtered_button"])
        self.save_results_button.config(text=texts["save_results_button"])
        self.paste_progress_label.config(text=texts["paste_progress_label"])
        self.current_code_label_widget.config(text=texts["current_code_label"])
        self.estimated_time_label_widget.config(text=texts["estimated_time_label"])
        self.elapsed_time_label_widget.config(text=texts["elapsed_time_label"])

        self._update_paste_speed_label_text()
        self._update_progress_labels()

        if self._is_reading_file: pass
        elif self.is_pasting:
             if self.is_paused: self._update_status_bar(texts["status_bar_paused"])
             else: self._update_status_bar(texts["status_bar_pasting"])
        elif self.total_df is not None or self.missing_df is not None: self._update_status_bar(texts["status_bar_ready"])
        else: self._update_status_bar(texts["status_bar_ready"])

        if hasattr(self, 'total_file_tooltip'): self.total_file_tooltip.text = texts["tooltip_total_file"]
        if hasattr(self, 'missing_file_tooltip'): self.missing_file_tooltip.text = texts["tooltip_missing_file"]

    def _create_main_widgets(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]

        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)

        self.file_frame = ttk.LabelFrame(main_frame, text=texts["file_selection_title"], padding="15")
        self.file_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.file_frame.columnconfigure(2, weight=1)

        self.select_total_button = ttk.Button(self.file_frame, text=texts["select_total_file_button"], command=self._select_total_file)
        self.select_total_button.grid(row=0, column=0, padx=5, pady=5)
        self.total_file_label_widget = ttk.Label(self.file_frame, text=texts["total_file_label"])
        self.total_file_label_widget.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.total_file_entry = ttk.Entry(self.file_frame, textvariable=self.total_file_path, state='readonly', width=60)
        self.total_file_entry.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        self.total_file_tooltip = Hovertip(self.total_file_entry, text=texts["tooltip_total_file"])

        self.select_missing_button = ttk.Button(self.file_frame, text=texts["select_missing_file_button"], command=self._select_missing_file)
        self.select_missing_button.grid(row=1, column=0, padx=5, pady=5)
        self.missing_file_label_widget = ttk.Label(self.file_frame, text=texts["missing_file_label"])
        self.missing_file_label_widget.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.missing_file_entry = ttk.Entry(self.file_frame, textvariable=self.missing_file_path, state='readonly', width=60)
        self.missing_file_entry.grid(row=1, column=2, padx=5, pady=5, sticky="ew")
        self.missing_file_tooltip = Hovertip(self.missing_file_entry, text=texts["tooltip_missing_file"])

        button_filter_frame = ttk.Frame(self.file_frame)
        button_filter_frame.grid(row=2, column=0, columnspan=3, pady=10)
        self.filter_button = ttk.Button(button_filter_frame, text=texts["filter_button"], command=self._start_filter_thread)
        self.filter_button.grid(row=0, column=0, padx=5)
        self.view_filtered_button = ttk.Button(button_filter_frame, text=texts["view_filtered_button"], command=self._show_filtered_codes)
        self.view_filtered_button.grid(row=0, column=1, padx=5)
        self.view_filtered_button.config(state=tk.DISABLED)
        self.save_results_button = ttk.Button(button_filter_frame, text=texts["save_results_button"], command=self._save_filtered_codes)
        self.save_results_button.grid(row=0, column=2, padx=5)
        self.save_results_button.config(state=tk.DISABLED)

        control_frame = ttk.LabelFrame(main_frame, text="Điều khiển Dán" if lang=="vi" else "粘贴控制", padding="15")
        control_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        control_frame.columnconfigure(1, weight=1)

        self.station_label = ttk.Label(control_frame, text=texts["station_selection_label"])
        self.station_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        station_radio_frame = ttk.Frame(control_frame)
        station_radio_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.station_a04_radio = ttk.Radiobutton(station_radio_frame, text=texts["station_a04"], variable=self.selected_station, value="A04", command=self._update_progress_labels)
        self.station_a04_radio.grid(row=0, column=0, padx=5)
        self.station_a07_radio = ttk.Radiobutton(station_radio_frame, text=texts["station_a07"], variable=self.selected_station, value="A07", command=self._update_progress_labels)
        self.station_a07_radio.grid(row=0, column=1, padx=5)
        self.station_a04_radio.config(state=tk.DISABLED)
        self.station_a07_radio.config(state=tk.DISABLED)

        self.paste_speed_label = ttk.Label(control_frame, text=texts["paste_speed_label"])
        self.paste_speed_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        speed_radio_frame = ttk.Frame(control_frame)
        speed_radio_frame.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.speed_ultrafast_radio = ttk.Radiobutton(speed_radio_frame, text=texts["speed_ultrafast"], variable=self.paste_speed_var, value="ultrafast", command=self._update_paste_speed)
        self.speed_ultrafast_radio.grid(row=0, column=0, padx=5)
        self.speed_fast_radio = ttk.Radiobutton(speed_radio_frame, text=texts["speed_fast"], variable=self.paste_speed_var, value="fast", command=self._update_paste_speed)
        self.speed_fast_radio.grid(row=0, column=1, padx=5)
        self.speed_medium_radio = ttk.Radiobutton(speed_radio_frame, text=texts["speed_medium"], variable=self.paste_speed_var, value="medium", command=self._update_paste_speed)
        self.speed_medium_radio.grid(row=0, column=2, padx=5)
        self.speed_slow_radio = ttk.Radiobutton(speed_radio_frame, text=texts["speed_slow"], variable=self.paste_speed_var, value="slow", command=self._update_paste_speed)
        self.speed_slow_radio.grid(row=0, column=3, padx=5)
        self.paste_speed_unit_label = ttk.Label(speed_radio_frame, text="")
        self.paste_speed_unit_label.grid(row=0, column=4, padx=5)

        button_paste_frame = ttk.Frame(control_frame)
        button_paste_frame.grid(row=2, column=0, columnspan=2, pady=10)
        self.start_paste_button = ttk.Button(button_paste_frame, text=texts["start_paste_button"], command=self._toggle_paste)
        self.start_paste_button.grid(row=0, column=0, padx=5)
        self.stop_paste_button = ttk.Button(button_paste_frame, text=texts["stop_paste_button"], command=self._stop_paste)
        self.stop_paste_button.grid(row=0, column=1, padx=5)
        self.stop_paste_button.config(state=tk.DISABLED)

        progress_frame = ttk.LabelFrame(main_frame, text="Tiến trình" if lang=="vi" else "进度", padding="15")
        progress_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        progress_frame.columnconfigure(1, weight=1)

        self.paste_progress_label = ttk.Label(progress_frame, text=texts["paste_progress_label"])
        self.paste_progress_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.progress_value_label = ttk.Label(progress_frame, text="0/0 (0%)")
        self.progress_value_label.grid(row=0, column=1, padx=5, pady=5, sticky="e")

        self.progressbar = ttk.Progressbar(progress_frame, orient="horizontal", length=400, mode="determinate")
        self.progressbar.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        self.current_code_label_widget = ttk.Label(progress_frame, text=texts["current_code_label"])
        self.current_code_label_widget.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.current_code_value_label = ttk.Label(progress_frame, textvariable=self.current_code_display, font=("Arial", 10, "bold"))
        self.current_code_value_label.grid(row=2, column=1, padx=5, pady=5, sticky="e")

        self.elapsed_time_label_widget = ttk.Label(progress_frame, text=texts["elapsed_time_label"])
        self.elapsed_time_label_widget.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.elapsed_time_value_label = ttk.Label(progress_frame, text="0 phút 0 giây" if lang=="vi" else "0 分钟 0 秒")
        self.elapsed_time_value_label.grid(row=3, column=1, padx=5, pady=5, sticky="e")

        self.estimated_time_label_widget = ttk.Label(progress_frame, text=texts["estimated_time_label"])
        self.estimated_time_label_widget.grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.estimated_time_value_label = ttk.Label(progress_frame, text=texts["waiting_for_paste"])
        self.estimated_time_value_label.grid(row=4, column=1, padx=5, pady=5, sticky="e")

        self.status_bar = ttk.Label(main_frame, textvariable=self.status_text, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))

        main_frame.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self._update_paste_speed()
        self._update_status_bar(texts["status_bar_ready"])

    def _update_status_bar(self, message):
        self.root.after(0, self.status_text.set, message)

    def _load_saved_paths(self):
         if 'total_file_path' in self.config and os.path.exists(self.config['total_file_path']):
              self._set_total_file_path(self.config['total_file_path'])
         if 'missing_file_path' in self.config and os.path.exists(self.config['missing_file_path']):
              self._set_missing_file_path(self.config['missing_file_path'])

    def _set_total_file_path(self, filepath):
         self._full_total_file_path = filepath
         self.total_file_path.set(os.path.basename(filepath))
         if hasattr(self, 'total_file_tooltip'): self.total_file_tooltip.text = filepath

    def _set_missing_file_path(self, filepath):
         self._full_missing_file_path = filepath
         self.missing_file_path.set(os.path.basename(filepath))
         if hasattr(self, 'missing_file_tooltip'): self.missing_file_tooltip.text = filepath

    def _select_total_file(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]
        filepath = filedialog.askopenfilename(
            title=texts["select_total_file_button"],
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filepath:
            self._set_total_file_path(filepath)
            self._update_status_bar(texts["status_bar_ready"])

    def _select_missing_file(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]
        filepath = filedialog.askopenfilename(
            title=texts["select_missing_file_button"],
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filepath:
            self._set_missing_file_path(filepath)
            self._update_status_bar(texts["status_bar_ready"])

    def _read_excel_threaded(self, filepath, file_type):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]
        df = None
        error = None
        try:
            self.root.after(0, self._update_status_bar, texts["status_bar_loading"].format(os.path.basename(filepath)))
            self._is_reading_file = True

            if file_type == "Tổng": df = pd.read_excel(filepath, engine='openpyxl')
            elif file_type == "Thiếu": df = pd.read_excel(filepath, engine='openpyxl', header=None)
            else: df = pd.read_excel(filepath, engine='openpyxl')

            if df is not None and df.empty:
                 error = "File trống."
                 message = texts["missing_file_empty"] if file_type == "Thiếu" else texts["reading_excel_error"].format(file=os.path.basename(filepath), error="File trống.")
                 self.root.after(0, show_message, "Cảnh báo", message, "warning")
                 log_action(f"File '{os.path.basename(filepath)}' trống.")
                 df = None

        except FileNotFoundError:
            error = "File not found"
            self.root.after(0, show_message, "Lỗi", texts["reading_excel_error"].format(file=os.path.basename(filepath), error="File not found"), "error")
            log_action(f"Lỗi: File không tìm thấy - {os.path.basename(filepath)}")
            df = None
        except pd.errors.EmptyDataError:
             error = "File trống"
             self.root.after(0, show_message, "Lỗi", texts["reading_excel_error"].format(file=os.path.basename(filepath), error="File trống"), "error")
             log_action(f"Lỗi: File trống - {os.path.basename(filepath)}")
             df = None
        except Exception as e:
            error = str(e)
            self.root.after(0, show_message, "Lỗi", texts["reading_excel_error"].format(file=os.path.basename(filepath), error=str(e)), "error")
            log_action(f"Lỗi khi đọc file Excel '{os.path.basename(filepath)}': {e}")
            df = None
        finally:
            self.root.after(0, self._handle_read_result, file_type, df, error)

    def _handle_read_result(self, file_type, df, error):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]

        if file_type == "Tổng":
            self.total_df = df
            if error or df is None:
                 self._is_reading_file = False
                 self._update_status_bar(texts["status_bar_error"].format(f"Read total file error: {error if error else 'Empty file'}"))
                 self._update_progress_labels()
                 self.filter_button.config(state=tk.NORMAL)
                 self._reset_filtered_codes()
                 self._reset_progress()
                 self._update_station_radio_states()
                 return

            log_action(texts["log_read_total"].format(os.path.basename(self._full_total_file_path)))

            if self._full_missing_file_path and os.path.exists(self._full_missing_file_path):
                 read_missing_thread = threading.Thread(target=self._read_excel_threaded, args=(self._full_missing_file_path, "Thiếu"))
                 read_missing_thread.daemon = True
                 read_missing_thread.start()
            else:
                 self._is_reading_file = False
                 show_message(texts["filter_button"], texts["please_select_files"], type="warning")
                 self._reset_filtered_codes()
                 self._reset_progress()
                 self._update_progress_labels()
                 self._update_status_bar(texts["status_bar_ready"])
                 self.filter_button.config(state=tk.NORMAL)
                 self._update_station_radio_states()

        elif file_type == "Thiếu":
            self.missing_df = df
            self._is_reading_file = False

            if error or df is None:
                 self._update_status_bar(texts["status_bar_error"].format(f"Read missing file error: {error if error else 'Empty file'}"))
                 self._update_progress_labels()
                 self.filter_button.config(state=tk.NORMAL)
                 self._reset_filtered_codes()
                 self._reset_progress()
                 self._update_station_radio_states()
                 return

            log_action(texts["log_read_missing"].format(os.path.basename(self._full_missing_file_path)))

            self._perform_filter_logic()

    def _start_filter_thread(self):
         lang = self.parent.current_language.get()
         texts = TEXTS[lang]
         total_file = self._full_total_file_path
         missing_file = self._full_missing_file_path

         if not total_file or not missing_file or not os.path.exists(total_file) or not os.path.exists(missing_file):
            show_message(texts["filter_button"], texts["please_select_files"], type="warning")
            self._reset_filtered_codes()
            self._reset_progress()
            self._update_progress_labels()
            self._update_status_bar(texts["status_bar_ready"])
            self._update_station_radio_states()
            return

         self.filter_button.config(state=tk.DISABLED)
         self._reset_filtered_codes()
         self._reset_progress()
         self._update_progress_labels()

         read_total_thread = threading.Thread(target=self._read_excel_threaded, args=(total_file, "Tổng"))
         read_total_thread.daemon = True
         read_total_thread.start()

    def _perform_filter_logic(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]

        if self.total_df is None or self.missing_df is None:
             self.filter_button.config(state=tk.NORMAL)
             self._update_status_bar(texts["status_bar_ready"])
             self._update_station_radio_states()
             return

        self._update_status_bar(texts["status_bar_filtering"])

        mes_code_col_name_total = "PCBID SN"
        upload_status_col_total_index = 15
        upload_true_value_check = True

        barcode_col_missing_index = 0
        station_col_missing_index = 1
        station_a04_value = "A04"
        station_a07_value = "A07"

        try:
            total_codes_true_set = set()
            try:
                mes_code_col_total_index = self.total_df.columns.get_loc(mes_code_col_name_total)
            except KeyError:
                 show_message("Lỗi", texts["invalid_excel_format"].format(file=os.path.basename(self._full_total_file_path)) + f"\nKhông tìm thấy cột '{mes_code_col_name_total}' trong file Tổng.", type="error")
                 log_action(f"Lỗi: Không tìm thấy cột '{mes_code_col_name_total}' trong File Tổng '{os.path.basename(self._full_total_file_path)}'.")
                 self._update_progress_labels()
                 self._update_status_bar(texts["status_bar_error"].format(f"Missing column '{mes_code_col_name_total}'"))
                 self.filter_button.config(state=tk.NORMAL)
                 self._update_station_radio_states()
                 return

            if self.total_df.shape[1] <= upload_status_col_total_index:
                 show_message("Lỗi", texts["invalid_excel_format"].format(file=os.path.basename(self._full_total_file_path)) + f"\nFile Tổng không có đủ số cột cần thiết để kiểm tra trạng thái tại cột {upload_status_col_total_index + 1}.", type="error")
                 log_action(f"Lỗi: File Tổng '{os.path.basename(self._full_total_file_path)}' không đủ số cột. Yêu cầu ít nhất {upload_status_col_total_index + 1} cột.")
                 self._update_progress_labels()
                 self._update_status_bar(texts["status_bar_error"].format("Total file missing status column"))
                 self.filter_button.config(state=tk.NORMAL)
                 self._update_station_radio_states()
                 return

            for index, row in self.total_df.iloc[1:].iterrows():
                try:
                     if len(row) > upload_status_col_total_index:
                         status_value = row[upload_status_col_total_index]
                         if (isinstance(status_value, bool) and status_value is True) or \
                            (isinstance(status_value, str) and str(status_value).strip().upper() == "TRUE"):

                              code = row[mes_code_col_total_index]
                              if pd.notna(code):
                                   cleaned_code = str(code).strip().replace('\n', '').replace('\r', '')
                                   if cleaned_code: total_codes_true_set.add(cleaned_code.upper())
                except Exception as e:
                     log_action(f"Lỗi xử lý dòng {index+2} trong File Tổng (lọc trạng thái TRUE): {e}. Dữ liệu dòng: {list(row)[:upload_status_col_total_index+2]}")

            log_action(texts["log_total_codes_true"].format(len(total_codes_true_set)))

            if not total_codes_true_set:
                 show_message("Thông báo", "Không có mã nào có trạng thái TRUE trong File Tổng. Không có mã nào để lọc hoặc dán.", type="info")
                 self._update_progress_labels()
                 self._update_status_bar(texts["status_bar_ready"])
                 self.filter_button.config(state=tk.NORMAL)
                 self._update_station_radio_states()
                 return

            missing_codes_a04_set = set()
            missing_codes_a07_set = set()
            log_action(texts["log_processing_missing_codes"])

            if self.missing_df.shape[1] <= max(barcode_col_missing_index, station_col_missing_index):
                 show_message("Lỗi", texts["invalid_excel_format"].format(file=os.path.basename(self._full_missing_file_path)) + f"\nFile Thiếu không có đủ số cột cần thiết (yêu cầu ít nhất {max(barcode_col_missing_index, station_col_missing_index) + 1} cột).", type="error")
                 log_action(f"Lỗi: File Thiếu '{os.path.basename(self._full_missing_file_path)}' không đủ số cột. Yêu cầu ít nhất {max(barcode_col_missing_index, station_col_missing_index) + 1} cột.")
                 self._update_progress_labels()
                 self._update_status_bar(texts["status_bar_error"].format("Missing file missing code/station columns"))
                 self.filter_button.config(state=tk.NORMAL)
                 self._update_station_radio_states()
                 return

            if not self.missing_df.empty:
                 for index, row in self.missing_df.iloc[1:].iterrows():
                    try:
                        if len(row) > max(barcode_col_missing_index, station_col_missing_index):
                            barcode_raw = row[barcode_col_missing_index]
                            station_raw = row[station_col_missing_index]

                            code = str(barcode_raw).strip().replace('\n', '').replace('\r', '').upper() if pd.notna(barcode_raw) else None
                            station = str(station_raw).strip().upper() if pd.notna(station_raw) else ""

                            if code:
                                if station == str(station_a04_value).upper(): missing_codes_a04_set.add(code)
                                elif station == str(station_a07_value).upper(): missing_codes_a07_set.add(code)

                    except Exception as e:
                        log_action(f"Lỗi xử lý dòng {index+2} trong File Thiếu: {e}. Dữ liệu dòng: {list(row)[:max(barcode_col_missing_index, station_col_missing_index)+2]}")
            else:
                 log_action(texts["log_missing_file_empty"])
                 show_message("Thông báo", texts["missing_file_empty"], type="info")
                 self._update_station_radio_states()
                 self._update_progress_labels()
                 self._update_status_bar(texts["status_bar_ready"])
                 self.filter_button.config(state=tk.NORMAL)
                 return


            log_action(f"File Thiếu: Đọc được {len(missing_codes_a04_set)} mã cho A04 và {len(missing_codes_a07_set)} mã cho A07 từ các dòng có trạm tương ứng.")

            count_a04 = 0
            count_a07 = 0

            if missing_codes_a04_set:
                 self.filtered_codes_a04 = list(total_codes_true_set - missing_codes_a04_set)
                 self.filtered_codes_a04.sort()
                 count_a04 = len(self.filtered_codes_a04)
                 log_action(texts["log_filter_result_a04"].format(count_a04))
            else:
                 log_action(texts["no_missing_codes_for_station"].format(station="A04"))

            if missing_codes_a07_set:
                 self.filtered_codes_a07 = list(total_codes_true_set - missing_codes_a07_set)
                 self.filtered_codes_a07.sort()
                 count_a07 = len(self.filtered_codes_a07)
                 log_action(texts["log_filter_result_a07"].format(count_a07))
            else:
                 log_action(texts["no_missing_codes_for_station"].format(station="A07"))

            filter_summary = texts["filter_success"]
            if count_a04 > 0: filter_summary += "\n" + texts["codes_found_label"].format(count=count_a04, station="A04")
            else: filter_summary += "\n" + texts["no_codes_found"].format(station="A04")
            if count_a07 > 0: filter_summary += "\n" + texts["codes_found_label"].format(count=count_a07, station="A07")
            else: filter_summary += "\n" + texts["no_codes_found"].format(station="A07")

            show_message(texts["filter_button"], filter_summary)
            log_action(f"Kết quả lọc: {filter_summary.replace('\n', ' | ')}")

            if count_a04 > 0 or count_a07 > 0:
                 self.view_filtered_button.config(state=tk.NORMAL)
                 self.save_results_button.config(state=tk.NORMAL)
            else:
                 self.view_filtered_button.config(state=tk.DISABLED)
                 self.save_results_button.config(state=tk.DISABLED)

            self._update_station_radio_states()

            selected_station = self.selected_station.get()
            if selected_station == "A04": self.total_codes_to_paste = len(self.filtered_codes_a04)
            elif selected_station == "A07": self.total_codes_to_paste = len(self.filtered_codes_a07)
            else: self.total_codes_to_paste = 0

            self._update_progress_labels()
            self._update_status_bar(texts["status_bar_ready"])

        except Exception as e:
            show_message("Lỗi Lọc", texts["filter_failed"].format(error=str(e)), type="error")
            log_action(f"Lỗi không xác định trong logic lọc: {e}")
            self._reset_filtered_codes()
            self._reset_progress()
            self._update_progress_labels()
            self._update_status_bar(texts["status_bar_error"].format(f"Filter general error: {e}"))
            self._update_station_radio_states()
        finally:
            self.filter_button.config(state=tk.NORMAL)

    def _reset_filtered_codes(self):
        self.filtered_codes_a04 = []
        self.filtered_codes_a07 = []
        if hasattr(self, 'view_filtered_button'): self.view_filtered_button.config(state=tk.DISABLED)
        if hasattr(self, 'save_results_button'): self.save_results_button.config(state=tk.DISABLED)

    def _update_station_radio_states(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]

        a04_has_codes = len(self.filtered_codes_a04) > 0
        a07_has_codes = len(self.filtered_codes_a07) > 0

        if a04_has_codes: self.station_a04_radio.config(state=tk.NORMAL)
        else: self.station_a04_radio.config(state=tk.DISABLED)

        if a07_has_codes: self.station_a07_radio.config(state=tk.NORMAL)
        else: self.station_a07_radio.config(state=tk.DISABLED)

        current_selection = self.selected_station.get()

        if not a04_has_codes and not a07_has_codes:
            self.start_paste_button.config(state=tk.DISABLED)
            show_message("Thông báo", texts["no_stations_with_codes"], type="info")
            log_action("Không có trạm nào có mã sau lọc.")

        elif current_selection == "A04" and not a04_has_codes and a07_has_codes:
             self.selected_station.set("A07")
             show_message("Thông báo", texts["auto_switch_station"].format(old_station="A04", new_station="A07"), type="info")
             log_action(texts["auto_switch_station"].format(old_station="A04", new_station="A07"))
             self.start_paste_button.config(state=tk.NORMAL)
        elif current_selection == "A07" and not a07_has_codes and a04_has_codes:
             self.selected_station.set("A04")
             show_message("Thông báo", texts["auto_switch_station"].format(old_station="A07", new_station="A04"), type="info")
             log_action(texts["auto_switch_station"].format(old_station="A07", new_station="A04"))
             self.start_paste_button.config(state=tk.NORMAL)
        elif (current_selection == "A04" and a04_has_codes) or (current_selection == "A07" and a07_has_codes):
             self.start_paste_button.config(state=tk.NORMAL)
        else:
             self.start_paste_button.config(state=tk.DISABLED)

        selected_station_after_update = self.selected_station.get()
        if selected_station_after_update == "A04": self.total_codes_to_paste = len(self.filtered_codes_a04)
        elif selected_station_after_update == "A07": self.total_codes_to_paste = len(self.filtered_codes_a07)
        else: self.total_codes_to_paste = 0

        self._update_progress_labels()

    def _update_paste_speed(self):
        self.current_speed_value = PASTE_SPEEDS[self.paste_speed_var.get()]
        self._update_paste_speed_label_text()
        self._update_progress_labels()

    def _update_paste_speed_label_text(self):
         lang = self.parent.current_language.get()
         texts = TEXTS[lang]
         if self.paste_speed_var.get() == "ultrafast": self.paste_speed_unit_label.config(text=texts["paste_speed_ultrafast_display"])
         else: self.paste_speed_unit_label.config(text=texts["paste_speed_unit_display"].format(speed=self.current_speed_value))

    def _toggle_paste(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]
        selected_station = self.selected_station.get()
        codes_to_paste = self.filtered_codes_a04 if selected_station == "A04" else self.filtered_codes_a07

        if not codes_to_paste:
             if not self.filtered_codes_a04 and not self.filtered_codes_a07: show_message(texts["start_paste_button"], texts["please_filter_first"], type="warning")
             else: show_message(texts["start_paste_button"], texts["no_codes_found"].format(station=selected_station), type="warning")
             self._update_status_bar(texts["status_bar_ready"])
             return

        if self.is_pasting and not self.is_paused:
            self.is_paused = True
            self._pause_event.set()
            self.paused_time = time.time()
            if self.start_time is not None: self._total_elapsed_time_at_pause += (self.paused_time - self.start_time)
            self.start_time = None
            self.root.after(0, lambda: self.start_paste_button.config(text=texts["start_paste_button"]))
            self.root.after(0, self._update_progress_labels)
            self.root.after(0, lambda: self._update_status_bar(texts["status_bar_paused"]))
            log_action(f"Tạm dừng quá trình dán cho Trạm {self.selected_station.get()}. Đã chạy: {str(timedelta(seconds=int(self._total_elapsed_time_at_pause)))}")

        elif self.is_pasting and self.is_paused:
            self.is_paused = False
            self._pause_event.clear()
            self.start_time = time.time()
            self.paused_time = None
            self.root.after(0, lambda: self.start_paste_button.config(text=texts["pause_paste_button"]))
            self.root.after(0, self._update_progress_labels)
            self.root.after(0, lambda: self._update_status_bar(texts["status_bar_pasting"]))
            log_action(f"Tiếp tục quá trình dán cho Trạm {self.selected_station.get()}.")

        elif not self.is_pasting:
            self._reset_progress_for_new_paste()
            self.total_codes_to_paste = len(codes_to_paste)
            self.is_pasting = True
            self.is_paused = False
            self._stop_event.clear()
            self._pause_event.clear()
            self.start_time = time.time()

            self.root.after(0, lambda: self.start_paste_button.config(text=texts["pause_paste_button"]))
            self.root.after(0, lambda: self.stop_paste_button.config(state=tk.NORMAL))

            self._paste_thread = threading.Thread(target=self._perform_paste, args=(codes_to_paste,))
            self._paste_thread.daemon = True
            self._paste_thread.start()

            log_action(f"Bắt đầu quá trình dán {self.total_codes_to_paste} mã cho Trạm {selected_station}.")
            self._update_progress_labels()
            self._update_status_bar(texts["status_bar_pasting"])

    def _stop_paste(self):
        if self.is_pasting:
            self._stop_event.set()
            self._pause_event.clear()
            lang = self.parent.current_language.get()
            texts = TEXTS[lang]
            log_action(f"Đã yêu cầu dừng quá trình dán.")
            self.root.after(0, lambda: self.start_paste_button.config(text=texts["start_paste_button"]))
            self.root.after(0, lambda: self.stop_paste_button.config(state=tk.DISABLED))
            self.root.after(0, lambda: self._update_status_bar(texts["status_bar_stopped"]))

    def _perform_paste(self, codes):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]
        i = self.current_code_index

        try:
            while i < len(codes):
                if self._stop_event.is_set():
                    break

                while self._pause_event.is_set():
                    time.sleep(0.05)
                    if self._stop_event.is_set():
                        break

                if self._stop_event.is_set():
                     break

                if not self._pause_event.is_set() and self.is_pasting:
                     self.root.after(0, self._update_progress_labels)
                     self.root.after(0, lambda: self._update_status_bar(texts["status_bar_pasting"]))

                     code = codes[i]
                     self.root.after(0, lambda c=code: self.current_code_display.set(c))

                     if not copy_to_clipboard(code):
                         log_action(f"Lỗi khi sao chép mã '{code}' vào clipboard. Dừng dán.")
                         self._stop_event.set()
                         break

                     time.sleep(POST_COPY_SHORT_DELAY)

                     try:
                         simulate_paste_and_enter()
                         log_action(f"Đã dán mã: {code}")
                     except Exception as e:
                         self.root.after(0, lambda: show_message("Lỗi Dán", texts["paste_error"].format(error=f"Lỗi mô phỏng phím: {e}"), type="error"))
                         log_action(f"Lỗi khi mô phỏng dán/enter cho mã '{code}': {e}. Dừng dán.")
                         self._stop_event.set()
                         break

                     self.current_code_index = i + 1

                     if (i + 1) < len(codes):
                          sleep_time = self.current_speed_value
                          if self.paste_speed_var.get() == "ultrafast":
                               time.sleep(POST_PASTE_SHORT_DELAY)
                          else:
                               start_sleep = time.time()
                               while (time.time() - start_sleep) < sleep_time:
                                    if self._pause_event.is_set() or self._stop_event.is_set(): break
                                    time.sleep(0)

                          if self._stop_event.is_set() or self._pause_event.is_set(): continue

                     i += 1

            log_action("Thread dán kết thúc vòng lặp.")
            self.root.after(0, self._finalize_paste, len(codes), self._stop_event.is_set())

        except Exception as e:
            self.root.after(0, lambda: show_message("Lỗi Dán", texts["paste_error"].format(error=f"Lỗi không xác định: {e}"), type="error"))
            log_action(f"Lỗi không xác định trong thread dán: {e}")
            self._stop_event.set()
            self.root.after(0, self._finalize_paste, len(codes), True)

    def _finalize_paste(self, total_codes_at_start, was_stopped):
         lang = self.parent.current_language.get()
         texts = TEXTS[lang]

         self.is_pasting = False
         self.is_paused = False
         self.start_time = None
         self.paused_time = None
         self._total_elapsed_time_at_pause = 0
         self.current_code_display.set("")

         self._update_progress_labels_after_thread_exit(was_stopped)

         self.start_paste_button.config(text=texts["start_paste_button"])
         self.stop_paste_button.config(state=tk.DISABLED)

         if was_stopped:
              show_message(texts["stop_paste_button"], texts["paste_stopped"])
              log_action(f"Quá trình dán đã dừng bởi người dùng hoặc do lỗi sau khi dán {self.current_code_index} mã.")
              self._update_status_bar(texts["status_bar_stopped"])
         else:
              show_message(texts["start_paste_button"], texts["paste_complete"].format(count=self.current_code_index, station=self.selected_station.get()))
              log_action(f"Hoàn thành dán {self.current_code_index} mã cho Trạm {self.selected_station.get()}.")
              self._update_status_bar(texts["status_bar_complete"])

         self.current_code_index = 0

    def _update_time_display_periodically(self):
        self._update_progress_labels()
        self.root.after(1000, self._update_time_display_periodically)

    def _update_progress_labels(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]
        total_codes_after_filter = self.total_codes_to_paste
        codes_processed = self.current_code_index

        if total_codes_after_filter > 0:
            progress_percentage = (codes_processed / total_codes_after_filter) * 100
            self.progressbar['value'] = progress_percentage
            self.progress_value_label.config(text=f"{codes_processed}/{total_codes_after_filter} ({progress_percentage:.1f}%)")
        else:
            self.progressbar['value'] = 0
            self.progress_value_label.config(text="0/0 (0%)")

        if self.is_pasting and not self.is_paused and self.start_time is not None:
            current_elapsed_time = self._total_elapsed_time_at_pause + (time.time() - self.start_time)
            elapsed_minutes = int(current_elapsed_time // 60)
            elapsed_seconds = int(current_elapsed_time % 60)
            self.elapsed_time_value_label.config(text=texts["elapsed_time_format"].format(minutes=elapsed_minutes, seconds=elapsed_seconds))

            if self.current_speed_value is not None and self.current_speed_value >= 0:
                 codes_remaining = total_codes_after_filter - codes_processed
                 if codes_remaining > 0 and self.current_speed_value > 0: # Chỉ tính ước tính nếu có mã còn lại và tốc độ > 0
                      estimated_remaining_time = codes_remaining * self.current_speed_value
                      estimated_completion_timestamp = time.time() + estimated_remaining_time
                      completion_datetime = datetime.fromtimestamp(estimated_completion_timestamp)
                      completion_time_str = completion_datetime.strftime("%H:%M:%S")
                      self.estimated_time_value_label.config(text=texts["estimated_completion_time"].format(time=completion_time_str))
                 elif codes_processed == total_codes_after_filter:
                      self.estimated_time_value_label.config(text="Hoàn thành!" if lang=="vi" else "完成！")
                 else: # Bao gồm cả trường hợp tốc độ = 0 (ultrafast) hoặc codes_remaining = 0
                      self.estimated_time_value_label.config(text=texts["calculating"])
            else: self.estimated_time_value_label.config(text=texts["calculating"])


        elif self.is_paused:
             elapsed_minutes = int(self._total_elapsed_time_at_pause // 60)
             elapsed_seconds = int(self._total_elapsed_time_at_pause % 60)
             self.elapsed_time_value_label.config(text=texts["elapsed_time_format"].format(minutes=elapsed_minutes, seconds=elapsed_seconds))
             self.estimated_time_value_label.config(text=texts["paste_paused"])
        else:
            self.elapsed_time_value_label.config(text="0 phút 0 giây" if lang=="vi" else "0 分钟 0 秒")
            self.estimated_time_value_label.config(text=texts["waiting_for_paste"])

    def _update_progress_labels_after_thread_exit(self, was_stopped):
         lang = self.parent.current_language.get()
         texts = TEXTS[lang]
         total_codes_after_filter = self.total_codes_to_paste
         codes_processed = self.current_code_index

         if total_codes_after_filter > 0:
              progress_percentage = (codes_processed / total_codes_after_filter) * 100
              self.progressbar['value'] = progress_percentage
              self.progress_value_label.config(text=f"{codes_processed}/{total_codes_after_filter} ({progress_percentage:.1f}%)")
         else:
              self.progressbar['value'] = 0
              self.progress_value_label.config(text="0/0 (0%)")

         elapsed_minutes = int(self._total_elapsed_time_at_pause // 60)
         elapsed_seconds = int(self._total_elapsed_time_at_pause % 60)
         self.elapsed_time_value_label.config(text=texts["elapsed_time_format"].format(minutes=elapsed_minutes, seconds=elapsed_seconds))

         if was_stopped: self.estimated_time_value_label.config(text=texts["paste_stopped"])
         else: self.estimated_time_value_label.config(text="Hoàn thành!" if lang=="vi" else "完成！")

         self.current_code_display.set("")

    def _reset_progress_for_new_paste(self):
        self.current_code_index = 0
        self.start_time = None
        self.paused_time = None
        self._total_elapsed_time_at_pause = 0
        self.is_pasting = False
        self.is_paused = False
        self._stop_event.clear()
        self._pause_event.clear()
        self.current_code_display.set("")

    def _reset_progress(self):
        self.current_code_index = 0
        self.total_codes_to_paste = 0
        self.start_time = None
        self.paused_time = None
        self._total_elapsed_time_at_pause = 0
        self.is_pasting = False
        self.is_paused = False
        self._stop_event.clear()
        self._pause_event.clear()

        self.progressbar['value'] = 0
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]
        self.progress_value_label.config(text="0/0 (0%)")
        self.elapsed_time_value_label.config(text="0 phút 0 giây" if lang=="vi" else "0 分钟 0 秒")
        self.estimated_time_value_label.config(text=texts["waiting_for_paste"])
        self.current_code_display.set("")

        if hasattr(self, 'start_paste_button'):
             self.start_paste_button.config(text=texts["start_paste_button"])
        if hasattr(self, 'stop_paste_button'):
             self.stop_paste_button.config(state=tk.DISABLED)


    def _show_filtered_codes(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]

        selected_station = self.selected_station.get()
        codes_to_show = self.filtered_codes_a04 if selected_station == "A04" else self.filtered_codes_a07

        if not codes_to_show:
             show_message(texts["view_filtered_button"], texts["no_codes_to_save"], type="warning")
             return

        view_window = tk.Toplevel(self.root)
        view_window.title(texts["view_filtered_window_title"].format(selected_station))
        view_window.geometry("300x400")
        view_window.transient(self.root)

        style = ttkthemes.ThemedStyle(view_window)
        style.set_theme("arc")

        code_list_text = scrolledtext.ScrolledText(view_window, wrap=tk.WORD, font=("Arial", 10))
        code_list_text.pack(padx=10, pady=10, fill="both", expand=True)

        for code in codes_to_show: code_list_text.insert(tk.END, str(code) + "\n")

        code_list_text.config(state=tk.DISABLED)
        ttk.Label(view_window, text=texts["codes_found_label"].format(count=len(codes_to_show), station=selected_station), font=("Arial", 10)).pack(pady=(0, 10))


    def _save_filtered_codes(self):
        lang = self.parent.current_language.get()
        texts = TEXTS[lang]

        selected_station = self.selected_station.get()
        codes_to_save = self.filtered_codes_a04 if selected_station == "A04" else self.filtered_codes_a07

        if not codes_to_save:
             show_message(texts["save_results_button"], texts["no_codes_to_save"], type="warning")
             return

        filepath = filedialog.asksaveasfilename(
            title=texts["save_file_dialog_title"],
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
            initialfile=f"Filtered_MES_Codes_{selected_station}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not filepath: return

        try:
            df_result = pd.DataFrame(codes_to_save, columns=['Filtered MES Code'])
            df_result.to_excel(filepath, index=False)

            show_message(texts["save_results_button"], texts["save_file_success"].format(count=len(codes_to_save), file=os.path.basename(filepath)))
            log_action(f"Đã lưu {len(codes_to_save)} mã lọc cho Trạm {selected_station} vào file '{os.path.basename(filepath)}'.")
            self._update_status_bar(texts["status_bar_ready"])

        except Exception as e:
            show_message("Lỗi Lưu File", texts["save_file_failed"].format(error=str(e)), type="error")
            log_action(f"Lỗi khi lưu file kết quả cho Trạm {selected_station}: {e}")
            self._update_status_bar(texts["status_bar_error"].format(f"Save file error: {e}"))

if __name__ == "__main__":
    root = ttkthemes.ThemedTk()
    root.set_theme("arc")

    try: root.attributes('-topmost', True)
    except Exception as e:
        print(f"Warning: Could not set window to be always on top: {e}")
        log_action(f"Warning: Could not set window to be always on top: {e}")

    app = AutoPasteTool(root)
    root.mainloop()
