import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import time
import pandas as pd
import re
import pygame

# 初始化 Pygame 混音器
pygame.mixer.init()


# 工具提示类
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event):
        x = event.widget.winfo_rootx() + 20
        y = event.widget.winfo_rooty() + 20
        self.tooltip = tk.Toplevel(event.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip, text=self.text, background="yellow", relief="solid", borderwidth=1)
        label.pack()

    def hide(self, event):
        if self.tooltip:
            self.tooltip.destroy()
        self.tooltip = None


# 全局变量
df = None
time_data = []
light_labels = []
number_labels = []
frame_dict = {}
light_data = []
index = 0
is_running = threading.Event()
current_file_path = ""
selected_sheet = ""
main_window = None


def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        sheet_names = pd.ExcelFile(file_path).sheet_names
        return file_path, sheet_names
    return None, []


def on_ok():
    global current_file_path, selected_sheet
    current_file_path = file_path_var.get()
    selected_sheet = sheet_selector.get()
    if current_file_path and selected_sheet:
        if re.search(r'.*history.*', selected_sheet, re.IGNORECASE):
            show_warning_dialog()
        else:
            dialog.withdraw()  # 隐藏对话框
            create_main_window()
    else:
        messagebox.showerror("Error", "Please select a file and a sheet.")


def on_cancel():
    root.destroy()


def show_warning_dialog():
    warning_dialog = tk.Toplevel(root)
    warning_dialog.title("Warning")

    warning_label = tk.Label(warning_dialog, text="This Sheet does not have the information of animation.", fg="red",
                             font=("Helvetica", 14))
    warning_label.pack(pady=10)

    button_frame = tk.Frame(warning_dialog)
    button_frame.pack(pady=10)

    yes_button = tk.Button(button_frame, text="Yes", command=warning_dialog.destroy)
    yes_button.pack(side=tk.LEFT, padx=5)

    no_button = tk.Button(button_frame, text="No", command=lambda: on_warning_no(warning_dialog))
    no_button.pack(side=tk.LEFT, padx=5)

    ToolTip(yes_button, "Close this dialog")
    ToolTip(no_button, "Close this dialog and load the sheet")


def on_warning_no(warning_dialog):
    warning_dialog.destroy()
    dialog.withdraw()  # 隐藏对话框
    create_main_window()


def create_main_window():
    global df, time_data, light_labels, number_labels, frame_dict, light_data, index, is_running, time_label, toggle_button, frame, main_window, selected_sheet

    # 如果主窗口已经存在，先销毁它
    if main_window:
        main_window.destroy()

    # 读取特定工作表中的数据
    df = pd.read_excel(current_file_path, sheet_name=selected_sheet)

    # 获取标题行数据
    title_row = df.columns

    # 获取时间数据，从第二行开始
    time_data = df.iloc[1:, 0].dropna().astype(str).tolist()

    # 检查时间数据是否为数字
    try:
        time_data = [int(value) for value in time_data if value.isdigit()]
    except ValueError:
        messagebox.showerror("Error", "Invalid time data in the selected sheet.")
        return

    # 创建主窗口
    main_window = tk.Toplevel(root)
    main_window.title("灯光动画显示器")

    # 绑定关闭事件
    main_window.protocol("WM_DELETE_WINDOW", on_main_window_close)

    # 在左上角显示工作表名称
    sheet_label = tk.Label(main_window, text=f"Current Sheet: {selected_sheet}", font=("Helvetica", 14))
    sheet_label.pack(anchor="nw", pady=10, padx=10)

    # 创建一个时间块
    global time_label
    time_label = tk.Label(main_window, text="Time: 0 ms", font=("Helvetica", 14))
    time_label.pack(pady=10)

    # 添加一个开关块
    toggle_button = tk.Button(main_window, text="关", command=toggle_running, font=("Helvetica", 14))
    toggle_button.pack(pady=10)
    ToolTip(toggle_button, "Start/Stop the animation")

    # 添加 reset 按钮
    reset_button = tk.Button(main_window, text="Reset", command=reset, font=("Helvetica", 14))
    reset_button.pack(pady=10)
    ToolTip(reset_button, "Reset the animation data")

    # 创建一个框架来放置灯光标签
    frame = tk.Frame(main_window)
    frame.pack(pady=10)

    # 清除旧的灯标签和数据
    light_labels.clear()
    number_labels.clear()
    frame_dict.clear()
    light_data.clear()
    index = 0

    def extract_group(name):
        match = re.match(r'(.+?)LED', name)
        if match:
            return match.group(1)
        match = re.search(r'LED(.+)', name)
        if match:
            suffix = match.group(1)
            letter_match = re.search(r'[A-Za-z]+', suffix)
            if letter_match:
                return letter_match.group(0)
        return "Unknown"

    # 创建灯光标签
    for col_index, cell_value in enumerate(title_row):
        if pd.notna(cell_value) and cell_value != '':
            group = extract_group(cell_value)

            if group not in frame_dict:
                frame_dict[group] = []

            if len(frame_dict[group]) == 0 or len(frame_dict[group][-1].winfo_children()) >= 20:
                new_frame = tk.Frame(frame)
                new_frame.pack(pady=10)
                frame_dict[group].append(new_frame)

            light_frame = tk.Frame(frame_dict[group][-1])
            light_frame.pack(side=tk.LEFT, padx=5)

            number_label = tk.Label(light_frame, text="0", font=("Helvetica", 10))
            number_label.pack()
            number_labels.append(number_label)

            light_label = tk.Label(light_frame, text=cell_value, width=10, height=5, bg='#000000',
                                   font=("Helvetica", 8))
            light_label.brightness_value = 0  # 初始化亮度值
            light_label.pack()
            light_labels.append(light_label)

            # 存储对应列的数据，从第二行开始
            column_data = []
            for val in df.iloc[1:, col_index]:
                try:
                    column_data.append(float(val))
                except ValueError:
                    continue
            light_data.append(column_data)

    is_running.set()

def on_main_window_close():
    global dialog, main_window, is_running
    is_running.clear()
    if main_window:
        main_window.destroy()
    dialog.deiconify()  # 重新显示对话框
    dialog.lift()


def update_brightness():
    global index, is_running, time_data, light_labels, number_labels, time_label, light_data
    while True:
        if is_running.is_set() and time_data:
            current_time = time_data[index % len(time_data)]
            time_label.config(text=f"Time: {current_time} ms")
            for i, label in enumerate(light_labels):
                if light_data[i] and label.winfo_exists():
                    actual_brightness = light_data[i][index % len(light_data[i])]
                    if actual_brightness == 0:
                        display_brightness = 0
                    elif actual_brightness < 20:
                        display_brightness = actual_brightness * 2
                    else:
                        display_brightness = actual_brightness
                    brightness = max(0, min(display_brightness * 2.55, 255))  # 确保亮度在0-255范围内
                    light_color = f'#{int(brightness):02x}{int(brightness):02x}00'  # 黄色灯光
                    label.config(bg=light_color)

                    # 更新数字标签，确保数字标签存在
                    if number_labels[i].winfo_exists():
                        number_labels[i].config(text=f"{actual_brightness:.2f}")

            index += 1
        time.sleep(0.001)  # 每毫秒更新一次

def reset():
    global index, is_running, time_label, toggle_button, light_labels, number_labels
    # 停止程序
    is_running.clear()
    toggle_button.config(text="开")

    # 重置所有数据
    time_label.config(text="Time: 0 ms")
    for label in number_labels:
        label.config(text="0")
    for label in light_labels:
        label.config(bg='#000000')
        label.brightness_value = 0
    index = 0  # 重置索引变量


def toggle_running():
    global is_running, toggle_button
    if is_running.is_set():
        is_running.clear()
        toggle_button.config(text="开")
    else:
        is_running.set()
        toggle_button.config(text="关")

# pyinstaller --onefile --noconsole --add-data "Worship of the Sun.mp3;." z13.py

def play_audio():
    # 替换为您的音频文件路径
    pygame.mixer.music.load("D:/HASCO/forEZ/Worship of the Sun.mp3")
    pygame.mixer.music.play()


def show_easter_egg():
    easter_egg_window = tk.Toplevel(root)
    easter_egg_window.title("Easter Egg")

    easter_egg_text = """Surprise! Playing: 祀日歌-无期迷途."""
    easter_egg_label = tk.Label(easter_egg_window, text=easter_egg_text, justify=tk.LEFT, padx=10, pady=10)
    easter_egg_label.pack()

    close_button = tk.Button(easter_egg_window, text="Close", command=easter_egg_window.destroy)
    close_button.pack(pady=10)

    play_audio()


def show_about():
    about_window = tk.Toplevel(root)
    about_window.title("About")

    about_text = """This application was created to visualize light animations.

    - Author: Hua Tong with Chatgpt
    - Contact: tonghaa698238@gmail.com
    - Date: 08/07/2024
    I am not a Human anymore! JOJO!
    Za! WaRuDuo!\(≧▽≦)/
    MudaMudaMudaMudaMudaMudaMudaMudaMudaMuda
    OraOraOraOraOraOraOraOraOraOraOraOraOraOra!
    """

    # 创建一个文本框
    text_box = tk.Text(about_window, wrap="word", padx=10, pady=10)
    text_box.insert("1.0", about_text)

    # 将特定单词转换为按钮
    def make_easter_egg_button(event):
        text_box.tag_remove("hidden_button", "1.0", tk.END)
        text_box.tag_add("hidden_button", "1.0", "1.4")
        text_box.tag_config("hidden_button", foreground="black", underline=True)
        text_box.tag_bind("hidden_button", "<Button-1>", lambda e: show_easter_egg())

    text_box.tag_config("hidden_button", foreground="black")
    text_box.tag_add("hidden_button", "1.0", "1.4")  # 假设隐藏按钮为文本的第一个单词
    text_box.tag_bind("hidden_button", "<Button-1>", lambda e: show_easter_egg())

    text_box.config(state=tk.DISABLED)
    text_box.pack()

    close_button = tk.Button(about_window, text="Close", command=about_window.destroy)
    close_button.pack(pady=10)


# 创建并启动后台线程来更新亮度
thread = threading.Thread(target=update_brightness)
thread.daemon = True  # 设置为守护线程，这样在主线程结束时该线程也会结束
thread.start()

# 主窗口
root = tk.Tk()
root.withdraw()  # 隐藏主窗口

# 弹出对话框
dialog = tk.Toplevel(root)
dialog.title("选择文件和工作表")

# 绑定关闭事件，确保点击右上角关闭按钮时的行为与点击“Cancel”按钮一致
dialog.protocol("WM_DELETE_WINDOW", on_cancel)

file_path_var = tk.StringVar()

# 选择文件按钮和文本框
file_button = tk.Button(dialog, text="选择文件", command=lambda: file_path_var.set(
    filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])))
file_button.pack(pady=10)
ToolTip(file_button, "Choose an Excel file")

file_path_entry = tk.Entry(dialog, textvariable=file_path_var, width=50)
file_path_entry.pack(pady=10)
ToolTip(file_path_entry, "File path of the selected Excel file")

# 选择工作表下拉菜单
sheet_selector = ttk.Combobox(dialog, state="readonly", font=("Helvetica", 14))
sheet_selector.pack(pady=10)
ToolTip(sheet_selector, "Select a sheet from the Excel file")


# 绑定选择文件事件
def on_file_select():
    file_path = file_path_var.get()
    if file_path:
        sheet_names = pd.ExcelFile(file_path).sheet_names
        sheet_selector['values'] = sheet_names
        if sheet_names:
            sheet_selector.current(0)


file_path_var.trace("w", lambda *args: on_file_select())

# OK和Cancel按钮
button_frame = tk.Frame(dialog)
button_frame.pack(pady=10)

ok_button = tk.Button(button_frame, text="OK", command=on_ok)
ok_button.pack(side=tk.LEFT, padx=5)
ToolTip(ok_button, "Load the selected sheet")

cancel_button = tk.Button(button_frame, text="Cancel", command=on_cancel)
cancel_button.pack(side=tk.LEFT, padx=5)
ToolTip(cancel_button, "Cancel and exit the application")

# 添加 About 按钮
about_button = tk.Button(dialog, text="About", command=show_about)
about_button.pack(side=tk.RIGHT, padx=5, pady=5)
ToolTip(about_button, "Show information about this application")

# 启动主事件循环
root.mainloop()
