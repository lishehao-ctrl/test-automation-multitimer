import os
import sys
import time
from datetime import datetime, timedelta

import pyvisa as visa
import tkinter as tk
import tkinter.font as font
from tkinter import ttk, filedialog, scrolledtext

from dmm_driver import instKS_34461A


class TerminalRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.originial_stdout = sys.stdout

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)

    def flush(self):
        pass


class UI(tk.Tk):
    data_type_mode = "Mode"
    data_type_ac_dc = "直流电交流电"
    data_type_range = "Range"
    data_type_sleep_time = "间隔时间"
    data_type_time_dur = "监测时间"
    data_type_time_dur_unit = "时长单位"
    data_type_usb_lan = "USB/LAN"
    data_type_visa_address = "Visa Address"
    data_type_user_input_btn_process = "结束"

    user_input_non_positive_alert = "请输入正数！！！"
    user_input_wrong_type_alert = "请输入数字！！！"
    user_input_miss_sleep_time_alert = "请输入触发间隔时间！！！"
    user_input_miss_time_dur_alert = "请输入监测时长！！！"
    user_input_miss_visa_address = "请输入Visa地址！！！"
    user_input_miss_ip_address = "请输入ip地址！！！"

    default_mode = "VOLT"
    default_ac_dc = "AC"
    default_range = "AUTO"
    default_time_dur_unit = "秒"
    default_usb_lan = "USB"
    default_fn = "Test_File.mat"
    show_selection_text_font = ("Microsoft YaHei UI", 18)
    default_text_font = ("Microsoft YaHei UI", 10)
    update_frequency = 100

    lable_for_show_selection = "你选中了："
    lable_for_mode_input = "请选择想要测量的Mode"
    lable_for_ac_dc_input = "请选择想要测量的AC/DC"
    lable_for_range_input = "请选择想要测量的Range"
    lable_for_sleep_time_input = "请输入想要的触发间隔时间(秒)"
    lable_for_time_dur_input = "请输入想要的监测时间(秒/分/时)"
    lable_for_usb_lan = "请选择设备连接模式"
    lable_for_visa_address = "请输入设备visa地址"
    lable_for_ip_address = "请输入设备ip地址   "
    label_for_filedialog_title = "选择保存路径和文件名"

    AC = "AC"
    DC = "DC"
    VOLTAGE = "VOLT"
    CURRENT = "CURR"
    time_unit_second = "秒"
    time_unit_minute = "分钟"
    time_unit_hour = "小时"
    usb = "USB"
    lan = "LAN"
    auto_detect = "自动识别"

    def __init__(self):
        super().__init__()
        self.change_font()
        self.screen_wdith = self.winfo_screenwidth()
        self.screen_height = self.winfo_screenheight()
        self.title("输入控制界面")

        fn = self.default_fn
        default_filepath = instKS_34461A.fn_relative(self, fn=fn)
        dir_path = os.path.dirname(default_filepath)
        new_file_name = datetime.now().strftime("%Y%m%d_%H_%M_%S") + "_" + fn
        self.file_path = os.path.join(dir_path, new_file_name)

    def change_font(self):
        self.default_font = font.nametofont("TkDefaultFont")
        self.default_font.configure(family=self.default_text_font[0], size=self.default_text_font[1])
        self.text_font = font.nametofont("TkTextFont")
        self.text_font.configure(family=self.default_text_font[0], size=self.default_text_font[1])

    def get_insts(self):
        rm = visa.ResourceManager()
        insts = rm.list_resources()
        return insts

    def generat_ui(self):
        self.usb_lan = tk.StringVar(value=self.default_usb_lan)
        self.var_usb_visa_address = tk.StringVar()
        self.var_visa_address = None

        self.lb_usb_lan = tk.Label(self, text=self.lable_for_usb_lan)
        self.lb_usb_lan.pack(anchor=tk.W)

        self.frame_usb_lan = tk.Frame(self)
        self.frame_usb_lan.pack(side=tk.LEFT, pady=5)

        self.frame_usb = tk.Frame(self)
        self.frame_usb.pack(anchor=tk.W)

        self.rb_btn_usb = tk.Radiobutton(
            self.frame_usb,
            text=self.auto_detect,
            variable=self.usb_lan,
            value=self.usb,
            command=lambda: self.show_selected(self.data_type_usb_lan),
        )
        self.rb_btn_usb.pack(side=tk.LEFT)

        self.lb_usb_visa_address = tk.Label(self.frame_usb, text=self.lable_for_visa_address)
        self.lb_usb_visa_address.pack(side=tk.LEFT)

        self.cmb_usb_visa_address = ttk.Combobox(self.frame_usb, textvariable=self.var_usb_visa_address, width=40)
        self.cmb_usb_visa_address["values"] = self.get_insts()
        self.cmb_usb_visa_address.pack(side=tk.LEFT, padx=10)
        if self.get_insts():
            self.cmb_usb_visa_address.set(self.get_insts()[0])

        self.btn_refresh = tk.Button(self.frame_usb, text="刷新", command=self.refresh_insts)
        self.btn_refresh.pack(side=tk.LEFT)

        self.frame_lan = tk.Frame(self)
        self.frame_lan.pack(anchor=tk.W)

        self.rb_btn_lan = tk.Radiobutton(
            self.frame_lan,
            text=self.lan,
            variable=self.usb_lan,
            value=self.lan,
            command=lambda: self.show_selected(self.data_type_usb_lan),
        )
        self.rb_btn_lan.pack(side=tk.LEFT)

        self.lb_lan_visa_address = tk.Label(self.frame_lan, text=self.lable_for_ip_address)
        self.lb_lan_visa_address.pack(side=tk.LEFT)

        self.txt_lan_visa_address_1 = tk.Text(self.frame_lan, width=10, height=1)
        self.txt_lan_visa_address_1.pack(side=tk.LEFT, padx=(10, 0))
        self.lb_lan_visa_address_1 = tk.Label(self.frame_lan, text=".")
        self.lb_lan_visa_address_1.pack(side=tk.LEFT)
        self.txt_lan_visa_address_2 = tk.Text(self.frame_lan, width=10, height=1)
        self.txt_lan_visa_address_2.pack(side=tk.LEFT, padx=(10, 0))
        self.lb_lan_visa_address_1 = tk.Label(self.frame_lan, text=".")
        self.lb_lan_visa_address_1.pack(side=tk.LEFT)
        self.txt_lan_visa_address_3 = tk.Text(self.frame_lan, width=10, height=1)
        self.txt_lan_visa_address_3.pack(side=tk.LEFT, padx=(10, 0))
        self.lb_lan_visa_address_1 = tk.Label(self.frame_lan, text=".")
        self.lb_lan_visa_address_1.pack(side=tk.LEFT)
        self.txt_lan_visa_address_4 = tk.Text(self.frame_lan, width=10, height=1)
        self.txt_lan_visa_address_4.pack(side=tk.LEFT, padx=(10, 0))

        self.var_mode = tk.StringVar(value=self.default_mode)
        self.var_ac_dc = tk.StringVar(value=self.default_ac_dc)
        self.var_range = tk.StringVar(value=self.default_range)
        self.var_sleep_time = None
        self.var_time_dur = None
        self.var_time_dur_unit = tk.StringVar(value=self.default_time_dur_unit)

        self.frame_mode_text = tk.Frame(self)
        self.frame_mode_text.pack(anchor=tk.W, pady=5)

        self.lb_mode = tk.Label(self.frame_mode_text, text=self.lable_for_mode_input)
        self.lb_mode.pack(side=tk.LEFT)

        self.frame_mode = tk.Frame(self)
        self.frame_mode.pack(pady=5, anchor=tk.W)

        self.rd_btn_mode_1 = tk.Radiobutton(
            self.frame_mode,
            text=self.VOLTAGE,
            variable=self.var_mode,
            value=self.VOLTAGE,
            command=lambda: self.show_selected(self.data_type_mode),
        )
        self.rd_btn_mode_1.pack(side=tk.LEFT)
        self.rd_btn_mode_2 = tk.Radiobutton(
            self.frame_mode,
            text=self.CURRENT,
            variable=self.var_mode,
            value=self.CURRENT,
            command=lambda: self.show_selected(self.data_type_mode),
        )
        self.rd_btn_mode_2.pack(side=tk.LEFT)

        self.lb_ac_dc = tk.Label(self, text=self.lable_for_ac_dc_input)
        self.lb_ac_dc.pack(pady=10, anchor=tk.W)

        self.frame_ac_dc = tk.Frame(self)
        self.frame_ac_dc.pack(pady=5, anchor=tk.W)

        self.rd_btn_ac_dc_1 = tk.Radiobutton(
            self.frame_ac_dc,
            text=self.AC,
            variable=self.var_ac_dc,
            value=self.AC,
            command=lambda: self.show_selected(self.data_type_ac_dc),
        )
        self.rd_btn_ac_dc_1.pack(side=tk.LEFT)
        self.rd_btn_ac_dc_2 = tk.Radiobutton(
            self.frame_ac_dc,
            text=self.DC,
            variable=self.var_ac_dc,
            value=self.DC,
            command=lambda: self.show_selected(self.data_type_ac_dc),
        )
        self.rd_btn_ac_dc_2.pack(side=tk.LEFT)

        self.lb_range = tk.Label(self, text=self.lable_for_range_input)
        self.lb_range.pack(pady=10, anchor=tk.W)

        self.frame_range = tk.Frame(self)
        self.frame_range.pack(pady=5, anchor=tk.W)

        self.rd_btn_range_1 = tk.Radiobutton(
            self.frame_range,
            text="AUTO",
            variable=self.var_range,
            value="AUTO",
            command=lambda: self.show_selected(self.data_type_range),
        )
        self.rd_btn_range_1.pack(side=tk.LEFT)
        self.rd_btn_range_2 = tk.Radiobutton(
            self.frame_range,
            text="0.1",
            variable=self.var_range,
            value="0.1",
            command=lambda: self.show_selected(self.data_type_range),
        )
        self.rd_btn_range_2.pack(side=tk.LEFT)
        self.rd_btn_range_3 = tk.Radiobutton(
            self.frame_range,
            text="1",
            variable=self.var_range,
            value="1",
            command=lambda: self.show_selected(self.data_type_range),
        )
        self.rd_btn_range_3.pack(side=tk.LEFT)
        self.rd_btn_range_4 = tk.Radiobutton(
            self.frame_range,
            text="10",
            variable=self.var_range,
            value="10",
            command=lambda: self.show_selected(self.data_type_range),
        )
        self.rd_btn_range_4.pack(side=tk.LEFT)
        self.rd_btn_range_5 = tk.Radiobutton(
            self.frame_range,
            text="100",
            variable=self.var_range,
            value="100",
            command=lambda: self.show_selected(self.data_type_range),
        )
        self.rd_btn_range_5.pack(side=tk.LEFT)
        self.rd_btn_range_6 = tk.Radiobutton(
            self.frame_range,
            text="1000",
            variable=self.var_range,
            value="1000",
            command=lambda: self.show_selected(self.data_type_range),
        )
        self.rd_btn_range_6.pack(side=tk.LEFT)
        self.rd_btn_range_7 = tk.Radiobutton(
            self.frame_range,
            text="10",
            variable=self.var_range,
            value="10",
            command=lambda: self.show_selected(self.data_type_range),
        )

        self.lb_sleep = tk.Label(self, text=self.lable_for_sleep_time_input)
        self.lb_sleep.pack(pady=10, anchor=tk.W)

        self.frame_sleep = tk.Frame(self)
        self.frame_sleep.pack(pady=5, anchor=tk.W)

        self.txt_sleep = tk.Text(self.frame_sleep, width=10, height=1)
        self.txt_sleep.pack(side=tk.LEFT, padx=5)

        self.lb_time_dur = tk.Label(self, text=self.lable_for_time_dur_input)
        self.lb_time_dur.pack(pady=10, anchor=tk.W)

        self.frame_time_dur = tk.Frame(self)
        self.frame_time_dur.pack(pady=5, anchor=tk.W)

        self.txt_time_dur = tk.Text(self.frame_time_dur, width=10, height=1)
        self.txt_time_dur.pack(side=tk.LEFT, padx=5)
        self.rd_btn_time_dur_unit_1 = tk.Radiobutton(
            self.frame_time_dur,
            text=self.time_unit_second,
            variable=self.var_time_dur_unit,
            value=self.time_unit_second,
            command=lambda: self.show_selected(self.data_type_time_dur_unit),
        )
        self.rd_btn_time_dur_unit_1.pack(side=tk.LEFT)
        self.rd_btn_time_dur_unit_2 = tk.Radiobutton(
            self.frame_time_dur,
            text=self.time_unit_minute,
            variable=self.var_time_dur_unit,
            value=self.time_unit_minute,
            command=lambda: self.show_selected(self.data_type_time_dur_unit),
        )
        self.rd_btn_time_dur_unit_2.pack(side=tk.LEFT)
        self.rd_btn_time_dur_unit_3 = tk.Radiobutton(
            self.frame_time_dur,
            text=self.time_unit_hour,
            variable=self.var_time_dur_unit,
            value=self.time_unit_hour,
            command=lambda: self.show_selected(self.data_type_time_dur_unit),
        )
        self.rd_btn_time_dur_unit_3.pack(side=tk.LEFT)

        self.lb_show_selected = tk.Label(self, text=self.lable_for_show_selection, font=(self.show_selection_text_font))
        self.lb_show_selected.pack(anchor=tk.W, padx=5)

        self.frame_btn_control = tk.Frame(self)
        self.frame_btn_control.pack(pady=5, anchor=tk.W)

        self.btn_file_path = tk.Button(
            self.frame_btn_control,
            width=20,
            height=2,
            text="文件保存地址",
            command=lambda: self.get_filepath(),
        )
        self.btn_file_path.pack(side=tk.LEFT, padx=5)

        self.btn_begin_test = tk.Button(
            self.frame_btn_control, width=20, height=2, text="开始测试", command=self.begin_measure
        )
        self.btn_begin_test.pack(side=tk.LEFT, padx=5)

        self.btn_terminate_test = tk.Button(
            self.frame_btn_control, width=20, height=2, text="停止测试", command=self.terminate
        )

        self.btn_exit = tk.Button(self.frame_btn_control, width=20, height=2, text="退出程序", command=sys.exit)
        self.btn_exit.pack(side=tk.LEFT, padx=5)

        self.show_terminal()

    def refresh_insts(self):
        self.cmb_usb_visa_address["values"] = self.get_insts()
        if self.get_insts():
            self.cmb_usb_visa_address.set(self.get_insts()[0])

    def show_terminal(self):
        self.text_area = scrolledtext.ScrolledText(
            self, wrap=tk.WORD, bg="black", fg="white", font=self.default_text_font
        )
        self.text_area.pack(padx=10, pady=5, anchor=tk.W, fill=tk.BOTH, expand=True)

        sys.stdout = TerminalRedirector(text_widget=self.text_area)
        sys.stderr = TerminalRedirector(text_widget=self.text_area)

    def show_remained_V(self):
        self.rd_btn_range_1.config(text="AUTO", variable=self.var_range, value="AUTO")
        self.rd_btn_range_2.config(text="0.1", variable=self.var_range, value="0.1")
        self.rd_btn_range_3.config(text="1", variable=self.var_range, value="1")
        self.rd_btn_range_4.config(text="10", variable=self.var_range, value="10")
        self.rd_btn_range_5.config(text="100", variable=self.var_range, value="100")
        self.rd_btn_range_6.config(text="1000", variable=self.var_range, value="1000")
        self.rd_btn_range_7.pack_forget()

    def show_remained_I(self):
        self.rd_btn_range_1.config(text="AUTO", variable=self.var_range, value="AUTO")
        self.rd_btn_range_2.config(text="0.0001", variable=self.var_range, value="0.0001")
        self.rd_btn_range_3.config(text="0.001", variable=self.var_range, value="0.001")
        self.rd_btn_range_4.config(text="0.01", variable=self.var_range, value="0.01")
        self.rd_btn_range_5.config(text="1", variable=self.var_range, value="1")
        self.rd_btn_range_6.config(text="3", variable=self.var_range, value="3")
        self.rd_btn_range_7.pack(side=tk.LEFT, padx=0)

    def get_filepath(self, fn="Test_File.mat"):
        self.file_path = filedialog.asksaveasfilename(
            defaultextension=".mat",
            filetypes=[("MAT files", "*.mat"), ("All files", "*.*")],
            initialfile=fn,
            title=self.label_for_filedialog_title,
        )

    def show_selected(self, data_type):
        var = None
        if data_type == self.data_type_mode:
            var = self.var_mode.get()
            if var == self.VOLTAGE:
                self.show_remained_V()
            elif var == self.CURRENT:
                self.show_remained_I()
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_ac_dc:
            var = self.var_ac_dc.get()
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_range:
            var = self.var_range.get()
            if var not in ("AUTO", "auto"):
                var = float(var)
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_sleep_time or data_type == self.data_type_time_dur:
            var_sleep_time = self.txt_sleep.get("1.0", tk.END).strip()
            var_time_dur = self.txt_time_dur.get("1.0", tk.END).strip()
            try:
                num_sleep_time = float(var_sleep_time)
                num_time_dur = float(var_time_dur)
                if num_sleep_time <= 0 or num_time_dur <= 0:
                    var = self.user_input_non_positive_alert
                    self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
                else:
                    self.var_sleep_time = num_sleep_time
                    self.var_time_dur = num_time_dur
            except Exception:
                if var_sleep_time and var_time_dur:
                    var = self.user_input_wrong_type_alert
                else:
                    var = []
                    if not var_sleep_time:
                        var.append(self.user_input_miss_sleep_time_alert)
                    if not var_time_dur:
                        var.append(self.user_input_miss_time_dur_alert)
                self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_time_dur_unit:
            var = self.var_time_dur_unit.get()
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_usb_lan:
            var = self.usb_lan.get()
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_visa_address:
            var_usb = self.var_usb_visa_address.get().replace(" ", "")
            var_lan = (
                self.txt_lan_visa_address_1.get("1.0", tk.END).replace(" ", "").rstrip("\n")
                + "."
                + self.txt_lan_visa_address_2.get("1.0", tk.END).replace(" ", "").rstrip("\n")
                + "."
                + self.txt_lan_visa_address_3.get("1.0", tk.END).replace(" ", "").rstrip("\n")
                + "."
                + self.txt_lan_visa_address_4.get("1.0", tk.END).replace(" ", "").rstrip("\n")
            )

            usb_selected = self.usb_lan.get() == self.usb
            lan_selected = self.usb_lan.get() == self.lan

            self.lan_visa_address_typed = (
                self.txt_lan_visa_address_1.get("1.0", tk.END).rstrip("\n")
                and self.txt_lan_visa_address_2.get("1.0", tk.END).rstrip("\n")
                and self.txt_lan_visa_address_3.get("1.0", tk.END).rstrip("\n")
                and self.txt_lan_visa_address_4.get("1.0", tk.END).rstrip("\n")
                and lan_selected
            )

            self.usb_visa_address_typed = var_usb and usb_selected

            if self.lan_visa_address_typed or self.usb_visa_address_typed:
                self.var_visa_address = f"TCPIP0::{var_lan}::INSTR" if lan_selected else var_usb
            else:
                if usb_selected:
                    var = self.user_input_miss_visa_address
                else:
                    var = self.user_input_miss_ip_address
                self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")

    def get_data(self, data_type):
        if data_type == self.data_type_mode:
            return self.var_mode.get()
        elif data_type == self.data_type_ac_dc:
            return self.var_ac_dc.get()
        elif data_type == self.data_type_range:
            return self.var_range.get()
        elif data_type == self.data_type_sleep_time:
            return self.var_sleep_time
        elif data_type == self.data_type_time_dur:
            return self.var_time_dur
        elif data_type == self.data_type_time_dur_unit:
            return self.var_time_dur_unit.get()
        elif data_type == self.data_type_visa_address:
            return self.var_visa_address

    def cal_run_time(self, time_unit, time_dur):
        time_in_second = 0
        if time_unit == self.time_unit_second:
            time_in_second = timedelta(seconds=time_dur).total_seconds()
        elif time_unit == self.time_unit_minute:
            time_in_second = timedelta(minutes=time_dur).total_seconds()
        else:
            time_in_second = timedelta(hours=time_dur).total_seconds()
        return time_in_second

    def begin_measure(self):
        self.show_selected(self.data_type_sleep_time)
        self.show_selected(self.data_type_time_dur)
        self.show_selected(self.data_type_visa_address)

        if self.var_sleep_time and self.var_time_dur and self.var_visa_address:
            self.is_terminated = False

            self.btn_terminate_test.pack(side=tk.LEFT, padx=5)
            self.btn_exit.pack_forget()
            self.btn_begin_test.pack_forget()
            self.btn_file_path.pack_forget()

            self.saved_visa_address = self.get_data(UI.data_type_visa_address)
            self.saved_mode_input = self.get_data(UI.data_type_mode)
            self.saved_ac_dc_input = self.get_data(UI.data_type_ac_dc)
            self.saved_range_input = self.get_data(UI.data_type_range)
            self.saved_sleep_time = self.get_data(UI.data_type_sleep_time)
            self.saved_time_dur = self.get_data(UI.data_type_time_dur)
            self.saved_time_dur_unit = self.get_data(UI.data_type_time_dur_unit)

            mt = instKS_34461A(visa_address=self.saved_visa_address)
            mt.inst_open()

            mt.set_mode(self.saved_mode_input, self.saved_ac_dc_input)
            mt.set_range(self.saved_range_input)

            print("主程序开始处理")

            total_runtime = self.cal_run_time(self.saved_time_dur_unit, self.saved_time_dur)
            self.time_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            start_time = time.time()

            count = 0

            self.time_stamps = []
            self.power_data = []

            self.time_stamps_path = None
            self.power_data_path = None

            is_delete_first_measure = False

            while True:
                time_since_start = time.time() - start_time
                if time_since_start >= total_runtime or self.is_terminated:
                    self.save_mat_file()
                    print(f"数据采集结束 程序已运行{time_since_start:.2f}{self.time_unit_second}")
                    break

                if count - 100 >= 0 and count % 100 == 0:
                    self.save_mat_file()

                self.time_stamps.append(time_since_start)
                self.time_measure_start = time.time()
                power = mt.measure()
                self.power_data.append(power)
                count += 1
                current_time = datetime.now().strftime("%m.%d %H:%M:%S")
                self.update()

                if not is_delete_first_measure:
                    is_delete_first_measure = True
                    start_time += time_since_start
                    time_since_start = 0

                print(f"[{current_time}] 执行任务{count}次...{power}")

                self.update_during_sleep(start_time, total_runtime, count, self.saved_sleep_time)

            self.btn_file_path.pack(side=tk.LEFT, padx=5)
            self.btn_begin_test.pack(side=tk.LEFT, padx=5)
            self.btn_exit.pack(side=tk.LEFT, padx=5)
            self.btn_terminate_test.pack_forget()

            mt.close()

    def terminate(self):
        self.is_terminated = True
        self.btn_file_path.pack(side=tk.LEFT, padx=5)
        self.btn_begin_test.pack(side=tk.LEFT, padx=5)
        self.btn_exit.pack(side=tk.LEFT, padx=5)
        self.btn_terminate_test.pack_forget()

    def update_during_sleep(self, start_time, total_runtime, count, sleep_time):
        interval = 1 / self.update_frequency
        sleep_count = int(sleep_time * self.update_frequency)
        time_end = self.time_measure_start + sleep_time
        for _ in range(sleep_count):
            self.update()
            time_since_start = time.time() - start_time
            if time_since_start >= total_runtime or self.is_terminated:
                break

            if time_since_start > (count * sleep_time):
                break

            current_time = time.time()
            remaining_time = time_end - current_time
            if remaining_time < interval:
                if remaining_time > 0:
                    time.sleep(interval)
                break

            time.sleep(interval)

    def save_mat_file(self):
        mat_var_time_stamps = "time_stamps"
        mat_var_power = "power"
        mat_var_config = "configuration"

        config = [
            f"{self.data_type_visa_address}: " + self.saved_visa_address,
            f"{self.data_type_mode}: " + self.saved_mode_input,
            f"{self.data_type_ac_dc}: " + self.saved_ac_dc_input,
            f"{self.data_type_range}: " + self.saved_range_input,
            f"{self.data_type_sleep_time}: " + str(self.saved_sleep_time) + self.default_time_dur_unit,
            f"{self.data_type_time_dur}: " + str(self.saved_time_dur) + self.saved_time_dur_unit,
            "开始时间: " + self.time_start,
            "保存时间: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ]

        data_to_save = {
            mat_var_time_stamps: list(self.time_stamps),
            mat_var_power: list(self.power_data),
            mat_var_config: list(config),
        }

        # Delegated to driver layer (may use scipy/numpy internally if available)
        if hasattr(instKS_34461A, "save_matfile"):
            instKS_34461A.save_matfile(self, fn=self.file_path, mm=data_to_save)
