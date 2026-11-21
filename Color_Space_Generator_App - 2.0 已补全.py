import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import json
import sqlite3
import math
import colorsys
from colormath.color_objects import sRGBColor, LabColor, CMYKColor, HSVColor
from colormath.color_conversions import convert_color
from datetime import datetime

class ColorSpaceGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("多色彩空间数值生成器")
        
        # 设置窗口大小和位置
        window_width = 550
        window_height = 650
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        # 创建界面元素
        self.create_widgets()
        
        # 用于跟踪生成进度
        self.generating = False
        self.current_progress = 0
        
        # 默认色彩空间
        self.color_space = "RGB"
        # 默认输出格式
        self.output_format = "CSV"
    
    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="多色彩空间数值生成器", font=("Microsoft YaHei", 16))
        title_label.pack(pady=10)
        
        # 色彩空间选择
        color_space_frame = tk.LabelFrame(self.root, text="选择色彩空间", font=("Microsoft YaHei", 10))
        color_space_frame.pack(pady=5, padx=10, fill=tk.X)
        
        self.color_space_var = tk.StringVar(value="RGB")
        color_spaces = ["RGB", "HSL/HSV", "CMYK", "YUV", "Lab", "HEX"]
        
        for i, space in enumerate(color_spaces):
            rb = tk.Radiobutton(
                color_space_frame,
                text=space,
                variable=self.color_space_var,
                value=space,
                font=("Microsoft YaHei", 9),
                command=self.update_input_fields
            )
            rb.grid(row=0, column=i, padx=5, pady=5)
        
        # 输出格式选择
        output_format_frame = tk.LabelFrame(self.root, text="输出格式", font=("Microsoft YaHei", 10))
        output_format_frame.pack(pady=5, padx=10, fill=tk.X)
        
        self.output_format_var = tk.StringVar(value="CSV")
        output_formats = ["CSV", "JSON", "SQLite"]
        
        for i, fmt in enumerate(output_formats):
            rb = tk.Radiobutton(
                output_format_frame,
                text=fmt,
                variable=self.output_format_var,
                value=fmt,
                font=("Microsoft YaHei", 9)
            )
            rb.grid(row=0, column=i, padx=5, pady=5)
        
        # 参数输入区域
        self.params_frame = tk.LabelFrame(self.root, text="参数设置", font=("Microsoft YaHei", 10))
        self.params_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        
        # 创建各色彩空间的输入字段
        self.create_rgb_fields()
        self.create_hsl_fields()
        self.create_cmyk_fields()
        self.create_yuv_fields()
        self.create_lab_fields()
        self.create_hex_fields()
        
        # 默认显示RGB字段
        self.show_rgb_fields()
        
        # 步长设置
        step_frame = tk.Frame(self.root)
        step_frame.pack(pady=5)
        tk.Label(step_frame, text="步长:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.step = tk.Entry(step_frame, width=5)
        self.step.pack(side=tk.LEFT, padx=5)
        self.step.insert(0, "1")
        
        # 进度条
        self.progress_frame = tk.Frame(self.root)
        self.progress_frame.pack(pady=10, fill=tk.X, padx=20)
        
        self.progress_label = tk.Label(
            self.progress_frame, 
            text="准备生成...", 
            font=("Microsoft YaHei", 9)
        )
        self.progress_label.pack()
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            orient=tk.HORIZONTAL,
            length=450,
            mode='determinate'
        )
        self.progress_bar.pack()
        
        # 生成按钮
        generate_btn = tk.Button(
            self.root, 
            text="生成文件", 
            font=("Microsoft YaHei", 12), 
            command=self.start_generation
        )
        generate_btn.pack(pady=10)
        
        # 停止按钮
        self.stop_btn = tk.Button(
            self.root,
            text="停止生成",
            font=("Microsoft YaHei", 10),
            command=self.stop_generation,
            state=tk.DISABLED
        )
        self.stop_btn.pack()
    
    def create_rgb_fields(self):
        self.rgb_frame = tk.Frame(self.params_frame)
        
        # R 通道设置
        r_frame = tk.Frame(self.rgb_frame)
        r_frame.pack(pady=5, fill=tk.X)
        tk.Label(r_frame, text="R 范围:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.r_min = tk.Entry(r_frame, width=5)
        self.r_min.pack(side=tk.LEFT, padx=5)
        tk.Label(r_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.r_max = tk.Entry(r_frame, width=5)
        self.r_max.pack(side=tk.LEFT, padx=5)
        self.r_min.insert(0, "0")
        self.r_max.insert(0, "255")
        
        # G 通道设置
        g_frame = tk.Frame(self.rgb_frame)
        g_frame.pack(pady=5, fill=tk.X)
        tk.Label(g_frame, text="G 范围:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.g_min = tk.Entry(g_frame, width=5)
        self.g_min.pack(side=tk.LEFT, padx=5)
        tk.Label(g_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.g_max = tk.Entry(g_frame, width=5)
        self.g_max.pack(side=tk.LEFT, padx=5)
        self.g_min.insert(0, "0")
        self.g_max.insert(0, "255")
        
        # B 通道设置
        b_frame = tk.Frame(self.rgb_frame)
        b_frame.pack(pady=5, fill=tk.X)
        tk.Label(b_frame, text="B 范围:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.b_min = tk.Entry(b_frame, width=5)
        self.b_min.pack(side=tk.LEFT, padx=5)
        tk.Label(b_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.b_max = tk.Entry(b_frame, width=5)
        self.b_max.pack(side=tk.LEFT, padx=5)
        self.b_min.insert(0, "0")
        self.b_max.insert(0, "255")
    
    def create_hsl_fields(self):
        self.hsl_frame = tk.Frame(self.params_frame)
        
        # H 通道设置
        h_frame = tk.Frame(self.hsl_frame)
        h_frame.pack(pady=5, fill=tk.X)
        tk.Label(h_frame, text="H 范围 (0-360):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.h_min = tk.Entry(h_frame, width=5)
        self.h_min.pack(side=tk.LEFT, padx=5)
        tk.Label(h_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.h_max = tk.Entry(h_frame, width=5)
        self.h_max.pack(side=tk.LEFT, padx=5)
        self.h_min.insert(0, "0")
        self.h_max.insert(0, "360")
        
        # S 通道设置
        s_frame = tk.Frame(self.hsl_frame)
        s_frame.pack(pady=5, fill=tk.X)
        tk.Label(s_frame, text="S 范围 (0-100):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.s_min = tk.Entry(s_frame, width=5)
        self.s_min.pack(side=tk.LEFT, padx=5)
        tk.Label(s_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.s_max = tk.Entry(s_frame, width=5)
        self.s_max.pack(side=tk.LEFT, padx=5)
        self.s_min.insert(0, "0")
        self.s_max.insert(0, "100")
        
        # L/V 通道设置
        l_frame = tk.Frame(self.hsl_frame)
        l_frame.pack(pady=5, fill=tk.X)
        tk.Label(l_frame, text="L/V 范围 (0-100):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.l_min = tk.Entry(l_frame, width=5)
        self.l_min.pack(side=tk.LEFT, padx=5)
        tk.Label(l_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.l_max = tk.Entry(l_frame, width=5)
        self.l_max.pack(side=tk.LEFT, padx=5)
        self.l_min.insert(0, "0")
        self.l_max.insert(0, "100")
    
    def create_cmyk_fields(self):
        self.cmyk_frame = tk.Frame(self.params_frame)
        
        # C 通道设置
        c_frame = tk.Frame(self.cmyk_frame)
        c_frame.pack(pady=5, fill=tk.X)
        tk.Label(c_frame, text="C 范围 (0-100):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.c_min = tk.Entry(c_frame, width=5)
        self.c_min.pack(side=tk.LEFT, padx=5)
        tk.Label(c_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.c_max = tk.Entry(c_frame, width=5)
        self.c_max.pack(side=tk.LEFT, padx=5)
        self.c_min.insert(0, "0")
        self.c_max.insert(0, "100")
        
        # M 通道设置
        m_frame = tk.Frame(self.cmyk_frame)
        m_frame.pack(pady=5, fill=tk.X)
        tk.Label(m_frame, text="M 范围 (0-100):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.m_min = tk.Entry(m_frame, width=5)
        self.m_min.pack(side=tk.LEFT, padx=5)
        tk.Label(m_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.m_max = tk.Entry(m_frame, width=5)
        self.m_max.pack(side=tk.LEFT, padx=5)
        self.m_min.insert(0, "0")
        self.m_max.insert(0, "100")
        
        # Y 通道设置
        y_frame = tk.Frame(self.cmyk_frame)
        y_frame.pack(pady=5, fill=tk.X)
        tk.Label(y_frame, text="Y 范围 (0-100):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.y_min = tk.Entry(y_frame, width=5)
        self.y_min.pack(side=tk.LEFT, padx=5)
        tk.Label(y_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.y_max = tk.Entry(y_frame, width=5)
        self.y_max.pack(side=tk.LEFT, padx=5)
        self.y_min.insert(0, "0")
        self.y_max.insert(0, "100")
        
        # K 通道设置
        k_frame = tk.Frame(self.cmyk_frame)
        k_frame.pack(pady=5, fill=tk.X)
        tk.Label(k_frame, text="K 范围 (0-100):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.k_min = tk.Entry(k_frame, width=5)
        self.k_min.pack(side=tk.LEFT, padx=5)
        tk.Label(k_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.k_max = tk.Entry(k_frame, width=5)
        self.k_max.pack(side=tk.LEFT, padx=5)
        self.k_min.insert(0, "0")
        self.k_max.insert(0, "100")
    
    def create_yuv_fields(self):
        self.yuv_frame = tk.Frame(self.params_frame)
        
        # Y 通道设置
        y_frame = tk.Frame(self.yuv_frame)
        y_frame.pack(pady=5, fill=tk.X)
        tk.Label(y_frame, text="Y 范围 (0-1):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.y_min_yuv = tk.Entry(y_frame, width=5)
        self.y_min_yuv.pack(side=tk.LEFT, padx=5)
        tk.Label(y_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.y_max_yuv = tk.Entry(y_frame, width=5)
        self.y_max_yuv.pack(side=tk.LEFT, padx=5)
        self.y_min_yuv.insert(0, "0")
        self.y_max_yuv.insert(0, "1")
        
        # U 通道设置
        u_frame = tk.Frame(self.yuv_frame)
        u_frame.pack(pady=5, fill=tk.X)
        tk.Label(u_frame, text="U 范围 (-0.5-0.5):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.u_min = tk.Entry(u_frame, width=5)
        self.u_min.pack(side=tk.LEFT, padx=5)
        tk.Label(u_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.u_max = tk.Entry(u_frame, width=5)
        self.u_max.pack(side=tk.LEFT, padx=5)
        self.u_min.insert(0, "-0.5")
        self.u_max.insert(0, "0.5")
        
        # V 通道设置
        v_frame = tk.Frame(self.yuv_frame)
        v_frame.pack(pady=5, fill=tk.X)
        tk.Label(v_frame, text="V 范围 (-0.5-0.5):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.v_min = tk.Entry(v_frame, width=5)
        self.v_min.pack(side=tk.LEFT, padx=5)
        tk.Label(v_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.v_max = tk.Entry(v_frame, width=5)
        self.v_max.pack(side=tk.LEFT, padx=5)
        self.v_min.insert(0, "-0.5")
        self.v_max.insert(0, "0.5")
    
    def create_lab_fields(self):
        self.lab_frame = tk.Frame(self.params_frame)
        
        # L 通道设置
        l_frame = tk.Frame(self.lab_frame)
        l_frame.pack(pady=5, fill=tk.X)
        tk.Label(l_frame, text="L 范围 (0-100):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.l_min_lab = tk.Entry(l_frame, width=5)
        self.l_min_lab.pack(side=tk.LEFT, padx=5)
        tk.Label(l_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.l_max_lab = tk.Entry(l_frame, width=5)
        self.l_max_lab.pack(side=tk.LEFT, padx=5)
        self.l_min_lab.insert(0, "0")
        self.l_max_lab.insert(0, "100")
        
        # a 通道设置
        a_frame = tk.Frame(self.lab_frame)
        a_frame.pack(pady=5, fill=tk.X)
        tk.Label(a_frame, text="a 范围 (-128-127):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.a_min = tk.Entry(a_frame, width=5)
        self.a_min.pack(side=tk.LEFT, padx=5)
        tk.Label(a_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.a_max = tk.Entry(a_frame, width=5)
        self.a_max.pack(side=tk.LEFT, padx=5)
        self.a_min.insert(0, "-128")
        self.a_max.insert(0, "127")
        
        # b 通道设置
        b_frame = tk.Frame(self.lab_frame)
        b_frame.pack(pady=5, fill=tk.X)
        tk.Label(b_frame, text="b 范围 (-128-127):", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.b_min_lab = tk.Entry(b_frame, width=5)
        self.b_min_lab.pack(side=tk.LEFT, padx=5)
        tk.Label(b_frame, text="到", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.b_max_lab = tk.Entry(b_frame, width=5)
        self.b_max_lab.pack(side=tk.LEFT, padx=5)
        self.b_min_lab.insert(0, "-128")
        self.b_max_lab.insert(0, "127")
    
    def create_hex_fields(self):
        self.hex_frame = tk.Frame(self.params_frame)
        
        # HEX 起始值
        hex_start_frame = tk.Frame(self.hex_frame)
        hex_start_frame.pack(pady=5, fill=tk.X)
        tk.Label(hex_start_frame, text="起始 HEX:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.hex_start = tk.Entry(hex_start_frame, width=10)
        self.hex_start.pack(side=tk.LEFT, padx=5)
        self.hex_start.insert(0, "000000")
        
        # HEX 结束值
        hex_end_frame = tk.Frame(self.hex_frame)
        hex_end_frame.pack(pady=5, fill=tk.X)
        tk.Label(hex_end_frame, text="结束 HEX:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.hex_end = tk.Entry(hex_end_frame, width=10)
        self.hex_end.pack(side=tk.LEFT, padx=5)
        self.hex_end.insert(0, "FFFFFF")
    
    def show_rgb_fields(self):
        self.hide_all_fields()
        self.rgb_frame.pack(fill=tk.BOTH, expand=True)
    
    def show_hsl_fields(self):
        self.hide_all_fields()
        self.hsl_frame.pack(fill=tk.BOTH, expand=True)
    
    def show_cmyk_fields(self):
        self.hide_all_fields()
        self.cmyk_frame.pack(fill=tk.BOTH, expand=True)
    
    def show_yuv_fields(self):
        self.hide_all_fields()
        self.yuv_frame.pack(fill=tk.BOTH, expand=True)
    
    def show_lab_fields(self):
        self.hide_all_fields()
        self.lab_frame.pack(fill=tk.BOTH, expand=True)
    
    def show_hex_fields(self):
        self.hide_all_fields()
        self.hex_frame.pack(fill=tk.BOTH, expand=True)
    
    def hide_all_fields(self):
        for frame in [self.rgb_frame, self.hsl_frame, self.cmyk_frame, 
                     self.yuv_frame, self.lab_frame, self.hex_frame]:
            frame.pack_forget()
    
    def update_input_fields(self):
        color_space = self.color_space_var.get()
        
        if color_space == "RGB":
            self.show_rgb_fields()
        elif color_space == "HSL/HSV":
            self.show_hsl_fields()
        elif color_space == "CMYK":
            self.show_cmyk_fields()
        elif color_space == "YUV":
            self.show_yuv_fields()
        elif color_space == "Lab":
            self.show_lab_fields()
        elif color_space == "HEX":
            self.show_hex_fields()
    
    def validate_inputs(self):
        color_space = self.color_space_var.get()
        
        try:
            if color_space == "RGB":
                r_min = int(self.r_min.get())
                r_max = int(self.r_max.get())
                g_min = int(self.g_min.get())
                g_max = int(self.g_max.get())
                b_min = int(self.b_min.get())
                b_max = int(self.b_max.get())
                
                if (r_min < 0 or r_max < 0 or g_min < 0 or g_max < 0 or b_min < 0 or b_max < 0 or
                    r_min > 255 or r_max > 255 or g_min > 255 or g_max > 255 or b_min > 255 or b_max > 255):
                    raise ValueError("RGB 值必须在 0-255 之间")
                
                return True, ("RGB", r_min, r_max, g_min, g_max, b_min, b_max)
            
            elif color_space == "HSL/HSV":
                h_min = int(self.h_min.get())
                h_max = int(self.h_max.get())
                s_min = int(self.s_min.get())
                s_max = int(self.s_max.get())
                l_min = int(self.l_min.get())
                l_max = int(self.l_max.get())
                
                if (h_min < 0 or h_max < 0 or s_min < 0 or s_max < 0 or l_min < 0 or l_max < 0 or
                    h_min > 360 or h_max > 360 or s_min > 100 or s_max > 100 or l_min > 100 or l_max > 100):
                    raise ValueError("H 范围 0-360, S/L 范围 0-100")
                
                return True, ("HSL", h_min, h_max, s_min, s_max, l_min, l_max)
            
            elif color_space == "CMYK":
                c_min = int(self.c_min.get())
                c_max = int(self.c_max.get())
                m_min = int(self.m_min.get())
                m_max = int(self.m_max.get())
                y_min = int(self.y_min.get())
                y_max = int(self.y_max.get())
                k_min = int(self.k_min.get())
                k_max = int(self.k_max.get())
                
                if (c_min < 0 or c_max < 0 or m_min < 0 or m_max < 0 or 
                    y_min < 0 or y_max < 0 or k_min < 0 or k_max < 0 or
                    c_min > 100 or c_max > 100 or m_min > 100 or m_max > 100 or 
                    y_min > 100 or y_max > 100 or k_min > 100 or k_max > 100):
                    raise ValueError("CMYK 值必须在 0-100 之间")
                
                return True, ("CMYK", c_min, c_max, m_min, m_max, y_min, y_max, k_min, k_max)
            
            elif color_space == "YUV":
                y_min = float(self.y_min_yuv.get())
                y_max = float(self.y_max_yuv.get())
                u_min = float(self.u_min.get())
                u_max = float(self.u_max.get())
                v_min = float(self.v_min.get())
                v_max = float(self.v_max.get())
                
                if (y_min < 0 or y_max < 0 or y_min > 1 or y_max > 1 or
                    u_min < -0.5 or u_max < -0.5 or u_min > 0.5 or u_max > 0.5 or
                    v_min < -0.5 or v_max < -0.5 or v_min > 0.5 or v_max > 0.5):
                    raise ValueError("Y 范围 0-1, U/V 范围 -0.5-0.5")
                
                return True, ("YUV", y_min, y_max, u_min, u_max, v_min, v_max)
            
            elif color_space == "Lab":
                l_min = int(self.l_min_lab.get())
                l_max = int(self.l_max_lab.get())
                a_min = int(self.a_min.get())
                a_max = int(self.a_max.get())
                b_min = int(self.b_min_lab.get())
                b_max = int(self.b_max_lab.get())
                
                if (l_min < 0 or l_max < 0 or l_min > 100 or l_max > 100 or
                    a_min < -128 or a_max < -128 or a_min > 127 or a_max > 127 or
                    b_min < -128 or b_max < -128 or b_min > 127 or b_max > 127):
                    raise ValueError("L 范围 0-100, a/b 范围 -128-127")
                
                return True, ("Lab", l_min, l_max, a_min, a_max, b_min, b_max)
            
            elif color_space == "HEX":
                hex_start = self.hex_start.get().strip().upper()
                hex_end = self.hex_end.get().strip().upper()
                
                if len(hex_start) != 6 or len(hex_end) != 6:
                    raise ValueError("HEX 值必须是6位字符")
                
                try:
                    int(hex_start, 16)
                    int(hex_end, 16)
                except ValueError:
                    raise ValueError("HEX 值必须是有效的十六进制数")
                
                return True, ("HEX", hex_start, hex_end)
            
            step = int(self.step.get())
            if step <= 0:
                raise ValueError("步长必须大于 0")
                
            return True, None
        except ValueError as e:
            messagebox.showerror("输入错误", str(e))
            return False, None
    
    def start_generation(self):
        valid, values = self.validate_inputs()
        if not valid:
            return
        
        # 根据输出格式选择保存对话框
        output_format = self.output_format_var.get()
        if output_format == "CSV":
            filetypes = [("CSV 文件", "*.csv"), ("所有文件", "*.*")]
            defaultext = ".csv"
        elif output_format == "JSON":
            filetypes = [("JSON 文件", "*.json"), ("所有文件", "*.*")]
            defaultext = ".json"
        else:  # SQLite
            filetypes = [("SQLite 数据库", "*.db"), ("所有文件", "*.*")]
            defaultext = ".db"
        
        # 询问保存位置
        file_path = filedialog.asksaveasfilename(
            defaultextension=defaultext,
            filetypes=filetypes,
            title=f"保存 {output_format} 文件"
        )
        
        if not file_path:
            return  # 用户取消了保存
        
        # 准备生成
        self.generating = True
        self.current_progress = 0
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0
        self.progress_label.config(text="正在生成... 0%")
        self.root.update()
        
        # 在后台线程中生成文件
        self.root.after(100, lambda: self.generate_file(file_path, output_format, *values))
    
    def stop_generation(self):
        self.generating = False
        self.progress_label.config(text="生成已停止")
        self.stop_btn.config(state=tk.DISABLED)
    
    def update_progress(self, progress):
        self.progress_bar["value"] = progress
        self.progress_label.config(text=f"正在生成... {progress:.1f}%")
        self.root.update()
    
    def generate_file(self, file_path, output_format, color_space, *args):
        try:
            if output_format == "CSV":
                self.generate_csv(file_path, color_space, *args)
            elif output_format == "JSON":
                self.generate_json(file_path, color_space, *args)
            elif output_format == "SQLite":
                self.generate_sqlite(file_path, color_space, *args)
            
            if self.generating:
                messagebox.showinfo("成功", f"文件已成功生成并保存到:\n{file_path}")
                self.progress_label.config(text="生成完成!")
                self.stop_btn.config(state=tk.DISABLED)
                self.generating = False
        except Exception as e:
            messagebox.showerror("错误", f"生成文件时出错:\n{str(e)}")
            self.progress_label.config(text="生成出错!")
            self.stop_btn.config(state=tk.DISABLED)
            self.generating = False
    
    def generate_csv(self, file_path, color_space, *args):
        with open(file_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            
            if color_space == "RGB":
                writer.writerow(["R", "G", "B", "HEX"])
                r_min, r_max, g_min, g_max, b_min, b_max = args
                step = int(self.step.get())
                
                total_combinations = ((r_max - r_min + 1) // step) * ((g_max - g_min + 1) // step) * ((b_max - b_min + 1) // step)
                current_count = 0
                
                for r in range(r_min, r_max + 1, step):
                    for g in range(g_min, g_max + 1, step):
                        for b in range(b_min, b_max + 1, step):
                            if not self.generating:
                                return
                            
                            writer.writerow([r, g, b, f"#{r:02X}{g:02X}{b:02X}"])
                            current_count += 1
                            
                            if current_count % 1000 == 0 or current_count == total_combinations:
                                progress = (current_count / total_combinations) * 100
                                self.update_progress(progress)
            
            elif color_space == "HSL":
                writer.writerow(["H", "S", "L", "R", "G", "B", "HEX"])
                h_min, h_max, s_min, s_max, l_min, l_max = args
                step = int(self.step.get())
                
                total_combinations = ((h_max - h_min + 1) // step) * ((s_max - s_min + 1) // step) * ((l_max - l_min + 1) // step)
                current_count = 0
                
                for h in range(h_min, h_max + 1, step):
                    for s in range(s_min, s_max + 1, step):
                        for l_val in range(l_min, l_max + 1, step):
                            if not self.generating:
                                return
                            
                            rgb = colorsys.hls_to_rgb(h/360, l_val/100, s/100)
                            r, g, b = [round(x*255) for x in rgb]
                            writer.writerow([h, s, l_val, r, g, b, f"#{r:02X}{g:02X}{b:02X}"])
                            current_count += 1
                            
                            if current_count % 1000 == 0 or current_count == total_combinations:
                                progress = (current_count / total_combinations) * 100
                                self.update_progress(progress)
            
            elif color_space == "CMYK":
                writer.writerow(["C", "M", "Y", "K", "R", "G", "B", "HEX"])
                c_min, c_max, m_min, m_max, y_min, y_max, k_min, k_max = args
                step = int(self.step.get())
                
                total_combinations = ((c_max - c_min + 1) // step) * ((m_max - m_min + 1) // step) * ((y_max - y_min + 1) // step) * ((k_max - k_min + 1) // step)
                current_count = 0
                
                for c in range(c_min, c_max + 1, step):
                    for m in range(m_min, m_max + 1, step):
                        for y in range(y_min, y_max + 1, step):
                            for k in range(k_min, k_max + 1, step):
                                if not self.generating:
                                    return
                                
                                cmyk = CMYKColor(c/100, m/100, y/100, k/100)
                                rgb = convert_color(cmyk, sRGBColor)
                                r, g, b = [round(x*255) for x in rgb.get_value_tuple()]
                                writer.writerow([c, m, y, k, r, g, b, f"#{r:02X}{g:02X}{b:02X}"])
                                current_count += 1
                                
                                if current_count % 1000 == 0 or current_count == total_combinations:
                                    progress = (current_count / total_combinations) * 100
                                    self.update_progress(progress)
            
            elif color_space == "YUV":
                writer.writerow(["Y", "U", "V", "R", "G", "B", "HEX"])
                y_min, y_max, u_min, u_max, v_min, v_max = args
                step = 0.01  # YUV使用固定步长
                
                total_combinations = int((y_max - y_min) / step) * int((u_max - u_min) / step) * int((v_max - v_min) / step)
                current_count = 0
                
                y_val = y_min
                while y_val <= y_max:
                    u_val = u_min
                    while u_val <= u_max:
                        v_val = v_min
                        while v_val <= v_max:
                            if not self.generating:
                                return
                            
                            # 转换为RGB
                            r = y_val + 1.4075 * (v_val - 0.5)
                            g = y_val - 0.3455 * (u_val - 0.5) - 0.7169 * (v_val - 0.5)
                            b = y_val + 1.7790 * (u_val - 0.5)
                            
                            r = max(0, min(1, r))
                            g = max(0, min(1, g))
                            b = max(0, min(1, b))
                            
                            r, g, b = [round(x*255) for x in (r, g, b)]
                            writer.writerow([round(y_val, 2), round(u_val, 2), round(v_val, 2), r, g, b, f"#{r:02X}{g:02X}{b:02X}"])
                            current_count += 1
                            
                            if current_count % 1000 == 0 or current_count == total_combinations:
                                progress = (current_count / total_combinations) * 100
                                self.update_progress(progress)
                            
                            v_val += step
                        u_val += step
                    y_val += step
            
            elif color_space == "Lab":
                writer.writerow(["L", "a", "b", "R", "G", "B", "HEX"])
                l_min, l_max, a_min, a_max, b_min, b_max = args
                step = int(self.step.get())
                
                total_combinations = ((l_max - l_min + 1) // step) * ((a_max - a_min + 1) // step) * ((b_max - b_min + 1) // step)
                current_count = 0
                
                for l_val in range(l_min, l_max + 1, step):
                    for a_val in range(a_min, a_max + 1, step):
                        for b_val in range(b_min, b_max + 1, step):
                            if not self.generating:
                                return
                            
                            lab = LabColor(l_val, a_val, b_val)
                            rgb = convert_color(lab, sRGBColor)
                            r, g, b = [round(x*255) for x in rgb.get_value_tuple()]
                            writer.writerow([l_val, a_val, b_val, r, g, b, f"#{r:02X}{g:02X}{b:02X}"])
                            current_count += 1
                            
                            if current_count % 1000 == 0 or current_count == total_combinations:
                                progress = (current_count / total_combinations) * 100
                                self.update_progress(progress)
            
            elif color_space == "HEX":
                writer.writerow(["HEX", "R", "G", "B"])
                hex_start, hex_end = args
                step = int(self.step.get())
                
                start_val = int(hex_start, 16)
                end_val = int(hex_end, 16)
                
                if start_val > end_val:
                    start_val, end_val = end_val, start_val
                
                total_combinations = (end_val - start_val + 1) // step
                current_count = 0
                
                for val in range(start_val, end_val + 1, step):
                    if not self.generating:
                        return
                    
                    hex_str = f"{val:06X}"
                    r = int(hex_str[0:2], 16)
                    g = int(hex_str[2:4], 16)
                    b = int(hex_str[4:6], 16)
                    writer.writerow([hex_str, r, g, b])
                    current_count += 1
                    
                    if current_count % 1000 == 0 or current_count == total_combinations:
                        progress = ((val - start_val) / (end_val - start_val)) * 100
                        self.update_progress(progress)
    
    def generate_json(self, file_path, color_space, *args):
        color_data = []
        
        if color_space == "RGB":
            r_min, r_max, g_min, g_max, b_min, b_max = args
            step = int(self.step.get())
            
            total_combinations = ((r_max - r_min + 1) // step) * ((g_max - g_min + 1) // step) * ((b_max - b_min + 1) // step)
            current_count = 0
            
            for r in range(r_min, r_max + 1, step):
                for g in range(g_min, g_max + 1, step):
                    for b in range(b_min, b_max + 1, step):
                        if not self.generating:
                            return
                        
                        color_data.append({
                            "color_space": "RGB",
                            "R": r,
                            "G": g,
                            "B": b,
                            "hex": f"#{r:02X}{g:02X}{b:02X}"
                        })
                        
                        current_count += 1
                        if current_count % 1000 == 0 or current_count == total_combinations:
                            progress = (current_count / total_combinations) * 100
                            self.update_progress(progress)
        
        elif color_space == "HSL":
            h_min, h_max, s_min, s_max, l_min, l_max = args
            step = int(self.step.get())
            
            total_combinations = ((h_max - h_min + 1) // step) * ((s_max - s_min + 1) // step) * ((l_max - l_min + 1) // step)
            current_count = 0
            
            for h in range(h_min, h_max + 1, step):
                for s in range(s_min, s_max + 1, step):
                    for l_val in range(l_min, l_max + 1, step):
                        if not self.generating:
                            return
                        
                        rgb = colorsys.hls_to_rgb(h/360, l_val/100, s/100)
                        r, g, b = [round(x*255) for x in rgb]
                        
                        color_data.append({
                            "color_space": "HSL",
                            "H": h,
                            "S": s,
                            "L": l_val,
                            "R": r,
                            "G": g,
                            "B": b,
                            "hex": f"#{r:02X}{g:02X}{b:02X}"
                        })
                        
                        current_count += 1
                        if current_count % 1000 == 0 or current_count == total_combinations:
                            progress = (current_count / total_combinations) * 100
                            self.update_progress(progress)
        
        elif color_space == "CMYK":
            c_min, c_max, m_min, m_max, y_min, y_max, k_min, k_max = args
            step = int(self.step.get())
            
            total_combinations = ((c_max - c_min + 1) // step) * ((m_max - m_min + 1) // step) * ((y_max - y_min + 1) // step) * ((k_max - k_min + 1) // step)
            current_count = 0
            
            for c in range(c_min, c_max + 1, step):
                for m in range(m_min, m_max + 1, step):
                    for y in range(y_min, y_max + 1, step):
                        for k in range(k_min, k_max + 1, step):
                            if not self.generating:
                                return
                            
                            cmyk = CMYKColor(c/100, m/100, y/100, k/100)
                            rgb = convert_color(cmyk, sRGBColor)
                            r, g, b = [round(x*255) for x in rgb.get_value_tuple()]
                            
                            color_data.append({
                                "color_space": "CMYK",
                                "C": c,
                                "M": m,
                                "Y": y,
                                "K": k,
                                "R": r,
                                "G": g,
                                "B": b,
                                "hex": f"#{r:02X}{g:02X}{b:02X}"
                            })
                            
                            current_count += 1
                            if current_count % 1000 == 0 or current_count == total_combinations:
                                progress = (current_count / total_combinations) * 100
                                self.update_progress(progress)
        
        elif color_space == "YUV":
            y_min, y_max, u_min, u_max, v_min, v_max = args
            step = 0.01  # YUV使用固定步长
            
            total_combinations = int((y_max - y_min) / step) * int((u_max - u_min) / step) * int((v_max - v_min) / step)
            current_count = 0
            
            y_val = y_min
            while y_val <= y_max:
                u_val = u_min
                while u_val <= u_max:
                    v_val = v_min
                    while v_val <= v_max:
                        if not self.generating:
                            return
                        
                        # 转换为RGB
                        r = y_val + 1.4075 * (v_val - 0.5)
                        g = y_val - 0.3455 * (u_val - 0.5) - 0.7169 * (v_val - 0.5)
                        b = y_val + 1.7790 * (u_val - 0.5)
                        
                        r = max(0, min(1, r))
                        g = max(0, min(1, g))
                        b = max(0, min(1, b))
                        
                        r, g, b = [round(x*255) for x in (r, g, b)]
                        
                        color_data.append({
                            "color_space": "YUV",
                            "Y": round(y_val, 2),
                            "U": round(u_val, 2),
                            "V": round(v_val, 2),
                            "R": r,
                            "G": g,
                            "B": b,
                            "hex": f"#{r:02X}{g:02X}{b:02X}"
                        })
                        
                        current_count += 1
                        if current_count % 1000 == 0 or current_count == total_combinations:
                            progress = (current_count / total_combinations) * 100
                            self.update_progress(progress)
                        
                        v_val += step
                    u_val += step
                y_val += step
        
        elif color_space == "Lab":
            l_min, l_max, a_min, a_max, b_min, b_max = args
            step = int(self.step.get())
            
            total_combinations = ((l_max - l_min + 1) // step) * ((a_max - a_min + 1) // step) * ((b_max - b_min + 1) // step)
            current_count = 0
            
            for l_val in range(l_min, l_max + 1, step):
                for a_val in range(a_min, a_max + 1, step):
                    for b_val in range(b_min, b_max + 1, step):
                        if not self.generating:
                            return
                        
                        lab = LabColor(l_val, a_val, b_val)
                        rgb = convert_color(lab, sRGBColor)
                        r, g, b = [round(x*255) for x in rgb.get_value_tuple()]
                        
                        color_data.append({
                            "color_space": "Lab",
                            "L": l_val,
                            "a": a_val,
                            "b": b_val,
                            "R": r,
                            "G": g,
                            "B": b,
                            "hex": f"#{r:02X}{g:02X}{b:02X}"
                        })
                        
                        current_count += 1
                        if current_count % 1000 == 0 or current_count == total_combinations:
                            progress = (current_count / total_combinations) * 100
                            self.update_progress(progress)
        
        elif color_space == "HEX":
            hex_start, hex_end = args
            step = int(self.step.get())
            
            start_val = int(hex_start, 16)
            end_val = int(hex_end, 16)
            
            if start_val > end_val:
                start_val, end_val = end_val, start_val
            
            total_combinations = (end_val - start_val + 1) // step
            current_count = 0
            
            for val in range(start_val, end_val + 1, step):
                if not self.generating:
                    return
                
                hex_str = f"{val:06X}"
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                
                color_data.append({
                    "color_space": "HEX",
                    "hex": hex_str,
                    "R": r,
                    "G": g,
                    "B": b
                })
                
                current_count += 1
                if current_count % 1000 == 0 or current_count == total_combinations:
                    progress = ((val - start_val) / (end_val - start_val)) * 100
                    self.update_progress(progress)
        
        # 写入JSON文件
        with open(file_path, 'w') as jsonfile:
            json.dump({
                "metadata": {
                    "color_space": color_space,
                    "generated_at": datetime.now().isoformat(),
                    "total_colors": len(color_data)
                },
                "colors": color_data
            }, jsonfile, indent=2)
    
    def generate_sqlite(self, file_path, color_space, *args):
        conn = sqlite3.connect(file_path)
        cursor = conn.cursor()
        
        # 创建表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS colors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            color_space TEXT,
            r INTEGER,
            g INTEGER,
            b INTEGER,
            hex TEXT,
            h REAL,
            s REAL,
            l REAL,
            c REAL,
            m REAL,
            y REAL,
            k REAL,
            y_val REAL,
            u REAL,
            v REAL,
            l_val REAL,
            a REAL,
            b_val REAL,
            generated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        
        if color_space == "RGB":
            r_min, r_max, g_min, g_max, b_min, b_max = args
            step = int(self.step.get())
            
            total_combinations = ((r_max - r_min + 1) // step) * ((g_max - g_min + 1) // step) * ((b_max - b_min + 1) // step)
            current_count = 0
            
            for r in range(r_min, r_max + 1, step):
                for g in range(g_min, g_max + 1, step):
                    for b in range(b_min, b_max + 1, step):
                        if not self.generating:
                            conn.close()
                            return
                        
                        cursor.execute('''
                        INSERT INTO colors (color_space, r, g, b, hex)
                        VALUES (?, ?, ?, ?, ?)
                        ''', (
                            "RGB",
                            r,
                            g,
                            b,
                            f"#{r:02X}{g:02X}{b:02X}"
                        ))
                        
                        current_count += 1
                        if current_count % 1000 == 0 or current_count == total_combinations:
                            progress = (current_count / total_combinations) * 100
                            self.update_progress(progress)
                            conn.commit()
        
        elif color_space == "HSL":
            h_min, h_max, s_min, s_max, l_min, l_max = args
            step = int(self.step.get())
            
            total_combinations = ((h_max - h_min + 1) // step) * ((s_max - s_min + 1) // step) * ((l_max - l_min + 1) // step)
            current_count = 0
            
            for h in range(h_min, h_max + 1, step):
                for s in range(s_min, s_max + 1, step):
                    for l_val in range(l_min, l_max + 1, step):
                        if not self.generating:
                            conn.close()
                            return
                        
                        rgb = colorsys.hls_to_rgb(h/360, l_val/100, s/100)
                        r, g, b = [round(x*255) for x in rgb]
                        
                        cursor.execute('''
                        INSERT INTO colors (color_space, h, s, l, r, g, b, hex)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            "HSL",
                            h,
                            s,
                            l_val,
                            r,
                            g,
                            b,
                            f"#{r:02X}{g:02X}{b:02X}"
                        ))
                        
                        current_count += 1
                        if current_count % 1000 == 0 or current_count == total_combinations:
                            progress = (current_count / total_combinations) * 100
                            self.update_progress(progress)
                            conn.commit()
        
        elif color_space == "CMYK":
            c_min, c_max, m_min, m_max, y_min, y_max, k_min, k_max = args
            step = int(self.step.get())
            
            total_combinations = ((c_max - c_min + 1) // step) * ((m_max - m_min + 1) // step) * ((y_max - y_min + 1) // step) * ((k_max - k_min + 1) // step)
            current_count = 0
            
            for c in range(c_min, c_max + 1, step):
                for m in range(m_min, m_max + 1, step):
                    for y in range(y_min, y_max + 1, step):
                        for k in range(k_min, k_max + 1, step):
                            if not self.generating:
                                conn.close()
                                return
                            
                            cmyk = CMYKColor(c/100, m/100, y/100, k/100)
                            rgb = convert_color(cmyk, sRGBColor)
                            r, g, b = [round(x*255) for x in rgb.get_value_tuple()]
                            
                            cursor.execute('''
                            INSERT INTO colors (color_space, c, m, y, k, r, g, b, hex)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', (
                                "CMYK",
                                c,
                                m,
                                y,
                                k,
                                r,
                                g,
                                b,
                                f"#{r:02X}{g:02X}{b:02X}"
                            ))
                            
                            current_count += 1
                            if current_count % 1000 == 0 or current_count == total_combinations:
                                progress = (current_count / total_combinations) * 100
                                self.update_progress(progress)
                                conn.commit()
        
        elif color_space == "YUV":
            y_min, y_max, u_min, u_max, v_min, v_max = args
            step = 0.01  # YUV使用固定步长
            
            total_combinations = int((y_max - y_min) / step) * int((u_max - u_min) / step) * int((v_max - v_min) / step)
            current_count = 0
            
            y_val = y_min
            while y_val <= y_max:
                u_val = u_min
                while u_val <= u_max:
                    v_val = v_min
                    while v_val <= v_max:
                        if not self.generating:
                            conn.close()
                            return
                        
                        # 转换为RGB
                        r = y_val + 1.4075 * (v_val - 0.5)
                        g = y_val - 0.3455 * (u_val - 0.5) - 0.7169 * (v_val - 0.5)
                        b = y_val + 1.7790 * (u_val - 0.5)
                        
                        r = max(0, min(1, r))
                        g = max(0, min(1, g))
                        b = max(0, min(1, b))
                        
                        r, g, b = [round(x*255) for x in (r, g, b)]
                        
                        cursor.execute('''
                        INSERT INTO colors (color_space, y_val, u, v, r, g, b, hex)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            "YUV",
                            round(y_val, 2),
                            round(u_val, 2),
                            round(v_val, 2),
                            r,
                            g,
                            b,
                            f"#{r:02X}{g:02X}{b:02X}"
                        ))
                        
                        current_count += 1
                        if current_count % 1000 == 0 or current_count == total_combinations:
                            progress = (current_count / total_combinations) * 100
                            self.update_progress(progress)
                            conn.commit()
                        
                        v_val += step
                    u_val += step
                y_val += step
        
        elif color_space == "Lab":
            l_min, l_max, a_min, a_max, b_min, b_max = args
            step = int(self.step.get())
            
            total_combinations = ((l_max - l_min + 1) // step) * ((a_max - a_min + 1) // step) * ((b_max - b_min + 1) // step)
            current_count = 0
            
            for l_val in range(l_min, l_max + 1, step):
                for a_val in range(a_min, a_max + 1, step):
                    for b_val in range(b_min, b_max + 1, step):
                        if not self.generating:
                            conn.close()
                            return
                        
                        lab = LabColor(l_val, a_val, b_val)
                        rgb = convert_color(lab, sRGBColor)
                        r, g, b = [round(x*255) for x in rgb.get_value_tuple()]
                        
                        cursor.execute('''
                        INSERT INTO colors (color_space, l_val, a, b_val, r, g, b, hex)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            "Lab",
                            l_val,
                            a_val,
                            b_val,
                            r,
                            g,
                            b,
                            f"#{r:02X}{g:02X}{b:02X}"
                        ))
                        
                        current_count += 1
                        if current_count % 1000 == 0 or current_count == total_combinations:
                            progress = (current_count / total_combinations) * 100
                            self.update_progress(progress)
                            conn.commit()
        
        elif color_space == "HEX":
            hex_start, hex_end = args
            step = int(self.step.get())
            
            start_val = int(hex_start, 16)
            end_val = int(hex_end, 16)
            
            if start_val > end_val:
                start_val, end_val = end_val, start_val
            
            total_combinations = (end_val - start_val + 1) // step
            current_count = 0
            
            for val in range(start_val, end_val + 1, step):
                if not self.generating:
                    conn.close()
                    return
                
                hex_str = f"{val:06X}"
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                
                cursor.execute('''
                INSERT INTO colors (color_space, hex, r, g, b)
                VALUES (?, ?, ?, ?, ?)
                ''', (
                    "HEX",
                    hex_str,
                    r,
                    g,
                    b
                ))
                
                current_count += 1
                if current_count % 1000 == 0 or current_count == total_combinations:
                    progress = ((val - start_val) / (end_val - start_val)) * 100
                    self.update_progress(progress)
                    conn.commit()
        
        conn.commit()
        conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = ColorSpaceGeneratorApp(root)
    root.mainloop()