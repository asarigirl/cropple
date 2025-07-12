import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
from PIL import Image, ImageTk, ImageFilter
import math
import json
import os
import tkinterdnd2
import re

try:
    RESAMPLE_LANCZOS = Image.Resampling.LANCZOS
    RESAMPLE_NEAREST = Image.Resampling.NEAREST
    RESAMPLE_BICUBIC = Image.Resampling.BICUBIC
    ROTATE_90 = Image.Transpose.ROTATE_90
    ROTATE_270 = Image.Transpose.ROTATE_270
except AttributeError: 
    RESAMPLE_LANCZOS = Image.LANCZOS
    RESAMPLE_NEAREST = Image.NEAREST
    RESAMPLE_BICUBIC = Image.BICUBIC
    ROTATE_90 = Image.ROTATE_90
    ROTATE_270 = Image.ROTATE_270

class CropApp:
    SETTINGS_FILE_NAME = ".cropple_settings.json"
    DEFAULT_ASPECT_W = "16"
    DEFAULT_ASPECT_H = "9"
    DEFAULT_BLUR_RADIUS = 70
    WINDOW_SIZE_RATIO = 0.75
    MIN_WINDOW_WIDTH = 780 
    MIN_WINDOW_HEIGHT_CONTROLS = 320 
    MIN_CANVAS_HEIGHT = 200          

    ASPECT_PRESETS = {
        "オリジナル": "original", "1:1": (1, 1), "カスタム": "custom", "自由選択": "free",
        "16:9": (16, 9), "9:16": (9, 16), "4:3": (4, 3), "3:4": (3, 4), 
        "3:2": (3, 2), "2:3": (2, 3), "5:4": (5, 4),"4:5": (4, 5), 
         "1:1.91": (100, 191)
    }
    PRESET_ORDER_ROW1 = ["オリジナル", "1:1", "16:9", "9:16", "4:3", "3:4"]
    PRESET_ORDER_ROW2 = ["3:2", "2:3", "5:4", "4:5", "1:1.91", "カスタム", "自由選択"]

    def __init__(self, master):
        self.master = master
        master.title("画像トリミング＆拡張アプリ-くろっぷる (cropple)")

        self.max_window_width = int(master.winfo_screenwidth() * self.WINDOW_SIZE_RATIO)
        self.max_window_height = int(master.winfo_screenheight() * self.WINDOW_SIZE_RATIO)

        self.settings_file_path = os.path.join(os.path.expanduser("~"), self.SETTINGS_FILE_NAME)
        self.image_path = None
        self.original_pil_image = None
        self.processed_pil_image = None
        self.active_pil_for_canvas = None
        self.display_pil_image = None
        self.tk_image = None
        self.image_on_canvas = None 
        self.rect = None 
        self.start_x = None
        self.start_y = None
        self._processing_aspect_change = False
        self.rotation_angle_var = tk.DoubleVar(value=0)
        self.rotation_fill_color_var = tk.StringVar(value="#CCCCCC")
        self.rotation_fill_mode_var = tk.StringVar(value="color") 
        self.extend_position_var = tk.StringVar(value="center")


        self.top_controls_area = ttk.Frame(master)
        self.top_controls_area.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10,0))

        group1_frame = ttk.Frame(self.top_controls_area)
        group1_frame.pack(fill=tk.X)
        self.mode_frame = ttk.LabelFrame(group1_frame, text="モード選択")
        self.mode_frame.pack(side=tk.LEFT, padx=(0,5), pady=5, fill=tk.Y, anchor='nw')
        self.mode = tk.StringVar(value="crop")
        self.crop_radio = ttk.Radiobutton(self.mode_frame, text="切り抜きモード", variable=self.mode, value="crop", command=self.on_mode_change)
        self.crop_radio.pack(anchor='w', padx=5)
        self.extend_radio = ttk.Radiobutton(self.mode_frame, text="拡張モード", variable=self.mode, value="extend", command=self.on_mode_change)
        self.extend_radio.pack(anchor='w', padx=5)

        self.aspect_ratio_outer_frame = ttk.LabelFrame(group1_frame, text="アスペクト比")
        self.aspect_ratio_outer_frame.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        self.aspect_choice_var = tk.StringVar(value="16:9")
        self.aspect_preset_row1_frame = ttk.Frame(self.aspect_ratio_outer_frame)
        self.aspect_preset_row1_frame.pack(fill="x", anchor="w", pady=(5,0))
        for preset_name in self.PRESET_ORDER_ROW1:
            rb = ttk.Radiobutton(self.aspect_preset_row1_frame, text=preset_name, variable=self.aspect_choice_var, value=preset_name, command=self.on_aspect_choice_change)
            rb.pack(side=tk.LEFT, padx=3, pady=2)
        self.aspect_preset_row2_frame = ttk.Frame(self.aspect_ratio_outer_frame)
        self.aspect_preset_row2_frame.pack(fill="x", anchor="w")
        for preset_name in self.PRESET_ORDER_ROW2:
            rb = ttk.Radiobutton(self.aspect_preset_row2_frame, text=preset_name, variable=self.aspect_choice_var, value=preset_name, command=self.on_aspect_choice_change)
            rb.pack(side=tk.LEFT, padx=3, pady=2)
        self.custom_aspect_inputs_frame = ttk.Frame(self.aspect_ratio_outer_frame)
        self.custom_aspect_inputs_frame.pack(pady=(5,5), fill="x", anchor="w")
        self.aspect_w_var = tk.StringVar(); self.aspect_h_var = tk.StringVar() 
        self.aspect_w_var.trace_add("write", self.on_custom_aspect_entry_write)
        self.aspect_h_var.trace_add("write", self.on_custom_aspect_entry_write)
        ttk.Label(self.custom_aspect_inputs_frame, text="カスタム幅:").pack(side=tk.LEFT, padx=(5,2), pady=2)
        self.aspect_w_entry = ttk.Entry(self.custom_aspect_inputs_frame, textvariable=self.aspect_w_var, width=5, state=tk.DISABLED)
        self.aspect_w_entry.pack(side=tk.LEFT, padx=(0,10), pady=2)
        ttk.Label(self.custom_aspect_inputs_frame, text="カスタム高さ:").pack(side=tk.LEFT, padx=(0,2), pady=2)
        self.aspect_h_entry = ttk.Entry(self.custom_aspect_inputs_frame, textvariable=self.aspect_h_var, width=5, state=tk.DISABLED)
        self.aspect_h_entry.pack(side=tk.LEFT, padx=(0,10), pady=2)

        group2_frame = ttk.Frame(self.top_controls_area)
        group2_frame.pack(fill=tk.X)
        self.blur_frame = ttk.LabelFrame(group2_frame, text="拡張モード設定")
        self.blur_frame.pack(side=tk.LEFT, padx=(0,5), pady=5, fill=tk.Y, anchor='nw')

        # ぼかし半径コントロール用のサブフレーム
        blur_radius_subframe = ttk.Frame(self.blur_frame)
        blur_radius_subframe.pack(anchor='w')
        ttk.Label(blur_radius_subframe, text="ぼかし半径:").pack(side=tk.LEFT, padx=(5,2), pady=5)
        self.blur_radius_var = tk.IntVar()
        self.blur_radius_scale = ttk.Scale(blur_radius_subframe, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.blur_radius_var, length=150, command=lambda s: self.blur_radius_var.set(int(float(s))))
        self.blur_radius_scale.pack(side=tk.LEFT, padx=5, pady=5)
        self.blur_radius_label = ttk.Label(blur_radius_subframe, text="")
        self.blur_radius_label.pack(side=tk.LEFT, padx=(0,5), pady=5)
        self.blur_radius_var.trace_add("write", lambda *args: self.blur_radius_label.config(text=str(self.blur_radius_var.get())))

        # 画像配置コントロール用のサブフレーム
        position_controls_subframe = ttk.Frame(self.blur_frame)
        position_controls_subframe.pack(anchor='w')
        ttk.Label(position_controls_subframe, text="画像配置:").pack(side=tk.LEFT, padx=(5, 5), pady=(0,5))
        self.pos_center_radio = ttk.Radiobutton(position_controls_subframe, text="中央", variable=self.extend_position_var, value="center")
        self.pos_center_radio.pack(side=tk.LEFT, pady=(0,5))
        self.pos_top_radio = ttk.Radiobutton(position_controls_subframe, text="上", variable=self.extend_position_var, value="top")
        self.pos_top_radio.pack(side=tk.LEFT, pady=(0,5))
        self.pos_bottom_radio = ttk.Radiobutton(position_controls_subframe, text="下", variable=self.extend_position_var, value="bottom")
        self.pos_bottom_radio.pack(side=tk.LEFT, pady=(0,5))
        self.pos_left_radio = ttk.Radiobutton(position_controls_subframe, text="左", variable=self.extend_position_var, value="left")
        self.pos_left_radio.pack(side=tk.LEFT, pady=(0,5))
        self.pos_right_radio = ttk.Radiobutton(position_controls_subframe, text="右", variable=self.extend_position_var, value="right")
        self.pos_right_radio.pack(side=tk.LEFT, pady=(0,5))

        self.rotation_frame = ttk.LabelFrame(group2_frame, text="回転")
        self.rotation_frame.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        rotation_buttons_frame = ttk.Frame(self.rotation_frame)
        rotation_buttons_frame.pack(fill=tk.X)
        self.rotate_left_button = ttk.Button(rotation_buttons_frame, text="左90°", command=lambda: self.apply_rotation_transpose(ROTATE_270), state=tk.DISABLED)
        self.rotate_left_button.pack(side=tk.LEFT, padx=5, pady=2)
        self.rotate_right_button = ttk.Button(rotation_buttons_frame, text="右90°", command=lambda: self.apply_rotation_transpose(ROTATE_90), state=tk.DISABLED)
        self.rotate_right_button.pack(side=tk.LEFT, padx=5, pady=2)
        self.reset_rotation_button = ttk.Button(rotation_buttons_frame, text="回転リセット", command=self.reset_all_rotation, state=tk.DISABLED)
        self.reset_rotation_button.pack(side=tk.LEFT, padx=5, pady=2)
        
        free_rotation_frame = ttk.Frame(self.rotation_frame)
        free_rotation_frame.pack(fill=tk.X, pady=(5,0))
        self.rotation_slider = ttk.Scale(free_rotation_frame, from_=-45.0, to=45.0, orient=tk.HORIZONTAL, variable=self.rotation_angle_var, length=120, command=self.on_rotation_slider_change_preview, state=tk.DISABLED)
        self.rotation_slider.pack(side=tk.LEFT, padx=5, pady=2)
        self.rotation_label = ttk.Label(free_rotation_frame, text="0.0°", width=5, anchor="w")
        self.rotation_label.pack(side=tk.LEFT, padx=2, pady=2)
        self.rotation_angle_var.trace_add("write", self._update_rotation_label)
        self.apply_rotation_button = ttk.Button(free_rotation_frame, text="自由回転適用", command=self.apply_free_rotation, state=tk.DISABLED)
        self.apply_rotation_button.pack(side=tk.LEFT, padx=5, pady=2)
        
        rotation_fill_options_frame = ttk.Frame(self.rotation_frame) 
        rotation_fill_options_frame.pack(fill=tk.X, pady=(5,0))
        self.rotation_fill_mode_frame = ttk.Frame(rotation_fill_options_frame)
        self.rotation_fill_mode_frame.pack(side=tk.LEFT, padx=(0,5), pady=2)
        ttk.Label(self.rotation_fill_mode_frame, text="余白:").pack(side=tk.LEFT, pady=2)
        self.fill_mode_color_radio = ttk.Radiobutton(self.rotation_fill_mode_frame, text="指定色", variable=self.rotation_fill_mode_var, value="color", command=self._on_rotation_fill_mode_change)
        self.fill_mode_color_radio.pack(side=tk.LEFT, padx=2)
        self.fill_mode_transparent_radio = ttk.Radiobutton(self.rotation_fill_mode_frame, text="透過", variable=self.rotation_fill_mode_var, value="transparent", command=self._on_rotation_fill_mode_change)
        self.fill_mode_transparent_radio.pack(side=tk.LEFT, padx=2)
        # 注釈ラベルの作成
        self.transparent_note_label = ttk.Label(self.rotation_fill_mode_frame, text="(プレビューは背景色, PNG保存で透過)", font=("Arial", 7))
        # pack は _on_rotation_fill_mode_change で制御

        self.fill_color_subframe = ttk.Frame(rotation_fill_options_frame) 
        self.fill_color_subframe.pack(side=tk.LEFT, pady=2)
        self.rotation_fill_color_button = ttk.Button(self.fill_color_subframe, text="色選択", command=self.choose_rotation_fill_color, width=7)
        self.rotation_fill_color_button.pack(side=tk.LEFT, padx=2)
        self.rotation_fill_color_preview = tk.Label(self.fill_color_subframe, text="    ", relief=tk.SUNKEN, width=4)
        self.rotation_fill_color_preview.pack(side=tk.LEFT, padx=2)
        self.rotation_fill_color_entry = ttk.Entry(self.fill_color_subframe, textvariable=self.rotation_fill_color_var, width=9)
        self.rotation_fill_color_entry.pack(side=tk.LEFT, padx=2)
        self.rotation_fill_color_var.trace_add("write", self._update_rotation_fill_preview)
        
        self.settings_preview_buttons_frame = ttk.Frame(self.top_controls_area)
        self.settings_preview_buttons_frame.pack(pady=(5,0), fill="x")
        self.save_settings_button = ttk.Button(self.settings_preview_buttons_frame, text="設定を保存する", command=self.save_settings)
        self.save_settings_button.pack(side=tk.LEFT, padx=5, pady=2)
        self.default_settings_button = ttk.Button(self.settings_preview_buttons_frame, text="初期設定に戻す", command=self.apply_default_settings_to_ui)
        self.default_settings_button.pack(side=tk.LEFT, padx=5, pady=2)
        self.update_preview_button = ttk.Button(self.settings_preview_buttons_frame, text="拡張プレビュー更新", command=self.update_preview_action, state=tk.DISABLED)
        self.update_preview_button.pack(side=tk.LEFT, padx=5, pady=2)
        
        self.separator_after_settings = ttk.Separator(self.top_controls_area, orient=tk.HORIZONTAL)
        self.separator_after_settings.pack(fill='x', pady=5)

        self.file_reset_buttons_frame = ttk.Frame(self.top_controls_area)
        self.file_reset_buttons_frame.pack(pady=(0,10), fill="x")
        self.load_button = ttk.Button(self.file_reset_buttons_frame, text="画像を開く", command=self.load_image_dialog)
        self.load_button.pack(side=tk.LEFT, padx=5, pady=2)
        self.reset_image_button = ttk.Button(self.file_reset_buttons_frame, text="画像リセット", command=self.reset_image_processing, state=tk.DISABLED)
        self.reset_image_button.pack(side=tk.LEFT, padx=5, pady=2)
        self.execute_button = ttk.Button(self.file_reset_buttons_frame, text="実行して画像を保存", command=self.execute_action)
        self.execute_button.pack(side=tk.LEFT, padx=5, pady=2)

        self.canvas_frame = ttk.Frame(master)
        self.canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0,10))
        self.canvas = tk.Canvas(self.canvas_frame, cursor="cross", bg="lightgrey")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.drop_target_register(tkinterdnd2.DND_FILES)
        self.canvas.dnd_bind('<<Drop>>', self.handle_drop)
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)

        self.load_settings()
        self._display_image_on_canvas() 
        self.on_mode_change()
        self.on_aspect_choice_change()
        self._update_rotation_fill_preview()
        self._on_rotation_fill_mode_change()

    def _update_rotation_label(self, *args):
        self.rotation_label.config(text=f"{self.rotation_angle_var.get():.1f}°")

    def on_rotation_slider_change_preview(self, value_str):
        pass

    def choose_rotation_fill_color(self):
        color_code = colorchooser.askcolor(title="回転時の背景色を選択", initialcolor=self.rotation_fill_color_var.get())
        if color_code and color_code[1]:
            self.rotation_fill_color_var.set(color_code[1])

    def _update_rotation_fill_preview(self, *args):
        color = self.rotation_fill_color_var.get()
        try: self.rotation_fill_color_preview.config(bg=color)
        except tk.TclError: self.rotation_fill_color_preview.config(bg="SystemButtonFace")

    def _on_rotation_fill_mode_change(self): 
        has_image = bool(self.processed_pil_image)
        if self.rotation_fill_mode_var.get() == "color" and has_image:
            self.rotation_fill_color_button.config(state=tk.NORMAL)
            self.rotation_fill_color_entry.config(state=tk.NORMAL)
            self.transparent_note_label.pack_forget() # 指定色時は注釈非表示
        else: 
            self.rotation_fill_color_button.config(state=tk.DISABLED)
            self.rotation_fill_color_entry.config(state=tk.DISABLED)
            if self.rotation_fill_mode_var.get() == "transparent" and has_image:
                 self.transparent_note_label.pack(side=tk.LEFT, padx=2, pady=2) # 透過選択時で画像ありなら表示
            else:
                 self.transparent_note_label.pack_forget() # それ以外は非表示
        self._update_rotation_fill_preview() 

    def apply_rotation_transpose(self, transpose_mode):
        if not self.processed_pil_image: return
        self.processed_pil_image = self.processed_pil_image.transpose(transpose_mode)
        self.rotation_angle_var.set(0) 
        self.active_pil_for_canvas = self.processed_pil_image
        self._display_image_on_canvas()
        if self.aspect_choice_var.get() == "オリジナル": self.on_aspect_choice_change()

    def apply_free_rotation(self):
        if not self.processed_pil_image: return
        angle_to_apply = self.rotation_angle_var.get()
        if abs(angle_to_apply) < 0.1: return
        
        image_to_rotate = self.processed_pil_image 
        fill_color_tuple = None

        if self.rotation_fill_mode_var.get() == "transparent":
            if image_to_rotate.mode != 'RGBA':
                image_to_rotate = image_to_rotate.convert('RGBA')
            fill_color_tuple = (0,0,0,0) 
        else: # "color"
            fill_color_hex = self.rotation_fill_color_var.get()
            try:
                r,g,b = tuple(int(fill_color_hex.lstrip('#')[i:i+2],16) for i in (0,2,4))
                fill_color_tuple = (r,g,b,255) if image_to_rotate.mode == 'RGBA' or image_to_rotate.mode == 'LA' else (r,g,b) # LAモードも考慮
            except ValueError:
                messagebox.showerror("色指定エラー", "背景色のHEXコードが無効です。デフォルトのグレーを使用します。")
                fill_color_tuple = (200,200,200,255) if image_to_rotate.mode == 'RGBA' or image_to_rotate.mode == 'LA' else (200,200,200)
        
        try:
            # rotate前に image_to_rotate を self.processed_pil_image に代入しないように注意
            rotated_image = image_to_rotate.rotate(
                -angle_to_apply, resample=RESAMPLE_BICUBIC, expand=True, fillcolor=fill_color_tuple
            )
            # 回転後の画像を processed_pil_image に設定
            self.processed_pil_image = rotated_image

        except Exception as e: messagebox.showerror("回転エラー", f"自由回転処理中にエラー: {e}"); return
        
        self.rotation_angle_var.set(0)
        self.active_pil_for_canvas = self.processed_pil_image
        self._display_image_on_canvas()
        if self.aspect_choice_var.get() == "オリジナル": self.on_aspect_choice_change()

    def reset_all_rotation(self): 
        if not self.original_pil_image: return
        self.processed_pil_image = self.original_pil_image.copy()
        self.rotation_angle_var.set(0)
        self.active_pil_for_canvas = self.processed_pil_image
        self._display_image_on_canvas()
        if self.aspect_choice_var.get() == "オリジナル": self.on_aspect_choice_change()

    def reset_image_processing(self):
        if not self.original_pil_image: messagebox.showwarning("リセット不可", "画像が読み込まれていません。"); return
        self.processed_pil_image = self.original_pil_image.copy()
        self.active_pil_for_canvas = self.processed_pil_image
        self.rotation_angle_var.set(0)
        if self.rect: self.canvas.delete(self.rect); self.rect = None
        self._display_image_on_canvas()
        self.on_mode_change() 
        self.aspect_choice_var.set("16:9")
        self.on_aspect_choice_change()
        messagebox.showinfo("画像リセット", "画像の状態を読み込み直後にリセットしました（設定は維持されます）。")

    def handle_drop(self, event):
        filepaths_str = event.data; raw_paths = []
        if filepaths_str.startswith('{') and filepaths_str.endswith('}'):
            content_inside_outer_braces = filepaths_str[1:-1]
            matches = re.findall(r'\{([^}]+)\}|([^}{}\s]+)', content_inside_outer_braces)
            for m in matches: raw_paths.append(m[0] if m[0] else m[1])
            if not raw_paths and content_inside_outer_braces: raw_paths.append(content_inside_outer_braces)
        else: raw_paths = filepaths_str.split(' ')
        cleaned_paths = []
        for p in raw_paths:
            p_cleaned = p.strip('"').strip("'")
            if os.path.exists(p_cleaned) and os.path.isfile(p_cleaned): cleaned_paths.append(p_cleaned)
        if cleaned_paths:
            for path_to_load in cleaned_paths:
                if path_to_load.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    self.load_image(path_to_load); return 
            messagebox.showwarning("ドロップエラー", "ドロップされた有効な画像ファイルが見つかりませんでした。")
        else: messagebox.showwarning("ドロップエラー", f"ドロップされたファイルパスを解析できませんでした。\nData: '{filepaths_str}'")

    def load_image_dialog(self):
        path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.gif"), ("All files", "*.*")])
        if path: self.load_image(path)

    def load_image(self, path):
        if not path: return
        try:
            self.original_pil_image = Image.open(path)
            if self.original_pil_image.mode not in ['RGB', 'L', 'RGBA', 'LA']: self.original_pil_image = self.original_pil_image.convert('RGB')
            elif self.original_pil_image.mode == 'P': self.original_pil_image = self.original_pil_image.convert('RGBA' if 'transparency' in self.original_pil_image.info else 'RGB')
            self.processed_pil_image = self.original_pil_image.copy()
            self.active_pil_for_canvas = self.processed_pil_image
            self.rotation_angle_var.set(0) 
        except Exception as e:
            messagebox.showerror("エラー", f"画像を開けませんでした: {path}\n{e}")
            self.original_pil_image=None; self.processed_pil_image=None; self.active_pil_for_canvas=None
            self._display_image_on_canvas(); self.on_mode_change()
            return
        self.image_path = path
        self._display_image_on_canvas()
        self.on_mode_change()

    def _display_image_on_canvas(self):
        self.master.update_idletasks()
        controls_height = self.top_controls_area.winfo_reqheight()
        canvas_frame_padx_sum = 20; canvas_frame_pady_sum = 15
        min_controls_width = self.MIN_WINDOW_WIDTH - canvas_frame_padx_sum - 20
        canvas_max_allowable_width = self.max_window_width - canvas_frame_padx_sum
        canvas_max_allowable_height = self.max_window_height - controls_height - canvas_frame_pady_sum - 20
        if not self.active_pil_for_canvas:
            placeholder_w=max(100,min_controls_width); placeholder_h=max(100, self.MIN_WINDOW_HEIGHT_CONTROLS - controls_height + self.MIN_CANVAS_HEIGHT - canvas_frame_pady_sum - 20 if self.MIN_WINDOW_HEIGHT_CONTROLS > controls_height else self.MIN_CANVAS_HEIGHT)
            self.canvas.config(width=placeholder_w,height=placeholder_h)
            if self.image_on_canvas: self.canvas.delete(self.image_on_canvas); self.image_on_canvas=None
            if self.tk_image: self.tk_image=None
            win_w=max(self.MIN_WINDOW_WIDTH,self.top_controls_area.winfo_reqwidth()+canvas_frame_padx_sum+20); win_h=max(self.MIN_WINDOW_HEIGHT_CONTROLS + self.MIN_CANVAS_HEIGHT, controls_height + placeholder_h + canvas_frame_pady_sum + 20)
            self.master.geometry(f"{min(win_w,self.max_window_width)}x{min(win_h,self.max_window_height)}")
            return
        img_width,img_height=self.active_pil_for_canvas.size
        ratio=1.0
        if img_width>0 and img_height>0: ratio=min(canvas_max_allowable_width/img_width,canvas_max_allowable_height/img_height,1.0)
        new_width=int(img_width*ratio); new_height=int(img_height*ratio)
        new_width=max(50,new_width); new_height=max(50,new_height)
        self.display_pil_image=self.active_pil_for_canvas.resize((new_width,new_height),RESAMPLE_LANCZOS)
        self.tk_image=ImageTk.PhotoImage(self.display_pil_image)
        self.canvas.config(width=new_width,height=new_height)
        if self.image_on_canvas: self.canvas.delete(self.image_on_canvas)
        if self.rect: self.canvas.delete(self.rect); self.rect=None; self.start_x=None; self.start_y=None
        self.image_on_canvas=self.canvas.create_image(0,0,anchor=tk.NW,image=self.tk_image)
        final_win_width=max(self.MIN_WINDOW_WIDTH, new_width+canvas_frame_padx_sum+20)
        final_win_height=max(self.MIN_WINDOW_HEIGHT_CONTROLS + self.MIN_CANVAS_HEIGHT, new_height+controls_height+canvas_frame_pady_sum+20)
        self.master.geometry(f"{min(final_win_width,self.max_window_width)}x{min(final_win_height,self.max_window_height)}")

    def on_custom_aspect_entry_write(self, *args):\n        if self._processing_aspect_change: return
        self.aspect_choice_var.set("カスタム")

    def on_aspect_choice_change(self, event=None):
        self._processing_aspect_change = True
        choice = self.aspect_choice_var.get()
        is_crop_mode = (self.mode.get() == "crop")

        if choice == "カスタム":
            self.aspect_w_entry.config(state=tk.NORMAL); self.aspect_h_entry.config(state=tk.NORMAL)
        else:
            self.aspect_w_entry.config(state=tk.DISABLED); self.aspect_h_entry.config(state=tk.DISABLED)
            if choice == "オリジナル":
                if self.processed_pil_image:
                    w,h=self.processed_pil_image.size; common=math.gcd(w,h)
                    self.aspect_w_var.set(str(w//common)); self.aspect_h_var.set(str(h//common))
                else: self.aspect_w_var.set(self.DEFAULT_ASPECT_W); self.aspect_h_var.set(self.DEFAULT_ASPECT_H)
            elif choice == "自由選択": pass
            elif choice in self.ASPECT_PRESETS:
                val = self.ASPECT_PRESETS[choice]
                if isinstance(val, tuple): self.aspect_w_var.set(str(val[0])); self.aspect_h_var.set(str(val[1]))
        
        if not is_crop_mode and choice == "自由選択":
            self.aspect_choice_var.set("16:9") 
            messagebox.showwarning("モードエラー", "拡張モードでは「自由選択」は使用できません。アスペクト比を16:9に戻しました。")
        self._processing_aspect_change = False

    def load_settings():
        try:
            with open(self.settings_file_path,'r') as f: settings=json.load(f)
            loaded_w=settings.get('aspect_w',self.DEFAULT_ASPECT_W); loaded_h=settings.get('aspect_h',self.DEFAULT_ASPECT_H)
            self.blur_radius_var.set(settings.get('blur_radius',self.DEFAULT_BLUR_RADIUS))
            self.rotation_fill_color_var.set(settings.get('rotation_fill_color', "#CCCCCC"))
            self.rotation_fill_mode_var.set(settings.get('rotation_fill_mode', "color"))
            self.extend_position_var.set(settings.get('extend_position', "center"))
            saved_aspect_choice = settings.get('aspect_choice', '16:9')
            if saved_aspect_choice in self.ASPECT_PRESETS: self.aspect_choice_var.set(saved_aspect_choice)
            else: self.aspect_choice_var.set("カスタム")
            self._processing_aspect_change=True
            if self.aspect_choice_var.get() == "カスタム" or self.ASPECT_PRESETS.get(self.aspect_choice_var.get()) == "original":
                self.aspect_w_var.set(loaded_w); self.aspect_h_var.set(loaded_h)
            self._processing_aspect_change=False
            self.on_aspect_choice_change()
        except(FileNotFoundError,json.JSONDecodeError): self.apply_default_settings_to_ui()
        self._update_rotation_fill_preview()
        self._on_rotation_fill_mode_change()

    def save_settings(self):
        settings={'aspect_w':self.aspect_w_var.get(), 'aspect_h':self.aspect_h_var.get(),
                  'aspect_choice': self.aspect_choice_var.get(),
                  'blur_radius':self.blur_radius_var.get(), 
                  'rotation_fill_color': self.rotation_fill_color_var.get(),
                  'rotation_fill_mode': self.rotation_fill_mode_var.get(),
                  'extend_position': self.extend_position_var.get()}
        try:
            with open(self.settings_file_path,'w') as f: json.dump(settings,f,indent=4)
            messagebox.showinfo("設定保存",f"設定を保存しました。\nパス: {self.settings_file_path}")
        except IOError as e: messagebox.showerror("設定保存エラー",f"設定の保存に失敗しました: {e}")

    def apply_default_settings_to_ui(self):
        self._processing_aspect_change=True
        self.aspect_w_var.set(self.DEFAULT_ASPECT_W); self.aspect_h_var.set(self.DEFAULT_ASPECT_H)
        self._processing_aspect_change=False
        self.blur_radius_var.set(self.DEFAULT_BLUR_RADIUS); self.aspect_choice_var.set("16:9")
        self.on_aspect_choice_change()
        self.rotation_angle_var.set(0); self.rotation_fill_color_var.set("#CCCCCC"); 
        self.rotation_fill_mode_var.set("color")
        self.extend_position_var.set("center")
        self._update_rotation_fill_preview()
        self._on_rotation_fill_mode_change()
        if self.original_pil_image: self.reset_image_processing()
        elif self.processed_pil_image: self.processed_pil_image=None; self.active_pil_for_canvas=None; self._display_image_on_canvas()

    def on_mode_change(self):
        mode=self.mode.get(); is_crop_mode=(mode=="crop"); is_extend_mode=(mode=="extend")
        has_image=bool(self.processed_pil_image)
        self.blur_radius_scale.config(state=tk.NORMAL if is_extend_mode else tk.DISABLED)
        pos_state = tk.NORMAL if is_extend_mode else tk.DISABLED
        self.pos_center_radio.config(state=pos_state)
        self.pos_top_radio.config(state=pos_state)
        self.pos_bottom_radio.config(state=pos_state)
        self.pos_left_radio.config(state=pos_state)
        self.pos_right_radio.config(state=pos_state)
        self.update_preview_button.config(state=tk.NORMAL if is_extend_mode and has_image else tk.DISABLED)
        rot_state = tk.NORMAL if has_image else tk.DISABLED
        self.rotate_left_button.config(state=rot_state); self.rotate_right_button.config(state=rot_state)
        self.reset_rotation_button.config(state=rot_state); self.rotation_slider.config(state=rot_state)
        self.apply_rotation_button.config(state=rot_state)
        self.fill_mode_color_radio.config(state=rot_state) 
        self.fill_mode_transparent_radio.config(state=rot_state) 
        self._on_rotation_fill_mode_change() 
        self.reset_image_button.config(state=tk.NORMAL if has_image else tk.DISABLED)
        self.canvas.config(cursor="cross" if is_crop_mode else "arrow")
        if self.rect: self.canvas.delete(self.rect); self.rect=None
        self.on_aspect_choice_change()
        if self.processed_pil_image: self.active_pil_for_canvas = self.processed_pil_image; self._display_image_on_canvas()
        else: self._display_image_on_canvas()

    def on_free_aspect_change(self): pass 

    def get_aspect_ratio_tuple(self):
        try:
            w_str=self.aspect_w_var.get(); h_str=self.aspect_h_var.get()
            if not w_str or not h_str: return None
            w=int(w_str); h=int(h_str)
            if w<=0 or h<=0: return None
            return w,h
        except ValueError: return None

    def on_button_press(self, event):
        if self.mode.get()!="crop" or not self.display_pil_image: return
        self.start_x=self.canvas.canvasx(event.x); self.start_y=self.canvas.canvasy(event.y)
        self.start_x=max(0,min(self.start_x,self.display_pil_image.width)); self.start_y=max(0,min(self.start_y,self.display_pil_image.height))
        if self.rect: self.canvas.delete(self.rect)
        self.rect=self.canvas.create_rectangle(self.start_x,self.start_y,self.start_x+1,self.start_y+1,outline='red',width=2)

    def on_mouse_drag(self, event):
        if self.mode.get()!="crop" or not self.rect or self.start_x is None: return
        cur_x=self.canvas.canvasx(event.x); cur_y=self.canvas.canvasy(event.y)
        canvas_w=self.display_pil_image.width; canvas_h=self.display_pil_image.height
        cur_x=max(0,min(cur_x,canvas_w)); cur_y=max(0,min(cur_y,canvas_h))
        end_x,end_y = cur_x,cur_y
        if self.aspect_choice_var.get() != "自由選択":
            aspect_tuple=self.get_aspect_ratio_tuple()
            if aspect_tuple:
                aspect_w_ratio,aspect_h_ratio = aspect_tuple
                delta_x=cur_x-self.start_x; delta_y=cur_y-self.start_y
                if delta_x==0 and delta_y==0: self.canvas.coords(self.rect,self.start_x,self.start_y,self.start_x+1,self.start_y+1); return
                potential_w_from_h=abs(delta_y)*(aspect_w_ratio/aspect_h_ratio); potential_h_from_w=abs(delta_x)/(aspect_w_ratio/aspect_h_ratio)
                if potential_w_from_h<=abs(delta_x): final_abs_h=abs(delta_y); final_abs_w=potential_w_from_h
                else: final_abs_w=abs(delta_x); final_abs_h=potential_h_from_w
                end_x=self.start_x+math.copysign(final_abs_w,delta_x); end_y=self.start_y+math.copysign(final_abs_h,delta_y)
        self.canvas.coords(self.rect,self.start_x,self.start_y,end_x,end_y)

    def execute_action(self):
        if not self.processed_pil_image: messagebox.showwarning("警告","まず画像を読み込んでください。"); return
        mode = self.mode.get()
        if mode=="crop": self.crop_image_action()
        elif mode=="extend": self.extend_image_and_save()

    def crop_image_action(self):
        if not self.processed_pil_image: messagebox.showwarning("切り抜き不可", "処理対象の画像が読み込まれていません。"); return
        image_to_save = None; operation_description = "現在の画像全体"
        if self.rect:
            coords = self.canvas.coords(self.rect)
            if coords and len(coords) == 4:
                source_for_crop = self.processed_pil_image 
                c_x1,c_y1,c_x2,c_y2=map(float,coords)
                norm_c_left=min(c_x1,c_x2); norm_c_top=min(c_y1,c_y2); norm_c_right=max(c_x1,c_x2); norm_c_bottom=max(c_y1,c_y2)
                if not self.display_pil_image: messagebox.showerror("エラー", "表示中の画像がありません。"); return
                disp_w=self.display_pil_image.width; disp_h=self.display_pil_image.height
                src_w,src_h=source_for_crop.size
                scale_x=src_w/disp_w if disp_w>0 else 1; scale_y=src_h/disp_h if disp_h>0 else 1
                crop_left=int(norm_c_left*scale_x); crop_top=int(norm_c_top*scale_y); crop_right=int(norm_c_right*scale_x); crop_bottom=int(norm_c_bottom*scale_y)
                crop_left=max(0,crop_left); crop_top=max(0,crop_top); crop_right=min(src_w,crop_right); crop_bottom=min(src_h,crop_bottom)
                if crop_left < crop_right and crop_top < crop_bottom:
                    cropped_pil = source_for_crop.crop((crop_left,crop_top,crop_right,crop_bottom))
                    self.processed_pil_image = cropped_pil; self.active_pil_for_canvas = self.processed_pil_image
                    self._display_image_on_canvas() 
                    if self.rect: self.canvas.delete(self.rect); self.rect = None
                    image_to_save = self.processed_pil_image; operation_description = "切り抜き後の画像"
                else: messagebox.showwarning("切り抜き範囲無効", "選択された切り抜き範囲が無効です。現在の画像全体を保存します。"); image_to_save = self.processed_pil_image
            else: image_to_save = self.processed_pil_image; messagebox.showinfo("情報", "切り抜き範囲の取得に失敗しました。現在の画像全体を保存します。")
        else: image_to_save = self.processed_pil_image; messagebox.showinfo("情報", "切り抜き範囲が選択されていません。現在の画像全体を保存します。")
        if not image_to_save: messagebox.showerror("エラー", "保存する画像がありません。"); return
        save_path=filedialog.asksaveasfilename(title=f"{operation_description}を保存", defaultextension=".png", filetypes=[("PNG files","*.png"),("JPEG files","*.jpg"),("All files","*.*")])
        if save_path:
            try: image_to_save.save(save_path); messagebox.showinfo("成功", f"{operation_description}を保存しました: {save_path}")
            except Exception as e: messagebox.showerror("エラー", f"画像の保存に失敗しました: {e}")

    def _generate_extended_image(self, source_image_for_processing):
        if not source_image_for_processing: return None
        aspect_tuple=self.get_aspect_ratio_tuple();
        if not aspect_tuple: return None
        target_aspect_w,target_aspect_h=aspect_tuple; target_ar=target_aspect_w/target_aspect_h
        orig_w,orig_h=source_image_for_canvas.size 
        current_ar = orig_w / orig_h if orig_h > 0 else float('inf')
        if abs(target_ar-current_ar)<1e-6 : return source_image_for_processing.copy()
        final_w,final_h = 0,0
        if target_ar > current_ar: final_h=orig_h; final_w=int(round(orig_h*target_ar))
        else: final_w=orig_w; final_h=int(round(orig_w/target_ar))
        if final_w<=0 or final_h<=0: return None
        output_mode = source_image_for_processing.mode
        initial_fill_for_extended = (0,0,0,0) if 'A' in output_mode else (200,200,200)
        extended_image = Image.new(output_mode, (final_w, final_h), initial_fill_for_extended)

        position = self.extend_position_var.get()
        # Y座標 (上下方向) の計算
        if final_h > orig_h: # 縦に拡張される場合
            if position == "top": paste_y_in_extended = 0
            elif position == "bottom": paste_y_in_extended = final_h - orig_h
            else: paste_y_in_extended = (final_h - orig_h) // 2 # "center", "left", "right"
        else: paste_y_in_extended = 0 # 横拡張 (final_h == orig_h)
        # X座標 (左右方向) の計算
        if final_w > orig_w: # 横に拡張される場合
            if position == "left": paste_x_in_extended = 0
            elif position == "right": paste_x_in_extended = final_w - orig_w
            else: paste_x_in_extended = (final_w - orig_w) // 2 # "center", "top", "bottom"
        else: paste_x_in_extended = 0 # 縦拡張 (final_w == orig_w)

        extended_image.paste(source_image_for_processing, (paste_x_in_extended, paste_y_in_extended), mask=source_image_for_processing if 'A' in output_mode else None)

        top_padding_fill=paste_y_in_extended; bottom_padding_fill=final_h-(paste_y_in_extended+orig_h)
        left_padding_fill=paste_x_in_extended; right_padding_fill=final_w-(paste_x_in_extended+orig_w)
        blur_radius_val = self.blur_radius_var.get()
        if top_padding_fill > 0:
            desired_source_thickness=max(1,top_padding_fill//2); actual_source_thickness=min(orig_h,desired_source_thickness)
            source_material=source_image_for_processing.crop((0,0,orig_w,actual_source_thickness))
            if actual_source_thickness<desired_source_thickness and desired_source_thickness>0: source_material=source_material.resize((orig_w,desired_source_thickness),RESAMPLE_LANCZOS)
            if blur_radius_val>0 and source_material.width>0 and source_material.height>0: blurred_material=source_material.filter(ImageFilter.GaussianBlur(blur_radius_val))
            else: blurred_material=source_material
            if blurred_material.width>0 and blurred_material.height>0: fill_content=blurred_material.resize((final_w,top_padding_fill),RESAMPLE_LANCZOS); extended_image.paste(fill_content,(0,0), mask=fill_content if 'A' in output_mode else None)
        if bottom_padding_fill > 0:
            desired_source_thickness=max(1,bottom_padding_fill//2); actual_source_thickness=min(orig_h,desired_source_thickness)
            source_material=source_image_for_processing.crop((0,orig_h-actual_source_thickness,orig_w,orig_h))
            if actual_source_thickness<desired_source_thickness and desired_source_thickness>0: source_material=source_material.resize((orig_w,desired_source_thickness),RESAMPLE_LANCZOS)
            if blur_radius_val>0 and source_material.width>0 and source_material.height>0: blurred_material=source_material.filter(ImageFilter.GaussianBlur(blur_radius_val))
            else: blurred_material=source_material
            if blurred_material.width>0 and blurred_material.height>0: fill_content=blurred_material.resize((final_w,bottom_padding_fill),RESAMPLE_LANCZOS); extended_image.paste(fill_content,(0, paste_y_in_extended + orig_h), mask=fill_content if 'A' in output_mode else None)
        if left_padding_fill > 0:
            desired_source_thickness=max(1,left_padding_fill//2); actual_source_thickness=min(orig_w,desired_source_thickness)
            source_material=source_image_for_processing.crop((0,0,actual_source_thickness,orig_h))
            if actual_source_thickness<desired_source_thickness and desired_source_thickness>0: source_material=source_material.resize((desired_source_thickness,orig_h),RESAMPLE_LANCZOS)
            if blur_radius_val>0 and source_material.width>0 and source_material.height>0: blurred_material=source_material.filter(ImageFilter.GaussianBlur(blur_radius_val))
            else: blurred_material=source_material
            if blurred_material.width>0 and blurred_material.height>0: fill_content=blurred_material.resize((left_padding_fill, final_h),RESAMPLE_LANCZOS); extended_image.paste(fill_content,(0,0), mask=fill_content if 'A' in output_mode else None)
        if right_padding_fill > 0:
            desired_source_thickness=max(1,right_padding_fill//2); actual_source_thickness=min(orig_w,desired_source_thickness)
            source_material=source_image_for_processing.crop((orig_w-actual_source_thickness,0,orig_w,orig_h))
            if actual_source_thickness<desired_source_thickness and desired_source_thickness>0: source_material=source_material.resize((desired_source_thickness,orig_h),RESAMPLE_LANCZOS)
            if blur_radius_val>0 and source_material.width>0 and source_material.height>0: blurred_material=source_material.filter(ImageFilter.GaussianBlur(blur_radius_val))
            else: blurred_material=source_material
            if blurred_material.width>0 and blurred_material.height>0: fill_content=blurred_material.resize((right_padding_fill,final_h),RESAMPLE_LANCZOS); extended_image.paste(fill_content,(paste_x_in_extended + orig_w, 0), mask=fill_content if 'A' in output_mode else None)
        extended_image.paste(source_image_for_processing,(paste_x_in_extended,paste_y_in_extended), mask=source_image_for_processing if 'A' in output_mode else None)
        return extended_image

    def update_preview_action(self):
        if self.mode.get()!="extend" or not self.processed_pil_image: messagebox.showwarning("プレビューエラー","拡張モードで画像を開いてからプレビューを更新してください。"); return
        preview_image=self._generate_extended_image(self.processed_pil_image)
        if preview_image: self.active_pil_for_canvas=preview_image; self._display_image_on_canvas()
        else: messagebox.showerror("プレビューエラー","プレビュー画像の生成に失敗しました。設定を確認してください。")

    def extend_image_and_save(self):
        if not self.processed_pil_image: messagebox.showwarning("警告","まず画像を読み込んでください。"); return
        final_image_to_save=self._generate_extended_image(self.processed_pil_image)
        if not final_image_to_save: messagebox.showerror("エラー","拡張画像の生成に失敗しました。設定を確認してください。"); return
        try:
            current_settings_ar_tuple=self.get_aspect_ratio_tuple()
            if current_settings_ar_tuple and self.processed_pil_image.height > 0 :
                current_settings_ar=current_settings_ar_tuple[0]/current_settings_ar_tuple[1]
                if abs(current_settings_ar-(self.processed_pil_image.width/self.processed_pil_image.height))<1e-6:
                     messagebox.showinfo("情報","画像は既に指定されたアスペクト比です。拡張処理はスキップされました。")
        except(TypeError,ZeroDivisionError,AttributeError):pass
        save_path=filedialog.asksaveasfilename(defaultextension=".png",filetypes=[("PNG files","*.png"),("JPEG files","*.jpg"),("All files","*.*")])
        if save_path:
            try: final_image_to_save.save(save_path); messagebox.showinfo("成功",f"画像を拡張して保存しました: {save_path}")
            except Exception as e: messagebox.showerror("エラー",f"画像の保存に失敗しました: {e}")

if __name__ == "__main__":
    root = tkinterdnd2.Tk()
    app = CropApp(root)
    initial_width = max(app.MIN_WINDOW_WIDTH, int(root.winfo_screenwidth() * 0.5))
    initial_height = max(app.MIN_WINDOW_HEIGHT_CONTROLS + app.MIN_CANVAS_HEIGHT, int(root.winfo_screenheight() * 0.6))
    root.geometry(f"{min(initial_width, app.max_window_width)}x{min(initial_height, app.max_window_height)}")
    root.mainloop()