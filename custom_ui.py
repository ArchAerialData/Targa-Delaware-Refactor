import customtkinter as ctk
from tkinter import filedialog
import os
import subprocess
import sys
import threading
import queue
from report_generator import OUTPUT_FOLDER_NAME

TRD_PIPELINES = {
    "RHNGL": "Red Hills NGL",
    "RRGPX": "RR GPX Lateral - 2",
    "RRR": "Road Runner Residue",
    "RRTW": "Road Runner TW",
    "RT": "Rojo Toro",
    "RB": "Rojo Banco",
    "RRDEI": "RR Double E Int",
}
from PIL import Image
from pathlib import Path

class ReportGUI(ctk.CTk):
    """CustomTkinter window for report generation with a modern layout."""

    def __init__(self, master_pilots, clients, generate_callback, *,
                 logo_path=None, initial_folder=None):
        super().__init__()
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.title("Report Generator")
        self.master_pilots = master_pilots
        self.clients = clients
        self.generate_callback = generate_callback
        self.progress_queue = queue.Queue()

        # Window Geometry (1600x900 logical pixels)
        self._scale = self.tk.call('tk', 'scaling')
        physical_w, physical_h = 1600, 900
        logical_w = int(physical_w / self._scale)
        logical_h = int(physical_h / self._scale)
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = int((screen_w - logical_w) / 2)
        y = int((screen_h - logical_h) / 2)
        self.geometry(f"{logical_w}x{logical_h}+{x}+{y}")

        # Main Frame with Split Layout
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True)

        # Header (Logo)
        if logo_path is None:
            logo_file = Path(__file__).parent / "Arch_Aerial_LOGO.jpg"
        else:
            logo_file = Path(logo_path)
        if logo_file.is_file():
            pil_logo = Image.open(logo_file)
            self._logo_aspect = pil_logo.width / pil_logo.height
            initial_w = int(physical_w * 0.25)
            initial_h = int(initial_w / self._logo_aspect)
            self._logo_ctk = ctk.CTkImage(
                light_image=pil_logo,
                size=(int(initial_w / self._scale), int(initial_h / self._scale)),
            )
            self.logo_label = ctk.CTkLabel(self.main_frame, image=self._logo_ctk, text="")
            self.logo_label.pack(pady=(20, 10))

        # Split Layout: Sidebar (Left) and Main Panel (Right)
        self.sidebar_frame = ctk.CTkFrame(self.main_frame, width=int(logical_w * 0.3), corner_radius=10)
        self.sidebar_frame.pack(side="left", fill="y", padx=(20, 10), pady=20)

        self.content_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.content_frame.pack(side="right", fill="both", expand=True, padx=(10, 20), pady=20)

        # Sidebar: Pilot Selection
        pilot_header = ctk.CTkLabel(self.sidebar_frame, text="Select Pilot(s) (max 3)", font=ctk.CTkFont(size=16, weight="bold"))
        pilot_header.pack(pady=(10, 5), padx=10, anchor="w")

        self.pilot_frame = ctk.CTkScrollableFrame(self.sidebar_frame, corner_radius=5)
        self.pilot_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.pilot_vars = []
        for p in master_pilots:
            var = ctk.BooleanVar()
            cb = ctk.CTkCheckBox(self.pilot_frame, text=p, variable=var, command=self._limit_pilots)
            cb.pack(anchor="w", pady=5, padx=10)
            self.pilot_vars.append((var, p))

        # Main Panel: Settings Card
        settings_card = ctk.CTkFrame(self.content_frame, corner_radius=10)
        settings_card.pack(fill="x", padx=20, pady=(20, 10))

        # File Path
        path_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        path_frame.pack(fill="x", padx=20, pady=10)
        path_label = ctk.CTkLabel(path_frame, text="File Path to Photos:", width=150, anchor="w")
        path_label.pack(side="left")
        self.path_var = ctk.StringVar(value=initial_folder or "")
        self.entry = ctk.CTkEntry(path_frame, textvariable=self.path_var, state="readonly")
        self.entry.pack(side="left", fill="x", expand=True, padx=(10, 0))
        browse_btn = ctk.CTkButton(path_frame, text="Browse…", command=self.browse_folder, width=100)
        browse_btn.pack(side="left", padx=(10, 0))

        # Client Selection
        client_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        client_frame.pack(fill="x", padx=20, pady=10)
        client_label = ctk.CTkLabel(client_frame, text="Select Client:", width=150, anchor="w")
        client_label.pack(side="left")
        self.client_var = ctk.StringVar(value=clients[0] if clients else "")
        self.client_menu = ctk.CTkOptionMenu(client_frame, variable=self.client_var, values=clients,
                                             command=self.update_cover_dir)
        self.client_menu.pack(side="left", fill="x", expand=True, padx=(10, 0))

        # Pipeline Selection for TRD
        self.pipeline_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        pipeline_label = ctk.CTkLabel(self.pipeline_frame, text="Select Pipeline:", width=150, anchor="w")
        pipeline_label.pack(side="left")
        self.pipeline_var = ctk.StringVar()
        pipeline_values = [f"{k} - {v}" for k, v in TRD_PIPELINES.items()]
        self.pipeline_menu = ctk.CTkOptionMenu(self.pipeline_frame, variable=self.pipeline_var, values=pipeline_values)
        if pipeline_values:
            self.pipeline_var.set(pipeline_values[0])
        self.pipeline_menu.pack(side="left", fill="x", expand=True, padx=(10, 0))

        self.pipeline_frame.pack_forget()


        # Cover Photo
        cover_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        cover_frame.pack(fill="x", padx=20, pady=10)
        cover_label = ctk.CTkLabel(cover_frame, text="Select Cover Photo:", width=150, anchor="w")
        cover_label.pack(side="left")
        self.cover_var = ctk.StringVar()
        self.cover_button = ctk.CTkButton(cover_frame, text="Browse…", command=self.browse_cover, width=100)
        self.cover_button.pack(side="left", padx=(10, 0))

        # Main Panel: Action Card
        action_card = ctk.CTkFrame(self.content_frame, corner_radius=10)
        action_card.pack(fill="x", padx=20, pady=10)

        self.gen_btn = ctk.CTkButton(action_card, text="Generate Reports", command=self._on_generate, fg_color="green")
        self.gen_btn.pack(pady=10)

        self.progress = ctk.CTkProgressBar(action_card)
        self.progress.pack(fill="x", padx=20, pady=10)
        self.progress.set(0)

        self.status_label = ctk.CTkLabel(action_card, text="")
        self.status_label.pack(pady=10)

        self.view_button = ctk.CTkButton(action_card, text="View Reports", command=self.open_reports, fg_color="blue")
        self.view_button.pack(pady=10)
        self.view_button.pack_forget()

        # Ensure pipeline frame visibility based on initial client
        self.update_cover_dir()

        # Bind resize event for logo
        self.bind("<Configure>", self._on_resize)

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.path_var.set(path)

    def update_cover_dir(self, *_):
        self.cover_var.set("")
        if self.client_var.get() == "TRD":
            self.pipeline_frame.pack(fill="x", padx=20, pady=10)
        else:
            self.pipeline_frame.pack_forget()

    def browse_cover(self):
        client = self.client_var.get()
        start_dir = Path(__file__).parent / "Clients" / client / "Cover Photos - DOCX"
        path = filedialog.askopenfilename(initialdir=start_dir, filetypes=[("DOCX", "*.docx")])
        if path:
            self.cover_var.set(path)

    def _limit_pilots(self):
        selected = [var for var, _ in self.pilot_vars if var.get()]
        if len(selected) > 3:
            for var, _ in self.pilot_vars[::-1]:
                if var.get():
                    var.set(False)
                    break

    def _collect_pilots(self):
        return [name for var, name in self.pilot_vars if var.get()][:3]

    def _on_generate(self):
        folder = self.path_var.get()
        client = self.client_var.get()
        pipeline = None
        if client == "TRD":
            selection = self.pipeline_var.get()
            pipeline = selection.split(" - ")[0] if selection else None
        cover = self.cover_var.get() or None
        pilots = self._collect_pilots()
        if not folder:
            self.status_label.configure(text="Select a folder")
            return
        if not pilots:
            self.status_label.configure(text="Select up to 3 pilots")
            return
        self.progress.set(0)
        self.status_label.configure(text="Backing up photos…")
        self.view_button.pack_forget()
        self.gen_btn.configure(state="disabled")
        self.update()

        thread = threading.Thread(target=self._generate_thread, args=(folder, client, pipeline, cover, pilots))
        thread.start()
        self._check_queue()

    def _generate_thread(self, folder, client, pipeline, cover, pilots):
        self.generate_callback(folder, client, pipeline, cover, pilots, self._queue_progress, self._queue_status)

    def _queue_progress(self, value):
        self.progress_queue.put(("progress", value))

    def _queue_status(self, text=""):
        self.progress_queue.put(("status", text))

    def _check_queue(self):
        try:
            while True:
                item = self.progress_queue.get_nowait()
                if item[0] == "progress":
                    self.progress.set(item[1])
                    if item[1] >= 1.0:
                        self.view_button.pack()
                        self.gen_btn.configure(state="normal")
                elif item[0] == "status":
                    self.status_label.configure(text=item[1])
        except queue.Empty:
            pass
        self.after(100, self._check_queue)

    def _on_resize(self, event):
        if event.widget is not self or not hasattr(self, "_logo_ctk"):
            return
        phys_w = int(self.winfo_width() * self._scale)
        new_phys_w = int(phys_w * 0.25)
        new_phys_h = int(new_phys_w / self._logo_aspect)
        new_log_w = int(new_phys_w / self._scale)
        new_log_h = int(new_phys_h / self._scale)
        if new_log_w < 1 or new_log_h < 1:
            return
        new_ctk = ctk.CTkImage(
            light_image=self._logo_ctk._light_image,
            size=(new_log_w, new_log_h),
        )
        self.logo_label.configure(image=new_ctk)
        self.logo_label.image = new_ctk

    def open_reports(self):
        folder = Path(self.path_var.get()) / OUTPUT_FOLDER_NAME
        path = str(folder)
        try:
            if os.name == 'nt':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception:
            pass
