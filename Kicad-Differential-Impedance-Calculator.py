import math
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import copy 
import re 
import csv 

# --- Dil Sözlükleri (Sadece İngilizce) ---
MESSAGES = {
    # General
    "TITLE": "Kicad-Differential-Impedance-Calculator (Version 1.0)",
    "DESIGNED_BY": "Designed by Muzaffer Nizam",
    "SUCCESS": "Data synchronized successfully!",
    "ERROR_SYNC": "Synchronization Error",
    "ERROR_INPUT": "Input Error",
    "ERROR_CALC": "Calculation Error",
    "ERROR_UNKNOWN": "An unexpected error occurred",
    "PLEASE_SAVE": "Please synchronize data before switching tabs.",
    
    # Tabs
    "TAB_STACKUP": "Stackup Editing (Dynamic)",
    "TAB_CALC": "Geometry & Calculation",
    "TAB_STANDARDS": "Standard Impedances",
    
    # Stackup Controls
    "COPPER_COUNT": "Copper Layer Count:",
    "TOTAL_THICKNESS": "Total PCB Thickness:",
    "MM": "mm",
    "EXPORT_CSV": "Export to CSV",
    "IMPORT_CSV": "Import from CSV",
    "SYNC_SAVE": "Save Stackup Data / Synchronize",
    
    # Stackup Table Headers
    "COL_NAME": "Layer Name",
    "COL_CLASS": "Class",
    "COL_THICKNESS": "Thickness (mm)",
    "COL_DK": "Dk (Er)",

    # Calculation Inputs
    "GROUP_GEOM": "Trace Geometry Parameters (mm)",
    "LABEL_W": "Trace Width (W):",
    "LABEL_GAP": "Gap Between Traces (Gap):",
    "LABEL_S": "Space to Ground (S):",
    "LABEL_Z0_TARGET": "Target Impedance (Z0, Ω):",
    "LABEL_TOLERANCE": "Tolerance (± %):",
    
    # Calculation Mapping
    "GROUP_MAPPING": "Calculation Parameters Mapping",
    "MAP_PARAM": "Formula Parameter",
    "MAP_SOURCE": "Data Source",
    "MAP_H": "Distance to Nearest Plane",
    "MAP_ER": "Effective (Average) Dk",
    
    # Calculation Results
    "CALC_LAYER": "Layer to Calculate:",
    "CALC_BUTTON": "Calculate Impedance",
    "GROUP_RESULTS": "Calculation Results",
    "LABEL_ZDIFF": "Differential Impedance (Zdiff):",
    "OHMS": "Ohms",
    "LABEL_STATUS": "Result:",
    "STATUS_OK": "Impedance PASS",
    "STATUS_FAIL": "Impedance FAIL",
    "LABEL_USED_PARAMS": "Used H, T, Er:",
    "LABEL_MODEL": "Model Selection:",
    "LABEL_CPWG": "CPWG Control:",
    "CPWG_IGNORED": "(Lateral Ground S ignored, since S > H)",
    "CPWG_APPLIED": "(Lateral Ground S EFFECTIVE! CPWG Correction Applied)",
    
    # Standards Table Headers
    "STD_INTERFACE": "Interface",
    "STD_NOMINAL_Z": "Nominal Differential Impedance (Ω)",
    "STD_TOLERANCE": "Typical Tolerance (± %)",
    "STD_NOTES": "Notes",
    
    # Standard Notes Content
    "NOTE_PCIE": "PCI-SIG specification; 100Ω targeted for PCB, 85Ω may be preferred in low-loss designs.",
    "NOTE_ETH": "IEEE 802.3; compatible with twisted pair cables, ±5% tighter tolerance possible.",
    "NOTE_USB2": "USB-IF specification; single-ended 45Ω.",
    "NOTE_USB3": "USB-IF; for SuperSpeed pairs, ±15% common.",
    "NOTE_CAN": "ISO 11898; ±20% for cable, ±10% recommended for PCB; common in automotive.",
    "NOTE_RS485": "EIA/TIA-485; standard for twisted pair cables, matched with termination resistor; industrial networks.",
    "NOTE_RS422": "EIA/TIA-422; generally 120Ω cables are used, 100Ω may be acceptable on PCB; for multi-drop.",
    "NOTE_LVDS": "TIA/EIA-644A; for general high-speed video/signal interfaces (e.g., display connections).",
    "NOTE_HDMI": "HDMI specification; for TMDS (Transition-Minimized Differential Signaling) pairs.",
    "NOTE_ETHCAT": "ETG specification; industrial Ethernet, based on twisted-pair.",
    "NOTE_MIL": "US military standard; avionics data bus, used in DO-254 compliant designs.",
    "NOTE_PROFIBUS": "IEC 61158; industrial automation, RS-485 based.",
}

# --- Sabit Renkler ve Şablonlar (Değişmedi) ---
COLOR_COPPER = "#CF8534"    
COLOR_SOLDER_MASK = "#228B22" 
COLOR_CORE = "#808080"        
COLOR_PREPREG = "#C0C0C0"     
COLOR_HEADER = "#2E8B57"      
COLOR_MAPPING_HEADER = "#2E8B57" 
COLOR_TEXT_LIGHT = "white"
COLOR_TEXT_DARK = "black" # Siyah Metin Rengi
COLOR_STANDARD_HEADER = "#36454F" 
COLOR_SYNC_BUTTON_BG = "#0056b3" 
COLOR_SYNC_BUTTON_FG = "black"   
COLOR_SUCCESS = "green"          

COPPER_TEMPLATE = ["", "Signal", "0.018", "", "Copper"] 
SOLDER_MASK_TOP = ["Top Solder", "Solder Mask", "0.01", "3.5", "Solder Mask"] 
SOLDER_MASK_BOTTOM = ["Bottom Solder", "Solder Mask", "0.01", "3.5", "Solder Mask"]
PREPREG_TEMPLATE = ["Dielectric", "Prepreg", "0.15", "4.1", "Prepreg"]
CORE_TEMPLATE = ["Dielectric", "Core", "0.51", "4.5", "Core"]

# --- Hesaplama Fonksiyonları (Değişmedi) ---
def calculate_wide_traces(W, S_track, T, H, Er):
    Z0_base = (87.0 / math.sqrt(Er + 1.41)) * math.log((5.98 * H) / (0.8 * W + T))
    coupling_factor = 1.0 - 0.48 * math.exp(-0.96 * S_track / H)
    return 2.0 * Z0_base * coupling_factor

def calculate_narrow_traces(W, S_track, T, H, Er):
    W_prime = W + (T / math.pi) * (1.0 + math.log((4.0 * math.pi * W) / T + 1.0))
    u = W_prime / H 
    a = (Er + 1.0) / 2.0
    b = (Er - 1.0) / 2.0
    c = math.pow(1.0 + 12.0 / u, -0.5) 
    Er_eff = a + b * c
    Z0 = (60.0 / math.sqrt(Er_eff)) * math.log(8.0 / u + u / 4.0)
    coupling_factor = (1.0 - 0.48 * math.exp(-0.96 * S_track / H))
    return 2.0 * Z0 * coupling_factor

# --- GUI Sınıfı ---

class ImpedanceCalculatorApp:
    def __init__(self, master):
        self.master = master
        
        # Tek Dil: İngilizce
        self.current_lang = MESSAGES 

        self.layer_options = [2, 4, 6, 8, 10, 12, 14, 16]
        self.num_layers_var = tk.StringVar(value=str(self.layer_options[2])) 
        self.num_layers_var.trace_add('write', self.on_layer_count_change) 

        self.W_var = tk.StringVar(value="0.2")    
        self.Gap_var = tk.StringVar(value="0.2")  
        self.S_var = tk.StringVar(value="1.0")    
        self.target_zdiff_var = tk.StringVar(value="100.0") 
        self.tolerance_percent_var = tk.StringVar(value="10.0") 

        self.stackup_data = [] 
        self.entry_vars = [] 
        self.signal_layers = []
        self.selected_layer = tk.StringVar() 
        
        self.Zdiff_result = tk.StringVar()
        self.tolerance_status_var = tk.StringVar(value="---") 
        self.model_info = tk.StringVar()
        self.cpwg_info = tk.StringVar()
        self.params_used = tk.StringVar()
        self.sync_status_var = tk.StringVar(value="")
        
        self.total_thickness_var = tk.StringVar(value="0.00") 

        self.table_frame = None 
        self.layer_select_combobox = None 
        self.zdiff_result_label = None 
        self.status_label = None 
        self.sync_status_label = None 
        self.notebook = None 

        self.setup_styles()
        self.vcmd_float = (master.register(self.validate_float_input), '%P')
        
        self.create_widgets()
        self.generate_stackup_data(int(self.num_layers_var.get()))
        self.update_main_title()

    # Dil Seçimi ve Yenileme Fonksiyonları kaldırıldı
    def update_main_title(self):
        self.master.title(self.current_lang["TITLE"])
        
    def validate_float_input(self, P):
        if P == "":
            return True
        if re.match(r'^\d*[,\.]?\d*$', P):
            return True
        return False

    # --- Görsel ve Renk Fonksiyonları (Değişmedi) ---
    
    def setup_styles(self):
        style = ttk.Style()
        
        style.configure("Cell.TLabel", padding=[5, 2], borderwidth=1, relief="solid", anchor="center")
        style.configure("Header.TLabel", background=COLOR_HEADER, foreground=COLOR_TEXT_LIGHT, font=('Arial', 9, 'bold'), borderwidth=1, relief="solid", anchor="center")
        style.configure("Copper.TLabel", background=COLOR_COPPER, foreground=COLOR_TEXT_DARK)
        style.configure("Solder.TLabel", background=COLOR_SOLDER_MASK, foreground=COLOR_TEXT_LIGHT)
        style.configure("Core.TLabel", background=COLOR_CORE, foreground=COLOR_TEXT_LIGHT)
        style.configure("Dielectric.TLabel", background=COLOR_PREPREG, foreground=COLOR_TEXT_DARK)
        style.configure("Total.TLabel", font=('Arial', 11, 'bold'), foreground='darkred') 
        style.configure("MappingHeader.TLabel", 
                        background=COLOR_MAPPING_HEADER, 
                        foreground=COLOR_TEXT_LIGHT, 
                        font=('Arial', 9, 'bold'), 
                        borderwidth=1, 
                        relief="solid", 
                        anchor="center")
        style.configure("Mapping.TLabel", background="white", borderwidth=1, relief="solid", anchor="w", padding=[5, 2], foreground='black')
        style.configure("Designer.TLabel", font=('Arial', 8, 'italic'), foreground='gray', anchor="e")
        style.configure("StandardHeader.TLabel", 
                        background=COLOR_STANDARD_HEADER, 
                        foreground=COLOR_TEXT_LIGHT, 
                        font=('Arial', 10, 'bold'), 
                        borderwidth=1, 
                        relief="solid", 
                        anchor="center",
                        padding=[5, 5])
        style.configure("StandardCell.TLabel", 
                        background="white", 
                        foreground="black", 
                        font=('Arial', 9), 
                        borderwidth=1, 
                        relief="solid", 
                        anchor="w",
                        padding=[5, 2])
        style.configure("Status.TLabel", font=('Arial', 12, 'bold'))
        
        style.configure("Sync.TButton", 
                        background=COLOR_SYNC_BUTTON_BG, 
                        foreground=COLOR_SYNC_BUTTON_FG,
                        font=('Arial', 9, 'bold'))
        style.map("Sync.TButton", 
                  background=[('active', '#004a99')], 
                  foreground=[('active', COLOR_SYNC_BUTTON_FG)])
        
        style.configure("SyncStatus.TLabel", 
                        foreground=COLOR_SUCCESS, 
                        font=('Arial', 9, 'bold'),
                        anchor="e")


    def get_color_style_by_type(self, layer_class, layer_name):
        if layer_class == "Solder Mask":
            return "Solder.TLabel"
        elif layer_class == "Copper":
            return "Copper.TLabel"
        elif layer_class == "Core":
            return "Core.TLabel"
        elif layer_class in ["Prepreg", "Dielectric"]:
            return "Dielectric.TLabel"
        return "Cell.TLabel"

    # --- Dinamik Stackup Oluşturma Mantığı (Değişmedi) ---

    def generate_stackup_data(self, num_copper_layers):
        
        copper_names = []
        for i in range(1, num_copper_layers + 1):
             if i == 1:
                 copper_names.append("1. Top Layer")
             elif i == num_copper_layers:
                 copper_names.append(f"{num_copper_layers}. Bottom Layer")
             else:
                 copper_names.append(f"{i}. Inner Layer {i-1}")
        
        new_stackup = []
        new_stackup.append(copy.deepcopy(SOLDER_MASK_TOP))
        
        for i, name in enumerate(copper_names):
            copper_layer = copy.deepcopy(COPPER_TEMPLATE)
            copper_layer[0] = name # Name
            
            layer_class = "Signal" 
            
            # --- KATMAN SINIFLANDIRMA MANTIĞI ---
            
            if num_copper_layers == 2:
                if i == 0: layer_class = "Signal"
                elif i == 1: layer_class = "Plane"
            
            elif num_copper_layers == 4:
                # L1=Signal, L2=Plane, L3=Plane, L4=Signal
                if i == 0: layer_class = "Signal"
                elif i == 1: layer_class = "Plane"
                elif i == 2: layer_class = "Plane"
                elif i == 3: layer_class = "Signal"
            
            elif num_copper_layers == 6:
                # L1=Signal, L2=Plane, L3=Signal, L4=Signal, L5=Plane, L6=Signal
                if i == 0: layer_class = "Signal"
                elif i == 1: layer_class = "Plane"
                elif i == 2: layer_class = "Signal"
                elif i == 3: layer_class = "Signal"
                elif i == 4: layer_class = "Plane"
                elif i == 5: layer_class = "Signal"
                
            elif num_copper_layers >= 8:
                # Signal, Plane, Signal, Plane, ...
                if (i + 1) % 2 != 0:
                    layer_class = "Signal"
                else:
                    layer_class = "Plane"
            
            copper_layer[1] = layer_class
            
            new_stackup.append(copper_layer)
            
            if i < num_copper_layers - 1:
                dielectric_layer = copy.deepcopy(CORE_TEMPLATE) if i == num_copper_layers//2 - 1 else copy.deepcopy(PREPREG_TEMPLATE)
                dielectric_layer[0] = f"Dielectric {i+2}"
                new_stackup.append(dielectric_layer)

        if num_copper_layers > 0:
            new_stackup.append(copy.deepcopy(SOLDER_MASK_BOTTOM))

        self.stackup_data = new_stackup
        
        self.signal_layers = [item[0] for item in self.stackup_data if item[1] == "Signal"] 
        if self.signal_layers:
             self.selected_layer.set(self.signal_layers[0])

        self.redraw_stackup_table()

    def on_layer_count_change(self, *args):
        try:
            num = int(self.num_layers_var.get())
            if num in self.layer_options:
                self.generate_stackup_data(num)
        except ValueError:
            pass 

    # --- METOT: Stackup Verilerini StringVar'lardan Ana Listeye Aktar ---
    def update_stackup_data(self):
        try:
            for i, row_vars in enumerate(self.entry_vars):
                
                self.stackup_data[i][0] = row_vars[0].get().strip()
                self.stackup_data[i][1] = row_vars[1].get().strip()
                self.stackup_data[i][2] = row_vars[2].get().strip()
                self.stackup_data[i][3] = row_vars[3].get().strip()
                
            self.signal_layers = [item[0] for item in self.stackup_data if item[1] == "Signal"] 
            if self.layer_select_combobox:
                self.layer_select_combobox['values'] = self.signal_layers
                if self.selected_layer.get() not in self.signal_layers and self.signal_layers:
                    self.selected_layer.set(self.signal_layers[0])
                elif not self.signal_layers:
                    self.selected_layer.set("")
            
            self.update_total_thickness()
            
            self.sync_status_var.set(self.current_lang["SUCCESS"])
            if self.sync_status_label:
                self.sync_status_label.config(foreground=COLOR_SUCCESS)
                self.master.after(2000, lambda: self.sync_status_var.set(""))

        except Exception as e:
            messagebox.showerror(self.current_lang["ERROR_SYNC"], f"{self.current_lang['ERROR_UNKNOWN']}: {e}")
            self.sync_status_var.set(self.current_lang["ERROR_SYNC"])
            if self.sync_status_label:
                self.sync_status_label.config(foreground="red")
            self.master.after(3000, lambda: self.sync_status_var.set(""))


    # --- Toplam Kalınlık Güncelleme Fonksiyonu (Değişmedi) ---

    def update_total_thickness(self, *args):
        total = 0.0
        for row_vars in self.entry_vars:
            thickness_str = row_vars[2].get()
            
            if thickness_str:
                try:
                    # Virgüllü veya noktalı girişi float'a çevir
                    value = float(thickness_str.replace(',', '.'))
                    total += value
                except ValueError:
                    pass
        
        self.total_thickness_var.set(f"{total:.3f}")

    # --- CSV İşlemleri ---

    def export_to_csv(self):
        self.update_stackup_data() 
        
        try:
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title=self.current_lang["EXPORT_CSV"]
            )
            
            if filepath:
                headers = ["Layer Number", self.current_lang["COL_NAME"], self.current_lang["COL_CLASS"], self.current_lang["COL_THICKNESS"], self.current_lang["COL_DK"]]
                
                with open(filepath, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=';') 
                    writer.writerow(headers)
                    
                    for i, row_data in enumerate(self.stackup_data):
                        row_list = [
                            f"{i+1:02d}", 
                            row_data[0],  
                            row_data[1],  
                            row_data[2].replace('.', ','),  
                            row_data[3].replace('.', ','),   
                            row_data[4] 
                        ]
                        writer.writerow(row_list)
                        
                    writer.writerow(["", self.current_lang["TOTAL_THICKNESS"], "", self.total_thickness_var.get(), self.current_lang["MM"]])

                messagebox.showinfo(self.current_lang["SUCCESS"], f"{self.current_lang['SUCCESS']} (CSV: {filepath})")
            
        except Exception as e:
            messagebox.showerror(self.current_lang["ERROR_SYNC"], f"{self.current_lang['ERROR_UNKNOWN']}: {e}")

    def clean_dk_value(self, value):
        value = str(value).strip()
        if not value:
            return ""

        cleaned = re.sub(r'[a-zA-Z\s]+', '', value) 
        
        if cleaned:
            return cleaned.replace('.', ',') 
        
        return ""


    def import_from_csv(self):
        try:
            filepath = filedialog.askopenfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title=self.current_lang["IMPORT_CSV"]
            )
            
            if not filepath:
                return

            data = []
            with open(filepath, 'r', newline='', encoding='utf-8') as f:
                
                try:
                    reader = csv.reader(f, delimiter=';')
                except:
                    f.seek(0)
                    reader = csv.reader(f, delimiter=',') 

                headers = next(reader) 
                for row in reader:
                    if len(row) > 3 and self.current_lang["TOTAL_THICKNESS"] not in row[1]: 
                        data.append(row)
            
            if not data or len(data) != len(self.stackup_data):
                 messagebox.showerror(self.current_lang["ERROR_INPUT"], f"Stackup size mismatch. File row count ({len(data)}) does not match existing stackup ({len(self.stackup_data)}).")
                 return

            for i, row in enumerate(data):
                
                name_new = row[1].strip() if len(row) > 1 else self.stackup_data[i][0]
                class_new = row[2].strip() if len(row) > 2 else self.stackup_data[i][1]
                thickness_new = row[3].strip() if len(row) > 3 else self.stackup_data[i][2]
                dk_raw = row[4].strip() if len(row) > 4 else self.stackup_data[i][3]
                
                dk_new = self.clean_dk_value(dk_raw)

                self.entry_vars[i][0].set(name_new) 
                self.entry_vars[i][1].set(class_new) 
                self.entry_vars[i][2].set(thickness_new)
                if self.stackup_data[i][4] in ["Prepreg", "Core", "Solder Mask"]:
                    self.entry_vars[i][3].set(dk_new) 
                
                self.stackup_data[i][0] = name_new
                self.stackup_data[i][1] = class_new
                self.stackup_data[i][2] = thickness_new
                self.stackup_data[i][3] = dk_new
            
            self.redraw_stackup_table() 
            self.update_stackup_data() 

            messagebox.showinfo(self.current_lang["SUCCESS"], self.current_lang["SUCCESS"])

        except Exception as e:
            messagebox.showerror(self.current_lang["ERROR_SYNC"], f"{self.current_lang['ERROR_UNKNOWN']}: {e}")


    # --- GUI Yaratma Metotları ---
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Üst kontrol kısmı (Dil seçeneği kaldırıldı)
        
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(pady=5, padx=5, fill="both", expand=True)

        # 1. Sekme: Stackup Düzenleme
        stackup_frame = ttk.Frame(self.notebook)
        self.notebook.add(stackup_frame, text=self.current_lang["TAB_STACKUP"])
        
        self.setup_stackup_controls(stackup_frame)
        
        self.table_frame = ttk.Frame(stackup_frame) 
        self.table_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Kaydet Buton Çerçevesi (Sağ alta hizalanacak)
        save_button_frame = ttk.Frame(stackup_frame)
        save_button_frame.pack(fill='x', pady=5, padx=5)
        
        button_and_status_container = ttk.Frame(save_button_frame)
        button_and_status_container.pack(side='right')

        # 1. Durum etiketi (Butonun üstünde)
        self.sync_status_label = ttk.Label(button_and_status_container, 
                                           textvariable=self.sync_status_var, 
                                           style="SyncStatus.TLabel",
                                           anchor="e") 
        self.sync_status_label.pack(side='top', fill='x', pady=(0, 2))

        # 2. Kaydet butonu (Butonun altında)
        ttk.Button(button_and_status_container, 
                   text=self.current_lang["SYNC_SAVE"], 
                   command=self.update_stackup_data, 
                   style="Sync.TButton").pack(side='top')
        
        # 2. Sekme: Geometri & Hesaplama
        calc_frame = ttk.Frame(self.notebook)
        self.notebook.add(calc_frame, text=self.current_lang["TAB_CALC"])
        self.setup_calculation_tab(calc_frame)
        
        # 3. Sekme: Standart Empedanslar
        standards_frame = ttk.Frame(self.notebook)
        self.notebook.add(standards_frame, text=self.current_lang["TAB_STANDARDS"])
        self.setup_standards_tab(standards_frame)
        
        # Tasarımcı Bilgisi (En alta eklenmiştir)
        designer_label = ttk.Label(main_frame, text=self.current_lang["DESIGNED_BY"], style="Designer.TLabel")
        designer_label.pack(fill='x', padx=5, pady=(0, 5))


    def setup_stackup_controls(self, frame):
        # Bu frame sadece katman sayısı ve Total Thickness/CSV kontrollerini içerir
        control_frame = ttk.Frame(frame)
        control_frame.pack(pady=10, fill='x', padx=5)
        
        left_control_frame = ttk.Frame(control_frame)
        left_control_frame.pack(side='left', padx=5)
        
        center_control_frame = ttk.Frame(control_frame)
        center_control_frame.pack(side='left', padx=15)
        
        right_control_frame = ttk.Frame(control_frame)
        right_control_frame.pack(side='right', padx=5)

        ttk.Label(left_control_frame, text=self.current_lang["COPPER_COUNT"]).pack(side='left', padx=5)
        ttk.Combobox(left_control_frame, 
                     textvariable=self.num_layers_var, 
                     values=self.layer_options, 
                     state="readonly", 
                     width=5).pack(side='left', padx=5)
        
        # CSV Butonları
        ttk.Button(center_control_frame, text=self.current_lang["EXPORT_CSV"], command=self.export_to_csv).pack(side='left', padx=5)
        ttk.Button(center_control_frame, text=self.current_lang["IMPORT_CSV"], command=self.import_from_csv).pack(side='left', padx=5)
                     
        # Sağ Taraf: Toplam Kalınlık Gösterimi
        ttk.Label(right_control_frame, text=self.current_lang["TOTAL_THICKNESS"]).pack(side='left')
        
        ttk.Label(right_control_frame, textvariable=self.total_thickness_var, style="Total.TLabel").pack(side='left', padx=(5, 0))
        ttk.Label(right_control_frame, text=self.current_lang["MM"], style="Total.TLabel").pack(side='left')


    def redraw_stackup_table(self):
        """Mevcut tabloyu siler ve yeni stackup verileriyle yeniden çizer."""
        if not self.table_frame:
            return

        for widget in self.table_frame.winfo_children():
            widget.destroy()
        
        self.entry_vars = []
        
        # Sütun Başlıkları
        headers = ["#", self.current_lang["COL_NAME"], self.current_lang["COL_CLASS"], self.current_lang["COL_THICKNESS"], self.current_lang["COL_DK"]]
        column_widths = [30, 150, 100, 100, 100]

        for col, header in enumerate(headers):
            ttk.Label(self.table_frame, text=header, style="Header.TLabel", width=column_widths[col]//10).grid(
                row=0, column=col, sticky="nsew"
            )
            self.table_frame.grid_columnconfigure(col, weight=1, minsize=column_widths[col])

        # Giriş Satırları
        for i, row_data in enumerate(self.stackup_data):
            # row_data: [Name (0), Class (1), Thickness (2), Dk (3), Layer_Type (4)]
            layer_name = row_data[0]
            layer_class = row_data[1] 
            layer_type = row_data[4] 
            style_name = self.get_color_style_by_type(layer_type, layer_name)
            
            # --- 0. Layer Number ---
            cell_0 = ttk.Label(self.table_frame, text=f"{i+1:02d}", style=style_name, anchor="center")
            cell_0.grid(row=i+1, column=0, sticky="nsew", padx=0, pady=0)
            
            row_vars = []
            
            # --- 1. Katman Adı (Düzenlenebilir Entry) ---
            var_name = tk.StringVar(value=layer_name)
            entry_name = ttk.Entry(self.table_frame, textvariable=var_name, width=column_widths[1]//10)
            entry_name.grid(row=i+1, column=1, sticky="nsew", padx=1, pady=1)
            
            # Copper Layer Name rengi her zaman siyah
            if layer_type == "Copper":
                 entry_name.config(background=COLOR_COPPER, foreground=COLOR_TEXT_DARK)
            else:
                 entry_name.config(background='white', foreground='black')

            row_vars.append(var_name)

            # --- 2. Sınıf (Combobox/Entry) ---
            if layer_type == "Copper":
                var_class = tk.StringVar(value=layer_class)
                combo_values = ["Signal", "Plane"] 
                
                combo_class = ttk.Combobox(self.table_frame, 
                                          textvariable=var_class, 
                                          values=combo_values, 
                                          state="readonly", 
                                          width=column_widths[2]//10)
                combo_class.grid(row=i+1, column=2, sticky="nsew", padx=1, pady=1)
                
                if layer_type == "Copper":
                    combo_class.config(background=COLOR_COPPER, foreground=COLOR_TEXT_DARK)
                
                row_vars.append(var_class)
            elif layer_type in ["Prepreg", "Core"]:
                var_class = tk.StringVar(value=layer_class)
                combo_class = ttk.Combobox(self.table_frame, 
                                          textvariable=var_class, 
                                          values=["Core", "Prepreg"], 
                                          state="readonly", 
                                          width=column_widths[2]//10)
                combo_class.grid(row=i+1, column=2, sticky="nsew", padx=1, pady=1)
                row_vars.append(var_class)
            else:
                cell_2 = ttk.Label(self.table_frame, text=layer_class, style=style_name, anchor="w")
                cell_2.grid(row=i+1, column=2, sticky="nsew", padx=0, pady=0)
                row_vars.append(tk.StringVar(value=layer_class))

            # --- 3. Thickness (Sadece Float/Virgüllü Sayı) ---
            var_thickness = tk.StringVar(value=row_data[2])
            entry_thickness = ttk.Entry(self.table_frame, textvariable=var_thickness, width=column_widths[3]//10, 
                                        validate='key', validatecommand=self.vcmd_float)
            entry_thickness.grid(row=i+1, column=3, sticky="nsew", padx=1, pady=1)
            
            if layer_type == "Copper":
                 entry_thickness.config(background=COLOR_COPPER, foreground=COLOR_TEXT_DARK)
            else:
                 entry_thickness.config(background='white', foreground='black') 
                 
            row_vars.append(var_thickness)
            
            var_thickness.trace_add('write', self.update_total_thickness) 
            
            # --- 4. Dk (Er) (Sadece Dielektrikler için Float/Virgüllü Sayı) ---
            if layer_type in ["Prepreg", "Core", "Solder Mask"]:
                var_dk = tk.StringVar(value=row_data[3])
                entry_dk = ttk.Entry(self.table_frame, textvariable=var_dk, width=column_widths[4]//10,
                                      validate='key', validatecommand=self.vcmd_float)
                entry_dk.grid(row=i+1, column=4, sticky="nsew", padx=1, pady=1)
                
                # *** DÜZELTME: Metin rengi, arka plan ne olursa olsun siyah yapılmalı. ***
                fg_color = COLOR_TEXT_DARK
                bg_color = COLOR_SOLDER_MASK if layer_type == "Solder Mask" else (COLOR_CORE if layer_type == "Core" else COLOR_PREPREG)
                
                entry_dk.config(background=bg_color, foreground=fg_color)
                
                row_vars.append(var_dk)
            else:
                cell_4 = ttk.Label(self.table_frame, text="", background=COLOR_COPPER)
                cell_4.grid(row=i+1, column=4, sticky="nsew", padx=0, pady=0)
                row_vars.append(tk.StringVar(value="")) 
            
            self.entry_vars.append(row_vars)

        self.update_total_thickness() 
        
        signal_layer_names = [item[0] for item in self.stackup_data if item[1] == "Signal"]
        
        if self.layer_select_combobox:
            self.layer_select_combobox['values'] = signal_layer_names
            if self.selected_layer.get() not in signal_layer_names and signal_layer_names:
                self.selected_layer.set(signal_layer_names[0])
            elif not signal_layer_names:
                self.selected_layer.set("")
                

    def setup_calculation_tab(self, frame):
        calc_content_frame = ttk.Frame(frame, padding="10")
        calc_content_frame.pack(fill="both", expand=True)

        top_row_frame = ttk.Frame(calc_content_frame)
        top_row_frame.pack(fill='x', pady=5)

        # Geometri Girişleri
        geom_input_frame = ttk.LabelFrame(top_row_frame, text=self.current_lang["GROUP_GEOM"])
        geom_input_frame.pack(side='left', padx=10, fill='y')
        
        geom_vars = [
            (self.current_lang["LABEL_W"], self.W_var),
            (self.current_lang["LABEL_GAP"], self.Gap_var),
            (self.current_lang["LABEL_S"], self.S_var),
            (self.current_lang["LABEL_Z0_TARGET"], self.target_zdiff_var),
            (self.current_lang["LABEL_TOLERANCE"], self.tolerance_percent_var)      
        ]
        
        for i, (label_text, var) in enumerate(geom_vars):
            ttk.Label(geom_input_frame, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky="w")
            ttk.Entry(geom_input_frame, textvariable=var, width=15, validate='key', validatecommand=self.vcmd_float).grid(row=i, column=1, padx=5, pady=5)
            
        # Parametre Haritalama
        mapping_frame = ttk.LabelFrame(top_row_frame, text=self.current_lang["GROUP_MAPPING"])
        mapping_frame.pack(side='left', padx=10, fill='both', expand=True)
        
        # Haritalama Tablosu İçeriği için Dinamik Metin Oluşturma
        map_param_w = self.current_lang["LABEL_W"].split(':')[0]
        map_param_gap = self.current_lang["LABEL_GAP"].split(':')[0]
        map_param_t = self.current_lang["COL_THICKNESS"].split(' ')[0]
        map_param_dk = self.current_lang["COL_DK"].split(' ')[0]
        map_param_s = self.current_lang["LABEL_S"].split(':')[0]
        
        map_source_input = self.current_lang["ERROR_INPUT"].split(' ')[0]
        map_source_thickness = self.current_lang["COL_THICKNESS"]
        map_source_map_h = self.current_lang["MAP_H"]
        map_source_map_er = self.current_lang["MAP_ER"]
        map_source_cpwg = self.current_lang["CPWG_IGNORED"].split('(')[1].split(')')[0]
        
        mapping_data = [
            (map_param_w, f"{map_param_w}_var ({map_param_w} ({map_source_input}))"),
            (map_param_gap, f"{map_param_gap}_var ({map_param_gap} ({map_source_input}))"),
            (f"{map_param_t} (T)", f"{map_source_thickness} (Stackup)"),
            (f"{map_param_t} (H)", map_source_map_h),
            (f"{map_param_dk} (Er)", map_source_map_er),
            (f"{map_param_s}", f"{map_param_s}_var ({map_source_cpwg})")
        ]

        ttk.Label(mapping_frame, text=self.current_lang["MAP_PARAM"], style="MappingHeader.TLabel").grid(row=0, column=0, sticky="nsew")
        ttk.Label(mapping_frame, text=self.current_lang["MAP_SOURCE"], style="MappingHeader.TLabel").grid(row=0, column=1, sticky="nsew")

        for i, (param, source) in enumerate(mapping_data):
            ttk.Label(mapping_frame, text=param, style="Mapping.TLabel").grid(row=i+1, column=0, sticky="nsew", padx=1, pady=1)
            ttk.Label(mapping_frame, text=source, style="Mapping.TLabel").grid(row=i+1, column=1, sticky="nsew", padx=1, pady=1)
        
        mapping_frame.grid_columnconfigure(0, weight=1)
        mapping_frame.grid_columnconfigure(1, weight=1)

        control_frame = ttk.Frame(calc_content_frame)
        control_frame.pack(pady=10, fill='x', padx=10)

        # Hesaplama Kontrolleri
        ttk.Label(control_frame, text=self.current_lang["CALC_LAYER"]).pack(side='left', padx=5)
        
        self.layer_select_combobox = ttk.Combobox(control_frame, 
                                                 textvariable=self.selected_layer, 
                                                 values=self.signal_layers, 
                                                 state="readonly", 
                                                 width=20)
        self.layer_select_combobox.pack(side='left', padx=5)

        ttk.Button(control_frame, text=self.current_lang["CALC_BUTTON"], command=self.calculate_impedance).pack(side='right', padx=5)

        # Sonuç Alanı
        result_frame = ttk.LabelFrame(calc_content_frame, text=self.current_lang["GROUP_RESULTS"])
        result_frame.pack(padx=10, pady=5, fill="x")
        
        ttk.Label(result_frame, text=self.current_lang["LABEL_ZDIFF"], font=('Arial', 10)).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.zdiff_result_label = ttk.Label(result_frame, textvariable=self.Zdiff_result, font=('Arial', 16, 'bold'), foreground='blue')
        self.zdiff_result_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(result_frame, text=self.current_lang["OHMS"], font=('Arial', 10)).grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(result_frame, text=self.current_lang["LABEL_STATUS"], font=('Arial', 10)).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        self.status_label = ttk.Label(result_frame, textvariable=self.tolerance_status_var, style="Status.TLabel")
        self.status_label.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Separator(result_frame, orient='horizontal').grid(row=2, column=0, columnspan=3, sticky="ew", pady=5)
        
        ttk.Label(result_frame, text=self.current_lang["LABEL_USED_PARAMS"]).grid(row=3, column=0, padx=5, pady=5, sticky="w")
        ttk.Label(result_frame, textvariable=self.params_used).grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        ttk.Label(result_frame, text=self.current_lang["LABEL_MODEL"]).grid(row=4, column=0, padx=5, pady=5, sticky="w")
        ttk.Label(result_frame, textvariable=self.model_info, wraplength=200).grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        ttk.Label(result_frame, text=self.current_lang["LABEL_CPWG"]).grid(row=5, column=0, padx=5, pady=5, sticky="w")
        ttk.Label(result_frame, textvariable=self.cpwg_info, wraplength=200).grid(row=5, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
    def setup_standards_tab(self, frame):
        """Üçüncü sekme için standart empedans tablosunu oluşturur."""
        
        # Standart Veriler MESSAGES'tan çekiliyor
        standards_data = [
            (self.current_lang["STD_INTERFACE"], self.current_lang["STD_NOMINAL_Z"], self.current_lang["STD_TOLERANCE"], self.current_lang["STD_NOTES"]),
            ("PCI Express (PCIe, Gen1-6)", "100 (or 85 for Gen2+)", "10", self.current_lang["NOTE_PCIE"]),
            ("Ethernet (1000BASE-T)", "100", "10", self.current_lang["NOTE_ETH"]),
            ("USB 2.0", "90", "15", self.current_lang["NOTE_USB2"]),
            ("USB 3.0 / 3.x", "90", "10-15", self.current_lang["NOTE_USB3"]),
            ("CAN Bus", "120", "10-20", self.current_lang["NOTE_CAN"]),
            ("RS-485", "120", "20", self.current_lang["NOTE_RS485"]),
            ("RS-422", "100-120", "20", self.current_lang["NOTE_RS422"]),
            ("LVDS", "100", "10", self.current_lang["NOTE_LVDS"]),
            ("HDMI", "100", "10", self.current_lang["NOTE_HDMI"]),
            ("EtherCAT", "100", "10", self.current_lang["NOTE_ETHCAT"]),
            ("MIL-STD-1553", "78", "10", self.current_lang["NOTE_MIL"]),
            ("Profibus DP", "150", "10-20", self.current_lang["NOTE_PROFIBUS"]),
        ]
        
        container = ttk.Frame(frame, padding="10")
        container.pack(fill="both", expand=True)

        for i, row_data in enumerate(standards_data):
            for j, cell_text in enumerate(row_data):
                if i == 0:
                    label = ttk.Label(container, text=cell_text, style="StandardHeader.TLabel", wraplength=200 if j == 3 else 100)
                    label.grid(row=i, column=j, sticky="nsew", padx=1, pady=1)
                else:
                    label = ttk.Label(container, text=cell_text, style="StandardCell.TLabel", wraplength=250 if j == 3 else 100)
                    label.grid(row=i, column=j, sticky="nsew", padx=1, pady=1)
                    
            container.grid_columnconfigure(j, weight=3 if j == 3 else 1)


    def get_float_or_error(self, var_name, var_value, can_be_zero=False):
        """String değeri float'a çevirir ve sıfırdan küçük/eşit olma durumunu kontrol eder."""
        try:
            # Gerekirse virgülü noktaya çevir (Hem 0.15 hem 0,15 kabul edilir)
            value = float(str(var_value).replace(',', '.')) 
            if not can_be_zero and value <= 0:
                 raise ValueError(f"'{var_name}' must be greater than zero.")
            return value
        except ValueError:
            raise ValueError(f"Please enter a valid numeric value for '{var_name}' (e.g., 0.15 or 0,15).")

    # YARDIMCI FONKSİYON: Bitişik Plane'i ve Dielektrikleri Bulma (KESİN KURAL SETİ)
    def find_nearest_plane_and_dielectric(self, signal_index):
        
        stackup_len = len(self.stackup_data)
        
        def get_dielectric_properties_to_plane(start_index, end_index, step):
            """Belirtilen indeksten başlayıp Plane'e kadar olan dielektrik özelliklerini hesaplar."""
            total_H = 0.0
            weighted_H_sum = 0.0
            
            # Start/end/step değerlerini Python'ın range() fonksiyonu ile uyumlu hale getir
            layer_indices = range(start_index, end_index + step, step) # step 1 veya -1 olabilir
            
            for i in layer_indices:
                item = self.stackup_data[i]
                
                if item[4] in ["Prepreg", "Core", "Solder Mask"]:
                    try:
                        thickness = self.get_float_or_error("", self.entry_vars[i][2].get(), can_be_zero=True)
                        er = self.get_float_or_error("", self.entry_vars[i][3].get(), can_be_zero=True)
                        
                        if thickness <= 0 or er <= 0:
                            raise ValueError 
                            
                        total_H += thickness
                        weighted_H_sum += thickness * er
                    except ValueError:
                        raise Exception(f"{self.current_lang['COL_THICKNESS']}/{self.current_lang['COL_DK']} value for layer ({item[0]}) is invalid or zero/negative.")
                elif item[1] == "Signal":
                     # Hata mesajı İngilizce
                     raise Exception(f"Another signal layer ({item[0]}) found between signal layer ({self.stackup_data[signal_index][0]}) and the Plane. Stackup error!")
            
            if total_H == 0:
                return 0.0, 1.0 
                
            final_Er = weighted_H_sum / total_H
            return total_H, final_Er


        is_top_layer = (self.stackup_data[signal_index][0] == "1. Top Layer")
        is_bottom_layer = (self.stackup_data[signal_index][0].endswith("Bottom Layer"))

        
        # --- ARAMA YÖNLERİ ---
        
        # 1. AŞAĞI YÖNDE ARAMA (Lower Plane)
        down_plane_index = -1
        current_index = signal_index + 1
        while current_index < stackup_len:
            if self.stackup_data[current_index][1] == "Plane":
                down_plane_index = current_index
                break
            if self.stackup_data[current_index][4] == "Copper" and self.stackup_data[current_index][1] != "Plane":
                 # Aşağıdaki katman bakır ama Plane değilse dur
                 break 
            current_index += 1
        
        H_down, Er_down, ref_down = (float('inf'), 0.0, None)

        if down_plane_index != -1:
            try:
                H_down, Er_down = get_dielectric_properties_to_plane(signal_index + 1, down_plane_index, 1) # Dielektrikleri al
                ref_down = self.stackup_data[down_plane_index][0]
            except Exception as e:
                raise e


        # 2. YUKARI YÖNDE ARAMA (Upper Plane)
        up_plane_index = -1
        current_index = signal_index - 1
        while current_index >= 0:
            if self.stackup_data[current_index][1] == "Plane":
                up_plane_index = current_index
                break
            if self.stackup_data[current_index][4] == "Copper" and self.stackup_data[current_index][1] != "Plane":
                 # Yukarıdaki katman bakır ama Plane değilse dur
                 break
            current_index -= 1
        
        H_up, Er_up, ref_up = (float('inf'), 0.0, None)

        if up_plane_index != -1:
            try:
                H_up, Er_up = get_dielectric_properties_to_plane(signal_index - 1, up_plane_index, -1) # Dielektrikleri al
                ref_up = self.stackup_data[up_plane_index][0]
            except Exception as e:
                raise e
        
        
        # --- KARAR VERME: HANGİ PLANE KULLANILACAK? ---
        
        if is_top_layer:
            if down_plane_index != -1 and H_down > 0:
                return H_down, Er_down, ref_down
            else:
                raise Exception("No 'Plane' layer found immediately below the Top Layer. Stackup error!")
                
        elif is_bottom_layer:
            if up_plane_index != -1 and H_up > 0:
                return H_up, Er_up, ref_up
            else:
                raise Exception("No 'Plane' layer found immediately above the Bottom Layer. Stackup error!")

        else:
            down_available = H_down != float('inf') and H_down > 0
            up_available = H_up != float('inf') and H_up > 0

            if not down_available and not up_available:
                raise Exception(f"{self.stackup_data[signal_index][0]} (Inner Layer): No 'Plane' layer found above or below. Stackup error!")
                
            elif down_available and not up_available:
                return H_down, Er_down, ref_down
                
            elif not down_available and up_available:
                return H_up, Er_up, ref_up
            
            elif down_available and up_available:
                # İki Plane de varsa, en yakını seçilir
                if H_down <= H_up:
                    return H_down, Er_down, ref_down
                else:
                    return H_up, Er_up, ref_up
        
        raise Exception("Unexpected error occurred during reference plane selection.")


    def calculate_impedance(self):
        if self.zdiff_result_label:
            self.zdiff_result_label.config(foreground="darkred")
        if self.status_label:
            self.status_label.config(foreground="darkred")
            self.tolerance_status_var.set(self.current_lang["ERROR_CALC"])
            
        try:
            self.update_stackup_data() 
            
            selected_layer_name = self.selected_layer.get()
            layer_index = -1
            
            for i, item in enumerate(self.stackup_data):
                if item[0] == selected_layer_name: 
                    layer_index = i
                    break
            
            if layer_index == -1 or selected_layer_name == "":
                raise Exception("Please select a valid Signal Layer for calculation.")
                
            if self.stackup_data[layer_index][1] != "Signal": 
                raise Exception(f"{self.current_lang['ERROR_INPUT']}: Layer Class ('{self.stackup_data[layer_index][1]}') must be 'Signal'.")

            T_str = self.entry_vars[layer_index][2].get()
            T = self.get_float_or_error(f"{selected_layer_name} {self.current_lang['COL_THICKNESS']} (T)", T_str)

            H, Er, reference_plane_name = self.find_nearest_plane_and_dielectric(layer_index)
            
            if H <= 0 or Er <= 0:
                # Hata mesajları İngilizce'ye çevrildi
                raise Exception(f"{self.current_lang['ERROR_INPUT']}: {self.current_lang['COL_THICKNESS']} (H={H:.3f}) or Dk (Er={Er:.2f}) is zero or negative. Check stackup parameters.")
            
            W_input = self.get_float_or_error(self.current_lang["LABEL_W"].split(':')[0], self.W_var.get())
            Gap_input = self.get_float_or_error(self.current_lang["LABEL_GAP"].split(':')[0], self.Gap_var.get())
            S_input = self.get_float_or_error(self.current_lang["LABEL_S"].split(':')[0], self.S_var.get())

            Target_Z0 = self.get_float_or_error(self.current_lang["LABEL_Z0_TARGET"].split(':')[0], self.target_zdiff_var.get())
            
            Tolerance_P_str = self.tolerance_percent_var.get()
            Tolerance_P = self.get_float_or_error(self.current_lang["LABEL_TOLERANCE"].split(':')[0], Tolerance_P_str, can_be_zero=True)
            if Tolerance_P < 0:
                 raise ValueError(self.current_lang["LABEL_TOLERANCE"] + " percentage cannot be negative.")


            WH_ratio = W_input / H
            # Metin oluşturma işlemi İngilizce olarak ayarlandı
            model_base = self.current_lang['LABEL_MODEL'].split(':')[0]
            model_used = f"{model_base}: Reference Plane is {reference_plane_name}. Model: "
            cpwg_note = self.current_lang["CPWG_IGNORED"]
            
            if WH_ratio >= 1.0:
                model_used += f"Regime 1: WIDE Trace (W/H={WH_ratio:.2f}) - 'Best-Fit' Formula"
                Zdiff_result = calculate_wide_traces(W_input, Gap_input, T, H, Er)
            else:
                model_used += f"Regime 2: NARROW Trace (W/H={WH_ratio:.2f}) - 'Academic' Formula"
                Zdiff_result = calculate_narrow_traces(W_input, Gap_input, T, H, Er)
            
            if S_input < H:
                cpwg_note = self.current_lang["CPWG_APPLIED"]
                coplanar_factor = math.pow(S_input / (S_input + 0.5 * W_input), 0.1)
                Zdiff_result *= coplanar_factor

            self.Zdiff_result.set(f"{Zdiff_result:.2f}")
            
            Tolerance_Factor = Tolerance_P / 100.0
            Lower_Limit = Target_Z0 * (1.0 - Tolerance_Factor)
            Upper_Limit = Target_Z0 * (1.0 + Tolerance_Factor)

            if Lower_Limit <= Zdiff_result <= Upper_Limit:
                result_color = "green"
                status_text = f"{self.current_lang['STATUS_OK']} (Target range: {Lower_Limit:.2f}Ω - {Upper_Limit:.2f}Ω)"
                status_color = "green"
            else:
                result_color = "red"
                status_text = f"{self.current_lang['STATUS_FAIL']} (Target range: {Lower_Limit:.2f}Ω - {Upper_Limit:.2f}Ω)"
                status_color = "red"
                
            self.zdiff_result_label.config(foreground=result_color)
            self.tolerance_status_var.set(status_text)
            self.status_label.config(foreground=status_color)

            self.model_info.set(model_used)
            self.cpwg_info.set(cpwg_note)
            self.params_used.set(f"H={H:.3f} mm ({self.current_lang['MAP_H']}), T={T:.3f} mm, Er={Er:.2f} ({self.current_lang['MAP_ER']})")

        except ValueError as e:
            messagebox.showerror(self.current_lang["ERROR_INPUT"], f"{self.current_lang['ERROR_INPUT']}: {e}")
            self.Zdiff_result.set("ERROR")
            self.model_info.set("---")
            self.cpwg_info.set("---")
            self.params_used.set("---")
            self.tolerance_status_var.set(self.current_lang["ERROR_INPUT"])
        except Exception as e:
            messagebox.showerror(self.current_lang["ERROR_CALC"], f"{self.current_lang['ERROR_UNKNOWN']}: {e}")
            self.Zdiff_result.set("ERROR")
            self.model_info.set("---")
            self.cpwg_info.set("---")
            self.params_used.set("---")
            self.tolerance_status_var.set(self.current_lang["ERROR_CALC"])


# Ana Pencereyi Oluşturma ve Uygulamayı Başlatma
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = ImpedanceCalculatorApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Critical error occurred during application startup: {e}")