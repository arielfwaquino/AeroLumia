import sys
sys.dont_write_bytecode = True
import os, time, re, threading, winsound, ast, json, unicodedata, ctypes
import customtkinter as ctk
import pyautogui, win32gui, win32con, keyboard
from PIL import Image, ImageGrab
from bs4 import BeautifulSoup, NavigableString

from motor_core import *

F_N = ("Arial", 12)
F_B = ("Arial", 12, "bold")
F_H = ("Arial", 16, "bold")
F_BTN = ("Arial", 13, "bold")

if getattr(sys, 'frozen', False): 
    RESOURCE_DIR = sys._MEIPASS
    CONFIG_DIR = os.path.dirname(sys.executable) 
else: 
    RESOURCE_DIR = CONFIG_DIR = os.path.dirname(os.path.abspath(__file__))

DICT_LOAD_ERROR = False

def super_normalize(text):
    if not text: return ""
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9\s]', ' ', text.lower())).strip()

def load_external_data():
    global DICT_LOAD_ERROR
    txt_loaded, prompt, masks, frases, dicionario = [], "", "", "", {}
    db_path = next((p for p in [os.path.join(CONFIG_DIR, 'AeroDatabase.txt'), os.path.join(RESOURCE_DIR, 'AeroDatabase.txt')] if os.path.exists(p)), None)
            
    if db_path:
        try:
            with open(db_path, 'r', encoding='utf-8-sig') as f:
                blocks = re.split(r'\[###_(PROMPT|MASKS|FRASES|DICTIONARY)_###\]', f.read())
                txt_loaded.append(f"AeroDatabase.txt ({'Ext.' if db_path.startswith(CONFIG_DIR) and CONFIG_DIR != RESOURCE_DIR else 'Emb.'})")
                raw_dic = None
                for i in range(1, len(blocks), 2):
                    marker, body = blocks[i], blocks[i+1].strip()
                    if marker == "PROMPT" and body: prompt = body
                    elif marker == "MASKS" and body: masks = body
                    elif marker == "FRASES" and body: frases = body
                    elif marker == "DICTIONARY" and body:
                        try: raw_dic = json.loads(body)
                        except:
                            try: raw_dic = ast.literal_eval(body)
                            except: DICT_LOAD_ERROR = True
                            
                if raw_dic:
                    for key, value in raw_dic.items():
                        if key.startswith("_"): dicionario[key] = value
                        else:
                            dicionario[key] = {}
                            for kw, weight in value.items():
                                if norm_kw := super_normalize(kw): dicionario[key][norm_kw] = weight
        except: DICT_LOAD_ERROR = True
    return prompt, masks, frases, dicionario, txt_loaded

HIDDEN_PROMPT, HIDDEN_MASKS, HIDDEN_FRASES, HIDDEN_DICTIONARY, LOADED_FILES = load_external_data()

def extract_glued_mask(raw_text):
    if not HIDDEN_MASKS: return raw_text
    raw_blocks = re.split(r'\n\s*\n', raw_text.strip())
    if not raw_blocks: return raw_text
    last_raw_block = raw_blocks[-1]
    raw_words_list = list(re.finditer(r'[a-zA-ZÀ-ÿ0-9]+', last_raw_block))
    raw_words_norm = [super_normalize(w.group()) for w in raw_words_list]
    best_cut_pos = 0
    for mask_section in re.split(r'@@\s*', HIDDEN_MASKS):
        if not mask_section.strip(): continue
        last_mask_block = re.split(r'\n\s*\n', mask_section.strip())[-1].strip()
        mask_words_list = list(re.finditer(r'[a-zA-ZÀ-ÿ0-9]+', last_mask_block))
        mask_words_norm = [super_normalize(w.group()) for w in mask_words_list]
        mw_len = len(mask_words_norm)
        if mw_len < 5 or len(raw_words_norm) < mw_len: continue
        if raw_words_norm[:mw_len] == mask_words_norm:
            cut_pos = raw_words_list[mw_len - 1].end()
            if cut_pos > best_cut_pos: best_cut_pos = cut_pos
    if best_cut_pos > 0:
        user_text = last_raw_block[best_cut_pos:].strip(' \t\n\r.:,;-')
        if user_text: raw_blocks[-1] = user_text
        else: raw_blocks.pop()
        return '\n\n'.join(raw_blocks).strip()
    return raw_text

def apply_mask_filter(raw_text):
    if not HIDDEN_MASKS: return raw_text
    raw_text = extract_glued_mask(raw_text)
    trash_blocks = set()
    for mask_section in re.split(r'@@\s*', HIDDEN_MASKS):
        if not mask_section.strip(): continue
        for b in re.split(r'\n\s*\n', mask_section.strip()):
            norm_b = super_normalize(b)
            if len(norm_b) > 5: trash_blocks.add(norm_b)
    
    title_regex = HIDDEN_DICTIONARY.get("_REGEX_RULES", {}).get("title", r'(?i)(RM |TC )')
    lines = raw_text.strip().split('\n')
    if lines and re.search(title_regex, lines[0]): raw_text = '\n'.join(lines[1:]).strip()
    
    input_blocks = re.split(r'\n\s*\n', raw_text)
    kept_blocks = []
    mask_blocks_original = [b.strip() for ms in re.split(r'@@\s*', HIDDEN_MASKS) if ms.strip() for b in re.split(r'\n\s*\n', ms.strip()) if len(super_normalize(b)) > 5]
    mask_blocks_original.sort(key=len, reverse=True)
    trash_blocks_norm = {super_normalize(b) for b in mask_blocks_original}
    
    def create_regex_from_block(block): return r'\s+'.join([re.escape(w) for w in block.split()])
    
    for b in input_blocks:
        if not b.strip(' \t\n\r.:,;-…'): continue
        norm_b = super_normalize(b)
        if len(norm_b) <= 5: 
            kept_blocks.append(b); continue
        if norm_b in trash_blocks_norm: continue 
        modified_b = b
        for mask_block in mask_blocks_original:
            if len(mask_block.split()) < 3: continue 
            pattern = create_regex_from_block(mask_block)
            if match_start := re.match(r'^\s*' + pattern + r'\s*', modified_b, re.IGNORECASE):
                modified_b = modified_b[match_start.end():]; continue
            if match_end := re.search(r'\s*' + pattern + r'\s*$', modified_b, re.IGNORECASE):
                modified_b = modified_b[:match_end.start()]
        if len(super_normalize(modified_b)) > 5: kept_blocks.append(modified_b.strip())
    return '\n\n'.join(kept_blocks).strip()

def find_joint(raw_text, title_line=""):
    text_padded = " " + super_normalize(raw_text) + " "
    scores = {k: 0 for k in HIDDEN_DICTIONARY.keys() if not k.startswith("_")}
    for joint, keywords in HIDDEN_DICTIONARY.items():
        if joint.startswith("_"): continue
        if isinstance(keywords, dict):
            for kw, weight in keywords.items():
                if re.search(r'\b' + re.escape(kw) + r'\b', text_padded): scores[joint] += weight
    if title_line:
        title_words = super_normalize(title_line).split()
        for joint, data in HIDDEN_DICTIONARY.get("_BOOSTS", {}).items():
            if any(w in title_words for w in data.get("words", [])):
                scores[joint] = scores.get(joint, 0) + data.get("score", 0)
    max_hits = max(scores.values()) if scores else 0
    if max_hits == 0: return None
    best_joints = [j for j, s in scores.items() if s == max_hits]
    return best_joints[0] if best_joints else None

class LaudoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.original_clipboard = ""
        clear_clipboard_safely(retries=2)
        
        self.title("AeroLumia")
        self.geometry("1060x700") 
        self.resizable(False, False)
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        self.abort_flag = self.mouse_jail_active = False
        self.jail_x = self.jail_y = 0
        self.overlay_window = None
        threading.Thread(target=self._mouse_jail_loop, daemon=True).start()
        
        if os.path.exists(ipath := os.path.join(RESOURCE_DIR, 'proj_ia_icone.ico')):
            try: self.iconbitmap(ipath)
            except: pass

        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=(5, 10))
        self.main_frame.grid_columnconfigure(0, weight=55, uniform="col")
        self.main_frame.grid_columnconfigure(1, weight=45, uniform="col")
        self.main_frame.grid_rowconfigure(2, weight=1)

        self.header_left = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_left.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 5))
        self.header_right = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_right.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 5))

        if os.path.exists(img_path := os.path.join(RESOURCE_DIR, 'proj_IA_input.png')):
            self.lbl_hamster = ctk.CTkLabel(self.header_left, text="", image=ctk.CTkImage(Image.open(img_path), size=(90, 90)))
        else: self.lbl_hamster = ctk.CTkLabel(self.header_left, text="[IMG]", font=F_N)
        self.lbl_hamster.pack(side="left", anchor="center", padx=(0, 15))
        ctk.CTkLabel(self.header_left, text="INPUT - Dados Brutos", font=F_H).pack(side="left", anchor="center")

        if os.path.exists(img_path_perfil := os.path.join(RESOURCE_DIR, 'proj_IA_output.png')):
            self.lbl_hamster_direita = ctk.CTkLabel(self.header_right, text="", image=ctk.CTkImage(Image.open(img_path_perfil), size=(96, 96)))
        else: self.lbl_hamster_direita = ctk.CTkLabel(self.header_right, text="[IMG]", font=F_N)
        self.lbl_hamster_direita.pack(side="left", anchor="center", padx=(0, 15), pady=(12, 0))
        ctk.CTkLabel(self.header_right, text="OUTPUT - Pré-visualização", font=F_H).pack(side="left", anchor="center", pady=(3, 0))

        self.control_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.control_frame.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))
     
        ctk.CTkLabel(self.control_frame, text="Artic.:", font=F_B).pack(side="left", padx=(0, 5))
        
        self.lista_articulacoes_rm = ["AUTO", "ANTEPÉ", "BACIA", "COLUNA CERVICAL", "COLUNA DORSAL", "COLUNA LOMBAR", "COTOVELO", "JOELHO", "MÃO", "MEMBROS", "OMBRO", "PUNHO", "QUADRIL", "SACRO CÓCCIX SACROILÍACAS", "TORNOZELO"]
        self.lista_articulacoes_tc = self.lista_articulacoes_rm + ["MSK (Genérico)"]
        
        self.combo_articulacao = ctk.CTkComboBox(self.control_frame, values=self.lista_articulacoes_rm, width=130)
        self.combo_articulacao.pack(side="left", padx=(0, 15)); self.combo_articulacao.set("AUTO")
        
        self.seg_modality = ctk.CTkSegmentedButton(self.control_frame, values=["RM", "TC"], font=F_BTN, height=30, selected_color="#205c8a", selected_hover_color="#18476a", command=self.on_modality_change)
        self.seg_modality.pack(side="left", padx=(0, 15)); self.seg_modality.set("RM")
        
        self.chk_tradutor_var = ctk.IntVar(value=0)
        self.chk_tradutor = ctk.CTkCheckBox(self.control_frame, text="TRADUTOR", variable=self.chk_tradutor_var, font=F_B)
        self.chk_tradutor.pack(side="left", padx=(0, 15))

        self.chk_frases_var = ctk.IntVar(value=0)
        self.chk_frases = ctk.CTkCheckBox(self.control_frame, text="FRASES PADRÃO", variable=self.chk_frases_var, font=F_B, command=self.on_frases_toggle)
        self.chk_frases.pack(side="left", padx=(0, 0))

        self.control_frame_right = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.control_frame_right.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(0, 10))
        self.control_frame_right.lift()

        self.chk_365_var = ctk.IntVar(value=0)
        
        self.radio_free = ctk.CTkRadioButton(self.control_frame_right, text="COPILOT FREE", variable=self.chk_365_var, value=0, font=F_B, text_color="black", fg_color="#205c8a", hover_color="#18476a")
        self.radio_free.pack(side="left", padx=(0, 15))
        
        self.radio_365 = ctk.CTkRadioButton(self.control_frame_right, text="COPILOT 365 - EXPERIMENTAL", variable=self.chk_365_var, value=1, font=F_B, text_color="black", fg_color="#205c8a", hover_color="#18476a")
        self.radio_365.pack(side="left", padx=(0, 0))

        self.txt_input = ctk.CTkTextbox(self.main_frame, wrap="word", width=0, font=ctk.CTkFont(*F_N))
        self.txt_input.grid(row=2, column=0, sticky="nsew", padx=(0, 10)); self.txt_input.focus_set()
        
        self.txt_output = ctk.CTkTextbox(self.main_frame, state="disabled", fg_color="#f5f5f5", wrap="word", cursor="arrow", width=0, font=ctk.CTkFont(*F_N))
        self.txt_output.grid(row=2, column=1, sticky="nsew", padx=(10, 0))
        
        self.tk_text_core = self.txt_output._textbox
        self.tk_text_core.tag_configure("bold", font=ctk.CTkFont(*F_B))
        self.tk_text_core.tag_configure("italic", font=ctk.CTkFont("Arial", 12, slant="italic"))
        for ev in ["<Button-1>", "<B1-Motion>", "<Double-Button-1>", "<Triple-Button-1>", "<Control-c>", "<Control-x>", "<Control-a>"]: self.tk_text_core.bind(ev, lambda e: "break")

        self.bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_frame.pack(fill="x", padx=20, pady=(10, 25)) 
        
        self.frame_botoes = ctk.CTkFrame(self.bottom_frame, fg_color="transparent")
        self.frame_botoes.pack(pady=10, anchor="center")
        
        self.btn_gerar_pacs = ctk.CTkButton(self.frame_botoes, text="GERAR LAUDO - RIS/PACS (Alt+Shift+Z)", width=350, height=50, font=("Arial", 14, "bold"), command=self.btn_gerar_pacs_click)
        self.btn_gerar_pacs.pack(side="left", padx=10) 
        
        self.btn_gerar_input = ctk.CTkButton(self.frame_botoes, text="GERAR LAUDO - CAIXA DE INPUT", width=250, height=40, font=F_BTN, command=self.btn_gerar_input_click)
        self.btn_gerar_input.pack(side="left", padx=10)
        
        self.btn_novo = ctk.CTkButton(self.frame_botoes, text="RESET / ABORTAR", width=120, height=40, font=F_BTN, command=self.trigger_global_abort, fg_color="#d9534f", hover_color="#c9302c")
        self.btn_novo.pack(side="left", padx=10)
        
        self.lbl_status = ctk.CTkLabel(self.bottom_frame, text="", font=F_N)
        self.lbl_status.pack(pady=(30, 5)); self.selected_joint = None
        
        try: 
            keyboard.add_hotkey('alt+shift+z', self.on_global_hotkey)
            keyboard.add_hotkey('ctrl+esc', self.trigger_global_abort)
            keyboard.add_hotkey('pause', self.trigger_global_abort)
            keyboard.add_hotkey('end', self.trigger_global_abort)
        except: pass
        
        if DICT_LOAD_ERROR: self.update_status("ERRO NO JSON DO AeroDatabase.txt.", "red", size=13, bold=True)
        elif LOADED_FILES: self.update_status(f"Sistema Inicializado. Lidos: {', '.join(LOADED_FILES)}", "green")
        else: self.update_status("Sistema Inicializado.", "green")

    def silent_cleanup(self):
        self.abort_flag = False; self.selected_joint = None
        self.txt_output.configure(state="normal"); self.txt_output.delete("1.0", "end"); self.txt_output.configure(state="disabled")

    def auto_translate(self, texto, modo):
        chave = "_TRANSLATE_RM_TC" if modo == "RM_PARA_TC" else "_TRANSLATE_TC_RM"
        for padrao, substituto in HIDDEN_DICTIONARY.get(chave, []): texto = re.sub(padrao, substituto, texto)
        return re.sub(r' {2,}', ' ', re.sub(r'\n{3,}', '\n\n', texto)).strip()

    def on_modality_change(self, value):
        current_val = self.combo_articulacao.get()
        if value == "TC": self.combo_articulacao.configure(values=self.lista_articulacoes_tc)
        else:
            self.combo_articulacao.configure(values=self.lista_articulacoes_rm)
            if current_val == "MSK (Genérico)": self.combo_articulacao.set("AUTO")

    def on_frases_toggle(self):
        if self.chk_frases_var.get() == 1:
            warn_win = ctk.CTkToplevel(self)
            warn_win.title("Aviso"); warn_win.geometry("450x150"); warn_win.attributes('-topmost', True); warn_win.grab_set() 
            ctk.CTkLabel(warn_win, text="Ocasionalmente o Copilot pode apresentar erros\ncom prompts longos.", font=("Arial", 14, "bold"), text_color="#d9534f").pack(expand=True, pady=(20, 10))
            ctk.CTkButton(warn_win, text="OK", font=("Arial", 14, "bold"), command=warn_win.destroy, width=120).pack(pady=(0, 20))

    def _mouse_jail_loop(self):
        while True:
            if self.mouse_jail_active and not self.abort_flag: ctypes.windll.user32.SetCursorPos(int(self.jail_x), int(self.jail_y))
            time.sleep(0.01)

    def enable_idiot_proof_shield(self):
        try:
            gui_info = GUITHREADINFO(); gui_info.cbSize = ctypes.sizeof(GUITHREADINFO)
            hwnd_active = ctypes.windll.user32.GetForegroundWindow()
            thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd_active, None)
            moved_to_caret = False
            if ctypes.windll.user32.GetGUIThreadInfo(thread_id, ctypes.byref(gui_info)) and gui_info.hwndCaret:
                pt = wintypes.POINT(gui_info.rcCaret.left, gui_info.rcCaret.top); ctypes.windll.user32.ClientToScreen(gui_info.hwndCaret, ctypes.byref(pt))
                self.jail_x, self.jail_y = pt.x, pt.y; moved_to_caret = True
            if not moved_to_caret and hwnd_active:
                rect = win32gui.GetWindowRect(hwnd_active)
                self.jail_x, self.jail_y = rect[2] - 150, rect[1] + 250
        except: self.jail_x, self.jail_y = pyautogui.position()
        self.mouse_jail_active = True; self.after(0, self._show_ghost_overlay)

    def disable_idiot_proof_shield(self):
        self.mouse_jail_active = False; self.after(0, self._hide_ghost_overlay)

    def _show_ghost_overlay(self):
        if self.overlay_window: return
        self.overlay_window = ctk.CTkToplevel(self)
        self.overlay_window.overrideredirect(True); self.overlay_window.attributes('-topmost', True)
        self.overlay_window.geometry(f"{self.winfo_screenwidth()}x80+0+{self.winfo_screenheight() - 80}")
        self.overlay_window.configure(fg_color="#8B0000")
        ctk.CTkLabel(self.overlay_window, text="⚠️ ROBÔ TRABALHANDO - NÃO MEXA NO MOUSE OU TECLADO (Pressione PAUSE ou END para Abortar) ⚠️", font=("Arial", 22, "bold"), text_color="white").pack(expand=True, fill="both")
        hwnd = win32gui.GetParent(self.overlay_window.winfo_id()); style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED)
        win32gui.SetLayeredWindowAttributes(hwnd, 0, 230, win32con.LWA_ALPHA) 

    def _hide_ghost_overlay(self):
        if self.overlay_window: self.overlay_window.destroy(); self.overlay_window = None

    def sleep_safe(self, seconds):
        end_time = time.time() + seconds
        while time.time() < end_time:
            if self.abort_flag: return False
            time.sleep(0.05)
        return True

    def trigger_global_abort(self):
        self.abort_flag = True; self.disable_idiot_proof_shield(); release_stuck_modifiers(); self.after(0, self.reset_ui_completely)

    def reset_ui_completely(self):
        self.txt_output.configure(state="normal"); self.txt_output.delete("1.0", "end"); self.txt_output.configure(state="disabled")
        self.combo_articulacao.set("AUTO"); self.change_ui_state("normal"); self.selected_joint = None
        if hasattr(self, 'original_clipboard') and self.original_clipboard: robust_clipboard_copy_text(self.original_clipboard)
        self.update_status("SISTEMA ABORTADO/RESETADO PELO USUÁRIO.", "red", 14, True)

    def show_error_window(self, title, msg, btn_text, geo="400x200", click_edge=True):
        self.abort_flag = True; self.disable_idiot_proof_shield(); release_stuck_modifiers(); self.reset_ui_completely()
        self.lift(); self.attributes('-topmost', True); self.after(500, lambda: self.attributes('-topmost', False)); self.focus_force()
        try: winsound.MessageBeep(winsound.MB_ICONHAND) 
        except: pass
        err_win = ctk.CTkToplevel(self); err_win.title(title); err_win.geometry(geo); err_win.attributes('-topmost', True); err_win.grab_set() 
        ctk.CTkLabel(err_win, text=msg, font=("Arial", 16, "bold"), text_color="#B22222").pack(expand=True, pady=(20, 10))
        def on_ok():
            err_win.destroy()
            if click_edge and (hwnd := get_window_by_title("edge")):
                try:
                    if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd); time.sleep(0.3)
                    rect = win32gui.GetWindowRect(hwnd); pyautogui.click(rect[2] - 150, rect[1] + 250)
                except: pass
        ctk.CTkButton(err_win, text=btn_text, font=("Arial", 14, "bold"), command=on_ok, width=150, height=40).pack(pady=(0, 20))

    def show_login_error_ui(self):
        msg = "Acesso ao Copilot bloqueado - Login necessário\n\nSelecione entrar com GOOGLE ou APPLE\n\n(NÃO USE \"Copilot 365\")\n\nEntão digite \"teste\" no copilot, aceite cookies e prove que é humano"
        self.show_error_window("Atenção: Login Necessário", msg, "OK, ENTENDI", "700x400")

    def show_login_error_365_ui(self):
        msg = "Acesso ao Copilot 365 bloqueado\n\nLogin necessário\n\nNa sequência, digite \"teste\" no Copilot 365, aceite os cookies e prove que é humano"
        self.show_error_window("Atenção: Login Necessário", msg, "OK, CIENTE", "700x400")

    def show_cookie_error_ui(self):
        self.show_error_window("Atenção: Cookies", "Por favor aceite os cookies", "OK", "450x150")

    def on_global_hotkey(self):
        if self.btn_gerar_pacs.cget("state") == "normal": self.after(0, self.btn_gerar_pacs_click)

    def apply_msk_style_visually(self):
        self.update_status("[ VISÃO ] Escaneando área para aplicar estilo...", "orange")
        if not self.sleep_safe(0.6): return False
        
        anchor1, anchor2 = get_anchor_paths()
        if not anchor1 or not anchor2: return False
        
        try:
            vscreen_x, vscreen_y = ctypes.windll.user32.GetSystemMetrics(76), ctypes.windll.user32.GetSystemMetrics(77)
            try: full_img = ImageGrab.grab(all_screens=True)
            except TypeError: full_img = ImageGrab.grab()
            
            hwnd = win32gui.GetForegroundWindow()
            offset_x, offset_y = vscreen_x, vscreen_y
            screen_img = full_img
            
            if hwnd:
                rect = win32gui.GetWindowRect(hwnd)
                if rect[2] > rect[0] and rect[3] > rect[1]:
                    crop_left = max(0, rect[0] - vscreen_x)
                    crop_top = max(0, rect[1] - vscreen_y)
                    crop_right = min(full_img.width, rect[2] - vscreen_x)
                    crop_bottom = min(full_img.height, rect[3] - vscreen_y)
                    
                    if crop_right > crop_left and crop_bottom > crop_top:
                        screen_img = full_img.crop((crop_left, crop_top, crop_right, crop_bottom))
                        offset_x = vscreen_x + crop_left
                        offset_y = vscreen_y + crop_top
            
            if self.abort_flag: return False
            
            box1 = pyautogui.locate(anchor1, screen_img, confidence=0.7, grayscale=True)
            box2 = pyautogui.locate(anchor2, screen_img, confidence=0.7, grayscale=True)
            
            if box1 and box2:
                click_x = offset_x + box1.left + (box1.width // 2) + int(((box2.left + (box2.width // 2)) - (box1.left + (box1.width // 2))) * 0.34)
                click_y = offset_y + box1.top + (box1.height // 2)
                
                pyautogui.moveTo(x=click_x, y=click_y)
                time.sleep(0.1)
                pyautogui.click() 
                
                if not self.sleep_safe(0.3): return False
                robust_clipboard_copy_text('MSK'); pyautogui.hotkey('ctrl', 'v')
                if not self.sleep_safe(0.2): return False
                hardware_enter(); return True
        except: pass
        return False

    def change_ui_state(self, state):
        for w in [self.btn_gerar_pacs, self.btn_gerar_input, self.combo_articulacao, self.seg_modality, self.chk_tradutor, self.chk_frases, self.radio_free, self.radio_365]: w.configure(state=state)

    def finaliza_rotina(self, msg, color, success=True):
        self.disable_idiot_proof_shield(); release_stuck_modifiers()
        if self.abort_flag: return 
        if success:
            self.seg_modality.set("RM"); self.on_modality_change("RM")
            self.chk_tradutor_var.set(0); self.chk_frases_var.set(0)
        self.change_ui_state("normal"); self.combo_articulacao.set("AUTO")
        self.update_status(msg, color, 24, True)

    def update_status(self, text, color="black", size=12, bold=False):
        self.lbl_status.configure(text=text, text_color=color, font=("Arial", size, "bold") if bold else ("Arial", size))
        self.update_idletasks()

    def render_html_to_output(self, html_string):
        self.txt_output.configure(state="normal"); self.txt_output.delete("1.0", "end")
        def nl(chars="\n\n"):
            txt = self.tk_text_core.get("1.0", "end-1c")
            if not txt: return
            if chars == "\n\n" and not txt.endswith("\n\n"): self.tk_text_core.insert("end", "\n" if txt.endswith("\n") else "\n\n")
            elif chars == "\n" and not txt.endswith("\n"): self.tk_text_core.insert("end", "\n")

        def walk_tree(node, active_tags):
            if isinstance(node, NavigableString):
                text = re.sub(r' {2,}', ' ', str(node).replace('\n', ' ').replace('\r', '').replace('\t', ' '))
                if text.startswith(' ') and (not (c := self.tk_text_core.get("1.0", "end-1c")) or c.endswith('\n')): text = text[1:]
                if text: self.tk_text_core.insert("end", text, tuple(active_tags))
            else:
                current_tags = active_tags.copy()
                is_bold = False
                if node.name in ['b', 'strong']: is_bold = True
                elif node.get('style'):
                    if match := re.search(r'font-weight\s*:\s*([a-z0-9]+)', node.get('style').lower()):
                        val = match.group(1)
                        if val in ['bold', 'bolder'] or (val.isdigit() and int(val) >= 600): is_bold = True
                if is_bold: current_tags.append("bold")
                if node.name in ['i', 'em']: current_tags.append("italic")
                if node.name in ['p', 'div']: nl("\n\n")
                elif node.name == 'li': nl("\n"); self.tk_text_core.insert("end", "• ", tuple(current_tags))
                elif node.name == 'br': nl("\n")
                for child in node: walk_tree(child, current_tags)
                if node.name in ['p', 'div']: nl("\n\n")
                elif node.name == 'li': nl("\n")

        walk_tree(BeautifulSoup(html_string, 'html.parser'), [])
        if (lw := len(c := self.tk_text_core.get("1.0", "end-1c")) - len(c.lstrip())) > 0: self.tk_text_core.delete("1.0", f"1.0+{lw}c")
        self.txt_output.configure(state="disabled")

    def request_manual_joint(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Definição Manual"); dialog.geometry("300x200"); dialog.attributes("-topmost", True); dialog.grab_set() 
        ctk.CTkLabel(dialog, text="Articulação não identificada.\nSelecione abaixo:").pack(pady=15)
        combo = ctk.CTkComboBox(dialog, values=["SELECIONE..."] + sorted([k for k in HIDDEN_DICTIONARY.keys() if not k.startswith("_")]))
        combo.set("SELECIONE..."); combo.pack(pady=10)
        def confirm(): 
            if combo.get() != "SELECIONE...": self.selected_joint = combo.get(); dialog.destroy()
        ctk.CTkButton(dialog, text="Confirmar", command=confirm).pack(pady=15); self.wait_window(dialog) 

    def extract_prompt_block(self, joint_name, modality):
        if not joint_name: return ""
        if joint_name == "MSK (Genérico)": joint_name = "MSK"
        joint_norm = super_normalize(joint_name)
        joint_keywords = ["sacro", "coccix", "sacroiliaca"] if joint_norm == "sacro coccix sacroiliacas" else [joint_norm]
        mod_keywords = ["rm", "ressonancia"] if modality == "RM" else ["tc", "tomografia"]
        old_style = super_normalize(f"{joint_name} {modality}")

        for block in re.split(r'@@', HIDDEN_PROMPT):
            if not block.strip(): continue
            header = super_normalize(block.strip().split('\n')[0])
            if (any(m in header for m in mod_keywords) and any(j in header for j in joint_keywords)) or header.startswith(old_style): return "@@" + block
        return ""
        
    def extract_frases_block(self, joint_name, modality):
        if not joint_name or not HIDDEN_FRASES: return ""
        if joint_name == "MSK (Genérico)": joint_name = "MSK"
        joint_norm = super_normalize(joint_name)
        joint_keywords = ["sacro", "coccix", "sacroiliaca"] if joint_norm == "sacro coccix sacroiliacas" else [joint_norm]
        mod_keywords = ["rm", "ressonancia"] if modality == "RM" else ["tc", "tomografia"]
        old_style = super_normalize(f"{joint_name} {modality}")
        
        for block in re.split(r'@\*@', HIDDEN_FRASES):
            if not block.strip(): continue
            lines = block.strip().split('\n')
            header = super_normalize(lines[0])
            if (any(m in header for m in mod_keywords) and any(j in header for j in joint_keywords)) or header.startswith(old_style): return '\n'.join(lines[1:]).strip()
        return ""

    def btn_gerar_input_click(self):
        raw_text = self.txt_input.get("1.0", "end-1c").strip()
        if not raw_text: return self.update_status("ERRO: Caixa vazia.", "red")
        self.original_clipboard = robust_clipboard_paste_text()
        self.silent_cleanup(); self.change_ui_state("disabled"); self.enable_idiot_proof_shield(); self.process_raw_text_and_run(raw_text)

    def btn_gerar_pacs_click(self):
        self.original_clipboard = robust_clipboard_paste_text()
        self.silent_cleanup(); self.change_ui_state("disabled"); self.update_status("[ A EXTRAIR DADOS DO VUE PACS... ]", "blue")
        threading.Thread(target=self.fetch_from_pacs_and_run, daemon=True).start()

    def fetch_from_pacs_and_run(self):
        if not self.sleep_safe(0.4): return
        hwnd = get_window_by_title("vue pacs")
        if not hwnd: return self.after(0, lambda: self.finaliza_rotina("ERRO: Vue PACS não encontrado.", "red", False))
        try:
            pyautogui.press('alt'); time.sleep(0.05)
            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW); win32gui.SetForegroundWindow(hwnd); time.sleep(0.6)
        except: pass 
        release_stuck_modifiers()
        try:
            pyautogui.hotkey('ctrl', 'shift', 't'); time.sleep(0.1); pyautogui.hotkey('ctrl', 'shift', 't'); time.sleep(0.1)
            self.after(0, self.enable_idiot_proof_shield)
            if not self.sleep_safe(0.1): return
            pyautogui.hotkey('ctrl', 'a'); time.sleep(0.1); pyautogui.hotkey('ctrl', 'c'); time.sleep(0.1); pyautogui.press('right')
        except: release_stuck_modifiers()
        if self.abort_flag: return
        raw_text = robust_clipboard_paste_text()
        if not raw_text or len(raw_text.strip()) < 5: return self.after(0, lambda: self.finaliza_rotina("ERRO: Texto vazio.", "red", False))
        self.after(0, lambda: self.process_raw_text_and_run(raw_text.strip(), update_ui=True))

    def process_raw_text_and_run(self, raw_text, update_ui=False):
        if self.abort_flag: return
        self.update_status("[ FILTRO DE MÁSCARAS E PREPARAÇÃO... ]", "blue")
        
        rules = HIDDEN_DICTIONARY.get("_REGEX_RULES", {})
        title_regex = rules.get("title", r'(?i)(RM |TC )')
        tc_regex = rules.get("tc", r'(?i)(tc\b)')
        rm_regex = rules.get("rm", r'(?i)(rm\b)')

        first_line = raw_text.strip().split('\n')[0] if raw_text.strip() else ""
        title_line = first_line if re.search(title_regex, first_line) else ""
        
        input_modality = "TC" if re.search(tc_regex, raw_text) else ("RM" if re.search(rm_regex, raw_text) else "RM")
        output_modality = self.seg_modality.get()
        
        if input_modality == "RM" and output_modality == "TC": raw_text = self.auto_translate(raw_text, "RM_PARA_TC")
        elif input_modality == "TC" and output_modality == "RM": raw_text = self.auto_translate(raw_text, "TC_PARA_RM")

        clean_input = apply_mask_filter(raw_text)
        if update_ui: self.after(0, lambda: self.txt_input.delete("1.0", "end") or self.txt_input.insert("1.0", clean_input))
            
        combo_val = self.combo_articulacao.get()
        self.selected_joint = combo_val if combo_val != "AUTO" else find_joint(clean_input, title_line=title_line)
            
        if not self.selected_joint:
            self.disable_idiot_proof_shield(); self.request_manual_joint(); self.enable_idiot_proof_shield()
            
        if not self.selected_joint or self.abort_flag: return self.after(0, lambda: self.finaliza_rotina("Operação cancelada.", "red", False))
        
        prompt_block = self.extract_prompt_block(self.selected_joint, output_modality)
        if not prompt_block and output_modality == "TC": prompt_block = self.extract_prompt_block("MSK", "TC")
        
        if self.chk_frases_var.get() == 1:
            frases_block = self.extract_frases_block(self.selected_joint, output_modality)
            if not frases_block and output_modality == "TC": frases_block = self.extract_frases_block("MSK", "TC")
            prompt_block = re.sub(r'<frasein[ií]cio>|<frasefim>', '', prompt_block.replace("<@insfra@>", frases_block), flags=re.IGNORECASE)
        else:
            prompt_block = re.sub(r'<frasein[ií]cio>.*?<frasefim>', '', prompt_block.replace("<@insfra@>", ""), flags=re.IGNORECASE | re.DOTALL)

        payload = f"{prompt_block}\n\n[DADOS BRUTOS]:\n{clean_input}\n\nREGRA FINAL: O laudo DEVE ser gerado entre as tags [[INI]] e [[FIM]]. Exemplo: [[INI]] texto do laudo [[FIM]]."
        if self.chk_tradutor_var.get() == 1: 
            if trans_rule := HIDDEN_DICTIONARY.get("_PROMPT_TRANSLATION", ""): payload += trans_rule
            
        robust_clipboard_copy_text(payload)
        threading.Thread(target=self.run_rpa_automation, args=(payload,), daemon=True).start()

    def run_rpa_automation(self, payload):
        if self.abort_flag: return
        is_365 = self.chk_365_var.get() == 1
        url_alvo = 'https://m365.cloud.microsoft/chat' if is_365 else 'https://copilot.microsoft.com/'
        ini_sleep = 9.0 if is_365 else 1.4
        err_time = 20.0 if is_365 else 8.0
        timeout = 55.0
        
        self.after(0, lambda: self.update_status(f"[ REDIRECIONANDO EDGE ({'365' if is_365 else 'FREE'})... ]", "orange"))
        reuse_or_open_edge(url_alvo)
        if not self.sleep_safe(ini_sleep): return
        if not getattr(self, 'first_run_done', False):
            if not self.sleep_safe(1.0): return 
            self.first_run_done = True
            
        robust_clipboard_copy_text(payload)
        start_time = time.time()
        sucesso = precisa_login = precisa_cookies = False
        
        try:
            if self.abort_flag: return
            pyautogui.hotkey('ctrl', 'v'); time.sleep(0.3)
            pyautogui.press('enter'); time.sleep(0.2); pyautogui.press('enter'); time.sleep(1.0)
            
            while (time.time() - start_time) < timeout:
                if self.abort_flag or not self.sleep_safe(1.5): return
                edge_hwnd = get_window_by_title("edge")
                if edge_hwnd:
                    rect = win32gui.GetWindowRect(edge_hwnd)
                    self.jail_x, self.jail_y = rect[2] - 150, rect[1] + 250
                
                pyautogui.scroll(-500) 
                if self.abort_flag: return
                
                pyautogui.click(self.jail_x, self.jail_y); time.sleep(0.1)
                clear_clipboard_safely(); pyautogui.hotkey('ctrl', 'a'); time.sleep(0.1); pyautogui.hotkey('ctrl', 'c'); time.sleep(0.1)
                copiado = robust_clipboard_paste_text()
                
                if (time.time() - start_time) > err_time and "REGRA FINAL" not in copiado: precisa_login = True; break
                if re.search(r'(?i)entrar no\s*copilot|entrar com a conta', copiado): precisa_login = True; break
                if re.search(r'(?i)Valorizamos a sua privacidade', copiado): precisa_cookies = True; break
                if [m for m in re.findall(r'\[\[INI\]\](.*?)\[\[FIM\]\]', copiado, re.I | re.S) if len(m.strip()) > 30]: sucesso = True; break
        except: release_stuck_modifiers()
        
        if self.abort_flag: return
        if precisa_login: self.after(0, self.show_login_error_365_ui if is_365 else self.show_login_error_ui)
        elif precisa_cookies: self.after(0, self.show_cookie_error_ui)
        elif sucesso: self.after(0, self.process_final_html)
        else: self.after(0, lambda: self.finaliza_rotina(f"ERRO: Timeout ({int(timeout)}s).", "red", False))

    def process_final_html(self):
        if self.abort_flag: return
        self.update_status("[ FORMATANDO HTML... ]", "blue")
        clean_html, plain_text = sanitize_and_clean_html(extract_html_from_clipboard() or "")
        if not clean_html:
            validos = [m for m in re.findall(r'\[\[INI\]\](.*?)\[\[FIM\]\]', robust_clipboard_paste_text(), re.I | re.S) if len(m.strip()) > 30]
            if validos:
                best_match = html.unescape(max(validos, key=len).strip()).replace('\xa0', ' ').replace('\u200b', '')
                best_match = re.sub(r'(?i)LESAOMAIOR:?\s*', '', best_match).replace("****", "").replace("** **", "")
                clean_html = plain_text = re.sub(r':\s*([A-Za-zÀ-ÿ])', r': \1', best_match)
                
        if clean_html and plain_text: 
            inject_dual_format_clipboard(clean_html, plain_text)
            self.render_html_to_output(clean_html)
            threading.Thread(target=self.paste_and_select_in_ris, args=(plain_text, clean_html), daemon=True).start()
        else: self.after(0, lambda: self.finaliza_rotina("ERRO: Falha no conteúdo.", "red", False))

    def paste_and_select_in_ris(self, plain_text, clean_html=None):
        if self.abort_flag: return
        self.after(0, lambda: self.update_status("[ COLANDO NO PACS... ]", "orange"))
        if not (hwnd := get_window_by_title("vue pacs")): return self.after(0, self.finalize_ui_success_manual)
            
        try:
            pyautogui.press('alt'); time.sleep(0.05)
            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW); win32gui.SetForegroundWindow(hwnd); time.sleep(0.6)
        except: pass 
            
        if not self.sleep_safe(0.5): return
        release_stuck_modifiers()
        try:
            if self.abort_flag: return
            pyautogui.hotkey('ctrl', 'shift', 't'); time.sleep(0.2); pyautogui.hotkey('ctrl', 'shift', 't'); time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'end'); time.sleep(0.1); pyautogui.press('enter'); pyautogui.press('enter'); time.sleep(0.1)
            pyautogui.keyDown('ctrl'); time.sleep(0.05); pyautogui.keyDown('shift'); time.sleep(0.05)
            pyautogui.press('v'); time.sleep(0.05); pyautogui.keyUp('shift'); pyautogui.keyUp('ctrl')
            time.sleep(0.5); self.after(0, lambda: self.update_status("[ REPOSICIONANDO CURSOR E SELECIONANDO... ]", "orange"))
            send_fast_left_arrows(len(plain_text.replace('\n\n', '\n')) + 1)
            if not self.sleep_safe(0.4): return
            select_with_hardware_scancodes()
        except: release_stuck_modifiers()
        
        if self.abort_flag: return
        self.mouse_jail_active = False; time.sleep(0.15) 
        estilo_ok = self.apply_msk_style_visually() if not self.abort_flag else False
        
        if clean_html and plain_text and not self.abort_flag: inject_dual_format_clipboard(clean_html, plain_text)
        if self.abort_flag: return
        self.disable_idiot_proof_shield()
        if estilo_ok and not self.abort_flag: self.after(0, lambda: self.finaliza_rotina("LAUDO E ESTILO PRONTOS!", "green", True))
        elif not self.abort_flag: self.after(0, lambda: self.finaliza_rotina("LAUDO PRONTO (Estilo manual necessário)", "black", True))

    def finalize_ui_success_manual(self):
        if self.abort_flag: return
        self.lift(); self.attributes('-topmost', True); self.after(500, lambda: self.attributes('-topmost', False))
        self.finaliza_rotina("COLE MANUALMENTE", "orange", True)

if __name__ == "__main__":
    app = LaudoApp()
    app.mainloop()
