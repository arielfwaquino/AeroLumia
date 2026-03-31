import sys, os, ctypes, time, re, threading, ast, subprocess, webbrowser, difflib, unicodedata, json, winsound, html
import customtkinter as ctk
import pyautogui, win32clipboard, win32gui, win32con, keyboard
from tkinter import messagebox
from ctypes import wintypes
from bs4 import BeautifulSoup, NavigableString
from PIL import Image, ImageGrab 

try: 
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try: 
        ctypes.windll.user32.SetProcessDPIAware()
    except: 
        pass

ULONG_PTR = ctypes.c_uint64 if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_uint32

class KEYBDINPUT(ctypes.Structure): 
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]

class MOUSEINPUT(ctypes.Structure): 
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]

class HARDWAREINPUT(ctypes.Structure): 
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD), ("wParamH", wintypes.WORD)]

class INPUT_UNION(ctypes.Union): 
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]

class INPUT(ctypes.Structure): 
    _fields_ = [("type", wintypes.DWORD), ("union", INPUT_UNION)]

class RECT(ctypes.Structure): 
    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

class GUITHREADINFO(ctypes.Structure): 
    _fields_ = [("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD), ("hwndActive", wintypes.HWND), ("hwndFocus", wintypes.HWND), ("hwndCapture", wintypes.HWND), ("hwndMenuOwner", wintypes.HWND), ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND), ("rcCaret", RECT)]

def release_stuck_modifiers():
    for k in ['ctrl', 'shift', 'alt', 'win']: 
        pyautogui.keyUp(k)
    for scan in [0x1D, 0x2A]: 
        inp = INPUT(type=1)
        inp.union.ki.wScan = scan
        inp.union.ki.dwFlags = 0x0008 | 0x0002
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

def send_fast_left_arrows(count):
    if count <= 0: 
        return
    VK_LEFT, KEYEVENTF_KEYUP, chunk_size = 0x25, 0x0002, 500 
    for chunk_start in range(0, count, chunk_size):
        actual_count = min(chunk_size, count - chunk_start)
        inputs = (INPUT * (actual_count * 2))()
        for i in range(actual_count):
            inputs[i*2].type = 1
            inputs[i*2].union.ki.wVk = VK_LEFT
            inputs[i*2].union.ki.dwFlags = 0
            inputs[i*2+1].type = 1
            inputs[i*2+1].union.ki.wVk = VK_LEFT
            inputs[i*2+1].union.ki.dwFlags = KEYEVENTF_KEYUP
        ctypes.windll.user32.SendInput(len(inputs), ctypes.byref(inputs), ctypes.sizeof(INPUT))
        time.sleep(0.002)

def select_with_hardware_scancodes():
    def press_scancode(scan, extend=False, up=False):
        inp = INPUT(type=1)
        inp.union.ki.wScan = scan
        inp.union.ki.dwFlags = 0x0008 | (0x0001 if extend else 0) | (0x0002 if up else 0)
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
    
    press_scancode(0x1D)
    press_scancode(0x2A)
    time.sleep(0.05) 
    press_scancode(0x4F, extend=True)
    time.sleep(0.05)
    press_scancode(0x4F, extend=True, up=True)
    time.sleep(0.05) 
    press_scancode(0x2A, up=True)
    press_scancode(0x1D, up=True)

def hardware_enter():
    inp = INPUT(type=1)
    inp.union.ki.wScan = 0x1C
    inp.union.ki.dwFlags = 0x0008 
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
    time.sleep(0.01) 
    inp.union.ki.dwFlags = 0x000A
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

if getattr(sys, 'frozen', False): 
    RESOURCE_DIR = sys._MEIPASS
    CONFIG_DIR = os.path.dirname(sys.executable) 
else: 
    RESOURCE_DIR = CONFIG_DIR = os.path.dirname(os.path.abspath(__file__))

HIDDEN_PROMPT_FALLBACK = "@@COLE SEU PROMPT AQUI"
HIDDEN_MASKS_FALLBACK = "@@COLE SUAS MASCARAS AQUI"
HIDDEN_FRASES_FALLBACK = "@*@COLE SUAS FRASES AQUI"
HIDDEN_DICTIONARY_FALLBACK = {}
DICT_LOAD_ERROR = False

def super_normalize(text):
    if not text: return ""
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')
    text = text.lower()
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()

def load_file_with_fallback(filename, fallback_value):
    for path in [os.path.join(CONFIG_DIR, filename), os.path.join(RESOURCE_DIR, filename)]:
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                c = f.read().strip()
                if c: 
                    return c, f"{filename} ({'Externo' if path.startswith(CONFIG_DIR) and CONFIG_DIR != RESOURCE_DIR else 'Embutido'})"
        except: 
            pass
    return fallback_value, None

def load_external_data():
    global DICT_LOAD_ERROR
    txt_loaded = []
    
    prompt, msg1 = load_file_with_fallback('hidden_prompt.txt', HIDDEN_PROMPT_FALLBACK)
    if msg1: txt_loaded.append(msg1)
    
    masks, msg2 = load_file_with_fallback('hidden_masks.txt', HIDDEN_MASKS_FALLBACK)
    if msg2: txt_loaded.append(msg2)
    
    frases, msg3 = load_file_with_fallback('hidden_frases.txt', HIDDEN_FRASES_FALLBACK)
    if msg3: txt_loaded.append(msg3)
    
    dic_str, msg4 = load_file_with_fallback('Hidden_Dictionary.txt', "{}")
    dicionario = {}
 
    if msg4: 
        try: 
            raw_dic = json.loads(dic_str)
            txt_loaded.append(msg4)
        except: 
            try:
                raw_dic = ast.literal_eval(dic_str)
                txt_loaded.append(msg4)
            except:
                txt_loaded.append("ERRO AO LER: Hidden_Dictionary.txt")
                raw_dic = HIDDEN_DICTIONARY_FALLBACK
                DICT_LOAD_ERROR = True
        
        for joint, keywords in raw_dic.items():
            dicionario[joint] = {}
            for kw, weight in keywords.items():
                norm_kw = super_normalize(kw)
                if norm_kw:
                    dicionario[joint][norm_kw] = weight
                    
    for joint in ["BACIA", "QUADRIL", "COLUNA LOMBAR", "SACRO CÓCCIX SACROILÍACAS"]:
        if joint not in dicionario: dicionario[joint] = {}
        dicionario[joint][super_normalize("doença de paget")] = 5
        dicionario[joint][super_normalize("paget")] = 5

    return prompt, masks, frases, dicionario, txt_loaded

HIDDEN_PROMPT, HIDDEN_MASKS, HIDDEN_FRASES, HIDDEN_DICTIONARY, LOADED_FILES = load_external_data()

def remove_accents(text):
    return unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')

def normalize_block(text):
    return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', '', remove_accents(text))).lower().strip()

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
        
        mask_blocks = re.split(r'\n\s*\n', mask_section.strip())
        last_mask_block = mask_blocks[-1].strip()
        
        mask_words_list = list(re.finditer(r'[a-zA-ZÀ-ÿ0-9]+', last_mask_block))
        mask_words_norm = [super_normalize(w.group()) for w in mask_words_list]
        
        if len(mask_words_norm) < 5: continue
        
        mw_len = len(mask_words_norm)
        if len(raw_words_norm) < mw_len: continue
        
        if raw_words_norm[:mw_len] == mask_words_norm:
            cut_pos = raw_words_list[mw_len - 1].end()
            if cut_pos > best_cut_pos:
                best_cut_pos = cut_pos
                
    if best_cut_pos > 0:
        user_text = last_raw_block[best_cut_pos:].strip(' \t\n\r.:,;-')
        if user_text:
            raw_blocks[-1] = user_text
        else:
            raw_blocks.pop()
        return '\n\n'.join(raw_blocks).strip()
        
    return raw_text

def apply_mask_filter(raw_text):
    if not HIDDEN_MASKS: 
        return raw_text
        
    raw_text = extract_glued_mask(raw_text)

    trash_blocks = set()
    for mask_section in re.split(r'@@\s*', HIDDEN_MASKS):
        if not mask_section.strip(): 
            continue
        for b in re.split(r'\n\s*\n', mask_section.strip()):
            norm_b = super_normalize(b)
            if len(norm_b) > 5:
                trash_blocks.add(norm_b)

    lines = raw_text.strip().split('\n')
    
    if lines and re.search(r'(?i)(RESSON[AÂ]NCIA|RM |TOMOGRAFIA|TC )', lines[0]):
        raw_text = '\n'.join(lines[1:]).strip()
        
    input_blocks = re.split(r'\n\s*\n', raw_text)
    kept_blocks = []
    
    mask_blocks_original = []
    for mask_section in re.split(r'@@\s*', HIDDEN_MASKS):
        if not mask_section.strip(): 
            continue
        for b in re.split(r'\n\s*\n', mask_section.strip()):
            if len(super_normalize(b)) > 5:
                mask_blocks_original.append(b.strip())
                
    mask_blocks_original.sort(key=len, reverse=True)
    trash_blocks_norm = {super_normalize(b) for b in mask_blocks_original}

    def create_regex_from_block(block):
        words = [re.escape(w) for w in block.split()]
        return r'\s+'.join(words)

    for b in input_blocks:
        # BUG FIX: Ignorar blocos que contêm apenas pontuação/reticências isoladas (como o "..." do meio da máscara)
        if not b.strip(' \t\n\r.:,;-…'):
            continue

        norm_b = super_normalize(b)
        
        if len(norm_b) <= 5: 
            kept_blocks.append(b)
            continue
      
        if norm_b in trash_blocks_norm:
            continue 
        
        modified_b = b
        for mask_block in mask_blocks_original:
            if len(mask_block.split()) < 3: continue 
            
            pattern = create_regex_from_block(mask_block)
            
            match_start = re.match(r'^\s*' + pattern + r'\s*', modified_b, re.IGNORECASE)
            if match_start:
                modified_b = modified_b[match_start.end():]
                continue
            
            match_end = re.search(r'\s*' + pattern + r'\s*$', modified_b, re.IGNORECASE)
            if match_end:
                modified_b = modified_b[:match_end.start()]
                
        if len(super_normalize(modified_b)) > 5:
            kept_blocks.append(modified_b.strip())
            
    filtered_text = '\n\n'.join(kept_blocks).strip()
    return filtered_text

def find_joint(raw_text, title_line=""):
    text_padded = " " + super_normalize(raw_text) + " "
    
    scores = {k: 0 for k in HIDDEN_DICTIONARY.keys()}
    
    for joint, keywords in HIDDEN_DICTIONARY.items():
        if isinstance(keywords, dict):
            for kw, weight in keywords.items():
                if re.search(r'\b' + re.escape(kw) + r'\b', text_padded):
                    scores[joint] += weight
        elif isinstance(keywords, list):
            for kw in keywords:
                if re.search(r'\b' + re.escape(kw) + r'\b', text_padded):
                    scores[joint] += 1

    if title_line:
        title_norm = super_normalize(title_line)
        title_words = title_norm.split()
        
        if "quadril" in title_words or "hemibacia" in title_words: scores["QUADRIL"] = scores.get("QUADRIL", 0) + 18
        if "bacia" in title_words or "pelve" in title_words: scores["BACIA"] = scores.get("BACIA", 0) + 14
        if "sacro" in title_words or "coccix" in title_words or "sacroiliaca" in title_words or "sacroiliacas" in title_words or "sacrococcigea" in title_words: scores["SACRO CÓCCIX SACROILÍACAS"] = scores.get("SACRO CÓCCIX SACROILÍACAS", 0) + 12
        if "mao" in title_words or "dedo" in title_words or "quirodactilo" in title_words or "quirodactilos" in title_words: scores["MÃO"] = scores.get("MÃO", 0) + 15
        if "punho" in title_words: scores["PUNHO"] = scores.get("PUNHO", 0) + 15
        if "pe" in title_words or "antepe" in title_words or "pododactilo" in title_words or "pododactilos" in title_words: scores["ANTEPÉ"] = scores.get("ANTEPÉ", 0) + 15
        if "tornozelo" in title_words: scores["TORNOZELO"] = scores.get("TORNOZELO", 0) + 15
        if "generica" in title_words or "coxa" in title_words or "perna" in title_words or "braco" in title_words or "antebraco" in title_words or "membros" in title_words: scores["MEMBROS"] = scores.get("MEMBROS", 0) + 15
        
        if "joelho" in title_words: scores["JOELHO"] = scores.get("JOELHO", 0) + 15
        if "cotovelo" in title_words: scores["COTOVELO"] = scores.get("COTOVELO", 0) + 15
        if "ombro" in title_words or "clavicula" in title_words or "escapula" in title_words or "plexo" in title_words: scores["OMBRO"] = scores.get("OMBRO", 0) + 15
        if "cervical" in title_words or "pescoco" in title_words: scores["COLUNA CERVICAL"] = scores.get("COLUNA CERVICAL", 0) + 15
        if "dorsal" in title_words or "toracica" in title_words: scores["COLUNA DORSAL"] = scores.get("COLUNA DORSAL", 0) + 15
        if "lombar" in title_words or "lombossacra" in title_words: scores["COLUNA LOMBAR"] = scores.get("COLUNA LOMBAR", 0) + 15

    max_hits = max(scores.values()) if scores else 0
    
    if max_hits == 0:
        return None
        
    best_joints = [j for j, s in scores.items() if s == max_hits]
    best_joint = best_joints[0] if best_joints else None
            
    return best_joint

def get_pacs_window():
    hwnds = []
    win32gui.EnumWindows(lambda hwnd, acc: acc.append(hwnd) if win32gui.IsWindowVisible(hwnd) and "vue pacs" in win32gui.GetWindowText(hwnd).lower() else True, hwnds)
    return hwnds[0] if hwnds else None

def get_edge_window():
    hwnds = []
    win32gui.EnumWindows(lambda hwnd, acc: acc.append(hwnd) if win32gui.IsWindowVisible(hwnd) and "edge" in win32gui.GetWindowText(hwnd).lower() else True, hwnds)
    return hwnds[0] if hwnds else None

def reuse_or_open_edge(url):
    hwnd = get_edge_window()
    if hwnd:
        try:
            if win32gui.IsIconic(hwnd): 
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            robust_clipboard_copy_text(url)
            pyautogui.hotkey('ctrl', 'l')
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.1)
            pyautogui.press('enter')
            return
        except: 
            pass
            
    paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", 
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe", 
        os.path.expanduser(r"~\AppData\Local\Microsoft\Edge\Application\msedge.exe")
    ]
    for p in paths:
        if os.path.exists(p):
            try: 
                return subprocess.Popen([p, url])
            except: 
                pass
                
    webbrowser.open_new_tab(url)

# =========================================================================
# PROTEÇÃO BLINDADA DO CLIPBOARD (Threading Lock)
# =========================================================================
CLIPBOARD_LOCK = threading.Lock()

def robust_clipboard_copy_text(text, retries=15):
    text = str(text).replace('\x00', '')
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
                    return True
                finally:
                    win32clipboard.CloseClipboard()
            except: 
                time.sleep(0.1)
        return False

def robust_clipboard_paste_text(retries=15):
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try:
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    else:
                        data = ""
                    return data
                finally:
                    win32clipboard.CloseClipboard()
            except: 
                time.sleep(0.1)
        return ""

def clear_clipboard_safely(retries=15):
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try:
                    win32clipboard.EmptyClipboard()
                finally:
                    win32clipboard.CloseClipboard()
                return
            except: 
                time.sleep(0.1)

def extract_html_from_clipboard(retries=15):
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try:
                    fmt = win32clipboard.RegisterClipboardFormat("HTML Format")
                    if win32clipboard.IsClipboardFormatAvailable(fmt):
                        html_data = win32clipboard.GetClipboardData(fmt).decode('utf-8', errors='ignore')
                    else:
                        html_data = None
                    return html_data
                finally:
                    win32clipboard.CloseClipboard()
            except: 
                time.sleep(0.1)
        return None

def inject_dual_format_clipboard(html_fragment, plain_text_fallback, retries=15):
    p_bytes = b"<html><body>"
    h_bytes = html_fragment.encode('utf-8')
    s_bytes = b"</body></html>"
    sh = 105
    sf = 105 + len(p_bytes)
    ef = 105 + len(p_bytes) + len(h_bytes)
    eh = 105 + len(p_bytes) + len(h_bytes) + len(s_bytes)
    
    payload = f"Version:0.9\r\nStartHTML:{sh:010d}\r\nEndHTML:{eh:010d}\r\nStartFragment:{sf:010d}\r\nEndFragment:{ef:010d}\r\n".encode('utf-8') + p_bytes + h_bytes + s_bytes
    
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat("HTML Format"), payload)
                    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text_fallback)
                    return
                finally:
                    win32clipboard.CloseClipboard()
            except: 
                time.sleep(0.1)

def sanitize_and_clean_html(raw_html):
    validos = [m for m in re.findall(r'\[\[INI\]\](.*?)\[\[FIM\]\]', raw_html, re.I | re.S) if len(m.strip()) > 30]
    if not validos: return None, None
    best_match = max(validos, key=len)
    
    best_match = html.unescape(best_match)
    best_match = best_match.replace('\xa0', ' ').replace('\u200b', '')
    best_match = re.sub(r'(?i)LESAOMAIOR:?\s*', '', best_match)
    best_match = best_match.replace("****", "").replace("** **", "")
    
    soup = BeautifulSoup(best_match.replace('\n', ' ').replace('\r', ''), 'html.parser')
    for tag in soup.find_all(['style', 'script', 'meta', 'svg', 'button']): 
        tag.decompose()
        
    clean_html_str = str(soup)
    text_soup = BeautifulSoup(clean_html_str, 'html.parser')
    
    for br in text_soup.find_all('br'): 
        br.replace_with('\n')
    for p in text_soup.find_all(['p', 'div']): 
        p.insert_after('\n\n')
        p.unwrap()
    for li in text_soup.find_all('li'): 
        li.insert_before('• ')
        li.insert_after('\n')
        li.unwrap()
        
    plain_text = re.sub(r'\n{3,}', '\n\n', '\n'.join([re.sub(r' {2,}', ' ', l.strip()) for l in text_soup.get_text(separator=" ").split('\n')])).strip()
    
    clean_html_str = re.sub(r':(\s*</(?:b|strong|i|em|span)>)?\s*([A-Za-zÀ-ÿ])', r':\1 \2', clean_html_str)
    plain_text = re.sub(r':\s*([A-Za-zÀ-ÿ])', r': \1', plain_text)

    return clean_html_str, plain_text

class LaudoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        clear_clipboard_safely(retries=2)
        
        self.title("AeroLumia")
        self.geometry("1060x700") 
        self.resizable(False, False)
        
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.abort_flag = False
        self.mouse_jail_active = False
        self.jail_x = 0
        self.jail_y = 0
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
            self.hamster_image = ctk.CTkImage(light_image=Image.open(img_path), dark_image=Image.open(img_path), size=(90, 90))
            self.lbl_hamster = ctk.CTkLabel(self.header_left, text="", image=self.hamster_image)
        else: 
            self.lbl_hamster = ctk.CTkLabel(self.header_left, text="[IMG FALTANDO]", font=("Arial", 12))
            
        self.lbl_hamster.pack(side="left", anchor="center", padx=(0, 15))
        
        self.lbl_title = ctk.CTkLabel(self.header_left, text="INPUT - Dados Brutos", font=("Arial", 16, "bold"))
        self.lbl_title.pack(side="left", anchor="center")

        if os.path.exists(img_path_perfil := os.path.join(RESOURCE_DIR, 'proj_IA_output.png')):
            self.hamster_perfil = ctk.CTkImage(light_image=Image.open(img_path_perfil), dark_image=Image.open(img_path_perfil), size=(96, 96))
            self.lbl_hamster_direita = ctk.CTkLabel(self.header_right, text="", image=self.hamster_perfil)
        else: 
            self.lbl_hamster_direita = ctk.CTkLabel(self.header_right, text="[IMG FALTANDO]", font=("Arial", 12))
            
        self.lbl_hamster_direita.pack(side="left", anchor="center", padx=(0, 15), pady=(12, 0))
        
        self.lbl_out_title = ctk.CTkLabel(self.header_right, text="OUTPUT - Pré-visualização", font=("Arial", 16, "bold"))
        self.lbl_out_title.pack(side="left", anchor="center", pady=(3, 0))

        self.control_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.control_frame.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))
     
        self.lbl_articulacao = ctk.CTkLabel(self.control_frame, text="Artic.:", font=("Arial", 12, "bold"))
        self.lbl_articulacao.pack(side="left", padx=(0, 5))
        
        self.lista_articulacoes_rm = [
            "AUTO", "ANTEPÉ", "BACIA", "COLUNA CERVICAL", "COLUNA DORSAL", 
            "COLUNA LOMBAR", "COTOVELO", "JOELHO", "MÃO", "MEMBROS", 
            "OMBRO", "PUNHO", "QUADRIL", "SACRO CÓCCIX SACROILÍACAS", "TORNOZELO"
        ]
        self.lista_articulacoes_tc = self.lista_articulacoes_rm + ["MSK (Genérico)"]
        
        self.combo_articulacao = ctk.CTkComboBox(self.control_frame, values=self.lista_articulacoes_rm, width=130)
        self.combo_articulacao.pack(side="left", padx=(0, 15))
        self.combo_articulacao.set("AUTO")
        
        self.seg_modality = ctk.CTkSegmentedButton(self.control_frame, values=["RM", "TC"], 
                                                   font=("Arial", 13, "bold"), height=30,
                                                   selected_color="#205c8a", selected_hover_color="#18476a",
                                                   command=self.on_modality_change)
        self.seg_modality.pack(side="left", padx=(0, 15))
        self.seg_modality.set("RM")
        
        self.chk_tradutor_var = ctk.IntVar(value=0)
        self.chk_tradutor = ctk.CTkCheckBox(self.control_frame, text="TRADUTOR (Inglês)", variable=self.chk_tradutor_var, font=("Arial", 12, "bold"))
        self.chk_tradutor.pack(side="left")

        self.chk_frases_var = ctk.IntVar(value=0)
        self.chk_frases = ctk.CTkCheckBox(self.control_frame, text="FRASES PADRÃO", variable=self.chk_frases_var, font=("Arial", 12, "bold"), command=self.on_frases_toggle)
        self.chk_frases.pack(side="left", padx=(15, 0))

        fonte_base = ctk.CTkFont(family="Arial", size=13)
        fonte_negrito = ctk.CTkFont(family="Arial", size=13, weight="bold")
        fonte_italico = ctk.CTkFont(family="Arial", size=13, slant="italic")
        
        self.txt_input = ctk.CTkTextbox(self.main_frame, wrap="word", width=0, font=fonte_base)
        self.txt_input.grid(row=2, column=0, sticky="nsew", padx=(0, 10))
        self.txt_input.focus_set()
        
        self.txt_output = ctk.CTkTextbox(self.main_frame, state="disabled", fg_color="#f5f5f5", wrap="word", cursor="arrow", width=0, font=fonte_base)
        self.txt_output.grid(row=2, column=1, sticky="nsew", padx=(10, 0))
        
        self.tk_text_core = self.txt_output._textbox
        self.tk_text_core.tag_configure("bold", font=fonte_negrito)
        self.tk_text_core.tag_configure("italic", font=fonte_italico)
        
        for ev in ["<Button-1>", "<B1-Motion>", "<Double-Button-1>", "<Triple-Button-1>", "<Control-c>", "<Control-x>", "<Control-a>"]: 
            self.tk_text_core.bind(ev, lambda e: "break")

        self.bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_frame.pack(fill="x", padx=20, pady=(10, 25)) 
        
        self.frame_botoes = ctk.CTkFrame(self.bottom_frame, fg_color="transparent")
        self.frame_botoes.pack(pady=10, anchor="center")
        
        self.btn_gerar_pacs = ctk.CTkButton(self.frame_botoes, text="GERAR LAUDO - RIS/PACS (Alt+Shift+Z)", width=350, height=50, font=("Arial", 14, "bold"), command=self.btn_gerar_pacs_click)
        self.btn_gerar_pacs.pack(side="left", padx=10) 
        
        self.btn_gerar_input = ctk.CTkButton(self.frame_botoes, text="GERAR LAUDO - CAIXA DE INPUT", width=250, height=40, font=("Arial", 13, "bold"), command=self.btn_gerar_input_click)
        self.btn_gerar_input.pack(side="left", padx=10)
        
        self.btn_novo = ctk.CTkButton(self.frame_botoes, text="RESET / ABORTAR", width=120, height=40, font=("Arial", 13, "bold"), command=self.trigger_global_abort, fg_color="#d9534f", hover_color="#c9302c")
        self.btn_novo.pack(side="left", padx=10)
        
        self.lbl_status = ctk.CTkLabel(self.bottom_frame, text="", font=("Arial", 12))
        self.lbl_status.pack(pady=(30, 5)) 
        self.selected_joint, self.block_hook = None, None
        
        try: 
            keyboard.add_hotkey('alt+shift+z', self.on_global_hotkey)
            keyboard.add_hotkey('ctrl+esc', self.trigger_global_abort)
            keyboard.add_hotkey('pause', self.trigger_global_abort)
            keyboard.add_hotkey('end', self.trigger_global_abort)
        except: 
            pass
        
        if DICT_LOAD_ERROR:
            self.update_status("ERRO NO JSON DO DICIONÁRIO (Verifique aspas e vírgulas). Usando emergência.", "red", size=13, bold=True)
        elif LOADED_FILES:
            self.update_status(f"Sistema Inicializado. Lidos: {', '.join(LOADED_FILES)}", "green")
        else:
            self.update_status("Sistema Inicializado (Usando prompts nativos. TXTs ausentes)", "green")

    def silent_cleanup(self):
        self.abort_flag = False
        self.selected_joint = None
        self.txt_output.configure(state="normal")
        self.txt_output.delete("1.0", "end")
        self.txt_output.configure(state="disabled")
        clear_clipboard_safely(retries=3)

    def auto_translate(self, texto, modo):
        if modo == "RM_PARA_TC":
            texto = re.sub(r'(?i)susceptibilidade magnética', 'endurecimento de feixe', texto)
            texto = re.sub(r'(?i)ressonância magnética|rm\b', 'tomografia computadorizada', texto)
            texto = re.sub(r'(?i)intensidade de sinal|sinal', 'atenuação', texto)
            texto = re.sub(r'(?i)hipersinal|sinal elevado', 'hiperatenuação', texto)
            texto = re.sub(r'(?i)hipossinal|sinal reduzido|baixo sinal', 'hipoatenuação', texto)
            texto = re.sub(r'(?i)isossinal', 'isoatenuação', texto)
            texto = re.sub(r'(?i)alteração de sinal', 'alteração de atenuação', texto)
            
            texto = re.sub(r'(?i)(aquisição de imagens( pesadas)? em|sequências( pesadas)? em).*?(?=\n|$)', '', texto)
            texto = re.sub(r'(?i)\b(stir|difusão|mapa adc|dp|neurografia)\b', '', texto)
            
            texto = re.sub(r'(?i)exposição óssea|afilamento condral', 'redução do espaço articular', texto)
            
            rm_words = r'\b(cartilagem|condropatia|lesão condral|fissura condral|erosão condral|labrum|lábio glenoidal|lábio acetabular|plica(s)?( sinovial| sinoviais)?|sinovite(s)?( reacional)?|menisco(s)?( medial| lateral)?|corno (anterior|posterior))\b'
            texto = re.sub(r'(?i)(,\s*|\.\s*|^|\n\s*)[^,\.\n]*?' + rm_words + r'[^,\.\n]*(?=[,\.\n]|$)', '', texto)
            
            edema_words = r'\b(edema\s+(ósseo|subcondral|da medular|endosteal))\b'
            texto = re.sub(r'(?i)(,\s*|\.\s*|^|\n\s*)[^,\.\n]*?' + edema_words + r'[^,\.\n]*(?=[,\.\n]|$)', '', texto)
            
            texto = re.sub(r'(?i)edema ósseo( contusional)?\s*(e\s*)?', '', texto)
            
            texto = re.sub(r'(?i)tenossinovite', 'pequena quantidade de líquido peritendíneo', texto)
            texto = re.sub(r'(?i)(edema|espessamento) periosteal', 'reação periosteal', texto)
            
            texto = re.sub(r'(?i).*?medula espinhal.*?\n', '', texto)
            texto = re.sub(r'(?i).*?cone medular.*?\n', '', texto)
            texto = re.sub(r'(?i).*?cauda equina.*?\n', '', texto)

            texto = re.sub(r'(?i)e atenuação preservad[oa]s?', 'preservados', texto)
            texto = re.sub(r'(?i)e atenuação normais', 'normais', texto)

            texto = re.sub(r',\s*,', ',', texto)
            texto = re.sub(r'\s+,\s+', ', ', texto)
            texto = re.sub(r'^\s*,\s*', '', texto) 
            texto = re.sub(r'\n\s*,\s*', '\n', texto) 
            texto = re.sub(r'\.\s*\.', '.', texto)
            texto = re.sub(r',\s*\.', '.', texto)

        elif modo == "TC_PARA_RM":
            texto = re.sub(r'(?i)endurecimento de feixe', 'susceptibilidade magnética', texto)
            texto = re.sub(r'(?i)tomografia computadorizada|tc\b', 'ressonância magnética', texto)
            texto = re.sub(r'(?i)atenuação', 'sinal', texto)
            texto = re.sub(r'(?i)hiperatenuação|hiperatenuante|hiperdenso', 'hipersinal', texto)
            texto = re.sub(r'(?i)hipoatenuação|hipoatenuante|hipodenso', 'hipossinal', texto)
            texto = re.sub(r'(?i)isoatenuação|isoatenuante|isodenso', 'isossinal', texto)
            
            texto = re.sub(r'(?i)(,\s*)?medindo\s+.*?(uh|unidades hounsfield)(?=[,\.\n])', '', texto)
            texto = re.sub(r'(?i)\b(uh|hounsfield)\b', '', texto)
            texto = re.sub(r'(?i)(,\s*)?(com\s+|associad[oa]s?\s+a\s+)?(gás|degeneração gasosa|fenômeno do vácuo|vácuo)( discal| intra-articular)?(?=[,\.\n])', '', texto)
            
        texto = re.sub(r'\n{3,}', '\n\n', texto)
        texto = re.sub(r' {2,}', ' ', texto)
        return texto.strip()

    def on_modality_change(self, value):
        current_val = self.combo_articulacao.get()
        if value == "TC":
            self.combo_articulacao.configure(values=self.lista_articulacoes_tc)
        else:
            self.combo_articulacao.configure(values=self.lista_articulacoes_rm)
            if current_val == "MSK (Genérico)":
                self.combo_articulacao.set("AUTO")

    def on_frases_toggle(self):
        if self.chk_frases_var.get() == 1:
            warn_win = ctk.CTkToplevel(self)
            warn_win.title("Aviso")
            warn_win.geometry("450x150")
            warn_win.attributes('-topmost', True)
            warn_win.grab_set() 
       
            lbl = ctk.CTkLabel(warn_win, text="Ocasionalmente o Copilot pode apresentar erros\ncom prompts longos.", font=("Arial", 14, "bold"), text_color="#d9534f")
            lbl.pack(expand=True, pady=(20, 10))
            btn = ctk.CTkButton(warn_win, text="OK", font=("Arial", 14, "bold"), command=warn_win.destroy, width=120)
            btn.pack(pady=(0, 20))

    def _mouse_jail_loop(self):
        while True:
            if self.mouse_jail_active and not self.abort_flag:
                ctypes.windll.user32.SetCursorPos(int(self.jail_x), int(self.jail_y))
            time.sleep(0.01)

    def enable_idiot_proof_shield(self):
        try:
            gui_info = GUITHREADINFO()
            gui_info.cbSize = ctypes.sizeof(GUITHREADINFO)
            hwnd_active = ctypes.windll.user32.GetForegroundWindow()
            thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd_active, None)
            moved_to_caret = False
            
            if ctypes.windll.user32.GetGUIThreadInfo(thread_id, ctypes.byref(gui_info)) and gui_info.hwndCaret:
                pt = wintypes.POINT(gui_info.rcCaret.left, gui_info.rcCaret.top)
                ctypes.windll.user32.ClientToScreen(gui_info.hwndCaret, ctypes.byref(pt))
                self.jail_x, self.jail_y = pt.x, pt.y
                moved_to_caret = True
 
            if not moved_to_caret and hwnd_active:
                rect = win32gui.GetWindowRect(hwnd_active)
                self.jail_x = rect[2] - 150
                self.jail_y = rect[1] + 250
        except: 
            self.jail_x, self.jail_y = pyautogui.position()
        
        self.mouse_jail_active = True
        self.after(0, self._show_ghost_overlay)

    def disable_idiot_proof_shield(self):
        self.mouse_jail_active = False
        self.after(0, self._hide_ghost_overlay)

    def _show_ghost_overlay(self):
        if self.overlay_window: return
        self.overlay_window = ctk.CTkToplevel(self)
        self.overlay_window.overrideredirect(True)
        self.overlay_window.attributes('-topmost', True)
        
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        y_pos = screen_h - 80 
        
        self.overlay_window.geometry(f"{screen_w}x80+0+{y_pos}")
        self.overlay_window.configure(fg_color="#8B0000")
        
        lbl = ctk.CTkLabel(self.overlay_window, text="⚠️ ROBÔ TRABALHANDO - NÃO MEXA NO MOUSE OU TECLADO (Pressione PAUSE ou END para Abortar) ⚠️", font=("Arial", 22, "bold"), text_color="white")
        lbl.pack(expand=True, fill="both")
        
        hwnd = win32gui.GetParent(self.overlay_window.winfo_id())
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED)
        win32gui.SetLayeredWindowAttributes(hwnd, 0, 230, win32con.LWA_ALPHA) 

    def _hide_ghost_overlay(self):
        if self.overlay_window: 
            self.overlay_window.destroy()
            self.overlay_window = None

    def sleep_safe(self, seconds):
        end_time = time.time() + seconds
        while time.time() < end_time:
            if self.abort_flag: 
                return False
            time.sleep(0.05)
        return True

    def trigger_global_abort(self):
        self.abort_flag = True
        self.disable_idiot_proof_shield()
        release_stuck_modifiers()
        self.after(0, self.reset_ui_completely)

    def reset_ui_completely(self):
        self.txt_input.delete("1.0", "end")
        self.txt_output.configure(state="normal")
        self.txt_output.delete("1.0", "end")
        self.txt_output.configure(state="disabled")
        self.combo_articulacao.set("AUTO")
        self.change_ui_state("normal")
        self.selected_joint = None
        self.txt_input.focus_set()
        self.update_status("SISTEMA ABORTADO/RESETADO PELO USUÁRIO.", "red", 14, True)

    def show_login_error_ui(self):
        self.abort_flag = True
        self.disable_idiot_proof_shield()
        release_stuck_modifiers()
        self.reset_ui_completely()
        self.lift()
        self.attributes('-topmost', True)
        self.after(500, lambda: self.attributes('-topmost', False))
        self.focus_force()
        try: 
            winsound.MessageBeep(winsound.MB_ICONHAND) 
        except: 
            pass
        
        err_win = ctk.CTkToplevel(self)
        err_win.title("Atenção: Login Necessário")
        err_win.geometry("650x300")
        err_win.attributes('-topmost', True)
        err_win.grab_set() 
        new_text = "O Copilot bloqueou o acesso pois você não está logado!\n\nFaça login no Copilot no EDGE (NÃO USE \"Copilot 365\").\n\nNa sequência digite \"teste\" no Copilot e aceite os cookies.\n\nEntão confirme que é humano."
        lbl = ctk.CTkLabel(err_win, text=new_text, font=("Arial", 18, "bold"), text_color="#B22222")
        lbl.pack(expand=True, pady=(20, 10))
        
        def on_login_ok():
            err_win.destroy()
            hwnd = get_edge_window()
            if hwnd:
                try:
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.3)
                    rect = win32gui.GetWindowRect(hwnd)
                    safe_x = rect[2] - 150
                    safe_y = rect[1] + 250
                    pyautogui.click(safe_x, safe_y)
                except: pass
                
        btn = ctk.CTkButton(err_win, text="OK, ENTENDI", font=("Arial", 16, "bold"), command=on_login_ok, width=200, height=45)
        btn.pack(pady=(0, 20))

    def show_cookie_error_ui(self):
        self.abort_flag = True
        self.disable_idiot_proof_shield()
        release_stuck_modifiers()
        self.reset_ui_completely()
        self.lift()
        self.attributes('-topmost', True)
        self.after(500, lambda: self.attributes('-topmost', False))
        self.focus_force()
        try: 
            winsound.MessageBeep(winsound.MB_ICONHAND) 
        except: 
            pass
        
        err_win = ctk.CTkToplevel(self)
        err_win.title("Atenção: Cookies")
        err_win.geometry("450x150")
        err_win.attributes('-topmost', True)
        err_win.grab_set() 
        lbl = ctk.CTkLabel(err_win, text="Por favor aceite os cookies", font=("Arial", 18, "bold"), text_color="#B22222")
        lbl.pack(expand=True, pady=(20, 10))
        
        def on_cookie_ok():
            err_win.destroy()
            hwnd = get_edge_window()
            if hwnd:
                try:
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.3)
                    rect = win32gui.GetWindowRect(hwnd)
                    safe_x = rect[2] - 150
                    safe_y = rect[1] + 250
                    pyautogui.click(safe_x, safe_y)
                except: pass
                
        btn = ctk.CTkButton(err_win, text="OK", font=("Arial", 16, "bold"), command=on_cookie_ok, width=150, height=40)
        btn.pack(pady=(0, 20))

    def on_global_hotkey(self):
        if self.btn_gerar_pacs.cget("state") == "normal": 
            self.after(0, self.btn_gerar_pacs_click)

    def apply_msk_style_visually(self):
        self.update_status("[ VISÃO ] Escaneando área para aplicar estilo...", "orange")
        if not self.sleep_safe(0.1): 
            return False
            
        anchor1 = os.path.join(RESOURCE_DIR, 'ancorabarra1.png')
        anchor2 = os.path.join(RESOURCE_DIR, 'ancorabarra2.png')
        if not os.path.exists(anchor1) or not os.path.exists(anchor2): 
            return False
            
        try:
            vscreen_x = ctypes.windll.user32.GetSystemMetrics(76)
            vscreen_y = ctypes.windll.user32.GetSystemMetrics(77)
            try: 
                screen_img = ImageGrab.grab(all_screens=True)
            except TypeError: 
                screen_img = ImageGrab.grab()
                
            if self.abort_flag: 
                return False
                
            box1 = pyautogui.locate(anchor1, screen_img, confidence=0.7, grayscale=True)
            box2 = pyautogui.locate(anchor2, screen_img, confidence=0.7, grayscale=True)
            
            if box1 and box2:
                click_x = box1.left + (box1.width // 2) + vscreen_x + int(((box2.left + (box2.width // 2) + vscreen_x) - (box1.left + (box1.width // 2) + vscreen_x)) * 0.34)
                click_y = box1.top + (box1.height // 2) + vscreen_y
                pyautogui.click(x=click_x, y=click_y) 
                
                if not self.sleep_safe(0.3): return False
                robust_clipboard_copy_text('MSK')
                pyautogui.hotkey('ctrl', 'v')
                if not self.sleep_safe(0.2): return False
                hardware_enter()
                return True
        except: 
            pass
        return False

    def change_ui_state(self, state):
        self.btn_gerar_pacs.configure(state=state)
        self.btn_gerar_input.configure(state=state)
        self.combo_articulacao.configure(state=state)
        self.seg_modality.configure(state=state)
        self.chk_tradutor.configure(state=state)
        self.chk_frases.configure(state=state)

    def finaliza_rotina(self, msg, color, success=True):
        self.disable_idiot_proof_shield()
        release_stuck_modifiers()
        if self.abort_flag: 
            return 
            
        if success:
            self.seg_modality.set("RM")
            self.on_modality_change("RM")
            self.chk_tradutor_var.set(0)
            self.chk_frases_var.set(0)
            
        self.change_ui_state("normal")
        self.combo_articulacao.set("AUTO")
        self.update_status(msg, color, 24, True)

    def update_status(self, text, color="black", size=12, bold=False):
        self.lbl_status.configure(text=text, text_color=color, font=("Arial", size, "bold") if bold else ("Arial", size))
        self.update_idletasks()

    def render_html_to_output(self, html_string):
        self.txt_output.configure(state="normal")
        self.txt_output.delete("1.0", "end")
        
        def nl(chars="\n\n"):
            txt = self.tk_text_core.get("1.0", "end-1c")
            if not txt: return
            if chars == "\n\n":
                if not txt.endswith("\n\n"):
                    self.tk_text_core.insert("end", "\n" if txt.endswith("\n") else "\n\n")
            elif chars == "\n":
                if not txt.endswith("\n"):
                    self.tk_text_core.insert("end", "\n")

        def walk_tree(node, active_tags):
            if isinstance(node, NavigableString):
                text = str(node).replace('\n', ' ').replace('\r', '').replace('\t', ' ')
                text = re.sub(r' {2,}', ' ', text)
                current_content = self.tk_text_core.get("1.0", "end-1c")
                
                if text.startswith(' ') and (not current_content or current_content.endswith('\n')):
                    text = text[1:]
                    
                if text:
                    self.tk_text_core.insert("end", text, tuple(active_tags))
            else:
                current_tags = active_tags.copy()
                if node.name in ['b', 'strong'] or (node.get('style') and 'font-weight' in node.get('style')): 
                    current_tags.append("bold")
                if node.name in ['i', 'em']: 
                    current_tags.append("italic")
                
                if node.name in ['p', 'div']: 
                    nl("\n\n")
                elif node.name == 'li': 
                    nl("\n")
                    self.tk_text_core.insert("end", "• ", tuple(current_tags))
                elif node.name == 'br': 
                    nl("\n")
                
                for child in node: 
                    walk_tree(child, current_tags)
                
                if node.name in ['p', 'div']: 
                    nl("\n\n")
                elif node.name == 'li': 
                    nl("\n")

        walk_tree(BeautifulSoup(html_string, 'html.parser'), [])
        
        content = self.tk_text_core.get("1.0", "end-1c")
        l_white = len(content) - len(content.lstrip())
        if l_white > 0:
            self.tk_text_core.delete("1.0", f"1.0+{l_white}c")
            
        self.txt_output.configure(state="disabled")

    def request_manual_joint(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Definição Manual")
        dialog.geometry("300x200")
        dialog.attributes("-topmost", True)
        dialog.grab_set() 
        ctk.CTkLabel(dialog, text="Articulação não identificada.\nSelecione abaixo:").pack(pady=15)
        
        values = ["SELECIONE..."] + sorted(list(HIDDEN_DICTIONARY.keys()))
        combo = ctk.CTkComboBox(dialog, values=values)
        combo.set("SELECIONE...")
        combo.pack(pady=10)
        
        def confirm(): 
            if combo.get() != "SELECIONE...":
                self.selected_joint = combo.get()
                dialog.destroy()
            
        ctk.CTkButton(dialog, text="Confirmar", command=confirm).pack(pady=15)
        self.wait_window(dialog) 

    def extract_prompt_block(self, joint_name, modality):
        if not joint_name: 
            return ""
        if joint_name == "MSK (Genérico)":
            joint_name = "MSK"
            
        joint_norm = super_normalize(joint_name)
        if joint_norm == "sacro coccix sacroiliacas":
            joint_keywords = ["sacro", "coccix", "sacroiliaca"]
        else:
            joint_keywords = [joint_norm]

        mod_keywords = ["rm", "ressonancia"] if modality == "RM" else ["tc", "tomografia"]
        old_style = super_normalize(f"{joint_name} {modality}")

        for block in re.split(r'@@', HIDDEN_PROMPT):
            if not block.strip(): 
                continue
            first_line = block.strip().split('\n')[0]
            header = super_normalize(first_line)
            
            has_mod = any(m in header for m in mod_keywords)
            has_joint = any(j in header for j in joint_keywords)
            
            if (has_mod and has_joint) or header.startswith(old_style):
                return "@@" + block
        return ""
        
    def extract_frases_block(self, joint_name, modality):
        if not joint_name or not HIDDEN_FRASES: 
            return ""
        if joint_name == "MSK (Genérico)":
            joint_name = "MSK"
            
        joint_norm = super_normalize(joint_name)
        if joint_norm == "sacro coccix sacroiliacas":
            joint_keywords = ["sacro", "coccix", "sacroiliaca"]
        else:
            joint_keywords = [joint_norm]

        mod_keywords = ["rm", "ressonancia"] if modality == "RM" else ["tc", "tomografia"]
        old_style = super_normalize(f"{joint_name} {modality}")
        
        for block in re.split(r'@\*@', HIDDEN_FRASES):
            if not block.strip(): 
                continue
            lines = block.strip().split('\n')
            first_line = lines[0]
            header = super_normalize(first_line)
            
            has_mod = any(m in header for m in mod_keywords)
            has_joint = any(j in header for j in joint_keywords)
            
            if (has_mod and has_joint) or header.startswith(old_style):
                return '\n'.join(lines[1:]).strip()
        return ""

    def btn_gerar_input_click(self):
        raw_text = self.txt_input.get("1.0", "end-1c").strip()
        if not raw_text: 
            return self.update_status("ERRO: Caixa vazia.", "red")
            
        self.silent_cleanup()
        self.txt_input.configure(state="normal")
        self.txt_input.delete("1.0", "end")
        self.txt_input.insert("1.0", raw_text)
        
        self.change_ui_state("disabled")
        self.enable_idiot_proof_shield()
        self.process_raw_text_and_run(raw_text)

    def btn_gerar_pacs_click(self):
        self.silent_cleanup()
        self.change_ui_state("disabled")
        self.update_status("[ A EXTRAIR DADOS DO VUE PACS... ]", "blue")
        threading.Thread(target=self.fetch_from_pacs_and_run, daemon=True).start()

    def fetch_from_pacs_and_run(self):
        if not self.sleep_safe(0.4): return
        hwnd = get_pacs_window()
        if not hwnd: 
            return self.after(0, lambda: self.finaliza_rotina("ERRO: Vue PACS não encontrado.", "red", False))
            
        try:
            pyautogui.press('alt')
            time.sleep(0.05)
            
            if win32gui.IsIconic(hwnd): 
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.6)
        except: 
            pass 
            
        release_stuck_modifiers()
        
        try:
            pyautogui.hotkey('ctrl', 'shift', 't')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'shift', 't')
            time.sleep(0.3)
            self.after(0, self.enable_idiot_proof_shield)
            if not self.sleep_safe(0.1): return
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            pyautogui.press('right')
        except: 
            release_stuck_modifiers()
            
        if self.abort_flag: return
        raw_text = robust_clipboard_paste_text()
        
        if not raw_text or len(raw_text.strip()) < 5: 
            return self.after(0, lambda: self.finaliza_rotina("ERRO: Texto vazio.", "red", False))
            
        self.after(0, lambda: self.process_raw_text_and_run(raw_text.strip(), update_ui=True))

    def process_raw_text_and_run(self, raw_text, update_ui=False):
        if self.abort_flag: return
        self.update_status("[ FILTRO DE MÁSCARAS E PREPARAÇÃO... ]", "blue")
        
        title_line = ""
        first_line = raw_text.strip().split('\n')[0] if raw_text.strip() else ""
        if re.search(r'(?i)(RESSON[AÂ]NCIA|RM |TOMOGRAFIA|TC )', first_line):
            title_line = first_line
        
        input_modality = "RM"
        if re.search(r'(?i)(tomografia|tc\b|hounsfield|uh\b|atenuação|atenuante)', raw_text):
            input_modality = "TC"
        elif re.search(r'(?i)(ressonância|rm\b|t1\b|t2\b|stir\b|sinal\b)', raw_text):
            input_modality = "RM"
            
        output_modality = self.seg_modality.get()
        
        if input_modality == "RM" and output_modality == "TC":
            raw_text = self.auto_translate(raw_text, "RM_PARA_TC")
        elif input_modality == "TC" and output_modality == "RM":
            raw_text = self.auto_translate(raw_text, "TC_PARA_RM")

        clean_input = apply_mask_filter(raw_text)
        
        if update_ui: 
            self.after(0, lambda: self.txt_input.delete("1.0", "end"))
            self.after(0, lambda: self.txt_input.insert("1.0", clean_input))
            
        combo_val = self.combo_articulacao.get()
        if combo_val != "AUTO":
            self.selected_joint = combo_val
        else:
            self.selected_joint = find_joint(clean_input, title_line=title_line)
            
        if not self.selected_joint:
            self.disable_idiot_proof_shield()
            self.request_manual_joint()
            self.enable_idiot_proof_shield()
            
        if not self.selected_joint or self.abort_flag: 
            return self.after(0, lambda: self.finaliza_rotina("Operação cancelada.", "red", False))
        
        prompt_block = self.extract_prompt_block(self.selected_joint, output_modality)
        
        if not prompt_block and output_modality == "TC":
            prompt_block = self.extract_prompt_block("MSK", "TC")
        
        if self.chk_frases_var.get() == 1:
            frases_block = self.extract_frases_block(self.selected_joint, output_modality)
            
            if not frases_block and output_modality == "TC":
                frases_block = self.extract_frases_block("MSK", "TC")
                
            prompt_block = prompt_block.replace("<@insfra@>", frases_block)
            prompt_block = re.sub(r'<frasein[ií]cio>|<frasefim>', '', prompt_block, flags=re.IGNORECASE)
        else:
            prompt_block = prompt_block.replace("<@insfra@>", "")
            prompt_block = re.sub(r'<frasein[ií]cio>.*?<frasefim>', '', prompt_block, flags=re.IGNORECASE | re.DOTALL)

        base_instruction = "REGRA FINAL: O laudo DEVE ser gerado entre as tags [[INI]] e [[FIM]]. Exemplo: [[INI]] texto do laudo [[FIM]]."
        
        payload = f"{prompt_block}\n\n[DADOS BRUTOS]:\n{clean_input}\n\n{base_instruction}"
        
        if self.chk_tradutor_var.get() == 1: 
            payload += "\n\n[INSTRUÇÃO DE TRADUÇÃO ABSOLUTA]: TRADUZA o laudo final INTEIRAMENTE para o INGLÊS. Isso inclui a tradução OBRIGATÓRIA de TODOS os títulos de tópicos (ex: 'Corpos vertebrais' -> 'Vertebral bodies', 'Nível' -> 'Level') e frases de fechamento exigidas nas regras. Mantenha RIGOROSAMENTE todas as formatações, negritos, tags HTML e a estrutura exigida (incluindo as tags [[INI]] e [[FIM]])."
            
        robust_clipboard_copy_text(payload)
        threading.Thread(target=self.run_rpa_automation, args=(payload,), daemon=True).start()

    def run_rpa_automation(self, payload):
        if self.abort_flag: return
        
        self.after(0, lambda: self.update_status("[ REDIRECIONANDO EDGE... ]", "orange"))
        reuse_or_open_edge('https://copilot.microsoft.com/')
        
        if not self.sleep_safe(1.4): return
        
        if not getattr(self, 'first_run_done', False):
            if not self.sleep_safe(1.0): return 
            self.first_run_done = True
            
        robust_clipboard_copy_text(payload)
        start_time = time.time()
        timeout = 45.0
        sucesso = False
        precisa_login = False
        precisa_cookies = False
        
        try:
            if self.abort_flag: return
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(1.0)
            
            while (time.time() - start_time) < timeout:
                if self.abort_flag or not self.sleep_safe(1.5): return
                
                edge_hwnd = get_edge_window()
                if edge_hwnd:
                    rect = win32gui.GetWindowRect(edge_hwnd)
                    self.jail_x = rect[2] - 150
                    self.jail_y = rect[1] + 250
                
                pyautogui.scroll(-500) 
                
                if self.abort_flag: return
                
                pyautogui.click(self.jail_x, self.jail_y) 
                time.sleep(0.1)
                
                clear_clipboard_safely()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.1)
                
                copiado = robust_clipboard_paste_text()
                
                if (time.time() - start_time) > 8.0 and "REGRA FINAL" not in copiado:
                    precisa_login = True
                    break
                
                if re.search(r'(?i)entrar no\s*copilot', copiado) or re.search(r'(?i)entrar com a conta', copiado):
                    precisa_login = True
                    break

                if re.search(r'(?i)Valorizamos a sua privacidade', copiado):
                    precisa_cookies = True
                    break
                    
                if [m for m in re.findall(r'\[\[INI\]\](.*?)\[\[FIM\]\]', copiado, re.I | re.S) if len(m.strip()) > 30]: 
                    sucesso = True
                    break
        except: 
            release_stuck_modifiers()
        
        if self.abort_flag: return
        
        if precisa_login: 
            self.after(0, self.show_login_error_ui)
        elif precisa_cookies: 
            self.after(0, self.show_cookie_error_ui)
        elif sucesso: 
            self.after(0, self.process_final_html)
        else: 
            self.after(0, lambda: self.finaliza_rotina(f"ERRO: Timeout ({int(timeout)}s).", "red", False))

    def process_final_html(self):
        if self.abort_flag: return
        self.update_status("[ FORMATANDO HTML... ]", "blue")
        clean_html, plain_text = sanitize_and_clean_html(extract_html_from_clipboard() or "")
        
        if not clean_html:
            validos = [m for m in re.findall(r'\[\[INI\]\](.*?)\[\[FIM\]\]', robust_clipboard_paste_text(), re.I | re.S) if len(m.strip()) > 30]
            if validos:
                best_match = max(validos, key=len).strip()
                
                best_match = html.unescape(best_match)
                best_match = best_match.replace('\xa0', ' ').replace('\u200b', '')
                best_match = re.sub(r'(?i)LESAOMAIOR:?\s*', '', best_match)
                best_match = best_match.replace("****", "").replace("** **", "")
                best_match = re.sub(r':\s*([A-Za-zÀ-ÿ])', r': \1', best_match)
                
                clean_html = plain_text = best_match
               
        if clean_html and plain_text: 
            inject_dual_format_clipboard(clean_html, plain_text)
            self.render_html_to_output(clean_html)
            threading.Thread(target=self.paste_and_select_in_ris, args=(plain_text,), daemon=True).start()
        else: 
            self.after(0, lambda: self.finaliza_rotina("ERRO: Falha no conteúdo.", "red", False))

    def paste_and_select_in_ris(self, plain_text):
        if self.abort_flag: return
        self.after(0, lambda: self.update_status("[ COLANDO NO PACS... ]", "orange"))
        
        if not (hwnd := get_pacs_window()): 
            return self.after(0, self.finalize_ui_success_manual)
            
        try:
            pyautogui.press('alt')
            time.sleep(0.05)
            
            if win32gui.IsIconic(hwnd): 
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.6)
        except: 
            pass 
            
        if not self.sleep_safe(0.5): return
        release_stuck_modifiers()
        
        try:
            if self.abort_flag: return
            pyautogui.hotkey('ctrl', 'shift', 't')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'shift', 't')
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'end')
            time.sleep(0.1)
            pyautogui.press('enter')
            pyautogui.press('enter')
            time.sleep(0.1)
            
            pyautogui.keyDown('ctrl')
            time.sleep(0.05)
            pyautogui.keyDown('shift')
            time.sleep(0.05)
            pyautogui.press('v')
            time.sleep(0.05)
            pyautogui.keyUp('shift')
            pyautogui.keyUp('ctrl')
            
            time.sleep(0.5)
            self.after(0, lambda: self.update_status("[ REPOSICIONANDO CURSOR E SELECIONANDO... ]", "orange"))
            
            chars_to_select = len(plain_text.replace('\n\n', '\n')) + 1
            send_fast_left_arrows(chars_to_select)
            
            if not self.sleep_safe(0.4): return
            select_with_hardware_scancodes()
            
        except: 
            release_stuck_modifiers()
        
        if self.abort_flag: return
        
        self.mouse_jail_active = False 
        time.sleep(0.15) 
        
        estilo_ok = False
        if not self.abort_flag:
            estilo_ok = self.apply_msk_style_visually()
            
        if self.abort_flag: return
        self.disable_idiot_proof_shield()
        
        if estilo_ok and not self.abort_flag: 
            self.after(0, lambda: self.finaliza_rotina("LAUDO E ESTILO PRONTOS!", "green", True))
        elif not self.abort_flag: 
            self.after(0, lambda: self.finaliza_rotina("LAUDO PRONTO (Estilo manual necessário)", "black", True))

    def finalize_ui_success_manual(self):
        if self.abort_flag: return
        self.lift()
        self.attributes('-topmost', True)
        self.after(500, lambda: self.attributes('-topmost', False))
        self.finaliza_rotina("COLE MANUALMENTE", "orange", True)

if __name__ == "__main__":
    app = LaudoApp()
    app.mainloop()
