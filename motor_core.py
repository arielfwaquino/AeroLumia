import sys, os, ctypes, time, re, threading, subprocess, webbrowser, html, json, ast, unicodedata
import pyautogui, win32clipboard, win32gui, win32con
import base64, tempfile
from ctypes import wintypes
from bs4 import BeautifulSoup, NavigableString

try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except: pass

# =========================================================================
# IMAGENS EMBUTIDAS EM BASE64 (Âncoras de Visão Computacional)
# =========================================================================
B64_ANCORA1 = b"iVBORw0KGgoAAAANSUhEUgAAACcAAAAYCAIAAAD27XMaAAAAA3NCSVQICAjb4U/gAAAAwHpUWHRSYXcgcHJvZmlsZSB0eXBlIEFQUDEAABiVfU9BDgMhCLz7in3CAIr6HNPYZpOmbfb/h2LUdk2aDhGQADO4W33UY79sr+N53e/VbQ0MOJ995gIgoUMAYlCL5jtGFLLMhhDHn3vUnCL8qa9gQbAOFaVorw80l/seP2mFzWxHwS9uNJEnDo5T71qneUdY62VoYhTfbuBRT7N/riMl06oSLRMhUcmmK4pIsqiC/zZIOi9/7wtse9KorXfTZ87g3u0wVgUhsgm7AAACFklEQVRIiWP89OkTA90B458/f+hvKxP9rWRgYGD5//8//W0dGL+O2BB++/ZtQUGBlJQUGwYoLi4m0kSiDPmFBOzt7eGaLS0tf/36NX36dEsrq1+kAOyGWFoiq0GE8N27dw8ePIjs6v9QwER8LOAyBMKACxJOTYwMf4m0kniAsFVZWdnOzg6rImSvnju5Eo9xeAxBsfU/ElixYkVmZpawsDADeshAFezdMFVTTe4/XrBixYrMzExMQ5ABSggLCQn19/c9ffp0x44dyOKPn39mYGDYvaXfyt6I4cUC/P4QEhLq7+/HNATFrwRDg4GBYfPe61u37bC1seT/sODAhg3EaMEPmDBC6P///////fvPwMDw79//x6//RxSsOHTn32cmBYYHCzau2HD3lRb+EEYGDDgCGVfZBBH8L6PtIy3C35Bh/+EDw4QTBRduMKQnxeHQgmEEaob5T2TO2Xbi4+kr9xUU+A/cYNhwgeHBD4ULNwSMjYyIsRI/wFkOZ+ZN2nLoZkKE/YUHDAdOPPzw4cODExMkxIW4OVmJ9CsEIIcwXJDx69evmEpff/rTMPN8QoT9gxcMB048vHDjwocLGxQ0FOJC/MI81BkZSfQaBmD88uULMv/X77/HLr2bt/3xvQv7Pnx4wMAgwPDjg4SyjLGBeWuuPQsLdWpGxs+fP1PFINJsJT6S0FqTfHx8ZNvKQrxfGSmPTxgAAM8F4TCneRV7AAAAAElFTkSuQmCC"
B64_ANCORA2 = b"iVBORw0KGgoAAAANSUhEUgAAACwAAAAYCAIAAAAODYjtAAAAA3NCSVQICAjb4U/gAAAAwXpUWHRSYXcgcHJvZmlsZSB0eXBlIEFQUDEAABiVfU9bCgMhDPz3FHuEycOsHkfKtiyUtuz9PxpR2xVKJ5jEIckk4bY9tmO/LK/jed3vW1gqGAiaNXMBkNAgADGoRvcNPQp55k1Y+59btJxW6KmuYEL0ChOj1V9rqC63OTpkhd18RsEvbdQlTxp8GfvOPI074syXvhOjaL2BO59G/RinYhZNTSSKCYw9EyFnnbEo+G9dpOny977IkiV1br6bPn2O8AYUaFY1XvpEowAAAgdJREFUSInNlr+q4kAUxo+rmUQIiKRIKVqIFlY2ViIp7HwBQdTGVvAlLCwUxEYQrUQrxUJEkkZSWFsIPoCFFhJEcrNByRbjhtlcMfEfe39NMl/OmXxz5pCJyzAMALhcLnAXt9ttG/M0Hueh2O4n+PWheR/iJ5kw7LDELBYL2xTnPF8JiqLeU4cPbUepVEIIFQoFh/EuXOrz+Xw/zuPx6LpuDmVZFgSBVEy2220oFML3+/3e5/PZmnh/JUajUaVSQQgBgCzLTlIeaEwn6LperVZzuVw2mwWAyWTygAkT9C/ko++evnuVJInn+Wg0mslkAKDf7yuKoqpqp9OJRCI0TXMcN51OrSvEF52AXJaJYRgagSRJFgWTTqe73a6mabvdDk8yHo9nsxkArFarzWYDAIlEwpJ1oyewj5tNd5/1ej2fz4vFIsMwPM9jURTFVCqlaVo4HA4EAo62g/RhwXY7BoNBPp//+kur1QKAdrutKEqz2YzFYgzD3Ex86QAjlcPh0Gg06vW6KQaDQbyecrk8HA5FUfT7/fF4/MZU2MtvOwzDUAlwT5AKy7IAwLKsJElY4TgOv4KiKK/XixCq1WrJZBIAer0emXv9WNl2AEJIVVVzuFwuBUEglVf4Eafo23riFR7YjtPpRCosy1qUp7lW4ok1vcsB/JB/zKsJmqb/ownX8Xj80NTO+QPfT2qI8WvDfwAAAABJRU5ErkJggg=="

def get_anchor_paths():
    path1 = os.path.join(tempfile.gettempdir(), 'aerolumia_ancora1.png')
    path2 = os.path.join(tempfile.gettempdir(), 'aerolumia_ancora2.png')
    try:
        if not os.path.exists(path1):
            with open(path1, 'wb') as f: f.write(base64.b64decode(B64_ANCORA1))
        if not os.path.exists(path2):
            with open(path2, 'wb') as f: f.write(base64.b64decode(B64_ANCORA2))
        return path1, path2
    except: return None, None

# =========================================================================
# FUNÇÕES DE HARDWARE, TECLADO E JANELAS
# =========================================================================
ULONG_PTR = ctypes.c_uint64 if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_uint32
class KEYBDINPUT(ctypes.Structure): _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]
class MOUSEINPUT(ctypes.Structure): _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]
class HARDWAREINPUT(ctypes.Structure): _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD), ("wParamH", wintypes.WORD)]
class INPUT_UNION(ctypes.Union): _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]
class INPUT(ctypes.Structure): _fields_ = [("type", wintypes.DWORD), ("union", INPUT_UNION)]
class RECT(ctypes.Structure): _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG), ("bottom", wintypes.LONG)]
class GUITHREADINFO(ctypes.Structure): _fields_ = [("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD), ("hwndActive", wintypes.HWND), ("hwndFocus", wintypes.HWND), ("hwndCapture", wintypes.HWND), ("hwndMenuOwner", wintypes.HWND), ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND), ("rcCaret", RECT)]

def release_stuck_modifiers():
    for k in ['ctrl', 'shift', 'alt', 'win']: pyautogui.keyUp(k)
    for scan in [0x1D, 0x2A]: 
        inp = INPUT(type=1)
        inp.union.ki.wScan = scan
        inp.union.ki.dwFlags = 0x0008 | 0x0002
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

def send_fast_left_arrows(count):
    if count <= 0: return
    for chunk_start in range(0, count, 500):
        actual_count = min(500, count - chunk_start)
        inputs = (INPUT * (actual_count * 2))()
        for i in range(actual_count):
            inputs[i*2].type = 1
            inputs[i*2].union.ki.wVk = 0x25
            inputs[i*2].union.ki.dwFlags = 0
            inputs[i*2+1].type = 1
            inputs[i*2+1].union.ki.wVk = 0x25
            inputs[i*2+1].union.ki.dwFlags = 0x0002
        ctypes.windll.user32.SendInput(len(inputs), ctypes.byref(inputs), ctypes.sizeof(INPUT))
        time.sleep(0.002)

def select_with_hardware_scancodes():
    def press_scancode(scan, extend=False, up=False):
        inp = INPUT(type=1)
        inp.union.ki.wScan = scan
        inp.union.ki.dwFlags = 0x0008 | (0x0001 if extend else 0) | (0x0002 if up else 0)
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
    press_scancode(0x1D); press_scancode(0x2A); time.sleep(0.05) 
    press_scancode(0x4F, extend=True); time.sleep(0.05); press_scancode(0x4F, extend=True, up=True); time.sleep(0.05) 
    press_scancode(0x2A, up=True); press_scancode(0x1D, up=True)

def hardware_enter():
    inp = INPUT(type=1)
    inp.union.ki.wScan = 0x1C
    inp.union.ki.dwFlags = 0x0008 
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
    time.sleep(0.01) 
    inp.union.ki.dwFlags = 0x000A
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

def get_window_by_title(title_substring):
    hwnds = []
    win32gui.EnumWindows(lambda hwnd, acc: acc.append(hwnd) if win32gui.IsWindowVisible(hwnd) and title_substring in win32gui.GetWindowText(hwnd).lower() else True, hwnds)
    return hwnds[0] if hwnds else None

def reuse_or_open_edge(url):
    hwnd = get_window_by_title("edge")
    if hwnd:
        try:
            if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd); time.sleep(0.3)
            robust_clipboard_copy_text(url); pyautogui.hotkey('ctrl', 'l'); time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v'); time.sleep(0.1); pyautogui.press('enter')
            return
        except: pass
    for p in [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", r"C:\Program Files\Microsoft\Edge\Application\msedge.exe", os.path.expanduser(r"~\AppData\Local\Microsoft\Edge\Application\msedge.exe")]:
        if os.path.exists(p):
            try: subprocess.Popen([p, url]); return
            except: pass
    webbrowser.open_new_tab(url)

# =========================================================================
# ÁREA DE TRANSFERÊNCIA E HTML
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
                finally: win32clipboard.CloseClipboard()
            except: time.sleep(0.1)
        return False

def robust_clipboard_paste_text(retries=15):
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try:
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                        return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    return ""
                finally: win32clipboard.CloseClipboard()
            except: time.sleep(0.1)
        return ""

def clear_clipboard_safely(retries=15):
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try: win32clipboard.EmptyClipboard()
                finally: win32clipboard.CloseClipboard()
                return
            except: time.sleep(0.1)

def extract_html_from_clipboard(retries=15):
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try:
                    fmt = win32clipboard.RegisterClipboardFormat("HTML Format")
                    if win32clipboard.IsClipboardFormatAvailable(fmt):
                        return win32clipboard.GetClipboardData(fmt).decode('utf-8', errors='ignore')
                    return None
                finally: win32clipboard.CloseClipboard()
            except: time.sleep(0.1)
        return None

def inject_dual_format_clipboard(html_fragment, plain_text_fallback, retries=15):
    p_bytes = b"<html><body>"
    h_bytes = html_fragment.encode('utf-8')
    s_bytes = b"</body></html>"
    sh = 105; sf = 105 + len(p_bytes); ef = sf + len(h_bytes); eh = ef + len(s_bytes)
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
                finally: win32clipboard.CloseClipboard()
            except: time.sleep(0.1)

def sanitize_and_clean_html(raw_html):
    validos = [m for m in re.findall(r'\[\[INI\]\](.*?)\[\[FIM\]\]', raw_html, re.I | re.S) if len(m.strip()) > 30]
    if not validos: return None, None
    best_match = html.unescape(max(validos, key=len)).replace('\xa0', ' ').replace('\u200b', '')
    best_match = re.sub(r'(?i)LESAOMAIOR:?\s*', '', best_match).replace("****", "").replace("** **", "")
    soup = BeautifulSoup(best_match.replace('\n', ' ').replace('\r', ''), 'html.parser')
    for tag in soup.find_all(['style', 'script', 'meta', 'svg', 'button']): tag.decompose()
    clean_html_str = str(soup)
    text_soup = BeautifulSoup(clean_html_str, 'html.parser')
    for br in text_soup.find_all('br'): br.replace_with('\n')
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
