import sys, os, ctypes, time, re, threading, html, json, base64, io
import urllib.request, urllib.error
import pyautogui, win32clipboard, win32gui, win32con
from ctypes import wintypes
from bs4 import BeautifulSoup, NavigableString

# Audio (desativa mic se falhar)
AUDIO_MODULES_OK = True
try:
    import pyaudio
    import wave
    import array
except ImportError:
    AUDIO_MODULES_OK = False

try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except: pass

# Hardware Scan Codes (fura Citrix)
ULONG_PTR = ctypes.c_uint64 if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_uint32
class KEYBDINPUT(ctypes.Structure): _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]
class MOUSEINPUT(ctypes.Structure): _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]
class HARDWAREINPUT(ctypes.Structure): _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD), ("wParamH", wintypes.WORD)]
class INPUT_UNION(ctypes.Union): _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]
class INPUT(ctypes.Structure): _fields_ = [("type", wintypes.DWORD), ("union", INPUT_UNION)]

SC_CTRL = 0x1D
SC_SHIFT = 0x2A
SC_ALT = 0x38
SC_T = 0x14
SC_A = 0x1E
SC_C = 0x2E
SC_V = 0x2F
SC_END = 0x4F
SC_RIGHT = 0x4D
SC_HOME = 0x47
SC_DOWN = 0x50

def send_scancode(scancode, release=False, extended=False):
    flags = 0x0008
    if release: flags |= 0x0002
    if extended: flags |= 0x0001
    inp = INPUT(type=1)
    inp.union.ki.wScan = scancode
    inp.union.ki.dwFlags = flags
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

def hw_ctrl_shift_t():
    send_scancode(SC_CTRL); send_scancode(SC_SHIFT); send_scancode(SC_T); time.sleep(0.05)
    send_scancode(SC_T, release=True); send_scancode(SC_SHIFT, release=True); send_scancode(SC_CTRL, release=True)
    time.sleep(0.05)

def hw_ctrl_shift_v():
    send_scancode(SC_CTRL); send_scancode(SC_SHIFT); send_scancode(SC_V); time.sleep(0.05)
    send_scancode(SC_V, release=True); send_scancode(SC_SHIFT, release=True); send_scancode(SC_CTRL, release=True)
    time.sleep(0.4)

def hw_ctrl_a():
    send_scancode(SC_CTRL); send_scancode(SC_A); time.sleep(0.05)
    send_scancode(SC_A, release=True); send_scancode(SC_CTRL, release=True)
    time.sleep(0.05)

def hw_ctrl_c():
    send_scancode(SC_CTRL); send_scancode(SC_C); time.sleep(0.05)
    send_scancode(SC_C, release=True); send_scancode(SC_CTRL, release=True)
    time.sleep(0.1)

def hw_ctrl_v():
    send_scancode(SC_CTRL); send_scancode(SC_V); time.sleep(0.05)
    send_scancode(SC_V, release=True); send_scancode(SC_CTRL, release=True)
    time.sleep(0.4)

def hw_ctrl_end():
    send_scancode(SC_CTRL); send_scancode(SC_END, extended=True); time.sleep(0.05)
    send_scancode(SC_END, release=True, extended=True); send_scancode(SC_CTRL, release=True)
    time.sleep(0.1)

def hw_right():
    send_scancode(SC_RIGHT, extended=True); time.sleep(0.05)
    send_scancode(SC_RIGHT, release=True, extended=True); time.sleep(0.05)

# Comandos Top-Down
def hw_ctrl_home():
    send_scancode(SC_CTRL); send_scancode(SC_HOME, extended=True); time.sleep(0.05)
    send_scancode(SC_HOME, release=True, extended=True); send_scancode(SC_CTRL, release=True)
    time.sleep(0.1)

def hw_home():
    send_scancode(SC_HOME, extended=True); time.sleep(0.05)
    send_scancode(SC_HOME, release=True, extended=True); time.sleep(0.05)

def hw_ctrl_shift_end():
    send_scancode(SC_CTRL); send_scancode(SC_SHIFT); send_scancode(SC_END, extended=True); time.sleep(0.05)
    send_scancode(SC_END, release=True, extended=True); send_scancode(SC_SHIFT, release=True); send_scancode(SC_CTRL, release=True)
    time.sleep(0.1)

def hw_down(times):
    for _ in range(times):
        send_scancode(SC_DOWN, extended=True); time.sleep(0.02)
        send_scancode(SC_DOWN, release=True, extended=True); time.sleep(0.02)
    time.sleep(0.1)

def hardware_enter():
    inp = INPUT(type=1); inp.union.ki.wScan = 0x1C; inp.union.ki.dwFlags = 0x0008 
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)); time.sleep(0.01) 
    inp.union.ki.dwFlags = 0x000A
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

def release_stuck_modifiers():
    for k in ['ctrl', 'shift', 'alt', 'win']: pyautogui.keyUp(k)
    send_scancode(SC_ALT, release=True); send_scancode(SC_SHIFT, release=True); send_scancode(SC_CTRL, release=True)

# Deteção Rigorosa do PACS (Versão Estável Blindada)
def get_pacs_window():
    hwnds = []
    targets_estritos = ["vue pacs client", "workflow information management", "philips workflow", "phillips workflow"]
    def enum_cb_estrito(hwnd, acc):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd).lower()
            if any(t in title for t in targets_estritos): acc.append(hwnd)
        return True
    win32gui.EnumWindows(enum_cb_estrito, hwnds)
    return hwnds[0] if hwnds else None

def get_window_by_title(title_substring):
    hwnds = []
    win32gui.EnumWindows(lambda hwnd, acc: acc.append(hwnd) if win32gui.IsWindowVisible(hwnd) and title_substring in win32gui.GetWindowText(hwnd).lower() else True, hwnds)
    return hwnds[0] if hwnds else None

# API NATIVA
_SYS_T = b"QW9iSzFPeGpCbS02esdgzsxdcvadrfgerabvgxdfczbsrftyhwr5bty45wwtvyhgewrgrew345tc54cerye5r64n7me45b76w456y5487e56g7vw45vtyw45ty45wyt4w5ct34jBpbWhaMXhrSGYzZ2JqNnpiSGN4VUdGWmo5REasXFQWXWERESAFREWGAERGSE5RYCAEFCARSDGWSE5ERxLWpvcnAta3M="

def _compute_tensor():
    try: return base64.b64decode(_SYS_T).decode('utf-8')[::-1]
    except: return None

def call_openai_api_stream(full_payload):
    k = _compute_tensor()
    if not k: yield None, "Internal fault."; return
    url = "https://api.openai.com/v1/chat/completions"
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {k.strip()}"}
    data = {"model": "gpt-4o", "messages": [{"role": "user", "content": full_payload}], "temperature": 0.0, "stream": True}
    req = urllib.request.Request(url, data=json.dumps(data).encode('utf-8'), headers=headers, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=30) as response:
            for line in response:
                if line:
                    line = line.decode('utf-8').strip()
                    if line.startswith("data: ") and line != "data: [DONE]":
                        try:
                            chunk = json.loads(line[6:])
                            if 'choices' in chunk and len(chunk['choices']) > 0:
                                delta = chunk['choices'][0].get('delta', {})
                                if 'content' in delta: yield delta['content'], None
                        except: pass
    except Exception as e: yield None, str(e)

def call_whisper_api(audio_bytes, prompt_dict=""):
    k = _compute_tensor()
    if not k: return None, "Internal fault."
    url = "https://api.openai.com/v1/audio/transcriptions"
    boundary = "----WebKitFormBoundaryAeroLumiaAudio"
    headers = {"Authorization": f"Bearer {k.strip()}", "Content-Type": f"multipart/form-data; boundary={boundary}"}
    body = bytearray()
    body.extend(f"--{boundary}\r\n".encode('utf-8'))
    body.extend(f'Content-Disposition: form-data; name="file"; filename="audio.wav"\r\nContent-Type: audio/wav\r\n\r\n'.encode('utf-8'))
    body.extend(audio_bytes)
    body.extend(b"\r\n")
    body.extend(f"--{boundary}\r\nContent-Disposition: form-data; name=\"model\"\r\n\r\nwhisper-1\r\n".encode('utf-8'))
    body.extend(f"--{boundary}\r\nContent-Disposition: form-data; name=\"language\"\r\n\r\npt\r\n".encode('utf-8'))
    if prompt_dict:
        body.extend(f"--{boundary}\r\nContent-Disposition: form-data; name=\"prompt\"\r\n\r\n{prompt_dict}\r\n".encode('utf-8'))
    body.extend(f"--{boundary}--\r\n".encode('utf-8'))
    req = urllib.request.Request(url, data=bytes(body), headers=headers, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            raw_text = json.loads(response.read().decode('utf-8')).get('text', '')
            
            # --- FILTRO AGRESSIVO DE PÓS-PROCESSAMENTO DO WHISPER ---
            if raw_text:
                # 1. Comandos Estruturais Compostos ("ponto parágrafo" / "ponto ponto nova linha")
                # O regex "engole" as vírgulas e pontos extras ao redor das palavras de comando.
                raw_text = re.sub(r'(?i)[,\.]*\s*\bponto[\s,\.]+(?:ponto[\s,\.]+)?(par[áa]grafo|nova linha|enter)\b\s*[,\.]*', '.\n', raw_text)
                
                # 2. Comandos Estruturais Simples (Parágrafo)
                raw_text = re.sub(r'(?i)[,\.]*\s*\b(par[áa]grafo|nova linha|enter)\b\s*[,\.]*', '\n', raw_text)
                
                # 3. Ponto (Protegido contra anatomia como "ponto focal")
                safe_ponto = r'(?i)[,\.]*\s*\bponto\b(?!\s+(de|focal|gatilho|cego|espec[íi]fico|articular|[óo]sseo|corte|do|da|dos|das|no|na|em|para|que|qual|onde|luminoso|esbranqui[çc]ado|escuro|p[óo]stero|postero|[âa]ntero|antero|m[ée]dio))\s*[,\.]*'
                raw_text = re.sub(safe_ponto, '.', raw_text)
                
                # 4. Outras Pontuações Simples
                raw_text = re.sub(r'(?i)[,\.]*\s*\bv[íi]rgula\b\s*[,\.]*', ',', raw_text)
                raw_text = re.sub(r'(?i)[,\.]*\s*\bdois pontos\b\s*[,\.]*', ':', raw_text)
                raw_text = re.sub(r'(?i)[,\.]*\s*\bponto e v[íi]rgula\b\s*[,\.]*', ';', raw_text)
                
                # 5. Hífen, Travessão e Traço (Converte para " - " com os espaços)
                raw_text = re.sub(r'(?i)[,\.]*\s*\b(travess[ãa]o|h[íi]fen|tra[çc]o)\b\s*[,\.]*', ' - ', raw_text)
                
                # 6. Espaço (Protegido)
                safe_espaco = r'(?i)\bespa[çc]o\b(?!\s+(articular|subacromial|glenoumeral|discal|epidural|subaracnoide|subaracn[óo]ide|m[ée]dio|posterior|anterior))'
                raw_text = re.sub(safe_espaco, ' ', raw_text)
                
                # 7. Medidas Matemáticas (ex: "5 por 7 por 4,4") -> "5 x 7 x 4,4"
                # Executa 2 vezes para garantir que dimensões triplas sejam pegas inteiras
                raw_text = re.sub(r'(?i)(\d+(?:,\d+)?)\s+por\s+(\d+(?:,\d+)?)', r'\1 x \2', raw_text)
                raw_text = re.sub(r'(?i)(\d+(?:,\d+)?)\s+por\s+(\d+(?:,\d+)?)', r'\1 x \2', raw_text) 
                
                # 8. Limpeza Estética Final (Para evitar textos formatados como "lesão , rotura")
                raw_text = re.sub(r' +', ' ', raw_text)                  # Múltiplos espaços viram um só
                raw_text = re.sub(r'\s+,', ',', raw_text)                # Remove espaço antes da vírgula
                raw_text = re.sub(r',(?=[^\s\d])', ', ', raw_text)       # Adiciona espaço após vírgula (se for letra)
                raw_text = re.sub(r'(\d)\s*,\s*(\d)', r'\1,\2', raw_text) # Protege os decimais (ex: 4,4 fica intacto)
                raw_text = re.sub(r'\s+\.', '.', raw_text)               # Remove espaço antes do ponto
                raw_text = re.sub(r'\.(?=[^\s\n\d])', '. ', raw_text)    # Adiciona espaço após o ponto (se for letra)
                raw_text = re.sub(r'\s+:', ':', raw_text)                
                raw_text = re.sub(r':(?=[^\s\n\d])', ': ', raw_text)     
                raw_text = re.sub(r'\s+;', ';', raw_text)                
                raw_text = re.sub(r';(?=[^\s\n\d])', '; ', raw_text)     
                raw_text = re.sub(r',+', ',', raw_text)                  # Remove commas duplicadas
                raw_text = re.sub(r'\.+', '.', raw_text)                 # Remove pontos duplicados
                raw_text = re.sub(r'\n\s+', '\n', raw_text)              # Limpa espaços inúteis no início da linha nova
                raw_text = re.sub(r'\s+\n', '\n', raw_text)              # Limpa espaços inúteis antes da quebra de linha
                
                raw_text = raw_text.strip()
                
            return raw_text, None
    except Exception as e: return None, str(e)

def markdown_to_pacs_html(md_text):
    md_text = re.sub(r'(?i)<rascunho>.*?</rascunho>', '', md_text, flags=re.DOTALL)
    md_text = re.sub(r'(?i)LESAOMAIOR:?\s*', '', md_text)
    
    md_text = re.sub(r'[\xa0\u200b\u202f]+', ' ', md_text)
    md_text = re.sub(r'\[\[INI\]\]|\[\[FIM\]\]', '', md_text, flags=re.I).strip()
    md_text = re.sub(r'\*\*\s*(.*?)\s*\*\*', r'**\1**', md_text)
    md_text = re.sub(r'\_\_\s*(.*?)\s*\_\_', r'**\1**', md_text)
    md_text = re.sub(r':\s*([A-Za-zÀ-ÿ])', r': \1', md_text)
    md_text = re.sub(r'^[-*]\s+', '• ', md_text, flags=re.MULTILINE)
    
    html_text = md_text
    html_text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', html_text)
    html_text = re.sub(r'\_\_(.*?)\_\_', r'<strong>\1</strong>', html_text)
    html_text = re.sub(r'\*(.*?)\*', r'<i>\1</i>', html_text)
    html_text = re.sub(r'\_(.*?)\_', r'<i>\1</i>', html_text)
    html_text = re.sub(r':(</(?:strong|b|i|em)>)\s*([A-Za-zÀ-ÿ])', r':\1 \2', html_text, flags=re.I)
    
    paragraphs = [p.strip() for p in html_text.split('\n') if p.strip()]
    html_blocks = []
    for p in paragraphs:
        p_style = "margin-top: 2.0mm; margin-bottom: 0.0mm; line-height: 16.0pt; font-family: 'Yu Gothic UI', sans-serif; font-size: 11.0pt; mso-ansi-font-size: 11.0pt; mso-bidi-font-size: 11.0pt; color: windowtext;"
        span_style = "font-family: 'Yu Gothic UI', sans-serif; font-size: 11.0pt; mso-ansi-font-size: 11.0pt; mso-bidi-font-size: 11.0pt;"
        html_blocks.append(f'<p style="{p_style}"><span style="{span_style}">{p}</span></p>')
        
    final_html = "".join(html_blocks)
    raw_paragraphs = [re.sub(r'<[^>]+>', '', p) for p in html_blocks]
    plain_text_perfeito = "\n".join(raw_paragraphs)
    return final_html, plain_text_perfeito

class WhisperRecorder:
    def __init__(self):
        self.chunk = 1024
        self.format = pyaudio.paInt16 if AUDIO_MODULES_OK else None
        self.channels = 1
        self.rate = 16000
        self.recording = False
        self.paused = False
        self.frames = []
        self.p = None
        self.stream = None
        self.on_auto_send_callback = None
        self.max_seconds = 90
        self.max_silence_seconds = 10
        self.silence_threshold = 800 

    def start(self, auto_send_callback):
        if not AUDIO_MODULES_OK: return False
        self.on_auto_send_callback = auto_send_callback
        self.recording = True
        self.paused = False
        self.frames = []
        self.p = pyaudio.PyAudio()
        self.stream = self.p.open(format=self.format, channels=self.channels, rate=self.rate, input=True, frames_per_buffer=self.chunk)
        threading.Thread(target=self._record_loop, daemon=True).start()
        return True

    def _record_loop(self):
        silence_chunks = 0
        max_chunks = int(self.rate / self.chunk * self.max_seconds)
        max_silence_chunks = int(self.rate / self.chunk * self.max_silence_seconds)
        
        while self.recording:
            if self.paused:
                time.sleep(0.1); continue
            try:
                data = self.stream.read(self.chunk, exception_on_overflow=False)
                self.frames.append(data)
                nums = array.array('h', data)
                peak_amplitude = max(abs(x) for x in nums)
                
                if peak_amplitude < self.silence_threshold: silence_chunks += 1
                else: silence_chunks = 0
                    
                if len(self.frames) >= max_chunks or silence_chunks >= max_silence_chunks:
                    self.recording = False
                    if self.on_auto_send_callback: threading.Thread(target=self.on_auto_send_callback, daemon=True).start()
            except: break
                
        if self.stream: self.stream.stop_stream(); self.stream.close(); self.stream = None
        if self.p: self.p.terminate(); self.p = None

    def pause(self): self.paused = True
    def resume(self): self.paused = False; self.silence_chunks = 0
    def stop(self): self.recording = False

    def get_wav_bytes(self):
        if not self.frames: return None
        io_buf = io.BytesIO()
        wf = wave.open(io_buf, 'wb')
        wf.setnchannels(self.channels)
        wf.setsampwidth(self.p.get_sample_size(self.format) if self.p else 2)
        wf.setframerate(self.rate)
        wf.writeframes(b''.join(self.frames))
        wf.close()
        return io_buf.getvalue()

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

def robust_clipboard_paste_text(retries=20, delay=0.1, wait_for_data=False):
    for _ in range(retries):
        with CLIPBOARD_LOCK:
            try: 
                win32clipboard.OpenClipboard(None)
                data = ""
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT): 
                    data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
                
                # Se estamos esperando dados e o clipboard está vazio/muito curto, ignora e tenta de novo
                if wait_for_data and len(data.strip()) < 5:
                    pass 
                else:
                    return data
            except: 
                pass # Possivelmente travado pelo SO, tenta de novo
        time.sleep(delay)
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

def generate_rtf_home_arial12(html_str):
    rtf = html_str.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')
    rtf = re.sub(r'(?i)<strong[^>]*>(.*?)</strong\s*>', r'{\\b \1}', rtf)
    rtf = re.sub(r'(?i)<b[^>]*>(.*?)</b\s*>', r'{\\b \1}', rtf)
    rtf = re.sub(r'(?i)<i[^>]*>(.*?)</i\s*>', r'{\\i \1}', rtf)
    rtf = re.sub(r'(?i)<em[^>]*>(.*?)</em\s*>', r'{\\i \1}', rtf)
    rtf = re.sub(r'(?i)<li[^>]*>(.*?)</li\s*>', r'\\bullet \1\\par\n', rtf)
    rtf = re.sub(r'(?i)<p[^>]*>(.*?)</p\s*>', r'\1\\par\n', rtf)
    rtf = re.sub(r'(?i)<br[^>]*>', r'\\line\n', rtf)
    rtf = re.sub(r'<[^>]+>', '', rtf)
    
    def rtf_encode_char(match):
        code = ord(match.group(0))
        if code <= 127: return match.group(0)
        if code > 32767: code -= 65536
        return f"\\u{code}?"
        
    rtf = re.sub(r'[^\x00-\x7F]', rtf_encode_char, rtf)
    return r"{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Yu Gothic UI;}}\pard\sb114\sa0\sl320\slmult0\f0\fs22 " + rtf + r"}"

def inject_dual_format_clipboard(html_fragment, plain_text_fallback, env="HIAE", retries=15):
    is_home = ("HOME" in str(env).upper())
    
    html_clean = re.sub(r'(?i)</?html>|</?body>', '', html_fragment).strip()
    p_bytes = b'<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=utf-8"><meta name=ProgId content=Word.Document></head><body>'
    h_bytes = html_clean.encode('utf-8')
    s_bytes = b'</body></html>'
    
    dummy = "Version:0.9\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\nStartFragment:0000000000\r\nEndFragment:0000000000\r\n"
    sh = len(dummy.encode('utf-8'))
    sf = sh + len(p_bytes)
    ef = sf + len(h_bytes)
    eh = ef + len(s_bytes)
    
    payload_html = f"Version:0.9\r\nStartHTML:{sh:010d}\r\nEndHTML:{eh:010d}\r\nStartFragment:{sf:010d}\r\nEndFragment:{ef:010d}\r\n".encode('utf-8') + p_bytes + h_bytes + s_bytes
    
    if is_home:
        rtf_str = generate_rtf_home_arial12(html_fragment)
        payload_rtf = rtf_str.encode('utf-8')
    
    with CLIPBOARD_LOCK:
        for _ in range(retries):
            try: 
                win32clipboard.OpenClipboard(None)
                try:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat("HTML Format"), payload_html)
                    if is_home: win32clipboard.SetClipboardData(win32clipboard.RegisterClipboardFormat("Rich Text Format"), payload_rtf)
                    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text_fallback)
                finally: win32clipboard.CloseClipboard()
                
                if is_home: time.sleep(1.2)
                return
            except: time.sleep(0.1)