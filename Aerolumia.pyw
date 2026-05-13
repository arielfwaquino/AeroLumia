import sys
import os

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

sys.dont_write_bytecode = True

import time, re, threading, winsound, ctypes, html, queue, json, datetime
import tkinter as tk
import customtkinter as ctk
import pyautogui, win32gui, win32con, keyboard
from PIL import Image
from bs4 import BeautifulSoup, NavigableString

from aero_core import *
from aero_texto import *

F_N = ("Arial", 10)
F_B = ("Arial", 10, "bold")
F_H = ("Arial", 13, "bold")
F_BTN = ("Arial", 14, "bold")
F_IO = ("Arial", 12)
F_CHK = ("Arial", 11.5, "bold")
F_WHISPER = ("Arial", 11.5, "bold")

PATH_ICONE = resource_path("proj_ia_icone.ico")
PATH_DATABASE = resource_path("AeroDatabase.txt")
PATH_INPUT_PNG = resource_path("proj_IA_input.png")
PATH_OUTPUT_PNG = resource_path("proj_IA_output.png")

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class WarningDialog(ctk.CTkToplevel):
    def __init__(self, master=None, title="Aviso", msg="Aviso do Sistema"):
        super().__init__(master)
        self.title(title)
        self.geometry("380x240")
        
        self.attributes('-topmost', True)
        self.transient(master)
        self.grab_set() 
        
        if master:
            self.update_idletasks()
            x = master.winfo_rootx() + (master.winfo_width() // 2) - 190
            y = master.winfo_rooty() + (master.winfo_height() // 2) - 120
            self.geometry(f"+{x}+{y}")
        
        if os.path.exists(PATH_ICONE):
            try: self.iconbitmap(PATH_ICONE)
            except: pass
                
        ctk.CTkLabel(self, text=title.upper(), font=("Arial", 20, "bold"), text_color="#d9534f").pack(pady=(25, 10))
        ctk.CTkLabel(self, text=msg, font=("Arial", 14)).pack(pady=10)
        btn = ctk.CTkButton(self, text="CIENTE", font=("Arial", 12, "bold"), width=120, command=self.on_ok)
        btn.pack(pady=(20, 10))
        
        self.bind("<Return>", lambda e: self.on_ok())
        self.bind("<Escape>", lambda e: self.on_ok())
        self.after(100, self.focus_force)

    def on_ok(self):
        self.destroy()

class LaudoApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.original_clipboard = ""
        self.title("AeroLumia Enterprise")
        self.geometry("835x525")
        self.resizable(False, False)
        
        self.abort_flag = self.is_streaming = self.is_recording = self.is_transcribing = False
        self.stream_queue = queue.Queue()
        self.stream_text_buffer = ""
        self.whisper = WhisperRecorder()
        self.current_source = 'RIS'

        if os.path.exists(PATH_ICONE):
            try: self.iconbitmap(PATH_ICONE)
            except: pass

        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=15, pady=(4, 8))
        self.main_frame.grid_columnconfigure(0, weight=53, uniform="col")
        self.main_frame.grid_columnconfigure(1, weight=47, uniform="col")
        self.main_frame.grid_rowconfigure(2, weight=1)

        self.header_left = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_left.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 5))
        self.header_right = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_right.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 5))

        if os.path.exists(PATH_INPUT_PNG):
            self.lbl_hamster = ctk.CTkLabel(self.header_left, text="", image=ctk.CTkImage(Image.open(PATH_INPUT_PNG), size=(68, 68)))
        else: 
            self.lbl_hamster = ctk.CTkLabel(self.header_left, text="[IMG]", font=F_N)
        self.lbl_hamster.pack(side="left", anchor="center", padx=(0, 15))
        ctk.CTkLabel(self.header_left, text="INPUT - Dados Brutos", font=F_H).pack(side="left", anchor="center")

        self.whisper_btn_frame = ctk.CTkFrame(self.header_left, fg_color="transparent")
        self.whisper_btn_frame.pack(side="left", padx=(10, 0), anchor="center")
        self.btn_whisper_record = ctk.CTkButton(self.whisper_btn_frame, text="GRAVA/PAUSA (Alt+Shift+W)", font=F_WHISPER, width=200, height=32, fg_color="#d9534f", text_color="white", hover_color="#c9302c", command=self.on_whisper_toggle)
        self.btn_whisper_record.pack(pady=(0, 4)) 
        self.btn_whisper_send = ctk.CTkButton(self.whisper_btn_frame, text="ENVIA WHISPER (Alt+Shift+E)", font=F_WHISPER, width=200, height=32, fg_color="#d9534f", text_color="white", hover_color="#c9302c", command=self.on_whisper_send)
        self.btn_whisper_send.pack(pady=0)

        if os.path.exists(PATH_OUTPUT_PNG):
            self.lbl_hamster_direita = ctk.CTkLabel(self.header_right, text="", image=ctk.CTkImage(Image.open(PATH_OUTPUT_PNG), size=(68, 68)))
        else: 
            self.lbl_hamster_direita = ctk.CTkLabel(self.header_right, text="[IMG]", font=F_N)
        self.lbl_hamster_direita.pack(side="left", anchor="center", padx=(0, 15))
        ctk.CTkLabel(self.header_right, text="OUTPUT - Pré-visualização", font=F_H).pack(side="left", anchor="center")

        self.control_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.control_frame.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))
        ctk.CTkLabel(self.control_frame, text="Artic", font=("Arial", 13, "bold")).pack(side="left", padx=(0, 4))
        
        self.lista_articulacoes_rm = ["AUTO", "ANTEPÉ", "BACIA", "COL CERVICAL", "COL DORSAL", "COL LOMBAR", "COTOVELO", "JOELHO", "MÃO", "MEMBROS", "OMBRO", "PUNHO", "QUADRIL", "SACRO CÓCCIX SACROILÍACAS", "TORNOZELO"]
        self.lista_articulacoes_tc = self.lista_articulacoes_rm + ["MSK (Genérico)"]
        
        self.combo_articulacao = ctk.CTkOptionMenu(self.control_frame, values=self.lista_articulacoes_rm, width=85, dynamic_resizing=False, fg_color="white", text_color="black", button_color="#e0e0e0", button_hover_color="#d0d0d0")
        self.combo_articulacao.pack(side="left", padx=(0, 8))
        self.combo_articulacao.set("AUTO")
        
        self.seg_modality = ctk.CTkSegmentedButton(self.control_frame, values=["RM", "TC"], font=("Arial", 12, "bold"), height=28, selected_color="#205c8a", selected_hover_color="#18476a", command=self.on_modality_change)
        self.seg_modality.pack(side="left", padx=(0, 8))
        self.seg_modality.set("RM")
        
        self.chk_tradutor_var = ctk.IntVar(value=0)
        self.chk_tradutor = ctk.CTkCheckBox(self.control_frame, text="TRADUTOR", variable=self.chk_tradutor_var, font=F_CHK, checkbox_width=18, checkbox_height=18, width=95)
        self.chk_tradutor.pack(side="left", padx=(0, 0))

        self.chk_frases_var = ctk.IntVar(value=0)
        self.chk_frases = ctk.CTkCheckBox(self.control_frame, text="FRASES PADRÃO", variable=self.chk_frases_var, font=F_CHK, checkbox_width=18, checkbox_height=18, command=self.on_frases_toggle)
        self.chk_frases.pack(side="left", padx=(0, 0))

        self.control_frame_right = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.control_frame_right.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(0, 10))
        self.control_frame_right.lift()

        self.paste_mode_var = ctk.StringVar(value="FULL") 
        
        self.radio_append = ctk.CTkRadioButton(self.control_frame_right, text="APPEND", variable=self.paste_mode_var, value="APPEND", font=F_CHK, text_color="black", fg_color="#205c8a", hover_color="#18476a")
        self.radio_append.pack(side="left", padx=(0, 15))
        
        self.radio_full = ctk.CTkRadioButton(self.control_frame_right, text="FULL REPORT", variable=self.paste_mode_var, value="FULL", font=F_CHK, text_color="black", fg_color="#205c8a", hover_color="#18476a")
        self.radio_full.pack(side="left", padx=(0, 15))

        self.txt_input = ctk.CTkTextbox(self.main_frame, wrap="word", width=0, font=ctk.CTkFont(*F_IO), undo=True, autoseparators=True, maxundo=-1)
        self.txt_input.grid(row=2, column=0, sticky="nsew", padx=(0, 10))
        self.txt_input.focus_set()
        self.txt_input._textbox.bind("<Button-3>", self.show_input_context_menu)
        self.txt_input._textbox.bind("<Control-y>", lambda e: self.txt_input._textbox.event_generate("<<Redo>>") or "break")
        
        self.original_input_fg = self.txt_input.cget("fg_color")
        self.original_border_width = self.txt_input.cget("border_width")
        self.original_border_color = self.txt_input.cget("border_color")
        
        self.recording_overlay = ctk.CTkFrame(self.txt_input, fg_color="#d9534f", corner_radius=8)
        self.recording_label = ctk.CTkLabel(self.recording_overlay, text="🎙️ GRAVANDO", font=("Arial", 16, "bold"), text_color="white")
        self.recording_label.pack(padx=20, pady=10)
        
        self.txt_output = ctk.CTkTextbox(self.main_frame, state="disabled", fg_color="#f5f5f5", wrap="word", cursor="arrow", width=0, font=ctk.CTkFont(*F_IO))
        self.txt_output.grid(row=2, column=1, sticky="nsew", padx=(10, 0))
        self.tk_text_core = self.txt_output._textbox
        self.tk_text_core.configure(spacing3=15)
        self.tk_text_core.tag_configure("bold", font=ctk.CTkFont(family="Arial", size=12, weight="bold"))
        self.tk_text_core.tag_configure("italic", font=ctk.CTkFont(family="Arial", size=12, slant="italic"))
        self.tk_text_core.bind("<Control-c>", self.on_custom_copy)
        self.tk_text_core.bind("<Button-3>", self.show_context_menu)

        self.working_overlay = ctk.CTkFrame(self.txt_output, fg_color="#5bc0de", corner_radius=8)
        self.working_label = ctk.CTkLabel(self.working_overlay, text="TRABALHANDO", font=("Arial", 16, "bold"), text_color="white")
        self.working_label.pack(padx=30, pady=15)

        self.bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_frame.pack(fill="x", padx=15, pady=(8, 19)) 
        self.center_wrapper = ctk.CTkFrame(self.bottom_frame, fg_color="transparent")
        self.center_wrapper.pack(expand=True)
        self.frame_botoes = ctk.CTkFrame(self.center_wrapper, fg_color="transparent")
        self.frame_botoes.pack(pady=(10, 0))
        
        self.btn_gerar_pacs = ctk.CTkButton(self.frame_botoes, text="GERAR LAUDO - RIS (Alt+Shift+Z)", width=290, height=42, font=F_BTN, command=self.btn_gerar_pacs_click)
        self.btn_gerar_pacs.pack(side="left", padx=8) 
        self.btn_gerar_input = ctk.CTkButton(self.frame_botoes, text="GERAR LAUDO - CAIXA DE INPUT", width=270, height=42, font=F_BTN, command=self.btn_gerar_input_click)
        self.btn_gerar_input.pack(side="left", padx=8)
        self.btn_novo = ctk.CTkButton(self.frame_botoes, text="RESET (Ctrl+Esc)", width=160, height=42, font=F_BTN, command=self.trigger_global_abort, fg_color="#d9534f", hover_color="#c9302c")
        self.btn_novo.pack(side="left", padx=8)
        
        self.lbl_status = ctk.CTkLabel(self.center_wrapper, text="", font=("Arial", 14, "bold"))
        self.lbl_status.pack(pady=(15, 5))
        self.selected_joint = None
        
        hotkeys = {
            'alt+shift+z': self.on_global_hotkey_append, 
            'alt+shift+w': self.on_whisper_toggle, 
            'alt+shift+e': self.on_whisper_send, 
            'ctrl+esc': self.trigger_global_abort
        }
        
        try:
            for k, func in hotkeys.items(): keyboard.add_hotkey(k, func)
        except: pass
        
        if DICT_LOAD_ERROR: self.update_status("ERRO NO JSON DO AeroDatabase.txt.", "red")
        elif LOADED_FILES: self.update_status("Sistema Inicializado.", "green")
        else: self.update_status("Sistema Inicializado.", "green")

    def detect_environment(self):
        windows = []
        try:
            def enum_cb(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if title: results.append(title)
                return True
            
            win32gui.EnumWindows(enum_cb, windows)
            
            targets = ["vue pacs client", "workflow information management", "philips workflow", "phillips workflow"]
            for title in windows:
                t_lower = title.lower()
                if any(t in t_lower for t in targets):
                    if "\\remote" in t_lower: return "HOME - EXPERIMENTAL"
                    else: return "HIAE"
        except: pass
        return "AGUARDANDO PACS"

    def format_joint_name(self, joint_name):
        if not joint_name or joint_name == "AUTO": return ""
        name = joint_name.upper()
        
        if "COLUNA CERVICAL" in name or "COL CERVICAL" in name: return "CERVICAL"
        if "SACRO CÓCCIX SACROILÍACAS" in name: return "SAC"
        return name.replace("COLUNA ", "").replace("COL ", "")

    def update_status(self, text, color="black", size=14, bold=True):
        if color == "orange": color = "#c05600"
        ambiente = self.detect_environment()
        
        art_info = ""
        if self.selected_joint and self.selected_joint != "AUTO": 
            art_info = f" - {self.format_joint_name(self.selected_joint)}"
        elif self.combo_articulacao.get() != "AUTO": 
            art_info = f" - {self.format_joint_name(self.combo_articulacao.get())}"
            
        texto_final = f"[{ambiente}{art_info}] {text}"
        self.lbl_status.configure(text=texto_final, text_color=color, font=("Arial", size, "bold") if bold else ("Arial", size))
        self.update_idletasks()

    def on_modality_change(self, value):
        current_val = self.combo_articulacao.get()
        if value == "TC": 
            self.combo_articulacao.configure(values=self.lista_articulacoes_tc)
        else:
            self.combo_articulacao.configure(values=self.lista_articulacoes_rm)
            if current_val == "MSK (Genérico)": self.combo_articulacao.set("AUTO")

    def on_frases_toggle(self):
        if self.chk_frases_var.get() == 1: 
            WarningDialog(self, title="Aviso - Frases Padrão", msg="Função experimental.\nPode apresentar instabilidades.")

    def _has_sel(self, widget):
        try: return bool(widget.tag_ranges("sel"))
        except: return False

    def safe_undo(self):
        try: self.txt_input._textbox.event_generate("<<Undo>>")
        except: pass

    def safe_redo(self):
        try: self.txt_input._textbox.event_generate("<<Redo>>")
        except: pass

    def show_input_context_menu(self, event):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Desfazer (Ctrl+Z)", command=self.safe_undo)
        menu.add_command(label="Refazer (Ctrl+Y)", command=self.safe_redo)
        menu.add_separator()
        if self._has_sel(self.txt_input._textbox):
            menu.add_command(label="Recortar (Ctrl+X)", command=lambda: self.txt_input._textbox.event_generate("<<Cut>>"))
            menu.add_command(label="Copiar (Ctrl+C)", command=lambda: self.txt_input._textbox.event_generate("<<Copy>>"))
        menu.add_command(label="Colar (Ctrl+V)", command=lambda: self.txt_input._textbox.event_generate("<<Paste>>"))
        menu.tk_popup(event.x_root, event.y_root)

    def show_context_menu(self, event):
        menu = tk.Menu(self, tearoff=0)
        if self._has_sel(self.tk_text_core):
            menu.add_command(label="Copiar texto selecionado", command=self.copy_partial_selection)
            menu.add_separator()
        menu.add_command(label="Copiar laudo inteiro", command=self.copy_whole_laudo)
        menu.add_command(label="Copiar laudo inteiro e enviar ao RIS", command=lambda: self.copy_and_send_to_ris(force_append=True))
        menu.tk_popup(event.x_root, event.y_root)

    def on_custom_copy(self, event=None):
        if self._has_sel(self.tk_text_core): self.copy_partial_selection()
        else: self.copy_whole_laudo()
        return "break"

    def copy_partial_selection(self):
        try: start, end = self.tk_text_core.index("sel.first"), self.tk_text_core.index("sel.last")
        except: return
        data = self.tk_text_core.dump(start, end, text=True, tag=True)
        active_tags, html_lines, plain_lines, c_html, c_plain = [], [], [], "", ""
        for key, val, _ in data:
            if key == "tagon":
                tag = val.replace('bold', 'strong')
                if tag not in active_tags: active_tags.append(tag)
                c_html += f"<{tag}>"
            elif key == "tagoff":
                tag = val.replace('bold', 'strong')
                if tag in active_tags: active_tags.remove(tag)
                c_html += f"</{tag}>"
            elif key == "text":
                parts = val.split('\n')
                for i, part in enumerate(parts):
                    if i > 0:
                        closing = "".join([f"</{t}>" for t in reversed(active_tags)])
                        html_lines.append(c_html + closing)
                        plain_lines.append(c_plain)
                        c_html = "".join([f"<{t}>" for t in active_tags])
                        c_plain = ""
                    if part: 
                        c_html += html.escape(part)
                        c_plain += part
        if c_plain or c_html:
            html_lines.append(c_html + "".join([f"</{t}>" for t in reversed(active_tags)]))
            plain_lines.append(c_plain)
        final_html = "".join([f'<p style="margin: 8px 0px 8px 0px;">{l}</p>' for l in html_lines if l.strip() and not re.fullmatch(r'^(<[^>]+>|\s)*$', l)])
        inject_dual_format_clipboard(final_html, "\n".join([p for p in plain_lines if p.strip()]), env=self.detect_environment())
        self.update_status("Texto selecionado copiado!", "green")

    def copy_whole_laudo(self):
        if hasattr(self, 'last_html') and self.last_html:
            inject_dual_format_clipboard(self.last_html, self.last_plain, env=self.detect_environment())
            self.update_status("Laudo copiado para a área de transferência!", "green")
        else: 
            self.update_status("Nenhum laudo disponível.", "red")

    def copy_and_send_to_ris(self, force_append=False):
        if hasattr(self, 'last_html') and self.last_html:
            self.update_status("Enviando ao RIS...", "orange")
            threading.Thread(target=self.paste_and_select_in_ris, args=(self.last_plain, self.last_html, force_append), daemon=True).start()
        else: 
            self.update_status("Nenhum laudo disponível.", "red")

    def on_whisper_toggle(self):
        if self.abort_flag or self.is_transcribing: return
        
        if not self.is_recording:
            if self.whisper.start(self.on_whisper_auto_send):
                self.is_recording = True
                self.after(0, self.ui_start_recording)
            else: 
                self.after(0, lambda: self.update_status("Erro: Módulos de áudio ausentes.", "red"))
        else:
            if self.whisper.paused:
                self.whisper.resume()
                self.after(0, lambda: self.recording_label.configure(text="🎙️ GRAVANDO"))
            else:
                self.whisper.pause()
                self.after(0, lambda: self.recording_label.configure(text="⏸️ PAUSADO"))

    def on_whisper_send(self):
        if not self.is_recording or self.is_transcribing: return
        self.on_whisper_auto_send()

    def on_whisper_auto_send(self):
        self.is_recording = False
        self.whisper.stop()
        self.after(0, self.ui_stop_recording_and_transcribe)

    def ui_start_recording(self):
        self.recording_label.configure(text="🎙️ GRAVANDO")
        self.recording_overlay.place(relx=0.5, rely=0.15, anchor="center")
        self.pulse_state = False
        self.pulse_recording_ui()

    def pulse_recording_ui(self):
        if not self.is_recording:
            self.txt_input.configure(fg_color=self.original_input_fg, border_width=self.original_border_width, border_color=self.original_border_color)
            return
        if self.whisper.paused: 
            self.txt_input.configure(fg_color="#e6e6e6", border_width=2, border_color="gray")
        else:
            self.pulse_state = not self.pulse_state
            bg_color = "#ffe6e6" if self.pulse_state else "#ffcccc"
            self.txt_input.configure(fg_color=bg_color, border_width=2, border_color="red")
        self.after(600, self.pulse_recording_ui)

    def ui_stop_recording_and_transcribe(self):
        self.txt_input.configure(fg_color=self.original_input_fg, border_width=self.original_border_width, border_color=self.original_border_color)
        wav_bytes = self.whisper.get_wav_bytes()
        if not wav_bytes:
            self.recording_overlay.place_forget()
            return self.update_status("Gravação vazia.", "orange")
            
        self.recording_label.configure(text="⏳ TRANSCREVENDO...")
        self.is_transcribing = True
        self.update_status("[ WHISPER... ]", "blue")
        threading.Thread(target=self.process_transcription, args=(wav_bytes,), daemon=True).start()

    def process_transcription(self, wav_bytes):
        text, err = call_whisper_api(wav_bytes, get_whisper_dictionary(self.combo_articulacao.get(), self.seg_modality.get()))
        self.after(0, lambda: self.ui_finish_transcription(text, err))

    def ui_finish_transcription(self, text, err):
        self.recording_overlay.place_forget()
        self.is_transcribing = False
        if err: return self.update_status(f"Erro Whisper: {err}", "red")
        if text:
            try:
                if self.txt_input._textbox.tag_ranges("sel"): 
                    self.txt_input._textbox.delete("sel.first", "sel.last")
            except: pass
            self.txt_input._textbox.edit_separator()
            self.txt_input.insert("insert", text.strip() + " ")
            self.txt_input._textbox.edit_separator()
            self.update_status("Transcrição concluída.", "green")

    def silent_cleanup(self):
        self.abort_flag = False
        self.selected_joint = None
        self.txt_output.configure(state="normal")
        self.txt_output.delete("1.0", "end")
        self.txt_output.configure(state="disabled")

    def sleep_safe(self, seconds):
        end_time = time.time() + seconds
        while time.time() < end_time:
            if self.abort_flag: return False
            time.sleep(0.05)
        return True

    def trigger_global_abort(self):
        self.abort_flag = True
        release_stuck_modifiers()
        self.after(0, self.reset_ui_completely)

    def reset_ui_completely(self):
        self.is_streaming = self.is_transcribing = self.is_recording = False
        if hasattr(self, 'whisper'): self.whisper.stop()
        for overlay in ['working_overlay', 'recording_overlay']:
            if hasattr(self, overlay): getattr(self, overlay).place_forget()
            
        self.txt_input.configure(fg_color=self.original_input_fg, border_width=self.original_border_width, border_color=self.original_border_color)
        self.txt_output.configure(state="normal")
        self.txt_output.delete("1.0", "end")
        self.txt_output.configure(state="disabled")
        self.combo_articulacao.set("AUTO")
        self.change_ui_state("normal")
        self.selected_joint = None
        
        if hasattr(self, 'original_clipboard') and self.original_clipboard: 
            robust_clipboard_copy_text(self.original_clipboard)
            
        self.update_status("SISTEMA ABORTADO.", "red")
        self.after(800, lambda: setattr(self, 'abort_flag', False))

    def show_error_window(self, title, msg, btn_text, geo="400x200"):
        self.abort_flag = True
        release_stuck_modifiers()
        self.reset_ui_completely()
        self.after(800, lambda: setattr(self, 'abort_flag', False))
        
        try: winsound.MessageBeep(winsound.MB_ICONHAND) 
        except: pass
        
        err_win = ctk.CTkToplevel(self)
        err_win.title(title)
        err_win.geometry(geo)
        err_win.attributes('-topmost', True)
        err_win.grab_set() 
        
        ctk.CTkLabel(err_win, text=msg, font=("Arial", 16, "bold"), text_color="#B22222").pack(expand=True, pady=(20, 10))
        
        def on_ok(event=None):
            err_win.destroy()
                
        ctk.CTkButton(err_win, text=btn_text, font=("Arial", 14, "bold"), command=on_ok, width=150, height=40).pack(pady=(0, 20))
        err_win.bind("<Return>", on_ok)
        err_win.bind("<Escape>", on_ok)
        err_win.after(100, err_win.focus_force)

    def bring_to_front(self):
        self.lift()
        self.attributes('-topmost', True)
        self.after(200, lambda: self.attributes('-topmost', False))
        self.focus_force()

    def on_global_hotkey_append(self):
        if self.btn_gerar_pacs.cget("state") == "normal": 
            self.after(0, self.btn_gerar_pacs_click)

    def change_ui_state(self, state):
        for w in [self.btn_gerar_pacs, self.btn_gerar_input, self.combo_articulacao, self.seg_modality, self.chk_tradutor, self.chk_frases, self.radio_append, self.radio_full]: 
            w.configure(state=state)

    def finaliza_rotina(self, msg, color, success=True):
        try: release_stuck_modifiers()
        except: pass
        if self.abort_flag: return 
        if success:
            self.seg_modality.set("RM")
            self.on_modality_change("RM")
            self.chk_tradutor_var.set(0)
            self.chk_frases_var.set(0)
            
        self.change_ui_state("normal")
        self.combo_articulacao.set("AUTO")
        self.update_status(msg, color)

    def render_html_to_output(self, html_string):
        self.txt_output.configure(state="normal")
        self.txt_output.delete("1.0", "end")
        
        def nl():
            txt = self.tk_text_core.get("1.0", "end-1c")
            if txt and not txt.endswith("\n"): 
                self.tk_text_core.insert("end", "\n")

        def walk_tree(node, active_tags):
            if isinstance(node, NavigableString):
                text = re.sub(r' {2,}', ' ', str(node).replace('\n', ' ').replace('\r', '').replace('\t', ' '))
                if text.startswith(' ') and (not (c := self.tk_text_core.get("1.0", "end-1c")) or c.endswith('\n')): 
                    text = text[1:]
                if text: 
                    self.tk_text_core.insert("end", text, tuple(active_tags))
            else:
                current_tags = active_tags.copy()
                is_bold = False
                
                if node.name in ['b', 'strong']: 
                    is_bold = True
                elif node.get('style'):
                    if match := re.search(r'font-weight\s*:\s*([a-z0-9]+)', node.get('style').lower()):
                        val = match.group(1)
                        if val in ['bold', 'bolder'] or (val.isdigit() and int(val) >= 600): 
                            is_bold = True
                            
                if is_bold: 
                    current_tags.append("bold")
                if node.name in ['i', 'em']: 
                    current_tags.append("italic")
                    
                if node.name in ['p', 'div']: 
                    nl()
                elif node.name == 'li': 
                    nl()
                    self.tk_text_core.insert("end", "• ", tuple(current_tags))
                elif node.name == 'br': 
                    nl()
                    
                for child in node: 
                    walk_tree(child, current_tags)
                    
                if node.name in ['p', 'div']: 
                    nl()
                elif node.name == 'li': 
                    nl()

        walk_tree(BeautifulSoup(html_string, 'html.parser'), [])
        if (lw := len(c := self.tk_text_core.get("1.0", "end-1c")) - len(c.lstrip())) > 0: 
            self.tk_text_core.delete("1.0", f"1.0+{lw}c")
        self.txt_output.configure(state="disabled")

    def request_manual_joint(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Definição Manual")
        dialog.geometry("300x200")
        dialog.attributes("-topmost", True)
        dialog.grab_set() 
        
        ctk.CTkLabel(dialog, text="Articulação não identificada.\nSelecione:").pack(pady=15)
        combo = ctk.CTkComboBox(dialog, values=["SELECIONE..."] + sorted([k for k in HIDDEN_DICTIONARY.keys() if not k.startswith("_")]))
        combo.set("SELECIONE...")
        combo.pack(pady=10)
        
        def confirm(event=None): 
            if combo.get() != "SELECIONE...": 
                self.selected_joint = combo.get()
                dialog.destroy()
                
        ctk.CTkButton(dialog, text="Confirmar", command=confirm).pack(pady=15)
        dialog.bind("<Return>", confirm)
        dialog.bind("<Escape>", lambda e: dialog.destroy())
        dialog.after(100, dialog.focus_force)
        self.wait_window(dialog) 

    def btn_gerar_input_click(self):
        self.current_source = 'INPUT_BOX'
        raw_text = self.txt_input.get("1.0", "end-1c").strip()
        if not raw_text: return self.update_status("ERRO: Caixa vazia.", "red")
        
        paste_mode = self.paste_mode_var.get()
        if paste_mode == "FULL":
            self.update_status("[ LENDO PACS JIT... ]", "blue")
            threading.Thread(target=self._read_pacs_silently_and_run, args=(raw_text,), daemon=True).start()
        else:
            self.last_raw_pacs_text = ""
            self.original_clipboard = robust_clipboard_paste_text()
            self.silent_cleanup()
            self.change_ui_state("disabled")
            self.process_raw_text_and_run(raw_text)

    def _extract_text_with_retries(self, max_attempts=3):
        for attempt in range(max_attempts):
            if self.abort_flag: return ""
            release_stuck_modifiers()
            clear_clipboard_safely()
            try:
                hw_ctrl_shift_t(); time.sleep(0.05)
                hw_ctrl_shift_t(); time.sleep(0.05)
                hw_ctrl_a(); time.sleep(0.1)
                hw_ctrl_c(); time.sleep(0.3)
                hw_right(); time.sleep(0.05)
            except: release_stuck_modifiers()
            
            raw = robust_clipboard_paste_text(retries=5, delay=0.1, wait_for_data=True)
            if raw and len(raw.strip()) > 0:
                return raw
                
            if attempt < max_attempts - 1:
                time.sleep(0.5)
        return ""

    def _read_pacs_silently_and_run(self, raw_text):
        hwnd = get_pacs_window()
        if hwnd:
            try:
                pyautogui.press('alt'); time.sleep(0.05)
                if win32gui.GetForegroundWindow() != hwnd:
                    if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                    win32gui.SetForegroundWindow(hwnd); time.sleep(0.4)
                else: time.sleep(0.05)
                
                pacs_text = self._extract_text_with_retries(max_attempts=3)
                self.last_raw_pacs_text = pacs_text 
                
                self.after(0, self.bring_to_front)
            except: pass
        else: self.last_raw_pacs_text = ""
            
        self.original_clipboard = robust_clipboard_paste_text()
        self.after(0, self.silent_cleanup)
        self.after(0, lambda: self.change_ui_state("disabled"))
        self.after(0, lambda: self.process_raw_text_and_run(raw_text))

    def btn_gerar_pacs_click(self):
        self.current_source = 'RIS'
        self._start_pacs_extraction()

    def _start_pacs_extraction(self):
        self.original_clipboard = robust_clipboard_paste_text()
        clear_clipboard_safely() 
        self.silent_cleanup()
        self.change_ui_state("disabled")
        self.update_status("[ EXTRAINDO PACS... ]", "blue")
        threading.Thread(target=self.fetch_from_pacs_and_run, daemon=True).start()

    def fetch_from_pacs_and_run(self):
        if not self.sleep_safe(0.2): return
        hwnd = get_pacs_window()
        if not hwnd: return self.after(0, lambda: self.finaliza_rotina("ERRO: PACS não encontrado.", "red", False))
        
        try:
            pyautogui.press('alt'); time.sleep(0.05)
            if win32gui.GetForegroundWindow() != hwnd:
                if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                win32gui.SetForegroundWindow(hwnd); time.sleep(0.4)
            else: time.sleep(0.05)
        except: pass 
        
        raw_text = self._extract_text_with_retries(max_attempts=3)
        self.after(0, self.bring_to_front)
        if self.abort_flag: return
        
        self.last_raw_pacs_text = raw_text 
        self.after(0, lambda: self.process_raw_text_and_run(raw_text, update_ui=True))

    def process_raw_text_and_run(self, raw_text, update_ui=False):
        raw_stripped = raw_text.strip()
        
        if len(raw_stripped) == 0:
            msg = "Falha ao capturar os dados do PACS.\n\nO sistema tentou copiar o texto 3 vezes sem sucesso.\nVerifique se a janela do laudo está aberta e visível."
            self.after(0, lambda: self.show_error_window("Falha de Extração", msg, "CIENTE", geo="600x210"))
            return
            
        elif len(raw_stripped) < 15:
            msg = f"Número mínimo de caracteres (15) não atingido.\nForam detectados apenas {len(raw_stripped)} caracteres.\n\nVerifique se o texto copiado está correto."
            self.after(0, lambda: self.show_error_window("Validação", msg, "CIENTE", geo="600x210"))
            return

        if self.abort_flag: return
        self.update_status("[ PREPARAÇÃO... ]", "blue")
        
        joint_selecionada = self.combo_articulacao.get()
        mapa_db = {
            "COL CERVICAL": "COLUNA CERVICAL",
            "COL DORSAL": "COLUNA DORSAL",
            "COL LOMBAR": "COLUNA LOMBAR",
            "SACRO CÓCCIX SACROILÍACAS": "SACROCÓCCIX" 
        }
        joint_db = mapa_db.get(joint_selecionada, joint_selecionada)
        
        payload, clean_input, final_joint = prepare_ai_payload(
            raw_text, 
            self.seg_modality.get(), 
            joint_db, 
            self.chk_frases_var.get() == 1, 
            self.chk_tradutor_var.get() == 1
        )
        
        self.saved_clean_input = clean_input.strip() if clean_input else ""
        
        if update_ui: 
            self.after(0, lambda: self.txt_input.delete("1.0", "end") or self.txt_input.insert("1.0", clean_input))
            
        if not final_joint:
            dialog = ctk.CTkToplevel(self)
            dialog.title("Manual")
            dialog.geometry("300x200")
            dialog.attributes("-topmost", True)
            dialog.grab_set() 
            
            ctk.CTkLabel(dialog, text="Articulação não identificada.\nSelecione:").pack(pady=15)
            combo = ctk.CTkComboBox(dialog, values=["SELECIONE..."] + sorted([k for k in HIDDEN_DICTIONARY.keys() if not k.startswith("_")]))
            combo.set("SELECIONE...")
            combo.pack(pady=10)
            
            def confirm(event=None): 
                if combo.get() != "SELECIONE...": 
                    self.selected_joint = combo.get()
                    dialog.destroy()
                    
            ctk.CTkButton(dialog, text="Confirmar", command=confirm).pack(pady=15)
            dialog.bind("<Return>", confirm)
            dialog.bind("<Escape>", lambda e: dialog.destroy())
            dialog.after(100, dialog.focus_force)
            self.wait_window(dialog)
            
        if not self.selected_joint and not final_joint or self.abort_flag: 
            return self.after(0, lambda: self.finaliza_rotina("Cancelado.", "red", False))

        self.selected_joint = final_joint

        self.txt_output.configure(state="normal")
        self.txt_output.delete("1.0", "end")
        self.txt_output.configure(state="disabled")
        self.working_overlay.place(relx=0.5, rely=0.38, anchor="center")
        self.stream_text_buffer = ""
        self.is_streaming = True
        self.after(50, self.process_stream_queue)
        threading.Thread(target=self.run_direct_api_stream_automation, args=(payload,), daemon=True).start()

    def run_direct_api_stream_automation(self, full_payload):
        if self.abort_flag: return
        self.after(0, lambda: self.update_status(f"[ GERANDO LAUDO... ]", "orange"))
        has_started = False
        error_msg = None
        
        for chunk, err in call_openai_api_stream(full_payload):
            if self.abort_flag: break
            if err: 
                error_msg = err
                break
            if chunk:
                if not has_started: 
                    self.stream_queue.put({"action": "clear"})
                    has_started = True
                self.stream_queue.put({"action": "insert", "text": chunk})
                self.stream_text_buffer += chunk

        if self.abort_flag: return
        if error_msg: self.stream_queue.put({"action": "error", "msg": error_msg})
        else: self.stream_queue.put({"action": "finish", "full_text": self.stream_text_buffer})

    def process_stream_queue(self):
        if not self.is_streaming: return
        try:
            while True:
                msg = self.stream_queue.get_nowait()
                action = msg.get("action")
                
                if action == "clear":
                    self.txt_output.configure(state="normal", text_color="black")
                    self.txt_output.delete("1.0", "end")
                    self.txt_output.configure(state="disabled")
                elif action == "insert":
                    self.txt_output.configure(state="normal")
                    self.txt_output.insert("end", msg.get("text"))
                    self.txt_output.see("end")
                    self.txt_output.configure(state="disabled")
                elif action == "error":
                    self.is_streaming = False
                    self.working_overlay.place_forget()
                    self.finaliza_rotina(f"ERRO: {msg.get('msg')}", "red", False)
                    return
                elif action == "finish":
                    self.is_streaming = False
                    self.working_overlay.place_forget() 
                    full_md = msg.get("full_text")
                    if not full_md: return self.finaliza_rotina("ERRO: Vazio", "red", False)
                    
                    self.update_status("[ FORMATANDO... ]", "blue")
                    clean_html_visual, _ = markdown_to_pacs_html(full_md)
                    self.render_html_to_output(clean_html_visual)
                    
                    paste_mode = self.paste_mode_var.get()
                    clipboard_md = full_md
                    
                    if paste_mode == "IN_OUT":
                        input_text = getattr(self, 'saved_clean_input', "")
                        if input_text:
                            clipboard_md = f"{input_text}\n\n{full_md}"
                            
                    clean_html_clip, plain_text_clip = markdown_to_pacs_html(clipboard_md)
                    
                    if clean_html_clip and plain_text_clip: 
                        self.last_html = clean_html_clip
                        self.last_plain = plain_text_clip
                        
                        if self.current_source == 'INPUT_BOX' and paste_mode in ["FULL", "IN_OUT"]:
                            threading.Thread(target=self._jit_read_and_paste, args=(plain_text_clip, clean_html_clip), daemon=True).start()
                        else:
                            threading.Thread(target=self.paste_and_select_in_ris, args=(plain_text_clip, clean_html_clip, False), daemon=True).start()
                    else: 
                        self.finaliza_rotina("ERRO Formatação.", "red", False)
                    return
        except queue.Empty: pass
        if self.is_streaming: self.after(50, self.process_stream_queue)

    def _jit_read_and_paste(self, plain_text, clean_html):
        self.update_status("[ LENDO PACS JIT... ]", "blue")
        hwnd = get_pacs_window()
        if not hwnd: 
            self.after(0, self.finalize_ui_success_manual)
            return
            
        try:
            pyautogui.press('alt'); time.sleep(0.05)
            if win32gui.GetForegroundWindow() != hwnd:
                if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                win32gui.SetForegroundWindow(hwnd); time.sleep(0.4)
            else: time.sleep(0.05)
            
            pacs_text = self._extract_text_with_retries(max_attempts=3)
            self.last_raw_pacs_text = pacs_text 
        except: pass
        
        self.paste_and_select_in_ris(plain_text, clean_html, force_append=False)

    def paste_and_select_in_ris(self, plain_text, clean_html=None, force_append=False):
        try:
            if self.abort_flag: return
            self.after(0, lambda: self.update_status("[ COLANDO PACS... ]", "orange"))
            hwnd = get_pacs_window()
            if not hwnd: 
                self.after(0, self.finalize_ui_success_manual)
                return
                
            try:
                pyautogui.press('alt'); time.sleep(0.05)
                if win32gui.GetForegroundWindow() != hwnd:
                    if win32gui.IsIconic(hwnd): win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                    win32gui.SetForegroundWindow(hwnd); time.sleep(0.4)
                else: time.sleep(0.05)
            except: pass 
                
            release_stuck_modifiers()
            try:
                if self.abort_flag: return
                paste_mode = "APPEND" if force_append else self.paste_mode_var.get()
                
                if paste_mode == "FULL" and getattr(self, 'last_raw_pacs_text', ""):
                    lines_to_drop = calculate_header_lines(self.last_raw_pacs_text)
                    inject_dual_format_clipboard(clean_html, plain_text, env=self.detect_environment())
                    time.sleep(0.2)
                    
                    hw_ctrl_shift_t(); time.sleep(0.05)
                    hw_ctrl_shift_t(); time.sleep(0.05)
                    hw_ctrl_home(); time.sleep(0.1)
                    if lines_to_drop > 0: hw_down(lines_to_drop); time.sleep(0.1)
                    hw_home(); time.sleep(0.1) 
                    hw_ctrl_shift_end(); time.sleep(0.15)
                    hw_ctrl_shift_v(); time.sleep(0.4)
                else:
                    inject_dual_format_clipboard(clean_html, plain_text, env=self.detect_environment())
                    time.sleep(0.2)
                    
                    hw_ctrl_shift_t(); time.sleep(0.05)
                    hw_ctrl_shift_t(); time.sleep(0.05)
                    hw_ctrl_end(); time.sleep(0.1)
                    hardware_enter(); time.sleep(0.05)
                    hardware_enter(); time.sleep(0.05)
                    hw_ctrl_shift_v(); time.sleep(0.4)
                    
                # Limpeza cirúrgica garantida após injeção via hardware scan codes.
                clear_clipboard_safely()
                
            except: release_stuck_modifiers()
            
            if self.abort_flag: return
            time.sleep(0.15) 
            self.after(0, lambda: self.finaliza_rotina("LAUDO PRONTO!", "green", True))
        except Exception as e:
            self.after(0, lambda: self.finaliza_rotina(f"ERRO PACS: {str(e)}", "red", False))

    def finalize_ui_success_manual(self):
        if self.abort_flag: return
        self.lift()
        self.attributes('-topmost', True)
        self.after(500, lambda: self.attributes('-topmost', False))
        self.finaliza_rotina("COLE MANUALMENTE", "orange", True)

if __name__ == "__main__":
    app = LaudoApp()
    app.mainloop()
