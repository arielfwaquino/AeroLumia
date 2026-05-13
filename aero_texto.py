import sys, os, re, json, ast, unicodedata

if getattr(sys, 'frozen', False): 
    RESOURCE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    CONFIG_DIR = os.path.dirname(sys.executable) 
else: 
    RESOURCE_DIR = CONFIG_DIR = os.path.dirname(os.path.abspath(__file__))

DICT_LOAD_ERROR = False

def super_normalize(text):
    if not text: return ""
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9\s]', ' ', text.lower())).strip()

def apply_smart_guillotine(raw_text):
    """
    Isola o Bloco de texto da Análise/Máscara usando as imagens nativas do PACS (" ") como âncora absoluta.
    """
    lines = raw_text.replace('\r', '').split('\n')
    
    # Procura todas as linhas que são EXATAMENTE um espaço (As imagens de barra azul do PACS)
    space_indices = [i for i, l in enumerate(lines) if l == " "]
    
    if not space_indices:
        return raw_text
        
    # Pega as imagens do cabeçalho (Técnica, Indicação, Análise formam no máximo 3)
    header_spaces = space_indices[:3]
    analise_image_idx = header_spaces[-1]
    
    # O texto da análise começa logo após a imagem da Análise, ignorando linhas vazias extras
    text_idx = analise_image_idx + 1
    while text_idx < len(lines) and lines[text_idx].strip() == "":
        text_idx += 1
        
    return '\n'.join(lines[text_idx:])

def calculate_header_lines(raw_text):
    """
    Motor Matemático por Âncora de Imagem.
    Informa ao RPA a linha exata da Análise, compensando imagens engolidas pelo Windows Clipboard.
    """
    lines = raw_text.replace('\r', '').split('\n')
    space_indices = [i for i, l in enumerate(lines) if l == " "]
    
    if not space_indices:
        return 0
        
    header_spaces = space_indices[:3]
    analise_image_idx = header_spaces[-1]
    
    text_idx = analise_image_idx + 1
    while text_idx < len(lines) and lines[text_idx].strip() == "":
        text_idx += 1
        
    # Se o clipboard engoliu uma imagem (ex: só achou 2 barras), sabemos que 1 linha foi perdida.
    # Adicionamos a diferença para que o hw_down bata com o número real do PACS.
    missing_images = 3 - len(header_spaces)
    
    return text_idx + missing_images

def load_external_data():
    global DICT_LOAD_ERROR
    txt_loaded, prompt, masks, frases, dicionario, whisper_dict = [], "", "", "", {}, ""
    db_path = next((p for p in [os.path.join(CONFIG_DIR, 'AeroDatabase.txt'), os.path.join(RESOURCE_DIR, 'AeroDatabase.txt')] if os.path.exists(p)), None)
            
    if db_path:
        try:
            with open(db_path, 'r', encoding='utf-8-sig') as f:
                blocks = re.split(r'\[###_(PROMPT|MASKS|FRASES|DICTIONARY|WHISPER_DICT)_###\]', f.read())
                txt_loaded.append(f"AeroDatabase.txt ({'Ext.' if db_path.startswith(CONFIG_DIR) and CONFIG_DIR != RESOURCE_DIR else 'Emb.'})")
                raw_dic = None
                for i in range(1, len(blocks), 2):
                    marker, body = blocks[i], blocks[i+1].strip()
                    if marker == "PROMPT" and body: prompt = body
                    elif marker == "MASKS" and body: masks = body
                    elif marker == "FRASES" and body: frases = body
                    elif marker == "WHISPER_DICT" and body: whisper_dict = body
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
    return prompt, masks, frases, dicionario, whisper_dict, txt_loaded

HIDDEN_PROMPT, HIDDEN_MASKS, HIDDEN_FRASES, HIDDEN_DICTIONARY, HIDDEN_WHISPER_DICT, LOADED_FILES = load_external_data()

def get_whisper_dictionary(joint_name="AUTO", modality="RM"):
    if not HIDDEN_WHISPER_DICT: return ""
    
    dict_geral = ""
    dict_especifico = ""
    
    for block in re.split(r'@@', HIDDEN_WHISPER_DICT):
        if not block.strip(): continue
        lines = block.strip().split('\n')
        header = super_normalize(lines[0])
        if "principal geral" in header or "principalgeral" in header:
            dict_geral = '\n'.join(lines[1:]).strip()
            break

    if not joint_name or joint_name == "AUTO": joint_name = "MSK"
    if joint_name == "MSK (Genérico)": joint_name = "MSK"
    
    joint_norm = super_normalize(joint_name)
    joint_keywords = ["sacro", "coccix", "sacroiliaca"] if joint_norm == "sacro coccix sacroiliacas" else [joint_norm]
    mod_keywords = ["rm", "ressonancia"] if modality == "RM" else ["tc", "tomografia"]
    old_style = super_normalize(f"{joint_name} {modality}")

    found_specific = False
    for block in re.split(r'@@', HIDDEN_WHISPER_DICT):
        if not block.strip(): continue
        lines = block.strip().split('\n')
        header = super_normalize(lines[0])
        if (any(m in header for m in mod_keywords) and any(j in header for j in joint_keywords)) or header.startswith(old_style): 
            dict_especifico = '\n'.join(lines[1:]).strip()
            found_specific = True
            break
            
    if not found_specific:
        for block in re.split(r'@@', HIDDEN_WHISPER_DICT):
            if not block.strip(): continue
            lines = block.strip().split('\n')
            if "msk" in super_normalize(lines[0]) and any(m in super_normalize(lines[0]) for m in mod_keywords):
                dict_especifico = '\n'.join(lines[1:]).strip()
                break

    final_dict = f"{dict_geral}, {dict_especifico}".strip(', ')
    return final_dict

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

def auto_translate(texto, modo):
    chave = "_TRANSLATE_RM_TC" if modo == "RM_PARA_TC" else "_TRANSLATE_TC_RM"
    for padrao, substituto in HIDDEN_DICTIONARY.get(chave, []): texto = re.sub(padrao, substituto, texto)
    return re.sub(r' {2,}', ' ', re.sub(r'\n{3,}', '\n\n', texto)).strip()

def extract_prompt_block(joint_name, modality):
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
    
def extract_frases_block(joint_name, modality):
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

def prepare_ai_payload(raw_text, ui_modality, ui_joint, use_phrases, use_translator):
    rules = HIDDEN_DICTIONARY.get("_REGEX_RULES", {})
    title_regex = rules.get("title", r'(?i)(RM |TC )')
    tc_regex = rules.get("tc", r'(?i)(tc\b)')
    rm_regex = rules.get("rm", r'(?i)(rm\b)')

    first_line = raw_text.strip().split('\n')[0] if raw_text.strip() else ""
    title_line = first_line if re.search(title_regex, first_line) else ""
    
    input_modality = "TC" if re.search(tc_regex, raw_text) else ("RM" if re.search(rm_regex, raw_text) else "RM")
    output_modality = ui_modality
    
    # 1. Poda da Análise e Filtro de Máscara DEVEM ocorrer no texto bruto original
    guillotined_text = apply_smart_guillotine(raw_text)
    clean_input = apply_mask_filter(guillotined_text)
    
    # 2. Somente após limpar o lixo nativo, aplicamos a tradução estrutural no texto residual 
    if input_modality == "RM" and output_modality == "TC": clean_input = auto_translate(clean_input, "RM_PARA_TC")
    elif input_modality == "TC" and output_modality == "RM": clean_input = auto_translate(clean_input, "TC_PARA_RM")

    final_joint = ui_joint if ui_joint != "AUTO" else find_joint(clean_input, title_line=title_line)
        
    if not final_joint: return None, clean_input, None 
        
    prompt_block = extract_prompt_block(final_joint, output_modality)
    if not prompt_block and output_modality == "TC": prompt_block = extract_prompt_block("MSK", "TC")
    
    if use_phrases:
        frases_block = extract_frases_block(final_joint, output_modality)
        if not frases_block and output_modality == "TC": frases_block = extract_frases_block("MSK", "TC")
        prompt_block = re.sub(r'<frasein[ií]cio>|<frasefim>', '', prompt_block.replace("<@insfra@>", frases_block), flags=re.IGNORECASE)
    else:
        prompt_block = re.sub(r'<frasein[ií]cio>.*?<frasefim>', '', prompt_block.replace("<@insfra@>", ""), flags=re.IGNORECASE | re.DOTALL)

    payload = f"{prompt_block}\n\n[DADOS BRUTOS]:\n{clean_input}\n\nREGRA FINAL: O laudo DEVE ser gerado entre as tags [[INI]] e [[FIM]]. Exemplo: [[INI]] texto do laudo [[FIM]]."
    
    if use_translator: 
        if trans_rule := HIDDEN_DICTIONARY.get("_PROMPT_TRANSLATION", ""): payload += f"\n\n{trans_rule}"

    return payload, clean_input, final_joint