from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from PIL import Image
import io
import re
import os
import fitz  # PyMuPDF

# ==========================================
# 🔲 フォント設定
# ==========================================
CUSTOM_FONT_FILE = "brush.ttf"
FONT_NAME = "KakizomeFont"

# ==========================================
# 📐 印刷用の位置設定
# ==========================================
OFFSET_X = 0.7 * mm 
OFFSET_Y = 1.3 * mm 

ZIP_Y = 148.0 * mm - 15.8 * mm 
ZIP_STEP = 7.3 * mm 
ZIP_X_LEFT_START = 46.0 * mm 
ZIP_X_RIGHT_START = ZIP_X_LEFT_START + (3 * ZIP_STEP) + (0.6 * mm)

HAGAKI_WIDTH = 100 * mm
HAGAKI_HEIGHT = 148 * mm

# ==========================================
# 📺 プレビュー画面専用の調整
# ==========================================
PREVIEW_ADJUST_X_MM = 0.0
PREVIEW_ADJUST_Y_MM = -4.0

# ==========================================

# フォント登録
try:
    if os.path.exists(CUSTOM_FONT_FILE):
        pdfmetrics.registerFont(TTFont(FONT_NAME, CUSTOM_FONT_FILE))
    else:
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))
        FONT_NAME = "HeiseiMin-W3"
except:
    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))
    FONT_NAME = "HeiseiMin-W3"


# --- 共通関数 ---
def get_zipcode_digits(address):
    zipcode = ""
    zip_match = re.search(r'\d{3}-?\d{4}', address)
    if zip_match:
        zipcode = zip_match.group()
        address = address.replace(zipcode, "").strip()
    digits = re.sub(r'[^0-9]', '', str(zipcode))
    return digits, address

def smart_split_address(address):
    lines = []
    LIMIT_1 = 17
    LIMIT_2 = 18
    
    blocks = re.split(r'[ 　]+', address.strip())
    current_line = ""
    
    for block in blocks:
        if not block: continue
        if current_line and (len(current_line) + len(block) + 1 <= LIMIT_1):
            current_line += " " + block
        else:
            if current_line:
                lines.append(current_line)
            if len(block) > LIMIT_1:
                while len(block) > LIMIT_2:
                    lines.append(block[:LIMIT_2])
                    block = block[LIMIT_2:]
                current_line = block
            else:
                current_line = block
    
    if current_line:
        lines.append(current_line)
    
    return lines[:3]

# --- PDF描画クラス ---
class VerticalTextRendererPDF:
    def __init__(self, canvas_obj, font_name):
        self.c = canvas_obj
        self.font_name = font_name
        self.trans_map = str.maketrans({
            '0': '〇', '1': '一', '2': '二', '3': '三', '4': '四',
            '5': '五', '6': '六', '7': '七', '8': '八', '9': '九',
            '-': '丨', 'ー': '丨', '－': '丨', '(': '︵', ')': '︶'
        })

    def draw_text(self, text, x, y_start, max_height, max_font_size, line_spacing=1.1):
        if not text: return
        clean_text = text.translate(self.trans_map)
        text_len = len(clean_text)
        if text_len == 0: return
        
        calc_size = max_height / (text_len * line_spacing)
        font_size = max(min(max_font_size, calc_size), 8)
        
        self.c.setFont(self.font_name, font_size)
        current_y = y_start
        char_step = font_size * line_spacing
        
        for char in clean_text:
            self.c.drawCentredString(x, current_y - font_size, char)
            current_y -= char_step

def generate_nengajo_pdf(target_records):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(HAGAKI_WIDTH, HAGAKI_HEIGHT))
    renderer = VerticalTextRendererPDF(c, FONT_NAME)
    
    for record in target_records:
        name = str(record.get("名前", ""))
        full_address = str(record.get("住所", ""))
        
        renmei = str(record.get("連名", "")).strip()
        if renmei.lower() == "nan": renmei = ""

        digits, address = get_zipcode_digits(full_address)

        # 1. 郵便番号
        if len(digits) >= 7:
            c.setFont(FONT_NAME, 14) 
            for i in range(3):
                x = ZIP_X_LEFT_START + (i * ZIP_STEP) + OFFSET_X
                y = ZIP_Y + OFFSET_Y
                c.drawCentredString(x, y, digits[i])
            for i in range(4):
                x = ZIP_X_RIGHT_START + (i * ZIP_STEP) + OFFSET_X
                y = ZIP_Y + OFFSET_Y
                c.drawCentredString(x, y, digits[3+i])

        # 2. 住所
        addr_lines = smart_split_address(address)
        line_configs = [
            (90 * mm, 125 * mm, 16),
            (82 * mm, 118 * mm, 14),
            (75 * mm, 118 * mm, 14)
        ]
        for i, line_text in enumerate(addr_lines):
            if i < len(line_configs):
                Lx, Ly, Lsize = line_configs[i]
                renderer.draw_text(line_text, Lx + OFFSET_X, Ly + OFFSET_Y, 100 * mm, Lsize)

        # 3. 名前・連名
        renmei_list = []
        if renmei:
            # ★★★ 修正ポイント：スペースでの分割を廃止 ★★★
            # 「・」や「,」や「、」だけで分割します。スペースは名前に残ります。
            split_items = re.split(r'[,、・]+', renmei)
            renmei_list = [x.strip() for x in split_items if x.strip()]

        total_people = 1 + len(renmei_list)
        col_spacing = 13 * mm
        start_x = (50 * mm) + ((total_people - 1) * col_spacing / 2) + OFFSET_X
        
        NAME_SIZE = 30 
        NAME_START_Y = 120 * mm 
        
        # 世帯主
        renderer.draw_text(name + " 様", start_x, NAME_START_Y + OFFSET_Y, 95 * mm, NAME_SIZE, line_spacing=1.15)
        
        # 連名
        current_x = start_x
        for r_name in renmei_list:
            current_x -= col_spacing
            renderer.draw_text(r_name + " 様", current_x, NAME_START_Y + OFFSET_Y, 95 * mm, NAME_SIZE, line_spacing=1.15)
        
        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer

def generate_preview_image(name, full_address, renmei=""):
    temp_record = [{"名前": name, "住所": full_address, "連名": renmei}]
    pdf_bytes = generate_nengajo_pdf(temp_record)
    
    doc = fitz.open(stream=pdf_bytes.getvalue(), filetype="pdf")
    page = doc.load_page(0)
    dpi = 300
    pix = page.get_pixmap(dpi=dpi, alpha=True)
    pdf_img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    bg_filename = "hagaki.png"
    base_img = Image.new("RGBA", pdf_img.size, (255, 255, 255, 255))

    if os.path.exists(bg_filename):
        try:
            user_bg = Image.open(bg_filename).convert("RGBA")
            user_bg = user_bg.resize(pdf_img.size, Image.Resampling.LANCZOS)
            base_img = Image.alpha_composite(base_img, user_bg)
        except:
            pass

    px_scale = dpi / 25.4
    shift_x = int(PREVIEW_ADJUST_X_MM * px_scale)
    shift_y = int(PREVIEW_ADJUST_Y_MM * px_scale)

    shifted_layer = Image.new("RGBA", base_img.size, (0, 0, 0, 0))
    shifted_layer.paste(pdf_img, (shift_x, -shift_y), mask=pdf_img) 

    combined = Image.alpha_composite(base_img, shifted_layer)
    return combined.convert("RGB")