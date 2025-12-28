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
# Google Fontsなどで入手した「brush.ttf」があれば使用
CUSTOM_FONT_FILE = "brush.ttf"
FONT_NAME = "KakizomeFont"

# ==========================================
# 📐 印刷用の位置設定 (成功した設定に戻しました！)
# ==========================================
# 1. 全体のズレ調整
OFFSET_X = 0.7 * mm 
OFFSET_Y = 1.3 * mm 

# 2. 郵便番号（上端からの位置）
ZIP_Y = 148.0 * mm - 15.8 * mm 

# 3. 数字の間隔
ZIP_STEP = 7.3 * mm 

# 4. 左3桁の開始位置
ZIP_X_LEFT_START = 46.0 * mm 

# 5. 右4桁の開始位置
# ★ここを「印刷で成功した計算式」に戻しました
ZIP_X_RIGHT_START = ZIP_X_LEFT_START + (3 * ZIP_STEP) + (0.6 * mm)

HAGAKI_WIDTH = 100 * mm
HAGAKI_HEIGHT = 148 * mm

# ==========================================
# 📺 プレビュー画面専用の調整 (印刷には影響しません)
# ==========================================
# 印刷は合っているのに画面だけズレる場合は、ここの数字で調整してください。
# ※プラスにすると右/下へ、マイナスにすると左/上へ動きます

PREVIEW_ADJUST_X_MM = 0.0  # 例: 画面上で右に7.5mmほどズラして表示
PREVIEW_ADJUST_Y_MM = -4.0 

# ==========================================

# フォント登録ロジック
try:
    if os.path.exists(CUSTOM_FONT_FILE):
        pdfmetrics.registerFont(TTFont(FONT_NAME, CUSTOM_FONT_FILE))
        print(f"成功: {CUSTOM_FONT_FILE} を読み込みました。")
    else:
        print(f"警告: {CUSTOM_FONT_FILE} が見つかりません。標準フォントを使います。")
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))
        FONT_NAME = "HeiseiMin-W3"
except Exception as e:
    print(f"フォントエラー: {e}")
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

def split_address(address):
    if len(address) < 16: return [address]
    keywords = ["区", "市", "郡"]
    split_index = 16
    for kw in keywords:
        idx = address.find(kw)
        if idx > 3 and (len(address) - idx) > 5:
            split_index = idx + 1
            break
    return [address[:split_index], address[split_index:]]

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
    """印刷用PDF作成（ここは印刷成功時のまま！）"""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(HAGAKI_WIDTH, HAGAKI_HEIGHT))
    renderer = VerticalTextRendererPDF(c, FONT_NAME)
    
    for record in target_records:
        name = str(record.get("名前", ""))
        full_address = str(record.get("住所", ""))
        digits, address = get_zipcode_digits(full_address)

        # 1. 郵便番号 (14pt)
        if len(digits) >= 7:
            c.setFont(FONT_NAME, 14) 
            # 左3桁
            for i in range(3):
                x = ZIP_X_LEFT_START + (i * ZIP_STEP) + OFFSET_X
                y = ZIP_Y + OFFSET_Y
                c.drawCentredString(x, y, digits[i])
            # 右4桁
            for i in range(4):
                x = ZIP_X_RIGHT_START + (i * ZIP_STEP) + OFFSET_X
                y = ZIP_Y + OFFSET_Y
                c.drawCentredString(x, y, digits[3+i])

        # 2. 住所
        addr_lines = split_address(address)
        renderer.draw_text(addr_lines[0], 90 * mm + OFFSET_X, 125 * mm + OFFSET_Y, 100 * mm, 16)
        if len(addr_lines) > 1:
            renderer.draw_text(addr_lines[1], 82 * mm + OFFSET_X, 125 * mm + OFFSET_Y, 100 * mm, 14)

        # 3. 名前 (34pt)
        renderer.draw_text(name + " 様", 50 * mm + OFFSET_X, 115 * mm + OFFSET_Y, 95 * mm, 34, line_spacing=1.15)
        
        c.showPage()

    c.save()
    buffer.seek(0)
    return buffer

def generate_preview_image(name, full_address):
    """
    プレビュー画像生成（印刷PDFを作ってから、画面用に位置をズラす）
    """
    # 1. 印刷と同じPDFデータを作成
    temp_record = [{"名前": name, "住所": full_address}]
    pdf_bytes = generate_nengajo_pdf(temp_record)
    
    # 2. 画像化
    doc = fitz.open(stream=pdf_bytes.getvalue(), filetype="pdf")
    page = doc.load_page(0)
    dpi = 300
    pix = page.get_pixmap(dpi=dpi, alpha=True)
    pdf_img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    # 3. 背景準備
    bg_filename = "hagaki.png"
    base_img = Image.new("RGBA", pdf_img.size, (255, 255, 255, 255))

    if os.path.exists(bg_filename):
        try:
            user_bg = Image.open(bg_filename).convert("RGBA")
            user_bg = user_bg.resize(pdf_img.size, Image.Resampling.LANCZOS)
            base_img = Image.alpha_composite(base_img, user_bg)
        except:
            pass

    # 4. プレビュー専用の位置補正
    # 印刷設定(OFFSET)はいじらず、画面表示だけここでズラします
    px_scale = dpi / 25.4
    shift_x = int(PREVIEW_ADJUST_X_MM * px_scale)
    shift_y = int(PREVIEW_ADJUST_Y_MM * px_scale)

    shifted_layer = Image.new("RGBA", base_img.size, (0, 0, 0, 0))
    shifted_layer.paste(pdf_img, (shift_x, -shift_y), mask=pdf_img) 

    combined = Image.alpha_composite(base_img, shifted_layer)
    return combined.convert("RGB")