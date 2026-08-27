import os
import io
import json
import logging
import requests
import telebot
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import google.generativeai as genai

# Logging sozlanishi
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Environment o'zgaruvchilarini olish
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not TELEGRAM_BOT_TOKEN or not GEMINI_API_KEY:
    raise ValueError("TELEGRAM_BOT_TOKEN va GEMINI_API_KEY atrof-muhit o'zgaruvchilarida sozlanishi shart!")

# API larni ishga tushirish
bot = telebot.TeleBot(TELEGRAM_BOT_TOKEN)
genai.configure(api_key=GEMINI_API_KEY)

# Modellarni belgilash (Asosiy va Zaxira modellar)
PRIMARY_MODEL = "gemini-3.5-flash"

SYSTEM_PROMPT = """
Siz professional prezentatsiya yaratuvchi AI assistentsiz. Foydalanuvchi taqdim etgan mavzu bo'yicha aynan 15 ta slayddan iborat structured JSON formatida ma'lumot qaytarishingiz kerak.

Qaytariladigan JSON strukturasi quyidagicha bo'lishi shart:
{
  "presentation_title": "Prezentatsiya sarlavhasi",
  "slides": [
    {
      "slide_number": 1,
      "title": "Slayd sarlavhasi",
      "content": ["Asosiy nuqta 1", "Asosiy nuqta 2", "Asosiy nuqta 3"],
      "image_keyword": "Aynan shu slaydga mos inglizcha bitta-ikkita qidiruv kalit so'zi (masalan: artificial intelligence, modern architecture, space rocket)"
    }
  ]
}

Qoidalar:
1. `slides` massivida aynan 15 ta element bo'lishi shart!
2. Har bir slayd mazmuni boy va tushunarli bo'lishi kerak.
3. Birinchi slayd kirish/sarlavha slaydi, oxirgi 15-slayd esa xulosa slaydi bo'lsin.
4. `image_keyword` aniq va vizual tasvirlanadigan inglizcha so'z bo'lishi shart.
"""

def generate_ai_content(prompt):
    """Gemini API orqali javob olish va 404 xatolarining oldini olish"""
    models_to_try = [PRIMARY_MODEL, "gemini-1.5-flash", "gemini-2.0-flash-exp"]
    
    for model_name in models_to_try:
        try:
            logger.info(f"Model sinab ko'rilmoqda: {model_name}")
            model = genai.GenerativeModel(
                model_name=model_name,
                generation_config={"response_mime_type": "application/json"}
            )
            response = model.generate_content([SYSTEM_PROMPT, prompt])
            return response.text
        except Exception as e:
            logger.warning(f"{model_name} modelida xatolik bo'ldi (404/Not Found bo'lishi mumkin): {e}")
            continue
            
    raise RuntimeError("Barcha Gemini modellari so'rovni bajarishda xatolik berdi.")

def fetch_image(keyword):
    """Unsplash API orqali mavzuga mos rasm yuklab olish"""
    try:
        url = f"https://source.unsplash.com/800x600/?{requests.utils.quote(keyword)}"
        res = requests.get(url, timeout=10)
        if res.status_code == 200:
            return io.BytesIO(res.content)
    except Exception as e:
        logger.error(f"Rasm yuklashda xatolik ({keyword}): {e}")
    
    try:
        fallback_res = requests.get("https://picsum.photos/800/600", timeout=10)
        if fallback_res.status_code == 200:
            return io.BytesIO(fallback_res.content)
    except Exception:
        return None

def build_powerpoint(data):
    """JSON ma'lumotlaridan slaydlardan iborat .pptx yaratish"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]
    slides_data = data.get("slides", [])
    total_slides = len(slides_data)

    for index, item in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)

        # 1. Slayd foni va Yuqori rangli chiziq
        top_bar = slide.shapes.add_shape(
            1, Inches(0), Inches(0), Inches(13.333), Inches(0.4)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(30, 64, 175)
        top_bar.line.fill.background()

        # 2. Sarlavha
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(1.0))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"{index + 1}. {item.get('title', 'Slayd')}"
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)

        # 3. Matn qismi
        text_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(6.5), Inches(4.8))
        tf_content = text_box.text_frame
        tf_content.word_wrap = True

        bullet_points = item.get("content", [])
        for i, point in enumerate(bullet_points):
            p_cont = tf_content.add_paragraph() if i > 0 else tf_content.paragraphs[0]
            p_cont.text = f"•  {point}"
            p_cont.font.size = Pt(18)
            p_cont.font.color.rgb = RGBColor(51, 65, 85)
            p_cont.space_after = Pt(14)

        # 4. Rasm qismi
        keyword = item.get("image_keyword", "presentation")
        img_stream = fetch_image(keyword)

        if img_stream:
            try:
                slide.shapes.add_picture(
                    img_stream, Inches(7.6), Inches(2.0), width=Inches(4.9), height=Inches(4.5)
                )
            except Exception as e:
                logger.error(f"Slaydga rasm joylashda xatolik: {e}")

        # Footer
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"{data.get('presentation_title', 'Taqdimot')} | Slayd {index + 1} / {total_slides}"
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(148, 163, 184)
        p_foot.alignment = PP_ALIGN.RIGHT

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(
        message,
        "Assalomu alaykum! Men sun'iy intellekt yordamida prezentatsiyalar tayyorlab beruvchi botman.\n\n"
        "Menga prezentatsiya mavzusini yuboring, men sizga mos tayyor PowerPoint taqdimotini yaratib beraman.\n\n"
        "Masalan:\n"
        "👉 *Sun'iy intellektning kelajagi*\n"
        "👉 *Ekologiya va atrof-muhit muhofazasi*"
    )

@bot.message_handler(func=lambda message: True)
def generate_presentation_handler(message):
    topic = message.text.strip()
    status_msg = bot.reply_to(message, "⏳ **Gemini AI slaydlaringizni tayyorlamoqda...**\n(Bu biroz vaqt olishi mumkin)")

    try:
        prompt = f"Mavzu: {topic}\nUshbu mavzu bo'yicha prezentatsiya ma'lumotlarini tayyorla."
        
        # 404 xatoligi kelib chiqmasligi uchun xavfsiz chaqiruv
        raw_json_text = generate_ai_content(prompt)
        data = json.loads(raw_json_text)
        
        bot.edit_message_text("🎨 **Slaydlar va rasmlar jamlanmoqda... PowerPoint (.pptx) fayli yaratilmoqda...**", 
                              chat_id=status_msg.chat.id, message_id=status_msg.message_id)

        pptx_file = build_powerpoint(data)

        filename = f"{topic[:20].strip()}_prezentatsiya.pptx"
        bot.send_document(
            message.chat.id,
            document=(filename, pptx_file, "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
            caption=f"✅ **Prezentatsiyangiz tayyor!**\n📌 Mavzu: *{topic}*"
        )
        bot.delete_message(chat_id=status_msg.chat.id, message_id=status_msg.message_id)

    except Exception as e:
        logger.error(f"Xatolik yuz berdi: {e}")
        bot.edit_message_text(
            "❌ Kechirasiz, prezentatsiya yaratishda xatolik yuz berdi. Iltimos, qaytadan urinib ko'ring yoki boshqa mavzu yozing.",
            chat_id=status_msg.chat.id,
            message_id=status_msg.message_id
        )

if __name__ == "__main__":
    logger.info("Bot ishga tushdi...")
    bot.infinity_polling()
