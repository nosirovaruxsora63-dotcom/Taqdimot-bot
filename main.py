import os
import io
import json
import logging
import aiohttp
import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import BufferedInputFile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from google import genai

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not TELEGRAM_BOT_TOKEN or not GEMINI_API_KEY:
    raise ValueError("TELEGRAM_BOT_TOKEN va GEMINI_API_KEY sozlanmagan!")

bot = Bot(token=TELEGRAM_BOT_TOKEN)
dp = Dispatcher()

# Yangi rasmiy SDK mijozini yaratish
client = genai.Client(api_key=GEMINI_API_KEY)

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
      "image_keyword": "Aynan shu slaydga mos inglizcha bitta kalit so'z (masalan: medicine, heart, surgery)"
    }
  ]
}

Qoidalar:
1. `slides` massivida aynan 15 ta element bo'lishi shart!
2. Birinchi slayd kirish, 15-slayd xulosa bo'lsin.
3. Javob faqat va faqat to'g'ri JSON formatida bo'lsin.
"""

async def generate_ai_content(prompt):
    """Google'ning eng yangi va barqaror modellari orqali xatosiz ma'lumot olish"""
    # Yangi SDK uchun rasmiy va ishlaydigan model nomlari
    models_to_try = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    
    for model_name in models_to_try:
        try:
            logger.info(f"Model sinab ko'rilmoqda: {model_name}")
            
            # Asinxron chaqiruv
            response = await asyncio.to_thread(
                client.models.generate_content,
                model=model_name,
                contents=f"{SYSTEM_PROMPT}\n\nMavzu: {prompt}",
            )
            
            if response.text:
                # JSON matnidan ortiqcha ```json belgilarini tozalaymiz
                text = response.text.strip()
                if text.startswith("```json"):
                    text = text[7:]
                if text.startswith("```"):
                    text = text[3:]
                if text.endswith("```"):
                    text = text[:-3]
                return text.strip()
        except Exception as e:
            logger.warning(f"{model_name} modelida xato: {e}")
            continue

    raise RuntimeError("404 yoki boshqa API xatolar tufayli biror model ishlamadi. API Kalitni tekshiring.")

async def fetch_image(keyword):
    """Rasm yuklab olish"""
    url = f"https://picsum.photos/800/600"
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, timeout=10) as resp:
                if resp.status == 200:
                    data = await resp.read()
                    return io.BytesIO(data)
    except Exception as e:
        logger.error(f"Rasm yuklashda xato: {e}")
    return None

async def build_powerpoint(data):
    """PowerPoint (.pptx) faylini yaratish"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    slides_data = data.get("slides", [])
    total_slides = len(slides_data)

    for index, item in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)

        # Yuqori ko'k chiziq
        top_bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.4))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(30, 64, 175)
        top_bar.line.fill.background()

        # Sarlavha
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(1.0))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"{index + 1}. {item.get('title', 'Slayd')}"
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)

        # Matn bo'limi
        text_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(6.5), Inches(4.8))
        tf_content = text_box.text_frame
        tf_content.word_wrap = True

        bullet_points = item.get("content", [])
        for i, point in enumerate(bullet_points):
            p_cont = tf_content.add_paragraph() if i > 0 else tf_content.paragraphs[0]
            p_cont.text = f"• {point}"
            p_cont.font.size = Pt(18)
            p_cont.font.color.rgb = RGBColor(51, 65, 85)
            p_cont.space_after = Pt(12)

        # Rasm qo'shish
        keyword = item.get("image_keyword", "presentation")
        img_stream = await fetch_image(keyword)
        if img_stream:
            try:
                slide.shapes.add_picture(img_stream, Inches(7.6), Inches(2.0), width=Inches(4.9), height=Inches(4.5))
            except Exception as e:
                logger.error(f"Rasm joylashda xato: {e}")

        # Footer (Pasti)
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

@dp.message(Command("start", "help"))
async def start_handler(message: types.Message):
    await message.answer("Salom! Menga prezentatsiya mavzusini yuboring, men sizga 15 slayddan iborat tayyor PowerPoint (.pptx) fayl tayyorlab beraman.")

@dp.message()
async def generate_handler(message: types.Message):
    topic = message.text.strip()
    status_msg = await message.answer("⏳ Slayd mazmuni va JSON tayyorlanmoqda...")

    try:
        raw_json = await generate_ai_content(topic)
        data = json.loads(raw_json)

        await bot.edit_message_text(
            "🎨 Slaydlar shakllantirilmoqda va rasmlar biriktirilmoqda...",
            chat_id=status_msg.chat.id,
            message_id=status_msg.message_id
        )

        pptx_stream = await build_powerpoint(data)
        file_bytes = pptx_stream.getvalue()
        filename = f"{topic[:20].strip()}_prezentatsiya.pptx"

        input_file = BufferedInputFile(file_bytes, filename=filename)

        await message.answer_document(
            document=input_file,
            caption=f"✅ **Prezentatsiyangiz tayyor!**\n📌 Mavzu: *{topic}*"
        )
        await bot.delete_message(chat_id=status_msg.chat.id, message_id=status_msg.message_id)

    except Exception as e:
        logger.error(f"Xatolik: {e}")
        await message.answer(f"❌ Xatolik yuz berdi: {e}")

async def main():
    logger.info("Bot ishga tushdi...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
