import os
import asyncio
import io
import urllib.parse
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.types import BufferedInputFile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from google import genai
import aiohttp

# --- TOKENS AND API KEYS ---
BOT_TOKEN = os.environ.get("BOT_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai_client = genai.Client(api_key=GEMINI_API_KEY)

# --- HELPER: UN SPLASH OR PLACEHOLDER IMAGE DOWNLOADER ---
async def fetch_image_bytes(query: str) -> bytes:
    """Mavzuga mos rasm qidirib yuklab beradi"""
    encoded_query = urllib.parse.quote(query)
    # Unsplash manbasidan mavzuga mos rasm olinadi
    url = f"https://source.unsplash.com/800x600/?{encoded_query}"
    
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, timeout=10) as resp:
                if resp.status == 200:
                    return await resp.read()
    except Exception as e:
        print(f"Rasm yuklashda xatolik: {e}")
    
    # Agar rasm topilmasa zaxira rasm
    fallback_url = "https://picsum.photos/800/600"
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(fallback_url, timeout=10) as resp:
                if resp.status == 200:
                    return await resp.read()
    except Exception:
        pass
    return None

# --- BOT HANDLERS ---
@dp.message(CommandStart())
async def start_cmd(message: types.Message):
    await message.answer(
        "Salom! Men taqdimot tayyorlab beruvchi botman.\n\n"
        "Menga taqdimot mavzusini yuboring, men har bir slaydda **rasmlar va chiroyli dizayn** bilan `.pptx` fayl tayyorlab beraman!"
    )

@dp.message(F.text)
async def generate_presentation(message: types.Message):
    topic = message.text.strip()
    status_msg = await message.answer("⏳ Taqdimot va rasmlar tayyorlanmoqda, iltimos kuting...")

    prompt = f"""
Siz professional taqdimot dizaynerisiz.
Mavzu: "{topic}"

Menga 15 ta slayddan iborat taqdimot matnini yarating. 
Har bir slayd uchun aniq quyidagi formatda javob bering:

SLIDE_TITLE: Slayd sarlavhasi
IMAGE_SEARCH: Rasm qidirish uchun inglizcha kalit so'z (masalan: artificial intelligence, business meeting, nature va h.k.)
SLIDE_CONTENT: 
- Birinchi muhim punkt
- Ikkinchi muhim punkt
- Uchinchi muhim punkt
---
"""

    try:
        # Gemini 3.6 Flash modeli orqali matn va rasm kalit so'zlarini olish
        response = ai_client.models.generate_content(
            model="gemini-3.6-flash",
            contents=prompt,
        )
        raw_text = response.text

        prs = Presentation()
        # Slayd o'lchamini 16:9 (keng) qilish
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

        blank_layout = prs.slide_layouts[6]
        slides_data = raw_text.split("---")

        for slide_text in slides_data:
            if not slide_text.strip():
                continue

            lines = slide_text.strip().split("\n")
            title = "Taqdimot"
            img_keyword = topic
            content_lines = []

            for line in lines:
                if line.startswith("SLIDE_TITLE:"):
                    title = line.replace("SLIDE_TITLE:", "").strip()
                elif line.startswith("IMAGE_SEARCH:"):
                    img_keyword = line.replace("IMAGE_SEARCH:", "").strip()
                elif line.startswith("SLIDE_CONTENT:"):
                    continue
                elif line.strip():
                    content_lines.append(line.strip())

            slide = prs.slides.add_slide(blank_layout)

            # 1. Sarlavha qo'shish
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = RGBColor(24, 43, 73)

            # 2. Matn qismini chap tomonga qo'shish
            content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.5), Inches(5.0))
            ctf = content_box.text_frame
            ctf.word_wrap = True
            
            for idx, text_line in enumerate(content_lines):
                cp = ctf.add_paragraph() if idx > 0 else ctf.paragraphs[0]
                cp.text = text_line
                cp.font.size = Pt(18)
                cp.font.color.rgb = RGBColor(50, 50, 50)
                cp.space_after = Pt(12)

            # 3. Rasmni o'ng tomonga joylashtirish
            img_bytes = await fetch_image_bytes(img_keyword)
            if img_bytes:
                image_stream = io.BytesIO(img_bytes)
                slide.shapes.add_picture(
                    image_stream, 
                    left=Inches(7.8), 
                    top=Inches(1.8), 
                    width=Inches(4.7), 
                    height=Inches(4.5)
                )

        # Faylni xotirada saqlash
        pptx_io = io.BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)

        input_file = BufferedInputFile(pptx_io.read(), filename=f"{topic}.pptx")
        await message.answer_document(input_file, caption=f"✨ **{topic}** bo'yicha rasmli taqdimotingiz tayyor!")
        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"Xatolik yuz berdi: {e}")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
