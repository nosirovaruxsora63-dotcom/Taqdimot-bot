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
import google.generativeai as genai
import aiohttp

# --- TOKENS AND API KEYS ---
BOT_TOKEN = os.environ.get("BOT_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# Gemini sozlamasi
genai.configure(api_key=GEMINI_API_KEY)

async def fetch_image_bytes(query: str) -> bytes:
    encoded_query = urllib.parse.quote(query)
    url = f"https://picsum.photos/800/600"
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, timeout=10) as resp:
                if resp.status == 200:
                    return await resp.read()
    except Exception:
        pass
    return None

@dp.message(CommandStart())
async def start_cmd(message: types.Message):
    await message.answer(
        "Salom! Men taqdimot tayyorlab beruvchi botman.\n\n"
        "Menga taqdimot mavzusini yuboring, men 15 ta slayddan iborat rasmli taqdimot tayyorlab beraman!"
    )

@dp.message(F.text)
async def generate_presentation(message: types.Message):
    topic = message.text.strip()
    status_msg = await message.answer("⏳ Taqdimot tayyorlanmoqda, iltimos kuting...")

    prompt = f"""
Siz professional taqdimot dizaynerisiz.
Mavzu: "{topic}"

Menga 15 ta slayddan iborat taqdimot matnini yarating. 
Har bir slayd uchun aniq quyidagi formatda javob bering:

SLIDE_TITLE: Slayd sarlavhasi
IMAGE_SEARCH: Rasm qidirish uchun kalit so'z
SLIDE_CONTENT: 
- Birinchi muhim punkt
- Ikkinchi muhim punkt
- Uchinchi muhim punkt
---
"""

    try:
        # Rasmiy va ishlaydigan model
        model = genai.GenerativeModel('gemini-2.0-flash')
        response = model.generate_content(prompt)
        raw_text = response.text

        prs = Presentation()
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

            # Sarlavha
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = RGBColor(24, 43, 73)

            # Matn
            content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.5), Inches(5.0))
            ctf = content_box.text_frame
            ctf.word_wrap = True
            
            for idx, text_line in enumerate(content_lines):
                cp = ctf.add_paragraph() if idx > 0 else ctf.paragraphs[0]
                cp.text = text_line
                cp.font.size = Pt(16)
                cp.font.color.rgb = RGBColor(50, 50, 50)
                cp.space_after = Pt(10)

            # Rasm
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

        pptx_io = io.BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)

        input_file = BufferedInputFile(pptx_io.read(), filename=f"{topic}.pptx")
        await message.answer_document(input_file, caption=f"✨ **{topic}** bo'yicha taqdimotingiz tayyor!")
        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"Xatolik yuz berdi: {e}")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
