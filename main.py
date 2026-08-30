import os
import json
import re
import asyncio
import requests
from io import BytesIO
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from pptx import Presentation
from pptx.util import Inches, Pt
from openai import OpenAI

BOT_TOKEN = os.getenv("BOT_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

# Groq (Llama 3) mijozi - 404 xatoliklarsiz ishlaydi
client = OpenAI(
    base_url="https://api.groq.com/openai/v1",
    api_key=GROQ_API_KEY,
)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

def create_ppt(data, filename="Taqdimot.pptx"):
    prs = Presentation()

    # 1-Slayd: Bosh sarlavha
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    slide1.shapes.title.text = data.get("title", "Taqdimot")
    slide1.placeholders[1].text = "Sun'iy intellekt tomonidan tayyorlandi"

    # 2-Slayd: Reja
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "Taqdimot Rejasi"
    reja_items = data.get("reja", [])
    reja_text = "\n".join([f"{i+1}. {item}" for i, item in enumerate(reja_items)])
    slide2.placeholders[1].text = reja_text

    # 3-dan 15-gacha Slaydlar (13 ta asosiy slayd)
    slides_data = data.get("slides", [])
    for index, item in enumerate(slides_data[:13]):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"{index + 1}. {item.get('title', '')}"

        # Matn qismi
        txBox = slide.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(4.5), Inches(4.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = item.get("content", "")
        p.font.size = Pt(16)

        # Rasm qismi
        keyword = item.get("image_keyword", "technology")
        img_url = f"https://image.pollinations.ai/prompt/{keyword}?width=600&height=500&nologo=true"
        
        try:
            res = requests.get(img_url, timeout=12)
            if res.status_code == 200:
                img_stream = BytesIO(res.content)
                slide.shapes.add_picture(img_stream, Inches(5.3), Inches(1.8), width=Inches(4.2))
        except Exception as img_err:
            print(f"Rasm yuklashda xato: {img_err}")

    prs.save(filename)
    return filename

@dp.message(Command("start"))
async def start_handler(message: types.Message):
    await message.answer("Assalomu alaykum! Mavzu yuboring. Men sizga 15 ta slayddan iborat, reja va rasmli taqdimot tayyorlab beraman.")

@dp.message()
async def generate_presentation_handler(message: types.Message):
    topic = message.text
    wait_msg = await message.answer(f"⏳ '{topic}' mavzusida 15 ta slayd tayyorlanmoqda...")

    prompt = f"""
    Menga '{topic}' mavzusida taqdimot uchun ma'lumotlarni faqat JSON formatida qaytar.
    Qoidalar:
    1. "title": Taqdimotning umumiy nomi (o'zbek tilida).
    2. "reja": Taqdimotning 5-6 ta banddan iborat rejasi (massiv ko'rinishida).
    3. "slides": Aniq 13 ta obyekt elementidan iborat massiv.
       Har bir slayd obyektida:
       - "title": Slayd sarlavhasi
       - "content": Slaydning asosiy mazmuni (3-5 ta batafsil gap)
       - "image_keyword": Slayd mazmuniga mos inglizcha 2-3 so'zdan iborat rasm kalit so'zi.
       
    Faqat JSON qaytar, boshqa hech qanday izoh yozma.
    """

    try:
        # Groq orqali Llama 3 modelidan foydalanish (404 xatolik chiqmaydi)
        completion = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        
        raw_response = completion.choices[0].message.content
        data = json.loads(raw_response)

        safe_topic = re.sub(r'[^\w\s-]', '', topic)[:15].strip()
        filename = f"{safe_topic}_taqdimot.pptx"
        create_ppt(data, filename)

        ppt_file = types.FSInputFile(filename)
        await message.answer_document(document=ppt_file, caption=f"✅ '{topic}' mavzusidagi 15 ta slayddan iborat taqdimot tayyor!")
        
        if os.path.exists(filename):
            os.remove(filename)

    except Exception as e:
        await message.answer(f"❌ Xatolik yuz berdi: {str(e)}\nIltimos, qayta urinib ko'ring.")
    finally:
        try:
            await wait_msg.delete()
        except:
            pass

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
