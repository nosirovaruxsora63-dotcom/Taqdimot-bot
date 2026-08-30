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
import google.generativeai as genai

BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

genai.configure(api_key=GEMINI_API_KEY)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# 404 xatolik bermaydigai aqlli Gemini chaqirish funksiyasi
def generate_gemini_content(prompt: str) -> str:
    # Ishlashi mumkin bo'lgan modellar ketma-ketligi
    models_to_try = [
        'gemini-2.0-flash',
        'gemini-1.5-flash',
        'gemini-1.5-pro',
        'gemini-pro'
    ]
    
    last_error = None
    for model_name in models_to_try:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            if response.text:
                return response.text
        except Exception as e:
            last_error = e
            # Agar 404 yoki model topilmadi xatosi bo'lsa, keyingi modelga o'tadi
            continue
            
    raise Exception(f"Barcha modellarda xatolik yuz berdi: {last_error}")

# 15 ta slayddan iborat PowerPoint yaratish
def create_ppt(data, filename="Taqdimot.pptx"):
    prs = Presentation()

    # 1-Slayd: Bosh sarlavha (Mavzu)
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    slide1.shapes.title.text = data.get("title", "Taqdimot")
    slide1.placeholders[1].text = "Sun'iy intellekt tomonidan tayyorlandi"

    # 2-Slayd: Reja
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "Taqdimot Rejasi"
    reja_items = data.get("reja", [])
    reja_text = "\n".join([f"{i+1}. {item}" for i, item in enumerate(reja_items)])
    slide2.placeholders[1].text = reja_text

    # 3-dan 15-gacha Slaydlar (Jami 13 ta asosiy slayd)
    slides_data = data.get("slides", [])
    for index, item in enumerate(slides_data[:13]):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"{index + 1}. {item.get('title', '')}"

        # Matn bloki (Chap tomonda)
        txBox = slide.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(4.5), Inches(4.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = item.get("content", "")
        p.font.size = Pt(16)

        # Rasm bloki (O'ng tomonda mavzuga mos rasm)
        keyword = item.get("image_keyword", "nature")
        img_url = f"https://image.pollinations.ai/prompt/{keyword}?width=600&height=500&nologo=true"
        
        try:
            res = requests.get(img_url, timeout=12)
            if res.status_code == 200:
                img_stream = BytesIO(res.content)
                slide.shapes.add_picture(img_stream, Inches(5.3), Inches(1.8), width=Inches(4.2))
        except Exception as img_err:
            print(f"Rasm yuklashda xato (o'tkazib yuborildi): {img_err}")

    prs.save(filename)
    return filename

@dp.message(Command("start"))
async def start_handler(message: types.Message):
    await message.answer("Assalomu alaykum! Taqdimot mavzusini yuboring. Men sizga 15 ta slayddan iborat, rejasili va rasmli taqdimot tayyorlab beraman.")

@dp.message()
async def generate_presentation_handler(message: types.Message):
    topic = message.text
    wait_msg = await message.answer(f"⏳ '{topic}' mavzusida 15 ta slayd, reja va rasmlar tayyorlanmoqda. Bu biroz vaqt olishi mumkin...")

    prompt = f"""
    Menga '{topic}' mavzusida taqdimot uchun ma'lumotlarni faqat to'g'ri va yaroqli JSON formatida qaytar.
    
    Qoidalar va struktura:
    1. "title": Taqdimotning umumiy nomi (o'zbek tilida).
    2. "reja": Taqdimotning 5-6 ta banddan iborat rejasi (massiv ko'rinishida).
    3. "slides": Aniq 13 ta obyekt apparatidan iborat massiv (chunki 1-slayd sarlavha, 2-slayd reja, keyingi 13 ta slayd bilan jami 15 ta bo'ladi).
       Har bir slayd obyektida:
       - "title": Slayd sarlavhasi
       - "content": Slaydning asosiy mazmuni (3-5 ta batafsil gap)
       - "image_keyword": Ushbu slayd mazmuniga mos keladigan inglizcha 2-3 so'zdan iborat rasm kalit so'zi (masalan: "medical doctor laboratory", "cloud computing server").
       
    Faqat JSON formatda javob ber, hech qanday qo'shimcha matn yoki izoh yozma.
    """

    try:
        # Gemini'dan xatosiz javob olish
        raw_response = generate_gemini_content(prompt)
        
        # JSON ni matn ichidan toza ajratib olish
        json_match = re.search(r'\{.*\}', raw_response, re.DOTALL)
        if json_match:
            clean_json = json_match.group(0)
        else:
            clean_json = raw_response

        data = json.loads(clean_json)

        # Taqdimot faylini yaratish
        safe_topic = re.sub(r'[^\w\s-]', '', topic)[:15].strip()
        filename = f"{safe_topic}_taqdimot.pptx"
        create_ppt(data, filename)

        # Telegramga yuborish
        ppt_file = types.FSInputFile(filename)
        await message.answer_document(document=ppt_file, caption=f"✅ '{topic}' mavzusidagi 15 ta slayddan iborat rasmli taqdimotingiz tayyor!")
        
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
