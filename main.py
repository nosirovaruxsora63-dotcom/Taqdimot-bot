import asyncio
import os
from aiogram import Bot, Dispatcher
from aiogram.types import Message
from aiogram.filters import CommandStart
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from google import genai
from aiohttp import web

# API Kalitlar
BOT_TOKEN = "8408520484:AAEqOi_Ymkh5g524-TVHdGv7yTLAG4Ja_Y4" 
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "AQ.Ab8RN6Iu6riqs8tEMVX_qo83q8CcYRaNLrCNliWfF9tz4fD3mg")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai_client = genai.Client(api_key=GEMINI_API_KEY)

@dp.message(CommandStart())
async def start_handler(message: Message):
    await message.answer("Salom! Menga slayd mavzusini yuboring, men sizga tayyor prezentatsiya faylini tayyorlab beraman.")

@dp.message()
async def generate_slide(message: Message):
    await message.answer("Slayd tayyorlanmoqda, iltimos kuting...")
    try:
        prompt = f"'{message.text}' mavzusida 5 ta slayd uchun qisqa va mazmunli matn tayyorlab ber."
        response = ai_client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt
        )
        
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        tf = txBox.text_frame
        tf.text = response.text
        
        file_path = "taqdimot.pptx"
        prs.save(file_path)
        
        from aiogram.types import FSInputFile
        doc = FSInputFile(file_path)
        await message.answer_document(doc, caption="Taqdimotingiz tayyor!")
    except Exception as e:
        await message.answer(f"Xatolik yuz berdi: {e}")

# Render portini aldash uchun veb-server
async def handle(request):
    return web.Response(text="Bot ishlayapti!")

async def main():
    app = web.Application()
    app.router.add_get("/", handle)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", 10000))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
