import asyncio
import io
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, BufferedInputFile
from aiogram.filters import CommandStart
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from google import genai

BOT_TOKEN = "8408520484:AAEqOi_Ymkh5g524-TVHdGv7yTLAG4Ja_Y4"
GEMINI_API_KEY = "AQ.Ab8RN6Iu6riqs8tEMVX_qo83q8CcYRaNLrCNliWfF9tz4fD3mg"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai_client = genai.Client(api_key=GEMINI_API_KEY)

@dp.message(CommandStart())
async def start_handler(message: Message):
    await message.answer("Salom! Menga slayd mavzusini yuboring, men sizga tayyor prezentatsiya faylini tayyorlab beraman.")

@dp.message(F.text)
async def generate_presentation(message: Message):
    topic = message.text
    wait_msg = await message.answer("Slayd tayyorlanmoqda, iltimos kuting...")

    prompt = f"""
    "{topic}" mavzusida 5 ta slayddan iborat prezentatsiya matnini tuzib ber.
    Har bir slayd uchun sarlavha va 3 ta asosiy punkt bo'lsin.
    Javobni AYNAN mana shu formatda ber, ortiqcha so'z yozma:

    SLAYD: Slayd sarlavhasi
    - Birinchi punkt
    - Ikkinchi punkt
    - Uchinchi punkt
    """

    try:
        response = ai_client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        content = response.text

        prs = Presentation()
        slides_data = content.strip().split("SLAYD:")

        for slide_text in slides_data:
            if not slide_text.strip():
                continue

            lines = slide_text.strip().split("\n")
            title_text = lines[0].strip()
            bullet_points = [line.strip().lstrip("- ") for line in lines[1:] if line.strip()]

            blank_slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_slide_layout)

            # Sarlavha
            tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(8.4), Inches(1.2))
            tf = tx_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = RGBColor(24, 43, 73)

            # Punktlar
            tx_box2 = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(8.4), Inches(4.5))
            tf2 = tx_box2.text_frame
            tf2.word_wrap = True

            for i, point in enumerate(bullet_points):
                p2 = tf2.add_paragraph() if i > 0 else tf2.paragraphs[0]
                p2.text = f"•  {point}"
                p2.font.size = Pt(20)
                p2.font.color.rgb = RGBColor(60, 60, 60)
                p2.space_after = Pt(14)

        pptx_io = io.BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)

        file = BufferedInputFile(pptx_io.read(), filename=f"{topic[:20]}.pptx")
        await message.answer_document(file, caption=f"✨ '{topic}' mavzusidagi slayd tayyor!")
        await wait_msg.delete()

    except Exception as e:
        await message.answer(f"Xatolik yuz berdi: {str(e)}")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())