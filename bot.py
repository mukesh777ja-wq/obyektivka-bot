import os
import asyncio
from pathlib import Path

from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, FSInputFile
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

BOT_TOKEN = os.environ["BOT_TOKEN"]
OUT = Path("output")
OUT.mkdir(exist_ok=True)

class Form(StatesGroup):
    fio = State()
    birth = State()
    birth_place = State()
    nationality = State()
    party = State()
    education = State()
    graduated = State()
    specialty = State()
    degree = State()
    languages = State()
    awards = State()
    elected = State()
    work = State()
    photo = State()
    relative = State()

def no_kb():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="YO'Q")]],
        resize_keyboard=True
    )

def set_cell(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(str(text))
    r.bold = bold
    r.font.name = "Times New Roman"
    r.font.size = Pt(8)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def make_doc(data, filename):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2)
    sec.right_margin = Cm(1.5)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MA’LUMOTNOMA")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(data["fio"])
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)

    fields = [
        ("Tug‘ilgan yili:", data["birth"]),
        ("Tug‘ilgan joyi:", data["birth_place"]),
        ("Millati:", data["nationality"]),
        ("Partiyaviyligi:", data["party"]),
        ("Ma’lumoti:", data["education"]),
        ("Tamomlagan:", data["graduated"]),
        ("Ma’lumoti bo‘yicha mutaxassisligi:", data["specialty"]),
        ("Ilmiy darajasi:", data["degree"]),
        ("Qaysi chet tillarini biladi:", data["languages"]),
        ("Davlat mukofotlari bilan taqdirlanganligi:", data["awards"]),
        ("Saylanadigan organlarda ishtiroki:", data["elected"]),
    ]
    for label, value in fields:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(label + " ")
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r = p.add_run(value)
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MEHNAT FAOLIYATI")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)

    for work in data.get("work", []):
        p = doc.add_paragraph(work)
        p.paragraph_format.space_after = Pt(1)

    doc.add_page_break()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f'{data["fio"]}ning yaqin qarindoshlari haqida')
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MA’LUMOT")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Qarindoshligi", "Familiyasi, ismi va otasining ismi",
               "Tug‘ilgan yili va joyi", "Ish joyi va lavozimi", "Turar joyi"]
    for i, h in enumerate(headers):
        set_cell(table.rows[0].cells[i], h, True)

    for rel in data.get("relatives", []):
        cells = table.add_row().cells
        vals = [rel[0], rel[1], rel[2], rel[3], rel[4]]
        for i, v in enumerate(vals):
            set_cell(cells[i], v)

    doc.save(filename)

dp = Dispatcher()

@dp.message(CommandStart())
async def start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(Form.fio)
    await message.answer(
        "Assalomu alaykum! Ma’lumotnoma tuzishni boshlaymiz.\n\n"
        "To‘liq FIO (Familiya Ism Sharifingiz) kiriting:"
    )

@dp.message(Form.fio)
async def fio(message: Message, state: FSMContext):
    await state.update_data(fio=message.text.strip())
    await state.set_state(Form.birth)
    await message.answer("Tug‘ilgan sana (kun.oy.yil)ni kiriting:")

@dp.message(Form.birth)
async def birth(message: Message, state: FSMContext):
    await state.update_data(birth=message.text.strip())
    await state.set_state(Form.birth_place)
    await message.answer("Tug‘ilgan joyingizni kiriting:")

@dp.message(Form.birth_place)
async def birth_place(message: Message, state: FSMContext):
    await state.update_data(birth_place=message.text.strip())
    await state.set_state(Form.nationality)
    await message.answer("Millatingizni kiriting:")

@dp.message(Form.nationality)
async def nationality(message: Message, state: FSMContext):
    await state.update_data(nationality=message.text.strip())
    await state.set_state(Form.party)
    await message.answer("Partiyaviyligingizni kiriting yoki YO‘Q bosing:", reply_markup=no_kb())

@dp.message(Form.party)
async def party(message: Message, state: FSMContext):
    await state.update_data(party=message.text.strip())
    await state.set_state(Form.education)
    await message.answer("Ma’lumotingizni kiriting:")

@dp.message(Form.education)
async def education(message: Message, state: FSMContext):
    await state.update_data(education=message.text.strip())
    await state.set_state(Form.graduated)
    await message.answer("Qaysi o‘quv yurtini qachon tamomlagansiz?")

@dp.message(Form.graduated)
async def graduated(message: Message, state: FSMContext):
    await state.update_data(graduated=message.text.strip())
    await state.set_state(Form.specialty)
    await message.answer("Ma’lumotingiz bo‘yicha mutaxassisligingizni kiriting:")

@dp.message(Form.specialty)
async def specialty(message: Message, state: FSMContext):
    await state.update_data(specialty=message.text.strip())
    await state.set_state(Form.degree)
    await message.answer("Ilmiy darajangizni kiriting yoki YO‘Q bosing:")

@dp.message(Form.degree)
async def degree(message: Message, state: FSMContext):
    await state.update_data(degree=message.text.strip())
    await state.set_state(Form.languages)
    await message.answer("Qaysi chet tillarini bilasiz? yoki YO‘Q bosing:")

@dp.message(Form.languages)
async def languages(message: Message, state: FSMContext):
    await state.update_data(languages=message.text.strip())
    await state.set_state(Form.awards)
    await message.answer("Davlat mukofotlari bilan taqdirlanganmisiz? Qaysi? yoki YO‘Q:")

@dp.message(Form.awards)
async def awards(message: Message, state: FSMContext):
    await state.update_data(awards=message.text.strip())
    await state.set_state(Form.elected)
    await message.answer("Saylanadigan organlarda deputat yoki a’zo bo‘lganmisiz? yoki YO‘Q:")

@dp.message(Form.elected)
async def elected(message: Message, state: FSMContext):
    await state.update_data(elected=message.text.strip(), work=[])
    await state.set_state(Form.work)
    await message.answer(
        "Mehnat faoliyatingizni kiriting.\n"
        "Masalan: 2018-2020 yy. — Tashkilot — lavozim.\n"
        "Har bir ish joyini alohida yuboring.\n"
        "Tugatish uchun YO‘Q bosing."
    )

@dp.message(Form.work)
async def work(message: Message, state: FSMContext):
    text = message.text.strip()
    data = await state.get_data()
    if text.upper() == "YO'Q" or text.upper() == "YO‘Q":
        await state.set_state(Form.photo)
        await message.answer("3x4 rasmingizni foto sifatida yuboring:")
        return
    arr = data.get("work", [])
    arr.append(text)
    await state.update_data(work=arr)
    await message.answer("Keyingi ish joyini kiriting yoki YO‘Q bosing.")

@dp.message(Form.photo, F.photo)
async def photo(message: Message, state: FSMContext):
    f = await message.bot.get_file(message.photo[-1].file_id)
    path = OUT / f"photo_{message.from_user.id}.jpg"
    await message.bot.download_file(f.file_path, destination=path)
    await state.update_data(photo=str(path), relatives=[])
    await state.set_state(Form.relative)
    await message.answer(
        "Qarindosh ma’lumotlarini quyidagi ko‘rinishda yuboring:\n"
        "Ota | F.I.Sh. | 1970-yil, joyi | Ish joyi va lavozimi | Turar joyi\n\n"
        "Barcha qarindoshlarni kiritib bo‘lgach: TAYYORLASH deb yozing."
    )

@dp.message(Form.photo)
async def photo_wrong(message: Message, state: FSMContext):
    await message.answer("Iltimos, rasmni foto sifatida yuboring.")

@dp.message(Form.relative)
async def relative(message: Message, state: FSMContext):
    text = message.text.strip()
    data = await state.get_data()

    if text.upper() == "TAYYORLASH":
        filename = OUT / f"Malumotnoma_{message.from_user.id}.docx"
        make_doc(data, filename)
        await message.answer_document(
            FSInputFile(filename),
            caption="Ma’lumotnomangiz tayyor."
        )
        await state.clear()
        return

    parts = [x.strip() for x in text.split("|")]
    if len(parts) != 5:
        await message.answer(
            "Format xato. 5 qism bo‘lishi kerak:\n"
            "Ota | F.I.Sh. | Tug‘ilgan yili va joyi | Ish joyi va lavozimi | Turar joyi"
        )
        return

    relatives = data.get("relatives", [])
    relatives.append(parts)
    await state.update_data(relatives=relatives)
    await message.answer("Qabul qilindi. Keyingi qarindoshni kiriting yoki TAYYORLASH deb yozing.")

async def health(request):
    return web.Response(text="OK")

async def run_bot():
    bot = Bot(BOT_TOKEN)
    await dp.start_polling(bot)

async def main():
    app = web.Application()
    app.router.add_get("/", health)
    app.router.add_get("/health", health)

    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", "10000"))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

    await run_bot()

if __name__ == "__main__":
    asyncio.run(main())
