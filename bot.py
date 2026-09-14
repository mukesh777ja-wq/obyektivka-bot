import os
from pathlib import Path
from datetime import datetime

from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

BOT_TOKEN = os.getenv("BOT_TOKEN", "BU_YERGA_BOT_TOKENINGIZNI_QOYING")
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
    relative_count = State()
    relative = State()

def no_kb():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="YO'Q")]],
        resize_keyboard=True,
        one_time_keyboard=True
    )

def ask_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Yana qarindosh qo'shish")],
            [KeyboardButton(text="Tayyorlash")]
        ],
        resize_keyboard=True
    )

def set_cell_text(cell, text, bold=False, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(str(text))
    r.bold = bold
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def set_doc_defaults(doc):
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2)
    sec.right_margin = Cm(1.5)
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

def add_field(p, label, value):
    r = p.add_run(label)
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    r2 = p.add_run(" " + value)
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(10)

def make_doc(data, filename):
    doc = Document()
    set_doc_defaults(doc)

    # 1-sahifa
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

    # Asosiy ma'lumotlar
    items = [
        ("Tug‘ilgan yili:", data["birth"]),
        ("Millati:", data["nationality"]),
        ("Ma’lumoti:", data["education"]),
        ("Ma’lumoti bo‘yicha mutaxassisligi:", data["specialty"]),
        ("Ilmiy darajasi:", data["degree"]),
        ("Qaysi chet tillarini biladi:", data["languages"]),
        ("Davlat mukofotlari bilan taqdirlanganligi (qanaqa?):", data["awards"]),
        ("Xalq deputatlari respublika, viloyat, shahar va tuman Kengashi deputatimi yoki boshqa saylanadigan organlarning a’zosimi (to‘liq ko‘rsatilishi lozim):", data["elected"]),
    ]
    for label, value in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        add_field(p, label, value)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MEHNAT FAOLIYATI")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)

    for row in data.get("work", []):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        add_field(p, "", row)

    # 2-sahifa
    doc.add_page_break()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f'{data["fio"]}ning yaqin qarindoshlari haqida')
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r.bold = True
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MA’LUMOT")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)

    headers = ["Qarindoshligi", "Familiyasi, ismi va otasining ismi",
               "Tug‘ilgan yili va joyi", "Ish joyi va lavozimi", "Turar joyi"]
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    widths = [2.1, 4.1, 3.2, 4.0, 5.0]
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, True, 8)

    for rel in data.get("relatives", []):
        cells = table.add_row().cells
        vals = [rel["type"], rel["fio"], rel["birth_place"], rel["job"], rel["address"]]
        for i, v in enumerate(vals):
            set_cell_text(cells[i], v, False, 8)

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
    await message.answer("Partiyaviyligingizni kiriting yoki YO‘Q tugmasini bosing:", reply_markup=no_kb())

@dp.message(Form.party)
async def party(message: Message, state: FSMContext):
    await state.update_data(party=message.text.strip())
    await state.set_state(Form.education)
    await message.answer("Ma’lumotingizni kiriting:", reply_markup=no_kb())

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
    await message.answer("Ilmiy darajangizni kiriting yoki YO‘Q bosing:", reply_markup=no_kb())

@dp.message(Form.degree)
async def degree(message: Message, state: FSMContext):
    await state.update_data(degree=message.text.strip())
    await state.set_state(Form.languages)
    await message.answer("Qaysi chet tillarini bilasiz? yoki YO‘Q bosing:", reply_markup=no_kb())

@dp.message(Form.languages)
async def languages(message: Message, state: FSMContext):
    await state.update_data(languages=message.text.strip())
    await state.set_state(Form.awards)
    await message.answer("Davlat mukofotlari bilan taqdirlanganmisiz? Qaysi? yoki YO‘Q:", reply_markup=no_kb())

@dp.message(Form.awards)
async def awards(message: Message, state: FSMContext):
    await state.update_data(awards=message.text.strip())
    await state.set_state(Form.elected)
    await message.answer("Saylanadigan organlarda deputat yoki a’zo bo‘lganmisiz? yoki YO‘Q:", reply_markup=no_kb())

@dp.message(Form.elected)
async def elected(message: Message, state: FSMContext):
    await state.update_data(elected=message.text.strip())
    await state.set_state(Form.work)
    await state.update_data(work=[])
    await message.answer(
        "Mehnat faoliyatingizni kiriting.\n"
        "Masalan: 2018-2020 yy. — Tashkilot nomi — lavozim.\n"
        "Har bir ish joyini alohida xabar qilib yuboring.\n"
        "Tugatish uchun: YO‘Q"
    )

@dp.message(Form.work)
async def work(message: Message, state: FSMContext):
    text = message.text.strip()
    data = await state.get_data()
    work = data.get("work", [])
    if text.upper() == "YO‘Q":
        await state.set_state(Form.photo)
        await message.answer("Endi 3x4 rasmingizni yuboring (foto sifatida):")
        return
    work.append(text)
    await state.update_data(work=work)
    await message.answer("Keyingi ish joyini kiriting yoki YO‘Q bosing.")

@dp.message(Form.photo, F.photo)
async def photo(message: Message, state: FSMContext):
    file = await message.bot.get_file(message.photo[-1].file_id)
    path = OUT / f"photo_{message.from_user.id}.jpg"
    await message.bot.download_file(file.file_path, destination=path)
    await state.update_data(photo=str(path), relatives=[])
    await state.set_state(Form.relative)
    await message.answer(
        "Qarindosh ma’lumotlarini kiritishni boshlaymiz.\n\n"
        "Masalan, birinchi bo‘lib OTA ma’lumotlarini yuboring.\n"
        "Format:\nQarindoshligi | F.I.Sh. | Tug‘ilgan yili va joyi | Ish joyi/lavozimi | Turar joyi"
    )

@dp.message(Form.photo)
async def photo_wrong(message: Message, state: FSMContext):
    await message.answer("Iltimos, 3x4 rasmingizni foto sifatida yuboring.")

@dp.message(Form.relative)
async def relative(message: Message, state: FSMContext):
    text = message.text.strip()
    data = await state.get_data()
    relatives = data.get("relatives", [])
    if text.lower() == "tayyorlash":
        filename = OUT / f"Malumotnoma_{message.from_user.id}.docx"
        make_doc(data, filename)
        await message.answer_document(
            document=__import__("aiogram").types.FSInputFile(filename),
            caption="Ma’lumotnomangiz tayyor."
        )
        await state.clear()
        return

    parts = [x.strip() for x in text.split("|")]
    if len(parts) != 5:
        await message.answer(
            "Iltimos, aynan 5 ta qism bilan yuboring:\n"
            "Ota | F.I.Sh. | 1970-yil, joyi | Ish joyi/lavozimi | Turar joyi"
        )
        return
    relatives.append({
        "type": parts[0], "fio": parts[1], "birth_place": parts[2],
        "job": parts[3], "address": parts[4]
    })
    await state.update_data(relatives=relatives)
    await message.answer(
        "Qabul qilindi. Keyingi qarindoshni shu formatda yuboring.\n"
        "Tugatgan bo‘lsangiz: Tayyorlash",
        reply_markup=ask_kb()
    )

async def main():
    if BOT_TOKEN == "BU_YERGA_BOT_TOKENINGIZNI_QOYING":
        raise RuntimeError("BOT_TOKEN ni muhit o‘zgaruvchisiga kiriting.")
    bot = Bot(BOT_TOKEN)
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
