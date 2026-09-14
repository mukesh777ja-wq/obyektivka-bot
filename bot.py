import os
import asyncio
from pathlib import Path

from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    Message, ReplyKeyboardMarkup, KeyboardButton, FSInputFile, Update
)
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
    from docx.enum.table import WD_ROW_HEIGHT_RULE

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(1.7)
    sec.right_margin = Cm(1.5)

    def style_run(run, size=10, bold=False):
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)
        run.bold = bold

    def style_paragraph(p, align=None, before=0, after=0, line=3):
        if align is not None:
            p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.line_spacing = line

    # ===== 1-BET =====
    p = doc.add_paragraph()
    style_paragraph(p, WD_ALIGN_PARAGRAPH.CENTER, after=2, line=3)
    style_run(p.add_run("MA’LUMOTNOMA"), 14, True)

    p = doc.add_paragraph()
    style_paragraph(p, WD_ALIGN_PARAGRAPH.CENTER, after=4, line=3)
    style_run(p.add_run(data["fio"]), 11, True)

    # Two-column block, matching the reference image.
    info = doc.add_table(rows=1, cols=2)
    info.autofit = False
    info.columns[0].width = Cm(8.0)
    info.columns[1].width = Cm(7.0)

    left = [
        ("Tug‘ilgan yili:", data["birth"]),
        ("Millati:", data["nationality"]),
        ("Ma’lumoti:", data["education"]),
        ("Ma’lumoti bo‘yicha mutaxassisligi:", data["specialty"]),
        ("Ilmiy darajasi:", data["degree"]),
        ("Qaysi chet tillarini biladi:", data["languages"]),
        ("Davlat mukofotlari bilan taqdirlanganligi (qanaqa):", data["awards"]),
        ("Xalq deputatlari respublika, viloyat, shahar va tuman Kengashlari deputatligi yoki boshqa saylanadigan organlarida a’zoligi (to‘liq ko‘rsatilishi lozim):", data["elected"]),
    ]
    right = [
        ("Tug‘ilgan joyi:", data["birth_place"]),
        ("Partiyaviyligi:", data["party"]),
        ("Tamomlagan:", data["graduated"]),
    ]

    for cell, items in zip(info.rows[0].cells, (left, right)):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        cell.text = ""
        for label, value in items:
            p = cell.add_paragraph() if cell.paragraphs[0].text else cell.paragraphs[0]
            style_paragraph(p, after=1, line=3)
            style_run(p.add_run(label + " "), 9, True)
            style_run(p.add_run(value), 9, False)

    # Remove visible borders from the info table.
    tblPr = info._tbl.tblPr
    borders = tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        from docx.oxml import OxmlElement
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        el = borders.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            borders.append(el)
        el.set(qn("w:val"), "nil")

    # Put the uploaded 3x4 photo at the top-right of the right column.
    photo = data.get("photo")
    if photo and Path(photo).exists():
        # 3x4 photo with a clear black border.
        cell = info.rows[0].cells[1]
        while len(cell.paragraphs) > 0:
            p = cell.paragraphs[0]
            p._element.getparent().remove(p._element)

        photo_tbl = cell.add_table(rows=1, cols=1)
        photo_tbl.autofit = False
        photo_cell = photo_tbl.cell(0, 0)
        photo_cell.width = Cm(3.2)
        photo_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        p = photo_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run()
        run.add_picture(photo, width=Cm(3.0), height=Cm(4.0))

        # Black border around the photo cell.
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        tcPr = photo_cell._tc.get_or_add_tcPr()
        borders = tcPr.first_child_found_in("w:tcBorders")
        if borders is None:
            borders = OxmlElement("w:tcBorders")
            tcPr.append(borders)
        for edge in ("top", "left", "bottom", "right"):
            tag = "w:" + edge
            el = borders.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                borders.append(el)
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "12")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")

        # Fixed labels are bold; values entered by the user remain regular.
        for label, value in right:
            p = cell.add_paragraph()
            style_paragraph(p, after=1, line=3)
            style_run(p.add_run(label + " "), 9, True)
            style_run(p.add_run(value), 9, False)

    p = doc.add_paragraph()
    style_paragraph(p, WD_ALIGN_PARAGRAPH.CENTER, before=3, after=2, line=3)
    style_run(p.add_run("MEHNAT FAOLIYATI"), 12, True)

    for work in data.get("work", []):
        p = doc.add_paragraph()
        style_paragraph(p, after=1, line=3)
        style_run(p.add_run(work), 9, False)

    # ===== 2-BET =====
    doc.add_page_break()

    p = doc.add_paragraph()
    style_paragraph(p, WD_ALIGN_PARAGRAPH.CENTER, after=1, line=3)
    style_run(p.add_run(f'{data["fio"]}ning yaqin qarindoshlari haqida'), 10, True)

    p = doc.add_paragraph()
    style_paragraph(p, WD_ALIGN_PARAGRAPH.CENTER, after=2, line=3)
    style_run(p.add_run("MA’LUMOT"), 12, True)

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    widths = [Cm(2.4), Cm(4.3), Cm(4.0), Cm(4.1), Cm(4.0)]
    headers = [
        "Qarindoshligi",
        "Familiyasi, ismi va otasining ismi",
        "Tug‘ilgan yili va joyi",
        "Ish joyi va lavozimi",
        "Turar joyi",
    ]

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.width = widths[i]
        set_cell(cell, h, True)

    # Header row also kept at 2 cm.
    table.rows[0].height = Cm(2)
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    for rel in data.get("relatives", []):
        row = table.add_row()
        row.height = Cm(2)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for i, v in enumerate(rel):
            cell = row.cells[i]
            cell.width = widths[i]
            set_cell(cell, v)

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
    if text.upper() in ("YO'Q", "YO‘Q"):
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

    if text.upper() == "TAYYORLASH":
        data = await state.get_data()
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

    data = await state.get_data()
    relatives = data.get("relatives", [])
    relatives.append(parts)
    await state.update_data(relatives=relatives)
    await message.answer(
        "Qabul qilindi. Keyingi qarindoshni kiriting yoki TAYYORLASH deb yozing."
    )

async def health(request):
    return web.Response(text="OK")

async def telegram_webhook(request):
    try:
        data = await request.json()
        update = Update.model_validate(data)
        await dp.feed_update(request.app["bot"], update)
        return web.Response(text="OK")
    except Exception as e:
        print(f"Webhook error: {e}")
        return web.Response(status=500, text="Webhook error")

async def on_startup(app):
    bot = app["bot"]

    external_url = os.environ.get("RENDER_EXTERNAL_URL")
    if not external_url:
        external_url = os.environ.get("EXTERNAL_URL")

    if not external_url:
        raise RuntimeError(
            "RENDER_EXTERNAL_URL topilmadi. Render Web Service URL manzili kerak."
        )

    webhook_url = external_url.rstrip("/") + "/telegram/webhook"
    await bot.set_webhook(webhook_url, drop_pending_updates=True)
    print(f"Telegram webhook set: {webhook_url}")

async def on_cleanup(app):
    bot = app["bot"]
    try:
        await bot.delete_webhook(drop_pending_updates=False)
    finally:
        await bot.session.close()

async def main():
    bot = Bot(BOT_TOKEN)

    app = web.Application()
    app["bot"] = bot
    app.router.add_get("/", health)
    app.router.add_get("/health", health)
    app.router.add_post("/telegram/webhook", telegram_webhook)

    app.on_startup.append(on_startup)
    app.on_cleanup.append(on_cleanup)

    port = int(os.environ.get("PORT", "10000"))
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

    print(f"Web service listening on port {port}")

    try:
        await asyncio.Event().wait()
    finally:
        await runner.cleanup()

if __name__ == "__main__":
    asyncio.run(main())
