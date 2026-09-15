import os
import asyncio
from pathlib import Path
from copy import deepcopy

from aiohttp import web
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, FSInputFile, Update

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BOT_TOKEN = os.environ["BOT_TOKEN"]
BASE = Path(__file__).resolve().parent
TEMPLATE = BASE / "template.docx"
OUT = BASE / "output"
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

def set_run(run, bold):
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.bold = bold

def clear_paragraph(p):
    for child in list(p._p):
        if child.tag != qn("w:pPr"):
            p._p.remove(child)

def put_label_value(p, label, value):
    clear_paragraph(p)
    r = p.add_run(label)
    set_run(r, True)
    r = p.add_run(str(value))
    set_run(r, False)

def put_two_fields(p, label1, value1, label2, value2, spaces=45):
    clear_paragraph(p)
    r = p.add_run(label1)
    set_run(r, True)
    r = p.add_run(str(value1))
    set_run(r, False)
    r = p.add_run(" " * spaces)
    set_run(r, False)
    r = p.add_run(label2)
    set_run(r, True)
    r = p.add_run(str(value2))
    set_run(r, False)

def set_photo_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
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

def add_photo_to_cell(cell, photo_path):
    # Clear the template's empty paragraphs.
    for p in list(cell.paragraphs):
        clear_paragraph(p)
    # A nested 1x1 table gives a clean black frame exactly around the 3x4 photo.
    tbl = cell.add_table(rows=1, cols=1)
    tbl.autofit = False
    pc = tbl.cell(0, 0)
    pc.width = Cm(3.2)
    pc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_photo_border(pc)
    p = pc.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    run.add_picture(str(photo_path), width=Cm(3.0), height=Cm(4.0))

def copy_row_format(src_row, dst_row):
    # Copy cell properties from the first data row of the template.
    for s, d in zip(src_row.cells, dst_row.cells):
        d._tc.get_or_add_tcPr().clear_content()
        d._tc.get_or_add_tcPr().extend(deepcopy(s._tc.get_or_add_tcPr()))
        d.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def fill_table_cell(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(str(text))
    r.font.name = "Times New Roman"
    r.font.size = Pt(8)
    r.bold = bold

def make_doc(data, filename):
    doc = Document(str(TEMPLATE))

    # Page 1: use the uploaded Word document as the actual template.
    p = doc.paragraphs[1]
    put_label_value(p, "", data["fio"])
    for run in p.runs:
        set_run(run, True)

    info = doc.tables[0]
    left = info.cell(0, 0)
    paras = left.paragraphs

    put_label_value(paras[0], "Tug‘ilgan yili: ", data["birth"])
    put_label_value(paras[1], "Tug‘ilgan joyi: ", data["birth_place"])
    put_two_fields(paras[2], "Millati: ", data["nationality"], "Partiyaviyligi: ", data["party"], spaces=43)
    put_label_value(paras[3], "Ma’lumoti: ", data["education"])
    put_label_value(paras[4], "Tamomlagan: ", data["graduated"])
    put_label_value(paras[5], "Ma’lumoti bo‘yicha mutaxassisligi: ", data["specialty"])
    put_two_fields(paras[6], "Qaysi chet tillarini biladi: ", data["languages"], "Ilmiy darajasi: ", data["degree"], spaces=38)
    put_label_value(paras[7], "Davlat mukofotlari bilan taqdirlanganligi (qanaqa): ", data["awards"])
    put_label_value(paras[8], "Xalq deputatlari respublika, viloyat, shahar va tuman Kengashlari deputatligi yoki boshqa saylanadigan organlarida a’zoligi (to‘liq ko‘rsatilishi lozim): ", data["elected"])

    # Preserve the template's paragraph spacing/line spacing.
    for p in paras:
        p.paragraph_format.line_spacing = 2.5

    photo_cell = info.cell(0, 1)
    add_photo_to_cell(photo_cell, data["photo"])

    # Page 1 headings already exist in the template.
    work_p = doc.paragraphs[3]
    work_p.text = ""
    for idx, work in enumerate(data.get("work", [])):
        if idx == 0:
            p = work_p
        else:
            p = doc.add_paragraph()
        r = p.add_run(work)
        set_run(r, False)
        p.paragraph_format.line_spacing = 2.5

    # Page 2 title uses the template paragraph, only the FIO changes.
    p = doc.paragraphs[5]
    put_label_value(p, "", f'{data["fio"]}ning yaqin qarindoshlari haqida')
    for run in p.runs:
        set_run(run, True)

    table = doc.tables[1]
    template_data_row = table.rows[1]

    # Remove all existing data rows, keep the header.
    while len(table.rows) > 1:
        tr = table.rows[-1]._tr
        tr.getparent().remove(tr)

    widths = [Cm(2.8), Cm(3.6), Cm(3.4), Cm(3.5), Cm(3.4)]

    for rel in data.get("relatives", []):
        row = table.add_row()
        # Copy row/cell properties from the template's original first data row.
        for s, d in zip(template_data_row.cells, row.cells):
            d._tc.get_or_add_tcPr().clear_content()
            d._tc.get_or_add_tcPr().extend(deepcopy(s._tc.get_or_add_tcPr()))
        row.height = Cm(2)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for i, v in enumerate(rel):
            row.cells[i].width = widths[i]
            fill_table_cell(row.cells[i], v, False)

    # Header stays bold as in the template.
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            for r in p.runs:
                set_run(r, True)

    # Exact 2 cm row height for every table row.
    for row in table.rows:
        row.height = Cm(2)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    doc.save(str(filename))

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
        await message.answer_document(FSInputFile(filename), caption="Ma’lumotnomangiz tayyor.")
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
    await message.answer("Qabul qilindi. Keyingi qarindoshni kiriting yoki TAYYORLASH deb yozing.")

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
    external_url = os.environ.get("RENDER_EXTERNAL_URL") or os.environ.get("EXTERNAL_URL")
    if not external_url:
        raise RuntimeError("RENDER_EXTERNAL_URL topilmadi.")
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
