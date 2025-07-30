
import json
import datetime
import pytz
from aiogram import Bot, Dispatcher, types
from apscheduler.schedulers.asyncio import AsyncIOScheduler
import asyncio
import os
from aiohttp import web
import pandas as pd

TOKEN = "8002116184:AAHtPEZX08V33fzxxH8BRVvnhJWaCE42g0Y"

CONFIG_FILE = "config.json"
EMPLOYEES_FILE = "employees.json"

bot = Bot(token=TOKEN)
dp = Dispatcher(bot)
scheduler = AsyncIOScheduler()
tz = pytz.timezone("Europe/Moscow")

if os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
        CHAT_ID = data.get("CHAT_ID")
else:
    CHAT_ID = None

def load_employees():
    if not os.path.exists(EMPLOYEES_FILE):
        return []
    with open(EMPLOYEES_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_employees(data):
    with open(EMPLOYEES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def save_chat_id(chat_id):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"CHAT_ID": chat_id}, f, ensure_ascii=False, indent=2)

async def check_birthdays():
    if not CHAT_ID:
        return
    today = datetime.datetime.now(tz).strftime("%d-%m")
    employees = load_employees()
    for emp in employees:
        if emp["birthday"] == today:
            msg = f"🎉 Сегодня день рождения у {emp['name']}! Поздравляем! 🎂"
            await bot.send_message(int(CHAT_ID), msg)

scheduler.add_job(check_birthdays, "cron", hour=8, minute=30)

@dp.message_handler(commands=["start"])
async def start_cmd(message: types.Message):
    await message.answer("Бот для поздравлений запущен!")

@dp.message_handler(commands=["get_chat_id"])
async def get_chat_id(message: types.Message):
    global CHAT_ID
    CHAT_ID = message.chat.id
    save_chat_id(CHAT_ID)
    await message.answer(f"✅ CHAT_ID этой группы сохранён: {CHAT_ID}")

# upload_excel + файл в одном сообщении
@dp.message_handler(commands=["upload_excel"], content_types=types.ContentType.DOCUMENT)
async def upload_excel_with_command(message: types.Message):
    if not message.document.file_name.endswith(".xlsx"):
        await message.answer("❌ Пришли Excel файл (.xlsx)")
        return
    file = await message.document.download()
    await process_excel(file.name, message)

# файл с подписью /upload_excel
@dp.message_handler(content_types=types.ContentType.DOCUMENT)
async def upload_excel_caption(message: types.Message):
    if message.caption and "/upload_excel" in message.caption:
        if not message.document.file_name.endswith(".xlsx"):
            await message.answer("❌ Пришли Excel файл (.xlsx)")
            return
        file = await message.document.download()
        await process_excel(file.name, message)

async def process_excel(file_path, message):
    try:
        df = pd.read_excel(file_path)
        employees = []
        for _, row in df.iterrows():
            name = str(row["name"]).strip()
            birthday = str(row["birthday"]).strip()
            employees.append({"name": name, "birthday": birthday})
        save_employees(employees)
        await message.answer(f"✅ Загружено сотрудников: {len(employees)}")
    except Exception as e:
        await message.answer(f"❌ Ошибка при загрузке Excel: {e}")

@dp.message_handler(commands=["list_employees"])
async def list_employees(message: types.Message):
    employees = load_employees()
    if not employees:
        await message.answer("📂 Список сотрудников пуст.")
        return
    text = "👥 Список сотрудников:\n"
    for emp in employees:
        text += f"- {emp['name']} — {emp['birthday']}\n"
    await message.answer(text)

@dp.message_handler(commands=["test_birthday"])
async def test_birthday(message: types.Message):
    today = datetime.datetime.now(tz).strftime("%d-%m")
    employees = load_employees()
    found = False
    for emp in employees:
        if emp["birthday"] == today:
            msg = f"🎉 ТЕСТ: Сегодня день рождения у {emp['name']}! Поздравляем! 🎂"
            await bot.send_message(message.chat.id, msg)
            found = True
    if not found:
        await message.answer("Сегодня нет дней рождения.")


@dp.message_handler(commands=["debug_time"])
async def debug_time(message: types.Message):
    current_time = datetime.datetime.now(tz)
    today = current_time.strftime("%d-%m")
    employees = load_employees()
    matches = [emp["name"] for emp in employees if emp["birthday"] == today]

    text = f"🕒 Текущее время (МСК): {current_time.strftime('%Y-%m-%d %H:%M:%S')}\n"
    text += f"📅 Сегодня бот проверяет дату: {today}\n"
    if matches:
        text += "🎉 Совпадения: " + ", ".join(matches)
    else:
        text += "Сегодня нет совпадений."

    await message.answer(text)


# web server
async def handle(request):
    return web.Response(text="Bot is running!")

async def start_webserver():
    app = web.Application()
    app.router.add_get("/", handle)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", int(os.environ.get("PORT", 10000)))
    await site.start()

async def main():
    scheduler.start()
    await start_webserver()
    await dp.start_polling()

if __name__ == "__main__":
    asyncio.run(main())
