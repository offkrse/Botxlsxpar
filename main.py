import os
import logging
import pandas as pd
import io
from aiogram import Bot, Dispatcher, Router, F
from aiogram.types import Message, BufferedInputFile
from aiogram.filters import Command
from aiogram.enums import ParseMode

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
router = Router()
dp.include_router(router)

@router.message(Command("start"))
async def cmd_start(message: Message):
    text = """
🤖 <b>Парсер Excel в TXT</b>

Отправьте Excel файлы (.xlsx, .xls) - я извлеку данные из первого столбца и верну текстовый файл.

<b>Особенности:</b>
• Беру данные из первого столбца (со 2-й строки)
• Поддерживаю множественную загрузку
• Автоматически объединяю данные
"""
    await message.answer(text, parse_mode=ParseMode.HTML)

class ExcelProcessor:
    def __init__(self):
        self.all_data = []
    
    def process_file(self, file_bytes: bytes) -> int:
        """Обрабатывает один файл, возвращает количество записей"""
        df = pd.read_excel(io.BytesIO(file_bytes), header=None)
        column_data = df.iloc[1:, 0].dropna().astype(str).tolist()
        self.all_data.extend(column_data)
        return len(column_data)
    
    def get_result_text(self) -> str:
        """Возвращает объединенный текст"""
        return "\n".join(self.all_data)

processor = ExcelProcessor()

@router.message(F.document)
async def handle_excel_files(message: Message):
    try:
        document = message.document
        file_name = document.file_name or "file"
        
        if not file_name.lower().endswith(('.xlsx', '.xls')):
            await message.answer("❌ Отправьте Excel файл (.xlsx или .xls)")
            return
        
        processing_msg = await message.answer(f"⏳ Обрабатываю {file_name}...")
        
        # Скачиваем и обрабатываем файл
        file = await bot.download(document)
        file_bytes = await file.read()
        
        records_count = processor.process_file(file_bytes)
        
        # Создаем результат для этого файла
        result_text = processor.get_result_text()
        result_file = BufferedInputFile(
            result_text.encode('utf-8'),
            filename=f"result_{file_name.split('.')[0]}.txt"
        )
        
        await message.answer_document(
            document=result_file,
            caption=f"✅ {file_name}\n📊 Записей: {records_count}\n📁 Всего: {len(processor.all_data)}"
        )
        
        await processing_msg.delete()
        
    except Exception as e:
        logger.error(f"Error: {e}")
        await message.answer(f"❌ Ошибка: {str(e)}")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
