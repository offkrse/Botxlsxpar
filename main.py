import os
import logging
import pandas as pd
import io
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from telegram.error import TelegramBadRequest

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Токен бота (установите в переменных окружения Render)
BOT_TOKEN = os.environ.get('BOT_TOKEN')

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /start"""
    welcome_text = """
🤖 Добро пожаловать в бот-парсер Excel файлов!

📊 **Как использовать:**
1. Отправьте мне один или несколько Excel файлов (.xlsx, .xls)
2. Я извлеку все данные из первого столбца (начиная со второй строки)
3. Верну вам текстовый файл с результатом

⚠️ **Важно:** 
- Файлы должны быть в формате Excel
- Данные берутся из первого столбца, начиная со второй строки
- Поддерживается обработка нескольких файлов одновременно
"""
    await update.message.reply_text(welcome_text)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /help"""
    help_text = """
📖 **Помощь по использованию бота:**

Просто отправьте мне Excel файлы, и я:
1. Извлеку все значения из первого столбца
2. Пропущу первую строку (заголовок)
3. Объединю данные из всех файлов
4. Верну текстовый файл с результатом

📁 **Поддерживаемые форматы:** .xlsx, .xls

💡 **Совет:** Вы можете отправить несколько файлов сразу!
"""
    await update.message.reply_text(help_text)

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик документов"""
    try:
        document = update.message.document
        
        # Проверяем, что это Excel файл
        file_extension = document.file_name.split('.')[-1].lower()
        if file_extension not in ['xlsx', 'xls']:
            await update.message.reply_text(
                "❌ Пожалуйста, отправьте файл в формате Excel (.xlsx или .xls)"
            )
            return
        
        # Отправляем сообщение о начале обработки
        processing_msg = await update.message.reply_text(
            f"⏳ Обрабатываю файл: {document.file_name}..."
        )
        
        # Скачиваем файл
        file = await context.bot.get_file(document.file_id)
        file_bytes = await file.download_as_bytearray()
        
        # Обрабатываем файл
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None)
            # Берём первый столбец, начиная со второй строки
            column_data = df.iloc[1:, 0].dropna().astype(str).tolist()
            
            if not column_data:
                await update.message.reply_text(
                    "❌ В файле не найдено данных для обработки"
                )
                return
            
            # Создаем текстовый файл в памяти
            text_content = "\n".join(column_data)
            text_file = io.BytesIO(text_content.encode('utf-8'))
            text_file.name = f"parsed_data_{document.file_name.split('.')[0]}.txt"
            
            # Отправляем результат
            await update.message.reply_document(
                document=text_file,
                caption=f"✅ Обработан файл: {document.file_name}\n"
                       f"📊 Извлечено записей: {len(column_data)}"
            )
            
            # Удаляем сообщение о обработке
            await processing_msg.delete()
            
        except Exception as e:
            logger.error(f"Error processing file: {e}")
            await update.message.reply_text(
                f"❌ Ошибка при обработке файла:\n{str(e)}"
            )
            await processing_msg.delete()
            
    except Exception as e:
        logger.error(f"Error in handle_document: {e}")
        await update.message.reply_text(
            "❌ Произошла ошибка при обработке файла. Попробуйте еще раз."
        )

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик ошибок"""
    logger.error(f"Update {update} caused error {context.error}")
    
    try:
        await update.message.reply_text(
            "❌ Произошла непредвиденная ошибка. Попробуйте еще раз."
        )
    except:
        pass

def main():
    """Основная функция для запуска бота"""
    if not BOT_TOKEN:
        logger.error("BOT_TOKEN not found in environment variables!")
        return
    
    # Создаем приложение
    application = Application.builder().token(BOT_TOKEN).build()
    
    # Добавляем обработчики
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    application.add_error_handler(error_handler)
    
    # Запускаем бота
    logger.info("Bot is starting...")
    application.run_polling()

if __name__ == '__main__':
    main()
