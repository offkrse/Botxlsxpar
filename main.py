import os
import logging
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import pandas as pd
from io import BytesIO

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Токен бота (замените на ваш токен)
BOT_TOKEN = os.getenv("BOT_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /start"""
    await update.message.reply_text(
        "Привет! Отправь мне один или несколько xlsx файлов.\n"
        "Я извлеку данные из первого столбца (пропуская первую строку) "
        "и создам txt файл с результатами."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик получения документов"""
    try:
        document = update.message.document
        
        # Проверяем, что файл имеет расширение .xlsx
        if not document.file_name.lower().endswith('.xlsx'):
            await update.message.reply_text("Пожалуйста, отправьте файл с расширением .xlsx")
            return
        
        # Скачиваем файл
        file = await context.bot.get_file(document.file_id)
        file_bytes = await file.download_as_bytearray()
        
        # Сохраняем файл во временный список
        if 'xlsx_files' not in context.user_data:
            context.user_data['xlsx_files'] = []
        
        context.user_data['xlsx_files'].append({
            'name': document.file_name,
            'data': file_bytes
        })
        
        await update.message.reply_text(f"Файл '{document.file_name}' получен! Отправьте следующий файл или введите /process для обработки.")
        
    except Exception as e:
        logger.error(f"Ошибка при обработке файла: {e}")
        await update.message.reply_text("Произошла ошибка при обработке файла. Попробуйте еще раз.")

async def process_files(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /process - обработка всех полученных файлов"""
    try:
        if 'xlsx_files' not in context.user_data or not context.user_data['xlsx_files']:
            await update.message.reply_text("Сначала отправьте xlsx файлы!")
            return
        
        all_data = []
        
        # Обрабатываем каждый файл
        for file_info in context.user_data['xlsx_files']:
            try:
                # Читаем xlsx файл
                df = pd.read_excel(BytesIO(file_info['data']))
                
                # Извлекаем данные из первого столбца, пропуская первую строку
                if len(df.columns) > 0:
                    column_data = df.iloc[1:, 0]  # Первый столбец, начиная со второй строки
                    # Фильтруем пустые значения и преобразуем в строки
                    valid_data = [str(item).strip() for item in column_data if pd.notna(item) and str(item).strip()]
                    all_data.extend(valid_data)
                    
                    logger.info(f"Из файла {file_info['name']} извлечено {len(valid_data)} записей")
                
            except Exception as e:
                logger.error(f"Ошибка при обработке файла {file_info['name']}: {e}")
                await update.message.reply_text(f"Ошибка при обработке файла {file_info['name']}")
                continue
        
        if not all_data:
            await update.message.reply_text("Не удалось извлечь данные из файлов.")
            return
        
        # Создаем txt файл
        txt_content = '\n'.join(all_data)
        txt_filename = "extracted_data.txt"
        
        # Сохраняем временный файл
        with open(txt_filename, 'w', encoding='utf-8') as f:
            f.write(txt_content)
        
        # Отправляем файл пользователю
        with open(txt_filename, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename=txt_filename,
                caption=f"Готово! Извлечено {len(all_data)} записей из {len(context.user_data['xlsx_files'])} файлов."
            )
        
        # Очищаем временные файлы и данные
        os.remove(txt_filename)
        context.user_data['xlsx_files'] = []
        
    except Exception as e:
        logger.error(f"Ошибка при обработке файлов: {e}")
        await update.message.reply_text("Произошла ошибка при обработке файлов.")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /help"""
    await update.message.reply_text(
        "Доступные команды:\n"
        "/start - начать работу с ботом\n"
        "/process - обработать все полученные xlsx файлы\n"
        "/help - показать эту справку\n\n"
        "Просто отправьте xlsx файлы боту, затем введите /process для получения результата."
    )

def main():
    """Основная функция"""
    # Создаем приложение
    application = Application.builder().token(BOT_TOKEN).build()
    
    # Добавляем обработчики
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("process", process_files))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    # Запускаем бота
    print("Бот запущен...")
    application.run_polling()

if __name__ == '__main__':
    main()
