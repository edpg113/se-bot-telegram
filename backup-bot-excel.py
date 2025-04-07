from telegram.ext import Application, CommandHandler
from dotenv import load_dotenv
import os
import json
import requests
import pandas as pd

load_dotenv()


async def start(update, context):
    await update.message.reply_text("Hallo! selamat datang di bot EDP!")

async def help_command(update, context):
    await update.message.reply_text("Ini adalah bot EDP. Gunakan /start untuk memulai.")

async def get_data(update, context):
    url = "http://api.open-notify.org/astros"
    response = requests.get(url)
    data = response.json()
    await update.message.reply_text(f"Data: {data}")

async def get_spek(update, context):
    if context.args:
        station = context.args[0].lower()
        with open('spek.json', 'r') as file:
            data = json.load(file)
        if station in data:
            spek = data[station]
            await update.message.reply_text(f"Spesifikasi {station}:\n" 
                                            f"CPU: {spek['CPU']}\n"
                                            f"RAM: {spek['RAM']}\n"
                                            f"Storage: {spek['Storage']}\n")
        else:
            await update.message.reply_text("Data tidak ditemukan.")
    else:
        await update.message.reply_text("Silakan masukkan nama stasiun setelah perintah /spek. (contoh: /spek station1)")

# Fungsi untuk mendapatkan spesifikasi dari file excel
def get_spek_excel():
    df = pd.read_excel('spek.xlsx')
    df = df.astype(str)
    # Membaca setiap baris di dalam file excel, mulai dari beris kedua (skip header)

    data = {row["Station"].strip().lower(): 
            {"CPU": row["CPU"], "RAM": row["RAM"], "Storage": row["Storage"]}
            for _, row in df.iterrows()}

    print("Data dari Excel: ", data)
    return data

async def get_spek_toko(update, context):
    spek = get_spek_excel()

    if context.args:
        station = context.args[0].lower()
        if station in spek:
                    speks = spek[station]
                    await update.message.reply_text(f"Berikut Data Spesifikasi :\n\n"
                                           f"Spesifikasi {station}:\n"
                       f"CPU: {speks['CPU']}\n"
                       f"RAM: {speks['RAM']}\n"
                       f"Storage: {speks['Storage']}"
)
        else:
            await update.message.reply_text("Data tidak ditemukan.")
    else:
        await update.message.reply_text("Silakan masukkan kode toko setelah perintah /spek. (contoh: /spek TCXH)")

if __name__ == '__main__':
    TOKEN_TELE = os.getenv("TOKEN")

    application = Application.builder().token(TOKEN_TELE).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("data", get_data))
    application.add_handler(CommandHandler("spek", get_spek_toko))

    print("Bot Telegram berjalan...")
    application.run_polling()
