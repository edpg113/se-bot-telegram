from telegram.ext import Application, CommandHandler
from dotenv import load_dotenv
import os
import json
import requests
import openpyxl

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

async def get_spek_toko(update, context):
    if context.args:
        kodeToko = context.args[0].lower()
        with open('spek-toko.json', 'r') as file:
            data = json.load(file)
            spek_list = data.get("spesifikasi", []) # Ambil array karyawan dari data
            for spesifikasi in spek_list:
                if kodeToko in spesifikasi["kode"].lower():
            #     data_karyawan = data[karyawan]
                    await update.message.reply_text(f"Berikut Data Spesifikasi :\n\n"
                                            f"Kode Toko : {spesifikasi['kode']}\n"
                                            f"Nama Toko : {spesifikasi['namaToko']}\n"
                                            f"Alamat : {spesifikasi['alamat']}\n"
                                            f"Jumlah Station : {spesifikasi['station']}\n"
                                            f"Ikios : {spesifikasi['ikios']}\n\n"
                                            f"Station 1 : {spesifikasi.get('station1', '')}\n"
                                            f"Station 2 : {spesifikasi.get('station2' , '')}\n"
                                            f"Station 3 : {spesifikasi.get('station3', '')}\n"
                                            f"Station 4 : {spesifikasi.get('station4', '')}\n"
                                            f"Station 5 : {spesifikasi.get('station5', '')}\n"
                                            f"Station 6 : {spesifikasi.get('station6', '')}\n"
                                            f"Station 7 : {spesifikasi.get('station7', '')}\n"
                                            f"Station 8 : {spesifikasi.get('station8', '')}\n"
                                            f"Station 9 : {spesifikasi.get('station9', '')}\n"
                                            f"Station 10 : {spesifikasi.get('station10', '')}\n\n"
)
                    return
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
    # application.add_handler(CommandHandler("spek", get_spek))
    application.add_handler(CommandHandler("spek", get_spek_toko))

    print("Bot Telegram berjalan...")
    application.run_polling()
