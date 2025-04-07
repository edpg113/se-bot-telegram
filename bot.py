from telegram.ext import Application, CommandHandler
from dotenv import load_dotenv
import os
import json
import requests
import pandas as pd

load_dotenv()

# /start
async def start(update, context):
    await update.message.reply_text("Hallo! selamat datang di bot EDP!")
# /help
async def help_command(update, context):
    await update.message.reply_text("Ini adalah bot EDP. Gunakan /start untuk memulai.")

# Fungsi untuk mendapatkan spesifikasi dari file excel
async def get_spek_excel():
    try:
        # Membaca data dari Excel dan menghapus spasi tersembunyi
        df = pd.read_excel('SE.xlsx', sheet_name="Laporan",dtype=str)  # Pastikan semua kolom dibaca sebagai string
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)  # Hapus spasi pada semua kolom

         # Hanya ambil kolom yang benar-benar terlihat di gambar
        visible_columns = ["KDTK", "NAMA", "STATION", "OS", "CPU_INFO", "CPU_USAGE","LAN SPEED", "SUHU", "BOOT_TIME", "AKTIVASI_WINDOWS", "PARTIAL_KEY_WINDOWS", "CEK KEY", "CEK BCA", "CEK MANDIRI", "EDP"]
        df = df[visible_columns]  # Abaikan kolom yang tersembunyi

        # Hapus baris yang kosong (jika ada baris hidden yang terbaca sebagai NaN)
        df.dropna(how="all", inplace=True)


        # print("📌 Nama kolom di Excel:", df.columns.tolist())
        # Cek apakah kolom yang dibutuhkan ada
        required_columns = {"KDTK", "NAMA", "STATION", "OS", "CPU_INFO", "CPU_USAGE", "LAN SPEED", "SUHU", "BOOT_TIME"}
        missing_columns = required_columns - set(df.columns)
        if missing_columns:
            print(f"❌ Kolom yang hilang: {missing_columns}")
            return {}

        # Konversi ke dictionary dengan key = KDTK
        data = {row["KDTK"].strip().upper(): {
            "KDTK": row["KDTK"],
            "Nama": row["NAMA"],
            "Station": row["STATION"],
            "OS": row["OS"],
            "CPU Info": row["CPU_INFO"],
            "CPU Usage": row["CPU_USAGE"],
            "LAN Speed": row["LAN SPEED"],
            "Suhu": row["SUHU"],
            "Boot Time": row["BOOT_TIME"],
            "Aktivasi Windows": row["AKTIVASI_WINDOWS"],
            "Partial Key Windows": row["PARTIAL_KEY_WINDOWS"],
            "Cek Key": row["CEK KEY"],
            "Cek BCA": row["CEK BCA"],
            "Cek Mandiri": row["CEK MANDIRI"],
            "EDP": row["EDP"],
        } for _, row in df.iterrows()}

        # print("✅ Data dari Excel:", data)  # Debugging
        return data

    except Exception as e:
        print(f"❌ Error membaca Excel: {e}")
        return {}

# /se
async def get_spek_toko(update, context):
    spek = await get_spek_excel()

  

    if context.args:
        station = context.args[0].strip().upper()
        
        if station in spek:
            speks = spek[station]
            hasil = f"""
Berikut Data Service Excellent Tanggal 23-12-2024 :\n\n
=============================================

Toko : {speks['KDTK']} - {speks['Nama']}
Station : {speks['Station']}
OS : {speks['OS']}
Processor : {speks['CPU Info']}
CPU Usage : {speks['CPU Usage']} %
LAN Speed : {speks['LAN Speed']}
Suhu : {speks['Suhu']} °C
Boot Time : {speks['Boot Time']}
Aktivasi Windows : {speks['Aktivasi Windows']}
Key Windows : {speks['Partial Key Windows']}
Cek Key : {speks['Cek Key']}
BCA : {speks['Cek BCA']}
Mandiri : {speks['Cek Mandiri']}
EDP : {speks['EDP']}

============================================="""

                    
            await update.message.reply_text(f"```{hasil}```", parse_mode='Markdown')
        else:
            await update.message.reply_text("Data tidak ditemukan.")
    else:
        await update.message.reply_text("Silakan masukkan kode toko setelah perintah /spek. (contoh: /spek TCXH)")

# fungsi untuk mencari key duplikat
async def find_duplicate_key():
    try:
        df = pd.read_excel('SE.xlsx', sheet_name="Laporan", dtype=str)  # Pastikan semua kolom dibaca sebagai string
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)  # Hapus spasi pada semua kolom
        if "PARTIAL_KEY_WINDOWS" not in df.columns:
            print("❌ Kolom 'PARTIAL_KEY_WINDOWS' tidak ditemukan.")
            return {}
        # Hitung jumlah kemunculan setiap key
        duplicate_keys = df["PARTIAL_KEY_WINDOWS"].value_counts()
        duplicate_keys = duplicate_keys[duplicate_keys > 1]

        # filter baris yang memiliki key duplikat
        df_dupes = df[df["PARTIAL_KEY_WINDOWS"].isin(duplicate_keys.index)]

        # kelompokan berdasarkan key
        result = {}
        for key, group in df_dupes.groupby("PARTIAL_KEY_WINDOWS"):
             result[key] = group[["KDTK", "NAMA", "STATION"]].to_dict(orient="records")
        return result
    except Exception as e:
        print(f"❌ Error mencari duplikat key : {e}")
        return {}
     
# /cekkey

async def show_duplicate_key(update, context):
     duplicates = await find_duplicate_key()
     if not duplicates:
          await update.message.reply_text("Tidak ada key duplikat ditemukan.")
          return
     
     msg = "Daftar Key Duplikat tanggal 23-12-2024 : \n\n"
     for key, entries in duplicates.items():
          msg += "-------------------------------------\n"
          msg += f"🔑 Key : {key}\n"
          for e in entries:
               msg += f"- Toko : {e['KDTK']} - {e['NAMA']}, Station : {e['STATION']}\n"
     
     await update.message.reply_text(msg)
        #   msg += "\n"
        #   message.append(msg)

    #  for msg in message:


# Fungsi utama untuk menjalankan bot
if __name__ == '__main__':
    TOKEN_TELE = os.getenv("TOKEN")

    application = Application.builder().token(TOKEN_TELE).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("se", get_spek_toko))
    application.add_handler(CommandHandler("cekkey", show_duplicate_key))

    print("Bot Telegram berjalan...")
    application.run_polling()
