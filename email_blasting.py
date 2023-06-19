import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from tkinter import messagebox

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(tk.END, file_path)

def send_emails():
    file_path = file_entry.get()
    email_pengirim = email_entry.get()
    password_pengirim = password_entry.get()
    judul_email = judul_entry.get()
    template_email = isi_text.get("1.0", tk.END)
    jumlah_email_str = jumlah_entry.get()  # Mengambil jumlah email yang diinput

    if not file_path:
        messagebox.showerror("Error", "Pilih file Excel terlebih dahulu!")
        return

    if not email_pengirim or not password_pengirim:
        messagebox.showerror("Error", "Masukkan email pengirim dan password!")
        return

    if not judul_email:
        messagebox.showerror("Error", "Masukkan judul email!")
        return

    if not template_email:
        messagebox.showerror("Error", "Masukkan template email!")
        return

    if not jumlah_email_str:
        messagebox.showerror("Error", "Isi jumlah email yang akan dikirim!")
        return

    try:
        jumlah_email = int(jumlah_email_str)
    except ValueError:
        messagebox.showerror("Error", "Jumlah email harus berupa angka!")
        return

    # Membaca file Excel
    try:
        df = pd.read_excel(file_path)
    except pd.errors.ParserError:
        messagebox.showerror("Error", "File Excel tidak valid!")
        return

    total_emails = min(len(df), jumlah_email)  # Menentukan total email yang akan dikirim
    progress_bar["maximum"] = total_emails

    # Mengirim email ke setiap penerima
    for index, row in df.iterrows():
        if index >= jumlah_email:  # Menghentikan pengiriman setelah mencapai jumlah email yang diinput
            break

        email_penerima = row['Email']
        nama = row['Nama']
        link = row['Link']

        # Mengganti placeholder dalam template dengan nilai yang sesuai
        isi_email = template_email.replace("{nama}", str(row['Nama'])).replace("{link}", str(row['Link']))

        # Membuat objek pesan email
        msg = MIMEText(isi_email)
        msg['Subject'] = judul_email
        msg['From'] = email_pengirim
        msg['To'] = email_penerima

        # Mengirim email menggunakan SMTP server
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_pengirim, password_pengirim)
            server.sendmail(email_pengirim, email_penerima, msg.as_string())
            server.quit()
        except smtplib.SMTPException:
            messagebox.showerror("Error", "Gagal mengirim email")

        # Memperbarui nilai progress bar
        progress_bar["value"] = index + 1
        window.update_idletasks()

    messagebox.showinfo("Info", "Email blasting selesai!")

# Membuat GUI
window = tk.Tk()
window.title("Email Blasting")

# Membuat style dengan menggunakan ttk
style = ttk.Style(window)
style.configure("TButton", padding=10)

# Frame untuk judul form
title_frame = ttk.Frame(window)
title_frame.grid(row=0, column=0, columnspan=2, pady=20)
title_label = ttk.Label(title_frame, text="Email Blasting", font=("Arial", 18))
title_label.pack()

# Frame untuk memilih file Excel
file_frame = ttk.Frame(window)
file_frame.grid(row=1, column=0, padx=20, pady=0, sticky=tk.W)
file_label = ttk.Label(file_frame, text="Pilih File Excel")
file_label.grid(row=0, padx=20, column=0, sticky=tk.W)
file_entry = ttk.Entry(file_frame)
file_entry.grid(row=0, padx=26, column=1)
browse_button = ttk.Button(file_frame, text="Browse", command=browse_file)
browse_button.grid(row=0, padx=10, pady=5, column=2, sticky=tk.W)
petunjuk_label = ttk.Label(file_frame, text="*Pastikan file Excel memiliki format .xlsx")
petunjuk_label.grid(row=1, columnspan=3, padx=20, pady=(0, 10), sticky=tk.W)

# Frame untuk informasi pengirim email
email_frame = ttk.Frame(window)
email_frame.grid(row=2, column=0, padx=20, pady=10, sticky=tk.W)
email_label = ttk.Label(email_frame, text="Email Pengirim")
email_label.grid(row=0, padx=20, column=0, sticky=tk.W)
email_entry = ttk.Entry(email_frame)
email_entry.grid(row=0, padx=21, column=1)

password_frame = ttk.Frame(window)
password_frame.grid(row=3, column=0, padx=20, pady=10, sticky=tk.W)
password_label = ttk.Label(password_frame, text="Password Pengirim")
password_label.grid(row=0, padx=20, column=0, sticky=tk.W)
password_entry = ttk.Entry(password_frame, show="*")
password_entry.grid(row=0, column=1)

# Frame untuk judul email
judul_frame = ttk.Frame(window)
judul_frame.grid(row=4, column=0, padx=20, pady=10, sticky=tk.W)
judul_label = ttk.Label(judul_frame, text="Judul Email")
judul_label.grid(row=0, padx=20, column=0, sticky=tk.W)
judul_entry = ttk.Entry(judul_frame)
judul_entry.grid(row=0, padx=41, column=1)

# Frame untuk template email
isi_frame = ttk.Frame(window)
isi_frame.grid(row=5, column=0, padx=20, pady=10, sticky=tk.W)
isi_label = ttk.Label(isi_frame, text="Template Email")
isi_label.grid(row=0, pady=10, column=0, sticky=tk.W)
isi_text = tk.Text(isi_frame, width=50, height=5)
isi_text.grid(row=1, column=0)

# Frame untuk jumlah email
jumlah_frame = ttk.Frame(window)
jumlah_frame.grid(row=6, column=0, padx=20, pady=10, sticky=tk.W)
jumlah_label = ttk.Label(jumlah_frame, text="Jumlah Email")
jumlah_label.grid(row=0, padx=20, column=0, sticky=tk.W)
jumlah_entry = ttk.Entry(jumlah_frame)
jumlah_entry.grid(row=0, padx=41, column=1)

# Tombol untuk mengirim email
send_button = ttk.Button(window, text="Kirim Email", command=send_emails)
send_button.grid(row=7, column=0, padx=20, pady=20)

# Progress bar
progress_bar = ttk.Progressbar(window, orient=tk.HORIZONTAL, length=200, mode='determinate')
progress_bar.grid(row=8, column=0, padx=20, pady=10)

window.mainloop()
