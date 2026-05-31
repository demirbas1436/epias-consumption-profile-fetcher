
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import fetch_all_multiple_factors as backend
import os

# Uygulama Teması
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class ProfilApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Pencere Ayarları
        self.title("Profil Katsayıları İndirici")
        self.geometry("600x650")
        self.resizable(False, True)
        
        # İkon Ayarla (Eğer dosya varsa)
        if os.path.exists("icon.ico"):
            self.iconbitmap("icon.ico")
            self.after(200, lambda: self.iconbitmap("icon.ico")) # Bazı durumlarda gecikmeli yükleme gerekebilir

        # Grid Yapılandırması
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure((0, 1, 2, 3, 4, 5), weight=0)
        self.grid_rowconfigure(6, weight=1)

        # Başlık
        self.header_label = ctk.CTkLabel(self, text="Profil Katsayıları İndirici", font=("Roboto Medium", 24))
        self.header_label.grid(row=0, column=0, padx=20, pady=(20, 30))

        # 0. Giriş Bilgileri
        self.login_frame = ctk.CTkFrame(self)
        self.login_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.user_label = ctk.CTkLabel(self.login_frame, text="Kullanıcı Adı:", font=("Roboto", 12))
        self.user_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.user_entry = ctk.CTkEntry(self.login_frame, width=180)
        self.user_entry.grid(row=0, column=1, padx=10, pady=10)
        self.user_entry.insert(0, "epias_mail_adresiniz") # Varsayılan

        self.pass_label = ctk.CTkLabel(self.login_frame, text="Şifre:", font=("Roboto", 12))
        self.pass_label.grid(row=0, column=2, padx=10, pady=10, sticky="w")

        self.pass_entry = ctk.CTkEntry(self.login_frame, show="*", width=120)
        self.pass_entry.grid(row=0, column=3, padx=10, pady=10)
        self.pass_entry.insert(0, "epias_sifreniz") # Varsayılan

        # 1. Dönem Seçimi (Modern Dropdown)
        self.date_frame = ctk.CTkFrame(self)
        self.date_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        self.date_label = ctk.CTkLabel(self.date_frame, text="Dönem:", font=("Roboto", 14))
        self.date_label.pack(side="left", padx=20, pady=10)

        # Yıl Seçimi (Manuel Giriş)
        ctk.CTkLabel(self.date_frame, text="Yıl:", font=("Roboto", 12)).pack(side="left", padx=(0, 5), pady=10)
        self.year_entry = ctk.CTkEntry(self.date_frame, width=80)
        self.year_entry.insert(0, "2025")
        self.year_entry.pack(side="left", padx=5, pady=10)

        # Ay Seçimi
        ctk.CTkLabel(self.date_frame, text="Ay:", font=("Roboto", 12)).pack(side="left", padx=(10, 5), pady=10)
        months = [f"{i:02d}" for i in range(1, 13)]
        self.month_combo = ctk.CTkComboBox(self.date_frame, values=months, width=70, state="readonly")
        self.month_combo.set("12")
        self.month_combo.pack(side="left", padx=5, pady=10)

        # Gün Seçimi
        ctk.CTkLabel(self.date_frame, text="Gün:", font=("Roboto", 12)).pack(side="left", padx=(10, 5), pady=10)
        days = [f"{i:02d}" for i in range(1, 32)]
        self.day_combo = ctk.CTkComboBox(self.date_frame, values=days, width=70, state="readonly")
        self.day_combo.set("01")
        self.day_combo.pack(side="left", padx=5, pady=10)

        # 2. Klasör Seçimi
        self.folder_frame = ctk.CTkFrame(self)
        self.folder_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        self.folder_btn = ctk.CTkButton(self.folder_frame, text="Klasör Seç", command=self.select_folder, width=120)
        self.folder_btn.pack(side="left", padx=20, pady=20)

        self.folder_path_label = ctk.CTkLabel(self.folder_frame, text="Dosya Yolu: Varsayılan (Uygulama Yanı)", font=("Roboto", 12), text_color="gray")
        self.folder_path_label.pack(side="left", padx=10, pady=20)
        
        self.selected_folder = ""

        # 3. Çalıştır Butonu
        self.run_btn = ctk.CTkButton(self, text="Verileri İndir ve Excel Oluştur", command=self.start_process, height=50, font=("Roboto Medium", 16))
        self.run_btn.grid(row=4, column=0, padx=20, pady=(20, 10), sticky="ew")

        # 4. Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=5, column=0, padx=20, pady=10, sticky="ew")
        self.progress_bar.set(0)

        # 5. Durum Log Alanı (Modern/Terminal Görünümü)
        self.log_frame = ctk.CTkFrame(self, bg_color="transparent")
        self.log_frame.grid(row=6, column=0, padx=20, pady=20, sticky="nsew")
        
        # Log Başlığı
        ctk.CTkLabel(self.log_frame, text="İşlem Kütüğü (Log)", font=("Roboto", 12, "bold")).pack(anchor="w", pady=(0, 5))

        self.status_textbox = ctk.CTkTextbox(
            self.log_frame, 
            height=120, 
            font=("Consolas", 12), 
            corner_radius=10,
            border_width=2,
            border_color="#3B8ED0", # Temaya uygun mavi ton
            fg_color="#1a1a1a",     # Koyu terminal arka planı
            text_color="#00ff00"    # Matrix/Hacker tarzı yeşil yazı veya beyaz "#ffffff"
        )
        self.status_textbox.pack(fill="both", expand=True)
        self.status_textbox.insert("0.0", ">>> Sistem Hazır.\n>>> Lütfen giriş yapıp dönem giriniz.\n")
        self.status_textbox.configure(state="disabled")

        # 6. Bana Ulaş Butonu
        self.contact_btn = ctk.CTkButton(
            self, 
            text="📧 Bana Ulaş", 
            command=self.show_contact_info,
            height=35,
            font=("Roboto", 12),
            fg_color="#2B2B2B",
            hover_color="#404040"
        )
        self.contact_btn.grid(row=7, column=0, padx=20, pady=(0, 20), sticky="ew")

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.selected_folder = folder_selected
            # Uzun yolları kısaltarak göster
            display_text = folder_selected
            if len(display_text) > 40:
                display_text = "..." + display_text[-37:]
            self.folder_path_label.configure(text=f"Yol: {display_text}", text_color=("black", "white"))
        else:
            self.selected_folder = ""
            self.folder_path_label.configure(text="Dosya Yolu: Varsayılan (Uygulama Yanı)", text_color="gray")

    def log(self, message):
        self.status_textbox.configure(state="normal")
        self.status_textbox.insert("end", message + "\n")
        self.status_textbox.see("end")
        self.status_textbox.configure(state="disabled")

    def update_progress(self, percent, message=None):
        self.progress_bar.set(percent)
        if message:
            self.log(message)

    def show_contact_info(self):
        """Bana Ulaş butonuna tıklandığında modern iletişim ekranı göster"""
        contact_window = ctk.CTkToplevel(self)
        contact_window.title("İletişim & Destek")
        contact_window.geometry("500x550")
        contact_window.resizable(False, False)
        
        # Pencereyi ana pencerenin üstünde tut
        contact_window.transient(self)
        contact_window.grab_set()
        
        # Başlık
        header_frame = ctk.CTkFrame(contact_window, fg_color="#1f538d", corner_radius=0)
        header_frame.pack(fill="x", pady=0)
        
        ctk.CTkLabel(
            header_frame, 
            text="� İletişim & Destek", 
            font=("Roboto Medium", 26),
            text_color="white"
        ).pack(pady=25)
        
        # İçerik Alanı
        content_frame = ctk.CTkFrame(contact_window, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=30, pady=20)
        
        # Açıklama
        ctk.CTkLabel(
            content_frame,
            text="Bu uygulama ile ilgili sorun yaşarsanız veya\nAPI değişiklikleri nedeniyle hata alırsanız benimle iletişime geçin.",
            font=("Roboto", 12),
            justify="center"
        ).pack(pady=(0, 20))
        
        # E-posta
        email_frame = ctk.CTkFrame(content_frame, fg_color="#ebebeb", corner_radius=10, border_width=1, border_color="#cccccc")
        email_frame.pack(fill="x", pady=8)
        ctk.CTkLabel(email_frame, text="📧", font=("Roboto", 18)).pack(side="left", padx=(15, 10), pady=12)
        ctk.CTkLabel(email_frame, text="E-posta:", font=("Roboto", 12, "bold"), text_color="#1a1a1a").pack(side="left", pady=12)
        def copy_email():
            self.clipboard_clear()
            self.clipboard_append("demirbas1436@gmail.com")
            email_btn.configure(text="✅ Kopyalandı!", text_color="#2ecc71")
            self.after(2000, lambda: email_btn.configure(text="demirbas1436@gmail.com", text_color="#1f538d"))

        email_btn = ctk.CTkButton(
            email_frame, 
            text="demirbas1436@gmail.com",
            command=copy_email,
            fg_color="transparent",
            hover_color="#d0d0d0",
            font=("Roboto", 11),
            text_color="#1f538d"
        )
        email_btn.pack(side="left", padx=5, pady=12)
        
        # Telefon
        phone_frame = ctk.CTkFrame(content_frame, fg_color="#ebebeb", corner_radius=10, border_width=1, border_color="#cccccc")
        phone_frame.pack(fill="x", pady=8)
        ctk.CTkLabel(phone_frame, text="📱", font=("Roboto", 18)).pack(side="left", padx=(15, 10), pady=12)
        ctk.CTkLabel(phone_frame, text="Telefon:", font=("Roboto", 12, "bold"), text_color="#1a1a1a").pack(side="left", pady=12)
        ctk.CTkLabel(phone_frame, text="05365689025", font=("Roboto", 11), text_color="#333333").pack(side="left", padx=5, pady=12)
        
        # LinkedIn
        linkedin_frame = ctk.CTkFrame(content_frame, fg_color="#ebebeb", corner_radius=10, border_width=1, border_color="#cccccc")
        linkedin_frame.pack(fill="x", pady=8)
        ctk.CTkLabel(linkedin_frame, text="💼", font=("Roboto", 18)).pack(side="left", padx=(15, 10), pady=12)
        ctk.CTkLabel(linkedin_frame, text="LinkedIn:", font=("Roboto", 12, "bold"), text_color="#1a1a1a").pack(side="left", pady=12)
        linkedin_btn = ctk.CTkButton(
            linkedin_frame, 
            text="Profili Ziyaret Et",
            command=lambda: self.open_url("https://tr.linkedin.com/in/muratdemirbas1436"),
            fg_color="#0077B5",
            hover_color="#005885",
            font=("Roboto", 11),
            text_color="white",
            width=140
        )
        linkedin_btn.pack(side="left", padx=5, pady=12)
        
        # Ayırıcı
        ctk.CTkFrame(content_frame, height=2, fg_color="#3B8ED0").pack(fill="x", pady=15)
        
        # GitHub - Açık Kaynak
        github_label = ctk.CTkLabel(
            content_frame,
            text="🌟 Bu Proje Açık Kaynak!",
            font=("Roboto Medium", 14)
        )
        github_label.pack(pady=(5, 10))
        
        github_frame = ctk.CTkFrame(content_frame, fg_color="#ebebeb", corner_radius=10, border_width=1, border_color="#cccccc")
        github_frame.pack(fill="x", pady=8)
        ctk.CTkLabel(github_frame, text="⭐", font=("Roboto", 18)).pack(side="left", padx=(15, 10), pady=12)
        ctk.CTkLabel(github_frame, text="GitHub:", font=("Roboto", 12, "bold"), text_color="#1a1a1a").pack(side="left", pady=12)
        github_btn = ctk.CTkButton(
            github_frame, 
            text="Kaynak Kodları Görüntüle",
            command=lambda: self.open_url("https://github.com/demirbas1436"),
            fg_color="#24292e",
            hover_color="#1a1e22",
            font=("Roboto", 11),
            text_color="white",
            width=180
        )
        github_btn.pack(side="left", padx=5, pady=12)
        
        # Alt bilgi
        ctk.CTkLabel(
            content_frame,
            text="Geri bildirimleriniz için teşekkürler! 🙏",
            font=("Roboto", 11, "italic"),
            text_color="gray"
        ).pack(pady=(15, 5))
        
        # Kapat butonu
        ctk.CTkButton(
            contact_window, 
            text="Kapat", 
            command=contact_window.destroy,
            width=150,
            height=35,
            font=("Roboto", 12)
        ).pack(pady=(0, 20))

    def open_url(self, url):
        """URL'yi varsayılan tarayıcıda aç"""
        import webbrowser
        webbrowser.open(url)

    def start_process(self):
        username = self.user_entry.get().strip()
        password = self.pass_entry.get().strip()
        
        # Dropdown'lardan tarihi al
        year = self.year_entry.get().strip()
        month = self.month_combo.get()
        day = self.day_combo.get()
        donem = f"{year}-{month}-{day}"
        
        if not username or not password:
            messagebox.showwarning("Uyarı", "Kullanıcı adı ve şifre boş olamaz.")
            return

        self.run_btn.configure(state="disabled", text="İşleniyor...")
        self.log(f"--- İşlem Başlatıldı: {donem} ---")
        self.progress_bar.set(0)
        
        # İşlemi arkaplanda çalıştır ki arayüz donmasın
        threading.Thread(target=self.run_backend_process, args=(donem, self.selected_folder, username, password), daemon=True).start()

    def run_backend_process(self, donem, folder, username, password):
        # Callback wrapper to be thread-safe
        def callback(percent, msg):
            self.after(0, lambda: self.update_progress(percent, msg))

        try:
            success, msg = backend.run_process(donem, folder, username, password, progress_callback=callback)
            # self.log(msg) # Callback zaten logluyor
            if success:
                self.after(0, lambda: messagebox.showinfo("Başarılı", msg))
            else:
                self.after(0, lambda: messagebox.showerror("Hata", msg))
        except Exception as e:
            self.log(f"Kritik Hata: {str(e)}")
            self.after(0, lambda: messagebox.showerror("Hata", f"Beklenmeyen bir hata oluştu:\n{e}"))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal", text="Verileri İndir ve Excel Oluştur"))

if __name__ == "__main__":
    app = ProfilApp()
    app.mainloop()
