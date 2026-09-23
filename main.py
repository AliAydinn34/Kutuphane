import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont
import sqlite3
import random
import csv
import codecs
import io
import urllib.request
import json
import ssl
import os
import shutil
from datetime import datetime

# --- SİSTEM DİZİNİ AYARLARI (BELGELER KLASÖRÜNE KAYIT) ---
APP_DIR = os.path.join(os.path.expanduser('~'), 'Documents', 'KutuphaneSistemi')
if not os.path.exists(APP_DIR):
    os.makedirs(APP_DIR)

DB_YOLU = os.path.join(APP_DIR, "kutuphane.db")
YEDEK_DIR = os.path.join(APP_DIR, "Yedekler")

# --- 1. Veritabanı Kurulumu ---
def veritabani_hazirla():
    baglanti = sqlite3.connect(DB_YOLU)
    imlec = baglanti.cursor()
    imlec.execute('''
        CREATE TABLE IF NOT EXISTS Kitaplar (
            ID INTEGER PRIMARY KEY, 
            Isim TEXT NOT NULL,
            Yazar TEXT NOT NULL,
            Yayinevi TEXT,
            BasimTarihi TEXT,
            SayfaSayisi TEXT,
            Olcu TEXT,
            Kondisyon TEXT,
            RafKonumu TEXT,
            Disarida BOOLEAN DEFAULT 0,
            EmanetAlan TEXT,
            EmanetTarihi TEXT
        )
    ''')
    
    try:
        imlec.execute("ALTER TABLE Kitaplar ADD COLUMN Etiketler TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass 
        
    try:
        imlec.execute("ALTER TABLE Kitaplar ADD COLUMN Dil TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass 
        
    imlec.execute('''
        CREATE TABLE IF NOT EXISTS Ayarlar (
            Anahtar TEXT PRIMARY KEY,
            Deger TEXT
        )
    ''')
    imlec.execute("INSERT OR IGNORE INTO Ayarlar (Anahtar, Deger) VALUES ('Baslik', 'Kütüphane Envanteri')")
    baglanti.commit()
    baglanti.close()

# --- 2. Arayüz ve Uygulama Sınıfı ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class KutuphaneUygulamasi(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Kütüphane Yönetim Sistemi")
        self.geometry("1300x700") 
        self.minsize(1200, 550)
        
        # --- YENİ: PENCERE VE GÖREV ÇUBUĞU SİMGESİ ---
        import sys
        if getattr(sys, 'frozen', False):
            self.iconbitmap(sys.executable) # Exe yapıldığında kendi dış simgesini kopyalar
        else:
            try:
                self.iconbitmap("logo.ico") # VS Code'da denerken klasördeki logoyu kullanır
            except:
                pass
        # ---------------------------------------------
        
        self.tooltip_pencere = None
        self.emanet_bilgileri = {}
        self.siralama_durumu = {"Yazar": False, "İsim": False, "Dil": False, "Basım": False, "Sayfa": False, "ID": False}
        self.aktif_siralama = None  
        self.coklu_secim_modu = False

        self.sabit_degerler = {"yazar": "", "isim": "", "dil": "", "yayinevi": "", "basim": "", "sayfa": "", "olcu": "", "kondisyon": "", "raf": ""}
        self.sabit_kilitler = {"yazar": False, "isim": False, "dil": False, "yayinevi": False, "basim": False, "sayfa": False, "olcu": False, "kondisyon": False, "raf": False}
      
        self.olcu_ref = None
        self.kondisyon_ref = None
        self.raf_ref = None

        self.bind("<Control-f>", lambda e: self.arama_kutusu.focus())
        self.bind("<Control-F>", lambda e: self.arama_kutusu.focus())
        self.bind("<Control-n>", lambda e: self.kitap_ekle_penceresi())
        self.bind("<Control-N>", lambda e: self.kitap_ekle_penceresi())
        self.bind("<Delete>", self.kisayol_sil)
        
        self.protocol("WM_DELETE_WINDOW", self.kapatirken_yedekle)

        # --- Üst Kısım: Genel Kapsayıcı ---
        self.ust_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.ust_frame.pack(pady=(15, 5), padx=20, fill="x")

        # Sol Taraf: Başlık
        self.baslik_frame = ctk.CTkFrame(self.ust_frame, fg_color="transparent")
        self.baslik_frame.pack(side="left", anchor="center")

        mevcut_baslik = self.baslik_oku()
        self.baslik = ctk.CTkLabel(self.baslik_frame, text=mevcut_baslik, font=("Arial", 28, "bold"))
        self.baslik.pack(side="left")
        
        self.baslik_duzenle_btn = ctk.CTkButton(self.baslik_frame, text="📝", font=("Arial", 16), width=35, height=35, 
                                                fg_color="transparent", hover_color="#3a3a3a", command=self.baslik_degistir_penceresi)
        self.baslik_duzenle_btn.pack(side="left", padx=10)
        
        # Sağ Taraf: İstatistik ve Excel İşlemleri
        self.sag_panel = ctk.CTkFrame(self.ust_frame, fg_color="transparent")
        self.sag_panel.pack(side="right", anchor="center")

        self.istatistik_frame = ctk.CTkFrame(self.sag_panel, fg_color="#2b2b2b", corner_radius=8)
        self.istatistik_frame.pack(fill="x", pady=(0, 5))
        
        self.lbl_toplam = ctk.CTkLabel(self.istatistik_frame, text="Toplam Kitap Sayısı: 0", font=("Arial", 13, "bold"))
        self.lbl_toplam.pack(pady=(5, 2))
        
        self.istatistik_alt_frame = ctk.CTkFrame(self.istatistik_frame, fg_color="transparent")
        self.istatistik_alt_frame.pack(pady=(0, 5))
        
        self.lbl_rafta = ctk.CTkLabel(self.istatistik_alt_frame, text="Rafta: 0", font=("Arial", 12, "bold"), text_color="white")
        self.lbl_rafta.pack(side="left", padx=15)
        
        self.lbl_disarida = ctk.CTkLabel(self.istatistik_alt_frame, text="Rafta Değil: 0", font=("Arial", 12, "bold"), text_color="#ffcccc")
        self.lbl_disarida.pack(side="left", padx=15)

        self.excel_frame = ctk.CTkFrame(self.sag_panel, fg_color="transparent")
        self.excel_frame.pack(fill="x")

        self.disa_aktar_btn = ctk.CTkButton(self.excel_frame, text="📤 Excel'e Aktar", font=("Arial", 11, "bold"), 
                                            width=110, height=28, command=self.excel_disa_aktar)
        self.disa_aktar_btn.pack(side="left", padx=(0, 5), expand=True)

        self.ice_aktar_btn = ctk.CTkButton(self.excel_frame, text="📥 Excel'den Yükle", font=("Arial", 11, "bold"), 
                                           width=110, height=28, command=self.excel_ice_aktar)
        self.ice_aktar_btn.pack(side="left", expand=True)

        # --- Kontrol Paneli (Sol Orta Kısım) ---
        self.kontrol_paneli = ctk.CTkFrame(self, fg_color="transparent")
        self.kontrol_paneli.pack(pady=10, padx=20, fill="x", anchor="w")
        
        self.ekle_butonu = ctk.CTkButton(self.kontrol_paneli, text="+ Yeni Kitap Ekle", font=("Arial", 14, "bold"), command=self.kitap_ekle_penceresi, width=140)
        self.ekle_butonu.pack(side="left", padx=(0, 10))
        
        self.dinamik_arama_alani = ctk.CTkFrame(self.kontrol_paneli, fg_color="transparent")
        self.dinamik_arama_alani.pack(side="left")
        
        self.arama_kutusu = ctk.CTkEntry(self.dinamik_arama_alani, placeholder_text="Kitap İsmi ile ara...", width=200)
        self.arama_kutusu.pack(side="left", padx=5)
        self.arama_kutusu.bind("<KeyRelease>", self.arama_yap)
        
        self.aralik_frame = ctk.CTkFrame(self.dinamik_arama_alani, fg_color="transparent")
        
        self.aralik_bas = ctk.CTkEntry(self.aralik_frame, placeholder_text="Örn: 1975", width=70)
        self.aralik_bas.pack(side="left")
        self.aralik_bas.bind("<KeyRelease>", self.arama_yap)
        
        ctk.CTkLabel(self.aralik_frame, text="-").pack(side="left", padx=5)
        
        self.aralik_bit = ctk.CTkEntry(self.aralik_frame, placeholder_text="2002", width=70)
        self.aralik_bit.pack(side="left")
        self.aralik_bit.bind("<KeyRelease>", self.arama_yap)

        self.durum_secici = ctk.CTkComboBox(self.dinamik_arama_alani, values=["Rafta", "Rafta değil"], command=self.arama_yap, width=120)
        
        self.kategoriler = [
            "Kitap İsmi", "Yazar", "Dil", "Yayınevi", "Basım Yılı (Aralık)", 
            "Sayfa Sayısı (Aralık)", "Kondisyon", "Raf Konumu", "Durum (Rafta/Dışarıda)", "Etiket İle Arama", "ID"
        ]
        self.arama_kategorisi = ctk.CTkComboBox(self.kontrol_paneli, values=self.kategoriler, width=170, command=self.kategori_degisti)
        self.arama_kategorisi.pack(side="left", padx=10)
        self.arama_kategorisi.set("Kitap İsmi")

        self.secim_frame = ctk.CTkFrame(self.kontrol_paneli, fg_color="transparent")
        self.secim_frame.pack(side="left", padx=(10, 0))

        self.coklu_sec_btn = ctk.CTkButton(self.secim_frame, text="Çoklu Seçim: Kapalı", font=("Arial", 11, "bold"), width=120, height=28, command=self.coklu_secim_toggle)
        self.coklu_sec_btn.pack(side="left", padx=(0, 5))

        self.tumunu_sec_btn = ctk.CTkButton(self.secim_frame, text="Tümünü Seç", font=("Arial", 11, "bold"), width=90, height=28, command=self.tumunu_sec)
        self.tumunu_sec_btn.pack(side="left")
        
        # --- Tablo Alanı ---
        self.tablo_alani = ctk.CTkFrame(self)
        self.tablo_alani.pack(fill="both", expand=True, padx=20, pady=(10, 20))
        
        self.tablo_olustur()
        self.verileri_yukle()

    # --- Hayalet Yedekleme Motoru ---
    def kapatirken_yedekle(self):
        try:
            if not os.path.exists(YEDEK_DIR):
                os.makedirs(YEDEK_DIR)
            tarih = datetime.now().strftime("%d_%m_%Y")
            hedef = os.path.join(YEDEK_DIR, f"yedek_{tarih}.db")
            shutil.copy2(DB_YOLU, hedef)
        except Exception:
            pass
        finally:
            self.destroy()

    def kisayol_sil(self, event):
        secili_satirlar = self.tablo.selection()
        if not secili_satirlar: return
        
        secili_id_listesi = [int(self.tablo.item(satir, "values")[0]) for satir in secili_satirlar]
        if len(secili_id_listesi) == 1:
            baslik = f"Kitap: {self.tablo.item(secili_satirlar[0], 'values')[2]}"
        else:
            baslik = f"{len(secili_id_listesi)} Kitap Seçili"
            
        self.kitap_sil(secili_id_listesi, baslik)

    # --- Çoklu Seçim Modülleri ---
    def coklu_secim_toggle(self):
        self.coklu_secim_modu = not self.coklu_secim_modu
        if self.coklu_secim_modu:
            self.coklu_sec_btn.configure(text="Çoklu Seçim: Açık", fg_color="#27ae60", hover_color="#2ecc71")
        else:
            self.coklu_sec_btn.configure(text="Çoklu Seçim: Kapalı", fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"], hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"])
            self.tablo.selection_remove(self.tablo.selection())

    def tumunu_sec(self):
        satirlar = self.tablo.get_children()
        secili_satirlar = self.tablo.selection()
        if len(satirlar) > 0 and len(secili_satirlar) == len(satirlar):
            self.tablo.selection_remove(satirlar)
        else:
            self.tablo.selection_set(satirlar)

    def tablo_ozel_tiklama(self, event):
        region = self.tablo.identify_region(event.x, event.y)
        if region in ("cell", "tree"):
            col = self.tablo.identify_column(event.x)
            satir_id = self.tablo.identify_row(event.y)
            if col == "#13":
                if satir_id:
                    if satir_id not in self.tablo.selection():
                        if not self.coklu_secim_modu:
                            self.tablo.selection_set(satir_id)
                        else:
                            self.tablo.selection_add(satir_id)
                return "break" 
            if self.coklu_secim_modu:
                if satir_id:
                    if satir_id in self.tablo.selection():
                        self.tablo.selection_remove(satir_id)
                    else:
                        self.tablo.selection_add(satir_id)
                return "break" 

    def tablo_tiklama(self, event):
        region = self.tablo.identify_region(event.x, event.y)
        if region == "cell":
            col = self.tablo.identify_column(event.x)
            if col == "#13": 
                satir_id = self.tablo.identify_row(event.y)
                if satir_id:
                    self.islem_menusu_goster(event)

    def cift_tikla_duzenle(self, event):
        secili_satirlar = self.tablo.selection()
        if len(secili_satirlar) == 1:
            kitap_id = int(self.tablo.item(secili_satirlar[0], "values")[0])
            self.kitap_guncelle_penceresi(kitap_id)

    def ort_tus_bas(self, event):
        self.tablo.scan_mark(event.x, event.y)

    def ort_tus_surukle(self, event):
        self.tablo.scan_dragto(event.x, event.y, gain=1)

    # --- Etiket Yönetimi Paneli ---
    def etiket_penceresi(self, id_listesi, baslik_metni):
        pencere = ctk.CTkToplevel(self)
        tek_kitap = len(id_listesi) == 1
        
        pencere.title("Etiket Yönetimi" if tek_kitap else "Toplu Etiketleme")
        self.merkezde_baslat(pencere, 400, 450)
        pencere.attributes('-topmost', True)
        pencere.grab_set()

        ctk.CTkLabel(pencere, text="Etiket Yönetimi" if tek_kitap else "Toplu Etiketleme", font=("Arial", 18, "bold"), text_color="#1f538d").pack(pady=(15, 5))
        ctk.CTkLabel(pencere, text=baslik_metni, font=("Arial", 12, "italic"), justify="center", wraplength=350).pack(pady=(0, 15))

        g_frame = ctk.CTkFrame(pencere, fg_color="transparent")
        g_frame.pack(fill="x", padx=30, pady=5)

        etiket_entry = ctk.CTkEntry(g_frame, placeholder_text="Örn: #Sosyoloji", height=32)
        etiket_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        liste_frame = ctk.CTkScrollableFrame(pencere, fg_color="#2b2b2b")
        liste_frame.pack(fill="both", expand=True, padx=30, pady=15)

        def etiket_ekle(event=None):
            yeni_etiket = etiket_entry.get().strip()
            if not yeni_etiket: return

            yeni_etiket = yeni_etiket.replace(" ", "")
            if not yeni_etiket.startswith("#"):
                yeni_etiket = "#" + yeni_etiket

            baglanti = sqlite3.connect(DB_YOLU)
            imlec = baglanti.cursor()
            
            for k_id in id_listesi:
                imlec.execute("SELECT Etiketler FROM Kitaplar WHERE ID=?", (k_id,))
                mevcut = imlec.fetchone()[0]
                if not mevcut: mevcut = ""

                mevcut_liste = mevcut.split()
                if yeni_etiket not in mevcut_liste:
                    mevcut_liste.append(yeni_etiket)
                    yeni_string = " ".join(mevcut_liste)
                    baglanti.execute("UPDATE Kitaplar SET Etiketler=? WHERE ID=?", (yeni_string, k_id))
            
            baglanti.commit()
            baglanti.close()
            etiket_entry.delete(0, tk.END)
            etiketleri_goster()
            self.arama_yap()

        ekle_btn = ctk.CTkButton(g_frame, text="Ekle", font=("Arial", 12, "bold"), width=60, height=32, fg_color="#27ae60", hover_color="#2ecc71", command=etiket_ekle)
        ekle_btn.pack(side="right")
        etiket_entry.bind("<Return>", etiket_ekle)

        def etiket_sil(silinecek_etiket):
            baglanti = sqlite3.connect(DB_YOLU)
            imlec = baglanti.cursor()
            imlec.execute("SELECT Etiketler FROM Kitaplar WHERE ID=?", (id_listesi[0],))
            mevcut = imlec.fetchone()[0]
            if mevcut:
                mevcut_liste = mevcut.split()
                if silinecek_etiket in mevcut_liste:
                    mevcut_liste.remove(silinecek_etiket)
                    yeni_string = " ".join(mevcut_liste)
                    baglanti.execute("UPDATE Kitaplar SET Etiketler=? WHERE ID=?", (yeni_string, id_listesi[0]))
                    baglanti.commit()
            baglanti.close()
            etiketleri_goster()
            self.arama_yap()

        def etiketleri_goster():
            for widget in liste_frame.winfo_children():
                widget.destroy()

            if not tek_kitap:
                ctk.CTkLabel(liste_frame, text="Toplu etiketleme modunda mevcut\netiketler görüntülenmez.\n\nEklediğiniz etiketler seçili\ntüm kitaplara işlenecektir.", font=("Arial", 12, "italic"), text_color="gray").pack(pady=40)
                return

            baglanti = sqlite3.connect(DB_YOLU)
            imlec = baglanti.cursor()
            imlec.execute("SELECT Etiketler FROM Kitaplar WHERE ID=?", (id_listesi[0],))
            mevcut = imlec.fetchone()[0]
            baglanti.close()

            if mevcut:
                for tag in mevcut.split():
                    t_frame = ctk.CTkFrame(liste_frame, fg_color="#1f538d", corner_radius=15)
                    t_frame.pack(side="top", fill="x", pady=4, padx=5)

                    lbl = ctk.CTkLabel(t_frame, text=tag, font=("Arial", 12, "bold"), text_color="white")
                    lbl.pack(side="left", padx=10, pady=5)

                    sil_btn = ctk.CTkButton(t_frame, text="✖", width=25, height=25, font=("Arial", 12, "bold"), fg_color="transparent", hover_color="#e74c3c", command=lambda t=tag: etiket_sil(t))
                    sil_btn.pack(side="right", padx=5)

        etiketleri_goster()

    def islem_menusu_goster(self, event):
        secili_satirlar = self.tablo.selection()
        if not secili_satirlar:
            return

        secili_id_listesi = []
        kitap_isimleri = []
        yazarlar = []
        for satir in secili_satirlar:
            degerler = self.tablo.item(satir, "values")
            secili_id_listesi.append(int(degerler[0]))
            yazarlar.append(degerler[1])
            kitap_isimleri.append(degerler[2])

        menu = tk.Menu(self, tearoff=0, bg="#2b2b2b", fg="white", font=("Arial", 11))
        
        if len(secili_satirlar) == 1:
            baslik_metni = f"Kitap: {kitap_isimleri[0][:20]}..." if len(kitap_isimleri[0])>20 else f"Kitap: {kitap_isimleri[0]}"
            durum_metni = "Durum Güncelle"
            sil_metni = "Sil"
        else:
            baslik_metni = f"{len(secili_satirlar)} Kitap Seçili"
            durum_metni = f"Durum Güncelle ({len(secili_satirlar)} Kitap)"
            sil_metni = f"Seçili {len(secili_satirlar)} Kitabı Sil"

        menu.add_command(label=durum_metni, command=lambda: self.durum_guncelle_penceresi(secili_id_listesi, baslik_metni))
        
        if len(secili_satirlar) == 1:
            menu.add_command(label="Bilgileri Güncelle", command=lambda: self.kitap_guncelle_penceresi(secili_id_listesi[0]))
            menu.add_command(label="Etiketler", command=lambda: self.etiket_penceresi(secili_id_listesi, f"{yazarlar[0]}\n{kitap_isimleri[0]}"))
        else:
            menu.add_command(label="Toplu Etiketle", command=lambda: self.etiket_penceresi(secili_id_listesi, f"Seçili {len(secili_id_listesi)} kitabın tamamına etiket ekle"))
            
        menu.add_separator()
        menu.add_command(label=sil_metni, command=lambda: self.kitap_sil(secili_id_listesi, baslik_metni))
        menu.tk_popup(event.x_root, event.y_root)

    # --- Excel Export / Import Modülleri ---
    def excel_disa_aktar(self):
        dosya_yolu = filedialog.asksaveasfilename(defaultextension=".csv",
                                                  title="Listeyi Excel (CSV) Olarak Kaydet",
                                                  filetypes=[("Excel CSV Dosyası", "*.csv")])
        if not dosya_yolu:
            return
            
        try:
            with open(dosya_yolu, mode='w', newline='', encoding='utf-8-sig') as dosya:
                yazici = csv.writer(dosya, delimiter=';') 
                
                basliklar = ["ID", "Yazar", "İsim", "Dil", "Yayınevi", "Basım Yılı", "Sayfa Sayısı", "Ölçü", "Kondisyon", "Raf Konumu", "Durum", "Emanet Alan", "Emanet Tarihi"]
                yazici.writerow(basliklar)
                
                for k in self.tablo.get_children(""):
                    degerler = self.tablo.item(k, "values")
                    kitap_id = int(degerler[0])
                    durum = degerler[10]
                    
                    emanet_kisi = ""
                    emanet_tarihi = ""
                    
                    if durum == "Rafta değil":
                        bilgi = self.emanet_bilgileri.get(kitap_id, "")
                        if bilgi:
                            parcalar = bilgi.split("'a teslim edildi.\n")
                            if len(parcalar) == 2:
                                i_alan = parcalar[0].split("'a teslim edildi.\n")
                                emanet_kisi = i_alan[0] if i_alan else parcalar[0]
                                emanet_tarihi = parcalar[1]
                                
                    satir_verisi = [degerler[0], degerler[1], degerler[2], degerler[3], degerler[4], degerler[5], degerler[6], degerler[7], degerler[8], degerler[9], durum, emanet_kisi, emanet_tarihi]
                    yazici.writerow(satir_verisi)
                    
            messagebox.showinfo("Başarılı", "Tablodaki liste başarıyla Excel formatında dışa aktarıldı!")
        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında bir hata oluştu:\n{e}")

    def excel_ice_aktar(self):
        bilgi_metni = ("İçe aktaracağınız dosya, daha önce bu programdan dışa aktarılmış veya "
                       "aynı başlık sırasına sahip bir Excel (.csv) dosyası olmalıdır.\n\n"
                       "Dosyadaki kitaplar sisteme YENİ KİTAP olarak, yeni ID numaralarıyla eklenecektir.\nDevam edilsin mi?")
        
        if not messagebox.askyesno("Bilgi", bilgi_metni):
            return
            
        dosya_yolu = filedialog.askopenfilename(title="Excel (CSV) Dosyası Seç", filetypes=[("Excel CSV Dosyası", "*.csv")])
        if not dosya_yolu:
            return
            
        try:
            with open(dosya_yolu, 'rb') as f:
                ham_veri = f.read()
                
            icerik = ""
            try:
                icerik = ham_veri.decode('utf-8-sig') 
            except UnicodeDecodeError:
                try:
                    icerik = ham_veri.decode('cp1254') 
                except UnicodeDecodeError:
                    icerik = ham_veri.decode('latin-1', errors='replace') 
                    
            if not icerik.strip():
                messagebox.showwarning("Uyarı", "Seçilen dosya boş veya okunamadı!")
                return
                
            ilk_satir = icerik.split('\n')[0]
            ayirici = ';'
            if ';' in ilk_satir:
                ayirici = ';'
            elif ',' in ilk_satir:
                ayirici = ','
            elif '\t' in ilk_satir:
                ayirici = '\t'
                
            f_io = io.StringIO(icerik)
            okuyucu = csv.reader(f_io, delimiter=ayirici)
            next(okuyucu, None) # Başlık satırını atlar
            
            baglanti = sqlite3.connect(DB_YOLU)
            imlec = baglanti.cursor()
            eklenen_sayi = 0
            
            # YENİ EKLENEN AKILLI FORMAT KONTROLÜ
            ilk_satir_kucuk = ilk_satir.lower()
            hucre_sayisi = len(ilk_satir.split(ayirici))
            # Eğer başlıkta küçük/büyük fark etmeksizin "dil" geçiyorsa veya 13+ sütun varsa yeni formattır
            yeni_format_mi = "dil" in ilk_satir_kucuk or hucre_sayisi >= 13

            for satir in okuyucu:
                if not satir or len(satir) < 3: 
                    continue
                    
                # Yeni format için eksik sütun tamamlama limiti 13'e çıkarıldı
                while len(satir) < 13:
                    satir.append("")
                    
                yazar = satir[1].strip()
                isim = satir[2].strip()
                
                if not isim or not yazar: 
                    continue
                
                # Geriye Dönük Uyumluluk (Akıllı Kontrol)
                if yeni_format_mi:
                    dil = satir[3].strip()
                    yayinevi = satir[4].strip()
                    basim = satir[5].strip()
                    sayfa = satir[6].strip()
                    olcu = satir[7].strip()
                    kondisyon = satir[8].strip()
                    raf = satir[9].strip()
                else:
                    dil = ""
                    yayinevi = satir[3].strip()
                    basim = satir[4].strip()
                    sayfa = satir[5].strip()
                    olcu = satir[6].strip()
                    kondisyon = satir[7].strip()
                    raf = satir[8].strip()
                
                while True:
                    yeni_id = random.randint(1, 9999)
                    imlec.execute("SELECT ID FROM Kitaplar WHERE ID=?", (yeni_id,))
                    if not imlec.fetchone():
                        break
                        
                baglanti.execute('''
                    INSERT INTO Kitaplar (ID, Isim, Yazar, Dil, Yayinevi, BasimTarihi, SayfaSayisi, Olcu, Kondisyon, RafKonumu)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (yeni_id, isim, yazar, dil, yayinevi, basim, sayfa, olcu, kondisyon, raf))
                
                eklenen_sayi += 1
                
            baglanti.commit()
            baglanti.close()
            self.arama_yap() 
            
            if eklenen_sayi > 0:
                messagebox.showinfo("Başarılı", f"{eklenen_sayi} adet kitap başarıyla içe aktarıldı!")
            else:
                messagebox.showwarning("Uyarı", "Dosyadan hiçbir kitap okunamadı.\nBaşlıkların ve sütunların doğruluğunu kontrol edin.")
            
        except Exception as e:
            messagebox.showerror("Hata", f"İçe aktarma sırasında beklenmedik bir hata oluştu.\n\nDetay: {e}")

    def sutuna_gore_sirala(self, sutun, toggle=True):
        if toggle:
            ters_mi = self.siralama_durumu[sutun]
            self.siralama_durumu[sutun] = not ters_mi
            self.aktif_siralama = sutun  
        else:
            ters_mi = not self.siralama_durumu[sutun]
            
        for col in self.siralama_durumu.keys():
            self.tablo.heading(col, text=col)
            
        ok = " ▼" if ters_mi else " ▲"
        self.tablo.heading(sutun, text=sutun + ok)
        
        liste = [(self.tablo.set(k, sutun), k) for k in self.tablo.get_children("")]
        
        if sutun in ["Basım", "Sayfa", "ID"]:
            def sayiya_cevir(deger):
                try: return int(deger[0])
                except ValueError: return 0
            liste.sort(key=sayiya_cevir, reverse=ters_mi)
        else:
            liste.sort(key=lambda t: t[0].lower(), reverse=ters_mi)
            
        for index, (_, k) in enumerate(liste):
            self.tablo.move(k, '', index)

    def istatistikleri_guncelle(self):
        baglanti = sqlite3.connect(DB_YOLU)
        imlec = baglanti.cursor()
        
        imlec.execute("SELECT COUNT(*) FROM Kitaplar")
        toplam = imlec.fetchone()[0]
        
        imlec.execute("SELECT COUNT(*) FROM Kitaplar WHERE Disarida = 1")
        disarida = imlec.fetchone()[0]
        
        baglanti.close()
        
        rafta = toplam - disarida
        
        self.lbl_toplam.configure(text=f"Toplam Kitap Sayısı: {toplam}")
        self.lbl_rafta.configure(text=f"Rafta: {rafta}")
        self.lbl_disarida.configure(text=f"Rafta Değil: {disarida}")

    def merkezde_baslat(self, pencere, genislik, yukseklik, saga_kaydir=False):
        self.update_idletasks()
        pencere.update_idletasks()
        
        ana_x = self.winfo_x()
        ana_y = self.winfo_y()
        ana_g = self.winfo_width()
        ana_yük = self.winfo_height()
        
        if saga_kaydir:
            x = ana_x + int(ana_g * 0.6) - int(genislik / 2)
            y = ana_y + int(ana_yük * 0.5) - int(yukseklik / 2)
        else:
            x = ana_x + int(ana_g / 2) - int(genislik / 2)
            y = ana_y + int(ana_yük / 2) - int(yukseklik / 2)
            
        pencere.geometry(f"{genislik}x{yukseklik}+{x}+{y}")

    def baslik_oku(self):
        baglanti = sqlite3.connect(DB_YOLU)
        imlec = baglanti.cursor()
        imlec.execute("SELECT Deger FROM Ayarlar WHERE Anahtar='Baslik'")
        sonuc = imlec.fetchone()
        baglanti.close()
        return sonuc[0] if sonuc else "Kütüphane Envanteri"

    def baslik_degistir_penceresi(self):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Envanter İsmini Değiştir")
        self.merkezde_baslat(pencere, 400, 200) 
        pencere.attributes('-topmost', True)
        pencere.grab_set()
        ctk.CTkLabel(pencere, text="Yeni Envanter İsmi:", font=("Arial", 14, "bold")).pack(pady=15)
        yeni_baslik_entry = ctk.CTkEntry(pencere, width=300)
        yeni_baslik_entry.insert(0, self.baslik.cget("text")) 
        yeni_baslik_entry.pack(pady=5)
        
        def kaydet():
            yeni_isim = yeni_baslik_entry.get().strip()
            if not yeni_isim:
                messagebox.showwarning("Uyarı", "Başlık alanı boş bırakılamaz!")
                pencere.attributes('-topmost', True)
                return
            baglanti = sqlite3.connect(DB_YOLU)
            baglanti.execute("UPDATE Ayarlar SET Deger=? WHERE Anahtar='Baslik'", (yeni_isim,))
            baglanti.commit()
            baglanti.close()
            self.baslik.configure(text=yeni_isim)
            pencere.destroy()

        ctk.CTkButton(pencere, text="Güncelle", fg_color="#27ae60", hover_color="#2ecc71", command=kaydet).pack(pady=20)

    def kategori_degisti(self, secim):
        self.arama_kutusu.pack_forget()
        self.aralik_frame.pack_forget()
        self.durum_secici.pack_forget()
        self.arama_kutusu.delete(0, tk.END)
        self.aralik_bas.delete(0, tk.END)
        self.aralik_bit.delete(0, tk.END)
        
        if secim in ["Basım Yılı (Aralık)", "Sayfa Sayısı (Aralık)"]:
            self.aralik_bas.configure(placeholder_text="Örn: 1975" if secim == "Basım Yılı (Aralık)" else "Örn: 100")
            self.aralik_bit.configure(placeholder_text="2002" if secim == "Basım Yılı (Aralık)" else "500")
            self.aralik_frame.pack(side="left", padx=5)
        elif secim == "Durum (Rafta/Dışarıda)":
            self.durum_secici.set("Rafta") 
            self.durum_secici.pack(side="left", padx=5)
        elif secim == "Etiket İle Arama":
            self.arama_kutusu.configure(placeholder_text="Örn: #Sosyoloji")
            self.arama_kutusu.insert(0, "#") 
            self.arama_kutusu.pack(side="left", padx=5)
        elif secim == "ID":
            self.arama_kutusu.configure(placeholder_text="4 Haneli ID numarası girin...")
            self.arama_kutusu.pack(side="left", padx=5)
        else:
            self.arama_kutusu.configure(placeholder_text=f"{secim} ile ara...")
            self.arama_kutusu.pack(side="left", padx=5)
        self.arama_yap() 

    def arama_yap(self, *args):
        kat = self.arama_kategorisi.get()
        if kat in ["Basım Yılı (Aralık)", "Sayfa Sayısı (Aralık)"]:
            self.verileri_yukle(kategori=kat, aralik_bas=self.aralik_bas.get().strip(), aralik_bit=self.aralik_bit.get().strip())
        elif kat == "Durum (Rafta/Dışarıda)":
            durum = self.durum_secici.get()
            self.verileri_yukle(kategori=kat, durum_secimi=durum)
        else:
            metin = self.arama_kutusu.get().strip()
            if kat == "Etiket İle Arama" and metin == "#": self.verileri_yukle(arama_metni="", kategori=kat)
            else: self.verileri_yukle(arama_metni=metin, kategori=kat)

    def tablo_olustur(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", borderwidth=0, rowheight=35)
        style.configure("Treeview.Heading", background="#1f538d", foreground="white", font=("Arial", 11, "bold"))
        style.map('Treeview.Heading', background=[('active', '#1f538d')], foreground=[('active', 'white')])
        style.map('Treeview', background=[('selected', '#14375e')])

        # Dikey Kaydırma Çubuğu
        self.scroll_y = ctk.CTkScrollbar(self.tablo_alani, orientation="vertical")
        self.scroll_y.pack(side="right", fill="y")
        
        # YENİ: Yatay Kaydırma Çubuğu
        self.scroll_x = ctk.CTkScrollbar(self.tablo_alani, orientation="horizontal")
        self.scroll_x.pack(side="bottom", fill="x")

        sutunlar = ("ID", "Yazar", "İsim", "Dil", "Yayınevi", "Basım", "Sayfa", "Ölçü", "Kondisyon", "Raf", "Durum", "Bilgi", "İşlem")
        
        # YENİ: xscrollcommand parametresi eklendi
        self.tablo = ttk.Treeview(self.tablo_alani, columns=sutunlar, show="headings", 
                                  yscrollcommand=self.scroll_y.set, 
                                  xscrollcommand=self.scroll_x.set)
        
        self.scroll_y.configure(command=self.tablo.yview)
        self.scroll_x.configure(command=self.tablo.xview) # YENİ: Yatay kaydırma bağlantısı
        
        basliklar_ve_genislik = {"ID": 60, "Yazar": 170, "İsim": 230, "Dil": 90, "Yayınevi": 130, "Basım": 70, "Sayfa": 50, "Ölçü": 60, "Kondisyon": 90, "Raf": 50, "Durum": 100, "Bilgi": 40, "İşlem": 50}
        
        for sutun, genislik in basliklar_ve_genislik.items():
            gosterilen_baslik = "" if sutun == "Bilgi" else ("İşlem" if sutun == "İşlem" else sutun)
            if sutun in self.siralama_durumu.keys() or sutun == "ID":
                if sutun not in self.siralama_durumu: self.siralama_durumu[sutun] = False
                self.tablo.heading(sutun, text=gosterilen_baslik, command=lambda c=sutun: self.sutuna_gore_sirala(c))
            else:
                self.tablo.heading(sutun, text=gosterilen_baslik)
            self.tablo.column(sutun, width=genislik, anchor="center")
            
        self.tablo.pack(fill="both", expand=True)
        self.tablo.tag_configure('disarida', foreground='#ffcccc')
        
        self.tablo.bind("<Double-1>", self.cift_tikla_duzenle)
        self.tablo.bind("<ButtonRelease-1>", self.tablo_tiklama)
        self.tablo.bind("<Button-1>", self.tablo_ozel_tiklama) 
        self.tablo.bind("<Motion>", self.tooltip_kontrol)
        self.tablo.bind("<Leave>", self.tooltip_gizle)
        self.tablo.bind("<Button-2>", self.ort_tus_bas)
        self.tablo.bind("<B2-Motion>", self.ort_tus_surukle)

    def sutun_genisliklerini_otomatik_ayarla(self):
        font = tkfont.Font(family="Arial", size=11)
        max_sinirlar = {"İsim": 420, "Yazar": 230, "Yayınevi": 180, "Dil": 120}
        sabitler = {"Bilgi": 40, "İşlem": 50, "ID": 60, "Durum": 100}
        esnek_sutunlar = ["İsim", "Yazar", "Yayınevi", "Dil"]

        for col in self.tablo["columns"]:
            if col in sabitler:
                self.tablo.column(col, width=sabitler[col], stretch=False)
                continue
                
            baslik_metni = col + " ▼"
            en_genis_piksel = font.measure(baslik_metni) + 30 
            
            for item in self.tablo.get_children(''):
                hucre_verisi = str(self.tablo.set(item, col))
                hucre_piksel = font.measure(hucre_verisi) + 30
                if hucre_piksel > en_genis_piksel:
                    en_genis_piksel = hucre_piksel
            
            if col in max_sinirlar and en_genis_piksel > max_sinirlar[col]:
                en_genis_piksel = max_sinirlar[col]
                
            if col in esnek_sutunlar:
                self.tablo.column(col, width=en_genis_piksel, minwidth=en_genis_piksel, stretch=True)
            else:
                self.tablo.column(col, width=en_genis_piksel, minwidth=en_genis_piksel, stretch=False)

    def verileri_yukle(self, arama_metni="", kategori="Kitap İsmi", aralik_bas="", aralik_bit="", durum_secimi=""):
        for satirlar in self.tablo.get_children():
            self.tablo.delete(satirlar)
            
        self.emanet_bilgileri.clear()
        
        for col in self.siralama_durumu.keys():
            self.tablo.heading(col, text=col)
            
        baglanti = sqlite3.connect(DB_YOLU)
        imlec = baglanti.cursor()
        
        sorgu = "SELECT ID, Isim, Yazar, Dil, Yayinevi, BasimTarihi, SayfaSayisi, Olcu, Kondisyon, RafKonumu, Disarida, EmanetAlan, EmanetTarihi FROM Kitaplar"
        parametreler = []
        
        if kategori == "Basım Yılı (Aralık)":
            if aralik_bas.isdigit() or aralik_bit.isdigit():
                sorgu += " WHERE"
                kosullar = []
                if aralik_bas.isdigit():
                    kosullar.append(" CAST(BasimTarihi AS INTEGER) >= ? ")
                    parametreler.append(int(aralik_bas))
                if aralik_bit.isdigit():
                    kosullar.append(" CAST(BasimTarihi AS INTEGER) <= ? ")
                    parametreler.append(int(aralik_bit))
                sorgu += " AND ".join(kosullar)
        elif kategori == "Sayfa Sayısı (Aralık)":
            if aralik_bas.isdigit() or aralik_bit.isdigit():
                sorgu += " WHERE"
                kosullar = []
                if aralik_bas.isdigit():
                    kosullar.append(" CAST(SayfaSayisi AS INTEGER) >= ? ")
                    parametreler.append(int(aralik_bas))
                if aralik_bit.isdigit():
                    kosullar.append(" CAST(SayfaSayisi AS INTEGER) <= ? ")
                    parametreler.append(int(aralik_bit))
                sorgu += " AND ".join(kosullar)
        elif kategori == "Durum (Rafta/Dışarıda)":
            if durum_secimi == "Rafta":
                sorgu += " WHERE Disarida = 0"
            elif durum_secimi == "Rafta değil":
                sorgu += " WHERE Disarida = 1"
        elif arama_metni:
            sutun_haritasi = {"Yazar": "Yazar", "Kitap İsmi": "Isim", "Dil": "Dil", "Yayınevi": "Yayinevi", "Kondisyon": "Kondisyon", "Raf Konumu": "RafKonumu", "ID": "ID"}
            if kategori in sutun_haritasi:
                if kategori == "ID":
                    try:
                        sorgu += " WHERE ID LIKE ?"
                        parametreler.append(f"%{int(arama_metni)}%")
                    except ValueError:
                        pass 
                else:
                    sorgu += f" WHERE {sutun_haritasi[kategori]} LIKE ?"
                    parametreler.append(f"%{arama_metni}%")
            elif kategori == "Etiket İle Arama":
                sorgu += " WHERE Etiketler LIKE ?"
                arama_terimi = arama_metni if arama_metni.startswith("#") else f"#{arama_metni}"
                parametreler.append(f"%{arama_terimi}%")

        imlec.execute(sorgu, parametreler)
        
        for satir in imlec.fetchall():
            kitap_id = satir[0]
            kitap_id_gorsel = f"{kitap_id:04d}"
            
            if satir[10]:
                durum_metni = "Rafta değil"
                bilgi_simgesi = "❓"
                etiket = ('disarida',)
                self.emanet_bilgileri[kitap_id] = f"{satir[11]}'a teslim edildi.\n{satir[12]}"
            else:
                durum_metni = "Rafta"
                bilgi_simgesi = ""
                etiket = ()
                
            eklenecek_satir = (kitap_id_gorsel, satir[2], satir[1], satir[3], satir[4], satir[5], satir[6], satir[7], satir[8], satir[9], durum_metni, bilgi_simgesi, "...")
            self.tablo.insert("", "end", values=eklenecek_satir, tags=etiket)
            
        baglanti.close()
        self.istatistikleri_guncelle()
        
        if self.aktif_siralama:
            self.sutuna_gore_sirala(self.aktif_siralama, toggle=False)
            
        self.sutun_genisliklerini_otomatik_ayarla()

    def kitap_sil(self, id_listesi, baslik_metni):
        mesaj = f"'{baslik_metni}' sistemden tamamen silinecek. Emin misiniz?" if len(id_listesi) == 1 else f"Seçili {len(id_listesi)} kitap sistemden tamamen silinecek. Emin misiniz?"
        if messagebox.askyesno("Onay", mesaj):
            baglanti = sqlite3.connect(DB_YOLU)
            for k_id in id_listesi:
                baglanti.execute("DELETE FROM Kitaplar WHERE ID=?", (k_id,))
            baglanti.commit()
            baglanti.close()
            self.arama_yap()

    def tooltip_kontrol(self, event):
        region = self.tablo.identify_region(event.x, event.y)
        if region == "cell":
            col = self.tablo.identify_column(event.x)
            satir = self.tablo.identify_row(event.y)
            if col == "#12" and satir:
                tags = self.tablo.item(satir, "tags")
                if 'disarida' in tags:
                    kitap_id = int(self.tablo.item(satir, "values")[0])
                    bilgi = self.emanet_bilgileri.get(kitap_id, "")
                    if bilgi:
                        self.tooltip_goster(event.x_root, event.y_root, bilgi)
                        return
        self.tooltip_gizle()

    def tooltip_goster(self, x, y, metin):
        if self.tooltip_pencere: return
        self.tooltip_pencere = tk.Toplevel(self)
        self.tooltip_pencere.wm_overrideredirect(True)
        self.tooltip_pencere.wm_geometry(f"+{x+15}+{y+10}")
        label = tk.Label(self.tooltip_pencere, text=metin, justify='left', background="#2b2b2b", foreground="#f39c12", relief='solid', borderwidth=1, font=("Arial", 11, "bold"))
        label.pack(ipadx=8, ipady=5)

    def tooltip_gizle(self, event=None):
        if self.tooltip_pencere:
            self.tooltip_pencere.destroy()
            self.tooltip_pencere = None

    def durum_guncelle_penceresi(self, id_listesi, baslik_metni):
        pencere = ctk.CTkToplevel(self)
        pencere.title("Durum Güncelle")
        self.merkezde_baslat(pencere, 350, 350, saga_kaydir=True)
        pencere.attributes('-topmost', True)
        pencere.grab_set()
        
        ctk.CTkLabel(pencere, text=baslik_metni, font=("Arial", 14, "bold"), wraplength=300).pack(pady=15)
        isim_entry = ctk.CTkEntry(pencere, placeholder_text="Emanet Alan Kişi", width=250)
        isim_entry.pack(pady=10)
        
        tarih_frame = ctk.CTkFrame(pencere, fg_color="transparent")
        tarih_frame.pack(pady=10)
        gun_entry = ctk.CTkEntry(tarih_frame, placeholder_text="Gün", width=60)
        gun_entry.pack(side="left", padx=5)
        ay_entry = ctk.CTkEntry(tarih_frame, placeholder_text="Ay", width=60)
        ay_entry.pack(side="left", padx=5)
        yil_entry = ctk.CTkEntry(tarih_frame, placeholder_text="Yıl", width=80)
        yil_entry.pack(side="left", padx=5)

        def emanet_ver():
            alan_kisi = isim_entry.get().strip()
            if not alan_kisi or not gun_entry.get(): return
            tarih = f"{gun_entry.get()}.{ay_entry.get()}.{yil_entry.get()}"
            baglanti = sqlite3.connect(DB_YOLU)
            for k_id in id_listesi:
                baglanti.execute("UPDATE Kitaplar SET Disarida=1, EmanetAlan=?, EmanetTarihi=? WHERE ID=?", (alan_kisi, tarih, k_id))
            baglanti.commit()
            baglanti.close()
            pencere.destroy()
            self.arama_yap()

        def iade_al():
            baglanti = sqlite3.connect(DB_YOLU)
            for k_id in id_listesi:
                baglanti.execute("UPDATE Kitaplar SET Disarida=0, EmanetAlan=NULL, EmanetTarihi=NULL WHERE ID=?", (k_id,))
            baglanti.commit()
            baglanti.close()
            pencere.destroy()
            self.arama_yap()

        ctk.CTkButton(pencere, text="Teslim Edildi İşaretle", fg_color="#d35400", hover_color="#e67e22", command=emanet_ver).pack(pady=(15,5))
        ctk.CTkButton(pencere, text="İade Edildi (Rafa Koy)", fg_color="#27ae60", hover_color="#2ecc71", command=iade_al).pack(pady=5)

    def kitap_ekle_penceresi(self):
        self._kitap_formu_olustur("Yeni Kitap Kaydı", "Sisteme Kaydet", is_yeni=True)

    def kitap_guncelle_penceresi(self, kitap_id):
        baglanti = sqlite3.connect(DB_YOLU)
        imlec = baglanti.cursor()
        imlec.execute("SELECT Isim, Yazar, Dil, Yayinevi, BasimTarihi, SayfaSayisi, Olcu, Kondisyon, RafKonumu FROM Kitaplar WHERE ID=?", (kitap_id,))
        mevcut_veri = imlec.fetchone()
        baglanti.close()
        self._kitap_formu_olustur("Kitap Bilgilerini Güncelle", "Değişiklikleri Kaydet", is_yeni=False, kitap_id=kitap_id, mevcut_veri=mevcut_veri)

    def _kitap_formu_olustur(self, baslik_metni, buton_metni, is_yeni=True, kitap_id=None, mevcut_veri=None):
        pencere = ctk.CTkToplevel(self)
        pencere.title(baslik_metni)
        
        self.merkezde_baslat(pencere, 450, 780)  
        pencere.attributes('-topmost', True)
        pencere.grab_set()
        
        ctk.CTkLabel(pencere, text=baslik_metni, font=("Arial", 20, "bold")).pack(pady=(15, 10))

        if is_yeni:
            isbn_frame = ctk.CTkFrame(pencere, fg_color="transparent")
            isbn_frame.pack(pady=(0, 5), fill="x", padx=30)
            
            ctk.CTkLabel(isbn_frame, text="Barkod/ISBN Okut (Otomatik Doldur):", font=("Arial", 11, "bold"), text_color="#f39c12").pack(anchor="w")
            
            entry_btn_frame = ctk.CTkFrame(isbn_frame, fg_color="transparent")
            entry_btn_frame.pack(fill="x", pady=2)
            
            isbn_entry = ctk.CTkEntry(entry_btn_frame, placeholder_text="Okutun veya yazın...", height=32)
            isbn_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
            
            durum_lbl = ctk.CTkLabel(pencere, text="", font=("Arial", 11, "italic"))
            durum_lbl.pack(pady=(0, 5))

        def satir_olustur(parent, placeholder, key, mevcut_val=""):
            frame = ctk.CTkFrame(parent, fg_color="transparent")
            frame.pack(pady=4) 

            entry = ctk.CTkEntry(frame, placeholder_text=placeholder, width=310)
            entry.pack(side="left", padx=(0, 5))

            if mevcut_val:
                entry.insert(0, mevcut_val)
            elif is_yeni and self.sabit_kilitler[key]:
                entry.insert(0, self.sabit_degerler[key])

            if is_yeni:
                kilitli_mi = self.sabit_kilitler[key]
                renk = "#27ae60" if kilitli_mi else "#1f538d"
                hover = "#2ecc71" if kilitli_mi else "#14375e"
                ikon = "🔒" if kilitli_mi else "🔓"
                
                btn = ctk.CTkButton(frame, text=ikon, width=35, font=("Arial", 14), fg_color=renk, hover_color=hover)
                btn.pack(side="left")

                def toggle_kilit(k=key, b=btn, e=entry):
                    self.sabit_kilitler[k] = not self.sabit_kilitler[k]
                    if self.sabit_kilitler[k]:
                        b.configure(text="🔒", fg_color="#27ae60", hover_color="#2ecc71")
                        self.sabit_degerler[k] = e.get().strip() 
                    else:
                        b.configure(text="🔓", fg_color="#1f538d", hover_color="#14375e")
                        self.sabit_degerler[k] = ""

                btn.configure(command=toggle_kilit)
            else:
                entry.configure(width=350)

            return entry

        yazar_entry = satir_olustur(pencere, "Yazarı", "yazar", mevcut_veri[1] if mevcut_veri else "")
        isim_entry = satir_olustur(pencere, "Kitap İsmi", "isim", mevcut_veri[0] if mevcut_veri else "")
        dil_entry = satir_olustur(pencere, "Dili (Örn: Türkçe)", "dil", mevcut_veri[2] if mevcut_veri else "")
        yayinevi_entry = satir_olustur(pencere, "Yayınevi", "yayinevi", mevcut_veri[3] if mevcut_veri else "")
        basim_entry = satir_olustur(pencere, "Basım Tarihi/Yılı", "basim", mevcut_veri[4] if mevcut_veri else "")
        sayfa_entry = satir_olustur(pencere, "Sayfa Sayısı", "sayfa", mevcut_veri[5] if mevcut_veri else "")
        olcu_entry = satir_olustur(pencere, "Ölçü (Örn: 14x21)", "olcu", mevcut_veri[6] if mevcut_veri else "")
        kondisyon_entry = satir_olustur(pencere, "Kondisyon (Örn: İyi, Yıpranmış)", "kondisyon", mevcut_veri[7] if mevcut_veri else "")
        raf_entry = satir_olustur(pencere, "Raf Konumu (Örn: 8)", "raf", mevcut_veri[8] if mevcut_veri else "")

        self.olcu_ref = olcu_entry
        self.kondisyon_ref = kondisyon_entry
        self.raf_ref = raf_entry

        if is_yeni:
            def internetten_ara(event=None):
                ham_kod = isbn_entry.get().strip()
                if not ham_kod: return
                
                kod = ham_kod.replace("-", "").replace(" ", "")
                durum_lbl.configure(text="🔍 İnternette sorgulanıyor...", text_color="#f1c40f")
                pencere.update_idletasks()
                
                bulunan_veri = {"Isim": "", "Yazar": "", "Dil": "", "Yayinevi": "", "Basim": "", "Sayfa": ""}
                hata_mesaji = ""
                bulundu_mu = False

                try:
                    try:
                        ctx = ssl._create_unverified_context()
                        url = f"https://www.googleapis.com/books/v1/volumes?q=isbn:{kod}"
                        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=4, context=ctx) as response:
                            data = json.loads(response.read().decode('utf-8'))
                        
                        if "items" in data:
                            info = data["items"][0]["volumeInfo"]
                            bulunan_veri["Isim"] = info.get("title", "")
                            bulunan_veri["Yazar"] = ", ".join(info.get("authors", []))
                            bulunan_veri["Yayinevi"] = info.get("publisher", "")
                            
                            dil_kodu = info.get("language", "").lower()
                            if dil_kodu == "tr": bulunan_veri["Dil"] = "Türkçe"
                            elif dil_kodu == "en": bulunan_veri["Dil"] = "İngilizce"
                            else: bulunan_veri["Dil"] = dil_kodu.upper()
                            
                            tarih = info.get("publishedDate", "")
                            bulunan_veri["Basim"] = tarih[:4] if len(tarih) >= 4 else tarih
                            bulunan_veri["Sayfa"] = str(info.get("pageCount", ""))
                            bulundu_mu = True
                    except Exception as e:
                        if "429" not in str(e): hata_mesaji = str(e)
                            
                    if not bulundu_mu:
                        try:
                            ctx = ssl._create_unverified_context()
                            url = f"https://openlibrary.org/api/books?bibkeys=ISBN:{kod}&format=json&jscmd=data"
                            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                            with urllib.request.urlopen(req, timeout=4, context=ctx) as response:
                                dataOL = json.loads(response.read().decode('utf-8'))
                            
                            key = f"ISBN:{kod}"
                            if key in dataOL:
                                infoOL = dataOL[key]
                                bulunan_veri["Isim"] = infoOL.get("title", "")
                                if "authors" in infoOL:
                                    bulunan_veri["Yazar"] = ", ".join([y.get("name", "") for y in infoOL["authors"]])
                                if "publishers" in infoOL:
                                    bulunan_veri["Yayinevi"] = ", ".join([y.get("name", "") for y in infoOL["publishers"]])
                                tarihOL = infoOL.get("publish_date", "")
                                bulunan_veri["Basim"] = str(tarihOL)[-4:] if len(str(tarihOL)) >= 4 else str(tarihOL)
                                bulunan_veri["Sayfa"] = str(infoOL.get("number_of_pages", ""))
                                bulundu_mu = True
                        except Exception as e:
                            if not hata_mesaji: hata_mesaji = str(e)

                    if bulundu_mu:
                        if not self.sabit_kilitler["isim"] and bulunan_veri["Isim"]: 
                            isim_entry.delete(0, tk.END); isim_entry.insert(0, bulunan_veri["Isim"])
                        if not self.sabit_kilitler["yazar"] and bulunan_veri["Yazar"]: 
                            yazar_entry.delete(0, tk.END); yazar_entry.insert(0, bulunan_veri["Yazar"])
                        if not self.sabit_kilitler["dil"] and bulunan_veri["Dil"]: 
                            dil_entry.delete(0, tk.END); dil_entry.insert(0, bulunan_veri["Dil"])
                        if not self.sabit_kilitler["yayinevi"] and bulunan_veri["Yayinevi"]: 
                            yayinevi_entry.delete(0, tk.END); yayinevi_entry.insert(0, bulunan_veri["Yayinevi"])
                        if not self.sabit_kilitler["basim"] and bulunan_veri["Basim"]: 
                            basim_entry.delete(0, tk.END); basim_entry.insert(0, bulunan_veri["Basim"])
                        if not self.sabit_kilitler["sayfa"] and bulunan_veri["Sayfa"]: 
                            sayfa_entry.delete(0, tk.END); sayfa_entry.insert(0, bulunan_veri["Sayfa"])
                        
                        durum_lbl.configure(text="✅ Künye dolduruldu! Düzenleyip kaydedebilirsiniz.", text_color="#2ecc71")
                    else:
                        if "429" in hata_mesaji:
                            durum_lbl.configure(text="❌ Google kotası doldu! Az sonra tekrar deneyin.", text_color="#e74c3c")
                        else:
                            durum_lbl.configure(text="❌ Barkod bulunamadı! Verileri elle girin.", text_color="#f1c40f")
                            
                except Exception as main_e:
                    durum_lbl.configure(text=f"❌ Hata: {str(main_e)[:40]}", text_color="#e74c3c")

            ara_btn = ctk.CTkButton(entry_btn_frame, text="🔍 Bul", font=("Arial", 12, "bold"), width=60, height=32, command=internetten_ara)
            ara_btn.pack(side="right")
            
            isbn_entry.bind("<Return>", internetten_ara)
            pencere.after(100, isbn_entry.focus)

        b_frame = ctk.CTkFrame(pencere, fg_color="transparent")
        b_frame.pack(pady=20)

        def veriyi_isle(event=None):
            isim = isim_entry.get().strip()
            yazar = yazar_entry.get().strip()
            if not isim or not yazar:
                messagebox.showwarning("Uyarı", "Kitap İsmi ve Yazar alanları zorunludur!")
                pencere.attributes('-topmost', True)
                return
                
            baglanti = sqlite3.connect(DB_YOLU)
            if not is_yeni: 
                baglanti.execute('''
                    UPDATE Kitaplar SET Isim=?, Yazar=?, Dil=?, Yayinevi=?, BasimTarihi=?, SayfaSayisi=?, Olcu=?, Kondisyon=?, RafKonumu=?
                    WHERE ID=?
                ''', (isim, yazar, dil_entry.get().strip(), yayinevi_entry.get().strip(), basim_entry.get().strip(), sayfa_entry.get().strip(), olcu_entry.get().strip(), kondisyon_entry.get().strip(), raf_entry.get().strip(), kitap_id))
                baglanti.commit()
                baglanti.close()
                pencere.destroy()
                self.arama_yap()
            else: 
                if self.sabit_kilitler["yazar"]: self.sabit_degerler["yazar"] = yazar
                if self.sabit_kilitler["isim"]: self.sabit_degerler["isim"] = isim
                if self.sabit_kilitler["dil"]: self.sabit_degerler["dil"] = dil_entry.get().strip()
                if self.sabit_kilitler["yayinevi"]: self.sabit_degerler["yayinevi"] = yayinevi_entry.get().strip()
                if self.sabit_kilitler["basim"]: self.sabit_degerler["basim"] = basim_entry.get().strip()
                if self.sabit_kilitler["sayfa"]: self.sabit_degerler["sayfa"] = sayfa_entry.get().strip()
                if self.sabit_kilitler["olcu"]: self.sabit_degerler["olcu"] = olcu_entry.get().strip()
                if self.sabit_kilitler["kondisyon"]: self.sabit_degerler["kondisyon"] = kondisyon_entry.get().strip()
                if self.sabit_kilitler["raf"]: self.sabit_degerler["raf"] = raf_entry.get().strip()

                imlec = baglanti.cursor()
                while True:
                    yeni_id = random.randint(1, 9999)
                    imlec.execute("SELECT ID FROM Kitaplar WHERE ID=?", (yeni_id,))
                    if not imlec.fetchone(): 
                        break
                        
                baglanti.execute('''
                    INSERT INTO Kitaplar (ID, Isim, Yazar, Dil, Yayinevi, BasimTarihi, SayfaSayisi, Olcu, Kondisyon, RafKonumu)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (yeni_id, isim, yazar, dil_entry.get().strip(), yayinevi_entry.get().strip(), basim_entry.get().strip(), sayfa_entry.get().strip(), olcu_entry.get().strip(), kondisyon_entry.get().strip(), raf_entry.get().strip()))
                
                baglanti.commit()
                baglanti.close()
                self.arama_yap()
                
                if not self.sabit_kilitler["isim"]: isim_entry.delete(0, tk.END)
                if not self.sabit_kilitler["yazar"]: yazar_entry.delete(0, tk.END)
                if not self.sabit_kilitler["dil"]: dil_entry.delete(0, tk.END)
                if not self.sabit_kilitler["yayinevi"]: yayinevi_entry.delete(0, tk.END)
                if not self.sabit_kilitler["basim"]: basim_entry.delete(0, tk.END)
                if not self.sabit_kilitler["sayfa"]: sayfa_entry.delete(0, tk.END)
                if not self.sabit_kilitler["olcu"]: olcu_entry.delete(0, tk.END)
                if not self.sabit_kilitler["kondisyon"]: kondisyon_entry.delete(0, tk.END)
                if not self.sabit_kilitler["raf"]: raf_entry.delete(0, tk.END)

                isbn_entry.delete(0, tk.END)
                isbn_entry.focus()
                durum_lbl.configure(text=f"✅ '{isim[:15]}...' Eklendi! Sıradakine geçin.", text_color="#2ecc71")

        ctk.CTkButton(b_frame, text=buton_metni, font=("Arial", 14, "bold"), fg_color="#27ae60", hover_color="#2ecc71", text_color="white", command=veriyi_isle).pack(side="left", padx=5)
            
        pencere.bind("<Control-s>", veriyi_isle)
        pencere.bind("<Control-S>", veriyi_isle)

if __name__ == "__main__":
    veritabani_hazirla()
    uygulama = KutuphaneUygulamasi()
    uygulama.mainloop()