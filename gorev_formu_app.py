import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import json
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import webbrowser
import tempfile

class GorevFormuApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Delta Proje - Görev Formu")
        self.root.geometry("800x600")
        self.root.configure(bg='#f5f5f5')
        
        # Form verileri
        self.form_data = {}
        self.current_step = 0
        self.form_no = self.get_next_form_no()
        
        # Ana frame
        self.main_frame = tk.Frame(root, bg='white', padx=30, pady=30)
        self.main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # İlk adımı göster
        self.show_step()
    
    def get_next_form_no(self):
        """Form numarasını al veya oluştur"""
        config_file = 'form_config.json'
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                last_no = config.get('last_form_no', 0)
        else:
            last_no = 0
        
        next_no = last_no + 1
        
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump({'last_form_no': next_no}, f)
        
        return str(next_no).zfill(5)
    
    def clear_frame(self):
        """Frame'i temizle"""
        for widget in self.main_frame.winfo_children():
            widget.destroy()
    
    def show_step(self):
        """Mevcut adımı göster"""
        self.clear_frame()
        
        if self.current_step == 0:
            self.step_0_form_bilgileri()
        elif self.current_step == 1:
            self.step_1_gorevli_personel()
        elif self.current_step == 2:
            self.step_2_avans_taseron()
        elif self.current_step == 3:
            self.step_3_gorev_tanimi()
        elif self.current_step == 4:
            self.step_4_gorev_yeri()
        elif self.current_step == 5:
            self.step_5_saat_bilgileri()
        elif self.current_step == 6:
            self.step_6_arac_bilgisi()
        elif self.current_step == 7:
            self.step_7_hazirlayan()
        elif self.current_step == 8:
            self.show_summary()
    
    def step_0_form_bilgileri(self):
        """Adım 0: Form Bilgileri"""
        tk.Label(self.main_frame, text="GÖREV FORMU", font=('Arial', 24, 'bold'), 
                bg='white', fg='#d32f2f').pack(pady=20)
        
        info_frame = tk.Frame(self.main_frame, bg='#fff9c4', relief='solid', borderwidth=2)
        info_frame.pack(pady=20, padx=50, fill='x')
        
        tk.Label(info_frame, text=f"FORM NO: {self.form_no}", font=('Arial', 16, 'bold'),
                bg='#fff9c4', fg='#d32f2f').pack(pady=10)
        
        tk.Label(info_frame, text=f"TARİH: {datetime.now().strftime('%d.%m.%Y')}", 
                font=('Arial', 12), bg='#fff9c4').pack(pady=5)
        
        tk.Label(info_frame, text="DOK.NO: F-001", font=('Arial', 12), 
                bg='#fff9c4').pack(pady=5)
        
        tk.Label(info_frame, text="REV.NO/TRH: 00 / 06.05.24", font=('Arial', 12),
                bg='#fff9c4').pack(pady=5)
        
        self.form_data['form_no'] = self.form_no
        self.form_data['tarih'] = datetime.now().strftime('%d.%m.%Y')
        
        self.add_navigation_buttons()
    
    def step_1_gorevli_personel(self):
        """Adım 1: Görevli Personel"""
        tk.Label(self.main_frame, text="GÖREVLİ PERSONEL", font=('Arial', 18, 'bold'),
                bg='white').pack(pady=20)
        
        tk.Label(self.main_frame, text="En fazla 5 personel seçebilirsiniz", 
                font=('Arial', 10), bg='white', fg='gray').pack()
        
        personel_listesi = ["Personel 1", "Personel 2", "Personel 3", "Personel 4", 
                           "Personel 5", "Personel 6", "Personel 7"]
        
        self.personel_vars = []
        personel_frame = tk.Frame(self.main_frame, bg='white')
        personel_frame.pack(pady=20)
        
        for i in range(5):
            frame = tk.Frame(personel_frame, bg='#ffeb3b', pady=5, padx=10)
            frame.pack(fill='x', pady=5)
            
            tk.Label(frame, text=f"Personel {i+1}:", font=('Arial', 11),
                    bg='#ffeb3b').pack(side='left', padx=5)
            
            var = tk.StringVar()
            combo = ttk.Combobox(frame, textvariable=var, values=personel_listesi,
                                width=30, state='readonly')
            combo.pack(side='left', padx=5)
            self.personel_vars.append(var)
        
        self.add_navigation_buttons()
    
    def step_2_avans_taseron(self):
        """Adım 2: Avans Tutarı ve Taşeron Şirket"""
        tk.Label(self.main_frame, text="AVANS TUTARI ve TAŞERON ŞİRKET", 
                font=('Arial', 18, 'bold'), bg='white').pack(pady=20)
        
        # Avans tutarı
        avans_frame = tk.Frame(self.main_frame, bg='#4dd0e1', pady=15, padx=20)
        avans_frame.pack(fill='x', pady=10, padx=50)
        
        tk.Label(avans_frame, text="Avans Tutarı:", font=('Arial', 12, 'bold'),
                bg='#4dd0e1').pack()
        self.avans_entry = tk.Entry(avans_frame, font=('Arial', 12), width=30)
        self.avans_entry.pack(pady=5)
        
        # Taşeron şirket
        taseron_frame = tk.Frame(self.main_frame, bg='#ffeb3b', pady=15, padx=20)
        taseron_frame.pack(fill='x', pady=10, padx=50)
        
        tk.Label(taseron_frame, text="Taşeron Şirket:", font=('Arial', 12, 'bold'),
                bg='#ffeb3b').pack()
        
        sirket_listesi = ["Şirket 1", "Şirket 2", "Şirket 3", "Şirket 4", "Şirket 5"]
        self.taseron_var = tk.StringVar()
        taseron_combo = ttk.Combobox(taseron_frame, textvariable=self.taseron_var,
                                     values=sirket_listesi, width=30)
        taseron_combo.pack(pady=5)
        
        self.add_navigation_buttons()
    
    def step_3_gorev_tanimi(self):
        """Adım 3: Görevin Tanımı"""
        tk.Label(self.main_frame, text="GÖREVİN TANIMI", font=('Arial', 18, 'bold'),
                bg='white').pack(pady=20)
        
        text_frame = tk.Frame(self.main_frame, bg='#ffeb3b', pady=15, padx=20)
        text_frame.pack(fill='both', expand=True, pady=10, padx=50)
        
        self.gorev_tanimi_text = tk.Text(text_frame, font=('Arial', 11), 
                                         height=15, wrap='word')
        self.gorev_tanimi_text.pack(fill='both', expand=True, pady=5)
        
        self.add_navigation_buttons()
    
    def step_4_gorev_yeri(self):
        """Adım 4: Görev Yeri"""
        tk.Label(self.main_frame, text="GÖREV YERİ", font=('Arial', 18, 'bold'),
                bg='white').pack(pady=20)
        
        text_frame = tk.Frame(self.main_frame, bg='#ffeb3b', pady=15, padx=20)
        text_frame.pack(fill='both', expand=True, pady=10, padx=50)
        
        self.gorev_yeri_text = tk.Text(text_frame, font=('Arial', 11),
                                       height=15, wrap='word')
        self.gorev_yeri_text.pack(fill='both', expand=True, pady=5)
        
        self.add_navigation_buttons()
    
    def step_5_saat_bilgileri(self):
        """Adım 5: Saat Bilgileri"""
        tk.Label(self.main_frame, text="SAAT - İŞÇİLİKLERİ", font=('Arial', 18, 'bold'),
                bg='white').pack(pady=20)
        
        # Yola çıkış
        frame1 = tk.Frame(self.main_frame, bg='#ffeb3b', pady=10, padx=15)
        frame1.pack(fill='x', pady=5, padx=50)
        tk.Label(frame1, text="Yola Çıkış Tarihi ve Saati:", font=('Arial', 11, 'bold'),
                bg='#ffeb3b').pack()
        self.yola_cikis_tarih = tk.Entry(frame1, font=('Arial', 10), width=15)
        self.yola_cikis_tarih.pack(side='left', padx=5)
        self.yola_cikis_tarih.insert(0, datetime.now().strftime('%d.%m.%Y'))
        self.yola_cikis_saat = tk.Entry(frame1, font=('Arial', 10), width=10)
        self.yola_cikis_saat.pack(side='left', padx=5)
        self.yola_cikis_saat.insert(0, "08:00")
        
        # Dönüş
        frame2 = tk.Frame(self.main_frame, bg='#ffeb3b', pady=10, padx=15)
        frame2.pack(fill='x', pady=5, padx=50)
        tk.Label(frame2, text="Dönüş Tarihi ve Saati:", font=('Arial', 11, 'bold'),
                bg='#ffeb3b').pack()
        self.donus_tarih = tk.Entry(frame2, font=('Arial', 10), width=15)
        self.donus_tarih.pack(side='left', padx=5)
        self.donus_tarih.insert(0, datetime.now().strftime('%d.%m.%Y'))
        self.donus_saat = tk.Entry(frame2, font=('Arial', 10), width=10)
        self.donus_saat.pack(side='left', padx=5)
        self.donus_saat.insert(0, "17:00")
        
        # Çalışma başlangıç
        frame3 = tk.Frame(self.main_frame, bg='#ffeb3b', pady=10, padx=15)
        frame3.pack(fill='x', pady=5, padx=50)
        tk.Label(frame3, text="Çalışma Başlangıç Tarihi ve Saati:", 
                font=('Arial', 11, 'bold'), bg='#ffeb3b').pack()
        self.calisma_baslangic_tarih = tk.Entry(frame3, font=('Arial', 10), width=15)
        self.calisma_baslangic_tarih.pack(side='left', padx=5)
        self.calisma_baslangic_tarih.insert(0, datetime.now().strftime('%d.%m.%Y'))
        self.calisma_baslangic_saat = tk.Entry(frame3, font=('Arial', 10), width=10)
        self.calisma_baslangic_saat.pack(side='left', padx=5)
        self.calisma_baslangic_saat.insert(0, "09:00")
        
        # Çalışma bitiş
        frame4 = tk.Frame(self.main_frame, bg='#ffeb3b', pady=10, padx=15)
        frame4.pack(fill='x', pady=5, padx=50)
        tk.Label(frame4, text="Çalışma Bitiş Tarihi ve Saati:", 
                font=('Arial', 11, 'bold'), bg='#ffeb3b').pack()
        self.calisma_bitis_tarih = tk.Entry(frame4, font=('Arial', 10), width=15)
        self.calisma_bitis_tarih.pack(side='left', padx=5)
        self.calisma_bitis_tarih.insert(0, datetime.now().strftime('%d.%m.%Y'))
        self.calisma_bitis_saat = tk.Entry(frame4, font=('Arial', 10), width=10)
        self.calisma_bitis_saat.pack(side='left', padx=5)
        self.calisma_bitis_saat.insert(0, "16:00")
        
        # Mola süresi
        frame5 = tk.Frame(self.main_frame, bg='#ffeb3b', pady=10, padx=15)
        frame5.pack(fill='x', pady=5, padx=50)
        tk.Label(frame5, text="Toplam Mola Süresi (dakika):", 
                font=('Arial', 11, 'bold'), bg='#ffeb3b').pack()
        self.mola_suresi = tk.Entry(frame5, font=('Arial', 10), width=10)
        self.mola_suresi.pack(padx=5)
        self.mola_suresi.insert(0, "60")
        
        self.add_navigation_buttons()
    
    def step_6_arac_bilgisi(self):
        """Adım 6: Araç Plaka No"""
        tk.Label(self.main_frame, text="ARAÇ PLAKA NO", font=('Arial', 18, 'bold'),
                bg='white').pack(pady=20)
        
        arac_frame = tk.Frame(self.main_frame, bg='#fff9c4', pady=15, padx=20)
        arac_frame.pack(fill='x', pady=10, padx=50)
        
        plaka_listesi = ["34 ABC 123", "06 XYZ 456", "41 DEF 789", "35 GHI 321"]
        self.arac_var = tk.StringVar()
        arac_combo = ttk.Combobox(arac_frame, textvariable=self.arac_var,
                                 values=plaka_listesi, width=30, state='readonly')
        arac_combo.pack(pady=10)
        
        self.add_navigation_buttons()
    
    def step_7_hazirlayan(self):
        """Adım 7: Hazırlayan - Görevlendiren"""
        tk.Label(self.main_frame, text="HAZIRLAYAN - GÖREVLENDİREN", 
                font=('Arial', 18, 'bold'), bg='white').pack(pady=20)
        
        hazir_frame = tk.Frame(self.main_frame, bg='#fff9c4', pady=15, padx=20)
        hazir_frame.pack(fill='x', pady=10, padx=50)
        
        tk.Label(hazir_frame, text="Adı - Soyadı:", font=('Arial', 12, 'bold'),
                bg='#fff9c4').pack()
        
        personel_listesi = ["Personel 1", "Personel 2", "Personel 3", "Personel 4"]
        self.hazirlayan_var = tk.StringVar()
        hazir_combo = ttk.Combobox(hazir_frame, textvariable=self.hazirlayan_var,
                                   values=personel_listesi, width=30, state='readonly')
        hazir_combo.pack(pady=10)
        
        self.add_navigation_buttons()
    
    def show_summary(self):
        """Özet ekranı - HTML şablonunda göster"""
        # HTML dosyası oluştur
        html_content = self.generate_html_summary()
        
        # Geçici dosya oluştur
        temp_file = os.path.join(tempfile.gettempdir(), f'gorev_formu_ozet_{self.form_no}.html')
        with open(temp_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # Tarayıcıda aç
        webbrowser.open('file://' + temp_file)
        
        # Tkinter penceresinde sadece kaydet butonu göster
        tk.Label(self.main_frame, text="Form Özeti Tarayıcıda Açıldı", 
                font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=30)
        
        tk.Label(self.main_frame, text="Lütfen tarayıcıdaki formu kontrol edin.\nEğer her şey doğruysa aşağıdaki KAYDET butonuna basın.", 
                font=('Arial', 12), bg='white', justify='center').pack(pady=20)
        
        # Kaydet butonu
        btn_frame = tk.Frame(self.main_frame, bg='white')
        btn_frame.pack(pady=30)
        
        tk.Button(btn_frame, text="◀ Geri", font=('Arial', 12),
                 command=self.prev_step, bg='#e0e0e0', width=12).pack(side='left', padx=10)
        
        tk.Button(btn_frame, text="💾 KAYDET", font=('Arial', 14, 'bold'),
                 command=self.save_to_excel, bg='#4caf50', fg='white',
                 width=15, height=2).pack(side='left', padx=10)
    
    def generate_html_summary(self):
        """HTML özet sayfası oluştur"""
        personel_html = ""
        for i, personel in enumerate(self.form_data.get('personel_listesi', []), 1):
            if personel:
                personel_html += f'<option selected>{personel}</option>'
        
        html = f"""<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Görev Formu - Önizleme</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: Arial, sans-serif;
            padding: 20px;
            background: #f5f5f5;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 20px;
        }}
        
        .header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 30px;
            border-bottom: 2px solid #333;
            padding-bottom: 15px;
        }}
        
        .logo-section {{
            display: flex;
            align-items: center;
            gap: 15px;
        }}
        
        .logo {{
            width: 60px;
            height: 60px;
            background: #d32f2f;
            clip-path: polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%);
        }}
        
        .logo-text .delta {{
            font-size: 24px;
            color: #d32f2f;
            font-weight: bold;
        }}
        
        .logo-text .subtitle {{
            font-size: 11px;
            color: #666;
        }}
        
        .form-title {{
            font-size: 32px;
            font-weight: bold;
            letter-spacing: 2px;
        }}
        
        .form-info {{
            background: #fff9c4;
            padding: 10px;
            border: 1px solid #333;
        }}
        
        .form-info-row {{
            display: grid;
            grid-template-columns: 100px 1fr;
            border-bottom: 1px solid #333;
            padding: 5px 0;
        }}
        
        .form-info-row:last-child {{
            border-bottom: none;
        }}
        
        .form-info-label {{
            font-weight: bold;
            padding: 5px;
            border-right: 1px solid #333;
        }}
        
        .form-info-value {{
            padding: 5px 10px;
            display: flex;
            align-items: center;
        }}
        
        #formNo {{
            font-size: 20px;
            font-weight: bold;
            color: #d32f2f;
        }}
        
        .main-grid {{
            display: grid;
            grid-template-columns: 1fr 300px;
            gap: 0;
            border: 2px solid #333;
            margin-bottom: 20px;
        }}
        
        .left-section {{
            border-right: 2px solid #333;
        }}
        
        .row-section {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            border-bottom: 2px solid #333;
        }}
        
        .field-box {{
            border-right: 2px solid #333;
            padding: 15px;
            background: #ffeb3b;
        }}
        
        .field-box:last-child {{
            border-right: none;
        }}
        
        .field-box.cyan {{
            background: #4dd0e1;
        }}
        
        .field-label {{
            font-weight: bold;
            font-size: 12px;
            margin-bottom: 8px;
            display: block;
        }}
        
        .field-value {{
            padding: 8px;
            background: white;
            border: 1px solid #333;
            min-height: 30px;
        }}
        
        .task-definition {{
            padding: 15px;
            background: #ffeb3b;
            min-height: 200px;
        }}
        
        .right-section {{
            background: #ffeb3b;
            padding: 15px;
        }}
        
        .time-tracking {{
            margin: 30px 0;
            border: 2px solid #333;
        }}
        
        .time-tracking-title {{
            background: #fff9c4;
            padding: 10px;
            font-weight: bold;
            border-bottom: 2px solid #333;
        }}
        
        .time-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            border-bottom: 2px solid #333;
        }}
        
        .time-cell-header {{
            border-right: 2px solid #333;
            padding: 10px;
            text-align: center;
            background: #fff9c4;
            font-size: 11px;
            font-weight: bold;
        }}
        
        .time-cell-header:last-child {{
            border-right: none;
        }}
        
        .time-row {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            border-bottom: 2px solid #333;
        }}
        
        .time-input-cell {{
            border-right: 2px solid #333;
            padding: 15px;
            background: #ffeb3b;
            text-align: center;
        }}
        
        .time-input-cell:last-child {{
            border-right: none;
        }}
        
        .vehicle-section {{
            padding: 15px;
            background: #fff9c4;
        }}
        
        .signature-section {{
            margin: 30px 0;
        }}
        
        .signature-row {{
            display: flex;
            gap: 0;
        }}
        
        .signature-box {{
            flex: 1;
            border: 2px solid #333;
            border-right: none;
            padding: 15px;
            min-height: 100px;
            background: #fff9c4;
        }}
        
        .signature-box:last-child {{
            border-right: 2px solid #333;
            background: white;
        }}
        
        .signature-label {{
            font-weight: bold;
            margin-bottom: 10px;
        }}
        
        select {{
            width: 100%;
            padding: 5px;
            border: none;
            background: white;
            font-size: 14px;
        }}
        
        .text-content {{
            padding: 10px;
            background: white;
            border: 1px solid #333;
            min-height: 100px;
            white-space: pre-wrap;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="logo-section">
                <div class="logo"></div>
                <div class="logo-text">
                    <div class="delta">delta proje</div>
                    <div class="subtitle">hidrolik & pnömatik</div>
                </div>
            </div>
            <div class="form-title">GÖREV FORMU</div>
            <div class="form-info">
                <div class="form-info-row">
                    <div class="form-info-label">FORM NO</div>
                    <div class="form-info-value">
                        <span id="formNo">{self.form_data.get('form_no', '')}</span>
                    </div>
                </div>
                <div class="form-info-row">
                    <div class="form-info-label">TARİH</div>
                    <div class="form-info-value">{self.form_data.get('tarih', '')}</div>
                </div>
                <div class="form-info-row">
                    <div class="form-info-label">DOK.NO</div>
                    <div class="form-info-value">F-001</div>
                </div>
                <div class="form-info-row">
                    <div class="form-info-label">REV.NO/TRH</div>
                    <div class="form-info-value">00 / 06.05.24</div>
                </div>
            </div>
        </div>

        <div class="main-grid">
            <div class="left-section">
                <div class="row-section">
                    <div class="field-box">
                        <label class="field-label">GÖREVLİ PERSONEL</label>
                        <select size="5" style="height: auto;">
                            {personel_html}
                        </select>
                    </div>
                    <div class="field-box cyan">
                        <label class="field-label">AVANSI TUTARI</label>
                        <div class="field-value">{self.form_data.get('avans_tutari', '')}</div>
                    </div>
                    <div class="field-box">
                        <label class="field-label">TAŞERON ŞİRKET</label>
                        <div class="field-value">{self.form_data.get('taseron_sirket', '')}</div>
                    </div>
                </div>
                <div class="task-definition">
                    <label class="field-label">GÖREVİN TANIMI</label>
                    <div class="text-content">{self.form_data.get('gorev_tanimi', '')}</div>
                </div>
            </div>
            <div class="right-section">
                <label class="field-label">G.YERİ</label>
                <div class="text-content" style="min-height: 280px;">{self.form_data.get('gorev_yeri', '')}</div>
            </div>
        </div>

        <div class="time-tracking">
            <div class="time-tracking-title">SAAT - İŞÇİLİKLERİ</div>
            <div class="time-grid">
                <div class="time-cell-header">YOLA ÇIKIŞ<br>TARİH ve SAAT</div>
                <div class="time-cell-header">DÖNÜŞ<br>TARİH ve SAAT</div>
                <div class="time-cell-header">ÇALIŞMA BAŞLANGIÇ<br>TARİH ve SAAT</div>
                <div class="time-cell-header">ÇALIŞMA BİTİŞ<br>TARİH ve SAAT</div>
                <div class="time-cell-header">TOPLAM MOLA<br>SÜRESİ (dakika)</div>
            </div>
            <div class="time-row">
                <div class="time-input-cell">
                    <div>{self.form_data.get('yola_cikis_tarih', '')}</div>
                    <div style="font-weight: bold; font-size: 16px; margin-top: 5px;">{self.form_data.get('yola_cikis_saat', '')}</div>
                </div>
                <div class="time-input-cell">
                    <div>{self.form_data.get('donus_tarih', '')}</div>
                    <div style="font-weight: bold; font-size: 16px; margin-top: 5px;">{self.form_data.get('donus_saat', '')}</div>
                </div>
                <div class="time-input-cell">
                    <div>{self.form_data.get('calisma_baslangic_tarih', '')}</div>
                    <div style="font-weight: bold; font-size: 16px; margin-top: 5px;">{self.form_data.get('calisma_baslangic_saat', '')}</div>
                </div>
                <div class="time-input-cell">
                    <div>{self.form_data.get('calisma_bitis_tarih', '')}</div>
                    <div style="font-weight: bold; font-size: 16px; margin-top: 5px;">{self.form_data.get('calisma_bitis_saat', '')}</div>
                </div>
                <div class="time-input-cell">
                    <div style="font-weight: bold; font-size: 20px;">{self.form_data.get('mola_suresi', '')} dk</div>
                </div>
            </div>
            <div class="vehicle-section">
                <label class="field-label">ARAÇ PLAKA NO</label>
                <div class="field-value">{self.form_data.get('arac_plaka', '')}</div>
            </div>
        </div>

        <div class="signature-section">
            <div style="background: #333; color: white; padding: 10px; font-weight: bold; margin-bottom: 1px;">
                HAZIRLAYAN - GÖREVLENDİREN
            </div>
            <div class="signature-row">
                <div class="signature-box">
                    <div class="signature-label">ADI - SOYADI</div>
                    <div class="field-value">{self.form_data.get('hazirlayan', '')}</div>
                </div>
                <div class="signature-box">
                    <div class="signature-label">İMZA</div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>"""
        return html
    
    def get_personel_summary(self):
        """Personel özetini oluştur"""
        personeller = self.form_data.get('personel_listesi', [])
        return '\n'.join([f"  - {p}" for p in personeller if p])
    
    def add_navigation_buttons(self):
        """Navigasyon butonları ekle"""
        btn_frame = tk.Frame(self.main_frame, bg='white')
        btn_frame.pack(side='bottom', pady=20)
        
        if self.current_step > 0:
            tk.Button(btn_frame, text="◀ Geri", font=('Arial', 12),
                     command=self.prev_step, bg='#e0e0e0', width=10).pack(side='left', padx=10)
        
        tk.Button(btn_frame, text="İleri ▶", font=('Arial', 12, 'bold'),
                 command=self.next_step, bg='#2196f3', fg='white', 
                 width=10).pack(side='left', padx=10)
    
    def next_step(self):
        """Bir sonraki adıma geç"""
        # Mevcut adımın verilerini kaydet
        if self.current_step == 1:
            personel_listesi = [var.get() for var in self.personel_vars if var.get()]
            self.form_data['personel_listesi'] = personel_listesi
        elif self.current_step == 2:
            self.form_data['avans_tutari'] = self.avans_entry.get()
            self.form_data['taseron_sirket'] = self.taseron_var.get()
        elif self.current_step == 3:
            self.form_data['gorev_tanimi'] = self.gorev_tanimi_text.get("1.0", "end-1c")
        elif self.current_step == 4:
            self.form_data['gorev_yeri'] = self.gorev_yeri_text.get("1.0", "end-1c")
        elif self.current_step == 5:
            self.form_data['yola_cikis_tarih'] = self.yola_cikis_tarih.get()
            self.form_data['yola_cikis_saat'] = self.yola_cikis_saat.get()
            self.form_data['donus_tarih'] = self.donus_tarih.get()
            self.form_data['donus_saat'] = self.donus_saat.get()
            self.form_data['calisma_baslangic_tarih'] = self.calisma_baslangic_tarih.get()
            self.form_data['calisma_baslangic_saat'] = self.calisma_baslangic_saat.get()
            self.form_data['calisma_bitis_tarih'] = self.calisma_bitis_tarih.get()
            self.form_data['calisma_bitis_saat'] = self.calisma_bitis_saat.get()
            self.form_data['mola_suresi'] = self.mola_suresi.get()
        elif self.current_step == 6:
            self.form_data['arac_plaka'] = self.arac_var.get()
        elif self.current_step == 7:
            self.form_data['hazirlayan'] = self.hazirlayan_var.get()
        
        self.current_step += 1
        self.show_step()
    
    def prev_step(self):
        """Bir önceki adıma dön"""
        if self.current_step > 0:
            self.current_step -= 1
            self.show_step()
    
    def save_to_excel(self):
        """Excel'e kaydet"""
        try:
            filename = f"gorev_formu_{self.form_no}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Görev Formu"
            
            # Başlık stilleri
            header_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
            border = Border(left=Side(style='thin'), right=Side(style='thin'),
                          top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Başlık
            ws.merge_cells('A1:B1')
            ws['A1'] = "DELTA PROJE - GÖREV FORMU"
            ws['A1'].font = Font(size=16, bold=True)
            ws['A1'].alignment = Alignment(horizontal='center')
            
            row = 3
            
            # Form bilgileri
            data_pairs = [
                ("Form No", self.form_data.get('form_no', '')),
                ("Tarih", self.form_data.get('tarih', '')),
                ("Avans Tutarı", self.form_data.get('avans_tutari', '')),
                ("Taşeron Şirket", self.form_data.get('taseron_sirket', '')),
            ]
            
            for label, value in data_pairs:
                ws[f'A{row}'] = label
                ws[f'A{row}'].font = Font(bold=True)
                ws[f'A{row}'].fill = header_fill
                ws[f'B{row}'] = value
                row += 1
            
            # Personel listesi
            row += 1
            ws[f'A{row}'] = "Görevli Personeller"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            row += 1
            
            for personel in self.form_data.get('personel_listesi', []):
                if personel:
                    ws[f'B{row}'] = personel
                    row += 1
            
            # Görev tanımı
            row += 1
            ws[f'A{row}'] = "Görevin Tanımı"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            row += 1
            ws.merge_cells(f'A{row}:B{row}')
            ws[f'A{row}'] = self.form_data.get('gorev_tanimi', '')
            ws[f'A{row}'].alignment = Alignment(wrap_text=True, vertical='top')
            
            # Görev yeri
            row += 2
            ws[f'A{row}'] = "Görev Yeri"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            row += 1
            ws.merge_cells(f'A{row}:B{row}')
            ws[f'A{row}'] = self.form_data.get('gorev_yeri', '')
            ws[f'A{row}'].alignment = Alignment(wrap_text=True, vertical='top')
            
            # Saat bilgileri
            row += 2
            ws[f'A{row}'] = "SAAT BİLGİLERİ"
            ws[f'A{row}'].font = Font(bold=True, size=12)
            ws[f'A{row}'].fill = header_fill
            row += 1
            
            time_data = [
                ("Yola Çıkış", f"{self.form_data.get('yola_cikis_tarih', '')} {self.form_data.get('yola_cikis_saat', '')}"),
                ("Dönüş", f"{self.form_data.get('donus_tarih', '')} {self.form_data.get('donus_saat', '')}"),
                ("Çalışma Başlangıç", f"{self.form_data.get('calisma_baslangic_tarih', '')} {self.form_data.get('calisma_baslangic_saat', '')}"),
                ("Çalışma Bitiş", f"{self.form_data.get('calisma_bitis_tarih', '')} {self.form_data.get('calisma_bitis_saat', '')}"),
                ("Mola Süresi", f"{self.form_data.get('mola_suresi', '')} dakika"),
            ]
            
            for label, value in time_data:
                ws[f'A{row}'] = label
                ws[f'A{row}'].font = Font(bold=True)
                ws[f'A{row}'].fill = header_fill
                ws[f'B{row}'] = value
                row += 1
            
            # Araç ve hazırlayan
            row += 1
            ws[f'A{row}'] = "Araç Plaka No"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('arac_plaka', '')
            row += 1
            
            ws[f'A{row}'] = "Hazırlayan"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('hazirlayan', '')
            
            # Sütun genişlikleri
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 50
            
            # Kaydet
            wb.save(filename)
            
            messagebox.showinfo("Başarılı", f"Form başarıyla kaydedildi!\n\nDosya: {filename}")
            
            # Yeni form için sıfırla
            self.reset_form()
            
        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme hatası: {str(e)}")
    
    def reset_form(self):
        """Formu sıfırla ve yeni form için hazırla"""
        self.form_data = {}
        self.current_step = 0
        self.form_no = self.get_next_form_no()
        self.show_step()


if __name__ == "__main__":
    root = tk.Tk()
    app = GorevFormuApp(root)
    root.mainloop()
