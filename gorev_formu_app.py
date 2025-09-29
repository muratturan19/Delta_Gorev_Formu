import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
import json
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import glob
from tkcalendar import DateEntry


class GorevFormuApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Delta Proje - Görev Formu Sistemi")
        self.root.geometry("800x600")
        self.root.configure(bg='#f5f5f5')

        # Mod: 'menu', 'new', 'edit'
        self.mode = 'menu'
        self.form_data = {}
        self.current_step = 0
        self.form_no = None
        self.is_readonly = False

        # Ana frame
        self.main_frame = tk.Frame(root, bg='white', padx=30, pady=30)
        self.main_frame.pack(fill='both', expand=True, padx=20, pady=20)

        # Ana menüyü göster
        self.show_main_menu()

    def get_next_form_no(self):
        """Yeni form numarası al"""
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

    def get_excel_filename(self, form_no):
        """Form numarasına göre Excel dosya adı"""
        return f"gorev_formu_{form_no}.xlsx"

    def load_form_from_excel(self, form_no):
        """Excel dosyasından formu yükle"""
        filename = self.get_excel_filename(form_no)
        if not os.path.exists(filename):
            return None

        try:
            wb = openpyxl.load_workbook(filename)
            ws = wb.active

            # Excel'den veri oku (basitleştirilmiş - gerçek implementasyon daha detaylı olmalı)
            data = {}
            for row in range(1, ws.max_row + 1):
                key = ws[f'A{row}'].value
                value = ws[f'B{row}'].value
                if key and value:
                    data[key] = value

            return data
        except Exception as e:
            messagebox.showerror("Hata", f"Form yüklenemedi: {str(e)}")
            return None

    def clear_frame(self):
        """Frame'i temizle"""
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def show_main_menu(self):
        """Ana menü ekranı"""
        self.clear_frame()
        self.mode = 'menu'

        # Logo/Başlık
        title = tk.Label(
            self.main_frame,
            text="🔧 DELTA PROJE\nGÖREV FORMU SİSTEMİ",
            font=('Arial', 24, 'bold'),
            bg='white',
            fg='#d32f2f'
        )
        title.pack(pady=50)

        # Butonlar frame
        button_frame = tk.Frame(self.main_frame, bg='white')
        button_frame.pack(expand=True)

        # Yeni Görev Oluştur butonu
        btn_new = tk.Button(
            button_frame,
            text="📝 YENİ GÖREV OLUŞTUR",
            font=('Arial', 16, 'bold'),
            bg='#4dd0e1',
            fg='black',
            width=25,
            height=3,
            command=self.start_new_form,
            cursor='hand2'
        )
        btn_new.pack(pady=15)

        # Görev Formu Çağır butonu
        btn_load = tk.Button(
            button_frame,
            text="📂 GÖREV FORMU ÇAĞIR",
            font=('Arial', 16, 'bold'),
            bg='#ffeb3b',
            fg='black',
            width=25,
            height=3,
            command=self.load_existing_form,
            cursor='hand2'
        )
        btn_load.pack(pady=15)

    def start_new_form(self):
        """Yeni form oluştur"""
        self.mode = 'new'
        self.form_data = {}
        self.current_step = 0
        self.form_no = self.get_next_form_no()
        self.is_readonly = False
        self.show_step()

    def load_existing_form(self):
        """Mevcut formu çağır"""
        self.clear_frame()

        # Başlık
        title = tk.Label(
            self.main_frame,
            text="📂 Form Çağır",
            font=('Arial', 20, 'bold'),
            bg='white',
            fg='#d32f2f'
        )
        title.pack(pady=30)

        # Açıklama
        info = tk.Label(
            self.main_frame,
            text="Tamamlanacak form numarasını seçin veya girin:",
            font=('Arial', 12),
            bg='white'
        )
        info.pack(pady=10)

        # Mevcut formları listele
        excel_files = glob.glob("gorev_formu_*.xlsx")
        form_numbers = []
        for file in excel_files:
            form_no = file.replace("gorev_formu_", "").replace(".xlsx", "")
            form_numbers.append(form_no)
        form_numbers.sort(reverse=True)

        # Form numarası seçimi
        entry_frame = tk.Frame(self.main_frame, bg='white')
        entry_frame.pack(pady=20)

        tk.Label(entry_frame, text="Form No:", font=('Arial', 14, 'bold'), bg='white').pack(side='left', padx=10)

        combo_state = 'readonly' if form_numbers else 'normal'

        form_no_combo = ttk.Combobox(
            entry_frame,
            font=('Arial', 14),
            width=15,
            values=form_numbers,
            state=combo_state
        )
        form_no_combo.pack(side='left', padx=10)
        if form_numbers:
            form_no_combo.current(0)
        form_no_combo.focus()

        def load_form():
            form_no = form_no_combo.get().strip().zfill(5)
            if not form_no:
                messagebox.showwarning("Uyarı", "Lütfen form numarası girin!")
                return

            filename = self.get_excel_filename(form_no)
            if not os.path.exists(filename):
                messagebox.showerror("Hata", f"Form {form_no} bulunamadı!\n\nDosya: {filename}")
                return

            # Formu yükle
            self.mode = 'edit'
            self.form_no = form_no
            self.load_partial_form(form_no)

        # Butonlar
        btn_frame = tk.Frame(self.main_frame, bg='white')
        btn_frame.pack(pady=30)

        tk.Button(
            btn_frame,
            text="✓ FORMU AÇ",
            font=('Arial', 12, 'bold'),
            bg='#4caf50',
            fg='white',
            width=15,
            command=load_form
        ).pack(side='left', padx=10)

        tk.Button(
            btn_frame,
            text="← Geri",
            font=('Arial', 12),
            bg='#ff9800',
            fg='white',
            width=15,
            command=self.show_main_menu
        ).pack(side='left', padx=10)

        # Enter tuşu ile aç
        form_no_combo.bind('<Return>', lambda e: load_form())

    def load_partial_form(self, form_no):
        """Kısmi dolu formu yükle ve devam et"""
        filename = self.get_excel_filename(form_no)

        try:
            wb = openpyxl.load_workbook(filename)
            ws = wb.active

            # Excel'deki tüm anahtar-değer çiftlerini oku
            raw_data = {}
            for key_cell, value_cell in ws.iter_rows(min_row=2, max_col=2, values_only=True):
                if key_cell:
                    raw_data[str(key_cell).strip()] = value_cell

            def parse_datetime_cell(value):
                """dd.mm.yyyy HH:MM formatındaki metni tarihe ve saate ayır."""
                tarih, saat = '', ''
                if isinstance(value, datetime):
                    tarih = value.strftime('%d.%m.%Y')
                    saat = value.strftime('%H:%M')
                elif isinstance(value, str):
                    cleaned = value.strip()
                    if cleaned:
                        parts = cleaned.split()
                        if len(parts) >= 2:
                            tarih = parts[0]
                            saat = parts[1]
                        elif ':' in cleaned:
                            saat = cleaned
                        else:
                            tarih = cleaned
                return tarih, saat

            def clean_mola_value(value):
                if isinstance(value, (int, float)):
                    return str(int(value))
                if isinstance(value, str):
                    return value.replace('dakika', '').strip()
                return ''

            self.form_data = {
                'form_no': form_no,
                'tarih': raw_data.get('Tarih', ''),
                'dok_no': raw_data.get('DOK.NO', ''),
                'rev_no': raw_data.get('REV.NO/TRH', ''),
                'avans': raw_data.get('Avans Tutarı', '') or '',
                'taseron': raw_data.get('Taşeron Şirket', '') or '',
                'gorev_tanimi': raw_data.get('Görevin Tanımı', '') or '',
                'gorev_yeri': raw_data.get('Görev Yeri', '') or '',
                'arac_plaka': raw_data.get('Araç Plaka No', '') or '',
                'hazirlayan': raw_data.get('Hazırlayan', '') or raw_data.get('Hazırlayan / Görevlendiren', '') or '',
            }

            for i in range(1, 6):
                key = f'Personel {i}'
                value = raw_data.get(key, '')
                if value:
                    self.form_data[f'personel_{i}'] = value

            # Tarih-saat alanlarını işle
            yola_tarih, yola_saat = parse_datetime_cell(raw_data.get('Yola Çıkış'))
            donus_tarih, donus_saat = parse_datetime_cell(raw_data.get('Dönüş'))
            calisma_baslangic_tarih, calisma_baslangic_saat = parse_datetime_cell(raw_data.get('Çalışma Başlangıç'))
            calisma_bitis_tarih, calisma_bitis_saat = parse_datetime_cell(raw_data.get('Çalışma Bitiş'))

            if yola_tarih:
                self.form_data['yola_cikis_tarih'] = yola_tarih
            if yola_saat:
                self.form_data['yola_cikis_saat'] = yola_saat
            if donus_tarih:
                self.form_data['donus_tarih'] = donus_tarih
            if donus_saat:
                self.form_data['donus_saat'] = donus_saat
            if calisma_baslangic_tarih:
                self.form_data['calisma_baslangic_tarih'] = calisma_baslangic_tarih
            if calisma_baslangic_saat:
                self.form_data['calisma_baslangic_saat'] = calisma_baslangic_saat
            if calisma_bitis_tarih:
                self.form_data['calisma_bitis_tarih'] = calisma_bitis_tarih
            if calisma_bitis_saat:
                self.form_data['calisma_bitis_saat'] = calisma_bitis_saat

            mola_value = clean_mola_value(raw_data.get('Toplam Mola'))
            if mola_value:
                self.form_data['mola_suresi'] = mola_value

            durum = (raw_data.get('DURUM') or '').strip().upper()

            if durum == 'TAMAMLANDI':
                self.current_step = 8  # Özet ekranını göster
            else:
                self.current_step = 5  # Saat bilgileri adımından başla

            self.show_step()

        except Exception as e:
            messagebox.showerror("Hata", f"Form okunamadı: {str(e)}")
            self.show_main_menu()

    def show_step(self):
        """Adımları göster"""
        self.clear_frame()

        # Mod kontrolü
        if self.mode == 'new':
            # Yeni form: 0-4 arası adımlar (Görev Yeri'ne kadar)
            if self.current_step > 4:
                self.save_partial_form()
                return

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
        """Adım 0: Form bilgileri"""
        readonly = self.mode == 'edit'

        tk.Label(self.main_frame, text="📋 Form Bilgileri", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        form_frame = tk.Frame(self.main_frame, bg='white')
        form_frame.pack(pady=20)

        # Form No
        tk.Label(form_frame, text="Form No:", font=('Arial', 12, 'bold'), bg='white').grid(row=0, column=0, sticky='w', pady=10)
        form_no_label = tk.Label(form_frame, text=self.form_no, font=('Arial', 12), bg='#e3f2fd', width=20, anchor='w')
        form_no_label.grid(row=0, column=1, padx=10, pady=10)

        # Tarih
        tk.Label(form_frame, text="Tarih:", font=('Arial', 12, 'bold'), bg='white').grid(row=1, column=0, sticky='w', pady=10)
        tarih_value = self.form_data.get('tarih', datetime.now().strftime('%d.%m.%Y'))
        tarih_label = tk.Label(form_frame, text=tarih_value, font=('Arial', 12), bg='#e3f2fd', width=20, anchor='w')
        tarih_label.grid(row=1, column=1, padx=10, pady=10)

        # DOK.NO
        tk.Label(form_frame, text="DOK.NO:", font=('Arial', 12, 'bold'), bg='white').grid(row=2, column=0, sticky='w', pady=10)
        self.dok_entry = tk.Entry(form_frame, font=('Arial', 12), width=20)
        self.dok_entry.insert(0, self.form_data.get('dok_no', 'F-001'))
        self.dok_entry.grid(row=2, column=1, padx=10, pady=10)

        # REV.NO/TRH
        tk.Label(form_frame, text="REV.NO/TRH:", font=('Arial', 12, 'bold'), bg='white').grid(row=3, column=0, sticky='w', pady=10)
        self.rev_entry = tk.Entry(form_frame, font=('Arial', 12), width=20)
        self.rev_entry.insert(0, self.form_data.get('rev_no', ''))
        self.rev_entry.grid(row=3, column=1, padx=10, pady=10)

        self.form_data['tarih'] = tarih_value

        self.add_navigation_buttons(readonly)

    def step_1_gorevli_personel(self):
        """Adım 1: Görevli personel"""
        readonly = self.mode == 'edit'

        tk.Label(self.main_frame, text="👥 Görevli Personel", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        personel_options = [
            "Ahmet Yılmaz", "Mehmet Demir", "Ali Kaya", "Veli Çelik",
            "Hasan Şahin", "Hüseyin Aydın", "İbrahim Özdemir", "Mustafa Arslan",
            "Emre Doğan", "Burak Yıldız"
        ]

        form_frame = tk.Frame(self.main_frame, bg='white')
        form_frame.pack(pady=20)

        self.personel_combos = []

        for i in range(5):
            tk.Label(form_frame, text=f"Personel {i+1}:", font=('Arial', 12, 'bold'), bg='white').grid(row=i, column=0, sticky='w', pady=10, padx=10)

            if readonly:
                value = self.form_data.get(f'personel_{i+1}', '')
                label = tk.Label(form_frame, text=value, font=('Arial', 12), bg='#f0f0f0', width=25, anchor='w')
                label.grid(row=i, column=1, padx=10, pady=10)
            else:
                combo = ttk.Combobox(form_frame, values=personel_options, font=('Arial', 12), width=23, state='readonly')
                combo.set(self.form_data.get(f'personel_{i+1}', ''))
                combo.grid(row=i, column=1, padx=10, pady=10)
                self.personel_combos.append(combo)

        if not readonly and self.personel_combos:

            def check_duplicate_personel(event=None):
                """Aynı personelin birden fazla seçilmesini engelle"""
                selected = []
                for combo in self.personel_combos:
                    value = combo.get()
                    if value:
                        if value in selected:
                            messagebox.showwarning(
                                "Uyarı",
                                f"'{value}' zaten seçilmiş!\nLütfen farklı bir personel seçin."
                            )
                            combo.set('')
                            return
                        selected.append(value)

            for combo in self.personel_combos:
                combo.bind('<<ComboboxSelected>>', check_duplicate_personel)

            # Var olan seçimleri doğrula
            check_duplicate_personel()

        self.add_navigation_buttons(readonly)

    def step_2_avans_taseron(self):
        """Adım 2: Avans ve Taşeron"""
        readonly = self.mode == 'edit'

        tk.Label(self.main_frame, text="💰 Avans ve Taşeron Bilgileri", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        form_frame = tk.Frame(self.main_frame, bg='white')
        form_frame.pack(pady=40)

        # Avans
        tk.Label(form_frame, text="Avans Tutarı:", font=('Arial', 12, 'bold'), bg='white').grid(row=0, column=0, sticky='w', pady=15)
        self.avans_entry = tk.Entry(form_frame, font=('Arial', 12), width=30)
        self.avans_entry.insert(0, self.form_data.get('avans', ''))
        self.avans_entry.grid(row=0, column=1, padx=10, pady=15)
        if readonly:
            self.avans_entry.config(state='readonly', bg='#f0f0f0')

        # Taşeron
        tk.Label(form_frame, text="Taşeron Şirket:", font=('Arial', 12, 'bold'), bg='white').grid(row=1, column=0, sticky='w', pady=15)

        taseron_options = ["Yok", "ABC İnşaat", "XYZ Teknik", "Marmara Mühendislik", "Anadolu Yapı"]

        if readonly:
            value = self.form_data.get('taseron', '')
            label = tk.Label(form_frame, text=value, font=('Arial', 12), bg='#f0f0f0', width=28, anchor='w')
            label.grid(row=1, column=1, padx=10, pady=15)
        else:
            self.taseron_combo = ttk.Combobox(form_frame, values=taseron_options, font=('Arial', 12), width=28)
            self.taseron_combo.set(self.form_data.get('taseron', ''))
            self.taseron_combo.grid(row=1, column=1, padx=10, pady=15)

        self.add_navigation_buttons(readonly)

    def step_3_gorev_tanimi(self):
        """Adım 3: Görev Tanımı"""
        readonly = self.mode == 'edit'

        tk.Label(self.main_frame, text="📝 Görevin Tanımı", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        self.gorev_tanimi_text = scrolledtext.ScrolledText(self.main_frame, font=('Arial', 11), width=70, height=15, wrap='word')
        self.gorev_tanimi_text.pack(pady=20, padx=20)
        self.gorev_tanimi_text.insert('1.0', self.form_data.get('gorev_tanimi', ''))

        if readonly:
            self.gorev_tanimi_text.config(state='disabled', bg='#f0f0f0')

        self.add_navigation_buttons(readonly)

    def step_4_gorev_yeri(self):
        """Adım 4: Görev Yeri"""
        readonly = self.mode == 'edit'

        tk.Label(self.main_frame, text="📍 Görev Yeri", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        self.gorev_yeri_text = scrolledtext.ScrolledText(self.main_frame, font=('Arial', 11), width=70, height=15, wrap='word')
        self.gorev_yeri_text.pack(pady=20, padx=20)
        self.gorev_yeri_text.insert('1.0', self.form_data.get('gorev_yeri', ''))

        if readonly:
            self.gorev_yeri_text.config(state='disabled', bg='#f0f0f0')

        self.add_navigation_buttons(readonly)

    def step_5_saat_bilgileri(self):
        """Adım 5: Saat bilgileri"""
        saat_list = [f"{i:02d}" for i in range(24)]
        dakika_list = [f"{i:02d}" for i in range(60)]

        tk.Label(self.main_frame, text="🕐 Saat ve Tarih Bilgileri", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        # Scroll frame
        canvas = tk.Canvas(self.main_frame, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='white')

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        form_frame = tk.Frame(scrollable_frame, bg='white')
        form_frame.pack(pady=10, padx=20)

        row = 0

        # Yola Çıkış
        tk.Label(form_frame, text="Yola Çıkış:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='w', pady=10)
        tk.Label(form_frame, text="Tarih:", bg='white').grid(row=row, column=1, sticky='e', padx=5)
        yola_cikis_tarih = DateEntry(form_frame, font=('Arial', 11), width=12, background='#4dd0e1',
                                      foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy',
                                      locale='tr_TR')
        yola_cikis_tarih.grid(row=row, column=2, padx=5)
        if self.form_data.get('yola_cikis_tarih'):
            try:
                yola_cikis_tarih.set_date(datetime.strptime(self.form_data.get('yola_cikis_tarih'), '%d.%m.%Y'))
            except:
                pass

        tk.Label(form_frame, text="Saat:", bg='white').grid(row=row, column=3, sticky='e', padx=5)

        # Saat frame
        saat_frame1 = tk.Frame(form_frame, bg='white')
        saat_frame1.grid(row=row, column=4, padx=5)

        yola_cikis_saat = ttk.Combobox(saat_frame1, values=saat_list, width=3, state='readonly', font=('Arial', 11))
        yola_cikis_saat.set('00')
        yola_cikis_saat.pack(side='left')
        tk.Label(saat_frame1, text=":", bg='white', font=('Arial', 11, 'bold')).pack(side='left')
        yola_cikis_dakika = ttk.Combobox(saat_frame1, values=dakika_list, width=3, state='readonly', font=('Arial', 11))
        yola_cikis_dakika.set('00')
        yola_cikis_dakika.pack(side='left')

        if self.form_data.get('yola_cikis_saat'):
            try:
                h, m = self.form_data.get('yola_cikis_saat', '00:00').split(':')
                yola_cikis_saat.set(h)
                yola_cikis_dakika.set(m)
            except:
                pass

        row += 1

        # Dönüş
        tk.Label(form_frame, text="Dönüş:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='w', pady=10)
        tk.Label(form_frame, text="Tarih:", bg='white').grid(row=row, column=1, sticky='e', padx=5)
        donus_tarih = DateEntry(form_frame, font=('Arial', 11), width=12, background='#4dd0e1',
                               foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy',
                               locale='tr_TR')
        donus_tarih.grid(row=row, column=2, padx=5)
        if self.form_data.get('donus_tarih'):
            try:
                donus_tarih.set_date(datetime.strptime(self.form_data.get('donus_tarih'), '%d.%m.%Y'))
            except:
                pass

        tk.Label(form_frame, text="Saat:", bg='white').grid(row=row, column=3, sticky='e', padx=5)

        saat_frame2 = tk.Frame(form_frame, bg='white')
        saat_frame2.grid(row=row, column=4, padx=5)

        donus_saat = ttk.Combobox(saat_frame2, values=saat_list, width=3, state='readonly', font=('Arial', 11))
        donus_saat.set('00')
        donus_saat.pack(side='left')
        tk.Label(saat_frame2, text=":", bg='white', font=('Arial', 11, 'bold')).pack(side='left')
        donus_dakika = ttk.Combobox(saat_frame2, values=dakika_list, width=3, state='readonly', font=('Arial', 11))
        donus_dakika.set('00')
        donus_dakika.pack(side='left')

        if self.form_data.get('donus_saat'):
            try:
                h, m = self.form_data.get('donus_saat', '00:00').split(':')
                donus_saat.set(h)
                donus_dakika.set(m)
            except:
                pass

        row += 1

        # Çalışma Başlangıç
        tk.Label(form_frame, text="Çalışma Başlangıç:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='w', pady=10)
        tk.Label(form_frame, text="Tarih:", bg='white').grid(row=row, column=1, sticky='e', padx=5)
        calisma_baslangic_tarih = DateEntry(form_frame, font=('Arial', 11), width=12, background='#4dd0e1',
                                           foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy',
                                           locale='tr_TR')
        calisma_baslangic_tarih.grid(row=row, column=2, padx=5)
        if self.form_data.get('calisma_baslangic_tarih'):
            try:
                calisma_baslangic_tarih.set_date(datetime.strptime(self.form_data.get('calisma_baslangic_tarih'), '%d.%m.%Y'))
            except:
                pass

        tk.Label(form_frame, text="Saat:", bg='white').grid(row=row, column=3, sticky='e', padx=5)

        saat_frame3 = tk.Frame(form_frame, bg='white')
        saat_frame3.grid(row=row, column=4, padx=5)

        calisma_baslangic_saat = ttk.Combobox(saat_frame3, values=saat_list, width=3, state='readonly', font=('Arial', 11))
        calisma_baslangic_saat.set('00')
        calisma_baslangic_saat.pack(side='left')
        tk.Label(saat_frame3, text=":", bg='white', font=('Arial', 11, 'bold')).pack(side='left')
        calisma_baslangic_dakika = ttk.Combobox(saat_frame3, values=dakika_list, width=3, state='readonly', font=('Arial', 11))
        calisma_baslangic_dakika.set('00')
        calisma_baslangic_dakika.pack(side='left')

        if self.form_data.get('calisma_baslangic_saat'):
            try:
                h, m = self.form_data.get('calisma_baslangic_saat', '00:00').split(':')
                calisma_baslangic_saat.set(h)
                calisma_baslangic_dakika.set(m)
            except:
                pass

        row += 1

        # Çalışma Bitiş
        tk.Label(form_frame, text="Çalışma Bitiş:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='w', pady=10)
        tk.Label(form_frame, text="Tarih:", bg='white').grid(row=row, column=1, sticky='e', padx=5)
        calisma_bitis_tarih = DateEntry(form_frame, font=('Arial', 11), width=12, background='#4dd0e1',
                                        foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy',
                                        locale='tr_TR')
        calisma_bitis_tarih.grid(row=row, column=2, padx=5)
        if self.form_data.get('calisma_bitis_tarih'):
            try:
                calisma_bitis_tarih.set_date(datetime.strptime(self.form_data.get('calisma_bitis_tarih'), '%d.%m.%Y'))
            except:
                pass

        tk.Label(form_frame, text="Saat:", bg='white').grid(row=row, column=3, sticky='e', padx=5)

        saat_frame4 = tk.Frame(form_frame, bg='white')
        saat_frame4.grid(row=row, column=4, padx=5)

        calisma_bitis_saat = ttk.Combobox(saat_frame4, values=saat_list, width=3, state='readonly', font=('Arial', 11))
        calisma_bitis_saat.set('00')
        calisma_bitis_saat.pack(side='left')
        tk.Label(saat_frame4, text=":", bg='white', font=('Arial', 11, 'bold')).pack(side='left')
        calisma_bitis_dakika = ttk.Combobox(saat_frame4, values=dakika_list, width=3, state='readonly', font=('Arial', 11))
        calisma_bitis_dakika.set('00')
        calisma_bitis_dakika.pack(side='left')

        if self.form_data.get('calisma_bitis_saat'):
            try:
                h, m = self.form_data.get('calisma_bitis_saat', '00:00').split(':')
                calisma_bitis_saat.set(h)
                calisma_bitis_dakika.set(m)
            except:
                pass

        row += 1

        # Mola Süresi
        tk.Label(form_frame, text="Toplam Mola Süresi:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='w', pady=10)
        mola_suresi = ttk.Spinbox(form_frame, from_=0, to=480, width=10, font=('Arial', 11))
        mola_suresi.set(self.form_data.get('mola_suresi', '0'))
        mola_suresi.grid(row=row, column=2, padx=5)
        tk.Label(form_frame, text="dakika", bg='white').grid(row=row, column=3, sticky='w', padx=5)

        # Widget'ları sakla
        self.yola_cikis_tarih_entry = yola_cikis_tarih
        self.yola_cikis_saat_combo = yola_cikis_saat
        self.yola_cikis_dakika_combo = yola_cikis_dakika
        self.donus_tarih_entry = donus_tarih
        self.donus_saat_combo = donus_saat
        self.donus_dakika_combo = donus_dakika
        self.calisma_baslangic_tarih_entry = calisma_baslangic_tarih
        self.calisma_baslangic_saat_combo = calisma_baslangic_saat
        self.calisma_baslangic_dakika_combo = calisma_baslangic_dakika
        self.calisma_bitis_tarih_entry = calisma_bitis_tarih
        self.calisma_bitis_saat_combo = calisma_bitis_saat
        self.calisma_bitis_dakika_combo = calisma_bitis_dakika
        self.mola_suresi_spin = mola_suresi

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.add_navigation_buttons(False, canvas_parent=True)

    def step_6_arac_bilgisi(self):
        """Adım 6: Araç bilgisi"""
        tk.Label(self.main_frame, text="🚗 Araç Bilgisi", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        form_frame = tk.Frame(self.main_frame, bg='white')
        form_frame.pack(pady=40)

        tk.Label(form_frame, text="Araç Plaka No:", font=('Arial', 12, 'bold'), bg='white').grid(row=0, column=0, sticky='w', pady=15)

        arac_options = [
            "34 ABC 123", "06 DEF 456", "41 GHI 789",
            "16 JKL 012", "35 MNO 345"
        ]

        self.arac_combo = ttk.Combobox(form_frame, values=arac_options, font=('Arial', 12), width=28, state='readonly')
        self.arac_combo.set(self.form_data.get('arac_plaka', ''))
        self.arac_combo.grid(row=0, column=1, padx=10, pady=15)

        self.add_navigation_buttons(False)

    def step_7_hazirlayan(self):
        """Adım 7: Hazırlayan"""
        tk.Label(self.main_frame, text="✍️ Hazırlayan / Görevlendiren", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        form_frame = tk.Frame(self.main_frame, bg='white')
        form_frame.pack(pady=40)

        tk.Label(form_frame, text="Ad Soyad:", font=('Arial', 12, 'bold'), bg='white').grid(row=0, column=0, sticky='w', pady=15)

        hazirlayan_options = [
            "Ahmet Yılmaz", "Mehmet Demir", "Ali Kaya",
            "Veli Çelik", "Hasan Şahin"
        ]

        self.hazirlayan_combo = ttk.Combobox(form_frame, values=hazirlayan_options, font=('Arial', 12), width=28, state='readonly')
        self.hazirlayan_combo.set(self.form_data.get('hazirlayan', ''))
        self.hazirlayan_combo.grid(row=0, column=1, padx=10, pady=15)

        self.add_navigation_buttons(False)

    def show_summary(self):
        """Özet ekranı"""
        # Önce tüm verileri topla
        self.collect_form_data()

        tk.Label(self.main_frame, text="📊 Form Özeti", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        # Ana canvas ve scrollbar
        canvas = tk.Canvas(self.main_frame, bg='white', highlightthickness=0, height=450)
        scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='white')

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Özet içeriği
        summary_frame = tk.Frame(scrollable_frame, bg='white', padx=30, pady=30)
        summary_frame.pack(fill='both', expand=True)

        # Başlık kutusu
        header_frame = tk.Frame(summary_frame, bg='#d32f2f', padx=15, pady=10)
        header_frame.pack(fill='x', pady=(0, 10))
        tk.Label(header_frame, text="DELTA PROJE - GÖREV FORMU",
                font=('Arial', 14, 'bold'), bg='#d32f2f', fg='white').pack()

        # Form Bilgileri
        self.add_summary_section(summary_frame, "📋 FORM BİLGİLERİ", [
            ("Form No", self.form_no),
            ("Tarih", self.form_data.get('tarih', '')),
            ("DOK.NO", self.form_data.get('dok_no', '')),
            ("REV.NO/TRH", self.form_data.get('rev_no', ''))
        ])

        # Görevli Personel
        personel_list = []
        for i in range(1, 6):
            personel = self.form_data.get(f'personel_{i}', '')
            if personel:
                personel_list.append((f"Personel {i}", personel))

        if personel_list:
            self.add_summary_section(summary_frame, "👥 GÖREVLİ PERSONEL", personel_list)

        # Mali Bilgiler
        mali_data = []
        if self.form_data.get('avans'):
            mali_data.append(("Avans Tutarı", self.form_data.get('avans', '')))
        if self.form_data.get('taseron'):
            mali_data.append(("Taşeron Şirket", self.form_data.get('taseron', '')))

        if mali_data:
            self.add_summary_section(summary_frame, "💰 MALİ BİLGİLER", mali_data)

        # Görev Detayları
        gorev_data = []
        if self.form_data.get('gorev_tanimi'):
            gorev_data.append(("Görevin Tanımı", self.form_data.get('gorev_tanimi', '')))
        if self.form_data.get('gorev_yeri'):
            gorev_data.append(("Görev Yeri", self.form_data.get('gorev_yeri', '')))

        if gorev_data:
            self.add_summary_section(summary_frame, "📝 GÖREV DETAYLARI", gorev_data)

        # Zaman Bilgileri
        zaman_data = []
        if self.form_data.get('yola_cikis_tarih'):
            zaman_data.append(("Yola Çıkış", f"{self.form_data.get('yola_cikis_tarih', '')} {self.form_data.get('yola_cikis_saat', '')}"))
        if self.form_data.get('donus_tarih'):
            zaman_data.append(("Dönüş", f"{self.form_data.get('donus_tarih', '')} {self.form_data.get('donus_saat', '')}"))
        if self.form_data.get('calisma_baslangic_tarih'):
            zaman_data.append(("Çalışma Başlangıç", f"{self.form_data.get('calisma_baslangic_tarih', '')} {self.form_data.get('calisma_baslangic_saat', '')}"))
        if self.form_data.get('calisma_bitis_tarih'):
            zaman_data.append(("Çalışma Bitiş", f"{self.form_data.get('calisma_bitis_tarih', '')} {self.form_data.get('calisma_bitis_saat', '')}"))
        if self.form_data.get('mola_suresi'):
            zaman_data.append(("Toplam Mola", f"{self.form_data.get('mola_suresi', '')} dakika"))

        if zaman_data:
            self.add_summary_section(summary_frame, "🕐 ZAMAN BİLGİLERİ", zaman_data)

        # Diğer Bilgiler
        diger_data = []
        if self.form_data.get('arac_plaka'):
            diger_data.append(("Araç Plaka No", self.form_data.get('arac_plaka', '')))
        if self.form_data.get('hazirlayan'):
            diger_data.append(("Hazırlayan / Görevlendiren", self.form_data.get('hazirlayan', '')))

        if diger_data:
            self.add_summary_section(summary_frame, "🚗 DİĞER BİLGİLER", diger_data)

        canvas.pack(side="left", fill="both", expand=True, padx=(0, 0))
        scrollbar.pack(side="right", fill="y")

        # Butonlar
        btn_frame = tk.Frame(self.main_frame, bg='white')
        btn_frame.pack(side='bottom', pady=20)

        tk.Button(
            btn_frame,
            text="💾 KAYDET",
            font=('Arial', 14, 'bold'),
            bg='#4caf50',
            fg='white',
            width=15,
            height=2,
            command=self.save_form,
            cursor='hand2'
        ).pack(side='left', padx=10)

        tk.Button(
            btn_frame,
            text="← Geri",
            font=('Arial', 12),
            bg='#ff9800',
            fg='white',
            width=15,
            command=self.previous_step,
            cursor='hand2'
        ).pack(side='left', padx=10)

    def add_summary_section(self, parent, title, data_list):
        """Özet bölümü ekle"""
        # Bölüm frame'i
        section_frame = tk.Frame(parent, bg='white', relief='solid', borderwidth=1)
        section_frame.pack(fill='x', pady=5)

        # Başlık
        title_frame = tk.Frame(section_frame, bg='#ffeb3b', padx=10, pady=5)
        title_frame.pack(fill='x')
        tk.Label(title_frame, text=title, font=('Arial', 12, 'bold'),
                bg='#ffeb3b', fg='#000', anchor='w').pack(fill='x')

        # Veri satırları
        for label, value in data_list:
            row_frame = tk.Frame(section_frame, bg='white', padx=10, pady=3)
            row_frame.pack(fill='x')

            # Label
            tk.Label(row_frame, text=f"{label}:", font=('Arial', 10, 'bold'),
                    bg='white', fg='#333', anchor='w', width=25).pack(side='left')

            # Value - uzun metinler için text widget
            if len(str(value)) > 50:
                value_text = tk.Text(row_frame, font=('Arial', 10), bg='#f5f5f5',
                                    height=3, width=50, wrap='word', relief='flat')
                value_text.insert('1.0', str(value))
                value_text.config(state='disabled')
                value_text.pack(side='left', fill='x', expand=True, padx=(5, 0))
            else:
                tk.Label(row_frame, text=str(value), font=('Arial', 10),
                        bg='#f5f5f5', fg='#000', anchor='w',
                        relief='flat', padx=8, pady=2).pack(side='left', fill='x', expand=True, padx=(5, 0))

    def collect_form_data(self):
        """Widget'lardan veri topla"""
        try:
            # Form bilgileri
            if hasattr(self, 'dok_entry') and self.dok_entry.winfo_exists():
                self.form_data['dok_no'] = self.dok_entry.get()
            if hasattr(self, 'rev_entry') and self.rev_entry.winfo_exists():
                self.form_data['rev_no'] = self.rev_entry.get()

            # Personel - combobox'lardan al
            if hasattr(self, 'personel_combos'):
                for i, combo in enumerate(self.personel_combos):
                    if combo.winfo_exists():
                        value = combo.get()
                        if value:
                            self.form_data[f'personel_{i+1}'] = value

            # Avans ve Taşeron
            if hasattr(self, 'avans_entry') and self.avans_entry.winfo_exists():
                self.form_data['avans'] = self.avans_entry.get()
            if hasattr(self, 'taseron_combo') and self.taseron_combo.winfo_exists():
                self.form_data['taseron'] = self.taseron_combo.get()

            # Görev tanımı ve yeri
            if hasattr(self, 'gorev_tanimi_text') and self.gorev_tanimi_text.winfo_exists():
                self.form_data['gorev_tanimi'] = self.gorev_tanimi_text.get('1.0', 'end-1c')
            if hasattr(self, 'gorev_yeri_text') and self.gorev_yeri_text.winfo_exists():
                self.form_data['gorev_yeri'] = self.gorev_yeri_text.get('1.0', 'end-1c')

            # Saat bilgileri
            if hasattr(self, 'yola_cikis_tarih_entry'):
                if self.yola_cikis_tarih_entry.winfo_exists():
                    self.form_data['yola_cikis_tarih'] = self.yola_cikis_tarih_entry.get_date().strftime('%d.%m.%Y')
                    h = int(self.yola_cikis_saat_combo.get() or 0)
                    m = int(self.yola_cikis_dakika_combo.get() or 0)
                    self.form_data['yola_cikis_saat'] = f"{h:02d}:{m:02d}"

                    self.form_data['donus_tarih'] = self.donus_tarih_entry.get_date().strftime('%d.%m.%Y')
                    h = int(self.donus_saat_combo.get() or 0)
                    m = int(self.donus_dakika_combo.get() or 0)
                    self.form_data['donus_saat'] = f"{h:02d}:{m:02d}"

                    self.form_data['calisma_baslangic_tarih'] = self.calisma_baslangic_tarih_entry.get_date().strftime('%d.%m.%Y')
                    h = int(self.calisma_baslangic_saat_combo.get() or 0)
                    m = int(self.calisma_baslangic_dakika_combo.get() or 0)
                    self.form_data['calisma_baslangic_saat'] = f"{h:02d}:{m:02d}"

                    self.form_data['calisma_bitis_tarih'] = self.calisma_bitis_tarih_entry.get_date().strftime('%d.%m.%Y')
                    h = int(self.calisma_bitis_saat_combo.get() or 0)
                    m = int(self.calisma_bitis_dakika_combo.get() or 0)
                    self.form_data['calisma_bitis_saat'] = f"{h:02d}:{m:02d}"

                    self.form_data['mola_suresi'] = self.mola_suresi_spin.get()

            # Araç ve hazırlayan
            if hasattr(self, 'arac_combo') and self.arac_combo.winfo_exists():
                self.form_data['arac_plaka'] = self.arac_combo.get()
            if hasattr(self, 'hazirlayan_combo') and self.hazirlayan_combo.winfo_exists():
                self.form_data['hazirlayan'] = self.hazirlayan_combo.get()

        except Exception as e:
            print(f"Veri toplama hatası: {e}")

    def save_partial_form(self):
        """Kısmi formu kaydet (Görev Yeri'ne kadar)"""
        self.collect_form_data()

        filename = self.get_excel_filename(self.form_no)

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Görev Formu"

            # Stil
            header_fill = PatternFill(start_color='FFEB3B', end_color='FFEB3B', fill_type='solid')
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            row = 1

            # Başlık
            ws[f'A{row}'] = "DELTA PROJE - GÖREV FORMU"
            ws[f'A{row}'].font = Font(size=16, bold=True, color='D32F2F')
            ws.merge_cells(f'A{row}:B{row}')
            row += 1

            # Form bilgileri
            ws[f'A{row}'] = "Form No"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_no
            row += 1

            ws[f'A{row}'] = "Tarih"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('tarih', '')
            row += 1

            ws[f'A{row}'] = "DOK.NO"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('dok_no', '')
            row += 1

            ws[f'A{row}'] = "REV.NO/TRH"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('rev_no', '')
            row += 1

            # Personel
            ws[f'A{row}'] = "Görevli Personel"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            row += 1

            for i in range(5):
                ws[f'A{row}'] = f"Personel {i+1}"
                ws[f'B{row}'] = self.form_data.get(f'personel_{i+1}', '')
                row += 1

            # Diğer bilgiler
            ws[f'A{row}'] = "Avans Tutarı"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('avans', '')
            row += 1

            ws[f'A{row}'] = "Taşeron Şirket"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('taseron', '')
            row += 1

            ws[f'A{row}'] = "Görevin Tanımı"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('gorev_tanimi', '')
            row += 1

            ws[f'A{row}'] = "Görev Yeri"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'] = self.form_data.get('gorev_yeri', '')
            row += 1

            # Durum
            ws[f'A{row}'] = "DURUM"
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'A{row}'].fill = PatternFill(start_color='FF9800', end_color='FF9800', fill_type='solid')
            ws[f'B{row}'] = "YARIM"
            ws[f'B{row}'].fill = PatternFill(start_color='FFC107', end_color='FFC107', fill_type='solid')

            # Sütun genişlikleri
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 60

            wb.save(filename)

            messagebox.showinfo(
                "Başarılı",
                f"Form oluşturuldu!\n\nForm No: {self.form_no}\nDosya: {filename}\n\nGörev tamamlandığında 'GÖREV FORMU ÇAĞIR' ile bu formu açıp kalan kısımları doldurun."
            )

            self.show_main_menu()

        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme hatası: {str(e)}")

    def save_form(self):
        """Tam formu kaydet"""
        self.collect_form_data()

        filename = self.get_excel_filename(self.form_no)

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Görev Formu"

            # Stil
            header_fill = PatternFill(start_color='FFEB3B', end_color='FFEB3B', fill_type='solid')

            row = 1

            # Başlık
            ws[f'A{row}'] = "DELTA PROJE - GÖREV FORMU"
            ws[f'A{row}'].font = Font(size=16, bold=True, color='D32F2F')
            ws.merge_cells(f'A{row}:B{row}')
            row += 1

            # Tüm bilgileri yaz
            data_map = [
                ("Form No", self.form_no),
                ("Tarih", self.form_data.get('tarih', '')),
                ("DOK.NO", self.form_data.get('dok_no', '')),
                ("REV.NO/TRH", self.form_data.get('rev_no', '')),
                ("", ""),
                ("Görevli Personel", ""),
            ]

            for label, value in data_map:
                if label:
                    ws[f'A{row}'] = label
                    ws[f'A{row}'].font = Font(bold=True)
                    ws[f'A{row}'].fill = header_fill
                    ws[f'B{row}'] = value
                row += 1

            # Personel listesi
            for i in range(5):
                ws[f'A{row}'] = f"Personel {i+1}"
                ws[f'B{row}'] = self.form_data.get(f'personel_{i+1}', '')
                row += 1

            row += 1

            # Diğer tüm alanlar
            all_data = [
                ("Avans Tutarı", self.form_data.get('avans', '')),
                ("Taşeron Şirket", self.form_data.get('taseron', '')),
                ("Görevin Tanımı", self.form_data.get('gorev_tanimi', '')),
                ("Görev Yeri", self.form_data.get('gorev_yeri', '')),
                ("", ""),
                ("Yola Çıkış", f"{self.form_data.get('yola_cikis_tarih', '')} {self.form_data.get('yola_cikis_saat', '')}"),
                ("Dönüş", f"{self.form_data.get('donus_tarih', '')} {self.form_data.get('donus_saat', '')}"),
                ("Çalışma Başlangıç", f"{self.form_data.get('calisma_baslangic_tarih', '')} {self.form_data.get('calisma_baslangic_saat', '')}"),
                ("Çalışma Bitiş", f"{self.form_data.get('calisma_bitis_tarih', '')} {self.form_data.get('calisma_bitis_saat', '')}"),
                ("Toplam Mola", f"{self.form_data.get('mola_suresi', '')} dakika"),
                ("", ""),
                ("Araç Plaka No", self.form_data.get('arac_plaka', '')),
                ("Hazırlayan", self.form_data.get('hazirlayan', '')),
                ("", ""),
                ("DURUM", "TAMAMLANDI"),
            ]

            for label, value in all_data:
                if label:
                    ws[f'A{row}'] = label
                    ws[f'A{row}'].font = Font(bold=True)
                    if label == "DURUM":
                        ws[f'A{row}'].fill = PatternFill(start_color='4CAF50', end_color='4CAF50', fill_type='solid')
                        ws[f'B{row}'].fill = PatternFill(start_color='81C784', end_color='81C784', fill_type='solid')
                    else:
                        ws[f'A{row}'].fill = header_fill
                    ws[f'B{row}'] = value
                row += 1

            # Sütun genişlikleri
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 60

            wb.save(filename)

            messagebox.showinfo(
                "Başarılı",
                f"Form başarıyla tamamlandı ve kaydedildi!\n\nForm No: {self.form_no}\nDosya: {filename}"
            )

            self.show_main_menu()

        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme hatası: {str(e)}")

    def add_navigation_buttons(self, readonly=False, canvas_parent=False):
        """Navigasyon butonları ekle"""
        parent = self.main_frame if not canvas_parent else self.root

        btn_frame = tk.Frame(parent, bg='white')
        if canvas_parent:
            btn_frame.pack(side='bottom', pady=20, fill='x')
        else:
            btn_frame.pack(side='bottom', pady=30, fill='x')

        # Butonları ortala
        center_frame = tk.Frame(btn_frame, bg='white')
        center_frame.pack(expand=True)

        if self.current_step > 0:
            btn_geri = tk.Button(
                center_frame,
                text="← Geri",
                font=('Arial', 13, 'bold'),
                bg='#ff9800',
                fg='white',
                width=15,
                height=2,
                command=self.previous_step,
                cursor='hand2',
                relief='raised',
                bd=3
            )
            btn_geri.pack(side='left', padx=15)

        if self.mode == 'new' and self.current_step >= 4:
            # Yeni form modunda Görev Yeri'nden sonra kaydet
            btn_kaydet = tk.Button(
                center_frame,
                text="💾 Kaydet",
                font=('Arial', 13, 'bold'),
                bg='#4caf50',
                fg='white',
                width=15,
                height=2,
                command=lambda: self.next_step(save_partial=True),
                cursor='hand2',
                relief='raised',
                bd=3
            )
            btn_kaydet.pack(side='left', padx=15)
        else:
            # Normal ilerleme
            btn_ileri = tk.Button(
                center_frame,
                text="İleri →",
                font=('Arial', 13, 'bold'),
                bg='#2196f3',
                fg='white',
                width=15,
                height=2,
                command=self.next_step,
                cursor='hand2',
                relief='raised',
                bd=3
            )
            btn_ileri.pack(side='left', padx=15)

    def next_step(self, save_partial=False):
        """Sonraki adım"""
        self.collect_form_data()

        if save_partial:
            self.save_partial_form()
            return

        self.current_step += 1
        self.show_step()

    def previous_step(self):
        """Önceki adım"""
        self.collect_form_data()

        if self.mode == 'edit' and self.current_step == 5:
            # Edit modunda geri dönmeye izin verme
            messagebox.showwarning("Uyarı", "Önceki adımlara dönüş yapılamaz!")
            return

        self.current_step -= 1
        self.show_step()


if __name__ == "__main__":
    root = tk.Tk()
    app = GorevFormuApp(root)
    root.mainloop()
