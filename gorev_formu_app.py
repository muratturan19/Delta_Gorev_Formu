import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
import glob
from tkcalendar import DateEntry

from core import form_service
from core.form_service import FormServiceError


class GorevFormuApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Delta Proje - Görev Formu Sistemi")
        self.root.geometry("800x600")
        self.root.minsize(900, 650)
        self.root.configure(bg='#f5f5f5')

        # Mod: 'menu', 'new', 'edit'
        self.mode = 'menu'
        self.form_data = {}
        self.current_step = 0
        self.form_no = None
        self.is_readonly = False
        self.nav_frame = None

        # Ana frame
        self.main_frame = tk.Frame(root, bg='white', padx=30, pady=30)
        self.main_frame.pack(fill='both', expand=True, padx=20, pady=20)

        # Ana menüyü göster
        self.show_main_menu()

    def clear_frame(self):
        """Frame'i temizle"""
        if self.nav_frame is not None and self.nav_frame.winfo_exists():
            self.nav_frame.destroy()
        self.nav_frame = None
        # Mouse wheel binding'leri temizle
        try:
            self.root.unbind_all("<MouseWheel>")
            self.root.unbind_all("<Button-4>")
            self.root.unbind_all("<Button-5>")
        except Exception:
            pass
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
        self.form_data = {'durum': 'YARIM'}
        self.current_step = 0
        try:
            self.form_no = form_service.get_next_form_no()
        except FormServiceError as exc:
            messagebox.showerror("Hata", str(exc))
            self.show_main_menu()
            return
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
        try:
            self.form_data = form_service.load_form_data(form_no)
        except FormServiceError as exc:
            messagebox.showerror("Hata", str(exc))
            self.show_main_menu()
            return

        self.mode = 'edit'
        self.form_no = form_no
        self.current_step = 0
        self.is_readonly = False
        self.show_step()

    def show_step(self):
        """Adımları göster"""
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

            combo_values = [''] + personel_options
            combo = ttk.Combobox(form_frame, values=combo_values, font=('Arial', 12), width=23, state='readonly')
            combo.set(self.form_data.get(f'personel_{i+1}', ''))
            combo.grid(row=i, column=1, padx=10, pady=10)
            self.personel_combos.append(combo)

        if self.personel_combos:

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

        self.add_navigation_buttons(False)

    def step_2_avans_taseron(self):
        """Adım 2: Avans ve Taşeron"""
        tk.Label(self.main_frame, text="💰 Avans ve Taşeron Bilgileri", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        form_frame = tk.Frame(self.main_frame, bg='white')
        form_frame.pack(pady=40)

        # Avans
        tk.Label(form_frame, text="Avans Tutarı:", font=('Arial', 12, 'bold'), bg='white').grid(row=0, column=0, sticky='w', pady=15)
        self.avans_entry = tk.Entry(form_frame, font=('Arial', 12), width=30)
        self.avans_entry.insert(0, self.form_data.get('avans', ''))
        self.avans_entry.grid(row=0, column=1, padx=10, pady=15)

        # Taşeron
        tk.Label(form_frame, text="Taşeron Şirket:", font=('Arial', 12, 'bold'), bg='white').grid(row=1, column=0, sticky='w', pady=15)

        taseron_options = ["Yok", "ABC İnşaat", "XYZ Teknik", "Marmara Mühendislik", "Anadolu Yapı"]

        combo_values = [''] + taseron_options
        self.taseron_combo = ttk.Combobox(form_frame, values=combo_values, font=('Arial', 12), width=28)
        self.taseron_combo.set(self.form_data.get('taseron', ''))
        self.taseron_combo.grid(row=1, column=1, padx=10, pady=15)

        self.add_navigation_buttons(False)

    def step_3_gorev_tanimi(self):
        """Adım 3: Görev Tanımı"""
        tk.Label(self.main_frame, text="📝 Görevin Tanımı", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        self.gorev_tanimi_text = scrolledtext.ScrolledText(self.main_frame, font=('Arial', 11), width=70, height=15, wrap='word')
        self.gorev_tanimi_text.pack(pady=20, padx=20)
        self.gorev_tanimi_text.insert('1.0', self.form_data.get('gorev_tanimi', ''))

        self.add_navigation_buttons(False)

    def step_4_gorev_yeri(self):
        """Adım 4: Görev Yeri"""
        tk.Label(self.main_frame, text="📍 Görev Yeri", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        self.gorev_yeri_text = scrolledtext.ScrolledText(self.main_frame, font=('Arial', 11), width=70, height=15, wrap='word')
        self.gorev_yeri_text.pack(pady=20, padx=20)
        self.gorev_yeri_text.insert('1.0', self.form_data.get('gorev_yeri', ''))

        self.add_navigation_buttons(False)

    def step_5_saat_bilgileri(self):
        """Adım 5: Saat bilgileri"""
        saat_list = [''] + [f"{i:02d}" for i in range(24)]
        dakika_list = [''] + [f"{i:02d}" for i in range(60)]

        tk.Label(self.main_frame, text="🕐 Saat ve Tarih Bilgileri", font=('Arial', 18, 'bold'), bg='white', fg='#d32f2f').pack(pady=20)

        # Scroll frame
        canvas = tk.Canvas(self.main_frame, bg='white', highlightthickness=0, height=450)
        scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='white')

        def _on_mousewheel_step5(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind_all("<MouseWheel>", _on_mousewheel_step5)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        form_frame = tk.Frame(scrollable_frame, bg='white')
        form_frame.pack(pady=10, padx=20, fill='both', expand=True)

        for col in range(5):
            weight = 0 if col == 4 else 1
            form_frame.grid_columnconfigure(col, weight=weight)

        row = 0

        # Yola Çıkış
        tk.Label(form_frame, text="Yola Çıkış:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='ew', pady=10)
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
        else:
            yola_cikis_tarih.delete(0, 'end')

        tk.Label(form_frame, text="Saat:", bg='white').grid(row=row, column=3, sticky='e', padx=5)

        # Saat frame
        saat_frame1 = tk.Frame(form_frame, bg='white')
        saat_frame1.grid(row=row, column=4, padx=5)

        yola_cikis_saat = ttk.Combobox(saat_frame1, values=saat_list, width=3, state='readonly', font=('Arial', 11))
        yola_cikis_saat.set('')
        yola_cikis_saat.pack(side='left')
        tk.Label(saat_frame1, text=":", bg='white', font=('Arial', 11, 'bold')).pack(side='left')
        yola_cikis_dakika = ttk.Combobox(saat_frame1, values=dakika_list, width=3, state='readonly', font=('Arial', 11))
        yola_cikis_dakika.set('')
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
        tk.Label(form_frame, text="Dönüş:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='ew', pady=10)
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
        else:
            donus_tarih.delete(0, 'end')

        tk.Label(form_frame, text="Saat:", bg='white').grid(row=row, column=3, sticky='e', padx=5)

        saat_frame2 = tk.Frame(form_frame, bg='white')
        saat_frame2.grid(row=row, column=4, padx=5)

        donus_saat = ttk.Combobox(saat_frame2, values=saat_list, width=3, state='readonly', font=('Arial', 11))
        donus_saat.set('')
        donus_saat.pack(side='left')
        tk.Label(saat_frame2, text=":", bg='white', font=('Arial', 11, 'bold')).pack(side='left')
        donus_dakika = ttk.Combobox(saat_frame2, values=dakika_list, width=3, state='readonly', font=('Arial', 11))
        donus_dakika.set('')
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
        tk.Label(form_frame, text="Çalışma Başlangıç:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='ew', pady=10)
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
        else:
            calisma_baslangic_tarih.delete(0, 'end')

        tk.Label(form_frame, text="Saat:", bg='white').grid(row=row, column=3, sticky='e', padx=5)

        saat_frame3 = tk.Frame(form_frame, bg='white')
        saat_frame3.grid(row=row, column=4, padx=5)

        calisma_baslangic_saat = ttk.Combobox(saat_frame3, values=saat_list, width=3, state='readonly', font=('Arial', 11))
        calisma_baslangic_saat.set('')
        calisma_baslangic_saat.pack(side='left')
        tk.Label(saat_frame3, text=":", bg='white', font=('Arial', 11, 'bold')).pack(side='left')
        calisma_baslangic_dakika = ttk.Combobox(saat_frame3, values=dakika_list, width=3, state='readonly', font=('Arial', 11))
        calisma_baslangic_dakika.set('')
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
        tk.Label(form_frame, text="Çalışma Bitiş:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='ew', pady=10)
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
        else:
            calisma_bitis_tarih.delete(0, 'end')

        tk.Label(form_frame, text="Saat:", bg='white').grid(row=row, column=3, sticky='e', padx=5)

        saat_frame4 = tk.Frame(form_frame, bg='white')
        saat_frame4.grid(row=row, column=4, padx=5)

        calisma_bitis_saat = ttk.Combobox(saat_frame4, values=saat_list, width=3, state='readonly', font=('Arial', 11))
        calisma_bitis_saat.set('')
        calisma_bitis_saat.pack(side='left')
        tk.Label(saat_frame4, text=":", bg='white', font=('Arial', 11, 'bold')).pack(side='left')
        calisma_bitis_dakika = ttk.Combobox(saat_frame4, values=dakika_list, width=3, state='readonly', font=('Arial', 11))
        calisma_bitis_dakika.set('')
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
        tk.Label(form_frame, text="Toplam Mola Süresi:", font=('Arial', 12, 'bold'), bg='white').grid(row=row, column=0, sticky='ew', pady=10)
        mola_suresi = ttk.Spinbox(form_frame, from_=0, to=480, width=10, font=('Arial', 11))
        mola_suresi.set(self.form_data.get('mola_suresi', '0'))
        mola_suresi.grid(row=row, column=2, padx=5)
        tk.Label(form_frame, text="dakika", bg='white').grid(row=row, column=3, sticky='ew', padx=5)

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

        self.arac_combo = ttk.Combobox(form_frame, values=[''] + arac_options, font=('Arial', 12), width=28, state='readonly')
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

        self.hazirlayan_combo = ttk.Combobox(form_frame, values=[''] + hazirlayan_options, font=('Arial', 12), width=28, state='readonly')
        self.hazirlayan_combo.set(self.form_data.get('hazirlayan', ''))
        self.hazirlayan_combo.grid(row=0, column=1, padx=10, pady=15)

        self.add_navigation_buttons(False)

    def show_summary(self):
        """Özet ekranı"""
        # Eski mouse wheel binding'leri temizle
        self.root.unbind_all("<MouseWheel>")
        self.root.unbind_all("<Button-4>")
        self.root.unbind_all("<Button-5>")
        # Önce tüm verileri topla
        self.collect_form_data()
        status = form_service.determine_form_status(self.form_data)
        self.form_data['durum'] = status.code

        tk.Label(
            self.main_frame,
            text="📊 Form Özeti",
            font=('Arial', 18, 'bold'),
            bg='white',
            fg='#d32f2f'
        ).pack(pady=20)

        # Ana canvas ve scrollbar
        canvas = tk.Canvas(self.main_frame, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='white')

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        # Özet içeriği
        summary_frame = tk.Frame(scrollable_frame, bg='#f0f4f7', padx=25, pady=25)
        summary_frame.pack(fill='both', expand=True)

        report_frame = tk.Frame(summary_frame, bg='white', bd=2, relief='solid')
        report_frame.pack(fill='both', expand=True)

        grid_frame = tk.Frame(report_frame, bg='white')
        grid_frame.pack(fill='both', expand=True, padx=3, pady=3)

        for col in range(6):
            if col in [0, 1]:
                grid_frame.grid_columnconfigure(col, weight=1, minsize=120)
            elif col == 4:
                grid_frame.grid_columnconfigure(col, weight=1, minsize=100)
            else:
                grid_frame.grid_columnconfigure(col, weight=2, minsize=150)

        def create_cell(row, column, text='', colspan=1, rowspan=1, bg='white', fg='#000',
                        font=('Arial', 10), anchor='w', padx=8, pady=6, wrap=None,
                        border=True, justify='left'):
            label = tk.Label(
                grid_frame,
                text=text,
                bg=bg,
                fg=fg,
                font=font,
                anchor=anchor,
                justify=justify,
                padx=padx,
                pady=pady,
                wraplength=wrap
            )
            if border:
                label.configure(relief='solid', borderwidth=1)
            else:
                label.configure(borderwidth=0)
            label.grid(row=row, column=column, columnspan=colspan, rowspan=rowspan, sticky='nsew')
            return label

        # Üst başlık
        create_cell(
            0,
            0,
            text="Delta Proje\nHidrolik & Pnömatik",
            colspan=2,
            bg='white',
            fg='#d32f2f',
            font=('Arial', 12, 'bold'),
            anchor='center',
            justify='center'
        )
        create_cell(
            0,
            2,
            text="GÖREV FORMU",
            colspan=2,
            bg='#0d47a1',
            fg='white',
            font=('Arial', 18, 'bold'),
            anchor='center'
        )
        create_cell(0, 4, "FORM NO", bg='#fff176', font=('Arial', 11, 'bold'), anchor='center')
        create_cell(0, 5, self.form_no or '-', bg='#bbdefb', font=('Arial', 11, 'bold'), anchor='center')

        tarih = self.form_data.get('tarih', '') or '-'
        dok_no = self.form_data.get('dok_no', '') or '-'
        rev_no = self.form_data.get('rev_no', '') or '-'

        create_cell(1, 4, "TARİH", bg='#fff176', font=('Arial', 11, 'bold'), anchor='center')
        create_cell(1, 5, tarih, bg='#e3f2fd', font=('Arial', 11), anchor='center')
        create_cell(2, 4, "DOK.NO", bg='#fff176', font=('Arial', 11, 'bold'), anchor='center')
        create_cell(2, 5, dok_no, bg='#e3f2fd', font=('Arial', 11), anchor='center')
        create_cell(3, 4, "REV.NO/TRH", bg='#fff176', font=('Arial', 11, 'bold'), anchor='center')
        create_cell(3, 5, rev_no, bg='#e3f2fd', font=('Arial', 11), anchor='center')

        for filler_row in range(1, 4):
            create_cell(filler_row, 0, '', colspan=4, bg='white')

        # Personel ve mali bilgiler
        current_row = 4
        create_cell(current_row, 0, "GÖREVLİ PERSONEL", colspan=2, bg='#fff176', font=('Arial', 12, 'bold'))
        create_cell(current_row, 2, "AVANS TUTARI", colspan=2, bg='#fff176', font=('Arial', 12, 'bold'))
        create_cell(current_row, 4, "TAŞERON ŞİRKET", colspan=2, bg='#fff176', font=('Arial', 12, 'bold'))

        personel_values = [self.form_data.get(f'personel_{i}', '') for i in range(1, 6)]
        avans_value = self.form_data.get('avans', '') or '-'
        taseron_value = self.form_data.get('taseron', '') or '-'

        for index in range(5):
            row_index = current_row + 1 + index
            create_cell(row_index, 0, f"Personel {index + 1}", bg='#fffde7', font=('Arial', 10, 'bold'))
            create_cell(row_index, 1, personel_values[index] or '-', bg='#f5f5f5', font=('Arial', 10))

            if index == 0:
                create_cell(row_index, 2, 'Tutar', bg='#fffde7', font=('Arial', 10, 'bold'))
                create_cell(row_index, 3, avans_value, bg='#f5f5f5', font=('Arial', 10))
                create_cell(row_index, 4, 'Şirket', bg='#fffde7', font=('Arial', 10, 'bold'))
                create_cell(row_index, 5, taseron_value, bg='#f5f5f5', font=('Arial', 10))
            else:
                create_cell(row_index, 2, '', bg='white')
                create_cell(row_index, 3, '', bg='white')
                create_cell(row_index, 4, '', bg='white')
                create_cell(row_index, 5, '', bg='white')

        current_row += 6
        create_cell(current_row, 0, '', colspan=6, bg='white', border=False, pady=4)

        current_row += 1
        create_cell(current_row, 0, "GÖREVİN TANIMI", colspan=4, bg='#fff176', font=('Arial', 12, 'bold'))
        create_cell(current_row, 4, "GÖREV YERİ", colspan=2, bg='#fff176', font=('Arial', 12, 'bold'))

        gorev_tanimi = self.form_data.get('gorev_tanimi', '') or '-'
        gorev_yeri = self.form_data.get('gorev_yeri', '') or '-'

        current_row += 1
        create_cell(current_row, 0, gorev_tanimi, colspan=4, bg='#f5f5f5', font=('Arial', 10), wrap=460)
        create_cell(current_row, 4, gorev_yeri, colspan=2, bg='#f5f5f5', font=('Arial', 10), wrap=220)

        current_row += 1
        create_cell(current_row, 0, '', colspan=6, bg='white', border=False, pady=4)

        current_row += 1
        create_cell(current_row, 0, "SEYAHAT / ÇALIŞMA BİLGİLERİ", colspan=6, bg='#fff176', font=('Arial', 12, 'bold'))

        def pair_row(row, label1, value1, label2='', value2=''):
            create_cell(row, 0, label1, bg='#fffde7', font=('Arial', 10, 'bold'))
            create_cell(row, 1, value1 or '-', bg='#f5f5f5', font=('Arial', 10))
            if label2:
                create_cell(row, 2, label2, bg='#fffde7', font=('Arial', 10, 'bold'))
                create_cell(row, 3, value2 or '-', bg='#f5f5f5', font=('Arial', 10))
            else:
                create_cell(row, 2, '', bg='white')
                create_cell(row, 3, '', bg='white')
            create_cell(row, 4, '', bg='white')
            create_cell(row, 5, '', bg='white')

        zaman_satirlari = [
            (
                "Yola Çıkış Tarihi",
                self.form_data.get('yola_cikis_tarih', ''),
                "Yola Çıkış Saati",
                self.form_data.get('yola_cikis_saat', '')
            ),
            (
                "Dönüş Tarihi",
                self.form_data.get('donus_tarih', ''),
                "Dönüş Saati",
                self.form_data.get('donus_saat', '')
            ),
            (
                "Çalışma Başlangıç",
                self.form_data.get('calisma_baslangic_tarih', ''),
                "Başlangıç Saati",
                self.form_data.get('calisma_baslangic_saat', '')
            ),
            (
                "Çalışma Bitiş",
                self.form_data.get('calisma_bitis_tarih', ''),
                "Bitiş Saati",
                self.form_data.get('calisma_bitis_saat', '')
            ),
            (
                "Toplam Mola",
                f"{self.form_data.get('mola_suresi', '')} dakika" if self.form_data.get('mola_suresi') else '',
                '',
                ''
            ),
        ]

        for offset, zaman in enumerate(zaman_satirlari, start=1):
            pair_row(current_row + offset, *zaman)

        current_row += len(zaman_satirlari) + 1
        create_cell(current_row, 0, '', colspan=6, bg='white', border=False, pady=4)

        current_row += 1
        create_cell(current_row, 0, "ARAÇ PLAKA NO", colspan=3, bg='#fff176', font=('Arial', 12, 'bold'))
        create_cell(current_row, 3, "", colspan=3, bg='white', border=False)

        arac_plaka = self.form_data.get('arac_plaka', '') or '-'

        current_row += 1
        create_cell(current_row, 0, arac_plaka, colspan=3, bg='#f5f5f5', font=('Arial', 10))
        create_cell(current_row, 3, '', colspan=3, bg='white', border=False)

        current_row += 1
        create_cell(current_row, 0, '', colspan=6, bg='white', border=False, pady=4)

        current_row += 1
        create_cell(current_row, 0, "HAZIRLAYAN / GÖREVLENDİREN", colspan=3, bg='#fff176', font=('Arial', 12, 'bold'))
        create_cell(current_row, 3, "İMZA", colspan=3, bg='#fff176', font=('Arial', 12, 'bold'))

        hazirlayan = self.form_data.get('hazirlayan', '') or '-'

        current_row += 1
        create_cell(current_row, 0, hazirlayan, colspan=3, bg='#f5f5f5', font=('Arial', 10))
        create_cell(current_row, 3, '', colspan=3, bg='white', border=False)

        current_row += 1
        create_cell(current_row, 0, '', colspan=6, bg='white', border=False, pady=4)

        current_row += 1
        status_bg = '#4caf50' if status.is_complete else '#ff9800'
        status_fg = 'white'
        status_value_bg = '#c8e6c9' if status.is_complete else '#ffe082'
        status_value_fg = '#1b5e20' if status.is_complete else '#bf360c'

        create_cell(current_row, 0, "DURUM", colspan=2, bg=status_bg, fg=status_fg, font=('Arial', 12, 'bold'), anchor='center')
        create_cell(current_row, 2, status.code, colspan=4, bg=status_value_bg, fg=status_value_fg, font=('Arial', 12, 'bold'), anchor='center')

        canvas.pack(side="left", fill="both", expand=True, pady=10)
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
                        self.form_data[f'personel_{i+1}'] = combo.get().strip()

            # Avans ve Taşeron
            if hasattr(self, 'avans_entry') and self.avans_entry.winfo_exists():
                self.form_data['avans'] = self.avans_entry.get().strip()
            if hasattr(self, 'taseron_combo') and self.taseron_combo.winfo_exists():
                self.form_data['taseron'] = self.taseron_combo.get().strip()

            # Görev tanımı ve yeri
            if hasattr(self, 'gorev_tanimi_text') and self.gorev_tanimi_text.winfo_exists():
                self.form_data['gorev_tanimi'] = self.gorev_tanimi_text.get('1.0', 'end-1c')
            if hasattr(self, 'gorev_yeri_text') and self.gorev_yeri_text.winfo_exists():
                self.form_data['gorev_yeri'] = self.gorev_yeri_text.get('1.0', 'end-1c')

            # Saat bilgileri
            if hasattr(self, 'yola_cikis_tarih_entry'):
                if self.yola_cikis_tarih_entry.winfo_exists():
                    self.form_data['yola_cikis_tarih'] = self.yola_cikis_tarih_entry.get().strip()
                    saat = self.yola_cikis_saat_combo.get().strip()
                    dakika = self.yola_cikis_dakika_combo.get().strip()
                    self.form_data['yola_cikis_saat'] = f"{saat}:{dakika}".strip(':') if (saat or dakika) else ''

                    self.form_data['donus_tarih'] = self.donus_tarih_entry.get().strip()
                    saat = self.donus_saat_combo.get().strip()
                    dakika = self.donus_dakika_combo.get().strip()
                    self.form_data['donus_saat'] = f"{saat}:{dakika}".strip(':') if (saat or dakika) else ''

                    self.form_data['calisma_baslangic_tarih'] = self.calisma_baslangic_tarih_entry.get().strip()
                    saat = self.calisma_baslangic_saat_combo.get().strip()
                    dakika = self.calisma_baslangic_dakika_combo.get().strip()
                    self.form_data['calisma_baslangic_saat'] = f"{saat}:{dakika}".strip(':') if (saat or dakika) else ''

                    self.form_data['calisma_bitis_tarih'] = self.calisma_bitis_tarih_entry.get().strip()
                    saat = self.calisma_bitis_saat_combo.get().strip()
                    dakika = self.calisma_bitis_dakika_combo.get().strip()
                    self.form_data['calisma_bitis_saat'] = f"{saat}:{dakika}".strip(':') if (saat or dakika) else ''

                    self.form_data['mola_suresi'] = self.mola_suresi_spin.get().strip()

            # Araç ve hazırlayan
            if hasattr(self, 'arac_combo') and self.arac_combo.winfo_exists():
                self.form_data['arac_plaka'] = self.arac_combo.get().strip()
            if hasattr(self, 'hazirlayan_combo') and self.hazirlayan_combo.winfo_exists():
                self.form_data['hazirlayan'] = self.hazirlayan_combo.get().strip()

        except Exception as e:
            print(f"Veri toplama hatası: {e}")

    def save_partial_form(self):
        """Kısmi formu kaydet (Görev Yeri'ne kadar)"""
        self.collect_form_data()
        try:
            filename, status = form_service.save_partial_form(self.form_no, self.form_data)
        except FormServiceError as exc:
            messagebox.showerror("Hata", str(exc))
            return

        self.form_data['durum'] = status.code

        messagebox.showinfo(
            "Başarılı",
            f"Form oluşturuldu!\n\nForm No: {self.form_no}\nDosya: {filename}\n\nGörev tamamlandığında 'GÖREV FORMU ÇAĞIR' ile bu formu açıp kalan kısımları doldurun."
        )

        self.show_main_menu()

    def save_form(self, stay_on_step=False):
        """Formu kaydet"""
        self.collect_form_data()
        try:
            filename, status = form_service.save_form(self.form_no, self.form_data)
        except FormServiceError as exc:
            messagebox.showerror("Hata", str(exc))
            return

        self.form_data['durum'] = status.code

        if stay_on_step:
            messagebox.showinfo(
                "Kaydedildi",
                f"Form {status.code} olarak kaydedildi.\n\nForm No: {self.form_no}\nDosya: {filename}"
            )
        else:
            messagebox.showinfo(
                "Başarılı",
                f"Form {status.code} olarak kaydedildi!\n\nForm No: {self.form_no}\nDosya: {filename}"
            )
            self.show_main_menu()

    def add_navigation_buttons(self, readonly=False, canvas_parent=False):
        """Navigasyon butonları ekle"""
        if self.nav_frame is not None and self.nav_frame.winfo_exists():
            self.nav_frame.destroy()

        parent = self.main_frame

        btn_frame = tk.Frame(parent, bg='white')
        pady = 20 if canvas_parent else 30
        btn_frame.pack(side='bottom', pady=pady, fill='x')

        # Butonları ortala
        center_frame = tk.Frame(btn_frame, bg='white')
        center_frame.pack(anchor='center', pady=10)

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

        btn_kaydet = tk.Button(
            center_frame,
            text="💾 Kaydet",
            font=('Arial', 13, 'bold'),
            bg='#4caf50',
            fg='white',
            width=15,
            height=2,
            command=self.handle_save_button,
            cursor='hand2',
            relief='raised',
            bd=3
        )
        btn_kaydet.pack(side='left', padx=15)

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

        self.nav_frame = btn_frame

    def handle_save_button(self):
        """Kaydet butonu davranışı"""
        self.collect_form_data()
        self.save_form(stay_on_step=False)

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

        self.current_step -= 1
        self.show_step()


if __name__ == "__main__":
    root = tk.Tk()
    app = GorevFormuApp(root)
    root.mainloop()
