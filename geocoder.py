import pandas as pd
from geopy.geocoders import Nominatim, Yandex
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import time
import re
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

class GeoCoderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Геокодер адресов v2.2")
        self.root.geometry("750x500")
        
        # Переменные
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.api_choice = tk.StringVar(value="osm")
        self.yandex_api_key = tk.StringVar()
        self.status = tk.StringVar(value="Готов к работе")
        self.address_column = tk.StringVar(value="Адрес") 
        self.city_column = tk.StringVar(value="Город")
        
        # GUI Элементы
        self.create_widgets()
    
    def create_widgets(self):
        style = ttk.Style()
        style.configure("TButton", padding=6, font=('Arial', 10))
        style.configure("TLabel", font=('Arial', 10))
        
        # Фрейм для загрузки файла
        file_frame = ttk.LabelFrame(self.root, text="1. Выберите файл Excel", padding=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(file_frame, text="Файл с адресами:").grid(row=0, column=0, sticky="w")
        self.input_entry = ttk.Entry(file_frame, textvariable=self.input_file, width=50)
        self.input_entry.grid(row=0, column=1, padx=5)
        self.browse_btn = ttk.Button(file_frame, text="Обзор...", command=self.browse_input_file)
        self.browse_btn.grid(row=0, column=2)
        
        ttk.Label(file_frame, text="Колонка с адресами:").grid(row=1, column=0, sticky="w", pady=(10,0))
        self.addr_col_entry = ttk.Entry(file_frame, textvariable=self.address_column, width=15)
        self.addr_col_entry.grid(row=1, column=1, sticky="w", padx=5)
        
        ttk.Label(file_frame, text="Колонка с городом:").grid(row=2, column=0, sticky="w", pady=(5,0))
        self.city_col_entry = ttk.Entry(file_frame, textvariable=self.city_column, width=15)
        self.city_col_entry.grid(row=2, column=1, sticky="w", padx=5)
        
        # Фрейм для настроек API
        api_frame = ttk.LabelFrame(self.root, text="2. Настройки геокодирования", padding=10)
        api_frame.pack(fill="x", padx=10, pady=5)
        
        self.osm_radio = ttk.Radiobutton(api_frame, text="OpenStreetMap (бесплатно, медленно)", 
                       variable=self.api_choice, value="osm")
        self.osm_radio.grid(row=0, column=0, sticky="w", pady=2)
        self.yandex_radio = ttk.Radiobutton(api_frame, text="Яндекс.Геокодер (платно, быстро)", 
                       variable=self.api_choice, value="yandex")
        self.yandex_radio.grid(row=1, column=0, sticky="w", pady=2)
        ttk.Label(api_frame, text="API-ключ Яндекс:").grid(row=2, column=0, sticky="w", pady=(10,0))
        self.api_entry = ttk.Entry(api_frame, textvariable=self.yandex_api_key, width=50)
        self.api_entry.grid(row=3, column=0, columnspan=2, sticky="we")
        
        # виджет для ошибок и исправлений
        settings_frame = ttk.Frame(self.root)
        settings_frame.pack(fill="x", padx=10, pady=5)

        self.auto_mode = tk.BooleanVar(value=True)
        self.correction_check = ttk.Checkbutton(settings_frame, text="Ручная коррекция проблемных адресов", 
                   variable=self.auto_mode)
        self.correction_check.pack(side="left")

        # Фрейм для запуска
        action_frame = ttk.Frame(self.root)
        action_frame.pack(fill="x", padx=10, pady=10)
        
        self.start_btn = ttk.Button(action_frame, text="Начать обработку", command=self.run_geocoding)
        self.start_btn.pack(side="left", padx=5)
        self.help_btn = ttk.Button(action_frame, text="Справка", command=self.show_help)
        self.help_btn.pack(side="right", padx=5)
        
        # Статус бар
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill="x", padx=10, pady=5)
        
        self.status_label = ttk.Label(status_frame, textvariable=self.status)
        self.status_label.pack(side="left")
        self.progress = ttk.Progressbar(status_frame, mode="determinate")
        self.progress.pack(fill="x", expand=True, padx=10)
        
        # Список виджетов, которые нужно отключать при обработке
        self.interactive_widgets = [
            self.input_entry, self.browse_btn, 
            self.addr_col_entry, self.city_col_entry,
            self.osm_radio, self.yandex_radio,
            self.api_entry, self.correction_check,
            self.start_btn, self.help_btn
        ]
    
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            base, ext = os.path.splitext(filename)
            self.output_file.set(f"{base}_with_coordinates{ext}")
    
    def get_geocoder(self):
        if self.api_choice.get() == "yandex":
            if not self.yandex_api_key.get():
                raise ValueError("Не указан API-ключ Яндекс")
            return Yandex(api_key=self.yandex_api_key.get())
        return Nominatim(user_agent="geoapi_excel")
    
    def preprocess_address(self, address, city):
        # Добавить новые шаблоны
        new_cases = {
            r"ТК\s?«([^»]+)»": r"торговый комплекс \1",
            r"ЖК\s?«([^»]+)»": r"жилой комплекс \1",
            r"мкр\.\s*(\d+[А-Я]?)": r"микрорайон \1",
            r"корп\.\s*(\d+)": r"корпус \1",
            "ЦУМ": "центральный универмаг"
        }
        for pattern, replacement in new_cases.items():
            address = re.sub(pattern, replacement, address)
        
        # Стандартизация городов
        city_mapping = {
            "ХМАО": "Ханты-Мансийск",
            "ЯНАО": "Салехард",
            "ЛНР": "Луганск"
        }
        city = city_mapping.get(city, city)

        """Предварительная обработка адреса"""
        # Специальные случаи (ручная корректировка)
        special_cases = {
            "мкр.1 (здание узла связи, ТЦ «МЕББЕРИ»)": f"{city}, микрорайон 1, торговый центр МЕББЕРИ",
            "д.55а": f"Селятино, дом 55а",
            "71-й километр МКАД, д.16А": "Москва, 71 км МКАД, дом 16А",
            "рынок": f"{city}, рынок",
            "ТЦ «д.торговли»": f"{city}, торговый центр",
            "пав. 27": f"{city}, павильон 27",
            "кв-л 2-й, д.7": f"{city}, квартал 2, дом 7",
            "мкр. 1, д.7А": f"{city}, микрорайон 1, дом 7А",
            "ст. Полтавская Красная, д.121": "станица Полтавская, улица Красная, дом 121"
        }
        
        # Проверка специальных случаев
        for orig, replacement in special_cases.items():
            if orig in address:
                return replacement
        
        # Удаление технических пометок в скобках
        address = re.sub(r'\([^)]*\)', '', address)
        
        # Удаление кавычек и лишних символов
        address = re.sub(r'["«»]', '', address).strip()
        
        # Добавление города, если отсутствует
        if city and city not in address:
            address = f"{city}, {address}"
        
        # Стандартизация сокращений
        replacements = {
            r'\bул\.': 'улица',
            r'\bд\.': 'дом',
            r'\bк\.': 'корпус',
            r'\bстр\.': 'строение',
            r'\bпав\.': ' павильон',
            r'\bмкр\.': 'микрорайон',
            r'\bпр-т\b': 'проспект',
            r'\bпр\.': 'проспект',
            r'\bпер\.': 'переулок',
            r'\bш\.': 'шоссе',
            r'\bпл\.': 'площадь',
            r'\bб-р': 'бульвар',
            r'\bкв-л': 'квартал',
            r'\bст\.': 'станица',
            r'\brp\.': 'рабочий поселок',
            r'\bТЦ\b': 'торговый центр',
            r'\bТРЦ\b': 'торгово-развлекательный центр',
            r'\bТК\b': 'торговый комплекс',
            r'\bЖК\b': 'жилой комплекс'
        }
        
        for pattern, replacement in replacements.items():
            address = re.sub(pattern, replacement, address)
        
        return address
    
    def geocode_address(self, address, geocoders, retries=3):
        """Геокодирование с использованием нескольких сервисов"""
        error_msg = "Адрес не найден"
        for geocoder in geocoders:
            if geocoder is None:
                continue
                
            for attempt in range(retries):
                try:
                    location = geocoder.geocode(address, timeout=15)
                    if location:
                        return location.latitude, location.longitude, str(geocoder).split('.')[-1].split(' ')[0], ""
                    time.sleep(0.5)
                except (GeocoderTimedOut, GeocoderServiceError) as e:
                    error_msg = f"Таймаут сервиса {geocoder}: {str(e)}"
                    if attempt < retries - 1:
                        time.sleep(2)
                        continue
                except Exception as e:
                    error_msg = f"Критическая ошибка: {str(e)}"
                    break
        return None, None, None, error_msg
    
    def set_widgets_state(self, state):
        """Устанавливает состояние интерактивных виджетов"""
        for widget in self.interactive_widgets:
            widget.config(state=state)
        self.root.config(cursor="watch" if state == "disabled" else "")
    
    def run_geocoding(self):
        if not self.input_file.get():
            messagebox.showerror("Ошибка", "Выберите файл для обработки!")
            return
        
        # Блокировка интерактивных элементов
        self.set_widgets_state("disabled")
        self.status.set("Начало обработки...")
        self.root.update()
        
        try:
            self.status.set("Чтение файла...")
            self.root.update()
            
            try:
                df = pd.read_excel(self.input_file.get())
            except Exception as e:
                messagebox.showerror("Ошибка чтения", f"Не удалось прочитать файл:\n{str(e)}")
                return
            
            # Проверка наличия необходимых колонок
            required_columns = [self.address_column.get(), self.city_column.get()]
            for col in required_columns:
                if col not in df.columns:
                    messagebox.showerror("Ошибка", 
                        f"Колонка '{col}' не найдена в файле!\nДоступные колонки: {list(df.columns)}")
                    return
            
            #остальной код геокодера доступен после покупки: tg: @@hardstone_national_anthem  | kwork: https://kwork.ru/user/dmitriynehaev
            #остальной код геокодера доступен после покупки: tg: @@hardstone_national_anthem  | kwork: https://kwork.ru/user/dmitriynehaev
            #остальной код геокодера доступен после покупки: tg: @@hardstone_national_anthem  | kwork: https://kwork.ru/user/dmitriynehaev
            #остальной код геокодера доступен после покупки: tg: @@hardstone_national_anthem  | kwork: https://kwork.ru/user/dmitriynehaev
            #остальной код геокодера доступен после покупки: tg: @@hardstone_national_anthem  | kwork: https://kwork.ru/user/dmitriynehaev
            #остальной код геокодера доступен после покупки: tg: @@hardstone_national_anthem  | kwork: https://kwork.ru/user/dmitriynehaev
            
            print(f"Размер исходных данных: {len(df)}")
            print(f"Обработано записей: {len(results)}")

            # Принудительное обрезание DataFrame
            df = df.reset_index(drop=True).iloc[:len(results)]
            total = len(results)
            
            df["Обработанный адрес"] = [r[2] for r in results]
            df["Широта"] = [r[0] for r in results]
            df["Долгота"] = [r[1] for r in results]
            
            # Сохранение результатов
            output_path = self.output_file.get()
            df.to_excel(output_path, index=False)

            assert len(df) == len(results), "Ошибка: Размеры данных не совпадают!"
            
            # Сохранение лога ошибок
            if failed_addresses:
                failed_df = pd.DataFrame(failed_addresses)
                error_path = output_path.replace(".xlsx", "_errors.xlsx")
                failed_df.to_excel(error_path, index=False)
            
            # Сохранение лога ручных исправлений
            if manual_fixes:
                fixes_df = pd.DataFrame(manual_fixes)
                fixes_path = output_path.replace(".xlsx", "_manual_fixes.xlsx")
                fixes_df.to_excel(fixes_path, index=False)
            
            # Статистика
            success_count = len([r for r in results if r[0] is not None])
            failed_count = total - success_count
            
            self.status.set(f"Готово! Успешно: {success_count}, Ошибки: {failed_count}")
            messagebox.showinfo("Результат", 
                f"Обработано {total} адресов\n"
                f"Успешно: {success_count}\n"
                f"Не найдено: {failed_count}\n\n"
                f"Основной файл: {output_path}\n"
                f"{'Лог ошибок: ' + error_path if failed_addresses else ''}"
                f"{'Лог ручных исправлений: ' + fixes_path if manual_fixes else ''}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
            self.status.set("Ошибка обработки")
        finally:
            self.progress["value"] = 0
            # Разблокировка интерфейса
            self.set_widgets_state("normal")
    
    def show_help(self):
        help_text = """Инструкция по улучшенному геокодированию v2.2:

1. Требования к входным данным:
   - Файл должен содержать колонку с адресом
   - Файл должен содержать колонку с городом
   - Адреса должны быть максимально полными

2. Особенности обработки:
   - Автоматическое добавление города к адресу
   - Стандартизация сокращений (ул., д., к. и т.д.)
   - Удаление технических пометок в скобках
   - Ручная корректировка сложных адресов
   - Использование нескольких геокодеров (Яндекс + OSM)

3. Результаты:
   - Основной файл с координатами
   - Файл с ошибками для проблемных адресов
   - Статистика успешно обработанных адресов

4. Рекомендации:
   - Для лучших результатов используйте Яндекс.Геокодер
   - Проверяйте и корректируйте адреса в файле ошибок
   - Для зарубежных адресов указывайте страну в колонке города

Для Яндекс.Геокодера:
- Получите ключ: https://developer.tech.yandex.ru/
- Вставьте его в поле API-ключа"""
        messagebox.showinfo("Справка", help_text)

class AddressFixDialog(tk.Toplevel):
    def __init__(self, parent, row_num, address, city, processed_address, error):
        super().__init__(parent)
        self.title(f"Проблема с адресом (строка {row_num})")
        self.geometry("700x580")
        self.result = None
        
        # Основные элементы
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # Информация о строке
        ttk.Label(main_frame, text=f"Строка №: {row_num}", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", pady=5)
        
        # Оригинальные данные
        ttk.Label(main_frame, text="Оригинальный адрес:").grid(row=1, column=0, sticky="w")
        ttk.Label(main_frame, text=address, wraplength=600).grid(row=2, column=0, sticky="w", padx=10)
        
        ttk.Label(main_frame, text="Оригинальный город:").grid(row=3, column=0, sticky="w", pady=(10,0))
        ttk.Label(main_frame, text=city).grid(row=4, column=0, sticky="w", padx=10)
        
        # Обработанный адрес
        ttk.Label(main_frame, text="Обработанный адрес:").grid(row=5, column=0, sticky="w", pady=(10,0))
        ttk.Label(main_frame, text=processed_address, wraplength=600, foreground="#555").grid(row=6, column=0, sticky="w", padx=10)
        
        # Ошибка
        ttk.Label(main_frame, text="Ошибка:").grid(row=7, column=0, sticky="w", pady=(10,0))
        ttk.Label(main_frame, text=error, wraplength=600, foreground="red").grid(row=8, column=0, sticky="w", padx=10)
        
        # Поля для исправления
        ttk.Label(main_frame, text="Исправьте адрес:", font=('Arial', 9, 'bold')).grid(row=9, column=0, sticky="w", pady=(15,0))
        self.address_entry = ttk.Entry(main_frame, width=80)
        self.address_entry.grid(row=10, column=0, sticky="we", padx=10, pady=5)
        self.address_entry.insert(0, address)
        
        ttk.Label(main_frame, text="Исправьте город:", font=('Arial', 9, 'bold')).grid(row=11, column=0, sticky="w", pady=(5,0))
        self.city_entry = ttk.Entry(main_frame, width=80)
        self.city_entry.grid(row=12, column=0, sticky="we", padx=10, pady=5)
        self.city_entry.insert(0, city)
        
        # Кнопки
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=13, column=0, pady=15)
        
        ttk.Button(btn_frame, text="Повторить попытку", command=self.retry).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Пропустить", command=self.skip).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Отменить обработку", command=self.cancel).pack(side="left", padx=5)
        
        # Советы
        advice_frame = ttk.LabelFrame(main_frame, text="Советы по улучшению адреса", padding=10)
        advice_frame.grid(row=14, column=0, sticky="we", pady=10)
        advice_text = """• Убедитесь, что указаны улица и номер дома
• Проверьте правильность названия города
• Для торговых центров укажите 'ТЦ Название'
• Для новостроек используйте 'ЖК Название'
• Уберите технические пометки в скобках
• Добавьте регион, если город не уникален"""
        ttk.Label(advice_frame, text=advice_text, justify="left").pack(anchor="w")

    def retry(self):
        self.result = ("retry", self.address_entry.get(), self.city_entry.get())
        self.destroy()

    def skip(self):
        self.result = ("skip", None, None)
        self.destroy()

    def cancel(self):
        self.result = ("cancel", None, None)
        self.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = GeoCoderApp(root)
    root.mainloop()