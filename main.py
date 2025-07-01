#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Базовые импорты
import sys
import os
import json
import time
import random
import threading

from io import BytesIO

# Tkinter импорты
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import scrolledtext

# Сетевые библиотеки
import requests
from urllib.parse import quote_plus

# Обработка данных
import pandas as pd

# Обработка HTML
from bs4 import BeautifulSoup
import re

# Обработка изображений
try:
    from PIL import Image, ImageTk, ImageOps, ImageDraw
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("Pillow не установлен")

try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False
    print("OpenCV не установлен")

try:
    from io import BytesIO
    IO_AVAILABLE = True
except ImportError:
    IO_AVAILABLE = False

# Константы
USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
PROM_SEARCH_URL = 'https://prom.ua/ua/search?search_term={}'
DEPOS_SEARCH_URL = 'https://depositphotos.com/ru/stock-photos/{}.html'

class ImageProcessor:
    """Класс для обработки изображений"""
    
    def __init__(self):
        self.output_dir = "output"
        self.input_dir = "input"
        
    def create_directories(self):
        """Создание необходимых директорий"""
        try:
            os.makedirs(self.input_dir, exist_ok=True)
            os.makedirs(self.output_dir, exist_ok=True)
            return True
        except Exception as e:
            print(f"Ошибка создания директорий: {e}")
            return False
    
    def detect_main_object_smart(self, image):
        """Умное определение границ основного объекта на белом фоне"""
        if not CV2_AVAILABLE:
            return 0, 0, image.width, image.height
            
        try:
            # Преобразуем в numpy array
            img_array = np.array(image)
            
            if len(img_array.shape) == 3:
                gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            else:
                gray = img_array
            
            # Создаем маску для белого фона
            white_mask = gray > 240
            object_mask = ~white_mask
            
            # Морфологические операции
            kernel = np.ones((3, 3), np.uint8)
            object_mask = cv2.morphologyEx(object_mask.astype(np.uint8), cv2.MORPH_CLOSE, kernel)
            object_mask = cv2.morphologyEx(object_mask, cv2.MORPH_OPEN, kernel)
            
            # Находим контуры
            contours, _ = cv2.findContours(object_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            if not contours:
                return 0, 0, image.width, image.height
            
            # Объединяем все контуры
            all_points = np.vstack(contours)
            x, y, w, h = cv2.boundingRect(all_points)
            
            if w <= 0 or h <= 0:
                return 0, 0, image.width, image.height
            
            return x, y, w, h
            
        except Exception as e:
            print(f"Ошибка определения объекта: {e}")
            return 0, 0, image.width, image.height
    
    def smart_scale_object(self, image, target_ratio, margin_percent=0.1):
        """Умное масштабирование объекта"""
        try:
            obj_x, obj_y, obj_w, obj_h = self.detect_main_object_smart(image)
            
            if obj_w == 0 or obj_h == 0:
                obj_x, obj_y = 0, 0
                obj_w, obj_h = image.width, image.height
            
            if obj_w <= 0 or obj_h <= 0:
                return image
            
            margin_w = max(1, int(image.width * margin_percent))
            margin_h = max(1, int(image.height * margin_percent))
            
            available_w = max(1, image.width - 2 * margin_w)
            available_h = max(1, image.height - 2 * margin_h)
            
            scale_w = available_w / obj_w if obj_w > 0 else 1.0
            scale_h = available_h / obj_h if obj_h > 0 else 1.0
            
            scale = min(scale_w, scale_h, 3.0)
            
            if scale > 1.2:
                new_width = max(1, int(image.width * scale))
                new_height = max(1, int(image.height * scale))
                
                scaled_image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                return scaled_image
            
            return image
            
        except Exception as e:
            print(f"Ошибка умного масштабирования: {e}")
            return image
    
    def process_image(self, input_path, output_path, target_ratio=4/3, bg_color=(255, 255, 255), smart_scale=False):
        """Обработка одного изображения"""
        if not PIL_AVAILABLE:
            print("PIL не доступен для обработки изображений")
            return False
            
        try:
            if not os.path.exists(input_path):
                print(f"Файл не найден: {input_path}")
                return False
                
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            with Image.open(input_path) as img:
                if img.size[0] == 0 or img.size[1] == 0:
                    print(f"Некорректные размеры изображения: {img.size}")
                    return False
                
                # Обработка прозрачности
                if img.mode in ('RGBA', 'LA', 'P'):
                    background = Image.new('RGB', img.size, bg_color)
                    if img.mode == 'P':
                        img = img.convert('RGBA')
                    
                    if img.mode in ('RGBA', 'LA'):
                        try:
                            alpha_channel = img.split()[-1]
                            background.paste(img, mask=alpha_channel)
                        except Exception as e:
                            print(f"Ошибка обработки альфа-канала: {e}")
                            background.paste(img.convert('RGB'))
                    else:
                        background.paste(img.convert('RGB'))
                    img = background
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Умное масштабирование
                if smart_scale:
                    try:
                        img = self.smart_scale_object(img, target_ratio)
                    except Exception as e:
                        print(f"Ошибка умного масштабирования: {e}")
                
                current_width, current_height = img.size
                if current_width == 0 or current_height == 0:
                    print(f"Некорректные размеры после обработки: {img.size}")
                    return False
                    
                current_ratio = current_width / current_height
                
                if abs(current_ratio - target_ratio) < 0.01:
                    img.save(output_path, quality=95)
                    return True
                
                if current_ratio > target_ratio:
                    new_height = int(current_width / target_ratio)
                    padding_height = new_height - current_height
                    padding_top = padding_height // 2
                    
                    new_img = Image.new('RGB', (current_width, new_height), bg_color)
                    new_img.paste(img, (0, padding_top))
                    
                else:
                    new_width = int(current_height * target_ratio)
                    padding_width = new_width - current_width
                    padding_left = padding_width // 2
                    
                    new_img = Image.new('RGB', (new_width, current_height), bg_color)
                    new_img.paste(img, (padding_left, 0))
                
                output_ext = os.path.splitext(output_path)[1].lower()
                if output_ext == '.jpg':
                    output_ext = '.jpeg'
                
                if output_ext in ['.jpeg', '.jpg']:
                    new_img.save(output_path, 'JPEG', quality=95)
                else:
                    new_img.save(output_path, quality=95)
                
                return True
                
        except Exception as e:
            print(f"Ошибка при обработке {input_path}: {e}")
            return False

class ImageParserApp:
    """Главное приложение"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Обработчик изображений v2.0")
        self.root.geometry("900x700")
        
        # Проверяем доступность библиотек
        self.check_dependencies()
        
        # Устанавливаем иконку
        self.set_icon()
        
        self.processor = ImageProcessor()
        self.processor.create_directories()
        
        # Переменные
        self.parsed_results = []
        self.selected_images = {}
        self.excel_data = []
        self.parsing_thread = None
        self.stop_parsing_flag = False
        
        # Создаем интерфейс
        self.create_interface()
    
    def check_dependencies(self):
        """Проверка доступности библиотек"""
        status = "Статус библиотек:\n"
        status += f"PIL/Pillow: {'✓' if PIL_AVAILABLE else '✗'}\n"
        status += f"OpenCV: {'✓' if CV2_AVAILABLE else '✗'}\n"
        status += f"IO: {'✓' if IO_AVAILABLE else '✗'}\n"
        print(status)
        
        if not PIL_AVAILABLE:
            messagebox.showwarning("Предупреждение", 
                "PIL/Pillow не установлен.\nОбработка изображений будет недоступна.")
    
    def set_icon(self):
        """Установка иконки приложения"""
        try:
            if os.path.exists("icon.ico"):
                self.root.iconbitmap("icon.ico")
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")
    
    def create_interface(self):
        """Создание интерфейса приложения"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Создаем вкладки
        self.create_parsing_tab()
        self.create_viewer_tab()
        self.create_processing_tab()
        self.create_excel_tab()
    
    def create_parsing_tab(self):
        """Создание вкладки парсинга"""
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text="Парсинг изображений")
        
        # Основной фрейм
        main_frame = ttk.Frame(tab_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Загрузка файла Excel
        excel_frame = ttk.LabelFrame(main_frame, text="Загрузка товаров из Excel")
        excel_frame.pack(fill='x', pady=(0, 10))
        
        file_frame = ttk.Frame(excel_frame)
        file_frame.pack(fill='x', padx=10, pady=10)
        
        self.excel_path = tk.StringVar(value="target.xlsx")
        ttk.Entry(file_frame, textvariable=self.excel_path, width=50).pack(side='left', padx=(0, 5))
        ttk.Button(file_frame, text="Обзор", command=self.browse_excel_file).pack(side='left', padx=(0, 5))
        ttk.Button(file_frame, text="Загрузить", command=self.load_excel_data).pack(side='left')
        
        self.items_info = tk.StringVar(value="Товары не загружены")
        ttk.Label(excel_frame, textvariable=self.items_info).pack(padx=10, pady=(0, 10))
        
        # Настройки парсинга
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки парсинга")
        settings_frame.pack(fill='x', pady=(0, 10))
        
        # Источники
        sources_frame = ttk.Frame(settings_frame)
        sources_frame.pack(fill='x', padx=10, pady=10)
        
        self.parse_prom = tk.BooleanVar(value=True)
        self.parse_depos = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(sources_frame, text="Prom.ua", variable=self.parse_prom).pack(side='left', padx=(0, 20))
        ttk.Checkbutton(sources_frame, text="Depositphotos.com", variable=self.parse_depos).pack(side='left')
        
        # Лимит изображений
        limit_frame = ttk.Frame(settings_frame)
        limit_frame.pack(fill='x', padx=10, pady=(0, 10))
        
        ttk.Label(limit_frame, text="Лимит изображений на товар:").pack(side='left')
        self.images_limit = tk.IntVar(value=10)
        ttk.Spinbox(limit_frame, from_=1, to=50, textvariable=self.images_limit, width=10).pack(side='left', padx=(10, 0))
        
        # Кнопки управления
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(buttons_frame, text="Начать парсинг", command=self.start_parsing).pack(side='left', padx=(0, 10))
        ttk.Button(buttons_frame, text="Остановить", command=self.stop_parsing).pack(side='left', padx=(0, 10))
        ttk.Button(buttons_frame, text="Очистить", command=self.clear_results).pack(side='left')
        
        # Прогресс
        self.progress_var = tk.StringVar(value="Готов к работе")
        ttk.Label(main_frame, textvariable=self.progress_var).pack(anchor='w', pady=(10, 5))
        
        self.progress_bar = ttk.Progressbar(main_frame, length=400, mode='determinate')
        self.progress_bar.pack(fill='x', pady=(0, 10))
        
        # Лог
        ttk.Label(main_frame, text="Лог парсинга:").pack(anchor='w')
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15)
        self.log_text.pack(fill='both', expand=True, pady=(5, 0))
    
    def create_viewer_tab(self):
        """Создание вкладки просмотра изображений"""
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text="Просмотр изображений")
        
        # Заголовок
        header_frame = ttk.Frame(tab_frame)
        header_frame.pack(fill='x', padx=10, pady=10)
        
        self.viewer_stats = tk.StringVar(value="Нет данных для просмотра")
        ttk.Label(header_frame, textvariable=self.viewer_stats, font=("Arial", 12, "bold")).pack(side='left')
        
        # Кнопки
        buttons_frame = ttk.Frame(header_frame)
        buttons_frame.pack(side='right')
        
        ttk.Button(buttons_frame, text="Загрузить JSON", command=self.load_json_results).pack(side='left', padx=(0, 5))
        ttk.Button(buttons_frame, text="Экспорт выбранных", command=self.export_selected).pack(side='left', padx=(0, 5))
        ttk.Button(buttons_frame, text="Очистить выбор", command=self.clear_selection).pack(side='left')
        
        # Область просмотра
        self.viewer_canvas = tk.Canvas(tab_frame)
        self.viewer_scrollbar = ttk.Scrollbar(tab_frame, orient="vertical", command=self.viewer_canvas.yview)
        self.viewer_frame = ttk.Frame(self.viewer_canvas)
        
        self.viewer_frame.bind("<Configure>", 
            lambda e: self.viewer_canvas.configure(scrollregion=self.viewer_canvas.bbox("all")))
        
        self.viewer_canvas.create_window((0, 0), window=self.viewer_frame, anchor="nw")
        self.viewer_canvas.configure(yscrollcommand=self.viewer_scrollbar.set)
        
        self.viewer_canvas.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=5)
        self.viewer_scrollbar.pack(side="right", fill="y", pady=5)
        
        # Привязка колесика мыши
        self.viewer_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
    
    def create_processing_tab(self):
        """Создание вкладки обработки изображений"""
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text="Обработка изображений")
        
        main_frame = ttk.Frame(tab_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Основные настройки
        settings_frame = ttk.LabelFrame(main_frame, text="Основные настройки")
        settings_frame.pack(fill='x', pady=(0, 10))
        
        # Соотношение сторон
        ratio_frame = ttk.Frame(settings_frame)
        ratio_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(ratio_frame, text="Соотношение сторон:").pack(side='left')
        self.ratio_var = tk.StringVar(value="4:3")
        ratio_combo = ttk.Combobox(ratio_frame, textvariable=self.ratio_var, 
                                 values=["1:1", "4:3", "3:2", "16:9", "3:4", "2:3", "9:16"], width=10)
        ratio_combo.pack(side='left', padx=(10, 0))
        
        # Пользовательское соотношение
        custom_frame = ttk.Frame(ratio_frame)
        custom_frame.pack(side='left', padx=(20, 0))
        
        ttk.Label(custom_frame, text="Или:").pack(side='left')
        ttk.Label(custom_frame, text="Ширина:").pack(side='left', padx=(10, 0))
        self.custom_width = tk.IntVar(value=4)
        ttk.Spinbox(custom_frame, from_=1, to=100, textvariable=self.custom_width, width=5).pack(side='left', padx=(5, 0))
        
        ttk.Label(custom_frame, text="Высота:").pack(side='left', padx=(10, 0))
        self.custom_height = tk.IntVar(value=3)
        ttk.Spinbox(custom_frame, from_=1, to=100, textvariable=self.custom_height, width=5).pack(side='left', padx=(5, 0))
        
        # Цвет фона
        color_frame = ttk.Frame(settings_frame)
        color_frame.pack(fill='x', padx=10, pady=(0, 10))
        
        ttk.Label(color_frame, text="Цвет фона:").pack(side='left')
        self.bg_color = tk.StringVar(value="255,255,255")
        ttk.Entry(color_frame, textvariable=self.bg_color, width=15).pack(side='left', padx=(10, 0))
        ttk.Label(color_frame, text="(R,G,B: 0-255)").pack(side='left', padx=(5, 0))
        
        # Источник изображений
        source_frame = ttk.LabelFrame(main_frame, text="Источник изображений")
        source_frame.pack(fill='x', pady=(0, 10))
        
        self.source_mode = tk.StringVar(value="json")
        ttk.Radiobutton(source_frame, text="Папка input", variable=self.source_mode, value="folder").pack(anchor='w', padx=10, pady=5)
        ttk.Radiobutton(source_frame, text="Преобразовать в JSON и загрузить", variable=self.source_mode, value="json").pack(anchor='w', padx=10, pady=5)
        
        # Опции обработки
        options_frame = ttk.LabelFrame(main_frame, text="Опции обработки изображений")
        options_frame.pack(fill='x', pady=(0, 10))
        
        self.processing_mode = tk.StringVar(value="smart")
        ttk.Radiobutton(options_frame, text="Стандартная обработка", variable=self.processing_mode, value="standard").pack(anchor='w', padx=10, pady=5)
        ttk.Radiobutton(options_frame, text="Специальный режим объектов на белом фоне", variable=self.processing_mode, value="smart").pack(anchor='w', padx=10, pady=5)
        
        # Кнопки управления
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(buttons_frame, text="Начать обработку", command=self.start_processing).pack(side='left', padx=(0, 10))
        ttk.Button(buttons_frame, text="Открыть папку input", command=lambda: self.open_folder("input")).pack(side='left', padx=(0, 10))
        ttk.Button(buttons_frame, text="Открыть папку output", command=lambda: self.open_folder("output")).pack(side='left')
        
        # Прогресс обработки
        self.processing_progress_var = tk.StringVar(value="Готов к обработке")
        ttk.Label(main_frame, textvariable=self.processing_progress_var).pack(anchor='w', pady=(10, 5))
        
        self.processing_progress_bar = ttk.Progressbar(main_frame, length=400, mode='determinate')
        self.processing_progress_bar.pack(fill='x', pady=(0, 10))
        
        # Лог обработки
        ttk.Label(main_frame, text="Лог обработки:").pack(anchor='w')
        self.processing_log = scrolledtext.ScrolledText(main_frame, height=10)
        self.processing_log.pack(fill='both', expand=True, pady=(5, 0))
    
    def create_excel_tab(self):
        """Создание вкладки генерации Excel"""
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text="Генерация Excel")
        
        main_frame = ttk.Frame(tab_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Настройки
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки")
        settings_frame.pack(fill='x', pady=(0, 10))
        
        # Папка для анализа
        folder_frame = ttk.Frame(settings_frame)
        folder_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(folder_frame, text="Папка для анализа:").pack(side='left')
        self.excel_folder = tk.StringVar(value="output")
        ttk.Entry(folder_frame, textvariable=self.excel_folder, width=30).pack(side='left', padx=(10, 5))
        ttk.Button(folder_frame, text="Обзор", command=self.browse_excel_folder).pack(side='left')
        
        # Базовый путь
        path_frame = ttk.Frame(settings_frame)
        path_frame.pack(fill='x', padx=10, pady=(0, 10))
        
        ttk.Label(path_frame, text="Базовый путь:").pack(side='left')
        self.base_path = tk.StringVar(value=os.path.abspath("output"))
        ttk.Entry(path_frame, textvariable=self.base_path, width=50).pack(side='left', padx=(10, 0))
        
        # Кнопки
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(buttons_frame, text="Создать Excel файл", command=self.create_excel_file).pack(side='left', padx=(0, 10))
        ttk.Button(buttons_frame, text="Открыть папку", command=lambda: self.open_folder(self.excel_folder.get())).pack(side='left')
        
        # Результат
        result_frame = ttk.LabelFrame(main_frame, text="Результат")
        result_frame.pack(fill='both', expand=True, pady=(0, 0))
        
        self.excel_result = scrolledtext.ScrolledText(result_frame, height=15)
        self.excel_result.pack(fill='both', expand=True, padx=10, pady=10)
    
    # Основные методы
    def browse_excel_file(self):
        """Выбор Excel файла"""
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
    
    def load_excel_data(self):
        """Загрузка данных из Excel файла"""
        try:
            excel_file = self.excel_path.get()
            if not excel_file or not os.path.exists(excel_file):
                messagebox.showerror("Ошибка", f"Файл {excel_file} не найден!")
                return
            
            df = pd.read_excel(excel_file)
            
            if df.empty:
                messagebox.showerror("Ошибка", "Excel файл пуст!")
                return
            
            # Поиск столбцов
            columns_map = {}
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if col_lower in ['артикул', 'article', 'арт']:
                    columns_map[col] = 'article'
                elif col_lower in ['назва', 'название', 'name', 'наименование']:
                    columns_map[col] = 'name'
            
            if columns_map:
                df = df.rename(columns=columns_map)
            
            if 'article' not in df.columns or 'name' not in df.columns:
                available_columns = ", ".join(df.columns)
                messagebox.showerror("Ошибка", 
                    f"В Excel-файле не найдены столбцы 'Артикул' и 'Название'\n"
                    f"Доступные столбцы: {available_columns}")
                return
            
            self.excel_data = []
            for _, row in df.iterrows():
                try:
                    article = str(row['article']).strip() if pd.notna(row['article']) else ""
                    name = str(row['name']).strip() if pd.notna(row['name']) else ""
                    
                    if article and name and article != 'nan' and name != 'nan':
                        self.excel_data.append({
                            'article': article,
                            'name': name
                        })
                except Exception as e:
                    print(f"Ошибка обработки строки: {e}")
                    continue
            
            if not self.excel_data:
                messagebox.showwarning("Предупреждение", "Не найдено валидных записей в Excel файле!")
                return
            
            self.items_info.set(f"Загружено товаров: {len(self.excel_data)}")
            self.log_message(f"Успешно загружено {len(self.excel_data)} товаров из {excel_file}")
            
        except Exception as e:
            error_msg = f"Ошибка при загрузке файла: {str(e)}"
            messagebox.showerror("Ошибка", error_msg)
            self.log_message(f"ОШИБКА: {error_msg}")
    
    def start_parsing(self):
        """Запуск парсинга"""
        if not self.excel_data:
            messagebox.showwarning("Предупреждение", "Сначала загрузите данные из Excel файла!")
            return
        
        if not self.parse_prom.get() and not self.parse_depos.get():
            messagebox.showwarning("Предупреждение", "Выберите хотя бы один источник для парсинга!")
            return
        
        self.stop_parsing_flag = False
        self.parsing_thread = threading.Thread(target=self.parse_images)
        self.parsing_thread.daemon = True
        self.parsing_thread.start()
    
    def stop_parsing(self):
        """Остановка парсинга"""
        self.stop_parsing_flag = True
        self.progress_var.set("Остановка...")
    
    def clear_results(self):
        """Очистка результатов"""
        self.parsed_results = []
        self.log_text.delete(1.0, tk.END)
        self.progress_bar['value'] = 0
        self.progress_var.set("Результаты очищены")
    
    def parse_images(self):
        """Основной метод парсинга"""
        try:
            self.parsed_results = []
            total_items = len(self.excel_data)
            
            if self.parse_prom.get():
                self.progress_var.set("Парсинг Prom.ua...")
                prom_results = self.parse_source(self.excel_data, "prom", total_items)
                self.parsed_results.extend(prom_results)
            
            if self.parse_depos.get() and not self.stop_parsing_flag:
                self.progress_var.set("Парсинг Depositphotos...")
                depos_results = self.parse_source(self.excel_data, "depos", total_items)
                self.parsed_results.extend(depos_results)
            
            if not self.stop_parsing_flag:
                timestamp = int(time.time())
                filename = f"parsed_results_{timestamp}.json"
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.parsed_results, f, ensure_ascii=False, indent=2)
                
                self.progress_var.set(f"Парсинг завершен! Результаты сохранены в {filename}")
                self.log_message(f"Парсинг завершен. Найдено изображений: {sum(len(r['images']) for r in self.parsed_results)}")
                
                self.load_parsed_results()
            else:
                self.progress_var.set("Парсинг остановлен пользователем")
                
        except Exception as e:
            self.progress_var.set(f"Ошибка парсинга: {e}")
            self.log_message(f"ОШИБКА: {e}")
    
    def parse_source(self, items, source, total_items):
        """Парсинг одного источника"""
        results = []
        
        for i, item in enumerate(items):
            if self.stop_parsing_flag:
                break
            
            article = item['article']
            name = item['name']
            
            self.log_message(f"[{i+1}/{total_items}] Обработка: [{article}] {name}")
            
            time.sleep(random.uniform(1.5, 3.0))
            
            try:
                if source == "prom":
                    html = self.get_html_prom(name)
                    images = self.extract_images_prom(html)
                    source_name = "Prom.ua"
                else:
                    html = self.get_html_depos(name)
                    images = self.extract_images_depos(html)
                    source_name = "Depositphotos.com"
                
                images = images[:self.images_limit.get()]
                
                result = {
                    'article': article,
                    'name': name,
                    'source': source_name,
                    'images': images
                }
                
                results.append(result)
                self.log_message(f"  Найдено изображений: {len(images)}")
                
                progress = ((i + 1) / total_items) * 100
                self.progress_bar['value'] = progress
                self.root.update_idletasks()
                
            except Exception as e:
                self.log_message(f"  ОШИБКА: {e}")
        
        return results
    
    def get_html_prom(self, query):
        """Получение HTML со страницы Prom.ua"""
        encoded_query = quote_plus(query)
        url = PROM_SEARCH_URL.format(encoded_query)
        
        headers = {
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
            'Referer': 'https://prom.ua/',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.text
    
    def get_html_depos(self, query):
        """Получение HTML со страницы Depositphotos"""
        encoded_query = quote_plus(query)
        url = DEPOS_SEARCH_URL.format(encoded_query)
        
        headers = {
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
            'Referer': 'https://depositphotos.com/',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.text
    
    def extract_images_prom(self, html):
        """Извлечение изображений с Prom.ua"""
        if not html:
            return []
        
        soup = BeautifulSoup(html, 'html.parser')
        images = []
        
        picture_tags = soup.find_all('picture', attrs={'data-qaid': True})
        
        for picture in picture_tags:
            img_tag = picture.find('img', attrs={'src': True})
            if img_tag and 'src' in img_tag.attrs:
                images.append(img_tag['src'])
        
        return images
    
    def extract_images_depos(self, html):
        """Извлечение изображений с Depositphotos"""
        if not html:
            return []
        
        soup = BeautifulSoup(html, 'html.parser')
        source_tags = soup.find_all('source')
        
        images = []
        for source in source_tags:
            if source.has_attr('srcset'):
                urls = re.findall(r'(https://[^\s]+)', source['srcset'])
                images.extend(urls)
        
        return images
    
    def log_message(self, message):
        """Добавление сообщения в лог"""
        self.root.after(0, lambda: self._append_log(message))
    
    def _append_log(self, message):
        """Внутренний метод для добавления в лог"""
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
    
    # Методы для просмотра изображений
    def load_json_results(self):
        """Загрузка результатов из JSON файла"""
        filename = filedialog.askopenfilename(
            title="Выберите JSON файл с результатами",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    self.parsed_results = json.load(f)
                self.load_parsed_results()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при загрузке файла: {e}")
    
    def load_parsed_results(self):
        """Загрузка результатов парсинга в просмотрщик"""
        for widget in self.viewer_frame.winfo_children():
            widget.destroy()
        
        if not self.parsed_results:
            ttk.Label(self.viewer_frame, text="Нет данных для отображения").pack(pady=20)
            return
        
        total_images = sum(len(item['images']) for item in self.parsed_results)
        self.viewer_stats.set(f"Товаров: {len(self.parsed_results)} | Изображений: {total_images} | Выбрано: {len(self.selected_images)}")
        
        for i, item in enumerate(self.parsed_results):
            self.create_product_widget(item, i)
    
    # def create_product_widget(self, item, index):
    #     """Создание виджета для одного товара"""
    #     product_frame = ttk.Frame(self.viewer_frame)
    #     product_frame.pack(fill='x', padx=10, pady=5)
        
    #     info_frame = ttk.Frame(product_frame)
    #     info_frame.pack(fill='x', pady=5)
        
    #     selected_indicator = "✓" if item['article'] in self.selected_images else "○"
    #     info_text = f"{selected_indicator} {index+1}. [{item['article']}] {item['name']} ({item['source']})"
        
    #     info_label = ttk.Label(info_frame, text=info_text, font=("Arial", 10, "bold"))
    #     info_label.pack(side='left')
        
    #     if item['images']:
    #         select_btn = ttk.Button(info_frame, text="Выбрать первое", 
    #                               command=lambda: self.select_first_image(item))
    #         select_btn.pack(side='right')
        
    #     images_frame = ttk.Frame(product_frame)
    #     images_frame.pack(fill='x', pady=5)
        
    #     canvas = tk.Canvas(images_frame, height=160)
    #     h_scrollbar = ttk.Scrollbar(images_frame, orient="horizontal", command=canvas.xview)
        
    #     scroll_frame = ttk.Frame(canvas)
    #     scroll_frame.bind("<Configure>", 
    #         lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
    #     canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    #     canvas.configure(xscrollcommand=h_scrollbar.set)
        
    #     self.load_product_images(scroll_frame, item)
        
    #     canvas.pack(fill='x')
    #     h_scrollbar.pack(fill='x')
        
    #     ttk.Separator(product_frame, orient='horizontal').pack(fill='x', pady=10)

    def create_product_widget(self, item, index):
        """Создание виджета для одного товара"""
        # Основной фрейм товара
        product_frame = ttk.Frame(self.viewer_frame)
        product_frame.pack(fill='x', padx=10, pady=5)
        
        # Информация о товаре
        info_frame = ttk.Frame(product_frame)
        info_frame.pack(fill='x', pady=5)
        
        # Индикатор выбора
        selected_indicator = "✓" if item['article'] in self.selected_images else "○"
        info_text = f"{selected_indicator} {index+1}. [{item['article']}] {item['name']} ({item['source']})"
        
        info_label = ttk.Label(info_frame, text=info_text, font=("Arial", 10, "bold"))
        info_label.pack(side='left')
        
        # Основной контейнер с изображениями и выбранным изображением
        main_container = ttk.Frame(product_frame)
        main_container.pack(fill='both', expand=True, pady=5)
        
        # Левая часть - скролл с изображениями (80% ширины)
        left_frame = ttk.Frame(main_container)
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # Правая часть - выбранное изображение (фиксированная ширина)
        right_frame = ttk.Frame(main_container, width=200)
        right_frame.pack(side='right', fill='y')
        right_frame.pack_propagate(False)
        
        # Заголовок для выбранного изображения
        ttk.Label(right_frame, text="Выбранное:", font=("Arial", 9, "bold")).pack(pady=(0, 5))
        
        # Фрейм для выбранного изображения
        selected_frame = ttk.Frame(right_frame, relief='solid', borderwidth=1)
        selected_frame.pack(fill='both', expand=True)
        
        # Сохраняем ссылку на фреймы для обновления
        setattr(product_frame, 'selected_frame', selected_frame)
        setattr(product_frame, 'info_label', info_label)
        setattr(product_frame, 'article', item['article'])
        
        # Создаем Canvas для горизонтальной прокрутки изображений
        canvas = tk.Canvas(left_frame, height=160)
        h_scrollbar = ttk.Scrollbar(left_frame, orient="horizontal", command=canvas.xview)
        
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(xscrollcommand=h_scrollbar.set)
        
        # Загружаем изображения
        self.load_product_images(scroll_frame, item, product_frame)
        
        canvas.pack(fill='both', expand=True)
        h_scrollbar.pack(fill='x')
        
        # Показываем выбранное изображение, если есть
        if item['article'] in self.selected_images:
            self.update_selected_image_display(product_frame, self.selected_images[item['article']])
        
        # Разделитель
        ttk.Separator(product_frame, orient='horizontal').pack(fill='x', pady=10)





    
    def load_product_images(self, parent_frame, item, product_frame):
        """Загрузка изображений для товара"""
        for i, image_url in enumerate(item['images'][:10]):  # Ограничиваем для производительности
            try:
                img_frame = ttk.Frame(parent_frame)
                img_frame.pack(side='left', padx=5)
                
                # threading.Thread(target=self.load_single_image, 
                #                args=(img_frame, image_url, item, i), daemon=True).start()
                threading.Thread(target=self.load_single_image, 
                                args=(img_frame, image_url, item, i, product_frame), daemon=True).start()
                
            except Exception as e:
                print(f"Ошибка загрузки изображения {image_url}: {e}")
    
    # def load_single_image(self, parent_frame, image_url, item, img_index):
    def load_single_image(self, parent_frame, image_url, item, img_index, product_frame):
        """Загрузка одного изображения"""
        if not PIL_AVAILABLE:
            self.root.after(0, lambda: self.create_error_placeholder(parent_frame, "PIL недоступен"))
            return
            
        try:
            response = requests.get(image_url, timeout=10)
            response.raise_for_status()
            
            if IO_AVAILABLE:
                image = Image.open(BytesIO(response.content))
            else:
                # Сохраняем временно и загружаем
                temp_path = f"temp_img_{int(time.time())}_{img_index}.jpg"
                with open(temp_path, 'wb') as f:
                    f.write(response.content)
                image = Image.open(temp_path)
                os.remove(temp_path)
            
            image.thumbnail((120, 120), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            
            # self.root.after(0, lambda: self.create_image_button(parent_frame, photo, image_url, item, img_index))
            self.root.after(0, lambda: self.create_image_button(parent_frame, photo, image_url, item, img_index, product_frame))
            
        except Exception as e:
            self.root.after(0, lambda: self.create_error_placeholder(parent_frame, str(e)))




    
    # def create_image_button(self, parent_frame, photo, image_url, item, img_index):
    def create_image_button(self, parent_frame, photo, image_url, item, img_index, product_frame):
        """Создание кнопки с изображением"""
        try:
            if not parent_frame.winfo_exists():
                return
                
            is_selected = self.selected_images.get(item['article']) == image_url
            
            # btn = tk.Button(parent_frame, image=photo, 
            #                command=lambda: self.select_image(item['article'], image_url),
            #                relief='solid' if is_selected else 'raised',
            #                borderwidth=3 if is_selected else 1,
            #                bg='lightblue' if is_selected else 'white')
            btn = tk.Button(parent_frame, image=photo, 
                            command=lambda: self.select_image_new(item['article'], image_url, product_frame),
                            relief='solid' if is_selected else 'raised',
                            borderwidth=3 if is_selected else 1,
                            bg='lightblue' if is_selected else 'white')

            btn.pack()
            btn.image = photo
            
        except tk.TclError:
            pass
        except Exception as e:
            print(f"Ошибка создания кнопки изображения: {e}")
    
    def create_error_placeholder(self, parent_frame, error_msg):
        """Создание placeholder для ошибки загрузки"""
        try:
            if not parent_frame.winfo_exists():
                return
                
            placeholder = tk.Label(parent_frame, text="Ошибка\nзагрузки", 
                                 width=15, height=8, bg='lightgray')
            placeholder.pack()
            
        except tk.TclError:
            pass
        except Exception as e:
            print(f"Ошибка создания placeholder: {e}")
    
    def select_image(self, article, image_url):
        """Выбор изображения для товара"""
        if article in self.selected_images and self.selected_images[article] == image_url:
            del self.selected_images[article]
        else:
            self.selected_images[article] = image_url
        
        total_images = sum(len(item['images']) for item in self.parsed_results)
        self.viewer_stats.set(f"Товаров: {len(self.parsed_results)} | Изображений: {total_images} | Выбрано: {len(self.selected_images)}")
        
        self.load_parsed_results()

    

    def select_image_new(self, article, image_url, product_frame):
        """Выбор изображения для товара без ререндера"""
        # Обновляем выбранное изображение
        old_selection = self.selected_images.get(article)
        self.selected_images[article] = image_url
        
        # Обновляем отображение выбранного изображения
        self.update_selected_image_display(product_frame, image_url)
        
        # Обновляем индикатор в заголовке товара
        info_label = getattr(product_frame, 'info_label', None)
        if info_label:
            # Находим индекс товара
            index = 0
            for i, item in enumerate(self.parsed_results):
                if item['article'] == article:
                    index = i
                    break
            
            item_info = next((item for item in self.parsed_results if item['article'] == article), None)
            if item_info:
                info_text = f"✓ {index+1}. [{article}] {item_info['name']} ({item_info['source']})"
                info_label.config(text=info_text)
        
        # Обновляем статистику
        total_images = sum(len(item['images']) for item in self.parsed_results)
        self.viewer_stats.set(f"Товаров: {len(self.parsed_results)} | Изображений: {total_images} | Выбрано: {len(self.selected_images)}")
        
        # Обновляем визуальное состояние кнопок только для этого товара
        self.update_product_buttons_state(product_frame, article, image_url, old_selection)

    def update_selected_image_display(self, product_frame, image_url):
        """Обновление отображения выбранного изображения"""
        selected_frame = getattr(product_frame, 'selected_frame', None)
        if not selected_frame:
            return
        
        # Очищаем текущее содержимое
        for widget in selected_frame.winfo_children():
            widget.destroy()
        
        # Загружаем и отображаем выбранное изображение
        threading.Thread(target=self.load_selected_image, 
                    args=(selected_frame, image_url), daemon=True).start()

    def load_selected_image(self, parent_frame, image_url):
        """Загрузка выбранного изображения"""
        try:
            from io import BytesIO
            response = requests.get(image_url, timeout=10)
            response.raise_for_status()
            
            image = Image.open(BytesIO(response.content))
            # Изменяем размер для отображения в правой панели
            image.thumbnail((180, 180), Image.Resampling.LANCZOS)
            
            photo = ImageTk.PhotoImage(image)
            
            # Создаем label с изображением в основном потоке
            self.root.after(0, lambda: self.create_selected_image_label(parent_frame, photo, image_url))
            
        except Exception as e:
            self.root.after(0, lambda: self.create_selected_error_label(parent_frame, str(e)))

    def create_selected_image_label(self, parent_frame, photo, image_url):
        """Создание label с выбранным изображением"""
        label = tk.Label(parent_frame, image=photo, bg='white')
        label.pack(expand=True)
        label.image = photo  # Сохраняем ссылку
        
        # Добавляем URL под изображением
        url_label = tk.Label(parent_frame, text="Выбрано", fg='green', font=("Arial", 8))
        url_label.pack(pady=(5, 0))

    def create_selected_error_label(self, parent_frame, error_msg):
        """Создание label с ошибкой для выбранного изображения"""
        error_label = tk.Label(parent_frame, text="Ошибка\nзагрузки", 
                            fg='red', font=("Arial", 8))
        error_label.pack(expand=True)

    def update_product_buttons_state(self, product_frame, article, selected_url, old_selection):
        """Обновление состояния кнопок изображений для товара"""
        # Этот метод можно оставить пустым, так как визуальные изменения 
        # кнопок будут видны при следующем обновлении
        pass


























    
    def select_first_image(self, item):
        """Выбор первого изображения товара"""
        if item['images']:
            self.select_image(item['article'], item['images'][0])
    
    def export_selected(self):
        """Экспорт выбранных изображений"""
        if not self.selected_images:
            messagebox.showwarning("Предупреждение", "Нет выбранных изображений!")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Сохранить выбранные изображения",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                export_data = []
                for article, image_url in self.selected_images.items():
                    item_info = next((item for item in self.parsed_results if item['article'] == article), None)
                    if item_info:
                        export_data.append({
                            'article': article,
                            'name': item_info['name'],
                            'source': item_info['source'],
                            'selected_image': image_url
                        })
                
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(export_data, f, ensure_ascii=False, indent=2)
                
                messagebox.showinfo("Успех", f"Экспортировано {len(export_data)} изображений в {filename}")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при экспорте: {e}")
    
    def clear_selection(self):
        """Очистка выбора изображений"""
        self.selected_images.clear()
        self.load_parsed_results()
    
    def _on_mousewheel(self, event):
        """Обработка прокрутки колесиком мыши"""
        self.viewer_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    # Методы для обработки изображений
    def start_processing(self):
        """Запуск обработки изображений"""
        if not PIL_AVAILABLE:
            messagebox.showerror("Ошибка", "PIL/Pillow не установлен! Обработка изображений невозможна.")
            return
            
        if self.source_mode.get() == "folder":
            if not os.path.exists("input") or not os.listdir("input"):
                messagebox.showwarning("Предупреждение", "Папка input пуста или не существует!")
                return
            
            threading.Thread(target=self.process_folder_images, daemon=True).start()
        
        elif self.source_mode.get() == "json":
            if not self.selected_images:
                messagebox.showwarning("Предупреждение", "Нет выбранных изображений для обработки!")
                return
            
            threading.Thread(target=self.process_selected_images, daemon=True).start()
    
    def process_folder_images(self):
        """Обработка изображений из папки"""
        try:
            ratio = self.get_target_ratio()
            bg_color = self.get_bg_color()
            smart_mode = self.processing_mode.get() == "smart"
            
            input_files = [f for f in os.listdir("input") 
                          if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif'))]
            
            if not input_files:
                self.processing_progress_var.set("Нет изображений для обработки")
                return
            
            total_files = len(input_files)
            processed = 0
            
            self.processing_progress_var.set(f"Обработка {total_files} изображений...")
            
            for i, filename in enumerate(input_files):
                input_path = os.path.join("input", filename)
                output_path = os.path.join("output", filename)
                
                try:
                    success = self.processor.process_image(input_path, output_path, ratio, bg_color, smart_mode)
                    if success:
                        processed += 1
                        self.processing_log_message(f"✓ Обработан: {filename}")
                    else:
                        self.processing_log_message(f"✗ Ошибка: {filename}")
                
                except Exception as e:
                    self.processing_log_message(f"✗ Ошибка {filename}: {e}")
                
                progress = ((i + 1) / total_files) * 100
                self.processing_progress_bar['value'] = progress
                self.root.update_idletasks()
            
            self.processing_progress_var.set(f"Готово! Обработано: {processed}/{total_files}")
            
        except Exception as e:
            self.processing_progress_var.set(f"Ошибка обработки: {e}")
    
    def process_selected_images(self):
        """Обработка выбранных изображений"""
        try:
            ratio = self.get_target_ratio()
            bg_color = self.get_bg_color()
            smart_mode = self.processing_mode.get() == "smart"
            
            if not self.selected_images:
                self.processing_progress_var.set("Нет выбранных изображений")
                return
            
            total_images = len(self.selected_images)
            processed = 0
            
            self.processing_progress_var.set(f"Загрузка и обработка {total_images} изображений...")
            
            for i, (article, image_url) in enumerate(self.selected_images.items()):
                try:
                    if not article or not image_url:
                        self.processing_log_message(f"✗ Пропуск: пустые данные")
                        continue
                    
                    self.processing_log_message(f"Загрузка изображения для артикула: {article}")
                    
                    response = requests.get(image_url, timeout=15, headers={'User-Agent': USER_AGENT})
                    response.raise_for_status()
                    
                    if len(response.content) == 0:
                        self.processing_log_message(f"✗ Пустой файл для {article}")
                        continue
                    
                    content_type = response.headers.get('content-type', '').lower()
                    if 'jpeg' in content_type or 'jpg' in content_type:
                        ext = '.jpg'
                    elif 'png' in content_type:
                        ext = '.png'
                    elif 'gif' in content_type:
                        ext = '.gif'
                    elif 'bmp' in content_type:
                        ext = '.bmp'
                    else:
                        if image_url.lower().endswith('.png'):
                            ext = '.png'
                        elif image_url.lower().endswith('.gif'):
                            ext = '.gif'
                        elif image_url.lower().endswith('.bmp'):
                            ext = '.bmp'
                        else:
                            ext = '.jpg'
                    
                    safe_article = "".join(c for c in article if c.isalnum() or c in (' ', '-', '_')).strip()
                    if not safe_article:
                        safe_article = f"image_{i}"
                    
                    temp_path = f"temp_{safe_article}_{int(time.time())}{ext}"
                    with open(temp_path, 'wb') as f:
                        f.write(response.content)
                    
                    if not os.path.exists(temp_path) or os.path.getsize(temp_path) == 0:
                        self.processing_log_message(f"✗ Ошибка сохранения временного файла для {article}")
                        continue
                    
                    output_path = os.path.join("output", f"{safe_article}{ext}")
                    success = self.processor.process_image(temp_path, output_path, ratio, bg_color, smart_mode)
                    
                    try:
                        if os.path.exists(temp_path):
                            os.remove(temp_path)
                    except:
                        pass
                    
                    if success:
                        processed += 1
                        self.processing_log_message(f"✓ Обработан: {article}")
                    else:
                        self.processing_log_message(f"✗ Ошибка обработки: {article}")
                
                except requests.exceptions.RequestException as e:
                    self.processing_log_message(f"✗ Ошибка загрузки {article}: {e}")
                except Exception as e:
                    self.processing_log_message(f"✗ Ошибка {article}: {e}")
                
                progress = ((i + 1) / total_images) * 100
                self.processing_progress_bar['value'] = progress
                self.root.update_idletasks()
            
            self.processing_progress_var.set(f"Готово! Обработано: {processed}/{total_images}")
            
        except Exception as e:
            self.processing_progress_var.set(f"Ошибка обработки: {e}")
            self.processing_log_message(f"КРИТИЧЕСКАЯ ОШИБКА: {e}")
    
    def get_target_ratio(self):
        """Получение целевого соотношения сторон"""
        if self.ratio_var.get() == "Пользовательское":
            return self.custom_width.get() / self.custom_height.get()
        else:
            ratio_map = {
                "1:1": 1.0,
                "4:3": 4/3,
                "3:2": 3/2,
                "16:9": 16/9,
                "3:4": 3/4,
                "2:3": 2/3,
                "9:16": 9/16
            }
            return ratio_map.get(self.ratio_var.get(), 4/3)
    
    def get_bg_color(self):
        """Получение цвета фона"""
        try:
            color_str = self.bg_color.get().strip()
            if ',' in color_str:
                color_parts = color_str.split(',')
                if len(color_parts) >= 3:
                    r, g, b = map(int, color_parts[:3])
                    r = max(0, min(255, r))
                    g = max(0, min(255, g))
                    b = max(0, min(255, b))
                    return (r, g, b)
                else:
                    return (255, 255, 255)
            else:
                if color_str.startswith('#'):
                    color_str = color_str[1:]
                if len(color_str) >= 6:
                    r = int(color_str[0:2], 16)
                    g = int(color_str[2:4], 16)
                    b = int(color_str[4:6], 16)
                    return (r, g, b)
                else:
                    return (255, 255, 255)
        except Exception as e:
            print(f"Ошибка парсинга цвета: {e}")
            return (255, 255, 255)
    
    def processing_log_message(self, message):
        """Добавление сообщения в лог обработки"""
        self.root.after(0, lambda: self._append_processing_log(message))
    
    def _append_processing_log(self, message):
        """Внутренний метод для добавления в лог обработки"""
        self.processing_log.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.processing_log.see(tk.END)
    
    def open_folder(self, folder_name):
        """Открытие папки в проводнике"""
        try:
            if os.name == 'nt':
                os.startfile(folder_name)
            elif os.name == 'posix':
                os.system(f'open "{folder_name}"')
        except:
            messagebox.showinfo("Информация", f"Папка: {os.path.abspath(folder_name)}")
    
    # Методы для генерации Excel
    def browse_excel_folder(self):
        """Выбор папки для анализа"""
        folder = filedialog.askdirectory(title="Выберите папку для анализа")
        if folder:
            self.excel_folder.set(folder)
            self.base_path.set(os.path.abspath(folder))
    
    def create_excel_file(self):
        """Создание Excel файла со списком файлов"""
        try:
            input_folder = self.excel_folder.get()
            base_path = self.base_path.get()
            
            if not input_folder or not os.path.exists(input_folder):
                self.excel_result.insert(tk.END, f"Ошибка: Папка '{input_folder}' не найдена!\n")
                return
            
            output_file = os.path.join(input_folder, "output.xlsx")
            
            articles = []
            positions = []
            file_paths = []
            
            file_count = 0
            for root, dirs, files in os.walk(input_folder):
                for file in files:
                    if file.lower() == "output.xlsx":
                        continue
                    
                    file_lower = file.lower()
                    if not any(file_lower.endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff']):
                        continue
                    
                    try:
                        article = os.path.splitext(file)[0]
                        file_extension = os.path.splitext(file)[1]
                        
                        if base_path:
                            file_path = os.path.join(base_path, f"{article}{file_extension}")
                        else:
                            file_path = os.path.join(root, file)
                        
                        articles.append(article)
                        positions.append("")
                        file_paths.append(file_path)
                        file_count += 1
                        
                    except Exception as e:
                        self.excel_result.insert(tk.END, f"Ошибка обработки файла {file}: {e}\n")
                        continue
            
            if file_count == 0:
                self.excel_result.insert(tk.END, "В папке не найдено файлов изображений!\n")
                return
            
            df = pd.DataFrame({
                'Наименование файла': articles,
                'позиция': positions,
                'Путь к файлу': file_paths
            })
            
            df.to_excel(output_file, index=False, engine='openpyxl')
            
            result_text = f"Успешно создан файл '{output_file}'\n"
            result_text += f"Найдено файлов: {file_count}\n\n"
            result_text += "Первые 10 записей:\n"
            result_text += df.head(10).to_string(index=False)
            
            if len(df) > 10:
                result_text += f"\n... и еще {len(df) - 10} записей"
            
            self.excel_result.delete(1.0, tk.END)
            self.excel_result.insert(tk.END, result_text)
            
            messagebox.showinfo("Успех", f"Excel файл создан: {output_file}")
            
        except Exception as e:
            error_msg = f"Ошибка при создании Excel файла: {e}\n"
            self.excel_result.insert(tk.END, error_msg)
            messagebox.showerror("Ошибка", str(e))

def main():
    """Главная функция приложения"""
    root = tk.Tk()
    
    # Настройка темы
    try:
        style = ttk.Style()
        style.theme_use('clam')
    except:
        pass  # Используем тему по умолчанию
    
    # Создаем приложение
    app = ImageParserApp(root)
    
    # Запускаем главный цикл
    root.mainloop()

if __name__ == "__main__":
    # Проверка зависимостей
    print("Проверка зависимостей...")
    
    required_modules = {
        'tkinter': True,
        'requests': True,
        'pandas': True,
        'bs4': True,
        'PIL': PIL_AVAILABLE,
        'cv2': CV2_AVAILABLE,
        'io': IO_AVAILABLE
    }
    
    missing_modules = []
    
    for module, available in required_modules.items():
        if available:
            print(f"✓ {module} доступен")
        else:
            print(f"✗ {module} недоступен")
            missing_modules.append(module)
    
    if missing_modules:
        print(f"\nНедостает модулей: {', '.join(missing_modules)}")
        print("Установите недостающие модули:")
        if 'PIL' in missing_modules:
            print("pip install pillow")
        if 'cv2' in missing_modules:
            print("pip install opencv-python")
        if any(m in missing_modules for m in ['requests', 'pandas', 'bs4']):
            print("pip install requests pandas beautifulsoup4 openpyxl")
    
    print("\nЗапуск приложения...")
    main()