import asyncio
from collections import defaultdict
from datetime import datetime
import glob
import openpyxl
import os
import pandas as pd
import sys
import time
from typing import (
    Any,
    DefaultDict,
    Dict,
    List,
    Optional,
    Tuple
)


class MedicalDevice:
    """
    Класс для хранения данных об устройстве.
    """

    def __init__(self,
                 device_id: str,
                 clinic_id: str,
                 clinic_name: str,
                 city: str,
                 department: str,
                 model: str,
                 serial_number: str,
                 install_date: Optional[datetime],
                 status: str,
                 warranty_until: Optional[datetime],
                 last_calibration_date: Optional[datetime],
                 last_service_date: Optional[datetime],
                 issues_reported_12mo: float,
                 failure_count_12mo: float,
                 uptime_pct: float,
                 issues_text: str,
                 source_file: str = '') -> None:
        """
        Инициализация медицинского устройства.

        :param device_id: уникальный идентификатор устройства.
        :param clinic_id: уникальный идентификатор клиники.
        :param clinic_name: название клиники.
        :param city: город установки.
        :param department: медицинское отделение.
        :param model: модель устройства.
        :param serial_number: серийный номер.
        :param install_date: дата установки.
        :param status: текущий статус устройства.
        :param warranty_until: дата окончания гарантии.
        :param last_calibration_date: дата последней калибровки.
        :param last_service_date: дата последнего обслуживания.
        :param issues_reported_12mo: количество проблем за 12 месяцев.
        :param failure_count_12mo: количество отказов за 12 месяцев.
        :param uptime_pct: процент работоспособности.
        :param issues_text: текстовое описание проблем.
        :param source_file: имя файла-источника.
        """

        self.device_id = device_id
        self.clinic_id = clinic_id
        self.clinic_name = clinic_name
        self.city = city
        self.department = department
        self.model = model
        self.serial_number = serial_number
        self.install_date = install_date
        self.status = status
        self.warranty_until = warranty_until
        self.last_calibration_date = last_calibration_date
        self.last_service_date = last_service_date
        self.issues_reported_12mo = float(issues_reported_12mo or 0)
        self.failure_count_12mo = float(failure_count_12mo or 0)
        self.uptime_pct = float(uptime_pct or 0)
        self.issues_text = issues_text
        self.source_file = source_file


class AsyncMedicalDeviceAnalyzer:
    """
    Асинхронный анализатор для нескольких файлов.
    """

    def __init__(self, folder_path: str) -> None:
        """
        Инициализация анализатора.

        :param folder_path: путь к папке с Excel файлами.
        """

        self.folder_path = folder_path
        self.devices = []
        self.current_date = datetime.now()

        # нормализация статусов
        self.status_mapping = {
            'operational': 'operational', 'op': 'operational', 'ok': 'operational', 'working': 'operational',
            'planned_installation': 'planned_installation', 'planned': 'planned_installation',
            'maintenance_scheduled': 'maintenance_scheduled', 'maintenance': 'maintenance_scheduled',
            'faulty': 'faulty', 'broken': 'faulty', 'error': 'faulty'
        }

    def parse_date(self, value: Any) -> Optional[datetime]:
        """
        Парсинг даты.

        :param value: значение для парсинга.

        :return: Optional[datetime]: объект datetime или None.
        """

        # проверяем пустое значение
        if not value or value == '':

            return None

        # если уже datetime, возвращаем как есть
        if isinstance(value, datetime):

            return value

        # пробуем разные форматы дат
        for fmt in ['%Y-%m-%d', '%d.%m.%Y', '%Y/%m/%d']:
            try:

                return datetime.strptime(str(value), fmt)

            except:
                pass # продолжаем со следующим форматом

        return None # ни один формат не подошел

    async def load_single_file(self, file_path: str) -> List[MedicalDevice]:
        """
        Загрузка одного файла Excel.

        :param file_path: путь к файлу.

        :return: List[MedicalDevice]: список устройств из файла.
        """

        # загружаем workbook в отдельном потоке
        loop = asyncio.get_event_loop()
        wb = await loop.run_in_executor(None, openpyxl.load_workbook, file_path, True)

        ws = wb.active # берем активный лист
        headers = [cell.value for cell in ws[1]] # читаем заголовки из первой строки

        devices = []
        # проходим по всем строкам начиная со второй
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                data = dict(zip(headers, row))

                # нормализуем статус устройства
                status = str(data.get('status', '')).lower().strip()
                normalized_status = self.status_mapping.get(status, status)

                # создаем объект MedicalDevice
                device = MedicalDevice(
                    device_id=str(data.get('device_id', '')),
                    clinic_id=str(data.get('clinic_id', '')),
                    clinic_name=str(data.get('clinic_name', '')),
                    city=str(data.get('city', '')),
                    department=str(data.get('department', '')),
                    model=str(data.get('model', '')),
                    serial_number=str(data.get('serial_number', '')),
                    install_date=self.parse_date(data.get('install_date')),
                    status=normalized_status,
                    warranty_until=self.parse_date(data.get('warranty_until')),
                    last_calibration_date=self.parse_date(data.get('last_calibration_date')),
                    last_service_date=self.parse_date(data.get('last_service_date')),
                    issues_reported_12mo=data.get('issues_reported_12mo', 0),
                    failure_count_12mo=data.get('failure_count_12mo', 0),
                    uptime_pct=data.get('uptime_pct', 0),
                    issues_text=str(data.get('issues_text', '')),
                    source_file=os.path.basename(file_path)
                )
                devices.append(device)

        wb.close() # закрываем workbook

        # выводим информацию о загруженном файле
        print(f"  Загружено {len(devices)} устройств из {os.path.basename(file_path)}")

        return devices

    async def load_all_files(self, file_pattern: str = "*.xlsx") -> List[MedicalDevice]:
        """
        Загрузка всех Excel файлов из папки.

        :param file_pattern: шаблон файлов для поиска.

        :return: List[MedicalDevice]: список всех устройств.
        """

        start_time = time.time()

        # проверяем существование папки
        if not os.path.exists(self.folder_path):
            print(f"Ошибка: Папка '{self.folder_path}' не существует!")

            return []

        # ищем все excel файлы по шаблону
        search_pattern = os.path.join(self.folder_path, file_pattern)
        files = glob.glob(search_pattern)

        # проверяем, есть ли файлы
        if not files:
            print(f"В папке '{self.folder_path}' не найдено Excel файлов!")
            print(f"Искали по шаблону: {search_pattern}")

            return []

        # выводим список найденных файлов
        print(f"\nНайдено {len(files)} файлов в папке '{self.folder_path}':")
        for f in files:
            print(f"  - {os.path.basename(f)}")

        # загружаем файлы параллельно (асинхронно)
        tasks = [self.load_single_file(file_path) for file_path in files]
        results = await asyncio.gather(*tasks)

        # объединяем все устройства из всех файлов
        all_devices = []
        for devices in results:
            all_devices.extend(devices)

        # удаляем дубликаты по device_id
        seen = set()
        unique = []
        for d in all_devices:
            if d.device_id not in seen:
                seen.add(d.device_id)
                unique.append(d)

        self.devices = unique  # сохраняем уникальные устройства

        # выводим итоговую статистику
        elapsed = time.time() - start_time
        print(f"\nИТОГО: Загружено {len(self.devices)} устройств из {len(files)} файлов (время: {elapsed:.2f} сек)")

        return self.devices

    async def filter_by_warranty(self) -> List[MedicalDevice]:
        """
        Фильтрация по гарантии.

        :return: List[MedicalDevice]: список устройств с истекающей гарантией (<=30 дней).
        """

        start_time = time.time()

        filtered = []  # список устройств с истекающей гарантией
        warranty_stats = defaultdict(int)  # статистика по категориям гарантии

        # проходим по всем устройствам
        for device in self.devices:
            if device.warranty_until:
                # рассчитываем сколько дней осталось до конца гарантии
                days_left = (device.warranty_until - self.current_date).days

                # распределяем по категориям
                if days_left <= 0:
                    warranty_stats['истекла'] += 1
                elif days_left <= 30:
                    warranty_stats['< 30 дней'] += 1
                    filtered.append(device) # добавляем в список с истекающей гарантией
                elif days_left <= 90:
                    warranty_stats['30-90 дней'] += 1
                elif days_left <= 180:
                    warranty_stats['90-180 дней'] += 1
                elif days_left <= 365:
                    warranty_stats['180-365 дней'] += 1
                else:
                    warranty_stats['больше года'] += 1
            else:
                warranty_stats['нет данных'] += 1

        # выводим статистику по гарантии
        print("\nСтатус гарантии:")
        for cat, count in sorted(warranty_stats.items()):
            print(f"  {cat}: {count}")

        print(f"\nУстройств с истекающей гарантией (<=30 дней): {len(filtered)}")

        elapsed = time.time() - start_time
        print(f"Время: {elapsed:.2f} сек")

        return filtered

    async def find_problem_clinics(self, top_n: int = 10) -> List[Dict[str, Any]]:
        """
        Поиск проблемных клиник.

        :param top_n: количество клиник для вывода.

        :return: List[Dict[str, Any]]: список топ клиник с проблемами.
        """

        start_time = time.time()

        # словарь для сбора статистики по клиникам
        stats = defaultdict(lambda: {'devices': 0, 'problems': 0, 'failures': 0, 'city': ''})

        for device in self.devices:
            key = device.clinic_name  # используем название клиники как ключ
            stats[key]['devices'] += 1  # увеличиваем счетчик устройств
            stats[key]['failures'] += device.failure_count_12mo  # суммируем отказы
            stats[key]['city'] = device.city  # сохраняем город

            # проверяем, есть ли проблемы у устройства
            if device.failure_count_12mo > 0 or device.issues_reported_12mo > 0:
                stats[key]['problems'] += 1

        # формируем результат
        clinics = []
        for name, data in stats.items():
            # рассчитываем процент проблемных устройств
            problem_percent = (data['problems'] / data['devices'] * 100) if data['devices'] > 0 else 0
            clinics.append({
                'clinic': name,
                'city': data['city'],
                'total_devices': data['devices'],
                'problem_devices': data['problems'],
                'total_failures': data['failures'],
                'problem_percent': round(problem_percent, 1)
            })

        # сортируем по количеству проблемных устройств (по убыванию)
        clinics.sort(key=lambda x: x['problem_devices'], reverse=True)
        top = clinics[:top_n]  # берем топ N клиник

        # выводим таблицу с топ клиниками
        print("\nТоп проблемных клиник:")
        print(f"{'Клиника':<30} {'Город':<15} {'Устройств':<10} {'Проблемных':<10} {'% проблем':<10}")
        print("-" * 75)

        for c in top:
            print(
                f"{c['clinic']:<30} {c['city']:<15} {c['total_devices']:<10} {c['problem_devices']:<10} {c['problem_percent']:<10}")

        elapsed = time.time() - start_time
        print(f"Время: {elapsed:.2f} сек")

        return top

    async def calibration_report(self) -> Dict[str, Dict[str, int]]:
        """
        Отчет по калибровке.

        :return: Dict[str, Dict[str, int]]: словарь с категориями калибровки.
        """

        start_time = time.time()

        # инициализируем категории для отчета
        categories = {
            '< 30 дней': {'count': 0, 'failures': 0},
            '30-90 дней': {'count': 0, 'failures': 0},
            '90-180 дней': {'count': 0, 'failures': 0},
            '180-365 дней': {'count': 0, 'failures': 0},
            '> 1 года': {'count': 0, 'failures': 0},
            'нет данных': {'count': 0, 'failures': 0}
        }

        # распределяем устройства по категориям
        for device in self.devices:
            if device.last_calibration_date:
                # рассчитываем сколько дней прошло с последней калибровки
                days = (self.current_date - device.last_calibration_date).days

                # определяем категорию по количеству дней
                if days < 30:
                    cat = '< 30 дней'
                elif days < 90:
                    cat = '30-90 дней'
                elif days < 180:
                    cat = '90-180 дней'
                elif days < 365:
                    cat = '180-365 дней'
                else:
                    cat = '> 1 года'

                categories[cat]['count'] += 1
                categories[cat]['failures'] += device.failure_count_12mo
            else:
                # нет данных о калибровке
                categories['нет данных']['count'] += 1
                categories['нет данных']['failures'] += device.failure_count_12mo

        # выводим отчет
        print("\nОтчет по калибровке:")
        print(f"{'Срок':<20} {'Устройств':<10} {'Отказов':<10}")
        print("-" * 40)

        for cat, data in categories.items():
            print(f"{cat:<20} {data['count']:<10} {data['failures']:<10}")

        elapsed = time.time() - start_time
        print(f"Время: {elapsed:.2f} сек")

        return categories

    async def create_pivot(self) -> List[Dict[str, Any]]:
        """
        Сводная таблица.

        :return: List[Dict[str, Any]]: список записей сводной таблицы.
        """

        start_time = time.time()

        # вложенный defaultdict для агрегации данных
        pivot_data = defaultdict(lambda: defaultdict(lambda: {'count': 0, 'failures': 0, 'uptime_sum': 0}))

        for device in self.devices:
            # группируем по клинике, городу и модели
            key = (device.clinic_name, device.city, device.model)
            pivot_data[key[0]][(key[1], key[2])]['count'] += 1  # количество устройств
            pivot_data[key[0]][(key[1], key[2])]['failures'] += device.failure_count_12mo  # сумма отказов
            pivot_data[key[0]][(key[1], key[2])]['uptime_sum'] += device.uptime_pct  # сумма uptime

        # формируем результат в виде списка словарей
        result = []
        for clinic, models in pivot_data.items():
            for (city, model), data in models.items():
                # рассчитываем средний uptime
                avg_uptime = data['uptime_sum'] / data['count'] if data['count'] > 0 else 0
                result.append({
                    'clinic': clinic,
                    'city': city,
                    'model': model,
                    'devices': data['count'],
                    'failures': data['failures'],
                    'avg_uptime': round(avg_uptime, 1)
                })

        print(f"\nСводная таблица: {len(result)} записей")

        elapsed = time.time() - start_time
        print(f"Время: {elapsed:.2f} сек")

        return result

    async def save_reports(self,
                           filtered: List[MedicalDevice],
                           clinics: List[Dict[str, Any]],
                           calibration: Dict[str, Dict[str, int]],
                           pivot: List[Dict[str, Any]],
                           output_file: str = 'medical_devices_report.xlsx') -> None:
        """
        Сохранение отчетов в Excel на разные листы.

        :param filtered: список устройств с истекающей гарантией.
        :param clinics: топ проблемных клиник.
        :param calibration: отчет по калибровке.
        :param pivot: сводная таблица.
        :param output_file: имя выходного файла.
        """

        start_time = time.time()

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

            # лист 1: устройства с истекающей гарантией
            filtered_data = []
            for device in filtered:
                filtered_data.append({
                    'device_id': device.device_id,
                    'clinic_name': device.clinic_name,
                    'city': device.city,
                    'model': device.model,
                    'status': device.status,
                    'warranty_until': device.warranty_until if device.warranty_until else None,
                    'source_file': device.source_file
                })
            df_filtered = pd.DataFrame(filtered_data)
            df_filtered.to_excel(writer, sheet_name='Фильтр по гарантии', index=False)

            # лист 2: топ проблемных клиник
            df_clinics = pd.DataFrame(clinics)
            df_clinics.to_excel(writer, sheet_name='Топ проблемных клиник', index=False)

            # лист 3: отчет по калибровке
            calibration_data = []
            for cat, data in calibration.items():
                calibration_data.append({
                    'срок_с_последней_калибровки': cat,
                    'количество_устройств': data['count'],
                    'количество_отказов': data['failures']
                })
            df_calibration = pd.DataFrame(calibration_data)
            df_calibration.to_excel(writer, sheet_name='Отчет по калибровке', index=False)

            # лист 4: сводная таблица
            df_pivot = pd.DataFrame(pivot)
            df_pivot.to_excel(writer, sheet_name='Сводная таблица', index=False)

            # лист 5: исходные данные (первые 5000 строк)
            devices_data = []
            for device in self.devices[:5000]:
                devices_data.append({
                    'device_id': device.device_id,
                    'clinic_id': device.clinic_id,
                    'clinic_name': device.clinic_name,
                    'city': device.city,
                    'department': device.department,
                    'model': device.model,
                    'serial_number': device.serial_number,
                    'install_date': device.install_date,
                    'status': device.status,
                    'warranty_until': device.warranty_until,
                    'last_calibration_date': device.last_calibration_date,
                    'last_service_date': device.last_service_date,
                    'issues_reported_12mo': device.issues_reported_12mo,
                    'failure_count_12mo': device.failure_count_12mo,
                    'uptime_pct': device.uptime_pct,
                    'issues_text': device.issues_text,
                    'source_file': device.source_file
                })
            df_devices = pd.DataFrame(devices_data)
            df_devices.to_excel(writer, sheet_name='Исходные данные', index=False)

            # лист 6: статистика по файлам
            file_stats = defaultdict(lambda: {'devices': 0, 'warranty_expiring': 0})
            for device in self.devices:
                file_stats[device.source_file]['devices'] += 1

            for device in filtered:
                file_stats[device.source_file]['warranty_expiring'] += 1

            stats_data = []
            for file, data in file_stats.items():
                stats_data.append({
                    'файл': file,
                    'всего_устройств': data['devices'],
                    'с_истекающей_гарантией': data['warranty_expiring']
                })
            df_stats = pd.DataFrame(stats_data)
            df_stats.to_excel(writer, sheet_name='Статистика по файлам', index=False)

        elapsed = time.time() - start_time
        print(f"\nОтчеты сохранены в файл: {output_file}")
        print(
            f"  Листы: Фильтр по гарантии, Топ проблемных клиник, Отчет по калибровке, Сводная таблица, Исходные данные, Статистика по файлам")
        print(f"Время сохранения: {elapsed:.2f} сек")


async def main() -> None:
    total_start = time.time()

    print("АСИНХРОННЫЙ АНАЛИЗАТОР МЕДИЦИНСКОГО ОБОРУДОВАНИЯ")

    # получаем путь к папке из аргумента командной строки или от пользователя
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
        print(f"\nПапка указана в аргументе: {folder_path}")
    else:
        folder_path = input("\nВведите путь к папке с Excel файлами: ").strip()

    # если путь не указан, используем текущую папку
    if not folder_path:
        folder_path = "."
        print(f"Используем текущую папку: {folder_path}")

    # проверяем существование папки
    if not os.path.exists(folder_path):
        print(f"\nОшибка: Папка '{folder_path}' не существует!")

        return

    # создаем экземпляр анализатора
    analyzer = AsyncMedicalDeviceAnalyzer(folder_path)

    # загружаем данные из всех файлов
    print("\n1. ЗАГРУЗКА ДАННЫХ ИЗ ФАЙЛОВ")
    print("-" * 40)
    await analyzer.load_all_files("*.xlsx")

    # проверяем, есть ли данные для анализа
    if not analyzer.devices:
        print("Нет данных для анализа!")
        return

    # выполняем все аналитические задачи параллельно
    print("\n2. ВЫПОЛНЕНИЕ АНАЛИЗА")
    print("-" * 40)
    results = await asyncio.gather(
        analyzer.filter_by_warranty(),
        analyzer.find_problem_clinics(),
        analyzer.calibration_report(),
        analyzer.create_pivot()
    )

    # сохраняем результаты
    print("\n3. СОХРАНЕНИЕ РЕЗУЛЬТАТОВ")
    print("-" * 40)
    await analyzer.save_reports(*results)

    # выводим общее время выполнения
    total_elapsed = time.time() - total_start
    print(f"ОБЩЕЕ ВРЕМЯ ВЫПОЛНЕНИЯ: {total_elapsed:.2f} секунд")


if __name__ == "__main__":
    asyncio.run(main())