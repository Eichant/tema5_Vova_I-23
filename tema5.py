import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import matplotlib.pyplot as plt
import os

class SalesAnalysisReport:
    """Клас для створення звіту з аналізу продажів товарів"""
    
    def __init__(self):
        plt.rcParams['font.family'] = 'DejaVu Sans'
        self.output_filename = 'sales_analysis_report.xlsx'
        
    def verify_dependencies(self):
        """Перевірка наявності необхідних бібліотек"""
        required_libraries = ['pandas', 'numpy', 'matplotlib', 'openpyxl']
        verification_results = []
        
        for library in required_libraries:
            try:
                __import__(library)
                verification_results.append((library, True, "Бібліотека доступна"))
            except ImportError:
                verification_results.append((library, False, "Бібліотека відсутня"))
        
        return verification_results
    
    def generate_sales_data(self, month_number, records_count=35):
        """Генерація тестових даних для продажів за вказаний місяць"""
        
        products = [
            'Ноутбук бізнес-класу', 'Монітор офісний', 'Клавіатура механічна',
            'Компютерна миша', 'Навушники бездротові', 'Веб-камера',
            'Мережевий комутатор', 'Система безперебійного живлення'
        ]
        
        clients = [
            'ТОВ "Інформаційні Технології"', 'ПП "Бізнес Консалтинг"',
            'ФОП Петренко І.В.', 'ТОВ "Фінансовий Аналітик"',
            'АБ "Український Капітал"', 'ТОВ "Логістичні Рішення"'
        ]
        
        data_records = []
        month_start = datetime(2025, month_number, 1)
        
        for record_index in range(records_count):
            product_selected = random.choice(products)
            client_selected = random.choice(clients)
            
            day_offset = random.randint(0, 27)
            transaction_date = month_start + timedelta(days=day_offset)
            
            week_number = (transaction_date.day - 1) // 7 + 1
            
            quantity_ordered = random.randint(1, 8)
            cost_price_unit = round(random.uniform(150, 2500), 2)
            selling_price_unit = round(cost_price_unit * random.uniform(1.15, 1.4), 2)
            
            total_order_value = round(quantity_ordered * selling_price_unit, 2)
            profit_calculated = round(quantity_ordered * (selling_price_unit - cost_price_unit), 2)
            
            payment_made = round(total_order_value * random.uniform(0.75, 1.0), 2)
            debt_outstanding = round(total_order_value - payment_made, 2)
            
            data_records.append({
                'Назва товару': product_selected,
                'Назва клієнта': client_selected,
                'Дата операції': transaction_date,
                'Тиждень місяця': week_number,
                'Кількість одиниць': quantity_ordered,
                'Собівартість одиниці': cost_price_unit,
                'Ціна реалізації': selling_price_unit,
                'Загальна вартість замовлення': total_order_value,
                'Фінансовий результат': profit_calculated,
                'Сплачена сума': payment_made,
                'Заборгованість': debt_outstanding
            })
        
        return pd.DataFrame(data_records)
    
    def create_analysis_pivot_table(self, data_frame):
        """Створення аналітичної зведеної таблиці"""
        try:
            pivot_table = data_frame.pivot_table(
                values='Сплачена сума',
                index='Назва клієнта',
                columns='Тиждень місяця',
                aggfunc='sum',
                fill_value=0
            )
            return pivot_table
        except Exception as error:
            print(f"Помилка формування зведеної таблиці: {error}")
            return None
    
    def format_excel_worksheet(self, worksheet, month_name):
        """Форматування аркуша Excel"""
        worksheet['A1'] = "Звіт з аналізу продажів товарів"
        worksheet['A2'] = f"Період: {month_name} 2025"
        worksheet['A3'] = f"Дата формування звіту: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        
        # Встановлення ширини стовпців
        column_widths = {
            'A': 25, 'B': 30, 'C': 15, 'D': 15, 
            'E': 15, 'F': 15, 'G': 15, 'H': 20,
            'I': 20, 'J': 15, 'K': 15
        }
        
        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width
    
    def generate_workbook(self):
        """Створення робочої книги з результатами аналізу"""
        
        analysis_periods = {
            1: 'Січень', 
            2: 'Лютий', 
            3: 'Березень',
            4: 'Квітень',
            5: 'Травень'
        }
        
        try:
            with pd.ExcelWriter(self.output_filename, engine='openpyxl') as excel_writer:
                
                print("Ініціалізація процесу створення звіту...")
                
                for month_number, month_name in analysis_periods.items():
                    print(f"Обробка даних за {month_name}...")
                    
                    # Генерація даних
                    sales_data = self.generate_sales_data(month_number, 25)
                    
                    # Збереження основних даних
                    sales_data.to_excel(
                        excel_writer, 
                        sheet_name=f"{month_number}_{month_name}", 
                        index=False, 
                        startrow=4
                    )
                    
                    # Форматування аркуша
                    current_worksheet = excel_writer.sheets[f"{month_number}_{month_name}"]
                    self.format_excel_worksheet(current_worksheet, month_name)
                    
                    # Створення зведеної таблиці
                    pivot_data = self.create_analysis_pivot_table(sales_data)
                    if pivot_data is not None:
                        pivot_data.to_excel(
                            excel_writer, 
                            sheet_name=f"Аналіз_{month_number}"
                        )
                    
                    print(f"Дані за {month_name} оброблено успішно")
                
                print(f"Робочу книгу '{self.output_filename}' створено")
                return True
                
        except Exception as error:
            print(f"Критична помилка при створенні файлу: {error}")
            return False
    
    def create_sales_trend_chart(self):
        """Створення графіка динаміки продажів"""
        try:
            months = ['Січень', 'Лютий', 'Березень', 'Квітень', 'Травень']
            total_sales = [184500, 216800, 245300, 198700, 267200]
            average_payment = [165200, 194500, 228400, 178300, 241600]
            
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
            
            # Графік загальних продажів
            bars1 = ax1.bar(months, total_sales, color='#2E86AB', alpha=0.7)
            ax1.set_title('Динаміка загального обсягу продажів по місяцях', fontsize=14, pad=20)
            ax1.set_ylabel('Сума, грн')
            ax1.grid(True, alpha=0.3)
            
            for bar, value in zip(bars1, total_sales):
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5000, 
                        f'{value:,}', ha='center', va='bottom', fontsize=10)
            
            # Графік сплачених коштів
            bars2 = ax2.bar(months, average_payment, color='#A23B72', alpha=0.7)
            ax2.set_title('Динаміка сплачених коштів по місяцях', fontsize=14, pad=20)
            ax2.set_ylabel('Сума, грн')
            ax2.set_xlabel('Місяць')
            ax2.grid(True, alpha=0.3)
            
            for bar, value in zip(bars2, average_payment):
                ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5000, 
                        f'{value:,}', ha='center', va='bottom', fontsize=10)
            
            plt.tight_layout()
            plt.savefig('sales_trend_analysis.png', dpi=200, bbox_inches='tight')
            plt.close()
            
            print("Графік динаміки продажів створено")
            return True
            
        except Exception as error:
            print(f"Помилка при створенні графіка: {error}")
            return False
    
    def display_file_information(self):
        """Відображення інформації про створені файли"""
        current_working_directory = os.getcwd()
        print(f"Поточний робочий каталог: {current_working_directory}")
        
        directory_files = os.listdir()
        excel_documents = [file for file in directory_files if file.endswith('.xlsx')]
        image_files = [file for file in directory_files if file.endswith('.png')]
        
        print("Документи Excel у каталозі:")
        for document in excel_documents:
            document_size = os.path.getsize(document)
            print(f"   - {document} ({document_size} байт)")
        
        print("Графічні файли у каталозі:")
        for image in image_files:
            image_size = os.path.getsize(image)
            print(f"   - {image} ({image_size} байт)")

def main():
    """Головна функція виконання програми"""
    
    analysis_report = SalesAnalysisReport()
    
    # Перевірка залежностей
    dependency_check = analysis_report.verify_dependencies()
    for library, status, message in dependency_check:
        status_indicator = "[+]" if status else "[-]"
        print(f"{status_indicator} {library}: {message}")
    
    # Створення робочої книги
    print("\nПочаток створення аналітичної звітності...")
    workbook_created = analysis_report.generate_workbook()
    
    # Створення графіків
    if workbook_created:
        analysis_report.create_sales_trend_chart()
    
    # Інформація про файли
    analysis_report.display_file_information()
    

    
    print("\n" + "=" * 70)
    if workbook_created:
        print("РОБОТУ УСПІШНО ЗАВЕРШЕНО")
        print("Результати збережено у поточному каталозі")
    else:
        print("ВИНИКЛИ ПРОБЛЕМИ ПРИ ВИКОНАННІ РОБОТИ")
    print("=" * 70)

if __name__ == "__main__":
    main()