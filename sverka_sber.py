import pandas as pd
from datetime import datetime
from typing import Dict, Optional


#todo доделать обработку по 2 файлу, сделать более понятныестроки ошибок, чтобы не запутаться, добавить аннотацию для основной функции

class Parser:
    def __init__(self, path_proverka, path_etolon, path_very_etolon):
        """
        Инициализация объекта класса Parser.
    
        :param path_proverka: Путь к файлу Excel с данными для проверки.
        :param path_etolon: Путь к эталонному файлу Excel.
        :param path_very_etolon: Путь к очень эталонному файлу Excel.
        """
        self.path_proverka = path_proverka
        self.path_etolon = path_etolon
        self.path_very_etolon = path_very_etolon
        
        self.read_in_xls()


    def read_in_xls(self):
        """
        Чтение данных из Excel-файлов и сохранение их в атрибуты объекта.
    
        Загрузка данных из трех Excel-файлов в атрибуты объекта:
        - Проверка данных (proverka)
        - Эталонные данные (etolon)
        - Очень эталонные данные (very_etolon)
    
        В эталонном файле (etolon) удаляются последние две строки, так как они содержат ненужную информацию.
        """
        self.proverka = pd.read_excel(self.path_proverka)
        
        # Чтение данных из эталонного файла и удаление последних двух строк
        # ( эти строки содержат служебную или не нужную информацию)
        self.etolon = pd.read_excel(self.path_etolon)[:-2]
        
        self.very_etolon = pd.read_excel(self.path_very_etolon)
        

    def sverka_proverka_etolon(self):
        self.no_snils_str = '\n\n\n\nУ этих людей нет снилса (либо в проверке либо в эталоне, либо там и там )\n\n'
        self.no_birth_data = '\n\n\n\nУ этих людей нет даты рождения \n\n'
        self.no_spec = '\n\n\n\nУ этих людей нет поля специальность \n\n'
        self.bad_date_in_proverka = '\n\n\n\nУ этих людей не совпадает дата в файле проверки \n\n'
        self.bad_agreement = '\n\n\n\nУ этих людей не совпадают данные договора \n\n'
        self.bad_snils = '\n\n\n\nУ этих людей не совпадают данные снилса \n\n'
        self.bad_date = '\n\n\n\nУ этих людей не совпадают даты рождения \n\n'
        
        self.bad_spec = '\n\n\n\nУ этих людей неправильно указана специальность \n\n'
        self.no_pers_in_first_file = '\n\n\n\n\n Этих людей нет в 1 файле(проводится поиск по очень этолонному файлу):\n\n'
        self.bad_agreement_very_etolon = '\n\n\n\nУ этих людей нет полных данных договора (либо в проверке либо в эталоне, либо там и там )\n\n'
        self.error = '\n\n\n\nУ этих людей какие - либо данные записаны в неправильном формате (либо в проверке либо в эталоне, либо там и там )\n\n'
        
        validated_person_count = 0
        
        for _, pers_in_proverka in self.proverka.iterrows():
            pers_name = pers_in_proverka['Фамилия, имя, отчество заемщика']
            is_find = False
            all_data_true = pers_in_proverka['День рождения (информация от банка)'] == pers_in_proverka['День рождения (информация от организации)']
            
            if not all_data_true:
                self.bad_date_in_proverka += f"{pers_name}\n"
                continue
            
            # Проверка в эталонном файле
            for _, pers_in_etolon in self.etolon.iterrows():
                fullName_etolon = f"{pers_in_etolon['Фамилия']} {pers_in_etolon['Имя']} {pers_in_etolon['Отчество']}"
                if pers_name == fullName_etolon:
                    is_find = True
                    validated_person_count += 1
                    self.check_birth_data(pers_in_proverka, pers_in_etolon, fullName_etolon)
                    self.check_snils(pers_in_proverka, pers_in_etolon, fullName_etolon)
                    self.check_specialty_code(pers_in_proverka, pers_in_etolon, fullName_etolon)
                    self.check_agreement(pers_in_proverka, pers_in_etolon, fullName_etolon)
                
            if not is_find:
                #print(f"Person not found in etalon file: {pers_name}")
                self.no_pers_in_first_file += f"{pers_name} --- этого человека нет в эталонном файле\n"
                
                # Проверка в очень эталонном файле
                for _, pers_in_very_etolon in self.very_etolon.iterrows():
                    fullName_very_etolon = f"{pers_in_very_etolon['FAM']} {pers_in_very_etolon['IM']} {pers_in_very_etolon['OT']}"
                    if pers_name == fullName_very_etolon:
                        validated_person_count += 1
                        is_find = True
                        self.check_agreement_very_etolon(pers_in_proverka, pers_in_very_etolon, fullName_very_etolon)
                        # еще 3 проверки
            
            if not is_find:
                self.error += f"{pers_name} --- нет ни в одном эталонном файле, либо имя написано в неправильном регистре \n"
        full_mess = (
            self.bad_date + self.no_snils_str + self.no_spec + self.no_birth_data + self.bad_date_in_proverka +
            self.bad_agreement + self.bad_snils + self.bad_spec + self.no_pers_in_first_file +
            self.bad_agreement_very_etolon + self.error
        )
        
        print(f"Validated person count: {validated_person_count}")
        print(full_mess)
        
    
    
    def check_agreement_very_etolon(
        self,
        pers_in_proverka: Dict[str, Optional[str]],
        pers_in_very_etolon: Dict[str, Optional[str]],
        fullName_very_etolon: str
    ) -> None:
        """
        Проверяет согласованность данных договора между файлом проверки и очень эталонным файлом.
    
        Аргументы:
        pers_in_proverka -- словарь с данными из первого источника (например, файл проверки)
        pers_in_very_etolon -- словарь с данными из эталонного источника
        fullName_very_etolon -- полное имя эталонного источника данных
    
        Возвращаемое значение:
        Нет
        """
        
        try:
            # Извлечение данных из словарей
            agreement_in_proverka = pers_in_proverka[
                'Реквизиты договора об образовании, заключенного при приеме на обучение за счет средств физического и (или) юридического лица  (дата, номер) (информация от организации)'
            ]
            agreement_num_very_etolon = pers_in_very_etolon['NOMDOG']
            agreement_date_in_very_etolon = pers_in_very_etolon['DATADOG']
            
            # Проверка на наличие NaN значений
            if pd.isna(agreement_in_proverka) or pd.isna(agreement_num_very_etolon) or pd.isna(agreement_date_in_very_etolon):
                self.bad_agreement_very_etolon += (
                    f"{fullName_very_etolon} --- нет даты договора или номера договора в одном из файлов (сверка с очень эталонным файлом)\n"
                )
                return
            
            # Поиск даты и номера в строке
            start_index = agreement_in_proverka.find("-") + 1
            end_index = agreement_in_proverka.find("/", start_index)
            
            if start_index == -1 or end_index == -1 or start_index > end_index:
                self.error += f"{fullName_very_etolon} --- неправильно заполнена дата\n"
                return
            
            # Преобразование данных в даты и номера
            try:
                date1 = datetime.strptime(str(agreement_in_proverka[:10]), "%d.%m.%Y")
            except ValueError:
                self.error += f"{fullName_very_etolon} --- ошибка преобразования даты из строки договора\n"
                return
            
            try:
                date2 = datetime.strptime(str(agreement_date_in_very_etolon), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                self.error += f"{fullName_very_etolon} --- ошибка преобразования даты из эталонного файла\n"
                return
            
            # Сравнение данных
            if date1 != date2 or str(agreement_in_proverka[start_index:end_index]).strip() != str(agreement_num_very_etolon).strip():
                self.bad_agreement_very_etolon += (
                    f"{fullName_very_etolon} --- не сходятся данные договора (сверка с очень эталонным файлом)\n"
                )
                
        except KeyError as e:
            self.error += f"{fullName_very_etolon} --- отсутствует ключ {e} в одном из словарей\n"
        except Exception as e:
            self.error += f"{fullName_very_etolon} --- неожиданная ошибка: {e}\n"
            
            

    def check_birth_data(
        self,
        pers_in_proverka: Dict[str, Optional[str]],
        pers_in_etolon: Dict[str, Optional[str]],
        fullName_etolon: str
    ) -> None:
        """
        Проверяет согласованность данных о дате рождения файлом проверки и очень эталонным файлом.
    
        Аргументы:
        pers_in_proverka -- словарь с данными из первого источника (например, файл проверки)
        pers_in_etolon -- словарь с данными из эталонного источника
        fullName_etolon -- полное имя эталонного источника данных
    
        Возвращаемое значение:
        Нет
        """
        
        data_format1 = "%d.%m.%Y"
        data_format2 = "%Y-%m-%d"
        
        # Проверка на наличие NaN значений
        if pd.isna(pers_in_proverka['День рождения (информация от организации)']) or pd.isna(pers_in_etolon['Дата рождения']):
            self.no_birth_data += f"{fullName_etolon} --- нет даты Рождения в одном из файлов (сверка с эталонным файлом)\n"
            return 
        
        try:
            # Преобразование данных в даты
            date1 = datetime.strptime(str(pers_in_proverka['День рождения (информация от организации)']), data_format1)
            date2 = datetime.strptime(str(pers_in_etolon['Дата рождения']), data_format2)
            
            # Сравнение данных
            if date1 != date2:
                self.bad_date += f"{fullName_etolon} --- неправильная дата рождения\n"
        except ValueError as e:
            self.error += f"{fullName_etolon} --- ошибка формата даты: {e}\n"
            return
        
    

    def check_agreement(
        self,
        pers_in_proverka: Dict[str, Optional[str]],
        pers_in_etolon: Dict[str, Optional[str]],
        fullName_etolon: str
    ) -> None:
        """
        Проверяет согласованность данных договора между двумя источниками.
    
        Аргументы:
        pers_in_proverka -- словарь с данными из первого источника (например, файл проверки)
        pers_in_etolon -- словарь с данными из эталонного источника
        fullName_etolon -- полное имя эталонного источника данных
    
        Возвращаемое значение:
        Нет
        """
        
        data_format = "%d.%m.%Y"
        
        try:
            # Извлечение данных из словарей
            agreement_in_proverka = pers_in_proverka['Реквизиты договора об образовании, заключенного при приеме на обучение за счет средств физического и (или) юридического лица  (дата, номер) (информация от организации)']
            agreement_num_etolon = pers_in_etolon['№ договора']
            agreement_date_in_etolon = pers_in_etolon['Дата договора']
            
            # Проверка на наличие NaN значений
            if pd.isna(agreement_in_proverka) or pd.isna(agreement_num_etolon) or pd.isna(agreement_date_in_etolon):
                self.bad_agreement += f"{fullName_etolon} --- нет даты договора или номера договора в одном из файлов (сверка с эталонным файлом)\n"
                return 
            
            # Преобразование данных в даты
            try:
                date1 = datetime.strptime(str(agreement_in_proverka[:10]), data_format)
                date2 = datetime.strptime(str(agreement_date_in_etolon), data_format)
            except ValueError as e:
                self.error += f"{fullName_etolon} --- ошибка формата даты: {e}\n"
                return 
            
            # Сравнение данных
            if date1 != date2 or agreement_in_proverka[11:] != agreement_num_etolon:
                self.bad_agreement += f"{fullName_etolon} --- не сходятся данные договора (сверка с эталонным файлом)\n"
                return 
                    
        except KeyError as e:
            self.error += f"{fullName_etolon} --- отсутствует ключ {e} в одном из словарей\n"
        except Exception as e:
            self.error += f"{fullName_etolon} --- неожиданная ошибка: {e}\n"
        
        
        
    def check_snils(
        self,
        pers_in_proverka: Dict[str, Optional[str]],
        pers_in_etolon: Dict[str, Optional[str]],
        fullName_etolon: str
    ) -> None:
        """
        Проверяет согласованность данных СНИЛС между двумя источниками.
    
        Аргументы:
        pers_in_proverka -- словарь с данными из первого источника (например, файл проверки)
        pers_in_etolon -- словарь с данными из эталонного источника
        fullName_etolon -- полное имя эталонного источника данных
    
        Возвращаемое значение:
        Нет
        """
        
        try:
            # Проверка на наличие NaN значений и пустых строк
            if pd.isna(pers_in_proverka['СНИЛС (информация от организации)']) or pers_in_etolon['СНИЛС'] == '':
                self.no_snils_str += f"{fullName_etolon} --- нет СНИЛС в одном из файлов (сверка с эталонным файлом)\n"
                return 
            
            # Очистка и сравнение СНИЛС
            snils_proverka = int(pers_in_proverka['СНИЛС (информация от организации)'])
            snils_etolon = int(pers_in_etolon['СНИЛС'].replace(" ", "").replace("-", ""))
            
            if snils_proverka != snils_etolon:
                self.bad_snils += f"{fullName_etolon} --- неправильный СНИЛС\n"
                return 

        except ValueError as e:
            self.error += f"{fullName_etolon} --- ошибка преобразования СНИЛС: {e}\n"
        except KeyError as e:
            self.error += f"{fullName_etolon} --- отсутствует ключ {e} в одном из словарей\n"
        except Exception as e:
            self.error += f"{fullName_etolon} --- неожиданная ошибка: {e}\n"

    
    def check_specialty_code(
        self,
        pers_in_proverka: Dict[str, Optional[str]],
        pers_in_etolon: Dict[str, Optional[str]],
        fullName_etolon: str
    ) -> None:
        """
        Проверяет согласованность кода специальности между проверкой и этолоном.
    
        Аргументы:
        pers_in_proverka -- словарь с данными из первого источника (например, файл проверки)
        pers_in_etolon -- словарь с данными из эталонного источника
        fullName_etolon -- полное имя эталонного источника данных
    
        Возвращаемое значение:
        Нет
        """
        
        try:
            # Проверка на наличие NaN значений
            if pd.isna(pers_in_proverka['Код направления подготовки/специальности (информация от организации)']) or pd.isna(pers_in_etolon['Специальность']):
                self.no_spec += f"{fullName_etolon} --- нет специальности (сверка с эталонным файлом)\n"
                return 
            
            # Извлечение и обработка данных о специальности
            target_etolon_spec = pers_in_etolon['Специальность'][:-3] if '.' in pers_in_etolon['Специальность'] else pers_in_etolon['Специальность']
            
            # Сравнение данных
            if target_etolon_spec[1:] != pers_in_proverka['Код направления подготовки/специальности (информация от организации)']:
                self.bad_spec += f"{fullName_etolon}\n"
                return 
        
        except KeyError as e:
            self.error += f"{fullName_etolon} --- отсутствует ключ {e} в одном из словарей\n"
        except Exception as e:
            self.error += f"{fullName_etolon} --- неожиданная ошибка: {e}\n"
        

# Путь к файлам
path_a = "/Users/user/Desktop/work/sverka_sber/data/файл проверки.xlsx"
path_b = "/Users/user/Desktop/work/sverka_sber/data/эталон.xls"
path_c = "/Users/user/Desktop/work/sverka_sber/data/очень эталон.xls"

# Создание и запуск экземпляра Parser
parser_instance = Parser(path_a, path_b, path_c)
parser_instance.sverka_proverka_etolon()
