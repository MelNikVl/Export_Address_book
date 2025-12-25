import csv
import re
from datetime import datetime
from time import sleep
from typing import List
from ldap3 import Server, Connection, ALL, SUBTREE
from tqdm import tqdm

def clean_string(s):
    """Очистка строки для CSV"""
    if s is None:
        return ""
    
    # Убираем переводы строк и лишние пробелы
    s = re.sub(r'[\r\n]+', ' ', s)
    s = s.strip()
    
    # Экранирование для CSV
    if ',' in s or '"' in s:
        s = s.replace('"', '""')
        s = f'"{s}"'
    
    return s

def get_global_catalog_users():
    # Параметры подключения
    gc_server = "------"  # Глобальный каталог (как пример xyz.int)
    base_dn = ""  # Для глобального каталога можно оставить пустым
    
    # Учетные данные
    username = "------"  # Имя пользователя
    password = "------"   # Пароль
    domain = "-----"           # Домен (как пример xyz.int)
    
    # Фильтр LDAP (как пример - @salutkbf.kz)
    ldap_filter = "(&(objectCategory=person)(objectClass=user)(mail=*--------))"
    
    # Атрибуты для получения
    attributes = ["name", "mail"]
    
    try:
        server = Server(gc_server, port=3268, get_info=ALL)
        user_upn = f"{username}@{domain}"
        
        print(f"Попытка подключения с UPN: {user_upn}")
        conn = Connection(server, user=user_upn, password=password, auto_bind=True)
        
        print("Подключение успешно!")
        
        # Поиск в глобальном каталоге
        conn.search(search_base=base_dn,
                   search_filter=ldap_filter,
                   search_scope=SUBTREE,
                   attributes=attributes,
                   paged_size=1000)
        users = []
        
        for entry in conn.entries:
            user = {}
            # Получаем атрибуты
            if hasattr(entry, 'name'):
                user['name'] = clean_string(str(entry.name))
            else:
                user['name'] = ""
                
            if hasattr(entry, 'mail'):
                user['mail'] = clean_string(str(entry.mail))
            else:
                user['mail'] = ""
            
            users.append(user)
        conn.unbind()
        
        return users
        
    except Exception as e:
        error_msg = str(e)
        print(f"Ошибка при подключении к глобальному каталогу: {error_msg}")
        print("\nВозможные причины ошибки:")
        print("1. Неверные учетные данные (имя пользователя или пароль)")
        print("2. Пользователь заблокирован или истек срок действия пароля")
        print("3. Нет прав доступа к глобальному каталогу")
        print("4. Неправильный формат имени пользователя")
        print("\nПопробуйте:")
        print(f"- Проверить учетные данные")
        print(f"- Использовать полный UPN формат: {username}@{domain}")
        print(f"- Или использовать полный DN пользователя")
        
        return []

def save_to_pst(users: List[dict]):
    """Сохранение пользователей в PST файл Outlook"""
    import win32com.client
    import os
    
    if not users:
        print("Нет данных для сохранения в PST")
        return
    
    outlook = None
    namespace = None
    pst_store = None

    try:
        # Создаем объект Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Генерируем имя PST файла с датой и временем
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pst_filename = f"contacts_{timestamp}.pst"
        pst_path = os.path.abspath(pst_filename)
        
        # Создаем новый PST файл
        namespace.AddStore(pst_path)
        print(f"Создан новый PST файл: {pst_path}")
        
        # Получаем последний добавленный PST (новый)
        stores = namespace.Stores
        pst_store = stores.Item(stores.Count)
        
        # Получаем папку контактов из PST
        contacts_folder = pst_store.GetDefaultFolder(10)  # 10 = olFolderContacts
        
        # Добавляем контакты с прогресс-баром
        added_count = 0
        skipped_count = 0
        
        print(f"\nДобавление контактов в PST...")
        for user in tqdm(users, desc="Обработка контактов", unit="контакт"):
            name = user.get('name', '').strip()
            mail = user.get('mail', '').strip()
            
            # Пропускаем записи без email
            if not mail:
                skipped_count += 1
                continue
            
            try:
                # Создаем новый контакт
                contact = contacts_folder.Items.Add("IPM.Contact")
                
                # Устанавливаем email
                contact.Email1Address = mail
                contact.Email1DisplayName = mail
                
                # Парсим имя (может быть "Фамилия Имя" или "Имя Фамилия")
                if name:
                    name_parts = name.split(maxsplit=1)
                    if len(name_parts) == 2:
                        # Предполагаем формат "Фамилия Имя"
                        contact.LastName = name_parts[0]
                        contact.FirstName = name_parts[1]
                        contact.FullName = name
                    else:
                        # Если одно слово, используем как фамилию
                        contact.LastName = name
                        contact.FullName = name
                else:
                    # Если имени нет, используем email
                    contact.FullName = mail
                
                # Сохраняем контакт
                contact.Save()
                added_count += 1
                
            except Exception as e:
                tqdm.write(f"Ошибка при добавлении контакта {mail}: {e}")
                skipped_count += 1
        
        print(f"\nСохранено в PST: {added_count} контактов")
        if skipped_count > 0:
            print(f"Пропущено: {skipped_count} записей (без email)")
        print(f"PST файл: {pst_path}")
        sleep(2)
        namespace.RemoveStore(contacts_folder)
        
    except Exception as e:
        print(f"Ошибка при сохранении в PST: {e}")
        import traceback
        traceback.print_exc()


# Использование
if __name__ == "__main__":
    # Получаем пользователей из глобального каталога
    users = get_global_catalog_users()
    
    if users:
        print(f"Найдено пользователей: {len(users)}")
        #user -> {"name": "", "mail": ""}
        save_to_pst(users)
    else:
        print("Пользователи не найдены")
