import random
import webbrowser
import pyttsx3
import speech_recognition
import wikipedia
from pyowm import OWM
import re
import pymorphy2
import urllib.parse
import os, sys
import imaplib
import email
from email.header import decode_header
from tkinter import *
import threading

# инициализация инструментов распознавания и ввода речи
recognizer = speech_recognition.Recognizer()
microphone = speech_recognition.Microphone()

engine = pyttsx3.init('sapi5')
engine.setProperty('voice', 'ru')


def record_and_recognize_audio():
    """
    Запись и распознавание аудио
    """
    with microphone:
        recognized_data = ""
        recognizer.adjust_for_ambient_noise(microphone, duration=2)  # Распознаватель на уровень шума в течение 2 секунд
        try:
            print("Listening...")  # delete
            # Запись аудио с микрофона в течение 5 секунд, с тайм-аутом 5 секунд
            audio = recognizer.listen(microphone, 5, 5)

        except speech_recognition.WaitTimeoutError:
            play_voice_assistant_speech("Проверте, включен ли ваш микрофон?")
            return

        # использование online-распознавания через Google
        try:
            print("Started recognition...")  # delete
            recognized_data = recognizer.recognize_google(audio, language="ru-RU").lower()
        except speech_recognition.UnknownValueError:
            pass

        # в случае проблем с доступом в Интернет происходит выброс ошибки
        except speech_recognition.RequestError:
            play_voice_assistant_speech("Проверте ваше подключение к интернету")
        return recognized_data


def play_voice_assistant_speech(text_to_speech):
    '''
    озвучка текста
    '''
    print(text_to_speech)  # delete
    engine.say(text_to_speech)
    text.insert(END, f"\nАссистент:\n{text_to_speech}")  # Добавление текста в текстовое поле
    engine.runAndWait()


def search_for_information_on_google(*args: tuple):
    '''
    поиск информации в google
    '''
    if not args[0]:  # если нет слов после команды
        return play_voice_assistant_speech('Скажите команду полностью')
    else:
        search_term = " ".join(args[0])
        encoded_search_term = urllib.parse.quote(search_term)  # Кодирование строки поиска для URL
        url = f"https://www.google.com/search?q={encoded_search_term}"
        webbrowser.open(url)

        return play_voice_assistant_speech('Открываю')


def search_for_information_on_youtube(*args: tuple):
    '''
    Поиск видео в youtube
    '''
    if not args[0]:  # если нет слов после команды
        return play_voice_assistant_speech('Скажите команду полностью')
    else:
        search_term = " ".join(args[0])
        encoded_search_term = urllib.parse.quote(search_term)  # Кодирование строки поиска для URL
        url = f"https://www.youtube.com/results?search_query={encoded_search_term}"
        webbrowser.open(url)

        return play_voice_assistant_speech('Открываю')


def stop_assistant(*args: tuple):
    play_voice_assistant_speech('Хорошо, до встречи!')
    finish()


def weather(city_name):
    '''
    Получение погоды
    '''
    try:
        owm = OWM('78fc9cb466b4aa6945d253a266eec9b5')  # Создание объекта OWM с API-ключом
        manager = owm.weather_manager()
        place = manager.weather_at_place(f"{city_name}")  # Получение погоды для заданного города
        res = place.weather  # Извлечение объекта weather с данными о погоде.
        value = int(res.temperature('celsius')['temp'])  # Получение температуры в градусах Цельсия
        return play_voice_assistant_speech(f'В городе {city_name.capitalize()} сейчас {value} градусов по Цельсию.')
    except Exception as e:
        return play_voice_assistant_speech(
            f'Не удалось получить погоду для города {city_name.capitalize()}. Ошибка: {e}')


def search_for_definition_on_wikipedia(*args: tuple):
    """
    Поиск в Wikipedia определения
    """
    if not args[0]:
        return play_voice_assistant_speech("Скажите команду полностью")

    search_term = " ".join(args[0])
    wikipedia.set_lang("ru")  # Установка русского языка

    try:
        page = wikipedia.page(search_term, auto_suggest=True)  # Поиск страница с автодополнением
        summary = wikipedia.summary(search_term, sentences=3)  # Краткое описание страницы
        play_voice_assistant_speech(f"Найдено в Википедии: \n{summary}")
    except wikipedia.exceptions.PageError:  # Обработка ошибки, если страница не найдена
        play_voice_assistant_speech("Не найдено в Wikipedia. Но найдено в гугл. Открываю")
        url = "https://google.com/search?q=" + search_term
        webbrowser.open(url)
    except Exception as e:
        play_voice_assistant_speech(f"Произошла ошибка: {e}")


def open_app(*args: tuple):
    """
    Открыть программу по ее названию.
    """
    if not args[0]:  # Если нет слов после команды
        return play_voice_assistant_speech('Скажите команду полностью')

    app = " ".join(args[0])
    try:
        if sys.platform == "win32":
            if (app == 'paint'):
                os.startfile('mspaint.exe')
            elif (app == 'word'):
                os.startfile('WINWORD.EXE')
            elif (app == 'google') or (app == 'гугл'):
                os.startfile('chrome.exe')
            elif (app == 'yandex') or (app == 'яндекс'):
                os.startfile('browser.exe')
            elif (app == 'powerpoint'):
                os.startfile('POWERPNT.EXE')
            elif (app == 'блокнот'):
                os.startfile('notepad.exe')
            else:
                return play_voice_assistant_speech(f"Приложение {app} не найдено")

        elif sys.platform == "darwin":
            if (app == 'mspaint.exe'):
                os.system('open -a Freeform')
            elif (app == 'WINWORD.EXE'):
                os.system('open -a /Applications/Microsoft\ Word.app')
            elif (app == 'chrome.exe'):
                os.system('open -a /Applications/Google\ Chrome.app')
            elif (app == 'POWERPNT.EXE'):
                os.system('open -a /Applications/Microsoft\ PowerPoint.app')
            elif (app == 'notepad.exe'):
                os.system('open -a TextEdit')
            else:
                return play_voice_assistant_speech(f"Приложение {app} не найдено")

        return play_voice_assistant_speech('Открываю')
    except FileNotFoundError:
        return play_voice_assistant_speech('Программа не найдена')


def get_desktop_path():
    """
    Получение пути к рабочему столу пользователя.
    """
    if sys.platform == "win32":
        return os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    elif sys.platform == "darwin":
        return os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')


def create_file(*args: tuple):
    """
    Создать файл с указанным именем на рабочем столе.
    """
    if not args[0]:  # Если нет слов после команды
        return 'Скажите команду полностью'

    filename = " ".join(args[0])
    desktop_path = get_desktop_path()  # Получение пути к рабочему столу
    file_path = os.path.join(desktop_path, filename)  # Формирование полного пути к файлу

    try:
        with open(file_path, 'w') as f:  # Открытие файла для записи (w) с именем file_path
            f.write("")  # Создаем пустой файл
        return play_voice_assistant_speech(f'Файл {filename} создан на рабочем столе')
    except Exception as e:
        return play_voice_assistant_speech(f'Ошибка при создании файла {filename}. Ошибка: {e}')


def create_folder(*args: tuple):
    """
    Создать папку с указанным именем на рабочем столе.
    """
    if not args[0]:  # Если нет слов после команды
        return 'Скажите команду полностью'

    foldername = " ".join(args[0])
    desktop_path = get_desktop_path()  # Получение пути к рабочему столу
    folder_path = os.path.join(desktop_path, foldername)  # Формирование полного пути к папке

    try:
        os.makedirs(folder_path, exist_ok=True)
        return play_voice_assistant_speech(f'Папка {foldername} создана на рабочем столе.')
    except Exception as e:
        return play_voice_assistant_speech(f'Ошибка при создании папки {foldername}. Ошибка: {e}')


def decode(msg):
    encoding = decode_header(msg)[0][1]  # извлечение кодировки
    msg = decode_header(msg)[0][0]  # извлечение декодированного текста
    if isinstance(msg, bytes):
        msg = msg.decode(encoding)  # декодирование байтов в строку
    return msg


def check_gmail(*args: tuple):
    try:
        # Параметры для подключения к почтовому серверу
        mail_pass = "cifaahkkbqimipco"  # Пароль
        username = "s73574903@gmail.com"
        imap_server = "imap.gmail.com"  # Задание адреса сервера IMAP для Gmail
        imap = imaplib.IMAP4_SSL(imap_server)  # Создание объекта imap для подключения к почтовому серверу IMAP
        imap.login(username, mail_pass)  # Авторизация
        imap.select(readonly=1)  # режим только для чтения
        '''
        Метод search  для поиска непрочитанных писем (UNSEEN) 
        retcode1  - код возврата операции поиска
        messages  - список идентификаторов непрочитанных писем
        '''
        (retcode1, messages) = imap.search(None, '(UNSEEN)')  # None - не используем "поиск по charset"
        if retcode1 == 'OK':  # статус операции
            play_voice_assistant_speech(f"Количество непрочитанных писем: {len(messages[0].split())}")
            for num in messages[0].split():
                redcode2, data = imap.fetch(num, '(RFC822)')  # извлекаем письмо в байтах
                msg = email.message_from_bytes(data[0][1])  # парсинг байтовых данных письма
                if redcode2 == 'OK':
                    play_voice_assistant_speech(
                        "\nТема: " + decode(msg["Subject"]) + '\nОтправитель: ' + ((msg["From"].split())[1]).strip(
                            "<>"))
        imap.close()

    except Exception as e:
        play_voice_assistant_speech(f"Произошла ошибка: {e}")

def play_random_music(*args: tuple):

    if sys.platform == "win32":
        music_dir = "C:\\Users\\Default\\Music"
    elif sys.platform == "darwin":
        music_dir = "Home\\Music"

    if os.path.isdir(music_dir):  # Проверка существования папки
        music_files = [f for f in os.listdir(music_dir) if
                       os.path.isfile(os.path.join(music_dir, f))]  # создание списка файлов в папке
        if music_files:
            random_music_file = random.choice(music_files)  # выбор случайного файла
            music_path = os.path.join(music_dir, random_music_file)  # полный путь к файлу
            try:
                os.startfile(music_path)
                return play_voice_assistant_speech(f'Играю {random_music_file}')
            except FileNotFoundError:
                return play_voice_assistant_speech('Файл не найден')
        else:
            return play_voice_assistant_speech('В папке музыки нет файлов')
    else:
        return play_voice_assistant_speech('Папка с музыкой не найдена')


commands = {
    ('искать', 'гугл', 'найди', 'найти'): search_for_information_on_google,
    ('погода', 'прогноз', 'погоду'): weather,
    ('стоп', 'хватит', 'пока'): stop_assistant,
    ('ютуб', 'youtube', 'ютубе'): search_for_information_on_youtube,
    ('википедия', 'wikipedia'): search_for_definition_on_wikipedia,
    ('открой', 'запусти'): open_app,
    ('проверь', 'check'): check_gmail,
    ('файл'): create_file,
    ('папка'): create_folder,
    ('включи', 'музыка', 'музыку'): play_random_music
}


def execute_command_with_name(command_name: str, *args: list):
    """
    Выполнение заданной пользователем команды с дополнительными аргументами
    :param command_name: название команды
    :param args: аргументы, которые будут переданы в функцию
    :return:
    """
    for key in commands.keys():
        if command_name in key:
            result = commands[key](*args)
            if result:  # Печатаем результат только если он не None
                print(result)
            break


def extract_city_name(voice_input):
    # Регулярное выражение для поиска фразы "погода в [город]" или "погода во [город]"
    match = re.search(
        r'(?:погода|скажи погоду|прогноз погоды|прогноз|скажи прогноз погоды|скажи прогноз) (?:в|во)? ([\w\s]+)',
        voice_input)
    if match:
        city = match.group(1).strip()
        return city
    else:
        match = re.search(
            r'(?:погода|скажи погоду|прогноз погоды|прогноз|скажи прогноз погоды|скажи прогноз) ([\w\s]+)', voice_input)
        if match:
            city = match.group(1).strip()
            return city
    return None


def convert_to_nominative(city):
    morph = pymorphy2.MorphAnalyzer()
    parsed = morph.parse(city)[0]
    return parsed.inflect({'nomn'}).word


work = 1


def main():
    global work

    while work:
        # старт записи речи с последующим выводом распознанной речи
        voice_input = record_and_recognize_audio()
        print(voice_input)  # delete
        text.insert(END, f"\nВы:\n{voice_input}", "right")

        if voice_input != None:
            city_in_prep_case = extract_city_name(voice_input)
            if city_in_prep_case:
                city_in_nom_case = convert_to_nominative(city_in_prep_case)
                weather(city_in_nom_case)
            else:
                # если не "погода в [город]", разбиваем ввод на команду и аргументы
                voice_input = voice_input.split(" ")
                command = voice_input[0]
                command_options = [str(input_part) for input_part in voice_input[1:len(voice_input)]]
                execute_command_with_name(command, command_options)


def thread_assistant():
    global work

    if button['text'] == 'Начать' or button['text'] == 'Продолжить':
        work = 1  # т.е. ассистент включен
        thread = threading.Thread(target=main)  # Создается новый поток с использованием функции threading.Thread
        # В качестве целевой функции main, которая будет выполняться в новом потоке
        thread.start()
        button['text'] = 'Пауза'

    elif button['text'] == 'Пауза':
        work = 0
        button['text'] = 'Продолжить'


root = Tk()  # создаем корневой объект - окно
root.title("")  # устанавливаем заголовок окна
root.geometry("300x400")  # устанавливаем размеры окна
root.config(bg="midnightblue")  # цвет фона окна

label = Label(text="Голосовой ассистент", font=("Arial", 14), background='midnightblue', foreground='white')
label.pack(pady=10)

text = Text(font=('Arial 10'), bg="lightsteelblue", wrap="word", height=15)  # текстовое поле для информации
text.pack(padx=15, fill=BOTH, expand=1)
text.tag_configure("right", justify="right")  # стиль выравнивания текста в текстовом поле

button = Button(root, text='Начать', command=thread_assistant, bg='lightsteelblue', width=10)
button.pack(pady=15)  # размещаем в окне


def finish():
    global work
    work = False
    root.destroy()  # ручное закрытие окна и всего приложения
    sys.exit(0)


root.protocol("WM_DELETE_WINDOW", finish)  # протокол: при закрытии окна запуск фун. finish
root.mainloop()  # Для отображения окна и взаимодействия с пользователем
