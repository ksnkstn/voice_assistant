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
