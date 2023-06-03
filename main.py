import requests
import openpyxl
from bs4 import BeautifulSoup
import pyttsx3
import speech_recognition as sr

def take_voice_input():
    # Initialize the recognizer
    r = sr.Recognizer()

    # Use the default microphone as the audio source
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1  # Set the pause threshold to control the end of speech
        audio = r.listen(source)

    try:
        print("Recognizing...")
        # Use the Google Web Speech API to recognize the audio
        text = r.recognize_google(audio, language='en')
        return text
    except sr.UnknownValueError:
        print("Sorry, I could not understand.")
        return None
    except sr.RequestError as e:
        print("Sorry, an error occurred: {0}".format(e))
        return None

def convert_text_to_speech(text):
    # Initialize the pyttsx3 engine
    engine = pyttsx3.init()
    
    # Convert text to speech
    engine.say(text)
    engine.runAndWait()

def get_soil_type(state):
    # Prepare the state name for URL
    state_url = state.lower().replace(' ', '-')

    # URL of the website with soil data
    url = f'https://www.soilhealth.dac.gov.in/NewHomePage/NBSS.aspx?state={state_url}'

    # Send a GET request to the website and retrieve the HTML content
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content
        soup = BeautifulSoup(response.content, 'html.parser')

        # Find the soil type element on the page
        soil_type_element = soup.find('td', {'class': 'main'})

        # Check if the soil type element was found
        if soil_type_element:
            # Extract the soil type text
            soil_type = soil_type_element.text.strip()
            return soil_type
        else:
            return None
    else:
        print('Error:', response.status_code)
        return None

def find_next_cell_value(file_path, search_string):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Iterate through the cells in column A
    for cell in sheet['A']:
        if cell.value == search_string:
            # Get the value from the next column
            next_column = sheet.cell(row=cell.row, column=cell.column + 1)
            return next_column.value

    return None

def fetch_weather(city):
    """
    City to weather
    :param city: City
    :return: weather
    """
    api_key = "32f40d1d539b58089d08ae7b223f0d17"
    #units_format = "&units=metric"
    c = 'cityandstate.xlsx'
    s = 'Stateandsoil.xlsx'
    state = find_next_cell_value(c, city)
    soil_type = get_soil_type(state)
    if soil_type:
        print(f"The soil type in {state} is: {soil_type}")
    else:
        soil = find_next_cell_value(s, state)

    base_url = "https://api.openweathermap.org/data/2.5/weather?q="
    complete_url = base_url + city + "&appid=" + api_key  

    response = requests.get(complete_url)

    city_weather_data = response.json()

    if city_weather_data["cod"] != "404":
        main_data = city_weather_data["main"]
        weather_description_data = city_weather_data["weather"][0]
        weather_description = weather_description_data["description"]
        current_temperature =round((main_data["temp"]-273.15))
        current_pressure = main_data["pressure"]
        current_humidity = main_data["humidity"]
        wind_data = city_weather_data["wind"]
        wind_speed = wind_data["speed"]
        if(wind_speed > 15):
            w = 'strong'
        elif(wind_speed > 10):
            w = 'medium'
        else:
            w = 'weak'

        final_response = f"""
        The weather in {city} is currently {weather_description} 
        with a temperature of {current_temperature} degree celcius, 
        humidity of {current_humidity} percent 
        and we have {w} wind, speed reaching {wind_speed} kilometers per hour
        soil is {soil} soil"""

        return final_response

    else:
        return "Sorry Sir, I couldn't find the city in my database. Please try again"

print("Enter your city: ")
convert_text_to_speech("Enter your city: ")
y = take_voice_input()
if (y == None):
    print("Please try again: ")
    convert_text_to_speech("Please try again: ")
    y = take_voice_input()
    if (y == None):
        convert_text_to_speech("Please try again mannually, its seems we are facing error")
        y = input("Please Enter your city: ")
print(y)
x = fetch_weather(y)
print(x)
convert_text_to_speech(x)
