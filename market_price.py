import pandas as pd
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

# Read the Excel file into a pandas DataFrame
df = pd.read_excel('market.xlsx')

# Get user input for commodity name and district
print("Enter the crop name: ")
convert_text_to_speech("Enter the crop name")
commodity_name = take_voice_input()
print(commodity_name)
print("Enter your district: ")
convert_text_to_speech("Enter your district: ")
district = take_voice_input()
print(district)
# Filter the DataFrame based on commodity name and district
filtered_data = df[(df['commodity_name'] == commodity_name) & (df['district'] == district)]

# Check if any records match the input
if filtered_data.empty:
    print("No matching records found.")
else:
    # Get the minimum and maximum prices
    min_price = filtered_data['min_price'].iloc[0]
    max_price = filtered_data['max_price'].iloc[0]
    modal_price = filtered_data['modal_price'].iloc[0]
    market = filtered_data['market'].iloc[0]

    # Print the minimum and maximum prices
    x = f"""In your nearest market, {market} the {commodity_name} crop, minimum price at which you should sell you yield is {min_price} and you sell it for maximum price of {max_price}. Also you need to keep in mind if you don't feel satisfied in with the selling price you can always sell it to government for the price {modal_price}."""
    print(x)
    convert_text_to_speech(x)
