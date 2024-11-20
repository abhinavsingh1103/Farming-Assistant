## Overview
This Python application provides weather details for a specified city and includes information about the soil type in the respective state. The application incorporates voice input and output functionality to make it user-friendly. It combines weather API data, web scraping, and Excel-based information retrieval to provide a comprehensive response.

---

## Features
- **Voice Interaction**: Users can input their city name via speech, and the app responds with speech output.
- **Weather Information**: Fetches real-time weather data using the OpenWeatherMap API.
- **Soil Information**: Determines the soil type of the state associated with the city.
- **Fallback Handling**: Allows manual input in case of voice recognition issues.

---

## Prerequisites
### Libraries
Install the required Python libraries using:
```bash
pip install requests beautifulsoup4 pyttsx3 SpeechRecognition openpyxl
```

### API Key
- Sign up for an API key at [OpenWeatherMap](https://openweathermap.org/) and replace the placeholder in the code:
  ```python
  api_key = "your_openweathermap_api_key"
  ```

### Files
- **`cityandstate.xlsx`**: A file mapping cities to their respective states.
- **`Stateandsoil.xlsx`**: A file mapping states to their soil types.

---

## How to Use

### Steps to Run
1. Clone the repository or download the project files.
2. Ensure the `cityandstate.xlsx` and `Stateandsoil.xlsx` files are correctly formatted and placed in the same directory as the script.
3. Run the script:
   ```bash
   python weather_soil_app.py
   ```
4. Speak the name of the city when prompted, or type it manually if needed.

### Interaction Flow
1. The program prompts you to enter your city name using voice input.
2. If the voice input is unclear, you'll be prompted to try again or enter the city manually.
3. The program fetches:
   - Weather data for the specified city.
   - Soil information for the city’s state (via scraping or Excel lookup).
4. The weather and soil details are displayed in text and read aloud.

---

## Input and Output Details

### Input
- **City Name**: Provided via voice or manual input.

### Output
Example response:
```
The weather in New York is currently clear sky 
with a temperature of 20 degrees Celsius, 
humidity of 60 percent, 
and we have medium wind, speed reaching 12 kilometers per hour. 
The soil is Sandy soil.
```

---

## Project Components

### Key Functions
1. **`take_voice_input()`**
   - Captures voice input using the microphone.
2. **`convert_text_to_speech(text)`**
   - Converts text output to speech for user interaction.
3. **`get_soil_type(state)`**
   - Scrapes soil type information from a government website.
4. **`find_next_cell_value(file_path, search_string)`**
   - Retrieves data from an Excel file.
5. **`fetch_weather(city)`**
   - Fetches weather and soil data for the specified city.

### External Resources
1. **[OpenWeatherMap API](https://openweathermap.org/)**: Fetches real-time weather data.
2. **Excel Files**:
   - `cityandstate.xlsx`: Maps cities to states.
   - `Stateandsoil.xlsx`: Maps states to soil types.

---

## Folder Structure
```
project/
│
├── weather_soil_app.py         # Main application script
├── cityandstate.xlsx           # Maps cities to states
├── Stateandsoil.xlsx           # Maps states to soil types
└── README.md                   # Documentation
```

---

## Error Handling
- **Voice Recognition**: If the voice input fails, the program prompts the user to retry or switch to manual input.
- **City Not Found**: Returns a friendly error message if the city is not available in the database.
- **Web Scraping Issues**: Fallback to Excel-based lookup if scraping fails.

---

## Future Improvements
1. Expand the database for city-state mapping.
2. Add error logging for debugging.
3. Improve web scraping to handle dynamic or JavaScript-rendered websites.
4. Add support for multiple languages in voice recognition and response.

---

## Contact
For any queries or contributions, please contact **[Your Name]** at **[Your Email]**.
