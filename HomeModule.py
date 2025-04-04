import sys
import os
import traceback
import asyncio
import aiohttp
import requests
import yfinance as yf
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QWidget, QLabel
)
from PyQt5.QtCore import Qt

try:
    from config import APIConfig, AppSettings
except ImportError:
    class APIConfig:
        OPENWEATHER = {"API_KEY": None, "BASE_URL": ""}
        ALPHAVANTAGE = {"API_KEY": None, "BASE_URL": ""}
        COMMODITIES = {"WHEAT_SYMBOL": "ZW=F", "CANOLA_SYMBOL": "RS=F"}
    class AppSettings:
        WEATHER_REFRESH_INTERVAL = 900000

class WeatherAPIErrors:
    INVALID_API_KEY = 401
    NOT_FOUND = 404
    RATE_LIMIT = 429

class AlphaVantageErrors:
    INVALID_API_KEY = 401
    RATE_LIMIT = 429

class WeatherCache:
    def __init__(self, ttl_seconds=300):
        self.cache = {}
        self.ttl = ttl_seconds

    def get(self, key):
        if key in self.cache:
            data, timestamp = self.cache[key]
            if datetime.now() - timestamp < timedelta(seconds=self.ttl):
                return data
            del self.cache[key]
        return None

    def set(self, key, value):
        self.cache[key] = (value, datetime.now())

class OpenWeatherClient:
    def __init__(self, api_key, base_url, cache=None):
        self.api_key = api_key
        self.base_url = base_url
        self.session = None
        self.cache = cache

    async def __aenter__(self):
        self.session = aiohttp.ClientSession()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        if self.session:
            await self.session.close()

    async def get_weather(self, city):
        if self.cache:
            cached_data = self.cache.get(city)
            if cached_data:
                print(f"CACHE HIT for {city}")
                return cached_data

        try:
            params = {
                'q': city,
                'appid': self.api_key,
                'units': 'metric'
            }
            async with self.session.get(self.base_url, params=params) as response:
                if response.status == 200:
                    data = await response.json()
                    if self.cache:
                        self.cache.set(city, data)
                    return data
                elif response.status == WeatherAPIErrors.INVALID_API_KEY:
                    print(f"Invalid API Key for {city}")
                elif response.status == WeatherAPIErrors.NOT_FOUND:
                    print(f"City not found: {city}")
                elif response.status == WeatherAPIErrors.RATE_LIMIT:
                    print(f"Rate limit exceeded for {city}")
                else:
                    print(f"Error {response.status} for {city}")
                return None
        except Exception as e:
            print(f"Request failed for {city}: {str(e)}")
            return None

class HomeModule(QWidget):
    def __init__(self, main_window=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setObjectName("HomeModule")
        self.weather_cache = WeatherCache()

    def _weather_icon(self, description: str) -> str:
        desc = description.lower()
        if "clear" in desc:
            return "☀️"
        elif "cloud" in desc:
            return "☁️"
        elif "rain" in desc:
            return "🌧️"
        elif "snow" in desc:
            return "❄️"
        elif "thunder" in desc:
            return "⛈️"
        elif "fog" in desc or "mist" in desc:
            return "🌫️"
        elif "drizzle" in desc:
            return "🌦️"
        else:
            return "🌡️"

    def validate_weather_config(self):
        api_key = getattr(APIConfig.OPENWEATHER, 'API_KEY', None)
        base_url = getattr(APIConfig.OPENWEATHER, 'BASE_URL', None)
        if not api_key:
            raise ValueError("OpenWeather API key is missing")
        if not base_url:
            raise ValueError("OpenWeather base URL is missing")
        return api_key, base_url

    async def _fetch_all_data(self):
        self._show_status_message("Refreshing dashboard data...", 0)
        print("DEBUG: Starting _fetch_all_data")
        tasks = [
            self.fetch_weather(),
            self.fetch_exchange_rate(),
            self.fetch_commodities()
        ]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        for task_name, result in zip(['weather', 'exchange rate', 'commodities'], results):
            if isinstance(result, Exception):
                print(f"ERROR: Failed to fetch {task_name} - {result}")
                traceback.print_exception(type(result), result, result.__traceback__)
        self._show_status_message("Dashboard data refreshed.")
        print("DEBUG: _fetch_all_data completed")

    async def fetch_weather(self):
        label = QLabel("🔄 Fetching weather...")
        timestamp = QLabel("Last updated: --:--")
        try:
            api_key, base_url = self.validate_weather_config()
        except ValueError as ve:
            label.setText(str(ve))
            return

        cities = ["Camrose,CA", "Wainwright,CA", "Killam,CA", "Provost,CA"]

        try:
            async with OpenWeatherClient(api_key, base_url, cache=self.weather_cache) as client:
                tasks = [client.get_weather(city) for city in cities]
                responses = await asyncio.gather(*tasks, return_exceptions=True)

                sorted_city_data = []
                for city, data in zip(cities, responses):
                    if isinstance(data, dict) and 'main' in data:
                        sorted_city_data.append((city, data))
                    else:
                        sorted_city_data.append((city, None))
                sorted_city_data.sort(key=lambda pair: pair[1]['main']['temp'] if pair[1] else float('-inf'), reverse=True)

                results = []
                for city, data in sorted_city_data:
                    city_name = city.split(',')[0]
                    if data is None:
                        results.append(f"❌ {city_name}: Data unavailable")
                        continue
                    try:
                        temp = data['main']['temp']
                        temp_min = data['main'].get('temp_min', temp)
                        temp_max = data['main'].get('temp_max', temp)
                        desc = data['weather'][0]['description']
                        icon = self._weather_icon(desc)
                        emoji = "🔥" if temp_max >= 30 else ("🧊" if temp_min <= -10 else "🌡️")
                        results.append(f"✅ {city_name}: {temp:.1f}°C (H:{temp_max:.1f}° / L:{temp_min:.1f}°), {desc.capitalize()} {emoji} {icon}")
                    except KeyError:
                        results.append(f"❌ {city_name}: Invalid data format")
        except Exception as e:
            print(f"ERROR: Weather fetch failed - {e}")
            traceback.print_exc()
            results = ["❌ Weather: Service unavailable"]

        now = datetime.now().strftime("%H:%M:%S")
        label.setText("\n".join(results))
        timestamp.setText(f"Last updated: {now}")

    async def fetch_exchange_rate(self):
        label = QLabel("🔄 Fetching exchange rate...")
        timestamp = QLabel("Last updated: --:--")
        api_key = getattr(APIConfig.ALPHAVANTAGE, 'API_KEY', None)
        base_url = getattr(APIConfig.ALPHAVANTAGE, 'BASE_URL', None)

        if not api_key or not base_url:
            label.setText("❌ USD-CAD: API configuration missing")
            return

        try:
            async with aiohttp.ClientSession() as session:
                params = {
                    "function": "CURRENCY_EXCHANGE_RATE",
                    "from_currency": "USD",
                    "to_currency": "CAD",
                    "apikey": api_key
                }
                async with session.get(base_url, params=params) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        rate = data.get("Realtime Currency Exchange Rate", {}).get("5. Exchange Rate")
                        if rate:
                            label.setText(f"✅ USD-CAD: {float(rate):.4f}")
                        else:
                            label.setText("❌ USD-CAD: Rate not found")
                    elif resp.status == AlphaVantageErrors.INVALID_API_KEY:
                        label.setText("❌ USD-CAD: Invalid API Key")
                    elif resp.status == AlphaVantageErrors.RATE_LIMIT:
                        label.setText("❌ USD-CAD: Rate limit exceeded")
                    else:
                        label.setText(f"❌ USD-CAD: API Error {resp.status}")
        except Exception as e:
            print(f"ERROR: fetch_exchange_rate() failed - {e}")
            traceback.print_exc()
            label.setText("❌ USD-CAD: Error")

        now = datetime.now().strftime("%H:%M:%S")
        timestamp.setText(f"Last updated: {now}")

    async def fetch_commodities(self):
        loop = asyncio.get_event_loop()
        wheat_label = QLabel("🔄 Fetching Wheat...")
        wheat_time = QLabel("Last updated: --:--")
        canola_label = QLabel("🔄 Fetching Canola...")
        canola_time = QLabel("Last updated: --:--")

        def get_price(symbol):
            try:
                ticker = yf.Ticker(symbol)
                hist = ticker.history(period="5d")
                if not hist.empty:
                    return hist['Close'].iloc[-1]
                return None
            except Exception as e:
                print(f"ERROR: get_price failed for {symbol} - {e}")
                return None

        wheat_symbol = getattr(APIConfig.COMMODITIES, 'WHEAT_SYMBOL', 'ZW=F')
        canola_symbol = getattr(APIConfig.COMMODITIES, 'CANOLA_SYMBOL', 'RS=F')

        try:
            wheat = await loop.run_in_executor(None, get_price, wheat_symbol)
            wheat_label.setText(f"✅ Wheat: ${wheat:,.2f} /bu" if wheat else "❌ Wheat: N/A")
        except Exception as e:
            print(f"ERROR: Wheat fetch failed - {e}")
            wheat_label.setText("❌ Wheat: Error")

        try:
            canola = await loop.run_in_executor(None, get_price, canola_symbol)
            canola_label.setText(f"✅ Canola: ${canola:,.2f} CAD/t" if canola else "❌ Canola: N/A")
        except Exception as e:
            print(f"ERROR: Canola fetch failed - {e}")
            canola_label.setText("❌ Canola: Error")

        now = datetime.now().strftime("%H:%M:%S")
        wheat_time.setText(f"Last updated: {now}")
        canola_time.setText(f"Last updated: {now}")
    # Should be defined below following the same pattern.

    # Add GUI layout + card creation and timers as needed.
