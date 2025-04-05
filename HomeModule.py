import sys
import os
import traceback
import asyncio
import aiohttp
import requests
import yfinance as yf
import json
import time
import math
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QWidget, QLabel, QGridLayout, QVBoxLayout, QHBoxLayout,
    QFrame, QSizePolicy, QGroupBox, QPushButton, QComboBox,
    QSpacerItem, QGraphicsDropShadowEffect, QProgressBar
)
from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread, QSize, QTimer, QRect
from PyQt5.QtGui import QColor, QFont, QIcon, QPalette, QPainter, QBrush, QPen, QLinearGradient

# Directly include API keys to ensure they're available
OPENWEATHER_API_KEY = "711ac00142aa78e1807ce84a8bf1582b"
OPENWEATHER_BASE_URL = "https://api.openweathermap.org/data/2.5/weather"
ALPHAVANTAGE_API_KEY = "PHNW69I8KX24I5PT"
ALPHAVANTAGE_BASE_URL = "https://www.alphavantage.co/query"
WHEAT_SYMBOL = "ZW=F"
BITCOIN_SYMBOL = "BTC-USD"  # Added Bitcoin symbol
# Expanded list of potential canola symbols to try
CANOLA_SYMBOLS = [
    "RS=F",           # Original symbol
    "CAD=F",          # Previously tried alternative
    "ICE:RS1!",       # ICE Canola Front Month
    "RSX23.NYM",      # Specific contract month format
    "ICE:RS",         # ICE Canola general
    "CED.TO",         # Canola ETF on Toronto exchange
    "CA:RS*1",        # Another variation
    "RSK23.NYM"       # Yet another contract format
]
# Default canola symbol to try first
CANOLA_SYMBOL = CANOLA_SYMBOLS[0]

# Constants for refresh intervals
WEATHER_REFRESH_HOURS = 1  # Refresh weather every hour
EXCHANGE_REFRESH_HOURS = 6  # Refresh exchange rate every 6 hours (4 times per day)
COMMODITIES_REFRESH_HOURS = 4  # Refresh commodities every 4 hours (6 times per day)

# Define data cache file paths
CACHE_DIR = "cache"
WEATHER_CACHE = os.path.join(CACHE_DIR, "weather_cache.json")
EXCHANGE_CACHE = os.path.join(CACHE_DIR, "exchange_cache.json")
COMMODITIES_CACHE = os.path.join(CACHE_DIR, "commodities_cache.json")

# UI Style Constants
COLORS = {
    "primary": "#367C2B",      # John Deere Green
    "secondary": "#FFDE00",    # John Deere Yellow
    "accent": "#2A5D24",       # Darker Green
    "background": "#F8F9FA",   # Light Gray
    "card_bg": "#FFFFFF",      # White
    "text_primary": "#333333", # Dark Gray
    "text_secondary": "#777777", # Medium Gray
    "success": "#28A745",      # Success Green
    "warning": "#FFC107",      # Warning Yellow
    "danger": "#DC3545",       # Danger Red
    "info": "#17A2B8",         # Info Blue
    "bitcoin": "#F7931A"       # Bitcoin Orange
}

# Import config if possible
try:
    from config import APIConfig, AppSettings
    print("Successfully imported config module")
    # Try to use config values but fall back to hardcoded values
    try:
        OPENWEATHER_API_KEY = APIConfig.OPENWEATHER["API_KEY"] or OPENWEATHER_API_KEY
        OPENWEATHER_BASE_URL = APIConfig.OPENWEATHER["BASE_URL"] or OPENWEATHER_BASE_URL
        ALPHAVANTAGE_API_KEY = APIConfig.ALPHAVANTAGE["API_KEY"] or ALPHAVANTAGE_API_KEY
        ALPHAVANTAGE_BASE_URL = APIConfig.ALPHAVANTAGE["BASE_URL"] or ALPHAVANTAGE_BASE_URL
        WHEAT_SYMBOL = APIConfig.COMMODITIES["WHEAT_SYMBOL"] or WHEAT_SYMBOL
        if hasattr(APIConfig.COMMODITIES, "CANOLA_SYMBOL") and APIConfig.COMMODITIES["CANOLA_SYMBOL"]:
            CANOLA_SYMBOL = APIConfig.COMMODITIES["CANOLA_SYMBOL"]
            # Add configured symbol to the start of our list if not already there
            if CANOLA_SYMBOL not in CANOLA_SYMBOLS:
                CANOLA_SYMBOLS.insert(0, CANOLA_SYMBOL)
    except (AttributeError, KeyError) as e:
        print(f"Warning: Some config values were missing: {e}")
except ImportError as e:
    print(f"WARNING: Could not import config module: {e}")
    print("Using hardcoded API keys instead")

print(f"Using OpenWeather API Key: {OPENWEATHER_API_KEY[:5]}...{OPENWEATHER_API_KEY[-5:]}")
print(f"Using AlphaVantage API Key: {ALPHAVANTAGE_API_KEY[:5]}...{ALPHAVANTAGE_API_KEY[-5:]}")
print(f"Available canola symbols to try: {CANOLA_SYMBOLS}")

class WeatherAPIErrors:
    INVALID_API_KEY = 401
    NOT_FOUND = 404
    RATE_LIMIT = 429

class AlphaVantageErrors:
    INVALID_API_KEY = 401
    RATE_LIMIT = 429

class DataCache:
    @staticmethod
    def ensure_cache_dir():
        if not os.path.exists(CACHE_DIR):
            os.makedirs(CACHE_DIR)
    
    @staticmethod
    def save_to_cache(data, cache_file):
        DataCache.ensure_cache_dir()
        try:
            with open(cache_file, 'w') as f:
                json.dump(data, f)
            print(f"Data saved to {cache_file}")
        except Exception as e:
            print(f"Error saving cache to {cache_file}: {e}")
    
    @staticmethod
    def load_from_cache(cache_file):
        try:
            if os.path.exists(cache_file):
                with open(cache_file, 'r') as f:
                    data = json.load(f)
                print(f"Data loaded from {cache_file}")
                return data
            else:
                print(f"Cache file {cache_file} not found")
        except Exception as e:
            print(f"Error loading cache from {cache_file}: {e}")
        return None
    
    @staticmethod
    def is_cache_expired(cache_file, hours):
        """Check if cache file is older than specified hours"""
        if not os.path.exists(cache_file):
            return True
        
        file_time = os.path.getmtime(cache_file)
        file_datetime = datetime.fromtimestamp(file_time)
        now = datetime.now()
        
        # Calculate if file is older than the specified hours
        return (now - file_datetime) > timedelta(hours=hours)

class HomeSignals(QObject):
    # Define signals for async operations
    weather_ready = pyqtSignal(list, str)
    exchange_ready = pyqtSignal(str, str, float)
    wheat_ready = pyqtSignal(str, str, float)
    canola_ready = pyqtSignal(str, str, float)
    bitcoin_ready = pyqtSignal(str, str, float)
    status_update = pyqtSignal(str)
    refresh_complete = pyqtSignal()

# Thread to handle data fetching
class DataFetcherThread(QThread):
    def __init__(self, home_module, fetch_weather=True, fetch_exchange=True, fetch_commodities=True):
        super().__init__()
        self.home_module = home_module
        self.fetch_weather = fetch_weather
        self.fetch_exchange = fetch_exchange
        self.fetch_commodities = fetch_commodities
        
    def run(self):
        # Create and start a new event loop in this thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(self.home_module._fetch_all_data(
                self.fetch_weather, self.fetch_exchange, self.fetch_commodities
            ))
        except Exception as e:
            print(f"Error in thread: {e}")
            traceback.print_exc()
        finally:
            loop.close()

# Thread for scheduled refresh
class ScheduledRefreshThread(QThread):
    refresh_signal = pyqtSignal(bool, bool, bool)  # weather, exchange, commodities
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.running = True
        self.last_weather_refresh = datetime.min
        self.last_exchange_refresh = datetime.min
        self.last_commodities_refresh = datetime.min
        
    def run(self):
        while self.running:
            now = datetime.now()
            
            # Check if it's time to refresh each data type
            refresh_weather = (now - self.last_weather_refresh) > timedelta(hours=WEATHER_REFRESH_HOURS)
            refresh_exchange = (now - self.last_exchange_refresh) > timedelta(hours=EXCHANGE_REFRESH_HOURS)
            refresh_commodities = (now - self.last_commodities_refresh) > timedelta(hours=COMMODITIES_REFRESH_HOURS)
            
            # If any data needs refreshing, emit the signal
            if refresh_weather or refresh_exchange or refresh_commodities:
                self.refresh_signal.emit(refresh_weather, refresh_exchange, refresh_commodities)
                
                # Update last refresh times for what will be refreshed
                if refresh_weather:
                    self.last_weather_refresh = now
                if refresh_exchange:
                    self.last_exchange_refresh = now
                if refresh_commodities:
                    self.last_commodities_refresh = now
            
            # Sleep for a while before checking again (1 minute)
            time.sleep(60)
    
    def stop(self):
        self.running = False

# Custom styled card widget
class StyledCard(QFrame):
    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet(f"""
            StyledCard {{
                background-color: {COLORS['card_bg']};
                border-radius: 8px;
                border: 1px solid #e0e0e0;
            }}
        """)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Add shadow effect
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 25))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)
        
        # Create layout
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(15, 15, 15, 15)
        
        # Title with styled font
        self.title_label = QLabel(title)
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        self.title_label.setStyleSheet(f"color: {COLORS['accent']};")
        
        # Content area
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(0, 10, 0, 0)
        
        # Add to main layout
        self.layout.addWidget(self.title_label)
        self.layout.addWidget(self.content_widget)
    
    def add_content(self, widget):
        self.content_layout.addWidget(widget)
    
    def add_footer(self, widget):
        widget.setAlignment(Qt.AlignRight)
        widget.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt;")
        self.layout.addWidget(widget)

# Custom styled button
class StyledButton(QPushButton):
    def __init__(self, text, icon_name=None, is_primary=True, parent=None):
        super().__init__(text, parent)
        
        # Set colors based on button type
        if is_primary:
            bg_color = COLORS['primary']
            hover_color = COLORS['accent']
            text_color = "white"
        else:
            bg_color = COLORS['secondary']
            hover_color = "#E5C700"  # Darker yellow
            text_color = COLORS['accent']
        
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {bg_color};
                color: {text_color};
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
            QPushButton:pressed {{
                background-color: {hover_color};
                padding-top: 9px;
                padding-bottom: 7px;
            }}
        """)
        
        # Add icon if provided
        if icon_name:
            icon_path = f"icons/{icon_name}.png"  # Adjust path as needed
            if os.path.exists(icon_path):
                self.setIcon(QIcon(icon_path))
                self.setIconSize(QSize(18, 18))

# Weather icon widget
class WeatherIconWidget(QWidget):
    def _draw_thermometer(self, painter):
        """Draw a thermometer icon for temperature display"""
        painter.setPen(QPen(QColor(200, 0, 0), 2))
        painter.drawLine(30, 15, 30, 45)
        painter.setBrush(QBrush(QColor(200, 0, 0)))
        painter.drawEllipse(25, 45, 10, 10)
        painter.drawRect(27, 25, 6, 20)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.description = "clear"
        self.temp = 0
        self.setMinimumSize(60, 60)
        self.setMaximumSize(60, 60)
    
    def update_weather(self, description, temp):
        self.description = description.lower()
        self.temp = temp
        self.update()
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Draw background
        bg_color = QColor(240, 240, 240)
        if self.temp < -10:
            bg_color = QColor(200, 220, 255)  # Cold blue
        elif self.temp > 30:
            bg_color = QColor(255, 220, 200)  # Hot red
            
        painter.setBrush(QBrush(bg_color))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(5, 5, 50, 50)
        
        # Draw weather icon
        if "clear" in self.description:
            self._draw_sun(painter)
        elif "cloud" in self.description:
            self._draw_cloud(painter)
        elif "rain" in self.description:
            self._draw_rain(painter)
        elif "snow" in self.description:
            self._draw_snow(painter)
        elif "thunder" in self.description:
            self._draw_thunder(painter)
        elif "fog" in self.description or "mist" in self.description:
            self._draw_fog(painter)
        else:
            self._draw_thermometer(painter)

# Price trend indicator widget
class TrendIndicator(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.trend = 0  # -1: down, 0: stable, 1: up
        self.percentage = 0.0
        self.setMinimumSize(80, 30)
        self.setMaximumSize(80, 30)
    
    def set_trend(self, trend, percentage=0.0):
        self.trend = trend
        self.percentage = percentage
        self.update()
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Choose color based on trend
        if self.trend > 0:
            color = QColor(COLORS['success'])
        elif self.trend < 0:
            color = QColor(COLORS['danger'])
        else:
            color = QColor(COLORS['text_secondary'])
        
        painter.setPen(QPen(color, 2))
        
        # Draw trend arrow
        if self.trend > 0:
            # Up arrow
            painter.drawLine(15, 20, 25, 10)
            painter.drawLine(25, 10, 35, 20)
        elif self.trend < 0:
            # Down arrow
            painter.drawLine(15, 10, 25, 20)
            painter.drawLine(25, 20, 35, 10)
        else:
            # Horizontal line
            painter.drawLine(15, 15, 35, 15)
        
        # Draw percentage text
        if self.percentage != 0:
            painter.setPen(QPen(color, 1))
            sign = "+" if self.trend > 0 else ""
            if self.trend == 0:
                sign = "±"
            painter.drawText(40, 20, f"{sign}{self.percentage:.1f}%")

# Circular progress gauge for visual representation
class CircularProgressGauge(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.value = 0
        self.maximum = 100
        self.color = QColor(COLORS['primary'])
        self.setMinimumSize(100, 100)
        self.setMaximumSize(100, 100)
        
    def set_value(self, value, maximum=None):
        self.value = value
        if maximum is not None:
            self.maximum = maximum
        self.update()
        
    def set_color(self, color):
        self.color = QColor(color)
        self.update()
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Draw background circle
        painter.setPen(QPen(QColor(230, 230, 230), 10))
        painter.drawEllipse(10, 10, 80, 80)
        
        # Draw progress arc
        if self.maximum > 0:
            angle = 360 * self.value / self.maximum
            painter.setPen(QPen(self.color, 10))
            # FIXED: Use QRect instead of passing individual parameters and convert angle to int
            angle_int = int(angle)
            painter.drawArc(QRect(10, 10, 80, 80), 90 * 16, -angle_int * 16)
        
        # Draw value text
        painter.setPen(QPen(QColor(COLORS['text_primary']), 1))
        font = painter.font()
        font.setPointSize(14)
        font.setBold(True)
        painter.setFont(font)
        
        if isinstance(self.value, float):
            text = f"{self.value:.1f}"
        else:
            text = str(self.value)
            
        x_position = int(50 - painter.fontMetrics().width(text) / 2)
        painter.drawText(x_position, 55, text)
class HomeModule(QWidget):
    def __init__(self, main_window=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setObjectName("HomeModule")
        
        # Set background color
        self.setStyleSheet(f"background-color: {COLORS['background']};")
        
        # Create signals for thread-safe UI updates
        self.signals = HomeSignals()
        
        # Create labels for UI elements
        self.status_label = QLabel("Dashboard ready")
        self.status_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-style: italic;")
        
        # Weather widgets
        self.weather_widgets = {}
        for city in ["Camrose", "Wainwright", "Killam", "Provost"]:
            self.weather_widgets[city] = {
                'icon': WeatherIconWidget(),
                'label': QLabel(f"{city}: Loading..."),
                'layout': QHBoxLayout()
            }
            self.weather_widgets[city]['label'].setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
            self.weather_widgets[city]['label'].setWordWrap(True)
            
            # Arrange horizontally
            self.weather_widgets[city]['layout'].addWidget(self.weather_widgets[city]['icon'])
            self.weather_widgets[city]['layout'].addWidget(self.weather_widgets[city]['label'], 1)
        
        self.weather_timestamp = QLabel("Last updated: --:--")
        
        # Exchange rate widgets
        self.exchange_label = QLabel("USD-CAD: Loading...")
        self.exchange_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
        self.exchange_timestamp = QLabel("Last updated: --:--")
        self.exchange_gauge = CircularProgressGauge()
        self.exchange_gauge.set_color(COLORS['info'])
        self.exchange_trend = TrendIndicator()
        
        # Wheat price widgets
        self.wheat_label = QLabel("Wheat: Loading...")
        self.wheat_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
        self.wheat_timestamp = QLabel("Last updated: --:--")
        self.wheat_gauge = CircularProgressGauge()
        self.wheat_gauge.set_color(COLORS['success'])
        self.wheat_trend = TrendIndicator()
        
        # Canola price widgets
        self.canola_label = QLabel("Canola: Loading...")
        self.canola_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
        self.canola_timestamp = QLabel("Last updated: --:--")
        self.canola_gauge = CircularProgressGauge()
        self.canola_gauge.set_color(COLORS['secondary'])
        self.canola_trend = TrendIndicator()
        
        # Bitcoin price widgets
        self.bitcoin_label = QLabel("Bitcoin: Loading...")
        self.bitcoin_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
        self.bitcoin_timestamp = QLabel("Last updated: --:--")
        self.bitcoin_gauge = CircularProgressGauge()
        self.bitcoin_gauge.set_color(COLORS['bitcoin'])
        self.bitcoin_trend = TrendIndicator()
        
        # Refresh options dropdown (now more subtle)
        self.refresh_combo = QComboBox()
        self.refresh_combo.addItem("Refresh All")
        self.refresh_combo.addItem("Refresh Weather Only")
        self.refresh_combo.addItem("Refresh Exchange Rate Only")
        self.refresh_combo.addItem("Refresh Commodities Only")
        self.refresh_combo.setStyleSheet(f"""
            QComboBox {{
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
                background-color: white;
                min-width: 150px;
                font-size: 10pt;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 20px;
            }}
        """)
        
        # Setup refresh animation timer - MOVED BEFORE _refresh_data() call
        self.refresh_animation_timer = QTimer(self)
        self.refresh_animation_timer.timeout.connect(self._update_refresh_animation)
        self.refresh_animation_value = 0
        self.is_refreshing = False
        
        # Set up the layout
        self._setup_ui()
        
        # Connect signals to update UI from async operations
        self.signals.weather_ready.connect(self._update_weather_ui)
        self.signals.exchange_ready.connect(self._update_exchange_ui)
        self.signals.wheat_ready.connect(self._update_wheat_ui)
        self.signals.canola_ready.connect(self._update_canola_ui)
        self.signals.bitcoin_ready.connect(self._update_bitcoin_ui)
        self.signals.status_update.connect(self.status_label.setText)
        self.signals.refresh_complete.connect(self._on_refresh_complete)
        
        # Set up the auto refresh thread
        self.scheduled_refresh = ScheduledRefreshThread(self)
        self.scheduled_refresh.refresh_signal.connect(self._scheduled_refresh)
        self.scheduled_refresh.start()
        
        # Initial data load
        self._load_cached_data()
        
        # Initial data fetch using a thread - AFTER timer initialization
        self._refresh_data()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        # Header with title and status
        header_layout = QHBoxLayout()
        dashboard_title = QLabel("AMS Dashboard")
        dashboard_title.setStyleSheet(f"""
            font-size: 18pt;
            font-weight: bold;
            color: {COLORS['accent']};
        """)
        header_layout.addWidget(dashboard_title)
        header_layout.addStretch()
        
        # More subtle refresh controls in the header
        refresh_layout = QHBoxLayout()
        refresh_layout.setSpacing(5)
        refresh_layout.addWidget(self.refresh_combo)
        
        # Simple link-style refresh button
        refresh_button = QPushButton("↻ Refresh")
        refresh_button.setStyleSheet(f"""
            QPushButton {{
                background: none;
                border: none;
                color: {COLORS['primary']};
                font-size: 10pt;
                text-decoration: underline;
                padding: 3px;
            }}
            QPushButton:hover {{
                color: {COLORS['accent']};
            }}
        """)
        refresh_button.clicked.connect(self._handle_refresh_button)
        refresh_layout.addWidget(refresh_button)
        
        # Add refresh controls to header
        header_layout.addLayout(refresh_layout)
        header_layout.addWidget(self.status_label)
        main_layout.addLayout(header_layout)
        
        # Create grid layout for cards with more spacing
        grid_layout = QGridLayout()
        grid_layout.setSpacing(15)
        
        # Weather card
        weather_card = StyledCard("Weather Conditions")
        weather_layout = QVBoxLayout()
        
        # Add each city's weather layout
        for city in self.weather_widgets:
            weather_layout.addLayout(self.weather_widgets[city]['layout'])
        
        # Add to the card
        weather_widget = QWidget()
        weather_widget.setLayout(weather_layout)
        weather_card.add_content(weather_widget)
        weather_card.add_footer(self.weather_timestamp)
        grid_layout.addWidget(weather_card, 0, 0, 1, 2)  # Span two columns
        
        # Exchange rate card
        exchange_card = StyledCard("Currency Exchange")
        exchange_layout = QHBoxLayout()
        exchange_text_layout = QVBoxLayout()
        exchange_text_layout.addWidget(self.exchange_label)
 
        exchange_text_layout.addWidget(self.exchange_trend)
        exchange_text_layout.addStretch()
        
        exchange_layout.addLayout(exchange_text_layout)
        exchange_layout.addWidget(self.exchange_gauge)
        
        exchange_widget = QWidget()
        exchange_widget.setLayout(exchange_layout)
        exchange_card.add_content(exchange_widget)
        exchange_card.add_footer(self.exchange_timestamp)
        grid_layout.addWidget(exchange_card, 1, 0)
        
        # Commodities layout with 2 rows, 2 columns
        commodities_layout = QGridLayout()
        commodities_layout.setSpacing(15)
        
        # Wheat card
        wheat_card = StyledCard("Wheat Price")
        wheat_layout = QHBoxLayout()
        wheat_text_layout = QVBoxLayout()
        wheat_text_layout.addWidget(self.wheat_label)
        wheat_text_layout.addWidget(self.wheat_trend)
        wheat_text_layout.addStretch()
        
        wheat_layout.addLayout(wheat_text_layout)
        wheat_layout.addWidget(self.wheat_gauge)
        
        wheat_widget = QWidget()
        wheat_widget.setLayout(wheat_layout)
        wheat_card.add_content(wheat_widget)
        wheat_card.add_footer(self.wheat_timestamp)
        commodities_layout.addWidget(wheat_card, 0, 0)
        
        # Canola card
        canola_card = StyledCard("Canola Price")
        canola_layout = QHBoxLayout()
        canola_text_layout = QVBoxLayout()
        canola_text_layout.addWidget(self.canola_label)
        canola_text_layout.addWidget(self.canola_trend)
        canola_text_layout.addStretch()
        
        canola_layout.addLayout(canola_text_layout)
        canola_layout.addWidget(self.canola_gauge)
        
        canola_widget = QWidget()
        canola_widget.setLayout(canola_layout)
        canola_card.add_content(canola_widget)
        canola_card.add_footer(self.canola_timestamp)
        commodities_layout.addWidget(canola_card, 0, 1)
        
        # Bitcoin card
        bitcoin_card = StyledCard("Bitcoin Price")
        bitcoin_layout = QHBoxLayout()
        bitcoin_text_layout = QVBoxLayout()
        bitcoin_text_layout.addWidget(self.bitcoin_label)
        bitcoin_text_layout.addWidget(self.bitcoin_trend)
        bitcoin_text_layout.addStretch()
        
        bitcoin_layout.addLayout(bitcoin_text_layout)
        bitcoin_layout.addWidget(self.bitcoin_gauge)
        
        bitcoin_widget = QWidget()
        bitcoin_widget.setLayout(bitcoin_layout)
        bitcoin_card.add_content(bitcoin_widget)
        bitcoin_card.add_footer(self.bitcoin_timestamp)
        commodities_layout.addWidget(bitcoin_card, 1, 0)
        
        # Add a small card with canola refresh button
        canola_refresh_card = StyledCard("Canola Data Fix")
        canola_refresh_layout = QVBoxLayout()
        canola_refresh_layout.setContentsMargins(0, 0, 0, 0)
        
        canola_refresh_info = QLabel("If canola prices aren't showing, try forcing a refresh")
        canola_refresh_info.setWordWrap(True)
        canola_refresh_info.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 10pt;")
        canola_refresh_layout.addWidget(canola_refresh_info)
        
        canola_refresh_button = StyledButton("Force Canola Refresh", is_primary=False)
        canola_refresh_button.clicked.connect(self.force_refresh_commodities)
        canola_refresh_layout.addWidget(canola_refresh_button)
        
        canola_refresh_widget = QWidget()
        canola_refresh_widget.setLayout(canola_refresh_layout)
        canola_refresh_card.add_content(canola_refresh_widget)
        commodities_layout.addWidget(canola_refresh_card, 1, 1)
        
        # Add commodities grid to main grid
        commodities_container = QWidget()
        commodities_container.setLayout(commodities_layout)
        grid_layout.addWidget(commodities_container, 1, 1, 2, 1)
        
        main_layout.addLayout(grid_layout)
        main_layout.addStretch()

    def _handle_refresh_button(self):
        """Handle refresh button click based on selected option"""
        option = self.refresh_combo.currentText()
        
        if option == "Refresh All":
            self._refresh_data(True, True, True)
        elif option == "Refresh Weather Only":
            self._refresh_data(True, False, False)
        elif option == "Refresh Exchange Rate Only":
            self._refresh_data(False, True, False)
        elif option == "Refresh Commodities Only":
            self._refresh_data(False, False, True)
            
        # Start refresh animation
        self.is_refreshing = True
        self.refresh_animation_timer.start(100)
    
    def _update_refresh_animation(self):
        """Update the animation during refresh"""
        if self.is_refreshing:
            self.refresh_animation_value = (self.refresh_animation_value + 10) % 100
            self.status_label.setText(f"Refreshing data {'.' * (self.refresh_animation_value // 25 + 1)}")
    
    def _scheduled_refresh(self, weather, exchange, commodities):
        """Handle scheduled refreshes from the timer thread"""
        self._refresh_data(weather, exchange, commodities)
    
    def force_refresh_commodities(self):
        """Force refresh of commodity data by bypassing cache check"""
        print("Forcing refresh of commodity data...")
        
        # Delete the commodities cache file if it exists
        if os.path.exists(COMMODITIES_CACHE):
            try:
                os.remove(COMMODITIES_CACHE)
                print(f"Deleted commodities cache file: {COMMODITIES_CACHE}")
            except Exception as e:
                print(f"Error deleting cache file: {e}")
        
        # Start a new thread to fetch ONLY commodities data
        self.fetcher_thread = DataFetcherThread(
            self, 
            fetch_weather=False, 
            fetch_exchange=False, 
            fetch_commodities=True
        )
        self.fetcher_thread.start()
        
        # Start refresh animation
        self.is_refreshing = True
        self.refresh_animation_timer.start(100)
        self.signals.status_update.emit("Forcing commodity data refresh...")
            
    # UI update slots connected to signals
    def _update_weather_ui(self, results, timestamp):
        """Update the weather UI with the fetched data"""
        # Process weather results to extract data for each city
        for result in results:
            for city in self.weather_widgets:
                if result.startswith(f"✅ {city}:"):
                    # Parse the weather data
                    # Format: "✅ City: temp°C (H:max° / L:min°), Description Emoji Icon"
                    parts = result.split(',')
                    temp_part = parts[0].split(':')[1].strip()
                    temp_value = float(temp_part.split('°')[0])
                    
                    description_part = parts[1].strip()
                    description = description_part.split(' ')[0]  # Just get the description word
                    
                    # Update the icon
                    self.weather_widgets[city]['icon'].update_weather(description, temp_value)
                    
                    # Style the label
                    styled_text = f"<span style='color:{COLORS['success']}; font-weight:bold'>{city}:</span> {temp_part} <br>{description_part}"
                    self.weather_widgets[city]['label'].setText(styled_text)
                    break
                elif result.startswith(f"❌ {city}:"):
                    # Handle error case
                    styled_text = f"<span style='color:{COLORS['danger']}; font-weight:bold'>{city}:</span> {result.split(':', 1)[1]}"
                    self.weather_widgets[city]['label'].setText(styled_text)
                    break
        
        self.weather_timestamp.setText(f"Last updated: {timestamp}")
        
    def _update_exchange_ui(self, rate_text, timestamp, rate_value=None):
        # Style the exchange rate text
        if rate_text.startswith("✅"):
            styled_text = f"<span style='color:{COLORS['success']}; font-size:14pt; font-weight:bold'>{rate_text}</span>"
            # Update the gauge if we have a value
            if rate_value is not None:
                self.exchange_gauge.set_value(rate_value)
                # Set trend based on typical USD-CAD range (1.25-1.45)
                if rate_value > 1.40:
                    self.exchange_trend.set_trend(1, (rate_value - 1.35) * 20)  # USD strong
                elif rate_value < 1.30:
                    self.exchange_trend.set_trend(-1, (1.35 - rate_value) * 20)  # CAD strong
                else:
                    self.exchange_trend.set_trend(0)  # Neutral
        else:
            styled_text = f"<span style='color:{COLORS['danger']}'>{rate_text}</span>"
            
        self.exchange_label.setText(styled_text)
        self.exchange_timestamp.setText(f"Last updated: {timestamp}")
        
    def _update_wheat_ui(self, wheat_text, timestamp, price_value=None):
        # Style the wheat price text
        if wheat_text.startswith("✅"):
            styled_text = f"<span style='color:{COLORS['success']}; font-size:14pt; font-weight:bold'>{wheat_text}</span>"
            # Update the gauge if we have a value
            if price_value is not None:
                self.wheat_gauge.set_value(price_value, 1000)  # Max around 1000
                # Set trend - this would ideally use historical data
                # For now use 500 as a baseline
                if price_value > 550:
                    self.wheat_trend.set_trend(1, (price_value - 500) / 5)
                elif price_value < 450:
                    self.wheat_trend.set_trend(-1, (500 - price_value) / 5)
                else:
                    self.wheat_trend.set_trend(0)
        else:
            styled_text = f"<span style='color:{COLORS['danger']}'>{wheat_text}</span>"
            
        self.wheat_label.setText(styled_text)
        self.wheat_timestamp.setText(f"Last updated: {timestamp}")
        
    def _update_canola_ui(self, canola_text, timestamp, price_value=None):
        # Style the canola price text
        if canola_text.startswith("✅"):
            styled_text = f"<span style='color:{COLORS['success']}; font-size:14pt; font-weight:bold'>{canola_text}</span>"
            # Update the gauge if we have a value
            if price_value is not None:
                self.canola_gauge.set_value(price_value, 1000)  # Max around 1000
                # Set trend
                if price_value > 700:
                    self.canola_trend.set_trend(1, (price_value - 650) / 10)
                elif price_value < 600:
                    self.canola_trend.set_trend(-1, (650 - price_value) / 10)
                else:
                    self.canola_trend.set_trend(0)
        elif canola_text.startswith("ℹ️"):
            styled_text = f"<span style='color:{COLORS['info']}; font-size:14pt; font-weight:bold'>{canola_text}</span>"
            # Update the gauge if we have a value (estimated)
            if price_value is not None:
                self.canola_gauge.set_value(price_value, 1000)
                self.canola_trend.set_trend(0)
        else:
            styled_text = f"<span style='color:{COLORS['danger']}'>{canola_text}</span>"
            
        self.canola_label.setText(styled_text)
        self.canola_timestamp.setText(f"Last updated: {timestamp}")
        
    def _update_bitcoin_ui(self, bitcoin_text, timestamp, price_value=None):
        # Style the bitcoin price text
        if bitcoin_text.startswith("✅"):
            styled_text = f"<span style='color:{COLORS['bitcoin']}; font-size:14pt; font-weight:bold'>{bitcoin_text}</span>"
            # Update the gauge if we have a value
            if price_value is not None:
                max_value = 100000  # Max BTC price for gauge scale
                self.bitcoin_gauge.set_value(price_value/1000, max_value/1000)  # Show in thousands
                # Set trend (this should use historical data ideally)
                if price_value > 65000:
                    self.bitcoin_trend.set_trend(1, (price_value - 60000) / 1000)
                elif price_value < 55000:
                    self.bitcoin_trend.set_trend(-1, (60000 - price_value) / 1000)
                else:
                    self.bitcoin_trend.set_trend(0)
        else:
            styled_text = f"<span style='color:{COLORS['danger']}'>{bitcoin_text}</span>"
            
        self.bitcoin_label.setText(styled_text)
        self.bitcoin_timestamp.setText(f"Last updated: {timestamp}")
    
    def _on_refresh_complete(self):
        """Handle completion of refresh"""
        self.is_refreshing = False
        self.refresh_animation_timer.stop()
        self.signals.status_update.emit("Dashboard data refreshed")

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

    def _load_cached_data(self):
        """Load data from cache files on startup"""
        # Load weather data
        weather_cache = DataCache.load_from_cache(WEATHER_CACHE)
        if weather_cache:
            self.signals.weather_ready.emit(weather_cache["results"], weather_cache["timestamp"])
        
        # Load exchange rate data
        exchange_cache = DataCache.load_from_cache(EXCHANGE_CACHE)
        if exchange_cache:
            # Extract rate value from text if available
            rate_value = None
            if "rate_text" in exchange_cache and exchange_cache["rate_text"].startswith("✅"):
                try:
                    rate_value = float(exchange_cache["rate_text"].split(":")[-1].strip())
                except:
                    rate_value = None
            
            self.signals.exchange_ready.emit(
                exchange_cache["rate_text"], 
                exchange_cache["timestamp"],
                rate_value
            )
        
        # Load commodities data
        commodities_cache = DataCache.load_from_cache(COMMODITIES_CACHE)
        if commodities_cache:
            # Extract wheat price value if available
            wheat_value = None
            if "wheat_text" in commodities_cache and commodities_cache["wheat_text"].startswith("✅"):
                try:
                    wheat_value = float(commodities_cache["wheat_text"].split("$")[1].split(" ")[0].replace(',', ''))
                except:
                    wheat_value = None
            
            # Extract canola price value if available
            canola_value = None
            if "canola_price" in commodities_cache:
                canola_value = commodities_cache["canola_price"]
            elif "canola_text" in commodities_cache and (commodities_cache["canola_text"].startswith("✅") or commodities_cache["canola_text"].startswith("ℹ️")):
                try:
                    canola_value = float(commodities_cache["canola_text"].split("$")[1].split(" ")[0].replace(',', ''))
                except:
                    canola_value = None
                    
            self.signals.wheat_ready.emit(
                commodities_cache["wheat_text"], 
                commodities_cache["timestamp"],
                wheat_value
            )
            
            self.signals.canola_ready.emit(
                commodities_cache["canola_text"], 
                commodities_cache["timestamp"],
                canola_value
            )
            
            # If we have bitcoin data in cache (for subsequent runs)
            if "bitcoin_text" in commodities_cache and "bitcoin_price" in commodities_cache:
                self.signals.bitcoin_ready.emit(
                    commodities_cache["bitcoin_text"],
                    commodities_cache["timestamp"],
                    commodities_cache["bitcoin_price"]
                )

    def _refresh_data(self, fetch_weather=True, fetch_exchange=True, fetch_commodities=True):
        # Update status label with animation
        self.is_refreshing = True
        self.refresh_animation_timer.start(100)
        
        # Check cache expiration before fetching data
        if fetch_weather and not DataCache.is_cache_expired(WEATHER_CACHE, WEATHER_REFRESH_HOURS):
            fetch_weather = False
            print("Weather cache is still valid, skipping fetch")
            
        if fetch_exchange and not DataCache.is_cache_expired(EXCHANGE_CACHE, EXCHANGE_REFRESH_HOURS):
            fetch_exchange = False
            print("Exchange rate cache is still valid, skipping fetch")
            
        if fetch_commodities and not DataCache.is_cache_expired(COMMODITIES_CACHE, COMMODITIES_REFRESH_HOURS):
            fetch_commodities = False
            print("Commodities cache is still valid, skipping fetch")
        
        # Only create and run thread if there's something to fetch
        if fetch_weather or fetch_exchange or fetch_commodities:
            # Create and run thread for data fetching
            self.fetcher_thread = DataFetcherThread(
                self, 
                fetch_weather=fetch_weather, 
                fetch_exchange=fetch_exchange, 
                fetch_commodities=fetch_commodities
            )
            self.fetcher_thread.start()
        else:
            self.is_refreshing = False
            self.refresh_animation_timer.stop()
            self.signals.status_update.emit("Using cached data (no refresh needed)")

    async def _fetch_all_data(self, fetch_weather=True, fetch_exchange=True, fetch_commodities=True):
        print("DEBUG: Starting _fetch_all_data")
        tasks = []
        
        if fetch_weather:
            tasks.append(self.fetch_weather())
        if fetch_exchange:
            tasks.append(self.fetch_exchange_rate())
        if fetch_commodities:
            tasks.append(self.fetch_commodities())
        
        if tasks:  # Only run tasks if there are any
            results = await asyncio.gather(*tasks, return_exceptions=True)
            for i, result in enumerate(results):
                if isinstance(result, Exception):
                    task_name = "weather" if i == 0 else "exchange rate" if i == 1 else "commodities"
                    print(f"ERROR: Failed to fetch {task_name} - {result}")
                    traceback.print_exception(type(result), result, result.__traceback__)
        
        self.signals.refresh_complete.emit()
        print("DEBUG: _fetch_all_data completed")

    async def fetch_weather(self):
        print(f"Fetching weather with API key: {OPENWEATHER_API_KEY}")
        
        if not OPENWEATHER_API_KEY:
            self.signals.weather_ready.emit(["OpenWeather API key is missing"], datetime.now().strftime("%H:%M:%S"))
            return

        cities = ["Camrose,CA", "Wainwright,CA", "Killam,CA", "Provost,CA"]
        results = []

        try:
            # Create a direct session instead of using the context manager
            session = aiohttp.ClientSession()
            
            for city in cities:
                try:
                    params = {
                        'q': city,
                        'appid': OPENWEATHER_API_KEY,
                        'units': 'metric'
                    }
                    print(f"Requesting weather for {city}")
                    async with session.get(OPENWEATHER_BASE_URL, params=params) as response:
                        print(f"Response status for {city}: {response.status}")
                        
                        if response.status == 200:
                            data = await response.json()
                            print(f"Weather data for {city}: {data}")
                            
                            try:
                                city_name = city.split(',')[0]
                                temp = data['main']['temp']
                                temp_min = data['main'].get('temp_min', temp)
                                temp_max = data['main'].get('temp_max', temp)
                                desc = data['weather'][0]['description']
                                icon = self._weather_icon(desc)
                                emoji = "🔥" if temp_max >= 30 else ("🧊" if temp_min <= -10 else "🌡️")
                                results.append(f"✅ {city_name}: {temp:.1f}°C (H:{temp_max:.1f}° / L:{temp_min:.1f}°), {desc.capitalize()} {emoji} {icon}")
                            except KeyError as ke:
                                print(f"KeyError for {city}: {ke}")
                                results.append(f"❌ {city_name}: Invalid data format")
                        else:
                            city_name = city.split(',')[0]
                            if response.status == WeatherAPIErrors.INVALID_API_KEY:
                                error_msg = "Invalid API Key"
                            elif response.status == WeatherAPIErrors.NOT_FOUND:
                                error_msg = "City not found"
                            elif response.status == WeatherAPIErrors.RATE_LIMIT:
                                error_msg = "Rate limit exceeded"
                            else:
                                error_msg = f"API Error {response.status}"
                            
                            results.append(f"❌ {city_name}: {error_msg}")
                except Exception as e:
                    city_name = city.split(',')[0]
                    print(f"Error fetching weather for {city}: {e}")
                    results.append(f"❌ {city_name}: {str(e)}")
            
            # Close the session when done
            await session.close()
                
        except Exception as e:
            print(f"ERROR: Weather fetch failed - {e}")
            traceback.print_exc()
            results = [f"❌ Weather: Service unavailable ({str(e)})"]

        # Use signals to update UI
        now = datetime.now().strftime("%H:%M:%S")
        self.signals.weather_ready.emit(results, now)
        
        # Cache the data
        weather_cache = {
            "results": results,
            "timestamp": now,
            "fetch_time": datetime.now().isoformat()
        }
        DataCache.save_to_cache(weather_cache, WEATHER_CACHE)

    async def fetch_exchange_rate(self):
        now = datetime.now().strftime("%H:%M:%S")
        print(f"Fetching exchange rate with API key: {ALPHAVANTAGE_API_KEY}")

        if not ALPHAVANTAGE_API_KEY:
            self.signals.exchange_ready.emit("❌ USD-CAD: API configuration missing", now, None)
            return

        try:
            # Manually handle URL and parameters to troubleshoot potential issues
            url = f"{ALPHAVANTAGE_BASE_URL}?function=CURRENCY_EXCHANGE_RATE&from_currency=USD&to_currency=CAD&apikey={ALPHAVANTAGE_API_KEY}"
            print(f"Sending request to: {url}")
            
            async with aiohttp.ClientSession() as session:
                async with session.get(url) as resp:
                    status = resp.status
                    print(f"Response status: {status}")
                    
                    if status == 200:
                        data = await resp.json()
                        print(f"Response data: {data}")
                        
                        # Check for error message in response
                        if "Error Message" in data:
                            result_text = f"❌ USD-CAD: {data['Error Message']}"
                            rate_value = None
                        elif "Information" in data:
                            result_text = f"❌ USD-CAD: {data['Information']}"
                            rate_value = None
                        elif "Note" in data:
                            result_text = f"❌ USD-CAD: {data['Note']}"
                            rate_value = None
                        else:
                            # Try to get exchange rate
                            rate = data.get("Realtime Currency Exchange Rate", {}).get("5. Exchange Rate")
                            if rate:
                                rate_value = float(rate)
                                result_text = f"✅ USD-CAD: {rate_value:.4f}"
                            else:
                                result_text = "❌ USD-CAD: Rate not found in response"
                                rate_value = None
                    elif status == AlphaVantageErrors.INVALID_API_KEY:
                        result_text = "❌ USD-CAD: Invalid API Key"
                        rate_value = None
                    elif status == AlphaVantageErrors.RATE_LIMIT:
                        result_text = "❌ USD-CAD: Rate limit exceeded"
                        rate_value = None
                    else:
                        result_text = f"❌ USD-CAD: API Error {status}"
                        rate_value = None
        except Exception as e:
            print(f"ERROR: fetch_exchange_rate() failed - {e}")
            traceback.print_exc()
            result_text = f"❌ USD-CAD: Error ({str(e)})"
            rate_value = None

        # Use signals to update UI
        self.signals.exchange_ready.emit(result_text, now, rate_value)
        
        # Cache the data
        exchange_cache = {
            "rate_text": result_text,
            "timestamp": now,
            "fetch_time": datetime.now().isoformat(),
            "rate_value": rate_value
        }
        DataCache.save_to_cache(exchange_cache, EXCHANGE_CACHE)

    async def get_canola_price_fallback(self):
        """Fallback method to get canola price by scraping a website"""
        print("Attempting fallback method to get canola price...")
        
        try:
            # Try to get data from a website that publishes canola prices
            # This example uses the Canadian Grain Commission website
            url = "https://www.grainscanada.gc.ca/en/grain-markets/canola-prices/"
            
            async with aiohttp.ClientSession() as session:
                async with session.get(url, timeout=10) as response:
                    if response.status == 200:
                        html = await response.text()
                        
                        # This is a very simple scraping example - in reality you'd need
                        # to adjust the parsing logic to match the actual website structure
                        # Look for a price pattern (e.g., "$700.00" or similar)
                        import re
                        price_matches = re.findall(r'\$(\d+\.\d+)', html)
                        
                        if price_matches:
                            # Take the first match as our price
                            price = float(price_matches[0])
                            return price
                        else:
                            print("No price pattern found in the webpage")
                    else:
                        print(f"Failed to fetch webpage: HTTP {response.status}")
        
        except Exception as e:
            print(f"Fallback method error: {e}")
            traceback.print_exc()
        
        # If all fails, use a hardcoded recent value with a warning
        print("Using hardcoded fallback value for canola")
        # This is your hard-coded fallback price, update it occasionally
        fallback_price = 695.25
        return fallback_price

    async def fetch_commodities(self):
        loop = asyncio.get_event_loop()
        now = datetime.now().strftime("%H:%M:%S")

        # Enhanced get_price function with debugging and NoneType handling
        def get_price(symbol):
            try:
                print(f"Fetching price for {symbol}")
                ticker = yf.Ticker(symbol)
                
                # FIXED: Check if info is None before using it
                try:
                    info = ticker.info
                    if info is None:
                        print(f"Ticker info for {symbol}: None returned")
                        # Don't return here, still try to get history
                    elif len(info) > 0:
                        print(f"Ticker info for {symbol}: Found {len(info)} fields")
                    else:
                        print(f"Ticker info for {symbol}: Empty dictionary")
                except Exception as e:
                    print(f"Error accessing ticker.info for {symbol}: {e}")
                    traceback.print_exc()
                    # Continue to get history, don't return None here
                
                # Get history and check if it has data
                try:
                    hist = ticker.history(period="5d")
                    if hist is None:
                        print(f"No history data returned for {symbol}")
                        return None
                    
                    if hist.empty:
                        print(f"Empty history dataframe for {symbol}")
                        return None
                    
                    # Check if 'Close' column exists
                    if 'Close' not in hist.columns:
                        print(f"No 'Close' column in history data for {symbol}. Columns: {hist.columns}")
                        return None
                    
                    # Check if we have any data in the Close column
                    if len(hist['Close']) == 0:
                        print(f"No closing prices in history data for {symbol}")
                        return None
                    
                    # Get the latest closing price
                    price = hist['Close'].iloc[-1]
                    print(f"Price for {symbol}: {price}")
                    return price
                except Exception as e:
                    print(f"Error getting history for {symbol}: {e}")
                    traceback.print_exc()
                    return None
            except Exception as e:
                print(f"ERROR: get_price failed for {symbol} - {e}")
                traceback.print_exc()
                return None
                # First try to get wheat price
        try:
            wheat = await loop.run_in_executor(None, get_price, WHEAT_SYMBOL)
            wheat_text = f"✅ Wheat: ${wheat:,.2f} /bu" if wheat else "❌ Wheat: N/A"
            wheat_value = wheat if wheat else None
        except Exception as e:
            print(f"ERROR: Wheat fetch failed - {e}")
            traceback.print_exc()
            wheat_text = f"❌ Wheat: Error ({str(e)})"
            wheat_value = None

        # Send wheat price to UI immediately
        self.signals.wheat_ready.emit(wheat_text, now, wheat_value)
        
        # Get Bitcoin price
        try:
            bitcoin = await loop.run_in_executor(None, get_price, BITCOIN_SYMBOL)
            bitcoin_text = f"✅ Bitcoin: ${bitcoin:,.2f} USD" if bitcoin else "❌ Bitcoin: N/A"
            bitcoin_value = bitcoin if bitcoin else None
        except Exception as e:
            print(f"ERROR: Bitcoin fetch failed - {e}")
            traceback.print_exc()
            bitcoin_text = f"❌ Bitcoin: Error ({str(e)})"
            bitcoin_value = None
            
        # Send Bitcoin price to UI
        self.signals.bitcoin_ready.emit(bitcoin_text, now, bitcoin_value)
        
        # Try multiple canola symbols until one works
        canola_price = None
        canola_text = "❌ Canola: N/A"
        used_symbol = None
        
        print(f"Attempting to fetch canola price using {len(CANOLA_SYMBOLS)} different symbols")
        for symbol in CANOLA_SYMBOLS:
            try:
                print(f"Trying canola symbol: {symbol}")
                canola = await loop.run_in_executor(None, get_price, symbol)
                if canola is not None:
                    canola_price = canola
                    canola_text = f"✅ Canola: ${canola:,.2f} CAD/t"
                    used_symbol = symbol
                    print(f"Found working canola symbol: {symbol}")
                    break
                else:
                    print(f"Symbol {symbol} returned None price")
            except Exception as e:
                print(f"Failed with symbol {symbol}: {e}")
                traceback.print_exc()
        
        # If all symbols failed, try a fallback method
        if canola_price is None:
            try:
                print("All canola symbols failed. Attempting fallback method...")
                # Use our fallback method to get a price
                canola_price = await self.get_canola_price_fallback()
                if canola_price is not None:
                    canola_text = f"ℹ️ Canola: ${canola_price:,.2f} CAD/t (Est.)"
                    used_symbol = "FALLBACK"
                else:
                    canola_text = f"❌ Canola: Data unavailable (tried all methods)"
            except Exception as e:
                print(f"Fallback method failed: {e}")
                traceback.print_exc()
                canola_text = f"❌ Canola: Data unavailable (tried {len(CANOLA_SYMBOLS)} symbols)"
        
        # Send canola price to UI
        self.signals.canola_ready.emit(canola_text, now, canola_price)
        
        # Cache the data with more diagnostic info
        commodities_cache = {
            "wheat_text": wheat_text,
            "wheat_price": wheat_value,
            "canola_text": canola_text,
            "canola_price": canola_price,
            "bitcoin_text": bitcoin_text,
            "bitcoin_price": bitcoin_value,
            "timestamp": now,
            "fetch_time": datetime.now().isoformat(),
            "canola_symbol_used": used_symbol,
            "canola_symbols_tried": CANOLA_SYMBOLS
        }
        DataCache.save_to_cache(commodities_cache, COMMODITIES_CACHE)
    
    def closeEvent(self, event):
        """Clean up resources when the widget is closed"""
        # Stop the scheduled refresh thread
        if hasattr(self, 'scheduled_refresh'):
            self.scheduled_refresh.stop()
            self.scheduled_refresh.wait()
        # Let the event continue to parent handlers
        super().closeEvent(event)
