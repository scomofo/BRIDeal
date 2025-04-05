import sys
import os
import traceback
import asyncio
import aiohttp
import requests
import yfinance as yf
import json
import time
import re
import math
import pandas as pd # Added for yfinance processing
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QWidget, QLabel, QGridLayout, QVBoxLayout, QHBoxLayout, QLayout, # Added QLayout
    QFrame, QSizePolicy, QGroupBox, QPushButton, QComboBox,
    QSpacerItem, QGraphicsDropShadowEffect, QProgressBar
)
# Added QPoint import
from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread, QSize, QTimer, QPoint
from PyQt5.QtGui import (
    QColor, QFont, QIcon, QPalette, QPainter, QBrush, QPen, QLinearGradient
)

# Constants defined outside the class for clarity, or keep inside if preferred structure
OPENWEATHER_API_KEY = "711ac00142aa78e1807ce84a8bf1582b"
OPENWEATHER_BASE_URL = "https://api.openweathermap.org/data/2.5/weather"
ALPHAVANTAGE_API_KEY = "PHNW69I8KX24I5PT"
ALPHAVANTAGE_BASE_URL = "https://www.alphavantage.co/query"
WHEAT_SYMBOL = "ZW=F"
BITCOIN_SYMBOL = "BTC-USD"  # Added Bitcoin symbol
# Expanded list of potential canola symbols to try
CANOLA_SYMBOLS = [
    "RS=F",         # Original symbol
    "CAD=F",        # Previously tried alternative
    "ICE:RS1!",     # ICE Canola Front Month
    "RSX23.NYM",    # Specific contract month format
    "ICE:RS",       # ICE Canola general
    "CED.TO",       # Canola ETF on Toronto exchange
    "CA:RS*1",      # Another variation
    "RSK23.NYM"     # Yet another contract format
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
    "primary": "#367C2B",     # John Deere Green
    "secondary": "#FFDE00",   # John Deere Yellow
    "accent": "#2A5D24",      # Darker Green
    "background": "#F8F9FA",  # Light Gray
    "card_bg": "#FFFFFF",     # White
    "text_primary": "#333333", # Dark Gray
    "text_secondary": "#777777", # Medium Gray
    "success": "#28A745",     # Success Green
    "warning": "#FFC107",     # Warning Yellow
    "danger": "#DC3545",      # Danger Red
    "info": "#17A2B8",        # Info Blue
    "bitcoin": "#F7931A"      # Bitcoin Orange
}

# Import config if possible
try:
    # Attempt to import config if it exists
    # from config import APIConfig, AppSettings # Example import
    print("Attempting to import config module...")
    # If successful, try to use config values but fall back to hardcoded values
    try:
        from config import APIConfig #, AppSettings
        print("Successfully imported config module")
        OPENWEATHER_API_KEY = APIConfig.OPENWEATHER.get("API_KEY", OPENWEATHER_API_KEY)
        OPENWEATHER_BASE_URL = APIConfig.OPENWEATHER.get("BASE_URL", OPENWEATHER_BASE_URL)
        ALPHAVANTAGE_API_KEY = APIConfig.ALPHAVANTAGE.get("API_KEY", ALPHAVANTAGE_API_KEY)
        ALPHAVANTAGE_BASE_URL = APIConfig.ALPHAVANTAGE.get("BASE_URL", ALPHAVANTAGE_BASE_URL)
        WHEAT_SYMBOL = APIConfig.COMMODITIES.get("WHEAT_SYMBOL", WHEAT_SYMBOL)
        # Use .get() for safer dictionary access
        config_canola_symbol = APIConfig.COMMODITIES.get("CANOLA_SYMBOL")
        if config_canola_symbol:
            CANOLA_SYMBOL = config_canola_symbol
            # Add configured symbol to the start of our list if not already there
            if CANOLA_SYMBOL not in CANOLA_SYMBOLS:
                CANOLA_SYMBOLS.insert(0, CANOLA_SYMBOL)
    except (AttributeError, KeyError, NameError) as e: # Catch NameError if APIConfig itself isn't defined
        print(f"Warning: Could not read some config values: {e}. Using defaults.")
except ImportError as e:
    print(f"WARNING: Could not import config module: {e}")
    print("Using hardcoded API keys and settings instead.")

print(f"Using OpenWeather API Key: {OPENWEATHER_API_KEY[:5]}...{OPENWEATHER_API_KEY[-5:]}")
print(f"Using AlphaVantage API Key: {ALPHAVANTAGE_API_KEY[:5]}...{ALPHAVANTAGE_API_KEY[-5:]}")
print(f"Available canola symbols to try: {CANOLA_SYMBOLS}")


class WeatherAPIErrors:
    INVALID_API_KEY = 401
    NOT_FOUND = 404
    RATE_LIMIT = 429

class AlphaVantageErrors:
    INVALID_API_KEY = 401 # Note: AlphaVantage often returns 200 OK with an error message inside JSON for invalid keys
    RATE_LIMIT = 429 # Usually indicated by a message in the JSON response


class DataCache:
    @staticmethod
    def ensure_cache_dir():
        if not os.path.exists(CACHE_DIR):
            try:
                os.makedirs(CACHE_DIR)
                print(f"Created cache directory: {CACHE_DIR}")
            except OSError as e:
                print(f"Error creating cache directory {CACHE_DIR}: {e}")


    @staticmethod
    def save_to_cache(data, cache_file):
        DataCache.ensure_cache_dir()
        try:
            with open(cache_file, 'w') as f:
                json.dump(data, f, indent=4) # Add indent for readability
            print(f"Data saved to {cache_file}")
        except Exception as e:
            print(f"Error saving cache to {cache_file}: {e}")
            traceback.print_exc()


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
                return None
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON from {cache_file}: {e}")
            return None
        except Exception as e:
            print(f"Error loading cache from {cache_file}: {e}")
            traceback.print_exc()
            return None

    @staticmethod
    def is_cache_expired(cache_file, hours):
        """Check if cache file is older than specified hours"""
        if not os.path.exists(cache_file):
            return True

        try:
            file_time = os.path.getmtime(cache_file)
            file_datetime = datetime.fromtimestamp(file_time)
            now = datetime.now()

            # Calculate if file is older than the specified hours
            return (now - file_datetime) > timedelta(hours=hours)
        except Exception as e:
            print(f"Error checking cache expiration for {cache_file}: {e}")
            return True # Assume expired if there's an error


class HomeSignals(QObject):
    # Define signals for async operations
    weather_ready = pyqtSignal(list, str)
    # Modify exchange_ready to accept None as the third argument
    exchange_ready = pyqtSignal(str, str, object)  # Use object to allow None or float
    wheat_ready = pyqtSignal(str, str, object)  # Apply same change to other signals
    canola_ready = pyqtSignal(str, str, object)
    bitcoin_ready = pyqtSignal(str, str, object)
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
        self.is_running = True

    def run(self):
        # Create and start a new event loop in this thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            if self.is_running:
                self.home_module.signals.status_update.emit("Starting data fetch...")
                loop.run_until_complete(self.home_module._fetch_all_data(
                    self.fetch_weather, self.fetch_exchange, self.fetch_commodities
                ))
                # Status update moved to _on_refresh_complete for better timing
                # self.home_module.signals.status_update.emit("Data fetch complete.")
            else:
                 self.home_module.signals.status_update.emit("Data fetch cancelled.")

        except Exception as e:
            print(f"Error in DataFetcherThread run method: {e}")
            traceback.print_exc()
            try:
                 # Try emitting error signal, might fail if QObject affinity is wrong
                 self.home_module.signals.status_update.emit(f"Error during fetch: {e}")
            except Exception as sig_e:
                 print(f"Could not emit status update signal from thread: {sig_e}")
        finally:
            try:
                # Gracefully shut down all running tasks in the loop
                tasks = asyncio.all_tasks(loop)
                if tasks:
                    for task in tasks:
                        task.cancel()
                    loop.run_until_complete(asyncio.gather(*tasks, return_exceptions=True))

                # Shutdown async generators
                loop.run_until_complete(loop.shutdown_asyncgens())
            except Exception as shutdown_e:
                print(f"Error during event loop shutdown: {shutdown_e}")
            finally:
                 if loop.is_running():
                      loop.stop() # Ensure loop stops if still running
                 loop.close()
                 asyncio.set_event_loop(None) # Clean up reference
                 print("DataFetcherThread event loop closed.")

        self.is_running = False

    def stop(self):
        """Safely stop the thread"""
        print("Stopping DataFetcherThread...")
        self.is_running = False
        # Don't wait here, let the run method finish naturally or handle cancellation
        # self.wait() # Blocking wait can cause issues if loop is busy


class ScheduledRefreshThread(QThread):
    refresh_signal = pyqtSignal(bool, bool, bool)  # weather, exchange, commodities

    def __init__(self, parent=None):
        super().__init__(parent)
        self.running = True
        self.last_weather_refresh = datetime.min
        self.last_exchange_refresh = datetime.min
        self.last_commodities_refresh = datetime.min
        print("ScheduledRefreshThread initialized.")

    def run(self):
        print("ScheduledRefreshThread started.")
        while self.running:
            now = datetime.now()
            refresh_needed = False # Flag to check if any refresh is needed

            # Check if it's time to refresh each data type
            refresh_weather = (now - self.last_weather_refresh) > timedelta(hours=WEATHER_REFRESH_HOURS)
            refresh_exchange = (now - self.last_exchange_refresh) > timedelta(hours=EXCHANGE_REFRESH_HOURS)
            refresh_commodities = (now - self.last_commodities_refresh) > timedelta(hours=COMMODITIES_REFRESH_HOURS)

            # Determine which need refresh and update last refresh time *before* emitting
            tasks_to_refresh = {}
            if refresh_weather:
                tasks_to_refresh['weather'] = True
                self.last_weather_refresh = now
                refresh_needed = True
                print(f"Scheduled refresh triggered for Weather at {now.strftime('%Y-%m-%d %H:%M:%S')}")
            if refresh_exchange:
                tasks_to_refresh['exchange'] = True
                self.last_exchange_refresh = now
                refresh_needed = True
                print(f"Scheduled refresh triggered for Exchange Rate at {now.strftime('%Y-%m-%d %H:%M:%S')}")
            if refresh_commodities:
                tasks_to_refresh['commodities'] = True
                self.last_commodities_refresh = now
                refresh_needed = True
                print(f"Scheduled refresh triggered for Commodities at {now.strftime('%Y-%m-%d %H:%M:%S')}")


            # If any data needs refreshing, emit the signal
            if refresh_needed:
                self.refresh_signal.emit(
                    tasks_to_refresh.get('weather', False),
                    tasks_to_refresh.get('exchange', False),
                    tasks_to_refresh.get('commodities', False)
                )

            # Sleep for a while before checking again (e.g., 1 minute)
            # Use QThread.sleep() for better integration if needed, but time.sleep is often fine
            count = 0
            sleep_interval = 1 # Check every second
            total_sleep = 60 # Total sleep time in seconds
            while self.running and count < total_sleep:
                 # QThread.msleep(sleep_interval * 1000) # Use Qt sleep
                 time.sleep(sleep_interval) # time.sleep is often okay for simple timers
                 count += sleep_interval

        print("ScheduledRefreshThread finished.")


    def stop(self):
        print("Stopping ScheduledRefreshThread...")
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
        # Sizing policy needs to be set on the instance, not the class
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred) # Prefer height based on content

        # Add shadow effect
        shadow = QGraphicsDropShadowEffect(self) # Pass parent to effect
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 30)) # Slightly darker shadow
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)

        # Create layout
        self.main_layout = QVBoxLayout(self) # Renamed for clarity
        self.main_layout.setContentsMargins(15, 15, 15, 15)
        self.main_layout.setSpacing(10) # Add some spacing

        # Title with styled font
        self.title_label = QLabel(title)
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        self.title_label.setStyleSheet(f"color: {COLORS['accent']}; border: none; background: none;") # Ensure no background/border override

        # Content area widget and its layout
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(0, 5, 0, 0) # Reduced margin above content
        self.content_layout.setSpacing(5)

        # Add title and content widget to main layout
        self.main_layout.addWidget(self.title_label)
        self.main_layout.addWidget(self.content_widget)
        self.main_layout.addStretch(1) # Add stretch to push footer down

        # Footer label for timestamp etc.
        self.footer_label = QLabel("")
        self.footer_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter) # Align right, vertical center
        self.footer_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; border: none; background: none; padding-top: 5px;") # Add padding above footer
        self.main_layout.addWidget(self.footer_label)


    def add_content(self, widget):
        """Adds a QWidget to the card's content area."""
        if isinstance(widget, QWidget):
             # Add the provided widget to the content_layout
             self.content_layout.addWidget(widget)
        elif isinstance(widget, QLayout):
             # If a layout is passed, add it to the content_layout
             # This might be needed if the content itself is just a layout
             print("Warning: Adding QLayout directly to StyledCard content. Consider wrapping in QWidget.")
             self.content_layout.addLayout(widget)
        else:
             print(f"Error: StyledCard.add_content expects a QWidget or QLayout, got {type(widget)}")


    def add_footer_text(self, text):
         """Sets the text of the footer label."""
         self.footer_label.setText(str(text))


# Custom styled button
class StyledButton(QPushButton):
    def __init__(self, text, icon_name=None, is_primary=True, parent=None):
        super().__init__(text, parent)
        self.setCursor(Qt.PointingHandCursor) # Make it look clickable

        # Set colors based on button type
        if is_primary:
            bg_color = COLORS['primary']
            hover_color = COLORS['accent']
            text_color = "white"
        else:
            bg_color = COLORS['secondary']
            hover_color = "#E5C700"  # Darker yellow
            text_color = COLORS['accent']

        # Improved stylesheet
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {bg_color};
                color: {text_color};
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 10pt; /* Explicit font size */
                outline: none; /* Remove focus outline */
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
            QPushButton:pressed {{
                background-color: {hover_color};
                padding-top: 9px; /* Simulate press */
                padding-bottom: 7px;
            }}
            QPushButton:disabled {{ /* Style for disabled state */
                background-color: #cccccc;
                color: #666666;
            }}
        """)

        # Add icon if provided
        if icon_name:
            icon_path = f"icons/{icon_name}.png"  # Adjust path as needed
            if os.path.exists(icon_path):
                self.setIcon(QIcon(icon_path))
                self.setIconSize(QSize(18, 18))
            else:
                 print(f"Warning: Icon not found at {icon_path}")


# Weather icon widget
class WeatherIconWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.description = "clear" # Default description
        self.temp = 0.0
        self.setMinimumSize(60, 60)
        self.setMaximumSize(60, 60)

    def update_weather(self, description, temp):
        self.description = str(description).lower() if description else "unknown"
        try:
             self.temp = float(temp)
        except (TypeError, ValueError):
             self.temp = 0.0 # Default if conversion fails
             print(f"Warning: Invalid temperature value received: {temp}")
        self.update() # Trigger repaint

    # Corrected paintEvent method in WeatherIconWidget class
    def paintEvent(self, event):
        """
        Custom paint event to draw the weather icon and temperature.

        Args:
            event (QPaintEvent): Paint event.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing) # Smoother text

        # Select and draw the appropriate weather icon based on description
        # Using a dictionary mapping for cleaner selection
        draw_func_map = {
            'clear': self._draw_sun,
            'cloud': self._draw_cloud, # Covers few, scattered, broken, overcast
            'rain': self._draw_rain,
            'snow': self._draw_snow,
            'thunder': self._draw_thunder,
            'fog': self._draw_fog,
            'mist': self._draw_fog,
            'drizzle': self._draw_rain, # Drizzle looks similar to light rain
            'haze': self._draw_fog, # Represent haze like fog
        }

        # Find the first matching keyword in the description
        draw_func = self._draw_thermometer # Default icon if no keyword matches
        description_lower = self.description.lower()
        for key, func in draw_func_map.items():
             # Check if the keyword is present as a whole word or substring
             if key in description_lower:
                  draw_func = func
                  break # Use the first match found

        # Call the selected drawing function
        draw_func(painter)

        # Add temperature text below the icon
        painter.setPen(QPen(QColor(COLORS['text_primary']), 1))
        font = painter.font()
        font.setPointSize(10)
        # font.setBold(True) # Optional bold text
        painter.setFont(font)
        temp_text = f"{self.temp:.0f}°" # Show integer temp
        text_rect = painter.fontMetrics().boundingRect(temp_text)
        # Center text horizontally, position below icon space (adjust 55 as needed)
        painter.drawText(int((self.width() - text_rect.width()) / 2), 55, temp_text)

        # Ensure NO code related to self.trend or drawing arrows is present here


    def _draw_sun(self, painter):
        # Center coordinates
        cx, cy = 30, 25
        radius = 12

        # Draw sun body
        painter.setBrush(QBrush(QColor(255, 200, 0)))
        painter.setPen(QPen(QColor(220, 180, 0), 1))
        painter.drawEllipse(cx - radius, cy - radius, radius * 2, radius * 2)

        # Draw rays
        painter.setPen(QPen(QColor(255, 200, 0), 2))
        num_rays = 8
        ray_len_inner = radius * 1.1
        ray_len_outer = radius * 1.6
        for i in range(num_rays):
            angle = i * math.pi * 2 / num_rays
            x1 = cx + ray_len_inner * math.cos(angle)
            y1 = cy + ray_len_inner * math.sin(angle)
            x2 = cx + ray_len_outer * math.cos(angle)
            y2 = cy + ray_len_outer * math.sin(angle)
            painter.drawLine(int(x1), int(y1), int(x2), int(y2))

    def _draw_cloud(self, painter):
         # Base cloud color
        cloud_color = QColor(220, 220, 220) # Slightly gray cloud
        outline_color = QColor(180, 180, 180)

        painter.setBrush(QBrush(cloud_color))
        painter.setPen(QPen(outline_color, 1))

        # Draw overlapping ellipses for cloud shape
        painter.drawEllipse(10, 25, 25, 15) # Bottom left
        painter.drawEllipse(25, 25, 25, 15) # Bottom right
        painter.drawEllipse(18, 15, 25, 20) # Top middle


    def _draw_rain(self, painter):
        self._draw_cloud(painter) # Draw cloud base
        painter.setPen(QPen(QColor(30, 144, 255), 2, cap_style=Qt.RoundCap)) # Blue rain drops, round ends
        # Draw slanted lines for rain drops below the cloud
        drop_start_y = 40
        drop_end_y = 50
        painter.drawLine(20, drop_start_y, 15, drop_end_y)
        painter.drawLine(30, drop_start_y, 25, drop_end_y)
        painter.drawLine(40, drop_start_y, 35, drop_end_y)


    def _draw_snow(self, painter):
        self._draw_cloud(painter) # Draw cloud base
        painter.setPen(QPen(QColor(180, 210, 255), 1)) # Light blue snowflake color
        painter.setBrush(QBrush(QColor(240, 248, 255))) # Almost white snowflake fill
        # Draw simple circles for snowflakes
        flake_radius = 3
        flake_y = 45
        painter.drawEllipse(18, flake_y, flake_radius*2, flake_radius*2)
        painter.drawEllipse(30, flake_y + 2, flake_radius*2, flake_radius*2) # Slightly offset
        painter.drawEllipse(42, flake_y, flake_radius*2, flake_radius*2)


    def _draw_thunder(self, painter):
        self._draw_cloud(painter) # Draw cloud base (maybe darker cloud?)
        painter.setBrush(QBrush(QColor(255, 215, 0))) # Gold lightning bolt
        painter.setPen(Qt.NoPen)
        # Define points for a simple lightning bolt shape
        points = [
            QPoint(30, 35), # Start top center
            QPoint(25, 45), # Down left
            QPoint(32, 45), # Right zig
            QPoint(27, 55)  # Down left zag
        ]
        # Convert points to QPolygonF or use appropriate draw method
        # Using drawPolygon which works with list of QPoint
        painter.drawPolygon(points)


    def _draw_fog(self, painter):
        painter.setPen(QPen(QColor(190, 190, 190), 3, cap_style=Qt.RoundCap)) # Thicker gray lines for fog/mist
        # Draw horizontal lines across the middle
        line_y_start = 25
        line_spacing = 8
        num_lines = 3
        line_x_start = 10
        line_x_end = 50
        for i in range(num_lines):
            y = line_y_start + i * line_spacing
            # Add slight randomness/waviness? (optional, more complex)
            painter.drawLine(line_x_start, y, line_x_end, y)

    def _draw_thermometer(self, painter):
        # Thermometer as a fallback/unknown icon
        bulb_cx, bulb_cy = 30, 45
        bulb_radius = 6
        stem_width = 8
        stem_height = 25
        stem_x = bulb_cx - stem_width // 2
        stem_y = bulb_cy - stem_height

        # Draw stem outline
        painter.setPen(QPen(QColor(160, 160, 160), 2))
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(stem_x, stem_y, stem_width, stem_height, stem_width / 2, stem_width / 2)

        # Draw bulb
        painter.setBrush(QBrush(QColor(200, 0, 0))) # Red bulb
        painter.setPen(QPen(QColor(150, 0, 0), 1))
        painter.drawEllipse(bulb_cx - bulb_radius, bulb_cy - bulb_radius, bulb_radius * 2, bulb_radius * 2)

        # Draw mercury level (partially up the stem)
        mercury_height = stem_height * 0.6 # Example level
        painter.setBrush(QBrush(QColor(200, 0, 0)))
        painter.setPen(Qt.NoPen)
        painter.drawRect(stem_x + 1, int(bulb_cy - mercury_height), stem_width - 2, int(mercury_height))


# Price trend indicator widget
class TrendIndicator(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.trend = 0  # -1: down, 0: stable, 1: up
        self.percentage = 0.0
        # Adjusted size for better fit with text
        self.setMinimumSize(80, 25)
        self.setMaximumSize(100, 25) # Allow slightly wider for longer percentages


    def set_trend(self, trend, percentage=0.0):
        try:
            self.trend = int(trend)
            self.percentage = float(percentage) if percentage is not None else 0.0
        except (ValueError, TypeError):
             print(f"Warning: Invalid trend/percentage value: trend={trend}, percentage={percentage}")
             self.trend = 0
             self.percentage = 0.0
        self.update() # Trigger repaint

    # Corrected paintEvent method in TrendIndicator class
    def paintEvent(self, event):
        """
        Custom paint event to draw the trend indicator arrow and percentage.

        Args:
            event (QPaintEvent): Paint event.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)

        arrow_width = 10
        arrow_height = 10
        arrow_x = 5 # Starting X for arrow
        arrow_y_center = self.height() // 2

        # Choose color based on trend
        if self.trend > 0:
            color = QColor(COLORS['success'])
        elif self.trend < 0:
            color = QColor(COLORS['danger'])
        else:
            color = QColor(COLORS['text_secondary'])

        # Create the pen first
        arrow_pen = QPen(color, 2)
        # Set cap and join styles using methods
        arrow_pen.setCapStyle(Qt.RoundCap)
        arrow_pen.setJoinStyle(Qt.RoundJoin)
        # Set the configured pen on the painter
        painter.setPen(arrow_pen)

        painter.setBrush(color) # Fill arrow

        # Draw trend arrow
        arrow_points = []
        if self.trend > 0:
            # Up arrow polygon points
            arrow_points = [
                QPoint(arrow_x, arrow_y_center + arrow_height // 2),
                QPoint(arrow_x + arrow_width, arrow_y_center + arrow_height // 2),
                QPoint(arrow_x + arrow_width // 2, arrow_y_center - arrow_height // 2)
            ]
        elif self.trend < 0:
            # Down arrow polygon points
            arrow_points = [
                QPoint(arrow_x, arrow_y_center - arrow_height // 2),
                QPoint(arrow_x + arrow_width, arrow_y_center - arrow_height // 2),
                QPoint(arrow_x + arrow_width // 2, arrow_y_center + arrow_height // 2)
            ]
        else:
            # Draw Horizontal line for stable
            # Create and configure a separate pen for the line
            stable_pen = QPen(color, 2)
            stable_pen.setCapStyle(Qt.RoundCap)
            painter.setPen(stable_pen) # Use the line pen
            painter.setBrush(Qt.NoBrush) # Don't fill the line
            line_y = self.height() // 2
            painter.drawLine(arrow_x, line_y, arrow_x + arrow_width, line_y)


        # Draw polygon if points exist (for up/down arrows)
        if arrow_points:
             # Using drawPolygon which works with list of QPoint
             painter.drawPolygon(arrow_points)


        # Draw percentage text next to the arrow
        if self.percentage is not None and abs(self.percentage) > 0.01: # Only show if significant
            # Use a standard pen for text
            text_pen = QPen(color, 1)
            painter.setPen(text_pen)
            font = painter.font()
            font.setPointSize(9) # Slightly smaller font
            # font.setBold(True) # Optional bold
            painter.setFont(font)

            sign = "+" if self.trend > 0 else ("-" if self.trend < 0 else "") # Use sign for clarity
            # Format with one decimal place, include sign
            text = f"{sign}{abs(self.percentage):.1f}%"

            text_x = arrow_x + arrow_width + 8 # Position text after arrow + spacing
            text_y = arrow_y_center + painter.fontMetrics().ascent() // 2 - 1 # Center vertically
            painter.drawText(text_x, text_y, text)
        elif self.trend == 0:
             # Optionally show "0.0%" or nothing for stable
             pass # Currently shows nothing for stable trend


# Circular progress gauge for visual representation
class CircularProgressGauge(QWidget):
    def __init__(self, parent=None):
        """
        Initialize the Circular Progress Gauge widget

        Args:
            parent (QWidget, optional): Parent widget. Defaults to None.
        """
        super().__init__(parent)
        self._value = 0.0
        self._maximum = 100.0
        self._minimum = 0.0 # Add minimum value support
        self._color = QColor(COLORS['primary'])
        self._text_format = "{:.1f}" # Default format for float value
        self._text_suffix = "" # Optional suffix like '%' or '°'

        # Set fixed size for the gauge
        gauge_size = 80 # Smaller gauge size
        self.setMinimumSize(gauge_size, gauge_size)
        self.setMaximumSize(gauge_size, gauge_size)


    def set_value(self, value):
        """
        Set the current value for the gauge.

        Args:
            value (float or int): Current value to display.
        """
        try:
            new_value = float(value)
            # Clamp value between min and max
            self._value = max(self._minimum, min(new_value, self._maximum))
            self.update() # Trigger a repaint
        except (TypeError, ValueError) as e:
            print(f"Error setting gauge value: {e}. Value: {value}")

    def set_maximum(self, maximum):
         """Set the maximum value."""
         try:
              self._maximum = float(maximum)
              self._value = max(self._minimum, min(self._value, self._maximum)) # Re-clamp current value
              self.update()
         except (TypeError, ValueError) as e:
              print(f"Error setting gauge maximum: {e}. Value: {maximum}")

    def set_minimum(self, minimum):
         """Set the minimum value."""
         try:
              self._minimum = float(minimum)
              self._value = max(self._minimum, min(self._value, self._maximum)) # Re-clamp current value
              self.update()
         except (TypeError, ValueError) as e:
              print(f"Error setting gauge minimum: {e}. Value: {minimum}")

    def set_range(self, minimum, maximum):
         """Set both minimum and maximum values."""
         try:
              self._minimum = float(minimum)
              self._maximum = float(maximum)
              # Ensure max is greater than min
              if self._maximum <= self._minimum:
                   print(f"Warning: Maximum ({self._maximum}) must be greater than minimum ({self._minimum}). Adjusting maximum.")
                   self._maximum = self._minimum + 1 # Set a default difference
              self._value = max(self._minimum, min(self._value, self._maximum)) # Re-clamp current value
              self.update()
         except (TypeError, ValueError) as e:
              print(f"Error setting gauge range: {e}. Min: {minimum}, Max: {maximum}")


    def set_color(self, color):
        """
        Set the color of the progress indicator.

        Args:
            color (str or QColor): Color name (from COLORS dict key) or QColor object.
        """
        # *** FIX: Check if color is a valid key in COLORS ***
        if isinstance(color, str) and color in COLORS:
            self._color = QColor(COLORS[color])
        elif isinstance(color, QColor):
            self._color = color
        else:
             print(f"Warning: Invalid color specified for gauge: {color}. Using default.")
             self._color = QColor(COLORS['primary'])
        self.update()

    def set_text_format(self, format_string, suffix=""):
         """Set the format string for displaying the value (e.g., "{:.0f}") and an optional suffix."""
         self._text_format = format_string
         self._text_suffix = suffix
         self.update()

    # Corrected paintEvent method in CircularProgressGauge class
    def paintEvent(self, event):
        """
        Custom paint event to draw the circular progress gauge.

        Args:
            event (QPaintEvent): Paint event.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)

        rect = self.rect().adjusted(5, 5, -5, -5) # Adjust drawing area slightly inward
        pen_width = 8 # Width of the progress/background arcs
        start_angle = 90 * 16  # Start at the top (12 o'clock)

        # Draw background circle (full arc)
        # *** FIX START ***
        bg_pen = QPen(QColor(230, 230, 230), pen_width) # Create pen
        bg_pen.setCapStyle(Qt.FlatCap)                 # Set cap style
        painter.setPen(bg_pen)                         # Apply pen
        # *** FIX END ***
        painter.drawArc(rect, 0, 360 * 16) # Draw full background arc

        # Draw progress arc
        # Calculate span angle based on value relative to min/max range
        value_range = self._maximum - self._minimum
        if value_range > 0:
            # Calculate proportion of the range the value represents
            proportion = (self._value - self._minimum) / value_range
            # Ensure proportion is within 0.0 to 1.0
            proportion = max(0.0, min(proportion, 1.0))
            span_angle = int(-360 * proportion * 16) # Negative for clockwise
        else:
            span_angle = 0 # No range, no progress shown


        if span_angle != 0: # Only draw if there's progress
            # *** FIX START ***
            progress_pen = QPen(self._color, pen_width) # Create pen
            progress_pen.setCapStyle(Qt.RoundCap)       # Set cap style
            painter.setPen(progress_pen)                # Apply pen
            # *** FIX END ***
            painter.drawArc(rect, start_angle, span_angle)

        # Draw value text
        painter.setPen(QPen(QColor(COLORS['text_primary']), 1))
        font = painter.font()
        font.setPointSize(12) # Slightly smaller font for smaller gauge
        font.setBold(True)
        painter.setFont(font)

        # Format text using the specified format string and add suffix
        try:
            # Handle format strings that might contain placeholders like '{:.1f}k'
            if '{' in self._text_format and '}' in self._text_format:
                 text = self._text_format.format(self._value)
            else:
                 # Assume it's a simple format specifier like '.1f' or '.0f'
                 text = f"{self._value:{self._text_format}}{self._text_suffix}"
        except (ValueError, TypeError) as format_e:
             text = "Err" # Fallback text if format fails
             print(f"Warning: Gauge text format '{self._text_format}' failed for value '{self._value}': {format_e}")


        # Center the text
        text_rect = painter.fontMetrics().boundingRect(text)
        painter.drawText(
            int((self.width() - text_rect.width()) / 2),
            int((self.height() - text_rect.height()) / 2 + painter.fontMetrics().ascent()),
            text
        )

# ==============================================================
# Main Home Module Widget
# ==============================================================
class HomeModule(QWidget):
    def __init__(self, main_window=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setObjectName("HomeModule")

        # Set background color for the module itself
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(COLORS['background']))
        self.setPalette(palette)

        # Create signals for thread-safe UI updates
        self.signals = HomeSignals()

        # --- UI Element Initialization ---
        self._init_ui_elements()

        # Setup refresh animation timer
        self.refresh_animation_timer = QTimer(self)
        self.refresh_animation_timer.timeout.connect(self._update_refresh_animation)
        self.refresh_animation_value = 0
        self._is_refreshing = False # Use a private variable convention

        # --- Layout Setup ---
        self._setup_ui_layout() # This now stores card references like self.weather_card

        # --- Signal Connections ---
        self._connect_signals()

        # --- Data Fetching Threads ---
        self.fetcher_thread = None # Initialize thread variable
        self.scheduled_refresh = ScheduledRefreshThread(self)
        self.scheduled_refresh.refresh_signal.connect(self._scheduled_refresh)

        # --- Initial Actions ---
        print("HomeModule Initializing...")
        self._load_cached_data() # Load cache first
        self.scheduled_refresh.start() # Start the timer thread
        # Perform an initial refresh for any expired/missing data AFTER UI is set up
        QTimer.singleShot(500, lambda: self._refresh_data(check_cache_first=False)) # Delay initial fetch slightly
        print("HomeModule Initialized.")


    def _init_ui_elements(self):
        """Initialize all UI widgets."""
        print("Initializing UI elements...")
        self.status_label = QLabel("Initializing dashboard...")
        self.status_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-style: italic;")
        self.status_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        # Weather widgets (using a dictionary for easier access)
        self.weather_widgets = {}
        self.weather_cities = ["Camrose", "Wainwright", "Killam", "Provost"]
        for city in self.weather_cities:
            city_label = QLabel(f"{city}: Loading...")
            city_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 10pt;") # Slightly smaller
            city_label.setWordWrap(True)
            city_icon = WeatherIconWidget()
            city_layout = QHBoxLayout()
            city_layout.addWidget(city_icon)
            city_layout.addWidget(city_label, 1) # Label takes expanding space
            city_layout.setSpacing(5) # Spacing between icon and label
            self.weather_widgets[city] = {
                'icon': city_icon,
                'label': city_label,
                'layout': city_layout
            }
        self.weather_timestamp = "Last updated: --:--" # Store as string, will be set on card footer

        # Exchange rate widgets
        self.exchange_label = QLabel("USD-CAD: Loading...")
        self.exchange_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
        self.exchange_timestamp = "Last updated: --:--"
        self.exchange_gauge = CircularProgressGauge()
        self.exchange_gauge.set_color('info') # Use key from COLORS dict
        self.exchange_gauge.set_range(1.20, 1.50) # Example range for USD/CAD
        self.exchange_gauge.set_text_format("{:.3f}") # Show 3 decimal places
        self.exchange_trend = TrendIndicator()

        # Wheat price widgets
        self.wheat_label = QLabel("Wheat: Loading...")
        self.wheat_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
        self.wheat_timestamp = "Last updated: --:--"
        self.wheat_gauge = CircularProgressGauge()
        self.wheat_gauge.set_color('success') # Use key from COLORS dict
        self.wheat_gauge.set_range(3.00, 12.00) # Range in dollars/bu (assuming conversion later)
        self.wheat_gauge.set_text_format(".2f") # Show dollars.cent format
        self.wheat_trend = TrendIndicator()

        # Canola price widgets
        self.canola_label = QLabel("Canola: Loading...")
        self.canola_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
        self.canola_timestamp = "Last updated: --:--"
        self.canola_gauge = CircularProgressGauge()
        self.canola_gauge.set_color('secondary') # Use key from COLORS dict
        self.canola_gauge.set_range(500, 1200) # Example range CAD/t
        self.canola_gauge.set_text_format(".0f") # Show integer value
        self.canola_trend = TrendIndicator()

        # Bitcoin price widgets
        self.bitcoin_label = QLabel("Bitcoin: Loading...")
        self.bitcoin_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 11pt;")
        self.bitcoin_timestamp = "Last updated: --:--"
        self.bitcoin_gauge = CircularProgressGauge()
        self.bitcoin_gauge.set_color('bitcoin') # Use key from COLORS dict
        self.bitcoin_gauge.set_range(10, 100) # Range in thousands of USD (e.g., 10k to 100k)
        self.bitcoin_gauge.set_text_format("{:.1f}k") # Show value in thousands
        self.bitcoin_trend = TrendIndicator()

        # Refresh options dropdown
        self.refresh_combo = QComboBox()
        self.refresh_combo.addItems([
            "Refresh All",
            "Refresh Weather Only",
            "Refresh Exchange Rate Only",
            "Refresh Commodities Only"
        ])
        self.refresh_combo.setStyleSheet(f"""
            QComboBox {{
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
                background-color: white;
                min-width: 160px; /* Slightly wider */
                font-size: 10pt;
                color: {COLORS['text_primary']};
            }}
            QComboBox::drop-down {{ /* Style the dropdown arrow area */
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: #ddd;
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
            }}
             QComboBox::down-arrow {{ /* The actual arrow */
                 /* image: url(icons/down_arrow.png); /* Optional: Custom arrow icon */
                 width: 10px;
                 height: 10px;
             }}
            QComboBox QAbstractItemView {{ /* Style the dropdown list */
                 border: 1px solid #ddd;
                 background-color: white;
                 selection-background-color: {COLORS['primary']};
                 color: {COLORS['text_primary']}; /* Ensure text is visible */
                 padding: 2px;
            }}
        """)
        # Button for triggering the refresh based on combo box
        self.refresh_button = QPushButton("↻ Refresh")
        self.refresh_button.setStyleSheet(f"""
            QPushButton {{
                background: none;
                border: none;
                color: {COLORS['primary']};
                font-size: 10pt;
                font-weight: bold; /* Make it slightly more prominent */
                text-decoration: none; /* Remove underline, looks cleaner */
                padding: 5px; /* Match combo box padding */
                margin-left: 5px; /* Space between combo and button */
            }}
            QPushButton:hover {{
                color: {COLORS['accent']};
                text-decoration: underline; /* Underline on hover */
            }}
             QPushButton:pressed {{
                 color: {COLORS['accent']};
             }}
             QPushButton:disabled {{ /* Style for disabled state */
                 color: #999999;
             }}
        """)
        self.refresh_button.setCursor(Qt.PointingHandCursor)
        self.refresh_button.clicked.connect(self._handle_refresh_button)

        # Button for forcing canola refresh
        self.canola_refresh_button = StyledButton("Force Canola Refresh", is_primary=False)
        self.canola_refresh_button.clicked.connect(self.force_refresh_commodities)
        self.canola_refresh_button.setToolTip("Try this if Canola data seems stuck or incorrect.")


    def _setup_ui_layout(self):
        """Setup the main layout and arrange widgets."""
        print("Setting up UI layout...")
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # --- Header ---
        header_layout = QHBoxLayout()
        dashboard_title = QLabel("AMS Dashboard")
        dashboard_title.setStyleSheet(f"""
            font-size: 18pt;
            font-weight: bold;
            color: {COLORS['accent']};
        """)
        header_layout.addWidget(dashboard_title)
        header_layout.addStretch()

        # Refresh controls in header
        refresh_controls_layout = QHBoxLayout()
        refresh_controls_layout.setSpacing(5)
        refresh_controls_layout.addWidget(self.refresh_combo)
        refresh_controls_layout.addWidget(self.refresh_button)
        header_layout.addLayout(refresh_controls_layout)

        # Add status label to header
        header_layout.addWidget(self.status_label)
        header_layout.setStretchFactor(dashboard_title, 0) # Title takes minimum space
        header_layout.setStretchFactor(refresh_controls_layout, 0) # Controls take minimum space
        header_layout.setStretchFactor(self.status_label, 1) # Status takes remaining space

        main_layout.addLayout(header_layout)

        # --- Main Content Grid ---
        grid_layout = QGridLayout()
        grid_layout.setSpacing(20) # Increased spacing between cards

        # --- Weather Card (Row 0, Col 0, Span 1x2) ---
        # Store reference to the card
        self.weather_card = StyledCard("Weather Conditions")
        weather_content_layout = QVBoxLayout()
        weather_content_layout.setSpacing(8) # Spacing between cities
        # Remove margins from the inner layout, card already has margins
        weather_content_layout.setContentsMargins(0,0,0,0)

        for city in self.weather_cities:
            # addLayout is correct here, adding the QHBoxLayout for each city
            weather_content_layout.addLayout(self.weather_widgets[city]['layout'])

        # Create a container QWidget for the weather layout
        weather_container_widget = QWidget()
        weather_container_widget.setLayout(weather_content_layout) # Set the layout on the widget

        # Add the container WIDGET (not the layout) to the card's content area
        self.weather_card.add_content(weather_container_widget)

        self.weather_card.add_footer_text(self.weather_timestamp) # Use method to set footer
        grid_layout.addWidget(self.weather_card, 0, 0, 1, 2) # Row 0, Col 0, Span 1 row, 2 cols

        # --- Exchange Rate Card (Row 1, Col 0) ---
        # Store reference
        self.exchange_card = StyledCard("Currency Exchange (USD/CAD)")
        exchange_content_layout = QHBoxLayout()
        exchange_text_layout = QVBoxLayout() # Layout for label and trend
        exchange_text_layout.addWidget(self.exchange_label)
        exchange_text_layout.addSpacing(5)
        exchange_text_layout.addWidget(self.exchange_trend)
        exchange_text_layout.addStretch()
        exchange_content_layout.addLayout(exchange_text_layout, 1) # Text takes more space
        exchange_content_layout.addWidget(self.exchange_gauge) # Gauge on the right

        # Container widget pattern (already correctly applied here)
        exchange_widget = QWidget()
        exchange_widget.setLayout(exchange_content_layout)
        self.exchange_card.add_content(exchange_widget)

        self.exchange_card.add_footer_text(self.exchange_timestamp)
        grid_layout.addWidget(self.exchange_card, 1, 0)

        # --- Commodities Container (Row 1, Col 1, Span 2 rows) ---
        commodities_grid = QGridLayout()
        commodities_grid.setSpacing(15)

        # --- Wheat Card (Commodities Grid: 0, 0) ---
        # Store reference
        self.wheat_card = StyledCard("Wheat Price (ZW=F)")
        wheat_content_layout = QHBoxLayout()
        wheat_text_layout = QVBoxLayout()
        wheat_text_layout.addWidget(self.wheat_label)
        wheat_text_layout.addSpacing(5)
        wheat_text_layout.addWidget(self.wheat_trend)
        wheat_text_layout.addStretch()
        wheat_content_layout.addLayout(wheat_text_layout, 1)
        wheat_content_layout.addWidget(self.wheat_gauge)

        # Container widget pattern
        wheat_widget = QWidget()
        wheat_widget.setLayout(wheat_content_layout)
        self.wheat_card.add_content(wheat_widget)

        self.wheat_card.add_footer_text(self.wheat_timestamp)
        commodities_grid.addWidget(self.wheat_card, 0, 0)

        # --- Canola Card (Commodities Grid: 0, 1) ---
        # Store reference
        self.canola_card = StyledCard("Canola Price")
        canola_content_layout = QHBoxLayout()
        canola_text_layout = QVBoxLayout()
        canola_text_layout.addWidget(self.canola_label)
        canola_text_layout.addSpacing(5)
        canola_text_layout.addWidget(self.canola_trend)
        canola_text_layout.addStretch()
        canola_content_layout.addLayout(canola_text_layout, 1)
        canola_content_layout.addWidget(self.canola_gauge)

        # Container widget pattern
        canola_widget = QWidget()
        canola_widget.setLayout(canola_content_layout)
        self.canola_card.add_content(canola_widget)

        self.canola_card.add_footer_text(self.canola_timestamp)
        commodities_grid.addWidget(self.canola_card, 0, 1)

        # --- Bitcoin Card (Commodities Grid: 1, 0) ---
        # Store reference
        self.bitcoin_card = StyledCard("Bitcoin Price (BTC-USD)")
        bitcoin_content_layout = QHBoxLayout()
        bitcoin_text_layout = QVBoxLayout()
        bitcoin_text_layout.addWidget(self.bitcoin_label)
        bitcoin_text_layout.addSpacing(5)
        bitcoin_text_layout.addWidget(self.bitcoin_trend)
        bitcoin_text_layout.addStretch()
        bitcoin_content_layout.addLayout(bitcoin_text_layout, 1)
        bitcoin_content_layout.addWidget(self.bitcoin_gauge)

        # Container widget pattern
        bitcoin_widget = QWidget()
        bitcoin_widget.setLayout(bitcoin_content_layout)
        self.bitcoin_card.add_content(bitcoin_widget)

        self.bitcoin_card.add_footer_text(self.bitcoin_timestamp)
        commodities_grid.addWidget(self.bitcoin_card, 1, 0)

        # --- Canola Refresh Card (Commodities Grid: 1, 1) ---
        canola_refresh_widget = QFrame()
        canola_refresh_widget.setStyleSheet(f"background-color: {COLORS['card_bg']}; border-radius: 8px; border: 1px solid #e0e0e0;")
        canola_refresh_layout = QVBoxLayout(canola_refresh_widget)
        canola_refresh_layout.setContentsMargins(10, 10, 10, 10)
        canola_refresh_layout.setSpacing(8)

        canola_refresh_info = QLabel("Canola data issues? Try forcing a refresh.")
        canola_refresh_info.setWordWrap(True)
        canola_refresh_info.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; border: none; background: none;")
        canola_refresh_layout.addWidget(canola_refresh_info)
        canola_refresh_layout.addWidget(self.canola_refresh_button)
        canola_refresh_layout.addStretch()

        canola_refresh_shadow = QGraphicsDropShadowEffect(canola_refresh_widget)
        canola_refresh_shadow.setBlurRadius(10)
        canola_refresh_shadow.setColor(QColor(0, 0, 0, 25))
        canola_refresh_shadow.setOffset(0, 1)
        canola_refresh_widget.setGraphicsEffect(canola_refresh_shadow)

        commodities_grid.addWidget(canola_refresh_widget, 1, 1)

        # Add the commodities grid layout to the main grid layout
        grid_layout.addLayout(commodities_grid, 1, 1, 2, 1) # Row 1, Col 1, Span 2 rows, 1 col

        # --- Add Grid to Main Layout ---
        main_layout.addLayout(grid_layout)
        main_layout.addStretch() # Add stretch at the bottom


    def _connect_signals(self):
        """Connect signals to slots."""
        print("Connecting signals...")
        self.signals.weather_ready.connect(self._update_weather_ui)
        self.signals.exchange_ready.connect(self._update_exchange_ui)
        self.signals.wheat_ready.connect(self._update_wheat_ui)
        self.signals.canola_ready.connect(self._update_canola_ui)
        self.signals.bitcoin_ready.connect(self._update_bitcoin_ui)
        self.signals.status_update.connect(self._update_status_label)
        self.signals.refresh_complete.connect(self._on_refresh_complete)

        # Connect combo box change to potentially update button text (optional)
        # self.refresh_combo.currentIndexChanged.connect(self._update_refresh_button_text)


    # --- Refresh Handling ---

    def _handle_refresh_button(self):
        """Handle refresh button click based on selected option."""
        if self._is_refreshing:
            print("Refresh already in progress.")
            return # Don't start another refresh if one is running

        option = self.refresh_combo.currentText()
        print(f"Manual refresh requested: {option}")

        fetch_flags = {'weather': False, 'exchange': False, 'commodities': False}

        if option == "Refresh All":
            fetch_flags = {'weather': True, 'exchange': True, 'commodities': True}
        elif option == "Refresh Weather Only":
            fetch_flags['weather'] = True
        elif option == "Refresh Exchange Rate Only":
            fetch_flags['exchange'] = True
        elif option == "Refresh Commodities Only":
            fetch_flags['commodities'] = True

        # Start animation before calling _refresh_data
        self._start_refresh_animation()
        # Pass flags to refresh data, allow cache check by default for manual refresh
        self._refresh_data(
             fetch_weather=fetch_flags['weather'],
             fetch_exchange=fetch_flags['exchange'],
             fetch_commodities=fetch_flags['commodities'],
             check_cache_first=True # Manual refresh usually respects cache intervals
        )

    def _scheduled_refresh(self, weather, exchange, commodities):
        """Handle scheduled refreshes from the timer thread."""
        if self._is_refreshing:
            print("Scheduled refresh skipped: another refresh is in progress.")
            return

        print(f"Scheduled refresh triggered: Weather={weather}, Exchange={exchange}, Commodities={commodities}")
        # Start animation before calling _refresh_data
        self._start_refresh_animation()
        # Scheduled refresh *forces* the fetch if the time has passed, so don't check cache again here
        self._refresh_data(weather, exchange, commodities, check_cache_first=False)


    def force_refresh_commodities(self):
        """Force refresh of commodity data by bypassing cache check."""
        if self._is_refreshing:
            print("Cannot force commodity refresh: another refresh is in progress.")
            return

        print("Forcing refresh of commodity data (ignoring cache)...")
        self._start_refresh_animation()
        self.signals.status_update.emit("Forcing commodity data refresh...")

        # Delete the commodities cache file to ensure fresh data attempt
        if os.path.exists(COMMODITIES_CACHE):
            try:
                os.remove(COMMODITIES_CACHE)
                print(f"Deleted commodities cache file: {COMMODITIES_CACHE}")
            except Exception as e:
                print(f"Warning: Error deleting cache file {COMMODITIES_CACHE}: {e}")

        # Start a new thread to fetch ONLY commodities data, ignoring cache check logic
        self._start_data_fetcher_thread(fetch_weather=False, fetch_exchange=False, fetch_commodities=True)


    def _refresh_data(self, fetch_weather=True, fetch_exchange=True, fetch_commodities=True, check_cache_first=True):
        """
        Initiates data refresh, optionally checking cache validity first.
        Starts the DataFetcherThread if any data needs fetching.
        """
        if self._is_refreshing and (fetch_weather or fetch_exchange or fetch_commodities):
             # Avoid starting a new fetch if already refreshing, unless it's a no-op call
             print("Refresh requested but already in progress. Ignoring.")
             return

        # Determine what actually needs fetching based on cache expiry (if requested)
        should_fetch_weather = fetch_weather
        should_fetch_exchange = fetch_exchange
        should_fetch_commodities = fetch_commodities

        if check_cache_first:
            print("Checking cache validity before fetching...")
            if fetch_weather and DataCache.is_cache_expired(WEATHER_CACHE, WEATHER_REFRESH_HOURS):
                print("Weather cache expired or missing.")
            elif fetch_weather:
                print("Weather cache is still valid, skipping fetch.")
                should_fetch_weather = False

            if fetch_exchange and DataCache.is_cache_expired(EXCHANGE_CACHE, EXCHANGE_REFRESH_HOURS):
                print("Exchange rate cache expired or missing.")
            elif fetch_exchange:
                print("Exchange rate cache is still valid, skipping fetch.")
                should_fetch_exchange = False

            if fetch_commodities and DataCache.is_cache_expired(COMMODITIES_CACHE, COMMODITIES_REFRESH_HOURS):
                print("Commodities cache expired or missing.")
            elif fetch_commodities:
                print("Commodities cache is still valid, skipping fetch.")
                should_fetch_commodities = False

        # Only create and run thread if there's something to fetch
        if should_fetch_weather or should_fetch_exchange or should_fetch_commodities:
            # Animation is started by the calling function (_handle_refresh_button, _scheduled_refresh)
            # self._start_refresh_animation() # Don't start animation here
            self.signals.status_update.emit("Starting data fetch...")
            self._start_data_fetcher_thread(should_fetch_weather, should_fetch_exchange, should_fetch_commodities)
        else:
            print("All required data is up-to-date in cache. No fetch needed.")
            # Ensure animation stops if it was somehow started, and update status
            self._stop_refresh_animation()
            self.signals.status_update.emit("Using cached data (refresh not needed)")


    def _start_data_fetcher_thread(self, weather, exchange, commodities):
         """Stops existing thread (if any) and starts a new DataFetcherThread."""
         if self.fetcher_thread and self.fetcher_thread.isRunning():
              print("Warning: Previous fetcher thread still running. Attempting to stop...")
              self.fetcher_thread.stop()
              # Give it a moment to potentially finish/stop before starting new one
              # self.fetcher_thread.wait(1000) # Optional wait with timeout

         print(f"Starting DataFetcherThread: W={weather}, E={exchange}, C={commodities}")
         self.fetcher_thread = DataFetcherThread(
              self,
              fetch_weather=weather,
              fetch_exchange=exchange,
              fetch_commodities=commodities
         )
         # Connect thread finished signal to clean up
         self.fetcher_thread.finished.connect(self._on_fetcher_thread_finished)
         self.fetcher_thread.start()


    def _on_fetcher_thread_finished(self):
         """Slot called when the DataFetcherThread finishes."""
         print("DataFetcherThread finished signal received.")
         # Refresh might not be fully complete yet (signals might still be processing)
         # _on_refresh_complete handles stopping animation and final status update
         # Optionally check exit status or perform other cleanup
         # self.fetcher_thread = None # Clear reference? Or keep for potential inspection


    def _start_refresh_animation(self):
         """Starts the refreshing animation and disables refresh button."""
         if not self._is_refreshing:
              print("Starting refresh animation.")
              self._is_refreshing = True
              self.refresh_animation_value = 0
              self.refresh_animation_timer.start(150) # Slower animation interval
              self._update_refresh_animation() # Initial update
              self.refresh_button.setEnabled(False) # Disable button during refresh
              self.refresh_combo.setEnabled(False)
              self.canola_refresh_button.setEnabled(False)


    def _stop_refresh_animation(self):
         """Stops the refreshing animation and enables refresh button."""
         if self._is_refreshing:
              print("Stopping refresh animation.")
              self._is_refreshing = False
              self.refresh_animation_timer.stop()
              # Status label update is handled by _on_refresh_complete or specific error signals
              # self._update_status_label("Dashboard ready" if not self.fetcher_thread or not self.fetcher_thread.isRunning() else "Finishing refresh...") # Avoid setting premature status
              self.refresh_button.setEnabled(True) # Re-enable button
              self.refresh_combo.setEnabled(True)
              self.canola_refresh_button.setEnabled(True)


    def _update_refresh_animation(self):
        """Update the animation dots in the status label during refresh."""
        if self._is_refreshing:
            self.refresh_animation_value = (self.refresh_animation_value + 1) % 4 # Cycle 0, 1, 2, 3
            dots = "." * self.refresh_animation_value
            # Update status without overwriting potential error messages
            current_text = self.status_label.text()
            if "Refreshing data" in current_text or self._is_refreshing: # Only update if in refreshing state
                 self._update_status_label(f"Refreshing data{dots}")


    def _on_refresh_complete(self):
        """Handle completion of the _fetch_all_data async tasks."""
        print("Signal refresh_complete received.")
        self._stop_refresh_animation()
        # Final status update - check if any errors occurred during the fetch
        # (This requires more complex state tracking or checking UI for error messages)
        # For now, set a generic 'refreshed' message if no obvious errors are visible.
        # A better approach would be to track success/failure per task.
        self.signals.status_update.emit("Dashboard data refreshed.")


    # --- UI Update Slots ---

    def _update_status_label(self, status_text):
         """Updates the status label text."""
         # Avoid overwriting critical error messages if needed (add logic here if required)
         self.status_label.setText(status_text)

    def _update_weather_ui(self, results, timestamp):
        """Update the weather UI with the fetched data."""
        print(f"Updating weather UI with {len(results)} results. Timestamp: {timestamp}")
        found_cities = set()

        for result_str in results:
            # Regex to parse success/error messages more robustly
            # Success: ✅ CityName: temp°C (Feels like feels°) Description Emoji Icon
            # Error:   ❌ CityName: Error Message
            success_match = re.match(r"✅\s*([^:]+):\s*(-?[\d.]+)\s*°C.*,\s*([a-zA-Z\s]+)", result_str)
            error_match = re.match(r"❌\s*([^:]+):\s*(.*)", result_str)

            city_name = None
            widget_data = None

            if success_match:
                city_name = success_match.group(1).strip()
                if city_name in self.weather_widgets:
                    widget_data = self.weather_widgets[city_name]
                    try:
                        temp_value = float(success_match.group(2))
                        description = success_match.group(3).strip()
                        widget_data['icon'].update_weather(description, temp_value)
                        # Use the full result string for the label for more detail
                        styled_text = f"<span style='font-weight:bold; color:{COLORS['accent']}'>{city_name}:</span> {result_str.split(':', 1)[1].strip()}"
                        widget_data['label'].setText(styled_text)
                        widget_data['label'].setToolTip(result_str) # Show full details on hover
                        found_cities.add(city_name)
                    except Exception as e:
                         print(f"Error parsing weather success message for {city_name}: {e} | String: {result_str}")
                         widget_data['label'].setText(f"{city_name}: Parse Error")
            elif error_match:
                city_name = error_match.group(1).strip()
                if city_name in self.weather_widgets:
                    widget_data = self.weather_widgets[city_name]
                    error_msg = error_match.group(2).strip()
                    styled_text = f"<span style='font-weight:bold; color:{COLORS['danger']}'>{city_name}:</span> {error_msg}"
                    widget_data['label'].setText(styled_text)
                    widget_data['label'].setToolTip(error_msg) # Show error on hover
                    widget_data['icon'].update_weather("unknown", 0) # Reset icon on error
                    found_cities.add(city_name)
            else:
                 print(f"Warning: Could not parse weather result string: {result_str}")


        # Handle cities that might not have been in the results
        for city in self.weather_cities:
             if city not in found_cities and city in self.weather_widgets:
                  print(f"Warning: No weather data received for {city}")
                  self.weather_widgets[city]['label'].setText(f"{city}: No data")
                  self.weather_widgets[city]['icon'].update_weather("unknown", 0)

        # Update the timestamp on the card using the stored reference
        if hasattr(self, 'weather_card') and self.weather_card:
             self.weather_card.add_footer_text(f"Last updated: {timestamp}")
        else:
             print("Could not find self.weather_card to update timestamp")


    def _update_exchange_ui(self, rate_text, timestamp, rate_value=None):
        """Update the Exchange Rate UI"""
        print(f"Updating Exchange UI: Value={rate_value}, Text='{rate_text}', Time={timestamp}")
        # Use the generic commodity update logic for consistency
        self._update_commodity_ui(
            label_widget=self.exchange_label,
            gauge_widget=self.exchange_gauge,
            trend_widget=self.exchange_trend,
            card_ref=getattr(self, 'exchange_card', None),
            prefix="USD/CAD", # Use prefix for formatting
            value_text=rate_text, # Pass original text for error handling
            timestamp=timestamp,
            price_value=rate_value,
            color_key='info',
            gauge_range=(1.20, 1.50),
            text_format="{:.4f}", # Exchange rate needs more precision
            trend_baseline=1.35,
            trend_sensitivity=1.0, # More sensitivity for smaller % changes
            currency_symbol="", # No currency symbol needed for exchange rate
            unit=""
        )


    def _update_commodity_ui(self, label_widget, gauge_widget, trend_widget, card_ref,
                              prefix, value_text, timestamp, price_value=None,
                              color_key='success', gauge_range=(0, 1), text_format="{:.2f}",
                              trend_baseline=None, trend_sensitivity=1.0, currency_symbol="$", unit=""):
         """Generic helper to update commodity UI elements."""
         print(f"Updating {prefix} UI: Value={price_value}, Text='{value_text}', Time={timestamp}")
         styled_text = value_text # Default
         trend = 0
         trend_percent = 0.0
         tooltip = value_text # Default tooltip

         # Determine color based on status prefix in value_text
         current_color_key = color_key # Default color
         if value_text.startswith("✅"):
              current_color_key = color_key # Use default success/assigned color
         elif value_text.startswith("❌"):
              current_color_key = 'danger'
         elif value_text.startswith("ℹ️"):
              current_color_key = 'info'


         if price_value is not None and isinstance(price_value, (float, int)):
             try:
                  # Format price using currency symbol and format string
                  # Handle cases where text_format already includes formatting (like 'k' for bitcoin)
                  if '{' in text_format and '}' in text_format:
                       formatted_price = text_format.format(price_value)
                  else:
                       # Apply basic currency formatting if text_format is simple
                       formatted_price = f"{currency_symbol}{price_value:{text_format.replace('$', '')}}" # Remove potential existing $

                  full_unit = f" {unit}" if unit else "" # Add unit if provided

                  # Construct styled text, using determined color
                  styled_text = f"<span style='font-size:14pt; font-weight:bold; color:{COLORS[current_color_key]};'>{prefix}: {formatted_price}{full_unit}</span>"
                  tooltip = f"{prefix}: {formatted_price}{full_unit} (Raw: {price_value})"

                  # Update gauge
                  gauge_widget.set_range(gauge_range[0], gauge_range[1])
                  gauge_widget.set_text_format(text_format, unit) # Pass unit to gauge if needed
                  gauge_widget.set_color(COLORS[current_color_key]) # Use determined color
                  gauge_widget.set_value(price_value)

                  # Calculate trend if baseline is provided
                  if trend_baseline is not None:
                       deviation = price_value - trend_baseline
                       # Define threshold relative to baseline (e.g., 0.5% change)
                       threshold = abs(trend_baseline * 0.005)
                       if abs(deviation) > threshold:
                            trend = 1 if deviation > 0 else -1
                            # Calculate percentage change from baseline, scale by sensitivity
                            if trend_baseline != 0: # Avoid division by zero
                                 trend_percent = (deviation / trend_baseline) * 100 * trend_sensitivity
                            else:
                                 trend_percent = 100.0 * trend_sensitivity if trend == 1 else -100.0 * trend_sensitivity


             except Exception as e:
                  print(f"Error updating {prefix} gauge/trend: {e}")
                  styled_text = f"<span style='color:{COLORS['danger']}'>{prefix}: Display Error</span>"
                  gauge_widget.set_value(gauge_widget._minimum) # Reset gauge
                  tooltip = f"{prefix}: Error displaying value {price_value}"

         else: # Handle error or info text directly if no valid price_value
             if value_text.startswith("❌"):
                 styled_text = f"<span style='color:{COLORS['danger']}'>{value_text}</span>"
             elif value_text.startswith("ℹ️"):
                 styled_text = f"<span style='color:{COLORS['info']}; font-size:14pt; font-weight:bold;'>{value_text}</span>"
             else: # Loading or unknown state
                 styled_text = value_text
             gauge_widget.set_value(gauge_widget._minimum) # Reset gauge on error/info without value

         label_widget.setText(styled_text)
         label_widget.setToolTip(tooltip)
         trend_widget.set_trend(trend, trend_percent)

         # Update timestamp on the card
         if card_ref: # Check if the card reference exists
             card_ref.add_footer_text(f"Last updated: {timestamp}")


    def _update_wheat_ui(self, wheat_text, timestamp, price_value=None):
        # Define Wheat specific parameters
        # Assuming yfinance ZW=F price_value is in USD CENTS per bushel.
        display_value = None
        gauge_range_dollars = (3.00, 12.00) # Range in dollars/bu
        baseline_dollars = 6.50 # Baseline in dollars/bu
        if price_value is not None:
             try:
                  # Convert cents to dollars for display/gauge
                  display_value = float(price_value) / 100.0
             except (ValueError, TypeError):
                   print(f"Wheat: Could not convert raw value {price_value} to dollars.")
                   display_value = None # Ensure it's None if conversion fails
                   wheat_text = "❌ Wheat: Invalid Value" # Override text on conversion error


        self._update_commodity_ui(
            label_widget=self.wheat_label,
            gauge_widget=self.wheat_gauge,
            trend_widget=self.wheat_trend,
            card_ref=getattr(self, 'wheat_card', None), # Get card ref safely
            prefix="Wheat",
            value_text=wheat_text, # Pass original text for status check
            timestamp=timestamp,
            price_value=display_value, # Use value in dollars
            color_key='success',
            gauge_range=gauge_range_dollars,
            text_format=".2f", # Format as float with 2 decimals (dollar value)
            currency_symbol="$",
            unit="/bu",
            trend_baseline=baseline_dollars,
            trend_sensitivity=0.2 # Wheat might be less volatile % wise
        )

    def _update_canola_ui(self, canola_text, timestamp, price_value=None):
         # Update card title based on symbol used if possible
         card_title = "Canola Price"
         match = re.search(r'\(([^)]+)\)', canola_text) # Extract symbol like (RS=F) or (Est.)
         if match:
              card_title += f" {match.group(0)}" # Add (Symbol) to title
         # Update title if self.canola_card exists
         if hasattr(self, 'canola_card'):
              self.canola_card.title_label.setText(card_title)

         # Assuming price_value is CAD per tonne
         self._update_commodity_ui(
            label_widget=self.canola_label,
            gauge_widget=self.canola_gauge,
            trend_widget=self.canola_trend,
            card_ref=getattr(self, 'canola_card', None),
            prefix="Canola",
            value_text=canola_text, # Pass original text
            timestamp=timestamp,
            price_value=price_value,
            color_key='secondary', # John Deere Yellow for Canola
            gauge_range=(500, 1200), # Approx range CAD/tonne
            text_format=".2f", # Format as float with 2 decimals
            currency_symbol="$",
            unit="CAD/t",
            trend_baseline=750, # Example baseline price CAD/t
            trend_sensitivity=0.25
        )


    def _update_bitcoin_ui(self, bitcoin_text, timestamp, price_value=None):
        # Bitcoin needs special handling for gauge (value in thousands)
        gauge_value_k = None
        label_price_value = None # Store the original price for the label
        if price_value is not None:
             try:
                  label_price_value = float(price_value)
                  gauge_value_k = label_price_value / 1000.0 # Convert to thousands for gauge
             except (ValueError, TypeError):
                  print(f"Bitcoin: Could not convert raw value {price_value} to float.")
                  label_price_value = None
                  gauge_value_k = None
                  bitcoin_text = "❌ Bitcoin: Invalid Value"


        self._update_commodity_ui(
            label_widget=self.bitcoin_label,
            gauge_widget=self.bitcoin_gauge,
            trend_widget=self.bitcoin_trend,
            card_ref=getattr(self, 'bitcoin_card', None),
            prefix="Bitcoin",
            value_text=bitcoin_text, # Pass original text
            timestamp=timestamp,
            price_value=gauge_value_k, # Pass value in K for gauge logic
            color_key='bitcoin',
            gauge_range=(10, 100), # Range in K USD (10k to 100k)
            text_format="{:.1f}k", # Format gauge as thousands
            currency_symbol="$", # Currency for label only
            unit="USD", # Unit for label only
            trend_baseline=60, # Baseline in K USD (60k)
            trend_sensitivity=1.0 # Bitcoin is volatile
        )
        # Override label text to ensure full $ amount formatting if successful
        if label_price_value is not None and bitcoin_text.startswith("✅"):
             try:
                 # Extract currency from original text if possible, default USD
                 currency = "USD"
                 match = re.search(r'\$([\d,]+\.\d{2})\s*([A-Z]{3})', bitcoin_text)
                 if match:
                      currency = match.group(2)

                 self.bitcoin_label.setText(f"<span style='font-size:14pt; font-weight:bold; color:{COLORS['bitcoin']};'>Bitcoin: ${label_price_value:,.2f} {currency}</span>")
                 self.bitcoin_label.setToolTip(f"Bitcoin: ${label_price_value:,.2f} {currency}")
             except Exception as e:
                  print(f"Error formatting bitcoin label: {e}")
                  # Keep the text generated by _update_commodity_ui as fallback


    # --- Caching Logic ---

    def _is_cache_expired(self, cache_file, hours):
        """Wrapper around DataCache method for instance use"""
        return DataCache.is_cache_expired(cache_file, hours)

    def _load_cached_data(self):
        """Load data from cache files on startup and update UI."""
        print("Loading cached data...")
        now_str = datetime.now().strftime("%H:%M:%S")

        # --- Load Weather ---
        weather_cache = DataCache.load_from_cache(WEATHER_CACHE)
        if weather_cache and not self._is_cache_expired(WEATHER_CACHE, WEATHER_REFRESH_HOURS):
            try:
                results = weather_cache.get('results')
                timestamp = weather_cache.get('timestamp', now_str)
                if isinstance(results, list) and results: # Basic validation
                    print("Using cached weather data.")
                    self.signals.weather_ready.emit(results, timestamp)
                else:
                    print("Invalid weather cache format. Will fetch.")
            except Exception as e:
                print(f"Error processing cached weather data: {e}")
        else:
            print("Weather cache missing or expired. Will fetch.")

        # --- Load Exchange Rate ---
        exchange_cache = DataCache.load_from_cache(EXCHANGE_CACHE)
        if exchange_cache and not self._is_cache_expired(EXCHANGE_CACHE, EXCHANGE_REFRESH_HOURS):
            try:
                rate_text = exchange_cache.get("rate_text", "Cache Error")
                timestamp = exchange_cache.get("timestamp", now_str)
                rate_value = exchange_cache.get("rate_value") # Can be None
                 # Ensure rate_value is float or None
                if rate_value is not None:
                     try:
                          rate_value = float(rate_value)
                     except (ValueError, TypeError):
                          print(f"Warning: Could not convert cached rate_value '{rate_value}' to float.")
                          rate_value = None

                print(f"Using cached exchange data. Value: {rate_value}")
                self.signals.exchange_ready.emit(rate_text, timestamp, rate_value)
            except Exception as e:
                print(f"Error processing cached exchange rate data: {e}")
        else:
            print("Exchange rate cache missing or expired. Will fetch.")

        # --- Load Commodities ---
        commodities_cache = DataCache.load_from_cache(COMMODITIES_CACHE)
        if commodities_cache and not self._is_cache_expired(COMMODITIES_CACHE, COMMODITIES_REFRESH_HOURS):
            print("Using cached commodities data.")
            timestamp = commodities_cache.get("timestamp", now_str)
            # Wheat
            try:
                wheat_text = commodities_cache.get("wheat_text", "Cache Error")
                wheat_value = commodities_cache.get("wheat_price") # Value stored is raw (cents?)
                if wheat_value is not None: wheat_value = float(wheat_value) # Convert
                self.signals.wheat_ready.emit(wheat_text, timestamp, wheat_value)
            except Exception as e:
                print(f"Error processing cached wheat data: {e}")
                self.signals.wheat_ready.emit("❌ Wheat: Cache Read Error", timestamp, None)

            # Canola
            try:
                canola_text = commodities_cache.get("canola_text", "Cache Error")
                canola_value = commodities_cache.get("canola_price")
                if canola_value is not None: canola_value = float(canola_value) # Convert
                self.signals.canola_ready.emit(canola_text, timestamp, canola_value)
            except Exception as e:
                print(f"Error processing cached canola data: {e}")
                self.signals.canola_ready.emit("❌ Canola: Cache Read Error", timestamp, None)

            # Bitcoin
            try:
                bitcoin_text = commodities_cache.get("bitcoin_text", "Cache Error")
                bitcoin_value = commodities_cache.get("bitcoin_price")
                if bitcoin_value is not None: bitcoin_value = float(bitcoin_value) # Convert
                self.signals.bitcoin_ready.emit(bitcoin_text, timestamp, bitcoin_value)
            except Exception as e:
                print(f"Error processing cached bitcoin data: {e}")
                self.signals.bitcoin_ready.emit("❌ Bitcoin: Cache Read Error", timestamp, None)
        else:
            print("Commodities cache missing or expired. Will fetch.")

        print("Finished loading cached data.")


    # --- Async Data Fetching ---

    async def _fetch_all_data(self, fetch_weather=True, fetch_exchange=True, fetch_commodities=True):
        """Gathers async tasks for fetching data."""
        print(f"DEBUG: Starting _fetch_all_data (W:{fetch_weather}, E:{fetch_exchange}, C:{fetch_commodities})")
        tasks = []
        task_names = [] # To identify results/errors

        if fetch_weather:
            tasks.append(asyncio.create_task(self.fetch_weather()))
            task_names.append("Weather")
        if fetch_exchange:
            tasks.append(asyncio.create_task(self.fetch_exchange_rate()))
            task_names.append("Exchange")
        if fetch_commodities:
            tasks.append(asyncio.create_task(self.fetch_commodities()))
            task_names.append("Commodities")

        if not tasks:
            print("DEBUG: No tasks to run in _fetch_all_data.")
            self.signals.refresh_complete.emit() # Still emit complete signal
            return

        print(f"DEBUG: Running {len(tasks)} tasks: {task_names}")
        # Use return_exceptions=True to handle individual task failures
        results = await asyncio.gather(*tasks, return_exceptions=True)

        print("DEBUG: asyncio.gather finished.")
        # Process results and errors
        for i, result in enumerate(results):
            task_name = task_names[i]
            if isinstance(result, Exception):
                print(f"ERROR: Task '{task_name}' failed in _fetch_all_data:")
                # Print exception traceback for detailed debugging
                traceback.print_exception(type(result), result, result.__traceback__)
                # Optionally emit a specific error signal or update UI status
                error_msg = f"❌ {task_name}: Fetch Error ({type(result).__name__})"
                now_str = datetime.now().strftime("%H:%M:%S")
                if task_name == "Weather":
                     # Ensure error is emitted as a list for weather_ready signal
                     self.signals.weather_ready.emit([f"❌ {city}: Fetch Error" for city in self.weather_cities], now_str)
                elif task_name == "Exchange":
                     self.signals.exchange_ready.emit(error_msg, now_str, None)
                elif task_name == "Commodities":
                     # Commodity errors handled internally, but emit generic errors here
                      self.signals.wheat_ready.emit("❌ Wheat: Fetch Error", now_str, None)
                      self.signals.canola_ready.emit("❌ Canola: Fetch Error", now_str, None)
                      self.signals.bitcoin_ready.emit("❌ Bitcoin: Fetch Error", now_str, None)
            else:
                print(f"DEBUG: Task '{task_name}' completed successfully.")
                # Successful results are handled by signals emitted within the fetch methods themselves

        print("DEBUG: Emitting refresh_complete signal.")
        self.signals.refresh_complete.emit()
        print("DEBUG: _fetch_all_data finished.")


    def _weather_icon(self, description: str) -> str:
        """Helper to get emoji for weather description (keep simple)."""
        desc = description.lower() if description else ""
        if "clear" in desc: return "☀️"
        if "sun" in desc and "cloud" in desc: return "⛅" # Partly cloudy
        if "few clouds" in desc: return "🌤️"
        if "scattered clouds" in desc: return "☁️"
        if "broken clouds" in desc: return "☁️"
        if "overcast clouds" in desc: return "🌥️"
        if "shower rain" in desc: return "🌦️"
        if "rain" in desc: return "🌧️"
        if "drizzle" in desc: return "🌦️"
        if "snow" in desc: return "❄️"
        if "thunder" in desc: return "⛈️"
        if "fog" in desc or "mist" in desc or "haze" in desc: return "🌫️"
        return "🌡️" # Default


    async def fetch_weather(self):
        """Fetches weather data for multiple cities using OpenWeatherMap."""
        print(f"Fetching weather...")
        if not OPENWEATHER_API_KEY or OPENWEATHER_API_KEY == "YOUR_API_KEY":
            msg = [f"❌ {city}: API Key Missing" for city in self.weather_cities]
            print("OpenWeather API key is missing or invalid.")
            self.signals.weather_ready.emit(msg, datetime.now().strftime("%H:%M:%S"))
            return

        results = []
        now_str = datetime.now().strftime("%H:%M:%S")
        # Limit concurrent connections to avoid hitting API limits quickly
        connector = aiohttp.TCPConnector(limit=4, ssl=False) # ssl=False might be needed sometimes, but use with caution
        async with aiohttp.ClientSession(connector=connector) as session:
             fetch_tasks = []
             for city_query in self.weather_cities: # Use the list defined in init
                 params = {
                     'q': f"{city_query},CA", # Add ,CA for Canadian cities
                     'appid': OPENWEATHER_API_KEY,
                     'units': 'metric'
                 }
                 task = asyncio.create_task(session.get(OPENWEATHER_BASE_URL, params=params, timeout=10))
                 fetch_tasks.append((city_query, task)) # Store city name with task

             city_results = {} # Store results per city
             # Use asyncio.gather to run all city requests concurrently
             responses = await asyncio.gather(*[t[1] for t in fetch_tasks], return_exceptions=True)

             # Process responses
             for i, response_or_exc in enumerate(responses):
                city_name = fetch_tasks[i][0] # Get city name associated with this response
                result_str = f"❌ {city_name}: Unknown Error" # Default error

                try:
                    if isinstance(response_or_exc, Exception):
                         # Handle exceptions like timeouts or connection errors
                         raise response_or_exc

                    # Process successful response
                    response = response_or_exc
                    status = response.status
                    print(f"Weather response status for {city_name}: {status}")

                    if status == 200:
                         data = await response.json()
                         try:
                              main_data = data.get('main', {})
                              weather_data = data.get('weather', [{}])[0] # Get first weather item

                              temp = main_data.get('temp')
                              # temp_min = main_data.get('temp_min', temp) # Optional H/L
                              # temp_max = main_data.get('temp_max', temp) # Optional H/L
                              desc = weather_data.get('description', 'N/A')
                              icon = self._weather_icon(desc)
                              feels_like = main_data.get('feels_like', temp) # Use feels_like temp

                              if temp is None or feels_like is None:
                                   result_str = f"❌ {city_name}: Temp data missing"
                              else:
                                   # Format string carefully
                                   result_str = (f"✅ {city_name}: {temp:.0f}°C "
                                                f"(Feels {feels_like:.0f}°), "
                                                f"{desc.capitalize()} {icon}")
                         except (KeyError, IndexError, TypeError) as parse_e:
                              print(f"Error parsing weather data for {city_name}: {parse_e}")
                              result_str = f"❌ {city_name}: Data parse error"
                    # Handle specific HTTP errors
                    elif status == WeatherAPIErrors.INVALID_API_KEY:
                         result_str = f"❌ {city_name}: Invalid API Key"
                    elif status == WeatherAPIErrors.NOT_FOUND:
                         result_str = f"❌ {city_name}: City not found"
                    elif status == WeatherAPIErrors.RATE_LIMIT:
                         result_str = f"❌ {city_name}: Rate limit hit"
                    else:
                         error_text = await response.text()
                         print(f"Weather API Error for {city_name} ({status}): {error_text[:200]}") # Log part of error text
                         result_str = f"❌ {city_name}: API Error {status}"

                # Handle exceptions from asyncio.gather or response processing
                except asyncio.TimeoutError:
                    print(f"Timeout fetching weather for {city_name}")
                    result_str = f"❌ {city_name}: Request Timeout"
                except aiohttp.ClientError as ce:
                     # Includes connection errors, etc.
                     print(f"ClientError fetching weather for {city_name}: {ce}")
                     result_str = f"❌ {city_name}: Network Error"
                except Exception as e:
                    print(f"Generic error processing weather for {city_name}: {e}")
                    traceback.print_exc()
                    result_str = f"❌ {city_name}: Fetch Error ({type(e).__name__})"
                finally:
                     # Ensure response is closed if it exists and is a response object
                     if 'response' in locals() and isinstance(response, aiohttp.ClientResponse) and not response.closed:
                          response.release()

                city_results[city_name] = result_str

        # Ensure results are in the original city order for the UI
        final_results = [city_results.get(city, f"❌ {city}: No response") for city in self.weather_cities]

        # Emit signal and save cache
        self.signals.weather_ready.emit(final_results, now_str)
        weather_cache_data = {
            "results": final_results,
            "timestamp": now_str,
            "fetch_time": datetime.now().isoformat()
        }
        DataCache.save_to_cache(weather_cache_data, WEATHER_CACHE)
        print("Weather fetch complete.")


    async def fetch_exchange_rate(self):
        """Fetches USD to CAD exchange rate using AlphaVantage."""
        print("Fetching exchange rate...")
        now_str = datetime.now().strftime("%H:%M:%S")
        result_text = "❌ USD-CAD: Unknown Error"
        rate_value = None

        if not ALPHAVANTAGE_API_KEY or ALPHAVANTAGE_API_KEY == "YOUR_API_KEY": # Basic check
            result_text = "❌ USD-CAD: AlphaVantage API key missing/invalid"
            print(result_text)
            self.signals.exchange_ready.emit(result_text, now_str, None)
            return

        params = {
            "function": "CURRENCY_EXCHANGE_RATE",
            "from_currency": "USD",
            "to_currency": "CAD",
            "apikey": ALPHAVANTAGE_API_KEY
        }

        try:
            # Use ssl=False cautiously if encountering SSL verification issues
            connector = aiohttp.TCPConnector(ssl=False)
            async with aiohttp.ClientSession(connector=connector) as session:
                print(f"Requesting Exchange Rate from: {ALPHAVANTAGE_BASE_URL}")
                async with session.get(ALPHAVANTAGE_BASE_URL, params=params, timeout=15) as resp: # Increased timeout
                    status = resp.status
                    print(f"Exchange Rate response status: {status}")
                    # Read response text first for better error diagnosis
                    response_text = await resp.text()
                    # print(f"Exchange Rate response text: {response_text[:500]}") # Debug: print partial response

                    if status == 200:
                        try:
                            data = json.loads(response_text) # Parse JSON from text

                            # Check for API error messages within the JSON response
                            if "Error Message" in data:
                                error_msg = data["Error Message"]
                                print(f"AlphaVantage API Error: {error_msg}")
                                result_text = f"❌ USD-CAD: API Error" # Keep UI msg short
                            elif "Information" in data: # Often indicates rate limit or usage issues
                                info_msg = data["Information"]
                                print(f"AlphaVantage API Info: {info_msg}")
                                result_text = f"❌ USD-CAD: API Limit/Info" # Keep UI msg short
                            elif "Note" in data: # Can indicate usage limit
                                note_msg = data["Note"]
                                print(f"AlphaVantage API Note: {note_msg}")
                                result_text = f"❌ USD-CAD: API Note" # Keep UI msg short
                            elif "Realtime Currency Exchange Rate" in data:
                                rate_data = data.get("Realtime Currency Exchange Rate", {})
                                rate_str = rate_data.get("5. Exchange Rate")
                                if rate_str:
                                    try:
                                        rate_value = float(rate_str)
                                        # Success text is generated in the UI update function
                                        result_text = f"✅ USD-CAD: {rate_value:.4f}" # Internal success marker
                                        print(f"Successfully fetched exchange rate: {rate_value}")
                                    except (ValueError, TypeError) as conv_e:
                                        print(f"Could not convert exchange rate '{rate_str}' to float: {conv_e}")
                                        result_text = f"❌ USD-CAD: Invalid rate format ({rate_str})"
                                else:
                                    result_text = "❌ USD-CAD: Rate field missing"
                            else:
                                print("Exchange rate data block not found in response.")
                                result_text = "❌ USD-CAD: Unexpected format"

                        except json.JSONDecodeError as json_e:
                             print(f"Failed to decode JSON response: {json_e}")
                             result_text = f"❌ USD-CAD: Invalid API response"
                             print(f"--- Raw Response Text Start ---\n{response_text[:500]}\n--- Raw Response Text End ---")

                    else: # Handle non-200 HTTP status codes
                        result_text = f"❌ USD-CAD: API HTTP Error {status}"

        except asyncio.TimeoutError:
            print("Timeout fetching exchange rate.")
            result_text = "❌ USD-CAD: Request Timeout"
        except aiohttp.ClientError as ce:
             print(f"ClientError fetching exchange rate: {ce}")
             result_text = f"❌ USD-CAD: Network Error"
        except Exception as e:
            print(f"Generic error fetching exchange rate: {e}")
            traceback.print_exc()
            result_text = f"❌ USD-CAD: Fetch Error"

        # Emit signal and save cache
        self.signals.exchange_ready.emit(result_text, now_str, rate_value)
        exchange_cache_data = {
            "rate_text": result_text,
            "rate_value": rate_value,
            "timestamp": now_str,
            "fetch_time": datetime.now().isoformat()
        }
        DataCache.save_to_cache(exchange_cache_data, EXCHANGE_CACHE)
        print("Exchange rate fetch complete.")


    async def get_canola_price_fallback(self):
        """Fallback method attempts to scrape a website for Canola prices."""
        # --- IMPORTANT ---
        # Web scraping is fragile and depends heavily on the target website's structure.
        # This is a placeholder and needs a real target URL and parsing logic.
        # It's also important to respect the website's terms of service (robots.txt).
        # Using a dedicated financial data provider is usually more reliable.
        print("Attempting fallback method to get canola price via scraping...")
        fallback_price = None
        # --- !!! REPLACE THIS URL !!! ---
        url = "https://www.grainscanada.gc.ca/en/grain-markets/canola-prices/" # Example URL (check if still valid/useful)
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36'} # Mimic browser

        try:
            connector = aiohttp.TCPConnector(ssl=False) # Often needed for scraping sites with varying SSL certs
            async with aiohttp.ClientSession(headers=headers, connector=connector) as session:
                async with session.get(url, timeout=15) as response:
                    print(f"Fallback scrape status for {url}: {response.status}")
                    if response.status == 200:
                        html = await response.text()
                        # --- Parsing Logic (NEEDS ADJUSTMENT FOR ACTUAL SITE) ---
                        # Example using regex (adjust pattern based on actual site):
                        # Look for something like <td class="price-value">$750.50</td>
                        # This pattern is highly specific and likely WRONG for the example URL.
                        match = re.search(r'>\s*\$(\d{3,4}\.\d{2})\s*<', html) # Generic search for $XXX.XX pattern
                        if match:
                            price_str = match.group(1)
                            try:
                                fallback_price = float(price_str.replace(',', '')) # Remove commas if present
                                print(f"Fallback scrape successful: Found price {fallback_price}")
                            except ValueError:
                                print(f"Fallback scrape: Found price text '{price_str}' but failed to convert.")
                        else:
                            print("Fallback scrape: Price pattern not found on page.")
                            # Alternative: Use BeautifulSoup4 if installed
                            # try:
                            #      from bs4 import BeautifulSoup
                            #      soup = BeautifulSoup(html, 'html.parser')
                            #      # Find element based on actual site structure (e.g., by ID, class, tag)
                            #      price_element = soup.find('span', class_='actual-price-class')
                            #      if price_element:
                            #           price_text = price_element.get_text(strip=True)
                            #           # Extract float from text...
                            # except ImportError:
                            #      print("BeautifulSoup4 not installed, cannot use for fallback.")
                            # except Exception as bs_e:
                            #      print(f"Error during BeautifulSoup parsing: {bs_e}")

                    else:
                        print(f"Fallback scrape failed: HTTP {response.status}")

        except asyncio.TimeoutError:
            print("Fallback scrape timed out.")
        except aiohttp.ClientError as ce:
             print(f"Fallback scrape network error: {ce}")
        except Exception as e:
            print(f"Fallback scrape error: {e}")
            traceback.print_exc()

        # Use hardcoded value only if scraping fails completely
        if fallback_price is None:
             hardcoded_fallback = 695.25 # Your last known good price
             print(f"Fallback scraping failed. Using hardcoded value: {hardcoded_fallback}")
             # return hardcoded_fallback # Uncomment to return hardcoded value
             return None # Prefer returning None if scraping fails
        else:
             return fallback_price


    async def fetch_commodities(self):
        """Fetches Wheat, Canola, and Bitcoin prices using yfinance."""
        print("Fetching commodity prices...")
        loop = asyncio.get_event_loop()
        now_str = datetime.now().strftime("%H:%M:%S")

        wheat_text = "❌ Wheat: Fetch Error"
        wheat_value = None
        canola_text = "❌ Canola: Fetch Error"
        canola_value = None
        canola_symbol_used = "N/A"
        bitcoin_text = "❌ Bitcoin: Fetch Error"
        bitcoin_value = None

        # --- Enhanced get_price function using yfinance Ticker ---
        def get_yfinance_data(symbol):
            """Fetches last closing price and potentially other info using yfinance."""
            print(f"yfinance: Fetching data for symbol '{symbol}'...")
            price = None
            info = {}
            try:
                ticker = yf.Ticker(symbol)

                # Try getting 'fast_info' first (less data, potentially quicker)
                try:
                     fast_info = ticker.fast_info
                     price = fast_info.get('lastPrice')
                     print(f"yfinance: Got fast_info for {symbol}. Last Price: {price}")
                     info = {k: fast_info.get(k) for k in ['currency', 'exchange', 'quoteType', 'lastPrice', 'previousClose']}
                except Exception as fast_info_e:
                     print(f"yfinance: Could not get fast_info for {symbol}: {fast_info_e}. Trying history...")
                     price = None # Reset price if fast_info failed

                # If fast_info didn't provide price, try history
                if price is None:
                    hist = ticker.history(period="5d", interval="1d", auto_adjust=True)

                    if hist is not None and not hist.empty and 'Close' in hist.columns:
                        last_close = hist['Close'].iloc[-1]
                        if pd.notna(last_close):
                            price = last_close
                            print(f"yfinance: Got price {price} from history for {symbol}")
                        else:
                            print(f"yfinance: Last closing price in history for {symbol} is NaN.")
                    else:
                        print(f"yfinance: Could not get valid history or 'Close' price for {symbol}. Trying info dict...")
                        # If history fails, try getting the full info dict as a last resort
                        try:
                             full_info = ticker.info
                             info = {k: full_info.get(k) for k in ['symbol', 'currency', 'exchange', 'quoteType', 'marketState', 'shortName', 'regularMarketPrice', 'previousClose', 'bid', 'ask']}
                             price_keys = ['regularMarketPrice', 'currentPrice', 'previousClose', 'bid', 'ask'] # Order of preference
                             for key in price_keys:
                                  if key in full_info and full_info[key] is not None and isinstance(full_info[key], (int, float)):
                                       price = float(full_info[key])
                                       print(f"yfinance: Got price {price} from full_info key '{key}' for {symbol}")
                                       break
                             if price is None:
                                  print(f"yfinance: No usable price found in full_info for {symbol}")

                        except Exception as full_info_e:
                             print(f"yfinance: Error getting full_info for {symbol}: {full_info_e}")

            except requests.exceptions.RequestException as req_e:
                 print(f"yfinance: Network error fetching data for {symbol}: {req_e}")
            except Exception as e:
                print(f"yfinance: Generic error fetching data for {symbol}: {e}")
                traceback.print_exc()

            # Ensure price is float or None
            if price is not None:
                 try:
                      price = float(price)
                 except (ValueError, TypeError):
                      print(f"yfinance: Could not convert final price '{price}' to float for {symbol}. Setting to None.")
                      price = None

            print(f"yfinance: Finished fetching for {symbol}. Price found: {price}")
            return price, info # Return both price and info dict

        # --- Fetch Wheat ---
        try:
            wheat_result = await loop.run_in_executor(None, get_yfinance_data, WHEAT_SYMBOL)
            wheat_value, wheat_info = wheat_result # wheat_value is raw price (likely cents)
            if wheat_value is not None:
                 # Assuming ZW=F value is in USD CENTS per bushel.
                 wheat_value_dollars = wheat_value / 100.0
                 wheat_text = f"✅ Wheat: ${wheat_value_dollars:,.2f} /bu"
            else:
                 wheat_text = f"❌ Wheat: Price N/A"
        except Exception as e:
            print(f"ERROR: Wheat fetch failed - {e}")
            traceback.print_exc()
            wheat_text = f"❌ Wheat: Error ({type(e).__name__})"
            wheat_value = None # Ensure value is None on error
        finally:
             # Emit wheat signal immediately after fetch attempt
             # Pass the RAW value (cents?) to the signal; conversion happens in UI update
             self.signals.wheat_ready.emit(wheat_text, now_str, wheat_value)

        # --- Fetch Bitcoin ---
        try:
            bitcoin_result = await loop.run_in_executor(None, get_yfinance_data, BITCOIN_SYMBOL)
            bitcoin_value, bitcoin_info = bitcoin_result
            if bitcoin_value is not None:
                 currency = bitcoin_info.get('currency', 'USD')
                 bitcoin_text = f"✅ Bitcoin: ${bitcoin_value:,.2f} {currency}"
            else:
                 bitcoin_text = f"❌ Bitcoin: Price N/A"
        except Exception as e:
            print(f"ERROR: Bitcoin fetch failed - {e}")
            traceback.print_exc()
            bitcoin_text = f"❌ Bitcoin: Error ({type(e).__name__})"
            bitcoin_value = None
        finally:
            # Emit bitcoin signal immediately
            self.signals.bitcoin_ready.emit(bitcoin_text, now_str, bitcoin_value)


        # --- Fetch Canola (Try multiple symbols) ---
        print(f"Attempting to fetch canola price using {len(CANOLA_SYMBOLS)} symbols...")
        canola_price_found = False
        for symbol in CANOLA_SYMBOLS:
            try:
                print(f"Trying canola symbol: {symbol}")
                canola_result = await loop.run_in_executor(None, get_yfinance_data, symbol)
                canola_value_attempt, canola_info = canola_result

                if canola_value_attempt is not None:
                    canola_value = canola_value_attempt
                    currency = canola_info.get('currency', 'CAD') # Assume CAD if not specified
                    # Assuming price is CAD per tonne
                    canola_text = f"✅ Canola ({symbol}): ${canola_value:,.2f} {currency}/t"
                    canola_symbol_used = symbol
                    canola_price_found = True
                    print(f"Found working canola symbol: {symbol} -> Price: {canola_value}")
                    break # Stop after finding the first working symbol
                else:
                    print(f"Symbol {symbol} returned None price.")

            except Exception as e:
                 print(f"Failed fetching canola with symbol '{symbol}': {e}")

        # --- Canola Fallback ---
        if not canola_price_found:
            print("All yfinance canola symbols failed. Attempting fallback method...")
            try:
                canola_fallback_value = await self.get_canola_price_fallback()
                if canola_fallback_value is not None:
                    canola_value = canola_fallback_value
                    canola_text = f"ℹ️ Canola (Est.): ${canola_value:,.2f} CAD/t"
                    canola_symbol_used = "FALLBACK/SCRAPE"
                    print(f"Canola fallback successful: Price {canola_value}")
                else:
                     print("Canola fallback method also failed.")
                     canola_text = f"❌ Canola: Data unavailable"
                     canola_symbol_used = "FAILED_ALL"
            except Exception as e:
                print(f"Canola fallback method itself raised an error: {e}")
                traceback.print_exc()
                canola_text = f"❌ Canola: Fallback Error"
                canola_symbol_used = "FALLBACK_ERROR"
                canola_value = None

        # Emit canola signal finally
        self.signals.canola_ready.emit(canola_text, now_str, canola_value)

        # --- Cache Results ---
        commodities_cache_data = {
            "wheat_text": wheat_text,
            "wheat_price": wheat_value, # Store the raw value from yfinance (cents?)
            "canola_text": canola_text,
            "canola_price": canola_value,
            "bitcoin_text": bitcoin_text,
            "bitcoin_price": bitcoin_value,
            "timestamp": now_str,
            "fetch_time": datetime.now().isoformat(),
            "canola_symbol_used": canola_symbol_used,
        }
        DataCache.save_to_cache(commodities_cache_data, COMMODITIES_CACHE)
        print("Commodities fetch complete.")


    # --- Cleanup ---
    def closeEvent(self, event):
        """Clean up resources when the widget (and application) is closed."""
        print("HomeModule closeEvent triggered.")

        # Stop the scheduled refresh thread gracefully
        if hasattr(self, 'scheduled_refresh') and self.scheduled_refresh.isRunning():
            print("Stopping scheduled refresh thread...")
            self.scheduled_refresh.stop()
            if not self.scheduled_refresh.wait(2000): # Wait briefly for it to finish
                 print("Warning: Scheduled refresh thread did not stop gracefully. Terminating.")
                 self.scheduled_refresh.terminate() # Force terminate if needed

        # Stop any running data fetcher thread
        if self.fetcher_thread and self.fetcher_thread.isRunning():
            print("Stopping active data fetcher thread...")
            self.fetcher_thread.stop() # Signal the thread to stop
            if not self.fetcher_thread.wait(3000): # Wait a bit longer for async tasks
                 print("Warning: Data fetcher thread did not stop gracefully. Terminating.")
                 self.fetcher_thread.terminate()

        print("HomeModule cleanup finished.")
        # Let the event continue to parent handlers (important for application exit)
        super().closeEvent(event)

# Example usage (for testing standalone)
if __name__ == '__main__':
    from PyQt5.QtWidgets import QApplication, QMainWindow

    # Dummy MainWindow class for testing
    class DummyMainWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Home Module Test")
            self.home_module = HomeModule(self)
            self.setCentralWidget(self.home_module)
            self.resize(800, 600)
            # Ensure cache directory exists for testing
            DataCache.ensure_cache_dir()

    app = QApplication(sys.argv)
    window = DummyMainWindow()
    window.show()
    sys.exit(app.exec_())
