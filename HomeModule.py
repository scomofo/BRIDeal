# -*- coding: utf-8 -*-
# HomeModule.py - v1.1 (Integrated UI update slots, fixed syntax, added logging/slots)
import sys
import os
import traceback
import asyncio
import time
import re
import math
import json
from datetime import datetime, timedelta
import logging
from typing import Optional, List, Dict, Any, Tuple

from PyQt5.QtWidgets import (
    QWidget, QLabel, QGridLayout, QVBoxLayout, QHBoxLayout, QLayout,
    QFrame, QSizePolicy, QGroupBox, QPushButton, QComboBox,
    QSpacerItem, QGraphicsDropShadowEffect, QProgressBar
)
from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread, QSize, QTimer, QPoint, pyqtSlot
from PyQt5.QtGui import (
    QColor, QFont, QIcon, QPalette, QPainter, QBrush, QPen, QLinearGradient
)

# Import new data sources
try:
    import finnhub  # For Finance data
except ImportError:
    logging.critical("CRITICAL: 'finnhub' library not found. Install with 'pip install finnhub-python'")
    finnhub = None

try:
    from weather import Tools as WeatherTools  # For Weather data (using Open-Meteo)
except ImportError:
    WeatherTools = None  # Set to None if import fails

# Attempt to import config safely
try:
    from config import apis, paths, settings
    FINNHUB_API_KEY = getattr(apis, 'FINNHUB_API_KEY', None)
    FX_SYMBOL_USDCAD = getattr(apis, 'FX_SYMBOL_USDCAD', 'OANDA:USD_CAD')
    WHEAT_SYMBOL = getattr(apis, 'WHEAT_SYMBOL', 'futures/W')
    CANOLA_SYMBOL = getattr(apis, 'CANOLA_SYMBOL', 'commodities/canola')
    BITCOIN_SYMBOL = getattr(apis, 'BITCOIN_SYMBOL', 'BINANCE:BTCUSDT')
    CACHE_DIR = getattr(paths, 'CACHE_DIR', './cache')
    WEATHER_REFRESH_HOURS = getattr(settings, 'WEATHER_REFRESH_INTERVAL', 1)
    EXCHANGE_REFRESH_HOURS = getattr(settings, 'EXCHANGE_REFRESH_INTERVAL', 6)
    COMMODITIES_REFRESH_HOURS = getattr(settings, 'COMMODITIES_REFRESH_INTERVAL', 4)
    API_TIMEOUT = getattr(settings, 'API_TIMEOUT', 15)
    WEATHER_CITIES = getattr(settings, 'WEATHER_CITIES', ["Calgary", "Edmonton", "Vancouver"])
except ImportError:
    logging.critical("CRITICAL: Missing required config.py module or some attributes. Using defaults.")
    FINNHUB_API_KEY = None
    FX_SYMBOL_USDCAD = 'OANDA:USD_CAD'
    WHEAT_SYMBOL = 'futures/W'
    CANOLA_SYMBOL = 'commodities/canola'
    BITCOIN_SYMBOL = 'BINANCE:BTCUSDT'
    CACHE_DIR = './cache'
    WEATHER_REFRESH_HOURS = 1
    EXCHANGE_REFRESH_HOURS = 6
    COMMODITIES_REFRESH_HOURS = 4
    API_TIMEOUT = 15
    WEATHER_CITIES = ["Calgary", "Edmonton", "Vancouver"]

# --- Setup Logger ---
logger = logging.getLogger(__name__)

# --- Constants and Setup ---
WEATHER_CACHE = os.path.join(CACHE_DIR, "weather_cache.json")
EXCHANGE_CACHE = os.path.join(CACHE_DIR, "exchange_cache.json")
COMMODITIES_CACHE = os.path.join(CACHE_DIR, "commodities_cache.json")

# Define colors (used in UI elements)
COLORS = {
    'primary': '#FFD700',
    'accent': '#003366',
    'text_primary': '#333333',
    'text_secondary': '#666666',
    'background': '#F8F9FA',
    'card_bg': '#FFFFFF',
    'success': '#28a745',
    'danger': '#dc3545',
    'warning': '#ffc107',
    'info': '#17a2b8',
    'secondary': '#FFD700',
    'bitcoin': '#F7931A'
}

# Cache directory setup
try:
    os.makedirs(CACHE_DIR, exist_ok=True)
    logger.info(f"Ensured cache directory: {CACHE_DIR}")
except OSError as e:
    logger.error(f"Error creating cache directory {CACHE_DIR}: {e}")

# Data cache utility class
class DataCache:
    @staticmethod
    def ensure_cache_dir():
        try:
            os.makedirs(CACHE_DIR, exist_ok=True)
        except OSError as e:
            logger.error(f"Error ensuring cache directory {CACHE_DIR}: {e}")

    @staticmethod
    def save_to_cache(data, cache_file):
        DataCache.ensure_cache_dir()
        try:
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            logger.info(f"Data saved to {os.path.basename(cache_file)}")
        except Exception as e:
            logger.error(f"Error saving cache to {os.path.basename(cache_file)}: {e}", exc_info=True)

    @staticmethod
    def load_from_cache(cache_file):
        try:
            if os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                logger.info(f"Data loaded from {os.path.basename(cache_file)}")
                return data
            else:
                logger.debug(f"Cache file {os.path.basename(cache_file)} not found.")
                return None
        except json.JSONDecodeError as e:
            logger.error(f"Error decoding JSON from {os.path.basename(cache_file)}: {e}")
            return None
        except Exception as e:
            logger.error(f"Error loading cache from {os.path.basename(cache_file)}: {e}", exc_info=True)
            return None

    @staticmethod
    def is_cache_expired(cache_file, hours):
        if not os.path.exists(cache_file):
            logger.debug(f"Cache file {os.path.basename(cache_file)} absent, expired.")
            return True
        try:
            file_mod_time = os.path.getmtime(cache_file)
            expiry_time = file_mod_time + hours * 3600
            is_expired = time.time() > expiry_time
            logger.debug(f"Cache check {os.path.basename(cache_file)}: Mod={datetime.fromtimestamp(file_mod_time):%H:%M}, Exp={datetime.fromtimestamp(expiry_time):%H:%M}, Expired={is_expired}")
            return is_expired
        except Exception as e:
            logger.error(f"Error checking cache expiration for {os.path.basename(cache_file)}: {e}")
            return True

class HomeSignals(QObject):
    weather_ready = pyqtSignal(list, str)
    exchange_ready = pyqtSignal(str, str, object)
    wheat_ready = pyqtSignal(str, str, object)
    canola_ready = pyqtSignal(str, str, object)
    bitcoin_ready = pyqtSignal(str, str, object)
    status_update = pyqtSignal(str)
    refresh_complete = pyqtSignal()

class DataFetcherThread(QThread):
    def __init__(self, home_module, fetch_weather=True, fetch_exchange=True, fetch_commodities=True):
        super().__init__()
        self.home_module = home_module
        self.fetch_weather = fetch_weather
        self.fetch_exchange = fetch_exchange
        self.fetch_commodities = fetch_commodities
        self._is_running = True
        self._loop = None
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")

    def run(self):
        self.logger.info(f"Run START (W:{self.fetch_weather}, E:{self.fetch_exchange}, C:{self.fetch_commodities})")
        try:
            self.logger.debug("Creating new asyncio event loop...")
            self._loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._loop)
            self.logger.debug("Event loop set.")
            if self._is_running:
                self.logger.debug("Emitting status: Starting fetch...")
                self.home_module.signals.status_update.emit("Starting data fetch...")
                self.logger.debug("Calling loop.run_until_complete(_fetch_all_data)...")
                results = self._loop.run_until_complete(self.home_module._fetch_all_data(
                    self.fetch_weather, self.fetch_exchange, self.fetch_commodities))
                self.logger.info(f"_fetch_all_data completed. Result count: {len(results) if isinstance(results, list) else 'N/A'}")
            else:
                self.home_module.signals.status_update.emit("Data fetch cancelled before start.")
        except Exception as e:
            self.logger.error(f"Error in run method: {e}", exc_info=True)
            self.home_module.signals.status_update.emit(f"Fetch Error: {e}")
        finally:
            self.logger.debug("Finally block entered.")
            if self._loop:
                self.logger.debug("Attempting loop shutdown...")
                try:
                    tasks = asyncio.all_tasks(self._loop)
                    to_cancel = [task for task in tasks if not task.done()]
                    if to_cancel:
                        self.logger.debug(f"Cancelling {len(to_cancel)} tasks...")
                        for task in to_cancel:
                            task.cancel()
                        self.logger.debug("Waiting for cancelled tasks...")
                        self._loop.run_until_complete(asyncio.gather(*to_cancel, return_exceptions=True))
                        self.logger.debug("Cancelled tasks gather complete.")
                    else:
                        self.logger.debug("No running tasks to cancel.")
                    self.logger.debug("Shutting down async generators...")
                    self._loop.run_until_complete(self._loop.shutdown_asyncgens())
                    self.logger.debug("Async generators shutdown complete.")
                except Exception as shutdown_e:
                    self.logger.error(f"Error during loop task cancellation/shutdown: {shutdown_e}", exc_info=True)
                finally:
                    if self._loop.is_running():
                        self.logger.debug("Stopping loop...")
                        self._loop.stop()
                    self.logger.debug("Closing loop...")
                    self._loop.close()
                    self.logger.info("Event loop closed.")
            else:
                self.logger.debug("No loop found in finally.")
            asyncio.set_event_loop(None)
            self._loop = None
            self._is_running = False
            self.logger.info("Run FINISHED.")

    def stop(self):
        self.logger.info("Stop requested...")
        self._is_running = False
        if self._loop and self._loop.is_running():
            self.logger.debug("Scheduling task cancellation via call_soon_threadsafe...")
            self._loop.call_soon_threadsafe(self._cancel_all_tasks_threadsafe)
        elif self._loop:
            self.logger.warning("Loop exists but not running during stop request.")
        else:
            self.logger.warning("Loop not found during stop request.")

    def _cancel_all_tasks_threadsafe(self):
        self.logger.debug("_cancel_all_tasks_threadsafe executing in loop thread.")
        if not self._loop:
            self.logger.warning("Loop is None in _cancel_all_tasks_threadsafe.")
            return
        tasks = asyncio.all_tasks(self._loop)
        to_cancel = [task for task in tasks if not task.done()]
        if to_cancel:
            self.logger.debug(f"Cancelling {len(to_cancel)} tasks from loop thread...")
            count = 0
            for task in to_cancel:
                if task.cancel():
                    count += 1
            self.logger.debug(f"{count} tasks marked for cancellation.")
        else:
            self.logger.debug("No running tasks found to cancel in loop thread.")

class ScheduledRefreshThread(QThread):
    refresh_signal = pyqtSignal(bool, bool, bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = True
        self.last_refresh_times = {'weather': datetime.min, 'exchange': datetime.min, 'commodities': datetime.min}
        self.refresh_intervals_hours = {
            'weather': WEATHER_REFRESH_HOURS,
            'exchange': EXCHANGE_REFRESH_HOURS,
            'commodities': COMMODITIES_REFRESH_HOURS
        }
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self.logger.info(f"Initialized with intervals (Hrs): W={WEATHER_REFRESH_HOURS}, E={EXCHANGE_REFRESH_HOURS}, C={COMMODITIES_REFRESH_HOURS}")

    def run(self):
        self.logger.info("Run START.")
        while self._running:
            now = datetime.now()
            refresh_tasks = {}
            for task_name, last_time in self.last_refresh_times.items():
                interval_hours = self.refresh_intervals_hours.get(task_name, 1)
                if interval_hours <= 0:
                    continue
                should_refresh = (now - last_time) > timedelta(hours=interval_hours)
                if should_refresh:
                    refresh_tasks[task_name] = True
                    self.last_refresh_times[task_name] = now
                    self.logger.info(f"TRIGGERED for {task_name.capitalize()} (Last: {last_time:%Y-%m-%d %H:%M}, Interval: {interval_hours}h)")
            if refresh_tasks:
                wf, ef, cf = (
                    refresh_tasks.get('weather', False),
                    refresh_tasks.get('exchange', False),
                    refresh_tasks.get('commodities', False)
                )
                self.refresh_signal.emit(wf, ef, cf)
            sleep_duration = 300
            count = 0
            while self._running and count < sleep_duration:
                time.sleep(1)
                count += 1
            if not self._running:
                self.logger.debug("Sleep interrupted by stop request.")
                break
        self.logger.info("Run FINISHED.")

    def stop(self):
        self.logger.info("Stop requested...")
        self._running = False

# --- Custom UI Widgets ---

class StyledCard(QFrame):
    """A styled container widget resembling a card."""
    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet(f"""
            StyledCard {{
                background-color: {COLORS['card_bg']};
                border-radius: 8px;
                border: 1px solid #e0e0e0;
            }}
        """)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(15, 15, 15, 15)
        self.main_layout.setSpacing(10)
        self.title_label = QLabel(title)
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        self.title_label.setStyleSheet(f"color: {COLORS['accent']}; border: none; background: none;")
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(0, 5, 0, 0)
        self.content_layout.setSpacing(5)
        self.footer_label = QLabel("")
        self.footer_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.footer_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 9pt; border: none; background: none; padding-top: 5px;")
        self.main_layout.addWidget(self.title_label)
        self.main_layout.addWidget(self.content_widget)
        self.main_layout.addStretch(1)
        self.main_layout.addWidget(self.footer_label)

    def add_content(self, widget: QWidget or QLayout):
        if isinstance(widget, QWidget):
            self.content_layout.addWidget(widget)
        elif isinstance(widget, QLayout):
            container = QWidget()
            container.setLayout(widget)
            self.content_layout.addWidget(container)
        else:
            self.logger.error(f"add_content expects QWidget or QLayout, got {type(widget)}")

    def add_footer_text(self, text):
        self.footer_label.setText(str(text))

    def clear_content(self):
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

class StyledButton(QPushButton):
    """A styled push button with primary and secondary styles."""
    def __init__(self, text, icon_path=None, is_primary=True, parent=None):
        super().__init__(text, parent)
        self.setCursor(Qt.PointingHandCursor)
        if is_primary:
            bg_color = COLORS['primary']
            hover_color = COLORS['accent']
            text_color = "white"
        else:
            bg_color = "#FFFFFF"
            hover_color = "#E5C700"
            text_color = COLORS['accent']
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {bg_color};
                color: {text_color};
                border: 1px solid {COLORS['accent'] if not is_primary else 'transparent'};
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 10pt;
                outline: none;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
                color: {'white' if is_primary else COLORS['accent']};
                border: 1px solid {COLORS['accent']};
            }}
            QPushButton:pressed {{
                background-color: {hover_color};
                padding-top: 9px;
                padding-bottom: 7px;
            }}
            QPushButton:disabled {{
                background-color: #cccccc;
                color: #666666;
                border: 1px solid #bbbbbb;
            }}
        """)
        if icon_path:
            if os.path.exists(icon_path):
                self.setIcon(QIcon(icon_path))
                self.setIconSize(QSize(16, 16))
            else:
                logger = logging.getLogger(__name__)
                logger.warning(f"Icon not found: {icon_path}")

class WeatherIconWidget(QWidget):
    """Draws a simple weather icon based on description and temperature."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self.description = "unknown"
        self.temp = 0.0
        self.setMinimumSize(60, 60)
        self.setMaximumSize(60, 60)

    def update_weather(self, description, temp):
        self.logger.debug(f"Updating icon: Desc='{description}', Temp={temp}")
        self.description = str(description).lower() if description else "unknown"
        try:
            self.temp = float(temp) if temp is not None else 0.0
        except (TypeError, ValueError):
            self.logger.warning(f"Invalid temperature value received: {temp}. Using 0.0.")
            self.temp = 0.0
        self.update()

    def paintEvent(self, event):
        """Handles the drawing of the weather icon."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)
        draw_func_map = {
            'clear': self._draw_sun,
            'sun': self._draw_sun,
            'cloud': self._draw_cloud,
            'overcast': self._draw_cloud,
            'rain': self._draw_rain,
            'drizzle': self._draw_rain,
            'snow': self._draw_snow,
            'thunder': self._draw_thunder,
            'fog': self._draw_fog,
            'mist': self._draw_fog,
            'haze': self._draw_fog,
        }
        draw_func = self._draw_thermometer
        description_lower = self.description.lower()
        for key, func in draw_func_map.items():
            if key in description_lower:
                draw_func = func
                break
        try:
            draw_func(painter)
        except Exception as draw_err:
            self.logger.error(f"Error drawing icon part ({self.description}): {draw_err}", exc_info=True)
            painter.end()
            painter.begin(self)
            painter.setRenderHint(QPainter.Antialiasing)
            self._draw_thermometer(painter)
        painter.setPen(QPen(QColor(COLORS['text_primary']), 1))
        font = painter.font()
        font.setPointSize(10)
        painter.setFont(font)
        temp_text = f"{self.temp:.0f}°"
        text_rect = painter.fontMetrics().boundingRect(temp_text)
        painter.drawText(int((self.width() - text_rect.width()) / 2), 55, temp_text)

    def _draw_sun(self, painter):
        painter.setBrush(QColor(COLORS['warning']))
        painter.setPen(Qt.NoPen)
        radius = 15
        center_point = QPoint(self.width() // 2, self.height() // 2 - 5)
        painter.drawEllipse(center_point, radius, radius)
        painter.setPen(QPen(QColor(COLORS['warning']), 2))
        num_rays = 8
        for i in range(num_rays):
            angle = math.pi / (num_rays / 2) * i
            x1 = center_point.x() + radius * 1.2 * math.cos(angle)
            y1 = center_point.y() + radius * 1.2 * math.sin(angle)
            x2 = center_point.x() + radius * 1.7 * math.cos(angle)
            y2 = center_point.y() + radius * 1.7 * math.sin(angle)
            painter.drawLine(int(x1), int(y1), int(x2), int(y2))

    def _draw_cloud(self, painter):
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(180, 180, 180))
        painter.drawEllipse(10, 15, 25, 25)
        painter.drawEllipse(25, 10, 30, 30)
        painter.drawEllipse(40, 18, 20, 20)
        painter.drawRect(15, 25, 40, 15)

    def _draw_rain(self, painter):
        self._draw_cloud(painter)
        painter.setPen(QPen(QColor(COLORS['info']), 1))
        for i in range(3):
            start_x = 25 + i * 10
            painter.drawLine(start_x, 40, start_x - 5, 50)

    def _draw_snow(self, painter):
        self._draw_cloud(painter)
        painter.setPen(QPen(Qt.white, 1))
        font = painter.font()
        font.setPointSize(14)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(20, 48, "*")
        painter.drawText(35, 45, "*")
        painter.drawText(45, 50, "*")

    def _draw_thunder(self, painter):
        self._draw_cloud(painter)
        painter.setPen(QPen(QColor(COLORS['warning']), 2))
        painter.setBrush(QColor(COLORS['warning']))
        points = [QPoint(35, 40), QPoint(30, 45), QPoint(33, 45), QPoint(28, 52)]
        painter.drawPolygon(points)

    def _draw_fog(self, painter):
        painter.setPen(QPen(QColor(190, 190, 190), 2))
        for y in range(20, 45, 8):
            painter.drawLine(10, y, self.width() - 10, y)

    def _draw_thermometer(self, painter):
        painter.setBrush(Qt.NoBrush)
        painter.setPen(QPen(QColor(COLORS['text_secondary']), 1))
        bulb_radius = 6
        stem_width = 8
        stem_height = 25
        center_x = self.width() // 2
        bulb_bottom_y = 45
        bulb_center_y = bulb_bottom_y - bulb_radius
        stem_top_y = bulb_center_y - stem_height
        stem_x = center_x - stem_width // 2
        painter.drawRect(stem_x, stem_top_y, stem_width, stem_height)
        painter.drawEllipse(QPoint(center_x, bulb_center_y), bulb_radius, bulb_radius)

class TrendIndicator(QWidget):
    """Displays an up/down/stable arrow and optional percentage change."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self.trend = 0
        self.percentage = 0.0
        self.setMinimumSize(80, 25)
        self.setMaximumSize(100, 25)

    def set_trend(self, trend, percentage=None):
        self.logger.debug(f"Setting trend: Trend={trend}, Percentage={percentage}")
        try:
            if trend > 0:
                self.trend = 1
            elif trend < 0:
                self.trend = -1
            else:
                self.trend = 0
            self.percentage = float(percentage) if percentage is not None else None
        except (ValueError, TypeError):
            self.logger.warning(f"Invalid trend/percentage value: Trend={trend}, Percentage={percentage}. Resetting.")
            self.trend = 0
            self.percentage = None
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)
        arrow_width = 10
        arrow_height = 10
        arrow_x = 5
        arrow_y_center = self.height() // 2
        if self.trend > 0:
            color = QColor(COLORS['success'])
        elif self.trend < 0:
            color = QColor(COLORS['danger'])
        else:
            color = QColor(COLORS['text_secondary'])
        arrow_pen = QPen(color, 2)
        arrow_pen.setCapStyle(Qt.RoundCap)
        arrow_pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(arrow_pen)
        painter.setBrush(color)
        arrow_points = []
        if self.trend > 0:
            arrow_points = [
                QPoint(arrow_x, arrow_y_center + arrow_height // 2),
                QPoint(arrow_x + arrow_width, arrow_y_center + arrow_height // 2),
                QPoint(arrow_x + arrow_width // 2, arrow_y_center - arrow_height // 2)
            ]
        elif self.trend < 0:
            arrow_points = [
                QPoint(arrow_x, arrow_y_center - arrow_height // 2),
                QPoint(arrow_x + arrow_width, arrow_y_center - arrow_height // 2),
                QPoint(arrow_x + arrow_width // 2, arrow_y_center + arrow_height // 2)
            ]
        else:
            stable_pen = QPen(color, 2)
            stable_pen.setCapStyle(Qt.RoundCap)
            painter.setPen(stable_pen)
            painter.setBrush(Qt.NoBrush)
            painter.drawLine(arrow_x, self.height() // 2, arrow_x + arrow_width, self.height() // 2)
        if arrow_points:
            painter.drawPolygon(arrow_points)
        if self.percentage is not None and abs(self.percentage) > 0.001:
            text_pen = QPen(color, 1)
            painter.setPen(text_pen)
            font = painter.font()
            font.setPointSize(9)
            painter.setFont(font)
            sign = "+" if self.trend > 0 else ("-" if self.trend < 0 else "")
            text = f"{sign}{abs(self.percentage):.1f}%"
            text_x = arrow_x + arrow_width + 8
            text_y = arrow_y_center + painter.fontMetrics().ascent() // 2 - 1
            painter.drawText(text_x, text_y, text)

class CircularProgressGauge(QWidget):
    """A circular gauge widget to display progress or a value."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self._value = 0.0
        self._maximum = 100.0
        self._minimum = 0.0
        self._color = QColor(COLORS['primary'])
        self._text_format = "{:.1f}"
        self._text_suffix = ""
        gauge_size = 80
        self.setMinimumSize(gauge_size, gauge_size)
        self.setMaximumSize(gauge_size, gauge_size)

    def set_value(self, value):
        try:
            new_value = float(value) if value is not None else self._minimum
            self._value = max(self._minimum, min(new_value, self._maximum))
            self.update()
        except (TypeError, ValueError) as e:
            self.logger.error(f"Error setting gauge value: {e}. Input: {value}")
            self._value = self._minimum
            self.update()

    def set_range(self, minimum, maximum):
        self.logger.debug(f"Setting range: Min={minimum}, Max={maximum}")
        try:
            min_val = float(minimum)
            max_val = float(maximum)
            if max_val <= min_val:
                self.logger.warning(f"Gauge range max <= min ({max_val} <= {min_val}). Adjusting max.")
                max_val = min_val + 1
            self._minimum = min_val
            self._maximum = max_val
            self.set_value(self._value)
            self.logger.debug(f"Gauge range set: Min={self._minimum}, Max={self._maximum}")
        except (TypeError, ValueError) as e:
            self.logger.error(f"Error setting gauge range: {e}. Min: {minimum}, Max: {maximum}", exc_info=True)

    def set_color(self, color):
        new_color = QColor(COLORS['primary'])
        try:
            if isinstance(color, QColor):
                new_color = color
            elif isinstance(color, str):
                if color.startswith('#'):
                    new_color = QColor(color)
                elif color in COLORS:
                    new_color = QColor(COLORS[color])
                else:
                    self.logger.warning(f"Invalid color name/key: '{color}'. Using default.")
            else:
                self.logger.warning(f"Invalid color type: {type(color)}. Using default.")
            self._color = new_color
            self.update()
        except Exception as e:
            self.logger.error(f"Error setting color '{color}': {e}", exc_info=True)
            self._color = QColor(COLORS['primary'])
            self.update()

    def set_text_format(self, format_string="{:.1f}", suffix=""):
        self._text_format = format_string
        self._text_suffix = suffix
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)
        rect = self.rect().adjusted(5, 5, -5, -5)
        pen_width = 8
        start_angle = 90 * 16
        bg_pen = QPen(QColor(230, 230, 230), pen_width)
        bg_pen.setCapStyle(Qt.FlatCap)
        painter.setPen(bg_pen)
        painter.drawArc(rect, 0, 360 * 16)
        value_range = self._maximum - self._minimum
        span_angle = 0
        if value_range > 0:
            proportion = max(0.0, min((self._value - self._minimum) / value_range, 1.0))
            span_angle = int(-360 * proportion * 16)
        if span_angle != 0:
            progress_pen = QPen(self._color, pen_width)
            progress_pen.setCapStyle(Qt.RoundCap)
            painter.setPen(progress_pen)
            painter.drawArc(rect, start_angle, span_angle)
        painter.setPen(QPen(QColor(COLORS['text_primary']), 1))
        font = painter.font()
        font.setPointSize(12)
        font.setBold(True)
        painter.setFont(font)
        try:
            formatted_value = self._text_format.format(self._value)
            text = f"{formatted_value}{self._text_suffix}"
        except Exception as e:
            text = "Err"
            self.logger.warning(f"Text format error: Format='{self._text_format}', Suffix='{self._text_suffix}', Value={self._value}. Error: {e}")
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
    """The main widget for the Home screen, displaying various data."""
    def __init__(self, main_window=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self.logger.info("Initializing HomeModule...")
        self.signals = HomeSignals()
        self.fetcher_thread = None
        self.scheduled_refresh = None
        self.finnhub_client = None
        if finnhub and FINNHUB_API_KEY:
            try:
                self.finnhub_client = finnhub.Client(api_key=FINNHUB_API_KEY)
                self.logger.info("Finnhub client initialized.")
            except Exception as e:
                self.logger.error(f"Failed to initialize Finnhub client: {e}", exc_info=True)
                self.finnhub_client = None
        else:
            self.logger.warning("Finnhub library not imported or API key missing. Finance data unavailable.")
        self.weather_tools = None
        if WeatherTools:
            try:
                self.weather_tools = WeatherTools()
                self.logger.info("WeatherTools initialized.")
            except Exception as e:
                self.logger.error(f"Failed to initialize WeatherTools: {e}", exc_info=True)
                self.weather_tools = None
        else:
            self.logger.warning("WeatherTools not imported. Weather data unavailable.")
        self.weather_cities = WEATHER_CITIES
        self.logger.info(f"Weather cities configured: {self.weather_cities}")
        self._init_ui()
        self._connect_signals()
        self._load_initial_data_from_cache()
        self._start_scheduled_refresh()
        QTimer.singleShot(1000, lambda: self.refresh_data(all_data=True))
        self.logger.info("HomeModule initialization complete.")

    def _init_ui(self):
        self.logger.debug("Initializing UI...")
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(10, 10, 10, 10)
        self.main_layout.setSpacing(15)
        top_layout = QHBoxLayout()
        self.refresh_button = StyledButton("Refresh All", is_primary=False)
        self.status_label = QLabel("Initializing...")
        self.status_label.setStyleSheet("color: #666666; font-style: italic;")
        top_layout.addWidget(self.status_label)
        top_layout.addStretch(1)
        top_layout.addWidget(self.refresh_button)
        self.main_layout.addLayout(top_layout)
        self.grid_layout = QGridLayout()
        self.grid_layout.setSpacing(15)
        self.main_layout.addLayout(self.grid_layout)
        self.weather_card = StyledCard("Weather")
        self.weather_layout = QHBoxLayout()
        self.weather_card.add_content(self.weather_layout)
        self.grid_layout.addWidget(self.weather_card, 0, 0, 1, 2)
        self.weather_widgets = {}
        for city in self.weather_cities:
            city_widget = QWidget()
            city_layout = QVBoxLayout(city_widget)
            city_layout.setContentsMargins(10, 5, 10, 5)
            city_label = QLabel(city)
            city_label.setFont(QFont("Arial", 10, QFont.Bold))
            city_label.setAlignment(Qt.AlignCenter)
            city_layout.addWidget(city_label)
            weather_icon = WeatherIconWidget()
            city_layout.addWidget(weather_icon)
            self.weather_widgets[city] = {
                'container': city_widget,
                'label': city_label,
                'icon': weather_icon
            }
            self.weather_layout.addWidget(city_widget)
        self.exchange_card = StyledCard("Exchange Rate (USD/CAD)")
        exchange_content = QWidget()
        exchange_layout = QVBoxLayout(exchange_content)
        exchange_layout.setContentsMargins(5, 5, 5, 5)
        exchange_rate_layout = QHBoxLayout()
        self.exchange_label = QLabel("Fetching...")
        self.exchange_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 14pt; font-weight: bold;")
        exchange_rate_layout.addWidget(self.exchange_label)
        exchange_rate_layout.addStretch(1)
        gauge_trend_layout = QHBoxLayout()
        self.exchange_gauge = CircularProgressGauge()
        self.exchange_gauge.set_range(1.20, 1.50)
        self.exchange_trend = TrendIndicator()
        gauge_trend_layout.addWidget(self.exchange_gauge)
        gauge_trend_layout.addWidget(self.exchange_trend)
        exchange_layout.addLayout(exchange_rate_layout)
        exchange_layout.addLayout(gauge_trend_layout)
        self.exchange_card.add_content(exchange_content)
        self.grid_layout.addWidget(self.exchange_card, 1, 0)
        self.wheat_card = StyledCard("Wheat Price")
        wheat_content = QWidget()
        wheat_layout = QVBoxLayout(wheat_content)
        wheat_layout.setContentsMargins(5, 5, 5, 5)
        wheat_price_layout = QHBoxLayout()
        self.wheat_label = QLabel("Fetching...")
        self.wheat_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 14pt; font-weight: bold;")
        wheat_price_layout.addWidget(self.wheat_label)
        wheat_price_layout.addStretch(1)
        wheat_gauge_trend_layout = QHBoxLayout()
        self.wheat_gauge = CircularProgressGauge()
        self.wheat_gauge.set_range(3.00, 12.00)
        self.wheat_trend = TrendIndicator()
        wheat_gauge_trend_layout.addWidget(self.wheat_gauge)
        wheat_gauge_trend_layout.addWidget(self.wheat_trend)
        wheat_layout.addLayout(wheat_price_layout)
        wheat_layout.addLayout(wheat_gauge_trend_layout)
        self.wheat_card.add_content(wheat_content)
        self.grid_layout.addWidget(self.wheat_card, 2, 0)
        self.canola_card = StyledCard("Canola Price")
        canola_content = QWidget()
        canola_layout = QVBoxLayout(canola_content)
        canola_layout.setContentsMargins(5, 5, 5, 5)
        canola_price_layout = QHBoxLayout()
        self.canola_label = QLabel("Fetching...")
        self.canola_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 14pt; font-weight: bold;")
        canola_price_layout.addWidget(self.canola_label)
        canola_price_layout.addStretch(1)
        canola_gauge_trend_layout = QHBoxLayout()
        self.canola_gauge = CircularProgressGauge()
        self.canola_gauge.set_range(500, 1200)
        self.canola_trend = TrendIndicator()
        canola_gauge_trend_layout.addWidget(self.canola_gauge)
        canola_gauge_trend_layout.addWidget(self.canola_trend)
        canola_layout.addLayout(canola_price_layout)
        canola_layout.addLayout(canola_gauge_trend_layout)
        self.canola_card.add_content(canola_content)
        self.grid_layout.addWidget(self.canola_card, 1, 1)
        self.bitcoin_card = StyledCard("Bitcoin Price")
        bitcoin_content = QWidget()
        bitcoin_layout = QVBoxLayout(bitcoin_content)
        bitcoin_layout.setContentsMargins(5, 5, 5, 5)
        bitcoin_price_layout = QHBoxLayout()
        self.bitcoin_label = QLabel("Fetching...")
        self.bitcoin_label.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 14pt; font-weight: bold;")
        bitcoin_price_layout.addWidget(self.bitcoin_label)
        bitcoin_price_layout.addStretch(1)
        bitcoin_gauge_trend_layout = QHBoxLayout()
        self.bitcoin_gauge = CircularProgressGauge()
        self.bitcoin_gauge.set_range(10, 100)
        self.bitcoin_trend = TrendIndicator()
        bitcoin_gauge_trend_layout.addWidget(self.bitcoin_gauge)
        bitcoin_gauge_trend_layout.addWidget(self.bitcoin_trend)
        bitcoin_layout.addLayout(bitcoin_price_layout)
        bitcoin_layout.addLayout(bitcoin_gauge_trend_layout)
        self.bitcoin_card.add_content(bitcoin_content)
        self.grid_layout.addWidget(self.bitcoin_card, 2, 1)
        self.main_layout.addStretch(1)
        self.logger.debug("UI Initialization complete.")

    def _connect_signals(self):
        self.logger.debug("Connecting signals...")
        self.signals.weather_ready.connect(self._update_weather_ui)
        self.signals.exchange_ready.connect(self._update_exchange_ui)
        self.signals.wheat_ready.connect(self._update_wheat_ui)
        self.signals.canola_ready.connect(self._update_canola_ui)
        self.signals.bitcoin_ready.connect(self._update_bitcoin_ui)
        self.signals.status_update.connect(self._update_status_label)
        self.signals.refresh_complete.connect(self.on_refresh_complete)
        self.refresh_button.clicked.connect(lambda: self.refresh_data(all_data=True))
        if self.scheduled_refresh:
            self.scheduled_refresh.refresh_signal.connect(self.trigger_scheduled_refresh)
        self.logger.debug("Signal connections complete.")

    def _load_initial_data_from_cache(self):
        self.logger.info("Loading initial data from cache...")
        self._update_status_label("Loading cached data...")
        if not DataCache.is_cache_expired(WEATHER_CACHE, WEATHER_REFRESH_HOURS):
            weather_data = DataCache.load_from_cache(WEATHER_CACHE)
            if weather_data and 'results' in weather_data and 'timestamp' in weather_data:
                self.logger.debug("Using cached weather data.")
                self._update_weather_ui(weather_data['results'], weather_data['timestamp'])
            else:
                self.logger.warning("Invalid weather cache format.")
        else:
            self.logger.info("Weather cache expired or missing.")
        if not DataCache.is_cache_expired(EXCHANGE_CACHE, EXCHANGE_REFRESH_HOURS):
            exchange_data = DataCache.load_from_cache(EXCHANGE_CACHE)
            if exchange_data and 'rate_text' in exchange_data and 'timestamp' in exchange_data:
                rate_value = exchange_data.get('rate_value')
                self.logger.debug(f"Using cached exchange data. Value: {rate_value}")
                self._update_exchange_ui(exchange_data['rate_text'], exchange_data['timestamp'], rate_value)
            else:
                self.logger.warning("Invalid exchange cache format.")
        else:
            self.logger.info("Exchange cache expired or missing.")
        if not DataCache.is_cache_expired(COMMODITIES_CACHE, COMMODITIES_REFRESH_HOURS):
            commodities_data = DataCache.load_from_cache(COMMODITIES_CACHE)
            if commodities_data and 'timestamp' in commodities_data:
                ts = commodities_data['timestamp']
                self.logger.debug("Using cached commodities data.")
                if 'wheat' in commodities_data:
                    w_data = commodities_data['wheat']
                    self._update_wheat_ui(w_data.get('text', 'Cached'), ts, w_data.get('value'))
                if 'canola' in commodities_data:
                    c_data = commodities_data['canola']
                    self._update_canola_ui(c_data.get('text', 'Cached'), ts, c_data.get('value'))
                if 'bitcoin' in commodities_data:
                    b_data = commodities_data['bitcoin']
                    self._update_bitcoin_ui(b_data.get('text', 'Cached'), ts, b_data.get('value'))
            else:
                self.logger.warning("Invalid or incomplete commodities cache format.")
        else:
            self.logger.info("Commodities cache expired or missing.")
        self._update_status_label("Ready.")

    def _start_scheduled_refresh(self):
        if self.scheduled_refresh and self.scheduled_refresh.isRunning():
            self.logger.warning("Scheduled refresh thread already running.")
            return
        self.logger.info("Starting scheduled refresh thread...")
        self.scheduled_refresh = ScheduledRefreshThread()
        self.scheduled_refresh.refresh_signal.connect(self.trigger_scheduled_refresh)
        self.scheduled_refresh.start()
        self.logger.info("Scheduled refresh thread started.")

    async def _fetch_all_data(self, fetch_weather=True, fetch_exchange=True, fetch_commodities=True):
        print(f"DEBUG: Starting _fetch_all_data (W:{fetch_weather}, E:{fetch_exchange}, C:{fetch_commodities})")
        tasks = []
        task_names = []
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
            self.signals.refresh_complete.emit()
            return []
        print(f"DEBUG: Running {len(tasks)} tasks: {task_names}")
        results = await asyncio.gather(*tasks, return_exceptions=True)
        print("DEBUG: asyncio.gather finished.")
        for i, result in enumerate(results):
            task_name = task_names[i]
            if isinstance(result, Exception):
                print(f"ERROR: Task '{task_name}' failed in _fetch_all_data:")
                traceback.print_exception(type(result), result, result.__traceback__)
                error_msg = f"❌ {task_name}: Fetch Error ({type(result).__name__})"
                now_str = datetime.now().strftime("%H:%M:%S")
                if task_name == "Weather":
                    self.signals.weather_ready.emit([f"❌ {city}: Fetch Error" for city in self.weather_cities], now_str)
                elif task_name == "Exchange":
                    self.signals.exchange_ready.emit(error_msg, now_str, None)
                elif task_name == "Commodities":
                    self.signals.wheat_ready.emit("❌ Wheat: Fetch Error", now_str, None)
                    self.signals.canola_ready.emit("❌ Canola: Fetch Error", now_str, None)
                    self.signals.bitcoin_ready.emit("❌ Bitcoin: Fetch Error", now_str, None)
            else:
                print(f"DEBUG: Task '{task_name}' completed successfully.")
        print("DEBUG: Emitting refresh_complete signal.")
        self.signals.refresh_complete.emit()
        print("DEBUG: _fetch_all_data finished.")
        return results

    def _update_status_label(self, status_text):
        self.status_label.setText(status_text)

    def _update_weather_ui(self, results, timestamp):
        print(f"Updating weather UI with {len(results)} results. Timestamp: {timestamp}")
        found_cities = set()
        for result_str in results:
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
                        styled_text = f"<span style='font-weight:bold; color:{COLORS['accent']}'>{city_name}:</span> {result_str.split(':', 1)[1].strip()}"
                        widget_data['label'].setText(styled_text)
                        widget_data['label'].setToolTip(result_str)
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
                    widget_data['label'].setToolTip(error_msg)
                    widget_data['icon'].update_weather("unknown", 0)
                    found_cities.add(city_name)
            else:
                print(f"Warning: Could not parse weather result string: {result_str}")
        for city in self.weather_cities:
            if city not in found_cities and city in self.weather_widgets:
                print(f"Warning: No weather data received for {city}")
                self.weather_widgets[city]['label'].setText(f"{city}: No data")
                self.weather_widgets[city]['icon'].update_weather("unknown", 0)
        if hasattr(self, 'weather_card') and self.weather_card:
            self.weather_card.add_footer_text(f"Last updated: {timestamp}")
        else:
            print("Could not find self.weather_card to update timestamp")

    def _update_exchange_ui(self, rate_text, timestamp, rate_value=None):
        print(f"Updating Exchange UI: Value={rate_value}, Text='{rate_text}', Time={timestamp}")
        self._update_commodity_ui(
            label_widget=self.exchange_label,
            gauge_widget=self.exchange_gauge,
            trend_widget=self.exchange_trend,
            card_ref=getattr(self, 'exchange_card', None),
            prefix="USD/CAD",
            value_text=rate_text,
            timestamp=timestamp,
            price_value=rate_value,
            color_key='info',
            gauge_range=(1.20, 1.50),
            text_format="{:.4f}",
            trend_baseline=1.35,
            trend_sensitivity=1.0,
            currency_symbol="",
            unit=""
        )

    def _update_commodity_ui(self, label_widget, gauge_widget, trend_widget, card_ref,
                              prefix, value_text, timestamp, price_value=None,
                              color_key='success', gauge_range=(0, 1), text_format="{:.2f}",
                              trend_baseline=None, trend_sensitivity=1.0, currency_symbol="$", unit=""):
        print(f"Updating {prefix} UI: Value={price_value}, Text='{value_text}', Time={timestamp}")
        styled_text = value_text
        trend = 0
        trend_percent = 0.0
        tooltip = value_text
        current_color_key = color_key
        if value_text.startswith("✅"):
            current_color_key = color_key
        elif value_text.startswith("❌"):
            current_color_key = 'danger'
        elif value_text.startswith("ℹ️"):
            current_color_key = 'info'
        if price_value is not None and isinstance(price_value, (float, int)):
            try:
                if '{' in text_format and '}' in text_format:
                    formatted_price = text_format.format(price_value)
                else:
                    formatted_price = f"{currency_symbol}{price_value:{text_format.replace('$', '')}}"
                full_unit = f" {unit}" if unit else ""
                styled_text = f"<span style='font-size:14pt; font-weight:bold; color:{COLORS[current_color_key]};'>{prefix}: {formatted_price}{full_unit}</span>"
                tooltip = f"{prefix}: {formatted_price}{full_unit} (Raw: {price_value})"
                gauge_widget.set_range(gauge_range[0], gauge_range[1])
                gauge_widget.set_text_format(text_format, unit)
                gauge_widget.set_color(COLORS[current_color_key])
                gauge_widget.set_value(price_value)
                if trend_baseline is not None:
                    deviation = price_value - trend_baseline
                    threshold = abs(trend_baseline * 0.005)
                    if abs(deviation) > threshold:
                        trend = 1 if deviation > 0 else -1
                        if trend_baseline != 0:
                            trend_percent = (deviation / trend_baseline) * 100 * trend_sensitivity
                        else:
                            trend_percent = 100.0 * trend_sensitivity if trend == 1 else -100.0 * trend_sensitivity
            except Exception as e:
                print(f"Error updating {prefix} gauge/trend: {e}")
                styled_text = f"<span style='color:{COLORS['danger']}'>{prefix}: Display Error</span>"
                gauge_widget.set_value(gauge_widget._minimum)
                tooltip = f"{prefix}: Error displaying value {price_value}"
        else:
            if value_text.startswith("❌"):
                styled_text = f"<span style='color:{COLORS['danger']}'>{value_text}</span>"
            elif value_text.startswith("ℹ️"):
                styled_text = f"<span style='color:{COLORS['info']}; font-size:14pt; font-weight:bold;'>{value_text}</span>"
            else:
                styled_text = value_text
            gauge_widget.set_value(gauge_widget._minimum)
        label_widget.setText(styled_text)
        label_widget.setToolTip(tooltip)
        trend_widget.set_trend(trend, trend_percent)
        if card_ref:
            card_ref.add_footer_text(f"Last updated: {timestamp}")

    def _update_wheat_ui(self, wheat_text, timestamp, price_value=None):
        display_value = None
        gauge_range_dollars = (3.00, 12.00)
        baseline_dollars = 6.50
        if price_value is not None:
            try:
                display_value = float(price_value) / 100.0
            except (ValueError, TypeError):
                print(f"Wheat: Could not convert raw value {price_value} to dollars.")
                display_value = None
                wheat_text = "❌ Wheat: Invalid Value"
        self._update_commodity_ui(
            label_widget=self.wheat_label,
            gauge_widget=self.wheat_gauge,
            trend_widget=self.wheat_trend,
            card_ref=getattr(self, 'wheat_card', None),
            prefix="Wheat",
            value_text=wheat_text,
            timestamp=timestamp,
            price_value=display_value,
            color_key='success',
            gauge_range=gauge_range_dollars,
            text_format=".2f",
            currency_symbol="$",
            unit="/bu",
            trend_baseline=baseline_dollars,
            trend_sensitivity=0.2
        )

    def _update_canola_ui(self, canola_text, timestamp, price_value=None):
        card_title = "Canola Price"
        match = re.search(r'\(([^)]+)\)', canola_text)
        if match:
            card_title += f" {match.group(0)}"
        if hasattr(self, 'canola_card'):
            self.canola_card.title_label.setText(card_title)
        self._update_commodity_ui(
            label_widget=self.canola_label,
            gauge_widget=self.canola_gauge,
            trend_widget=self.canola_trend,
            card_ref=getattr(self, 'canola_card', None),
            prefix="Canola",
            value_text=canola_text,
            timestamp=timestamp,
            price_value=price_value,
            color_key='secondary',
            gauge_range=(500, 1200),
            text_format=".2f",
            currency_symbol="$",
            unit="CAD/t",
            trend_baseline=750,
            trend_sensitivity=0.25
        )

    def _update_bitcoin_ui(self, bitcoin_text, timestamp, price_value=None):
        gauge_value_k = None
        label_price_value = None
        if price_value is not None:
            try:
                label_price_value = float(price_value)
                gauge_value_k = label_price_value / 1000.0
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
            value_text=bitcoin_text,
            timestamp=timestamp,
            price_value=gauge_value_k,
            color_key='bitcoin',
            gauge_range=(10, 100),
            text_format="{:.1f}k",
            currency_symbol="$",
            unit="USD",
            trend_baseline=60,
            trend_sensitivity=1.0
        )
        if label_price_value is not None and bitcoin_text.startswith("✅"):
            try:
                currency = "USD"
                match = re.search(r'\$([\d,]+\.\d{2})\s*([A-Z]{3})', bitcoin_text)
                if match:
                    currency = match.group(2)
                self.bitcoin_label.setText(f"<span style='font-size:14pt; font-weight:bold; color:{COLORS['bitcoin']};'>Bitcoin: ${label_price_value:,.2f} {currency}</span>")
                self.bitcoin_label.setToolTip(f"Bitcoin: ${label_price_value:,.2f} {currency}")
            except Exception as e:
                print(f"Error formatting bitcoin label: {e}")

    def on_refresh_complete(self):
        self.logger.info("Refresh cycle complete.")
        self.refresh_button.setEnabled(True)
        self._update_status_label("Ready.")

    def trigger_scheduled_refresh(self, fetch_weather, fetch_exchange, fetch_commodities):
        self.logger.info(f"Scheduler triggered refresh: W={fetch_weather}, E={fetch_exchange}, C={fetch_commodities}")
        self.refresh_data(fetch_weather, fetch_exchange, fetch_commodities)

    def refresh_data(self, fetch_weather=False, fetch_exchange=False, fetch_commodities=False, all_data=False):
        if all_data:
            fetch_weather = fetch_exchange = fetch_commodities = True
        if not (fetch_weather or fetch_exchange or fetch_commodities):
            self.logger.warning("refresh_data called with no data types selected.")
            return
        msg_parts = []
        if fetch_weather: msg_parts.append("Weather")
        if fetch_exchange: msg_parts.append("Exchange")
        if fetch_commodities: msg_parts.append("Commodities")
        msg = f"Refreshing {', '.join(msg_parts)}..."
        self._update_status_label(msg)
        self.refresh_button.setEnabled(False)
        if self.fetcher_thread and self.fetcher_thread.isRunning():
            self.logger.info("Stopping running fetch thread before starting new one.")
            self.fetcher_thread.stop()
            if self.fetcher_thread.wait(3000):
                self.logger.info("Previous fetcher thread stopped successfully.")
            else:
                self.logger.warning("Previous fetcher thread did not stop in time. Starting new one anyway.")
        self.fetcher_thread = DataFetcherThread(self, fetch_weather, fetch_exchange, fetch_commodities)
        self.fetcher_thread.start()
        self.logger.info(f"Started data fetcher thread (W:{fetch_weather}, E:{fetch_exchange}, C:{fetch_commodities})")

    async def fetch_weather(self):
        self.logger.info(f"Fetching weather for cities: {self.weather_cities}")
        results = []
        now_str = datetime.now().strftime("%H:%M:%S")
        if not self.weather_tools:
            self.logger.error("Weather tools not available - fetch failed")
            error_results = [f"❌ {city}: Weather API tools not available" for city in self.weather_cities]
            self.signals.weather_ready.emit(error_results, now_str)
            return error_results
        for city in self.weather_cities:
            try:
                simulated_result = f"✅ {city}: 20°C, Clear skies"
                results.append(simulated_result)
            except Exception as e:
                self.logger.error(f"Error fetching weather for {city}: {e}")
                results.append(f"❌ {city}: Fetch Error")
        self.signals.weather_ready.emit(results, now_str)
        return results

    async def fetch_exchange_rate(self):
        now_str = datetime.now().strftime("%H:%M:%S")
        try:
            simulated_rate = 1.3456
            result_text = f"✅ Rate: {simulated_rate}"
            self.signals.exchange_ready.emit(result_text, now_str, simulated_rate)
            return simulated_rate
        except Exception as e:
            self.logger.error(f"Error fetching exchange rate: {e}", exc_info=True)
            error_text = "❌ USD/CAD: Fetch Error"
            self.signals.exchange_ready.emit(error_text, now_str, None)
            return None

    async def fetch_commodities(self):
        now_str = datetime.now().strftime("%H:%M:%S")
        results = {}
        try:
            simulated_wheat = 650
            simulated_canola = 800
            simulated_bitcoin = 60000
            wheat_text = f"✅ Wheat: ${simulated_wheat / 100:.2f}/bu"
            canola_text = f"✅ Canola: ${simulated_canola:.2f} (RS=F)"
            bitcoin_text = f"✅ Bitcoin: ${simulated_bitcoin:,.2f} USD"
            self.signals.wheat_ready.emit(wheat_text, now_str, simulated_wheat)
            self.signals.canola_ready.emit(canola_text, now_str, simulated_canola)
            self.signals.bitcoin_ready.emit(bitcoin_text, now_str, simulated_bitcoin)
            results = {"wheat": simulated_wheat, "canola": simulated_canola, "bitcoin": simulated_bitcoin}
        except Exception as e:
            self.logger.error(f"Error fetching commodities: {e}", exc_info=True)
            self.signals.wheat_ready.emit("❌ Wheat: Fetch Error", now_str, None)
            self.signals.canola_ready.emit("❌ Canola: Fetch Error", now_str, None)
            self.signals.bitcoin_ready.emit("❌ Bitcoin: Fetch Error", now_str, None)
        return results
