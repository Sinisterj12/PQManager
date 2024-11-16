import json
import os
import logging
import sys
from pathlib import Path

class ConfigManager:
    def __init__(self):
        # Use proper app data directory for settings
        if getattr(sys, 'frozen', False):
            # Running as compiled exe
            app_data = os.path.join(os.environ.get('APPDATA', ''), 'PQManager')
        else:
            # Running in development
            app_data = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config')
            
        self.config_dir = app_data
        self.config_file = os.path.join(self.config_dir, 'settings.json')
        self.default_settings = {
            'selected_printer': '',
            'monitoring_interval': 10000,  # 10 seconds
            'start_minimized': False,
            'auto_start_monitoring': True
        }
        self._ensure_config_exists()

    def _ensure_config_exists(self):
        """Create config directory and file if they don't exist"""
        try:
            os.makedirs(self.config_dir, exist_ok=True)
            if not os.path.exists(self.config_file):
                self.save_settings(self.default_settings)
                logging.info(f"Created new settings file at {self.config_file}")
            else:
                logging.info(f"Using existing settings at {self.config_file}")
        except Exception as e:
            logging.error(f"Error ensuring config exists: {e}")

    def load_settings(self):
        """Load settings from config file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    settings = json.load(f)
                    logging.info(f"Loaded settings: {settings}")
                    return settings
        except Exception as e:
            logging.error(f"Error loading settings: {e}")
        
        logging.info("Using default settings")
        return self.default_settings.copy()

    def save_settings(self, settings):
        """Save settings to config file"""
        try:
            # Ensure directory exists
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            
            with open(self.config_file, 'w') as f:
                json.dump(settings, f, indent=4)
            logging.info(f"Saved settings: {settings}")
            return True
        except Exception as e:
            logging.error(f"Error saving settings: {e}")
            return False

    def update_setting(self, key, value):
        """Update a single setting"""
        settings = self.load_settings()
        settings[key] = value
        return self.save_settings(settings)
