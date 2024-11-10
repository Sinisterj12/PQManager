import json
import os
import logging

class ConfigManager:
    def __init__(self):
        self.config_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config')
        self.config_file = os.path.join(self.config_dir, 'settings.json')
        self.default_settings = {
            'selected_printer': '',
            'monitoring_interval': 15000,  # 15 seconds in milliseconds
            'start_minimized': False,
            'auto_start_monitoring': True
        }
        self._ensure_config_exists()

    def _ensure_config_exists(self):
        """Create config directory and file if they don't exist"""
        os.makedirs(self.config_dir, exist_ok=True)
        if not os.path.exists(self.config_file):
            self.save_settings(self.default_settings)

    def load_settings(self):
        """Load settings from config file"""
        try:
            with open(self.config_file, 'r') as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"Error loading settings: {e}")
            return self.default_settings

    def save_settings(self, settings):
        """Save settings to config file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(settings, f, indent=4)
            return True
        except Exception as e:
            logging.error(f"Error saving settings: {e}")
            return False

    def update_setting(self, key, value):
        """Update a single setting"""
        settings = self.load_settings()
        settings[key] = value
        return self.save_settings(settings)
