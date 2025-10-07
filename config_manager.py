import json
import os
import shutil

class ConfigManager:
    def __init__(self, configs_dir="configs", default_file="default.json"):
        self.configs_dir = configs_dir
        os.makedirs(self.configs_dir, exist_ok=True)
        self.default_path = os.path.join(self.configs_dir, default_file)
        if not os.path.exists(self.default_path):
            self.save_config({"headers": []}, self.default_path)
        self.config = self.load_config(self.default_path)

    def list_configs(self):
        return [f for f in os.listdir(self.configs_dir) if f.endswith(".json")]

    def load_config(self, path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def save_config(self, config, path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        self.config = config

    def rename_config(self, old_name, new_name):
        old_path = os.path.join(self.configs_dir, old_name)
        new_path = os.path.join(self.configs_dir, new_name)
        if os.path.exists(old_path):
            os.rename(old_path, new_path)

    def import_config(self, import_path):
        filename = os.path.basename(import_path)
        dest_path = os.path.join(self.configs_dir, filename)
        shutil.copy(import_path, dest_path)

    def export_config(self, filename, export_path):
        src_path = os.path.join(self.configs_dir, filename)
        shutil.copy(src_path, export_path)