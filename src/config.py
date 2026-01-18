import json
import os

class Config:
    def __init__(self):
        self.definitions = {}  # For Conditional Compilation (e.g., Win64)
        self.object_model = {
            "globals": {},
            "classes": {}
        }
        self.load_standard_model()

    def parse_defines(self, define_str):
        """Parses a string like 'WIN64=True,VBA7=True' into the definitions dict."""
        if not define_str:
            return
        
        pairs = define_str.split(',')
        for pair in pairs:
            if '=' in pair:
                key, value = pair.split('=', 1)
                # Parse boolean values
                if value.lower() == 'true':
                    self.definitions[key.strip().upper()] = True
                elif value.lower() == 'false':
                    self.definitions[key.strip().upper()] = False
                else:
                    self.definitions[key.strip().upper()] = value.strip()
            else:
                # Assume True if no value provided
                self.definitions[pair.strip().upper()] = True

    def load_standard_model(self):
        """Loads the built-in standard model."""
        base_path = os.path.dirname(os.path.abspath(__file__))
        std_model_path = os.path.join(base_path, 'std_model.json')
        if os.path.exists(std_model_path):
            self.load_model(std_model_path)
        else:
            print(f"Warning: Standard model not found at {std_model_path}")

    def load_model(self, filepath):
        """Loads an external JSON object model and merges it."""
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
            
            # Merge Globals
            if "globals" in data:
                self.object_model["globals"].update(data["globals"])
            
            # Merge Classes
            if "classes" in data:
                # Deep merge for classes is better, but simple update for now
                for cls_name, cls_def in data["classes"].items():
                    if cls_name in self.object_model["classes"]:
                        # Merge members
                        if "members" in cls_def:
                            existing_members = self.object_model["classes"][cls_name].get("members", {})
                            existing_members.update(cls_def["members"])
                            self.object_model["classes"][cls_name]["members"] = existing_members
                    else:
                        self.object_model["classes"][cls_name] = cls_def
                        
        except Exception as e:
            print(f"Error loading model {filepath}: {e}")

    def get_global(self, name):
        return self.object_model["globals"].get(name)

    def get_class(self, name):
        return self.object_model["classes"].get(name)
