import json
import os

class Config:
    def __init__(self):
        self.definitions = {}  # For Conditional Compilation (e.g., Win64)
        self.object_model = {
            "globals": {},
            "classes": {},
            "enums": {}
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
            try:
                self.load_model(std_model_path)
            except Exception as e:
                print(f"Warning: Failed to load standard model at {std_model_path}: {e}")
        else:
            print(f"Warning: Standard model not found at {std_model_path}")

    def load_model(self, filepath):
        """Loads an external JSON object model and merges it."""
        with open(filepath, 'r') as f:
            data = json.load(f)

        if not isinstance(data, dict):
            raise ValueError("Model must be a JSON object.")

        valid_sections = {"globals", "classes", "enums", "references"}
        if not any(k in data for k in valid_sections):
            raise ValueError(f"Model file must contain at least one of the following sections: {', '.join(valid_sections)}")

        # Merge Globals (Key Normalization)
        if "globals" in data:
            for name, defn in data["globals"].items():
                self.object_model["globals"][name.lower()] = defn

        # Merge Classes (Key Normalization)
        if "classes" in data:
            for cls_name, cls_def in data["classes"].items():
                lower_name = cls_name.lower()
                if lower_name in self.object_model["classes"]:
                    # Merge members
                    if "members" in cls_def:
                        existing_members = self.object_model["classes"][lower_name].get("members", {})
                        existing_members.update(cls_def["members"])
                        self.object_model["classes"][lower_name]["members"] = existing_members
                else:
                    self.object_model["classes"][lower_name] = cls_def

        # Merge References
        if "references" in data:
            if "references" not in self.object_model:
                    self.object_model["references"] = []
            existing_names = {r["name"] for r in self.object_model["references"]}
            for ref in data["references"]:
                if ref["name"] not in existing_names:
                    self.object_model["references"].append(ref)
                    existing_names.add(ref["name"])

        # Merge Enums (Key Normalization)
        if "enums" in data:
            for enum_name, members in data["enums"].items():
                self.object_model["enums"][enum_name.lower()] = members

    def get_global(self, name):
        return self.object_model["globals"].get(name.lower())

    def get_class(self, name):
        return self.object_model["classes"].get(name.lower())
