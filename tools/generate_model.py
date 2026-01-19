import json
import os
import sys
import traceback

try:
    import comtypes.client
    import comtypes
except ImportError:
    print("Error: 'comtypes' library is required. Install it using: pip install comtypes")
    sys.exit(1)

def generate_model():
    base_path = os.getcwd() 
    ref_file = "vba_references.json"
    
    if not os.path.exists(ref_file):
        possible_paths = [
            os.path.join(base_path, ref_file),
            os.path.join(base_path, "tools", ref_file),
            os.path.join(os.path.dirname(base_path), ref_file),
        ]
        found = False
        for p in possible_paths:
            if os.path.exists(p):
                ref_file = p
                found = True
                break
        if not found:
            print(f"Error: {ref_file} not found. Did you run the VBA Exporter?")
            sys.exit(1)

    print(f"Loading references from {ref_file}...")
    with open(ref_file, 'r') as f:
        ref_data = json.load(f)

    model = {
        "globals": {},
        "classes": {},
        "enums": {},
        "references": ref_data.get("references", [])
    }

    # PASS 1: Process Interfaces and Enums
    for ref in model["references"]:
        name = ref.get("name", "Unknown")
        guid = ref.get("guid")
        major = ref.get("major")
        minor = ref.get("minor")
        path = ref.get("path", "")
        
        print(f"Processing Library (Pass 1 - Interfaces/Enums): {name}...")
        
        try:
            mod = None
            try:
                mod = comtypes.client.GetModule((guid, major, minor))
            except Exception as e:
                if path and os.path.exists(path):
                    mod = comtypes.client.GetModule(path)
                else:
                    print(f"  Error: Could not load {name}")
                    continue
            
            if not mod: continue

            for attr_name in dir(mod):
                try:
                    attr = getattr(mod, attr_name)
                    
                    # 1. Interfaces / Dispatch Interfaces
                    if isinstance(attr, type) and (hasattr(attr, '_methods_') or hasattr(attr, '_disp_methods_')):
                        type_name = str(attr_name)
                        if type_name not in model["classes"]:
                            model["classes"][type_name] = { "type": "Class", "members": {} }
                        
                        methods_list = getattr(attr, '_methods_', [])
                        disp_methods_list = getattr(attr, '_disp_methods_', [])
                        all_methods = methods_list + disp_methods_list
                        
                        for m in all_methods:
                            m_name = None
                            if len(m) >= 3:
                                candidate = m[2]
                                if isinstance(candidate, str):
                                    m_name = candidate
                                elif isinstance(candidate, tuple) and len(candidate) > 0 and isinstance(candidate[0], str):
                                    m_name = candidate[0]
                            if not m_name and len(m) >= 2 and isinstance(m[1], str):
                                m_name = m[1]
                            
                            if m_name:
                                clean_name = str(m_name)
                                if clean_name.startswith("_get_"): clean_name = clean_name[5:]
                                elif clean_name.startswith("_set_"): clean_name = clean_name[5:]
                                elif clean_name.startswith("_put_"): clean_name = clean_name[5:]
                                    
                                model["classes"][type_name]["members"][clean_name] = {"type": "Variant"}

                                # Global Promotion
                                if name == "VBA" and not clean_name.startswith("_"):
                                    model["globals"][clean_name] = {"type": "Variant"}
                    
                    # 2. Enums (Heuristic: Classes with _case_insensitive_ or just integers)
                    # Comtypes often generates Enums as classes with class attributes for values
                    if isinstance(attr, type) and not hasattr(attr, '_methods_') and not hasattr(attr, '_disp_methods_') and not hasattr(attr, '_reg_clsid_'):
                         # Inspect attributes for integer values
                        enum_name = str(attr_name)
                        enum_values = {}
                        has_enum_values = False
                        for e_name in dir(attr):
                            if e_name.startswith("_"): continue
                            val = getattr(attr, e_name)
                            if isinstance(val, int):
                                enum_values[e_name] = val
                                has_enum_values = True
                        
                        if has_enum_values:
                             model["enums"][enum_name] = enum_values
                             # Also promote enum members to globals (implicit in VBA)
                             # Or store them in model["enums"] structure where keys are global
                             # For now, let's just make sure they are accessible via resolve_enum
                             pass

                except Exception:
                     pass
        except Exception:
             pass

    # PASS 2: Process CoClasses (Map to Interfaces)
    for ref in model["references"]:
        name = ref.get("name", "Unknown")
        guid = ref.get("guid")
        major = ref.get("major")
        minor = ref.get("minor")
        path = ref.get("path", "")

        print(f"Processing Library (Pass 2 - CoClasses): {name}...")
        try:
             # Fast reload/access
            mod = None
            try: mod = comtypes.client.GetModule((guid, major, minor))
            except: 
                if path and os.path.exists(path): mod = comtypes.client.GetModule(path)
            
            if not mod: continue

            for attr_name in dir(mod):
                try:
                    attr = getattr(mod, attr_name)
                     # Check for CoClass
                    if hasattr(attr, '_reg_clsid_'):
                         coclass_name = str(attr_name)
                         if coclass_name not in model["classes"]:
                             model["classes"][coclass_name] = { "type": "Class", "members": {} }
                         
                         # Find default interface
                         if hasattr(attr, '_com_interfaces_') and len(attr._com_interfaces_) > 0:
                             default_intf = attr._com_interfaces_[0]
                             intf_name = default_intf.__name__ # e.g. IVShape
                             
                             # Copy members
                             if intf_name in model["classes"]:
                                 src_members = model["classes"][intf_name]["members"]
                                 print(f"    Copying {len(src_members)} members from {intf_name} to {coclass_name}")
                                 for m_k, m_v in src_members.items():
                                     model["classes"][coclass_name]["members"][m_k] = m_v
                             else:
                                 print(f"    Warning: Interface {intf_name} not found for CoClass {coclass_name}")
                         else:
                             print(f"    Warning: No default interface found for CoClass {coclass_name}")
                                     
                except Exception as ex:
                    print(f"    Error processing CoClass {attr_name}: {ex}")
        except Exception:
            pass

    # Add standard globals
    model["globals"]["Visio"] = {"type": "IVApplication"}

    # Hardcoded Standard VBA Globals (Fallback)
    standard_globals = [
        "InStr", "Left", "Right", "Mid", "Len", "LenB", "LTrim", "RTrim", "Trim", 
        "UCase", "LCase", "Space", "String", "Format", "Replace", "Split", "Join",
        "MsgBox", "InputBox", "Shell", "DoEvents", "CreateObject", "GetObject",
        "CurDir", "Dir", "MkDir", "RmDir", "ChDir", "ChDrive", "Kill", "FileCopy", "Name", "FileLen", "GetAttr", "SetAttr",
        "FileDateTime", "FreeFile", "Open", "Close", "Print", "Write", "Input", "Line", "Loc", "LOF", "Seek", "EOF",
        "Sin", "Cos", "Tan", "Atn", "Sqr", "Exp", "Log", "Abs", "Sgn", "Fix", "Int", 
        "Rnd", "Randomize", "Timer", "Time", "Date", "Now", "Day", "Month", "Year", 
        "Hour", "Minute", "Second", "DateSerial", "TimeSerial", "DateValue", "TimeValue", "DateAdd", "DateDiff", "DatePart",
        "IsNumeric", "IsDate", "IsEmpty", "IsNull", "IsArray", "IsObject", "IsError", "IsMissing",
        "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng", "CLngLng", "CLngPtr", "CSng", "CStr", "CVar",
        "CVErr", "Val", "Str", "Hex", "Oct", "RGB", "QBColor", "VarType", "TypeName", "IIf", "Switch", "Choose", "Partition",
        "LBound", "UBound", "Array", "Filter", "Error", "Err",
        "Chr", "ChrW", "Asc", "AscW", "Environ", "InStrRev", "StrComp", "Round"
    ]


    for g in standard_globals:
        if g not in model["globals"]:
             model["globals"][g] = {"type": "Variant"}
    
    # Manual Fixes for Common VBA Constants
    common_consts = ["vbCrLf", "vbNullString", "vbTrue", "vbFalse", "vbRed", "vbGreen", "vbBlue", 
                     "vbBlack", "vbWhite", "vbInformation", "vbExclamation", "vbCritical", "vbYesNo", "vbYes", "vbNo", "vbOKOnly", 
                     "vbObjectError", "vbNullChar", "vbModeless", "vbCr", "vbLf", "vbTab", "vbBack", "vbFormFeed", "vbVerticalTab",
                     "vbNewLine", "vbDate", "vbBoolean", "vbByte", "vbCurrency", "vbDecimal", "vbDouble", "vbEmpty", "vbError", "vbInteger", "vbLong", "vbNull", "vbObject", "vbSingle", "vbString", "vbVariant",
                     "visServiceVersion140"]
    for c in common_consts:
        model["globals"][c] = {"type": "Long"}
        
    # Manual Fixes for Visio Constants (that might be missed or tricky Enums)
    visio_consts = ["visCustPropsAsk", "visCustPropsLangID", "visCustPropsCalendar", "visUserValue", "visUserPrompt", 
                    "visCenterViewDefault", "visSectionHyperlink", "visSectionUser", "visTagDefault", 
                    "visTagCnnctPt", "visTagCnnctPtABCD", "visTagCnnctNamed", "visTagCnnctNamedABCD",
                    "visSectionProp", "visCustPropsLabel", "visCustPropsPrompt", "visCustPropsType", 
                    "visCustPropsFormat", "visCustPropsValue", "visCustPropsSortKey", "visCustPropsInvis"]
    for c in visio_consts:
        model["globals"][c] = {"type": "Long"}

    # Promote IVApplication members to Global Scope (e.g. ActiveWindow, ActiveDocument)
    if "IVApplication" in model["classes"]:
        app_members = model["classes"]["IVApplication"]["members"]
        print(f"Promoting {len(app_members)} Application members to Global scope...")
        for m_name, m_def in app_members.items():
             if m_name not in model["globals"]:
                 model["globals"][m_name] = m_def
    
    # Ensure UserForm has Show/Hide and standard properties
    if "UserForm" in model["classes"]:
        uf_members = model["classes"]["UserForm"]["members"]
        common_uf_members = {
            "Show": "Sub", "Hide": "Sub", "Controls": "Object", 
            "Width": "Double", "Height": "Double", "Top": "Double", "Left": "Double",
            "ScrollHeight": "Double", "ScrollWidth": "Double",
            "InsideHeight": "Double", "InsideWidth": "Double"
        }
        for m, t in common_uf_members.items():
             if m not in uf_members:
                 uf_members[m] = {"type": t}

    # Manual Fixes for Excel constants used in repo
    xl_consts = ["xlValidateDecimal", "xlBetween", "xlUp", "xlToLeft"]
    for c in xl_consts:
        model["globals"][c] = {"type": "Long"}
    
    # Alias ThisDocument to Document (or IVShape for some items, but Document is best for top level)
    if "Document" in model["classes"]:
        model["globals"]["ThisDocument"] = {"type": "Document"}
    elif "IVDocument" in model["classes"]:
        model["globals"]["ThisDocument"] = {"type": "IVDocument"}

    # Help Document/IVDocument with missing but critical properties
    for doc_cls in ["Document", "IVDocument"]:
        if doc_cls in model["classes"]:
            doc_members = model["classes"][doc_cls]["members"]
            missing_doc = {"Path": "String", "VBProject": "Object", "CustomUI": "String", "Fullname": "String", "Name": "String"}
            for m, t in missing_doc.items():
                if m not in doc_members:
                    doc_members[m] = {"type": t}

    # Fix MSForms Controls (they often need IControl members)
    if "IControl" in model["classes"]:
        control_members = model["classes"]["IControl"]["members"]
        for cls in ["CommandButton", "TextBox", "Label", "ListBox", "ComboBox", "CheckBox", "OptionButton", "Frame", "Image"]:
            if cls in model["classes"]:
                for m_name, m_def in control_members.items():
                    if m_name not in model["classes"][cls]["members"]:
                        model["classes"][cls]["members"][m_name] = m_def


    output_file = "vba_model.json"
    print(f"Saving model to {output_file}...")
    try:
        with open(output_file, 'w') as f:
            json.dump(model, f, indent=2)
        print("Done.")
    except Exception as e:
        print(f"Error saving JSON: {e}")

if __name__ == "__main__":
    generate_model()
