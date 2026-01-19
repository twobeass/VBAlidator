import json
import sys
import os
import glob
try:
    import comtypes.client
    import comtypes.typeinfo
except ImportError:
    print("Error: 'comtypes' library is required.")
    print("Please install it: pip install comtypes")
    sys.exit(1)

def generate_model(json_path):
    print(f"Processing: {json_path}")
    
    with open(json_path, 'r') as f:
        data = json.load(f)
    
    if "references" not in data:
        print("No 'references' section found in JSON.")
        return

    if "classes" not in data:
        data["classes"] = {}

    for ref in data["references"]:
        name = ref.get("name")
        full_path = ref.get("fullpath")
        
        if not full_path or not os.path.exists(full_path):
            print(f"Skipping {name}: Path not found ({full_path})")
            continue
            
        print(f"Inspecting {name} ({full_path})...")
        
        try:
            # Load TypeLib
            tlib = comtypes.client.GetModule(full_path)
            
            # Since GetModule returns the python module wrapping it, we need the raw TypeLib inspection
            # Easier approach: Use LoadTypeLibEx
            tlib_obj = comtypes.typeinfo.LoadTypeLibEx(full_path)
            
            count = tlib_obj.GetTypeInfoCount()
            for i in range(count):
                ti_kind = tlib_obj.GetTypeInfoType(i)
                
                # We care about Interfaces, CoClasses, Modules, and Enums
                if ti_kind in (comtypes.typeinfo.TKIND_DISPATCH, comtypes.typeinfo.TKIND_INTERFACE, 
                             comtypes.typeinfo.TKIND_COCLASS, comtypes.typeinfo.TKIND_MODULE,
                             comtypes.typeinfo.TKIND_ENUM):
                    
                    ti_name = tlib_obj.GetDocumentation(i)[0]
                    ti = tlib_obj.GetTypeInfo(i)
                    attr = ti.GetTypeAttr()
                    
                    # Decide if this should be added to globals
                    # VBA.Global, VBA.Interaction, VBA.Strings, etc. are modules where members are global
                    # Also _HiddenInterface (found in VBA lib) contains String functions like Left/Right
                    is_global_module = (ti_kind == comtypes.typeinfo.TKIND_MODULE) or \
                                     (ti_name in ("Global", "_HiddenInterface"))
                    
                    if not is_global_module:
                        if ti_name not in data["classes"]:
                            data["classes"][ti_name] = {"members": {}}
                    
                    # Iterate vars/constants
                    for v_idx in range(attr.cVars):
                        desc = ti.GetVarDesc(v_idx)
                        names = ti.GetNames(desc.memid)
                        var_name = names[0]
                        
                        # Get value for constants/enums
                        val_type = "Variant"
                        if desc.varkind == comtypes.typeinfo.VAR_CONST:
                             # It's a constant/enum value
                             val_type = "Integer" # Simplified
                        
                        member_def = {"type": val_type}
                        
                        if is_global_module or ti_kind == comtypes.typeinfo.TKIND_ENUM:
                            # Promote to globals (e.g. vbCrLf, vbYes)
                            data["globals"][var_name] = member_def
                        else:
                            data["classes"][ti_name]["members"][var_name] = member_def
                        
                    # Iterate functions/methods
                    for f_idx in range(attr.cFuncs):
                        desc = ti.GetFuncDesc(f_idx)
                        names = ti.GetNames(desc.memid)
                        func_name = names[0]
                        
                        # Try to get return type name
                        ret_type = "Variant"
                        try:
                            # elemdescFunc.tdesc.vt usually
                            if desc.elemdescFunc.tdesc.vt == 26: # VT_PTR -> Object
                                # Try to resolve the pointed type
                                href = desc.elemdescFunc.tdesc.u.lptdesc.contents.u.hreftype
                                ref_ti = ti.GetRefTypeInfo(href)
                                ret_type = ref_ti.GetDocumentation(-1)[0]
                            elif desc.elemdescFunc.tdesc.vt == 8: # BSTR
                                ret_type = "String"
                            elif desc.elemdescFunc.tdesc.vt in (2, 3, 16, 17, 18, 19): # Integers
                                ret_type = "Integer"
                            elif desc.elemdescFunc.tdesc.vt == 11: # Boolean
                                ret_type = "Boolean"
                        except:
                            pass
                            
                        member_def = {"type": ret_type}
                        
                        if is_global_module:
                            data["globals"][func_name] = member_def
                        else:
                            data["classes"][ti_name]["members"][func_name] = member_def
        except Exception as e:
            print(f"Failed to process {name}: {e}")

    # Manually backfill ThisDocument if missing (Visio specific)
    # IVDocument has the Path property, not the generic Document
    data["globals"]["ThisDocument"] = {"type": "IVDocument"}
    
    # Patch VBE on Application
    # IVApplication usually misses this property in the TypeLib or it's named strangely
    if "IVApplication" in data["classes"]:
         data["classes"]["IVApplication"]["members"]["VBE"] = {"type": "VBE"}
    if "Application" in data["classes"]:
         data["classes"]["Application"]["members"]["VBE"] = {"type": "VBE"}
         
    # Ensure VBE type exists (minimal stub if missing)
    if "VBE" not in data["classes"]:
        data["classes"]["VBE"] = {
            "members": {
                "ActiveVBProject": {"type": "VBProject"} 
            }
        }
        
    # Ensure VBProject type exists
    if "VBProject" not in data["classes"]:
        data["classes"]["VBProject"] = {
            "members": {
               "References": {"type": "References"}
            }
        }

    # Write back
    with open(json_path, 'w') as f:
        json.dump(data, f, indent=2)
    print("Model updated successfully.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_typelib_json.py <path_to_vba_model.json>")
        sys.exit(1)
    
    generate_model(sys.argv[1])
