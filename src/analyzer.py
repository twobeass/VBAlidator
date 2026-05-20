from .parser import (  # noqa: F401
    DoNode,
    EraseNode,
    ForNode,
    IfNode,
    ProcedureNode,
    RedimNode,
    SelectNode,
    StatementNode,
    WithNode,
)

def _normalize_identifier(name):
    """Strip VBA legacy type-suffix and bracket-quoting from an identifier.

    - `Mid$` → `mid` (String)
    - `x%`   → `x`   (Integer)
    - `c@`   → `c`   (Currency)
    - `[A1]` → `a1`  (foreign-name escape)
    """
    if not name:
        return name
    n = name
    # Bracket-quoted: [foo] → foo
    if len(n) >= 2 and n[0] == '[' and n[-1] == ']':
        n = n[1:-1]
    # Trailing legacy type-suffix (only the unambiguous ones — &/!/# are
    # operators / preprocessor / date markers and never lex into IDENTIFIER).
    if n and n[-1] in '$%@':
        n = n[:-1]
    return n.lower()


class SymbolTable:
    def __init__(self, name, parent=None, scope_type='Block'):
        self.name = name
        self.parent = parent
        self.scope_type = scope_type
        self.symbols = {} # name -> {type: ..., kind: Var/Proc/Class, extra: ...}

    def define(self, name, type_name, kind, extra=None):
        self.symbols[_normalize_identifier(name)] = {"type": type_name, "kind": kind, "extra": extra}

    def resolve(self, name):
        key = _normalize_identifier(name)
        if key in self.symbols:
            return self.symbols[key]
        if self.parent:
            return self.parent.resolve(name)
        return None

class Analyzer:
    def __init__(self, config):
        self.config = config
        self.modules = []
        self.global_scope = SymbolTable("Global", scope_type='Global')
        self.errors = []
        self.udts = {} # name_lower -> TypeNode
        self.reference_names = set()
        self._current_labels = None
        self._current_proc_name = None
        self._current_def_type_map = {}
        self._loop_stack = []
        self._for_var_stack = []

        # Load Standard/Config Globals into Global Scope
        for name, defn in self.config.object_model.get("globals", {}).items():
            # Use 'returns' as type if available, otherwise 'type'
            type_name = defn.get("returns", defn.get("type", "Variant"))
            kind = defn.get("type", "Global")
            self.global_scope.define(name, type_name, kind, extra=defn)
            
        # Load Classes into Global Scope (as Types)
        for name in self.config.object_model.get("classes", {}).keys():
            self.global_scope.define(name, name, "Class")

        # Load References as Global Symbols (Treat as Objects/Libraries)
        if "references" in self.config.object_model:
            for ref in self.config.object_model["references"]:
                 self.reference_names.add(ref["name"].lower())
                 self.global_scope.define(ref["name"], ref["name"], "Library")

        # Load Enums into Global Scope
        for enum_name, members in self.config.object_model.get("enums", {}).items():
            self.global_scope.define(enum_name, enum_name, "Enum") # Type = Enum Name
            for member_name, val in members.items():
                self.global_scope.define(member_name, "Long", "EnumItem")
            for member_name, val in members.items():
                self.global_scope.define(member_name, "Long", "EnumItem")

    def add_module(self, module_node):
        self.modules.append(module_node)

    def analyze(self):
        # Pass 1: Populate Symbol Tables
        self.pass1_discovery()
        
        # Pass 2: Verify References
        self.pass2_resolution()
        
        return self.errors

    def pass1_discovery(self):
        for mod in self.modules:
            # Register module name itself (allows usage like Module1.Func)
            self.global_scope.define(mod.name, mod.name, mod.module_type)
            
            # Check for Predeclared ID (Global Instance for Classes/Forms)
            if mod.attributes.get('VB_PredeclaredId', 'False').lower() == 'true':
                 self.global_scope.define(mod.name, mod.name, mod.module_type)

            if mod.module_type == 'Module':
                for var in mod.variables:
                    if var.scope.lower() in ('public', 'global', 'friend'):
                        if getattr(var, 'is_enum_member', False):
                            kind = 'EnumItem'
                        elif getattr(var, 'is_const', False):
                            kind = 'Const'
                        else:
                            kind = 'Variable'
                        self.global_scope.define(var.name, var.type_name, kind)
                
                for proc in mod.procedures:
                    if proc.scope.lower() in ('public', 'friend'):
                         self.global_scope.define(proc.name, proc.return_type, 'Procedure', extra=proc)
                
                # Register Public Types
                for type_name, udt in mod.types.items():
                    if udt.scope.lower() in ('public', 'friend'):
                        self.global_scope.define(type_name, type_name, 'Type')
                        self.udts[type_name.lower()] = udt

            else:
                self.global_scope.define(mod.name, mod.name, 'Class')
                for type_name, udt in mod.types.items():
                     if udt.scope.lower() in ('public', 'friend'):
                         self.global_scope.define(type_name, type_name, 'Type')
                         self.udts[type_name.lower()] = udt

    def pass2_resolution(self):
        for mod in self.modules:
            mod_scope = SymbolTable(mod.name, parent=self.global_scope, scope_type=mod.module_type)

            for var in mod.variables:
                if getattr(var, 'is_enum_member', False):
                    kind = 'EnumItem'
                elif getattr(var, 'is_const', False):
                    kind = 'Const'
                else:
                    kind = 'Variable'
                mod_scope.define(var.name, var.type_name, kind)
            for proc in mod.procedures:
                mod_scope.define(proc.name, proc.return_type, 'Procedure', extra=proc)
            # Register Local/Private Types in Module Scope
            for type_name, udt in mod.types.items():
                mod_scope.define(type_name, type_name, 'Type')
                self.udts[type_name.lower()] = udt

            if mod.module_type in ('Form', 'Class'):
                 mod_scope.define('Me', mod.name, 'Variable')

            # Phase 2.3 — Property Get/Let/Set arity & type compatibility
            self._validate_property_arity(mod)

            # Phase 2.10 — argument-list well-formedness (VBA130-VBA136)
            self._validate_param_lists(mod)

            # Phase 2.12 — duplicate procedure names (VBA150)
            self._validate_duplicate_procedures(mod)

            # Phase 1.6 — declared-type resolution (VBA120/122/123)
            self._validate_declared_types(mod)

            # Phase 3.10 — object-module member restrictions (VBA370)
            self._validate_object_module_members(mod)

            # Phase 3.3 — Declare PtrSafe (64-bit) requirement
            self._validate_ptrsafe_declares(mod)

            # Phase 3.4 — Enum-member uniqueness within an enum
            self._validate_enum_uniqueness(mod)

            # Phase 3.6 — Option Explicit (style-warning, configurable)
            self._validate_option_explicit(mod)

            # Phase 3.1 — Implements <Interface> contract check
            self._validate_implements(mod, mod_scope)

            for proc in mod.procedures:
                self.analyze_procedure(proc, mod_scope, mod)

    def _validate_option_explicit(self, mod):
        """Modules without `Option Explicit` allow implicit (auto-Variant)
        variable creation, which is the #1 source of typo-induced bugs in
        VBA. Style-level severity: doesn't fail compilation but is a strong
        code-quality signal for AI-generated code reviews.
        """
        # Forms have an implicit Option Explicit-like behaviour for their
        # Begin/End block, but their `_Initialize` etc. modules still
        # benefit from the explicit declaration.
        opts = getattr(mod, 'options', {}) or {}
        if not opts.get('explicit', False):
            self.errors.append({
                "file": mod.filename,
                "line": 1,
                "rule_id": "VBA320",
                "severity": "warning",
                "message": (
                    f"Module '{mod.name}' is missing `Option Explicit`. "
                    f"Without it, typo'd variable names silently create new "
                    f"Variant variables — a common source of AI-generated bugs."
                ),
            })

    def _validate_implements(self, mod, mod_scope):
        """For each `Implements X` statement, verify the class provides a
        matching method `X_<Member>` for every public member of the
        interface. The interface itself is resolved either as a same-project
        class module (best effort) or as a library type.
        """
        impls = getattr(mod, 'implements', []) or []
        if not impls:
            return
        for raw in impls:
            iface_name = raw.split('.')[-1]
            iface = self._find_interface_module(iface_name)
            if iface is None:
                # Unknown interface — let the regular identifier resolver
                # (or a future library-type pass) handle it.
                continue
            for proc in iface.procedures:
                if proc.proc_type and proc.proc_type.lower().startswith(('sub', 'function', 'property')):
                    if proc.scope.lower() not in ('public', 'friend'):
                        continue
                    expected_name = f"{iface_name}_{proc.name}"
                    impl = next((p for p in mod.procedures
                                 if p.name.lower() == expected_name.lower()), None)
                    if impl is None:
                        self.errors.append({
                            "file": mod.filename,
                            "line": 0,
                            "rule_id": "VBA330",
                            "severity": "error",
                            "message": (
                                f"Class '{mod.name}' implements '{iface_name}' but is "
                                f"missing the required method '{expected_name}' "
                                f"(matching {proc.proc_type} {proc.name})."
                            ),
                        })
                    else:
                        # VBA153 — the implementing method exists but its
                        # parameter count disagrees with the interface member.
                        iface_argc = len(getattr(proc, 'args', []) or [])
                        impl_argc = len(getattr(impl, 'args', []) or [])
                        if iface_argc != impl_argc:
                            self.errors.append({
                                "file": mod.filename,
                                "line": getattr(impl, 'line', 0) or 0,
                                "rule_id": "VBA153",
                                "severity": "error",
                                "message": (
                                    f"'{expected_name}' does not match the interface "
                                    f"member '{iface_name}.{proc.name}': expected "
                                    f"{iface_argc} parameter(s), got {impl_argc}."
                                ),
                            })

    def _find_interface_module(self, iface_name):
        target = iface_name.lower()
        for mod in self.modules:
            if mod.module_type in ('Class', 'Form') and mod.name.lower() == target:
                return mod
        return None

    # ---- Phase 3.2: RaiseEvent target + argument count ------------------

    def _validate_raise_event(self, tokens, scope, filename, context):
        """`RaiseEvent <Name>(<args>?)` — only legal when <Name> is an
        Event declared in the same Class/Form module. Validate name and
        argument count.
        """
        if not tokens or len(tokens) < 2:
            return
        if tokens[0].type != 'IDENTIFIER' or tokens[0].value.lower() != 'raiseevent':
            return
        name_tok = tokens[1]
        if name_tok.type != 'IDENTIFIER':
            return

        # Look up the event by name in the module that owns the current
        # procedure (events are private to their declaring class).
        event = None
        for mod in self.modules:
            if mod.filename != filename:
                continue
            for proc in mod.procedures:
                if (proc.proc_type or '').lower() == 'event' and proc.name.lower() == name_tok.value.lower():
                    event = proc
                    break
            break
        if event is None:
            self.errors.append({
                "file": filename,
                "line": name_tok.line,
                "rule_id": "VBA340",
                "severity": "error",
                "message": (
                    f"RaiseEvent '{name_tok.value}' in '{context}': no matching "
                    f"`Event {name_tok.value}` declared in this module."
                ),
            })
            return

        # Count arguments — between the parens following the event name.
        # Token shape: [RaiseEvent, Name, '(', arg, ',', arg, ..., ')'].
        if len(tokens) < 3 or not (tokens[2].type == 'OPERATOR' and tokens[2].value == '('):
            actual = 0
        else:
            actual = self._count_call_args(tokens, 2)

        expected_min = sum(1 for a in event.args if not a.is_optional and not a.is_paramarray)
        expected_max = len(event.args)
        # ParamArray accepts unlimited.
        if any(a.is_paramarray for a in event.args):
            expected_max = 9999

        if actual < expected_min or actual > expected_max:
            self.errors.append({
                "file": filename,
                "line": name_tok.line,
                "rule_id": "VBA341",
                "severity": "error",
                "message": (
                    f"RaiseEvent '{name_tok.value}' arg count mismatch: "
                    f"event declared with {expected_min}{('..' + str(expected_max)) if expected_max != expected_min else ''} "
                    f"parameter(s), got {actual}."
                ),
            })

    def _count_call_args(self, tokens, lparen_idx):
        """Given a token list and the index of the opening `(`, return the
        number of comma-separated top-level arguments inside the parens.
        Empty parens → 0.
        """
        depth = 0
        count = 0
        seen_any = False
        for j in range(lparen_idx, len(tokens)):
            t = tokens[j]
            if t.type == 'OPERATOR' and t.value == '(':
                depth += 1
            elif t.type == 'OPERATOR' and t.value == ')':
                depth -= 1
                if depth == 0:
                    break
            elif depth == 1 and t.type == 'OPERATOR' and t.value == ',':
                count += 1
                seen_any = True
            elif depth == 1:
                seen_any = True
        if seen_any:
            return count + 1
        return 0

    def _validate_ptrsafe_declares(self, mod):
        """In 64-bit Office (the modern default since Office 2010 for x64
        and the current Microsoft 365 default) every `Declare` statement
        must carry the `PtrSafe` attribute. The check is gated by the
        `WIN64` / `VBA7` conditional-compilation defines so a 32-bit-only
        run can opt out by passing `--define WIN64=False`.
        """
        defs = self.config.definitions
        # If the user has explicitly set 32-bit, skip the check.
        if defs.get('WIN64') is False or defs.get('VBA7') is False:
            return
        for proc in mod.procedures:
            if not proc.is_declare:
                continue
            if proc.is_ptrsafe:
                continue
            self.errors.append({
                "file": mod.filename,
                "line": getattr(proc, 'declare_line', 0) or 0,
                "rule_id": "VBA300",
                "severity": "error",
                "message": (
                    f"Declare {proc.proc_type} '{proc.name}' is missing the "
                    f"`PtrSafe` attribute. 64-bit VBA (Office 2010 x64+, "
                    f"Microsoft 365) refuses to compile API declarations "
                    f"without it. Add `PtrSafe` after `Declare`."
                ),
            })

    def _validate_enum_uniqueness(self, mod):
        """Within a single Enum block all member names must be unique.
        VBA accepts duplicate *values* (`A = 1: B = 1`) but not duplicate
        *names*. The error here mirrors the IDE's compile-time message.
        """
        for type_name, type_node in mod.types.items():
            members = getattr(type_node, 'members', [])
            # Heuristic: enums are registered alongside UDTs in mod.types,
            # but enum members carry a literal Long type while UDT members
            # have arbitrary types — and Enum members are exposed
            # globally via pass1_discovery. We cover both: if a member
            # name repeats inside `members`, flag it.
            seen = {}
            for m in members:
                key = m.name.lower()
                first_seen = seen.get(key)
                if first_seen is not None:
                    self.errors.append({
                        "file": mod.filename,
                        "line": 0,
                        "rule_id": "VBA310",
                        "severity": "error",
                        "message": (
                            f"Enum / Type '{type_name}' has duplicate member name "
                            f"'{m.name}'."
                        ),
                    })
                else:
                    seen[key] = m

    def _validate_property_arity(self, mod):
        """Group properties by name and verify the Get/Let/Set contract:

        - Property Let / Property Set must have exactly Get-arg-count + 1
          parameters (the trailing parameter receives the assigned value).
        - The trailing parameter type of Let/Set should be assignable to
          the Property Get return type.
        - Property Set is only valid when the assigned-value parameter is
          object-typed; Property Let is the value-typed sibling.
        """
        from collections import defaultdict
        groups = defaultdict(dict)
        for proc in mod.procedures:
            ptype = (proc.proc_type or "").lower()
            if not ptype.startswith("property"):
                continue
            kind = ptype.split()[-1]  # 'get' | 'let' | 'set'
            if kind in ('get', 'let', 'set'):
                groups[proc.name.lower()][kind] = proc

        for prop_name, members in groups.items():
            get_proc = members.get('get')
            let_proc = members.get('let')
            set_proc = members.get('set')

            # Note: VBA *does* permit both Property Let AND Property Set on
            # the same property name as long as they accept different
            # value-parameter types (real-world libraries like JSONBag use
            # this for polymorphic properties), so we explicitly do NOT
            # flag the Let+Set combination.

            for accessor in (let_proc, set_proc):
                if accessor is None:
                    continue
                if not accessor.args:
                    self.errors.append({
                        "file": mod.filename,
                        "line": 0,
                        "rule_id": "VBA221",
                        "severity": "error",
                        "message": (
                            f"{accessor.proc_type} '{accessor.name}' must have at least one parameter "
                            f"(the assigned value)."
                        ),
                    })
                    continue

                if get_proc is not None:
                    expected = len(get_proc.args) + 1
                    actual = len(accessor.args)
                    if actual != expected:
                        self.errors.append({
                            "file": mod.filename,
                            "line": 0,
                            "rule_id": "VBA222",
                            "severity": "error",
                            "message": (
                                f"{accessor.proc_type} '{accessor.name}' must have {expected} parameter(s) "
                                f"(Property Get '{get_proc.name}' has {len(get_proc.args)}); got {actual}."
                            ),
                        })
                        continue

                # Set/Let semantic mismatch — independent of Get presence.
                last_arg = accessor.args[-1]
                kind = accessor.proc_type.split()[-1].lower()

                # VBA152 — a Property Get and its Let must agree on the value
                # type. Restricted to clearly-scalar primitives on both sides
                # to stay false-positive-safe (object hierarchies / Variant
                # coercions are intentionally not flagged).
                if get_proc is not None and kind == 'let':
                    get_t = (get_proc.return_type or 'Variant')
                    let_t = (last_arg.type_name or 'Variant')
                    if (self._is_clearly_scalar(get_t) and self._is_clearly_scalar(let_t)
                            and get_t.replace('()', '').strip().lower()
                                != let_t.replace('()', '').strip().lower()):
                        self.errors.append({
                            "file": mod.filename,
                            "line": getattr(accessor, 'line', 0) or 0,
                            "rule_id": "VBA152",
                            "severity": "error",
                            "message": (
                                f"Property '{accessor.name}' is inconsistent: Property Get "
                                f"returns '{get_t}' but Property Let takes '{let_t}'."
                            ),
                        })
                rhs_is_object = self._is_clearly_object(last_arg.type_name)
                rhs_is_scalar = self._is_clearly_scalar(last_arg.type_name)
                if kind == 'set' and rhs_is_scalar:
                    self.errors.append({
                        "file": mod.filename,
                        "line": 0,
                        "rule_id": "VBA223",
                        "severity": "error",
                        "message": (
                            f"Property Set '{accessor.name}' last parameter '{last_arg.name}' "
                            f"is scalar type '{last_arg.type_name}'; use Property Let instead."
                        ),
                    })
                elif kind == 'let' and rhs_is_object:
                    self.errors.append({
                        "file": mod.filename,
                        "line": 0,
                        "rule_id": "VBA224",
                        "severity": "error",
                        "message": (
                            f"Property Let '{accessor.name}' last parameter '{last_arg.name}' "
                            f"is object type '{last_arg.type_name}'; use Property Set instead."
                        ),
                    })

    def _validate_duplicate_procedures(self, mod):
        """VBA150 — two procedures share a name in the same module.

        Property Get/Let/Set legitimately share a name (different accessor
        kinds), so the dedup key includes the accessor kind for properties.
        Property and non-property members are grouped separately to avoid
        edge-case false positives.
        """
        seen = {}
        for proc in mod.procedures:
            ptype = (proc.proc_type or '').lower()
            if ptype == 'event':
                continue  # events have their own declaration namespace
            if ptype.startswith('property'):
                key = ('prop', proc.name.lower(), ptype.split()[-1])
            else:
                key = ('proc', proc.name.lower())
            if key in seen:
                self.errors.append({
                    "file": mod.filename,
                    "line": getattr(proc, 'line', 0) or 0,
                    "rule_id": "VBA150",
                    "severity": "error",
                    "message": (
                        f"Ambiguous name detected: '{proc.name}' is already defined "
                        f"in module '{mod.name}'."
                    ),
                })
            else:
                seen[key] = proc

    def _validate_param_lists(self, mod):
        """Phase 2.10 — argument-list well-formedness (VBA130-VBA136).

        Pure syntactic checks over each procedure's declared parameters; no
        type model required, so these never false-positive on valid code.
        """
        for proc in mod.procedures:
            args = getattr(proc, 'args', None) or []
            if not args:
                continue
            line = getattr(proc, 'line', 0) or 0
            ptype = proc.proc_type or "procedure"
            # The trailing value parameter of a Property Let/Set is implicitly
            # required (it receives the assigned RHS) and may legitimately
            # follow Optional index parameters — exempt it from VBA130.
            ptl = ptype.lower()
            is_value_param_proc = ptl.startswith('property') and ptl.split()[-1] in ('let', 'set')
            seen_optional = False
            paramarray_count = 0
            n = len(args)
            for idx, arg in enumerate(args):
                is_opt = getattr(arg, 'is_optional', False)
                is_pa = getattr(arg, 'is_paramarray', False)

                # VBA130 — once an Optional appears, every following parameter
                # must be Optional (or the trailing ParamArray).
                if is_opt:
                    seen_optional = True
                elif seen_optional and not is_pa and not (is_value_param_proc and idx == n - 1):
                    self.errors.append({
                        "file": mod.filename, "line": line,
                        "rule_id": "VBA130", "severity": "error",
                        "message": (
                            f"Parameter '{arg.name}' of {ptype} '{proc.name}' is "
                            f"required but follows an Optional parameter. Every "
                            f"parameter after the first Optional must also be Optional."
                        ),
                    })

                if is_pa:
                    paramarray_count += 1
                    # VBA131 — ParamArray must be the last parameter.
                    if idx != n - 1:
                        self.errors.append({
                            "file": mod.filename, "line": line,
                            "rule_id": "VBA131", "severity": "error",
                            "message": (
                                f"ParamArray '{arg.name}' must be the last parameter "
                                f"of {ptype} '{proc.name}'."
                            ),
                        })
                    # VBA132 — only one ParamArray allowed.
                    if paramarray_count > 1:
                        self.errors.append({
                            "file": mod.filename, "line": line,
                            "rule_id": "VBA132", "severity": "error",
                            "message": (
                                f"{ptype} '{proc.name}' declares more than one "
                                f"ParamArray; only one is allowed."
                            ),
                        })
                    # VBA134 — ParamArray element type must be Variant.
                    base = (arg.type_name or 'Variant').replace('()', '').strip()
                    if base and base.lower() != 'variant':
                        self.errors.append({
                            "file": mod.filename, "line": line,
                            "rule_id": "VBA134", "severity": "error",
                            "message": (
                                f"ParamArray '{arg.name}' must be declared as an array "
                                f"of Variant, not '{arg.type_name}'."
                            ),
                        })

                # VBA136 — array parameters cannot be passed ByVal.
                if (arg.type_name or '').endswith('()') and getattr(arg, 'mechanism', 'ByRef') == 'ByVal':
                    self.errors.append({
                        "file": mod.filename, "line": line,
                        "rule_id": "VBA136", "severity": "error",
                        "message": (
                            f"Array parameter '{arg.name}' of {ptype} '{proc.name}' "
                            f"must be passed ByRef, not ByVal."
                        ),
                    })

                # VBA137 — user-defined types may not be passed ByVal (enums
                # are Long under the hood and are exempt).
                if getattr(arg, 'mechanism', 'ByRef') == 'ByVal':
                    base_t = (arg.type_name or '').replace('()', '').strip().lower()
                    udt = self.udts.get(base_t)
                    if udt is not None and not getattr(udt, 'is_enum', False):
                        self.errors.append({
                            "file": mod.filename, "line": line,
                            "rule_id": "VBA137", "severity": "error",
                            "message": (
                                f"User-defined type '{arg.type_name}' parameter "
                                f"'{arg.name}' of {ptype} '{proc.name}' may not be "
                                f"passed ByVal; use ByRef."
                            ),
                        })

            # VBA133 — a ParamArray cannot coexist with Optional parameters.
            if paramarray_count and any(getattr(a, 'is_optional', False) for a in args):
                self.errors.append({
                    "file": mod.filename, "line": line,
                    "rule_id": "VBA133", "severity": "error",
                    "message": (
                        f"{ptype} '{proc.name}' combines Optional parameters with a "
                        f"ParamArray; this is not allowed."
                    ),
                })

    def analyze_procedure(self, proc, mod_scope, mod):
        proc_scope = SymbolTable(proc.name, parent=mod_scope, scope_type='Procedure')

        for arg in proc.args:
            proc_scope.define(arg.name, arg.type_name, 'Variable')

        # Pass 1.5 — collect all labels reachable inside this procedure so
        # GoTo / On Error GoTo / Resume / GoSub can be validated against
        # them in pass 2.
        labels = set()
        self._collect_labels(proc.body, labels)
        self._current_labels = labels
        self._current_proc_name = proc.name
        self._current_def_type_map = getattr(mod, "def_type_map", {}) or {}

        # Phase 2.11 — VBA175: duplicate label declarations in the procedure.
        self._validate_duplicate_labels(proc, mod.filename)

        # Phase 2.11 — loop context for Exit For / Exit Do (VBA170/171) and
        # For-counter checks (VBA172/173). Reset per procedure.
        self._loop_stack = []
        self._for_var_stack = []

        try:
            self.analyze_block(proc.body, proc_scope, mod.filename, proc.name, with_stack=[])
        finally:
            self._current_labels = None
            self._current_proc_name = None
            self._current_def_type_map = {}
            self._loop_stack = []
            self._for_var_stack = []

    def _validate_duplicate_labels(self, proc, filename):
        """VBA175 — the same line label declared twice in one procedure."""
        seen = {}
        ordered = []
        self._collect_labels_ordered(proc.body, ordered)
        for tok in ordered:
            key = tok.value.lower()
            if key in seen:
                self.errors.append({
                    "file": filename,
                    "line": tok.line,
                    "rule_id": "VBA175",
                    "severity": "error",
                    "message": (
                        f"Duplicate label '{tok.value}' in '{proc.name}'. "
                        f"A label may be declared only once per procedure."
                    ),
                })
            else:
                seen[key] = tok

    def _collect_labels_ordered(self, nodes, out):
        """Like `_collect_labels` but preserves order and keeps the token
        (not just the lowercased name) so duplicates can be reported."""
        from .parser import IfNode, WithNode, ForNode, DoNode, SelectNode, StatementNode
        for node in nodes:
            if isinstance(node, StatementNode):
                if self.is_label(node.tokens):
                    out.append(node.tokens[0])
            elif isinstance(node, IfNode):
                self._collect_labels_ordered(node.true_block, out)
                for _cond, blk in node.else_blocks:
                    self._collect_labels_ordered(blk, out)
                if node.else_block:
                    self._collect_labels_ordered(node.else_block, out)
            elif isinstance(node, WithNode):
                self._collect_labels_ordered(node.body, out)
            elif isinstance(node, (ForNode, DoNode)):
                self._collect_labels_ordered(node.body, out)
            elif isinstance(node, SelectNode):
                for case in node.cases:
                    self._collect_labels_ordered(case.body, out)

    def _validate_jump_target(self, tokens, filename, context):
        """Validate `GoTo`, `On Error GoTo`, `Resume`, `GoSub` against the
        per-procedure label registry built in `analyze_procedure`.

        Special forms accepted without a label:
        - `On Error GoTo 0`     → resets the error handler
        - `On Error GoTo -1`    → clears active error (Office VBA)
        - `On Error Resume Next`
        - `Resume` / `Resume Next` (no label)
        - `On <expr> GoTo lbl1, lbl2, ...` and `On <expr> GoSub …`
          (computed-GoTo jump table)
        """
        if not tokens or self._current_labels is None:
            return

        i = 0
        n = len(tokens)
        first = tokens[0].value.lower() if tokens[0].type == 'IDENTIFIER' else None

        # `On Error GoTo <target>` / `On Error Resume Next`
        if first == 'on' and n >= 2 and tokens[1].type == 'IDENTIFIER' and tokens[1].value.lower() == 'error':
            i = 2
            if i >= n:
                return
            kw = tokens[i].value.lower() if tokens[i].type == 'IDENTIFIER' else ''
            if kw == 'resume':
                # `On Error Resume Next` — no target to validate
                return
            if kw == 'goto':
                i += 1
                if i >= n:
                    return
                target_tok = tokens[i]
                # `On Error GoTo 0` / `On Error GoTo -1` — reset/clear handler
                if target_tok.type in ('INTEGER',):
                    return
                if target_tok.type == 'OPERATOR' and target_tok.value == '-':
                    # negative integer e.g. -1
                    return
                if target_tok.type == 'IDENTIFIER':
                    self._check_label_exists(target_tok, filename, context, "On Error GoTo")
                    return
            return

        # `On <expr> GoTo lbl1, lbl2, ...` / `On <expr> GoSub lbl1, lbl2`
        # (computed-GoTo jump table; the labels after the keyword are a
        # comma-separated list, not value references.)
        if first == 'on':
            j = self._find_on_jump_keyword(tokens)
            if j is not None:
                kind = "On GoTo" if tokens[j].value.lower() == 'goto' else "On GoSub"
                for tok in self._iter_label_list(tokens[j + 1:]):
                    self._check_label_exists(tok, filename, context, kind)
                return

        if first == 'goto':
            if n >= 2 and tokens[1].type == 'IDENTIFIER':
                self._check_label_exists(tokens[1], filename, context, "GoTo")
            elif n >= 2 and tokens[1].type == 'INTEGER':
                # Numeric labels like `GoTo 100` — VBA allows them; require numeric label registered
                # (We don't track numeric labels yet; tolerate to avoid false positives.)
                return
            return

        if first == 'resume':
            if n == 1:
                return  # bare `Resume`
            tok = tokens[1]
            if tok.type == 'IDENTIFIER':
                if tok.value.lower() == 'next':
                    return
                self._check_label_exists(tok, filename, context, "Resume")
            return

        if first == 'gosub':
            if n >= 2 and tokens[1].type == 'IDENTIFIER':
                self._check_label_exists(tokens[1], filename, context, "GoSub")
            return

    def _check_label_exists(self, token, filename, context, kind):
        name = token.value.lower()
        if name not in self._current_labels:
            self.errors.append({
                "file": filename,
                "line": token.line,
                "rule_id": "VBA201",
                "severity": "error",
                "message": (
                    f"{kind} target '{token.value}' is not a label in '{context}'. "
                    f"Declared labels: {sorted(self._current_labels) or 'none'}."
                ),
            })

    @staticmethod
    def _find_on_jump_keyword(tokens):
        """Locate the index of `GoTo`/`GoSub` in an `On <expr> GoTo …` /
        `On <expr> GoSub …` statement (the computed-GoTo form). Returns
        None if this isn't an On-jump-table statement.

        Skips identifiers/operators inside the `<expr>` portion. The
        keyword we're looking for is the bare `GoTo` or `GoSub`
        (NOT `On Error GoTo`, handled separately).
        """
        if not tokens or tokens[0].type != 'IDENTIFIER' or tokens[0].value.lower() != 'on':
            return None
        depth = 0
        for idx in range(1, len(tokens)):
            t = tokens[idx]
            if t.type == 'OPERATOR':
                if t.value == '(':
                    depth += 1
                elif t.value == ')':
                    depth -= 1
                continue
            if depth == 0 and t.type == 'IDENTIFIER':
                lv = t.value.lower()
                if lv in ('goto', 'gosub'):
                    return idx
                if lv == 'error':
                    return None  # handled by On Error branch
        return None

    @staticmethod
    def _iter_label_list(tokens):
        """Yield the identifier tokens from a comma-separated label list
        like `lbl1, lbl2, lblN`. Stops at end of token stream or at a
        colon (statement separator)."""
        for t in tokens:
            if t.type == 'OPERATOR':
                if t.value == ':':
                    return
                # commas and other operators are skipped
                continue
            if t.type == 'IDENTIFIER':
                yield t

    # ---- Phase 2.2: Set vs. Let assignment enforcement -------------------

    _SCALAR_TYPES = {
        "boolean", "byte", "integer", "long", "longlong", "longptr",
        "single", "double", "currency", "decimal", "date", "string",
    }
    _CALLABLE_KINDS = {"sub", "function", "procedure", "property", "library"}

    def _is_clearly_scalar(self, type_name):
        if not type_name:
            return False
        t = type_name.lower()
        if t in self._SCALAR_TYPES:
            return True
        # User-defined Type (UDT) is value-typed, not Object.
        if t in self.udts:
            return True
        return False

    def _is_clearly_object(self, type_name):
        if not type_name:
            return False
        t = type_name.lower()
        if t == "object":
            return True
        # Known class in the standard / loaded model.
        if self.config.object_model.get("classes", {}).get(t):
            return True
        return False

    _BUILTIN_TYPE_NAMES = {
        "boolean", "byte", "integer", "long", "longlong", "longptr",
        "single", "double", "currency", "decimal", "date", "string",
        "object", "variant", "any", "iunknown", "idispatch",
    }

    # Type-declaration suffix → implied type. Only $ / % / @ survive lexing
    # as part of the identifier token (& ! # are tokenised separately).
    _TYPE_SUFFIX_MAP = {'$': 'string', '%': 'integer', '@': 'currency'}

    # Universal VBA / stdole / MSForms intrinsic type & enum names that are
    # not part of every host model but are always available in the VBA
    # runtime. Keeps VBA120 quiet on these regardless of which `--host`
    # model (if any) is loaded.
    _INTRINSIC_TYPE_NAMES = {
        # VBA.* intrinsic enums
        "vbtristate", "vbcomparemethod", "vbcalltype", "vbdayofweek",
        "vbfirstweekofyear", "vbvartype", "vbmsgboxresult", "vbmsgboxstyle",
        "vbfileattribute", "vbappwinstyle", "vbstrconv", "vbqueryclose",
        "vbimestatus",
        # stdole picture / font / colour types
        "ole_color", "ole_handle", "ole_tristate", "ole_optexclusive",
        "ole_cancelbool", "ole_enabledefaultbool",
        "ole_xpos_pixels", "ole_ypos_pixels", "ole_xsize_pixels", "ole_ysize_pixels",
        "ole_xpos_himetric", "ole_ypos_himetric", "ole_xsize_himetric", "ole_ysize_himetric",
        "ole_xpos_container", "ole_ypos_container", "ole_xsize_container", "ole_ysize_container",
        "font", "ifont", "ifontdisp", "stdfont",
        "picture", "ipicture", "ipicturedisp", "stdpicture",
        # MSForms / VB form base types
        "form", "userform", "control", "frame", "page", "msforms",
    }

    def _is_known_type(self, type_name):
        """True if `type_name` resolves to a built-in, a source-declared
        UDT/Enum, a class/form module, a host-model class/enum, or a
        library-qualified name. Used by VBA120; deliberately lenient on
        dotted/qualified names to avoid false positives on un-modelled hosts.
        """
        if not type_name:
            return True
        t = type_name.strip()
        if t.endswith('()'):
            t = t[:-2].strip()
        if '*' in t:  # fixed-length string `String * N`
            t = t.split('*')[0].strip()
        if not t:
            return True
        low = t.lower()
        if low in self._BUILTIN_TYPE_NAMES or low in self._INTRINSIC_TYPE_NAMES:
            return True
        if low in self.udts:
            return True
        if self._is_class_type(t) or self._is_enum_type(t):
            return True
        if '.' in t:
            # Library-qualified (`Scripting.Dictionary`, `Excel.Range`): check
            # the trailing segment, else accept — we don't model every host.
            tail = t.rpartition('.')[2]
            if self._is_known_type(tail):
                return True
            return True
        sym = self.global_scope.resolve(t)
        if sym and (sym.get('kind') or '').lower() in ('type', 'class', 'enum', 'library', 'module', 'form'):
            return True
        for m in self.modules:
            if m.name.lower() == low:
                return True
        return False

    def _check_type_known(self, type_name, line, filename, context):
        """VBA120 — emit an 'undefined type' error if `type_name` is unknown."""
        if self._is_known_type(type_name):
            return
        base = (type_name or '').replace('()', '').split('*')[0].strip()
        self.errors.append({
            "file": filename,
            "line": line or 0,
            "rule_id": "VBA120",
            # Warning, not error: with an incomplete host model we cannot be
            # certain an unresolved name is a typo vs. an un-modelled host
            # type, so VBA120 surfaces the risk without gating compile-safety.
            "severity": "warning",
            "message": (
                f"User-defined type '{base}' is not defined (in {context}). "
                f"If it is a host type, load the matching model with `--host`."
            ),
        })

    def _validate_object_module_members(self, mod):
        """VBA370 — constants, arrays and Declare statements may not be Public
        members of an object module (Class / Form). Such members are only
        legal in standard modules."""
        if mod.module_type not in ('Class', 'Form'):
            return
        public_scopes = ('public', 'global')
        for var in mod.variables:
            if getattr(var, 'is_enum_member', False):
                continue
            if (var.scope or '').lower() not in public_scopes:
                continue
            kind = None
            if getattr(var, 'is_const', False):
                kind = "constant"
            elif (var.type_name or '').endswith('()'):
                kind = "array"
            if kind:
                self.errors.append({
                    "file": mod.filename, "line": 0,
                    "rule_id": "VBA370", "severity": "error",
                    "message": (
                        f"Public {kind} '{var.name}' is not allowed as a member of "
                        f"object module '{mod.name}'. Use a standard module, or make "
                        f"it Private."
                    ),
                })
        for proc in mod.procedures:
            if getattr(proc, 'is_declare', False) and (proc.scope or '').lower() in public_scopes:
                self.errors.append({
                    "file": mod.filename, "line": getattr(proc, 'line', 0) or 0,
                    "rule_id": "VBA370", "severity": "error",
                    "message": (
                        f"Public Declare '{proc.name}' is not allowed in object module "
                        f"'{mod.name}'. Declare statements in a class must be Private."
                    ),
                })

    def _validate_declared_types(self, mod):
        """Phase 1.6 — VBA120/122/123: resolve declared type names and reject
        empty Type/Enum blocks."""
        for tname, udt in mod.types.items():
            is_enum = getattr(udt, 'is_enum', False)
            if not udt.members:
                self.errors.append({
                    "file": mod.filename, "line": 0,
                    "rule_id": "VBA123" if is_enum else "VBA122",
                    "severity": "error",
                    "message": (
                        f"Enum '{tname}' has no members; an empty Enum is not allowed."
                        if is_enum else
                        f"Type '{tname}' has no members; a user-defined type must "
                        f"declare at least one member."
                    ),
                })
            elif not is_enum:
                for m in udt.members:
                    self._check_type_known(m.type_name, 0, mod.filename, f"Type {tname}")

        for var in mod.variables:
            if not getattr(var, 'is_enum_member', False):
                self._check_type_known(var.type_name, 0, mod.filename, f"module '{mod.name}'")

        for proc in mod.procedures:
            line = getattr(proc, 'line', 0)
            for arg in (proc.args or []):
                self._check_type_known(arg.type_name, line, mod.filename,
                                       f"{proc.proc_type} '{proc.name}'")
            pt = (proc.proc_type or '').lower()
            if pt.startswith('function') or pt == 'property get':
                self._check_type_known(proc.return_type, line, mod.filename,
                                       f"{proc.proc_type} '{proc.name}'")

    def _split_assignment(self, tokens):
        """If `tokens` is an assignment, return (lhs_tokens, eq_index, rhs_tokens, has_set).
        Otherwise return None.

        Recognised forms:
            <lhs> = <rhs>
            Set <lhs> = <rhs>
            Let <lhs> = <rhs>
        """
        if not tokens:
            return None
        # Strip trailing colon used as statement separator
        toks = [t for t in tokens if not (t.type == 'OPERATOR' and t.value == ':')]
        if not toks:
            return None

        has_set = False
        has_let = False
        start = 0
        first = toks[0]
        if first.type == 'IDENTIFIER' and first.value.lower() == 'set':
            has_set = True
            start = 1
        elif first.type == 'IDENTIFIER' and first.value.lower() == 'let':
            has_let = True
            start = 1

        # Find an `=` not preceded by `<`/`>`/`<=`/`>=` and not part of `:=` / `<>`.
        eq_index = None
        paren_depth = 0
        for idx in range(start, len(toks)):
            t = toks[idx]
            if t.type == 'OPERATOR' and t.value == '(':
                paren_depth += 1
            elif t.type == 'OPERATOR' and t.value == ')':
                paren_depth -= 1
            elif paren_depth == 0 and t.type == 'OPERATOR' and t.value == '=':
                # avoid named-arg form `name:=value` — but `:=` is its own token.
                eq_index = idx
                break
        if eq_index is None:
            return None

        lhs = toks[start:eq_index]
        rhs = toks[eq_index + 1:]
        if not lhs:
            return None
        return lhs, eq_index, rhs, has_set, has_let

    def _resolve_lhs_type(self, lhs_tokens, scope):
        """Best-effort type resolution for the assignment LHS.
        Returns (type_name, kind) or (None, None).

        P2.6 — walks the full dotted chain `a.b.c.d`. If any hop cannot
        be resolved (member not in UDT/class, or parent type is a
        permissive bag like `Object`/`Variant`), returns (None, None) so
        downstream validators skip rather than mis-typing the LHS.
        """
        if not lhs_tokens:
            return None, None
        first = lhs_tokens[0]
        if first.type != 'IDENTIFIER':
            return None, None
        sym = scope.resolve(first.value)
        if not sym:
            return None, None
        type_name = sym.get("type")
        kind = sym.get("kind")
        i = 1
        while i + 1 < len(lhs_tokens) and lhs_tokens[i].type == 'OPERATOR' and lhs_tokens[i].value == '.':
            member_tok = lhs_tokens[i + 1]
            if member_tok.type != 'IDENTIFIER':
                return None, None
            # Bail out on permissive bag-types: the member's real type
            # is unknowable from declaration alone.
            if (type_name or "").lower() in ("object", "variant", "unknown", ""):
                return None, None
            resolved = self.resolve_member(type_name or "", member_tok.value)
            if not resolved:
                return None, None
            type_name, kind, _extra = resolved
            i += 2
        return type_name, kind

    def _validate_set_vs_let(self, tokens, scope, filename, context):
        if not tokens:
            return
        parsed = self._split_assignment(tokens)
        if not parsed:
            return
        lhs, _eq, _rhs, has_set, has_let = parsed  # noqa: F841

        # Bare identifier (`x = …`) is the common form. Dotted chains
        # (`obj.member = …`) are now supported via _resolve_lhs_type after
        # P2.6 — but indexed targets (`a(0) = …`) and bang-operator LHS
        # still require expression-level evaluation we don't model, so
        # skip those to avoid false positives.
        if not lhs or lhs[0].type != 'IDENTIFIER':
            return

        # VBA164 — `Me` is read-only; it can never be an assignment target.
        if lhs[0].value.lower() == 'me' and (len(lhs) == 1):
            self.errors.append({
                "file": filename, "line": lhs[0].line,
                "rule_id": "VBA164", "severity": "error",
                "message": f"Invalid use of Me keyword: cannot assign to Me in '{context}'.",
            })
            return

        is_bare = len(lhs) == 1
        is_dotted_chain = (
            not is_bare
            and all(
                (t.type == 'IDENTIFIER') or (t.type == 'OPERATOR' and t.value == '.')
                for t in lhs
            )
        )
        if not is_bare and not is_dotted_chain:
            return

        type_name, kind = self._resolve_lhs_type(lhs, scope)
        if type_name is None or kind is None:
            return

        kind_l = (kind or "").lower()
        if kind_l in self._CALLABLE_KINDS:
            return

        # VBA160 — assigning to a constant or an Enum member is illegal.
        if kind_l in ("const", "enumitem"):
            what = "constant" if kind_l == "const" else "Enum member"
            self.errors.append({
                "file": filename, "line": lhs[0].line,
                "rule_id": "VBA160", "severity": "error",
                "message": (
                    f"Assignment to {what} '{''.join(t.value for t in lhs)}' is not "
                    f"permitted in '{context}'."
                ),
            })
            return

        if kind_l in ("type", "class", "property"):
            return

        line = lhs[0].line
        display = "".join(t.value for t in lhs)
        is_scalar = self._is_clearly_scalar(type_name)
        is_object = self._is_clearly_object(type_name)

        if has_set and is_scalar:
            self.errors.append({
                "file": filename,
                "line": line,
                "rule_id": "VBA210",
                "severity": "error",
                "message": (
                    f"`Set` used on non-object target '{display}' of type '{type_name}' "
                    f"in '{context}'. Use `Let` or assignment without `Set`."
                ),
            })
            return

        if not has_set and is_object and not has_let:
            self.errors.append({
                "file": filename,
                "line": line,
                "rule_id": "VBA211",
                "severity": "error",
                "message": (
                    f"Object assignment to '{display}' (type '{type_name}') "
                    f"requires `Set` in '{context}'."
                ),
            })

    # ----------------------------------------------------------------------

    def _collect_labels(self, nodes, out):
        """Recursively walk every block under a procedure body and record
        every line label so jump targets can be verified.
        """
        for node in nodes:
            if isinstance(node, StatementNode):
                if self.is_label(node.tokens):
                    out.add(node.tokens[0].value.lower())
            elif isinstance(node, IfNode):
                self._collect_labels(node.true_block, out)
                for _cond, blk in node.else_blocks:
                    self._collect_labels(blk, out)
                if node.else_block:
                    self._collect_labels(node.else_block, out)
            elif isinstance(node, WithNode):
                self._collect_labels(node.body, out)
            elif isinstance(node, ForNode):
                self._collect_labels(node.body, out)
            elif isinstance(node, DoNode):
                self._collect_labels(node.body, out)
            elif isinstance(node, SelectNode):
                for case in node.cases:
                    self._collect_labels(case.body, out)

    def analyze_block(self, nodes, scope, filename, context, with_stack):
        unreachable = False
        prev_node = None

        for node in nodes:
            if isinstance(node, StatementNode):
                # Check for Label
                if self.is_label(node.tokens):
                    unreachable = False

                # Check for Control Flow Boundary (e.g. End If, Else, Next) -> Reset unreachable
                # Heuristic: If we hit a block boundary, assume the jump was conditional or we merged back
                if unreachable and self.is_control_flow_boundary(node.tokens):
                    unreachable = False

                if unreachable:
                    if not self.is_ignorable(node.tokens):
                        self.errors.append({
                            "file": filename,
                            "line": node.tokens[0].line,
                            "rule_id": "VBA009",
                            "severity": "warning",
                            "message": f"Unreachable code detected in '{context}'."
                        })

                # Phase 2.1 — validate jump targets before normal analysis
                # so we surface bad jumps even if expression analysis later
                # bails out on the same line.
                self._validate_jump_target(node.tokens, filename, context)

                # Phase 2.2 — Set vs. Let on assignments
                self._validate_set_vs_let(node.tokens, scope, filename, context)

                # Phase 2.4 — Operator-type sanity (literal-only)
                self._validate_operator_types(node.tokens, filename, context)

                # Phase 3.2 — RaiseEvent target + arity
                self._validate_raise_event(node.tokens, scope, filename, context)

                # Check for Dim
                if node.tokens and node.tokens[0].value.lower() in ('dim', 'static', 'const'):
                     self.process_dim(node.tokens, scope, filename, context, with_stack)
                elif node.tokens and node.tokens[0].value.lower() == 'raiseevent':
                     # Suppress regular identifier resolution on the event name
                     # — events are only visible to their declaring class and
                     # _validate_raise_event has already vetted them.
                     pass
                else:
                     self.analyze_statement(node.tokens, scope, filename, context, with_stack)

                # Check for Exit Mismatch
                if node.tokens and node.tokens[0].value.lower() == 'exit':
                    if len(node.tokens) > 1:
                        exit_kind = node.tokens[1].value.lower()
                        # VBA170/171 — `Exit For` / `Exit Do` only valid inside
                        # the matching loop kind.
                        if exit_kind == 'for' and 'for' not in getattr(self, '_loop_stack', []):
                            self.errors.append({
                                "file": filename,
                                "line": node.tokens[0].line,
                                "rule_id": "VBA170",
                                "severity": "error",
                                "message": f"Exit For is not within a For...Next loop in '{context}'.",
                            })
                        elif exit_kind == 'do' and 'do' not in getattr(self, '_loop_stack', []):
                            self.errors.append({
                                "file": filename,
                                "line": node.tokens[0].line,
                                "rule_id": "VBA171",
                                "severity": "error",
                                "message": f"Exit Do is not within a Do...Loop in '{context}'.",
                            })
                        if exit_kind in ('sub', 'function', 'property'):
                            # Verify against context
                            # Resolve context in parent scope
                            # context is proc name
                            proc_sym = scope.parent.resolve(context)
                            if proc_sym and proc_sym.get('kind') == 'Procedure':
                                proc_def = proc_sym['extra']
                                proc_type_lower = proc_def.proc_type.lower()

                                mismatch = False
                                if exit_kind == 'sub':
                                    if not proc_type_lower.startswith('sub'):
                                        mismatch = True
                                elif exit_kind == 'function':
                                    if not proc_type_lower.startswith('function'):
                                        mismatch = True
                                elif exit_kind == 'property':
                                    if not proc_type_lower.startswith('property'):
                                        mismatch = True

                                if mismatch:
                                    self.errors.append({
                                        "file": filename,
                                        "line": node.tokens[0].line,
                                        "message": f"Exit {node.tokens[1].value} not allowed in {proc_def.proc_type}."
                                    })

                # Check for Jump
                if self.is_unconditional_jump(node.tokens):
                    # Check if conditional (e.g. "If x Then Exit Sub" split by colon)
                    is_conditional_jump = False
                    if prev_node and isinstance(prev_node, StatementNode):
                         # Check if on same line
                         if prev_node.tokens and node.tokens and prev_node.tokens[0].line == node.tokens[0].line:
                             # Check if prev starts with If
                             if prev_node.tokens[0].value.lower() == 'if':
                                 is_conditional_jump = True

                    if not is_conditional_jump:
                        unreachable = True
            
            elif isinstance(node, IfNode):
                # Analyze Condition
                self.analyze_statement(node.condition_tokens, scope, filename, context, with_stack)
                # Analyze True Block
                self.analyze_block(node.true_block, scope, filename, context, with_stack)
                # Analyze ElseIf Blocks
                for cond_tokens, block in node.else_blocks:
                     self.analyze_statement(cond_tokens, scope, filename, context, with_stack)
                     self.analyze_block(block, scope, filename, context, with_stack)
                # Analyze Else Block
                if node.else_block:
                     self.analyze_block(node.else_block, scope, filename, context, with_stack)
            
            elif isinstance(node, WithNode):
                if unreachable:
                     self.errors.append({
                         "file": filename,
                         "line": node.expr_tokens[0].line if node.expr_tokens else 0,
                         "rule_id": "VBA009",
                         "severity": "warning",
                         "message": f"Unreachable code detected (With block) in '{context}'."
                     })

                expr_type = self.resolve_expression_type(node.expr_tokens, scope, with_stack)
                new_stack = with_stack + [expr_type or 'Unknown']
                self.analyze_block(node.body, scope, filename, context, new_stack)

            elif isinstance(node, ForNode):
                # `For Each var In coll` and `For var = a To b [Step c]` —
                # analyze header tokens (including loop var, range, collection)
                # and recursively walk the body.
                if node.header_tokens:
                    self.analyze_statement(node.header_tokens, scope, filename, context, with_stack)

                var_name = node.var_token.value if node.var_token else None
                var_key = var_name.lower() if var_name else None

                # VBA173 — the same control variable is already driving an
                # enclosing For loop in this procedure.
                if var_key and var_key in getattr(self, '_for_var_stack', []):
                    self.errors.append({
                        "file": filename,
                        "line": getattr(node, 'line', 0) or (node.var_token.line if node.var_token else 0),
                        "rule_id": "VBA173",
                        "severity": "error",
                        "message": (
                            f"For loop control variable '{var_name}' is already in "
                            f"use by an enclosing loop in '{context}'."
                        ),
                    })

                # VBA174 — `For Each` control variable must be Variant or an
                # object type; primitive scalar counters are illegal.
                if node.kind == 'each' and var_name:
                    sym = scope.resolve(var_name)
                    vtype = sym.get('type') if sym else None
                    if vtype and self._is_clearly_scalar(vtype):
                        self.errors.append({
                            "file": filename,
                            "line": node.var_token.line,
                            "rule_id": "VBA174",
                            "severity": "error",
                            "message": (
                                f"For Each control variable '{var_name}' is typed "
                                f"'{vtype}'; it must be Variant or an object type."
                            ),
                        })

                # VBA172 — closing `Next x` names a variable other than this
                # loop's counter (skip the multi-close `Next j, i` form).
                if (node.kind == 'counter' and var_key and node.next_var_token
                        and not node.next_multi
                        and node.next_var_token.value.lower() != var_key):
                    self.errors.append({
                        "file": filename,
                        "line": node.next_var_token.line,
                        "rule_id": "VBA172",
                        "severity": "error",
                        "message": (
                            f"Next control variable '{node.next_var_token.value}' does "
                            f"not match the For counter '{var_name}' in '{context}'."
                        ),
                    })

                self._loop_stack.append('for')
                if var_key:
                    self._for_var_stack.append(var_key)
                self.analyze_block(node.body, scope, filename, context, with_stack)
                if var_key:
                    self._for_var_stack.pop()
                self._loop_stack.pop()

            elif isinstance(node, DoNode):
                # Do [While|Until cond] / Do … Loop While|Until / While … Wend
                # For bottom-tested loops the condition is evaluated *after*
                # the body — analyse the body first so any `Dim` inside the
                # body has registered the variable before the condition reads
                # it (regression for stdReg's `Loop While result = …` where
                # `result` is declared inline in the loop body).
                pos = getattr(node, 'condition_position', 'top')
                if pos == 'top' and node.condition_tokens:
                    self.analyze_statement(node.condition_tokens, scope, filename, context, with_stack)
                # `Exit Do` is only valid inside a real Do...Loop, not While...Wend.
                pushed_do = getattr(node, 'kind', 'do') == 'do'
                if pushed_do:
                    self._loop_stack.append('do')
                self.analyze_block(node.body, scope, filename, context, with_stack)
                if pushed_do:
                    self._loop_stack.pop()
                if pos == 'bottom' and node.condition_tokens:
                    self.analyze_statement(node.condition_tokens, scope, filename, context, with_stack)

            elif isinstance(node, SelectNode):
                # Selector expression
                if node.expr_tokens:
                    self.analyze_statement(node.expr_tokens, scope, filename, context, with_stack)
                # Per-Case body. Case header tokens (Is < 5, 1 To 10, value list, …)
                # are walked as expressions so type/identifier errors surface.
                for case in node.cases:
                    if not case.is_else and case.header_tokens:
                        self.analyze_statement(case.header_tokens, scope, filename, context, with_stack)
                    self.analyze_block(case.body, scope, filename, context, with_stack)

            elif isinstance(node, RedimNode):
                self._analyze_redim(node, scope, filename, context, with_stack)

            elif isinstance(node, EraseNode):
                self._analyze_erase(node, scope, filename, context, with_stack)

            prev_node = node

    def _analyze_redim(self, node, scope, filename, context, with_stack):
        """Validate ReDim targets:
        - The target must resolve to an array variable (declared `Dim x()`,
          `Dim x() As T`, or already-redimmed). Variant is permissive.
        - Dotted targets like `ReDim This.scopes(1 To N)` walk the member
          chain through `_resolve_lhs_type` (P2.6 chain walker) so UDT
          and class array members are recognised.
        - Dimension expressions must be analyzable.
        """
        for target in node.targets:
            # Backwards-compatible unpack: (name_token, dim_tokens, as_type)
            # or (name_token, dim_tokens, as_type, chain_tokens).
            chain_tokens = None
            if len(target) >= 4:
                name_token, dim_tokens, _as_type, chain_tokens = target[0], target[1], target[2], target[3]
            else:
                name_token, dim_tokens, _as_type = target[0], target[1], target[2]
            if name_token is None:
                continue
            name = name_token.value

            # Leading-dot target — `ReDim .points(1 To N)` inside a
            # `With X` block. We can't usefully walk the chain without
            # the with-stack anchor's actual UDT/Class, so stay silent
            # instead of emitting a phantom "Undefined identifier" on
            # the bare member name.
            if chain_tokens and chain_tokens[0].type == 'OPERATOR' and chain_tokens[0].value == '.':
                if dim_tokens:
                    self.analyze_statement(dim_tokens, scope, filename, context, with_stack)
                continue

            # Dotted target: walk the chain via the P2.6 LHS resolver.
            if chain_tokens and len(chain_tokens) > 1:
                resolved_type, resolved_kind = self._resolve_lhs_type(chain_tokens, scope)
                display = "".join(t.value for t in chain_tokens)
                if resolved_type is None:
                    # The chain root is unknown OR an intermediate hop
                    # is permissive (Object / Variant / unloaded ref) —
                    # in either case we can't usefully validate the
                    # array-ness, so stay silent rather than emit a
                    # false VBA101 on the leaf.
                    pass
                else:
                    target_type = (resolved_type or "").lower()
                    is_array = target_type.endswith("()") or "array" in target_type
                    is_variant = target_type in ("variant", "")
                    if not (is_array or is_variant):
                        self.errors.append({
                            "file": filename,
                            "line": name_token.line,
                            "rule_id": "VBA103",
                            "severity": "error",
                            "message": f"ReDim target '{display}' must be a dynamic array, got type '{resolved_type}' in '{context}'.",
                        })
                if dim_tokens:
                    self.analyze_statement(dim_tokens, scope, filename, context, with_stack)
                continue

            sym = scope.resolve(name)
            if sym is None:
                self.errors.append({
                    "file": filename,
                    "line": name_token.line,
                    "rule_id": "VBA101",
                    "severity": "error",
                    "message": f"Undefined identifier '{name}' in ReDim target inside '{context}'.",
                })
            else:
                target_type = (sym.get("type") or "").lower()
                kind = (sym.get("kind") or "").lower()
                is_array = target_type.endswith("()") or "array" in target_type
                is_variant = target_type in ("variant", "")
                if kind not in ("variable", "expression") and kind != "":
                    self.errors.append({
                        "file": filename,
                        "line": name_token.line,
                        "rule_id": "VBA102",
                        "severity": "error",
                        "message": f"ReDim target '{name}' is not an array variable (kind={sym.get('kind')}) in '{context}'.",
                    })
                elif not (is_array or is_variant):
                    self.errors.append({
                        "file": filename,
                        "line": name_token.line,
                        "rule_id": "VBA103",
                        "severity": "error",
                        "message": f"ReDim target '{name}' must be a dynamic array, got type '{sym.get('type')}' in '{context}'.",
                    })

            if dim_tokens:
                self.analyze_statement(dim_tokens, scope, filename, context, with_stack)

    def _analyze_erase(self, node, scope, filename, context, with_stack):
        """Validate Erase targets must be array variables (or Variant)."""
        for target in node.targets:
            # Backwards-compat: legacy parse_erase emitted a bare Token;
            # current parser emits a list (single-Token list for bare names,
            # longer for dotted chains like `Erase This.Leaf.points`).
            if isinstance(target, list):
                chain = target
            else:
                chain = [target]
            first_tok = chain[0]
            name = first_tok.value

            # Dotted target: walk the chain through `_resolve_lhs_type`.
            # An unresolvable hop returns (None, None) — stay silent in
            # that case (member chain may pass through a permissive bag
            # type like Variant/Object that we can't introspect).
            if len(chain) > 1:
                resolved_type, _kind = self._resolve_lhs_type(chain, scope)
                if resolved_type is None:
                    continue
                target_type = (resolved_type or "").lower()
                is_array = target_type.endswith("()") or "array" in target_type
                is_variant = target_type in ("variant", "")
                display = "".join(t.value for t in chain)
                if not (is_array or is_variant):
                    self.errors.append({
                        "file": filename,
                        "line": first_tok.line,
                        "rule_id": "VBA106",
                        "severity": "error",
                        "message": f"Erase target '{display}' must be an array, got type '{resolved_type}' in '{context}'.",
                    })
                continue

            sym = scope.resolve(name)
            if sym is None:
                self.errors.append({
                    "file": filename,
                    "line": first_tok.line,
                    "rule_id": "VBA104",
                    "severity": "error",
                    "message": f"Undefined identifier '{name}' in Erase target inside '{context}'.",
                })
                continue
            target_type = (sym.get("type") or "").lower()
            kind = (sym.get("kind") or "").lower()
            is_array = target_type.endswith("()") or "array" in target_type
            is_variant = target_type in ("variant", "")
            if kind not in ("variable", "expression") and kind != "":
                self.errors.append({
                    "file": filename,
                    "line": first_tok.line,
                    "rule_id": "VBA105",
                    "severity": "error",
                    "message": f"Erase target '{name}' is not an array variable (kind={sym.get('kind')}) in '{context}'.",
                })
            elif not (is_array or is_variant):
                self.errors.append({
                    "file": filename,
                    "line": first_tok.line,
                    "rule_id": "VBA106",
                    "severity": "error",
                    "message": f"Erase target '{name}' must be an array, got type '{sym.get('type')}' in '{context}'.",
                })

    # ---- Phase 2.4: Operator-type compatibility (literal-only) ----------

    # Strictly-arithmetic operators where a string operand is a guaranteed
    # type error. `+` is *not* in this list because VBA coerces it
    # bidirectionally between String and Number; `&` is string concat.
    _ARITH_OPS = {'-', '*', '/', '\\', '^'}
    _ARITH_KEYWORDS = {'mod'}

    def _validate_operator_types(self, tokens, filename, context):
        """Walk a statement's tokens and flag arithmetic operators that
        sit directly between a string literal and a numeric / string
        literal. The check is intentionally conservative — only literal
        operands are inspected, so it never fires on variables whose
        runtime type might be coerce-able.
        """
        if not tokens:
            return
        for i, tok in enumerate(tokens):
            if tok.type == 'OPERATOR' and tok.value in self._ARITH_OPS:
                self._check_arith_at(tokens, i, tok.value, filename, context)
            elif tok.type == 'IDENTIFIER' and tok.value.lower() in self._ARITH_KEYWORDS:
                self._check_arith_at(tokens, i, tok.value, filename, context)

    def _check_arith_at(self, tokens, idx, op_text, filename, context):
        # Find previous non-whitespace token (lhs) and next (rhs).
        lhs = tokens[idx - 1] if idx > 0 else None
        rhs = tokens[idx + 1] if idx + 1 < len(tokens) else None
        if not lhs or not rhs:
            return
        # Skip unary `-`: previous token is itself an operator/keyword.
        if op_text == '-':
            if lhs.type == 'OPERATOR' and lhs.value in {'(', ',', '=', '<', '>', '+', '-', '*', '/', '\\', '^', '<>', '<=', '>='}:
                return
            if lhs.type == 'IDENTIFIER' and lhs.value.lower() in {
                'and', 'or', 'not', 'xor', 'eqv', 'imp', 'mod', 'like',
                'is', 'then', 'to', 'step', 'in', 'else',
            }:
                return

        lhs_str = lhs.type == 'STRING'
        rhs_str = rhs.type == 'STRING'
        lhs_num = lhs.type in ('INTEGER', 'FLOAT', 'HEX', 'OCTAL')
        rhs_num = rhs.type in ('INTEGER', 'FLOAT', 'HEX', 'OCTAL')

        # Only fire when at least one side is a literal AND types disagree
        # in a way arithmetic cannot reconcile.
        if (lhs_str and rhs_num) or (lhs_num and rhs_str) or (lhs_str and rhs_str):
            offender = lhs if lhs_str else rhs
            self.errors.append({
                "file": filename,
                "line": offender.line,
                "rule_id": "VBA240",
                "severity": "error",
                "message": (
                    f"Type mismatch: arithmetic operator '{op_text}' between "
                    f"string literal and numeric literal in '{context}'. "
                    f"Use `&` for concatenation or `+` if you really mean "
                    f"VBA's bidirectional coercion."
                ),
            })

    # ----------------------------------------------------------------------

    # ---- Phase 2.5: Const-expression validation -------------------------

    # Built-in constant aliases shipped in the std model that *are* constants
    # even though their entry kind happens to be "String"/"Long"/etc.
    _CONST_KEYWORDS = {
        'true', 'false', 'nothing', 'empty', 'null',
    }
    _CONST_OPERATORS = {
        '+', '-', '*', '/', '\\', '^', '&', '=', '<>', '<', '>', '<=', '>=',
        '(', ')', '.', ',', ':',
    }
    _CONST_KEYWORD_OPS = {
        'and', 'or', 'not', 'xor', 'eqv', 'imp', 'mod', 'like',
        'is', 'true', 'false', 'nothing', 'empty', 'null',
    }

    def _is_constant_kind(self, sym):
        """A symbol counts as constant for the purposes of a Const RHS if
        it is itself another Const, an Enum member, or a literal-typed
        well-known global (vbCrLf, vbObjectError, …)."""
        if not sym:
            return False
        kind = (sym.get('kind') or '').lower()
        if kind in ('const', 'enumitem'):
            return True
        # Built-in scalar globals (vbCrLf, vbCritical, …) declared in
        # std_model with a scalar type and no callable kind. They behave
        # like constants in the VBA compiler.
        if kind in ('string', 'long', 'integer', 'boolean', 'double', 'variant'):
            return True
        return False

    def _validate_const_expression(self, expr_tokens, scope, filename, context, const_name):
        """Reject Const initialisers that reference variables or call functions."""
        if not expr_tokens:
            return
        i = 0
        n = len(expr_tokens)
        while i < n:
            tok = expr_tokens[i]
            if tok.type == 'IDENTIFIER':
                low = tok.value.lower().rstrip('$')
                # Reserved keywords that are valid in const expressions
                if low in self._CONST_KEYWORDS or low in self._CONST_KEYWORD_OPS:
                    i += 1
                    continue

                # Treat `<ident> ( ... )` as a function call → not constant.
                next_is_call = (
                    i + 1 < n
                    and expr_tokens[i + 1].type == 'OPERATOR'
                    and expr_tokens[i + 1].value == '('
                )
                sym = scope.resolve(tok.value)
                if next_is_call:
                    self.errors.append({
                        "file": filename,
                        "line": tok.line,
                        "rule_id": "VBA230",
                        "severity": "error",
                        "message": (
                            f"Const '{const_name}' initialiser calls '{tok.value}' — "
                            f"only constant expressions are allowed."
                        ),
                    })
                    return
                if sym and not self._is_constant_kind(sym):
                    self.errors.append({
                        "file": filename,
                        "line": tok.line,
                        "rule_id": "VBA231",
                        "severity": "error",
                        "message": (
                            f"Const '{const_name}' initialiser references non-constant "
                            f"'{tok.value}' (kind={sym.get('kind')})."
                        ),
                    })
                    return
                # Unknown identifier — already reported elsewhere; no extra noise.
            i += 1

    # ----------------------------------------------------------------------

    def _apply_def_type(self, name, current_type):
        """Resolve implicit DefInt/DefStr/… typing for an untyped variable
        whose first letter falls in the active per-module DefType map.
        """
        if current_type and current_type.lower() != 'variant':
            return current_type
        if not name or not self._current_def_type_map:
            return current_type
        first = name[0].lower()
        return self._current_def_type_map.get(first, current_type)

    def process_dim(self, tokens, scope, filename, context, with_stack):
        # Simplified Dim parser
        is_const = bool(tokens) and tokens[0].value.lower() == 'const'
        symbol_kind = 'Const' if is_const else 'Variable'
        # Track whether the current name has been given an explicit `As` —
        # used so DefType only applies when typing was implicit.
        explicit_as = False
        iterator = iter(tokens)
        next(iterator) # Skip Dim

        current_name = None
        current_type = 'Variant'
        is_array = False
        
        tokens_list = list(iterator)
        i = 0
        while i < len(tokens_list):
            t = tokens_list[i]
            if t.type == 'IDENTIFIER':
                if t.value.lower() == 'as':
                    explicit_as = True
                    saw_new = False
                    i += 1
                    type_parts = []
                    while i < len(tokens_list):
                        if tokens_list[i].value.lower() == 'new':
                            saw_new = True
                            i += 1
                            continue
                        
                        if tokens_list[i].type == 'IDENTIFIER':
                            type_parts.append(tokens_list[i].value)
                            i += 1
                            if i < len(tokens_list) and tokens_list[i].value == '.':
                                type_parts.append('.')
                                i += 1
                            else:
                                break
                        else:
                            break
                    current_type = "".join(type_parts)
                    if is_array:
                        current_type += "()"

                    # Phase 1.6 — VBA120: resolve the declared local type.
                    self._check_type_known(current_type, t.line, filename, context)

                    # VBA270 — a type-declaration suffix on the name must agree
                    # with the `As` type (`Dim x% As Long` is illegal). Only the
                    # $ / % / @ suffixes survive lexing as part of the identifier.
                    if current_name and current_name[-1] in self._TYPE_SUFFIX_MAP:
                        sfx_type = self._TYPE_SUFFIX_MAP[current_name[-1]]
                        base_ct = current_type.replace('()', '').split('*')[0].strip().lower()
                        if base_ct and base_ct != sfx_type:
                            self.errors.append({
                                "file": filename, "line": t.line,
                                "rule_id": "VBA270", "severity": "error",
                                "message": (
                                    f"Type-declaration character on '{current_name}' "
                                    f"implies {sfx_type.capitalize()} but it is declared "
                                    f"As {current_type} in '{context}'."
                                ),
                            })

                    # VBA165 — `New` requires a creatable class; `As New
                    # <primitive/Object/Variant>` is illegal.
                    if saw_new:
                        base = current_type.replace('()', '').split('*')[0].strip()
                        if base.lower() in self._BUILTIN_TYPE_NAMES:
                            self.errors.append({
                                "file": filename, "line": t.line,
                                "rule_id": "VBA165", "severity": "error",
                                "message": (
                                    f"Invalid use of New keyword: '{base}' is not a "
                                    f"creatable class in '{context}'."
                                ),
                            })

                    # Phase 2.8 — Fixed-length String only valid at module
                    # level (or inside UDTs). `Dim s As String * 10` inside
                    # a procedure is a hard VBA compile error.
                    if (
                        current_type.lower() == "string"
                        and i < len(tokens_list)
                        and tokens_list[i].type == 'OPERATOR'
                        and tokens_list[i].value == '*'
                    ):
                        line = tokens_list[i].line
                        # Consume `* <length>` regardless so we don't leak
                        # tokens into the next iteration.
                        i += 1
                        if i < len(tokens_list):
                            i += 1
                        self.errors.append({
                            "file": filename,
                            "line": line,
                            "rule_id": "VBA250",
                            "severity": "error",
                            "message": (
                                f"Fixed-length String declaration `As String * N` "
                                f"is not allowed at procedure level (only in modules "
                                f"or UDTs) for '{current_name}' in '{context}'."
                            ),
                        })

                    # If `=` follows the As-type, leave the actual define to
                    # the `=` branch so the assignment expression is analyzed
                    # (and so Const-expression validation can fire).
                    next_is_eq = (
                        i < len(tokens_list)
                        and tokens_list[i].type == 'OPERATOR'
                        and tokens_list[i].value == '='
                    )

                    if current_name and not next_is_eq:
                        if current_name.lower() in scope.symbols:
                            self.errors.append({
                                "file": filename,
                                "line": tokens[0].line,
                                "message": f"Duplicate declaration of identifier '{current_name}' in current scope."
                            })
                        else:
                            scope.define(current_name, current_type, symbol_kind)
                        current_name = None
                        current_type = 'Variant'
                        is_array = False
                else:
                    if current_name:
                        # Implicit Variant definition for the previous variable
                        t_type = "Variant"
                        if not explicit_as:
                            t_type = self._apply_def_type(current_name, t_type)
                        if is_array: t_type += "()"
                        scope.define(current_name, t_type, symbol_kind)
                        is_array = False # Reset for next

                    current_name = t.value
                    explicit_as = False
                    i += 1
                    
                    # Check for Array ()
                    if i < len(tokens_list) and tokens_list[i].value == '(':
                        is_array = True
                        depth = 1
                        i += 1
                        while i < len(tokens_list) and depth > 0:
                            if tokens_list[i].value == '(': depth += 1
                            elif tokens_list[i].value == ')': depth -= 1
                            i += 1

            elif t.type == 'OPERATOR' and t.value == '=':
                # Initialization
                if current_name:
                     if current_name.lower() in scope.symbols:
                         self.errors.append({
                             "file": filename,
                             "line": tokens[0].line,
                             "message": f"Duplicate declaration of identifier '{current_name}' in current scope."
                         })
                     else:
                         t_type = current_type
                         if not explicit_as:
                             t_type = self._apply_def_type(current_name, t_type)
                         if is_array and not t_type.endswith('()'): t_type += "()"
                         scope.define(current_name, t_type, symbol_kind)

                     # We defined the variable, now let's analyze the assignment expression
                     # We need to find where the expression ends (at comma or end of tokens)
                     expr_start = i + 1
                     expr_end = len(tokens_list)

                     # Look ahead for comma
                     depth_parens = 0
                     for k in range(expr_start, len(tokens_list)):
                          if tokens_list[k].value == '(': depth_parens += 1
                          elif tokens_list[k].value == ')': depth_parens -= 1
                          elif tokens_list[k].value == ',' and depth_parens == 0:
                               expr_end = k
                               break

                     expr_tokens = tokens_list[expr_start:expr_end]
                     self.analyze_statement(expr_tokens, scope, filename, context, with_stack)
                     if is_const:
                         self._validate_const_expression(expr_tokens, scope, filename, context, current_name)

                     current_name = None
                     current_type = 'Variant'
                     is_array = False
                     explicit_as = False

                     i = expr_end
                else:
                     i += 1

            elif t.value == ',':
                if current_name:
                    if current_name.lower() in scope.symbols:
                        self.errors.append({
                            "file": filename,
                            "line": tokens[0].line,
                            "message": f"Duplicate declaration of identifier '{current_name}' in current scope."
                        })
                    else:
                        t_type = current_type
                        if not explicit_as:
                            t_type = self._apply_def_type(current_name, t_type)
                        if is_array and not t_type.endswith('()'): t_type += "()"
                        scope.define(current_name, t_type, symbol_kind)
                    current_name = None
                    current_type = 'Variant'
                    is_array = False
                    explicit_as = False
                i += 1
            else:
                i += 1

        if current_name:
             if current_name.lower() in scope.symbols:
                 self.errors.append({
                     "file": filename,
                     "line": tokens[0].line,
                     "message": f"Duplicate declaration of identifier '{current_name}' in current scope."
                 })
             else:
                 t_type = current_type
                 if not explicit_as:
                     t_type = self._apply_def_type(current_name, t_type)
                 if is_array and not t_type.endswith('()'): t_type += "()"
                 scope.define(current_name, t_type, symbol_kind)

    def resolve_expression_type(self, tokens, scope, with_stack):
        return self.analyze_statement(tokens, scope, "", "", with_stack, report_errors=False)

    def analyze_statement(self, tokens, scope, filename, context, with_stack, report_errors=True):
        # `On Error GoTo <label>` / `On Error Resume Next` / `On Error
        # GoTo 0|-1` — fully validated by `_validate_jump_target`. Don't
        # re-walk the token stream as a value-bearing expression or
        # `Error` would be treated as a bare identifier (which it is in
        # other VBA contexts — see the comment in `analyze_expression_info`'s
        # KEYWORDS set — but here it's part of the On-Error syntax).
        if (
            tokens and tokens[0].type == 'IDENTIFIER' and tokens[0].value.lower() == 'on'
            and len(tokens) >= 2 and tokens[1].type == 'IDENTIFIER'
            and tokens[1].value.lower() == 'error'
        ):
            return None

        # `On <expr> GoTo lbl1, lbl2, ...` (computed GoTo). Analyse the
        # selector expression normally and skip the label list — those
        # identifiers are validated by `_validate_jump_target`, not as
        # value references (otherwise every label fires VBA001).
        if tokens and tokens[0].type == 'IDENTIFIER' and tokens[0].value.lower() == 'on':
            j = self._find_on_jump_keyword(tokens)
            if j is not None and j >= 2:
                self.analyze_expression_info(
                    tokens[1:j], scope, filename, context, with_stack,
                    report_errors=report_errors,
                )
                return None
        type_name, _, _ = self.analyze_expression_info(tokens, scope, filename, context, with_stack, report_errors=report_errors)
        return type_name

    def analyze_expression_info(self, tokens, scope, filename, context, with_stack, report_errors=True, allow_implicit_call=True):
        KEYWORDS = {
            'set', 'call', 'if', 'then', 'else', 'elseif', 'end', 'exit',
            # `error` is intentionally NOT here — VBA reserves it only in
            # the two-token forms `On Error ...` (handled by the dedicated
            # `on` + lookahead branch above analyze_statement) and `Error
            # <number>` raise-statement. Anywhere else it's a perfectly
            # valid identifier (and stdVBA's `stdAcc::AwaitForElement`
            # uses it as a local variable name).
            'on', 'goto', 'resume', 'do', 'loop', 'while', 'wend',
            'for', 'next', 'select', 'case', 'with', 'to', 'step', 'in',
            'byval', 'byref', 'optional', 'paramarray', 'true', 'false',
            'nothing', 'empty', 'null',
            'not', 'each', 'sub', 'function', 'property', 'const', 'dim', 'as',
            'type', 'boolean', 'integer', 'string', 'variant', 'object',
            'byte', 'long', 'single', 'double', 'currency', 'date', 'decimal',
            'and', 'or', 'xor', 'is', 'like', 'typeof', 'mod', 'new', 'print',
            'open', 'close', 'input', 'output', 'append', 'binary', 'random',
            'get', 'put', 'let', 'stop', 'len', 'mid', 'redim', 'preserve', 'erase',
            'friend', 'event', 'implements', 'raiseevent', 'gosub', 'return',
            'lset', 'rset', 'addressof',
            'defbool', 'defbyte', 'defint', 'deflong', 'defcur', 'defsng',
            'defdbl', 'defdec', 'defdate', 'defstr', 'defobj', 'defvar'
        }

        i = 0
        last_resolved_type = None
        last_resolved_kind = None
        last_resolved_name = None
        last_resolved_symbol = None
        expect_member = False
        prev_keyword = None
        
        implied_type = None

        while i < len(tokens):
            token = tokens[i]
            
            # Check for Implicit Call (Sub style)
            # Only legal at statement level — inside an expression (e.g.
            # `Round(Timer - t, 3)`) the comma is a Round-arg separator,
            # not a Timer-arg separator. Disabling the check in recursive
            # paren-contents avoids attributing `(Timer - t, 3)` to Timer.
            if allow_implicit_call and last_resolved_kind in ('Function', 'Procedure', 'Global') and last_resolved_name and not expect_member:
                 is_arg_start = False
                 if token.type in ('STRING', 'INTEGER', 'FLOAT'): is_arg_start = True
                 elif token.type == 'IDENTIFIER' and token.value.lower() not in KEYWORDS: is_arg_start = True
                 elif token.type == 'OPERATOR' and token.value.lower() in ('-', 'not', 'byval', 'byref'): is_arg_start = True

                 if is_arg_start:
                      arg_tokens = tokens[i:]
                      if report_errors and last_resolved_symbol:
                           self.validate_signature(last_resolved_name, last_resolved_symbol, arg_tokens, filename, token.line, context, scope, with_stack)

                      args = self.split_args(arg_tokens)
                      for arg in args:
                          self.analyze_expression_info(arg, scope, filename, context, with_stack, report_errors=report_errors, allow_implicit_call=False)
                      break
            
            if token.type == 'OPERATOR':
                val = token.value.lower()
                if val not in ('.', '!', '(', ')', ','):
                    if val == '&':
                        implied_type = 'String'
                    elif val in ('=', '<>', '<', '>', '<=', '>=', 'like', 'is'):
                         implied_type = 'Boolean'
                    elif val in ('+', '-', '*', '/', '^', 'mod', '\\'):
                         if last_resolved_type not in ('Integer', 'Long', 'Single', 'Double', 'Currency', 'Byte'):
                              implied_type = 'Double'
                         else:
                              implied_type = last_resolved_type

                    last_resolved_symbol = None
                    last_resolved_name = None
                    last_resolved_type = None
                    last_resolved_kind = None
                    expect_member = False
                    prev_keyword = None
                    i += 1
                    continue

            if token.type == 'IDENTIFIER':
                name = token.value
                last_resolved_name = name
                
                if name.lower() in KEYWORDS and not expect_member:
                    prev_keyword = name.lower()
                    last_resolved_type = None
                    last_resolved_kind = None
                    last_resolved_symbol = None
                    i += 1
                    continue
                
                # Check for Label Definition "Label:" or Named Argument "Arg:="
                if i + 1 < len(tokens) and tokens[i+1].type == 'OPERATOR':
                    if tokens[i+1].value == ':':
                        # A `Label:` only exists at the START of a statement
                        # (i == 0). In `Sub S(): Dim x: x = Foo: End Sub`
                        # the parser hands us colon-separated statements
                        # whose last tokens include a trailing `:` (the
                        # statement separator). Treating every `IDENT :`
                        # in the stream as a label-definition makes the
                        # analyser skip identifier resolution of `Foo`,
                        # which is the silent-failure UAT §3 caught.
                        # Real label: only when i == 0 AND there are no
                        # other tokens after the colon, OR the trailing
                        # tokens are themselves just more colons.
                        is_label_position = (i == 0)
                        if is_label_position:
                            i += 2
                            prev_keyword = None
                            last_resolved_type = None
                            last_resolved_kind = None
                            continue
                        # Statement-separator colon — drop it and continue
                        # resolving the previous identifier normally. We
                        # do NOT advance past the colon yet; the next
                        # iteration will handle it as an OPERATOR.
                    if tokens[i+1].value == ':=':
                        # Named Argument - skip it and the operator
                        i += 2
                        last_resolved_type = None
                        last_resolved_kind = None
                        continue

                if prev_keyword in ('goto', 'resume', 'gosub'):
                    prev_keyword = None
                    i += 1
                    continue

                if expect_member and last_resolved_type:
                    current_module_name = next((m.name for m in self.modules if m.filename == filename), None)
                    member_type, member_kind, member_extra = self.resolve_member(last_resolved_type, name, current_module_name) or (None, None, None)
                    if not member_type:
                        if not self._is_permissive_chain_type(last_resolved_type):

                            if report_errors:
                                self.errors.append({
                                    "file": filename,
                                    "line": token.line,
                                    "message": f"Member '{name}' not found in type '{last_resolved_type}' inside '{context}'."
                                })
                    last_resolved_type = member_type or 'Unknown'
                    # Propagate Kind if possible, or assume Unknown
                    # If resolved to a Variable member, it's a Variable.
                    last_resolved_kind = member_kind or 'Unknown'

                    # Store member metadata in a transient symbol dict for validation
                    if member_extra:
                        last_resolved_symbol = {"type": last_resolved_type, "kind": last_resolved_kind, "extra": member_extra}
                    else:
                        last_resolved_symbol = None

                    expect_member = False
                else:
                    sym = scope.resolve(name)
                    last_resolved_symbol = sym
                    if not sym:
                        # Dynamic ENUM Lookup
                        enum_val = self.resolve_enum(name)
                        if enum_val is not None:
                             last_resolved_type = 'Long'
                             last_resolved_kind = 'EnumItem'
                        else:
                            last_resolved_symbol = None
                            # HEURISTIC: If inside a Form, assume undefined identifier is an implicit Control
                            is_in_form = False
                            curr = scope
                            while curr:
                                if curr.scope_type == 'Form':
                                    is_in_form = True
                                    break
                                curr = curr.parent

                            if is_in_form:
                                last_resolved_type = 'Object'
                                last_resolved_kind = 'Control'
                            else:
                                if report_errors:
                                    self.errors.append({
                                        "file": filename,
                                        "line": token.line,
                                        "message": f"Undefined identifier '{name}' in '{context}'."
                                    })
                                last_resolved_type = 'Unknown'
                                last_resolved_kind = 'Unknown'
                    if not sym:
                        # ... (existing fallback logic)
                        # ...
                        # ...
                                last_resolved_type = 'Unknown'
                                last_resolved_kind = 'Unknown'
                    else:
                        last_resolved_type = sym['type']
                        last_resolved_kind = sym.get('kind', 'Unknown')
                        last_resolved_symbol = sym
                
                prev_keyword = None
                i += 1
            
            elif token.type == 'OPERATOR':
                if token.value == '.':
                    if last_resolved_type is None:
                        if with_stack:
                            last_resolved_type = with_stack[-1]
                            # Assume With object is a variable reference?
                            # Usually yes (Reference to object).
                            last_resolved_kind = 'Variable'
                        else:
                            if report_errors:
                                self.errors.append({
                                    "file": filename,
                                    "line": token.line,
                                    "message": f"Invalid or unexpected . reference without With block in '{context}'."
                                })
                            last_resolved_type = 'Unknown'
                            last_resolved_kind = 'Unknown'
                    expect_member = True
                    i += 1
                elif token.value == '(':
                    depth = 1
                    i += 1
                    start_index = i
                    while i < len(tokens) and depth > 0:
                        if tokens[i].type == 'OPERATOR':
                            if tokens[i].value == '(': depth += 1
                            elif tokens[i].value == ')': depth -= 1
                        i += 1
                    
                    end_index = i - 1
                    
                    # Check if we are invoking something that isn't callable
                    # Must be a Variable (not a Procedure/Function) AND have a non-array scalar type
                    if last_resolved_kind == 'Variable' and last_resolved_type in ('String', 'Integer', 'Long', 'Boolean', 'Double', 'Currency', 'Date', 'Single', 'Byte'):
                         # We allow 'Unknown', 'Variant', 'Object', 'LongPtr', 'Any' and user defined types.
                         # Also ignore if prev_keyword was AddressOf or similar
                         if report_errors and prev_keyword != 'addressof':
                            self.errors.append({
                                "file": filename,
                                "line": token.line,
                                "message": f"Expected Array or Procedure, got variable '{prev_keyword or 'Unknown'}' of type '{last_resolved_type}'."
                            })

                    # Recursively analyze the content inside the parentheses
                    inner_type = 'Variant'
                    inferred_ret_type = None

                    sub_tokens = []
                    if end_index > start_index:
                        sub_tokens = tokens[start_index : end_index]

                    # Hook for CreateObject("ProgID") / GetObject(, "ProgID")
                    # — infer the return type from the ProgID string. Try
                    # both the full ProgID (matches namespaced classes
                    # like `Shell.Application`) and the bare last segment
                    # (matches `Dictionary` from `Scripting.Dictionary`).
                    if last_resolved_name and last_resolved_name.lower() in ('createobject', 'getobject'):
                            prog_id_tok = None
                            if last_resolved_name.lower() == 'createobject':
                                if len(sub_tokens) > 0 and sub_tokens[0].type == 'STRING':
                                    prog_id_tok = sub_tokens[0]
                            else:
                                # GetObject(pathname, [class]) — the ProgID,
                                # if any, is the second positional argument.
                                args_split = self.split_args(sub_tokens)
                                if len(args_split) >= 2 and args_split[1] and args_split[1][0].type == 'STRING':
                                    prog_id_tok = args_split[1][0]
                            if prog_id_tok is not None:
                                prog_id = prog_id_tok.value.strip('"')
                                bare = prog_id.split('.')[-1] if '.' in prog_id else None
                                for candidate in (prog_id, bare):
                                    if candidate and self.config.get_class(candidate):
                                        inferred_ret_type = candidate
                                        break

                    # Hook for Signature Validation.
                    #
                    # Skip when this looks like a default-property `Item`
                    # call: `obj.children(idx)` where `children` is a
                    # 0-arg Property returning a collection-like type
                    # (Collection / Dictionary / anything with `.Item`).
                    # VBA implicitly rewrites that to
                    # `obj.children.Item(idx)`, so the args belong to
                    # `Item`, not the property. Without this exemption
                    # we get a phantom "Expected at most 0, got 1" on
                    # idiomatic library code (stdAcc.cls).
                    is_default_property_item_call = (
                        sub_tokens
                        and isinstance(last_resolved_symbol, dict)
                        and last_resolved_type
                        and last_resolved_type not in ('Unknown', 'Variant', 'Object')
                        and self.resolve_member(last_resolved_type, 'Item') is not None
                    )
                    if report_errors and last_resolved_symbol and not is_default_property_item_call:
                            self.validate_signature(last_resolved_name, last_resolved_symbol, sub_tokens, filename, token.line, context, scope, with_stack)

                    if sub_tokens:
                        # Inner expression analysis
                        inner_type, _, _ = self.analyze_expression_info(sub_tokens, scope, filename, context, with_stack, report_errors=report_errors, allow_implicit_call=False)
                    
                    # Determine result type
                    if last_resolved_type is None:
                        # Grouping (Expression)
                        last_resolved_type = inner_type
                        last_resolved_kind = 'Expression' # Grouping is always an Expression
                    else:
                        # Function/Array Call
                        if inferred_ret_type:
                            last_resolved_type = inferred_ret_type
                            last_resolved_kind = 'Expression'
                        elif last_resolved_type.endswith('()'):
                             # `arr()` with empty parens is VBA's explicit
                             # pass-whole-array syntax, not an indexed access —
                             # the type stays the array, not the element. Real
                             # element access always has at least one index
                             # token inside the parens.
                             if not sub_tokens:
                                 # last_resolved_type and kind unchanged
                                 pass
                             else:
                                 last_resolved_type = last_resolved_type[:-2]
                                 # Array Element Access -> Preserves Variable-ness if base was Variable
                                 if last_resolved_kind == 'Variable':
                                     last_resolved_kind = 'Variable'
                                 else:
                                     last_resolved_kind = 'Expression'
                        else:
                            # Default Property Logic (e.g. Selection(1) -> Selection.Item(1))
                            # If the type is an object and has an "Item" member, resolve to that type.
                            current_module_name = next((m.name for m in self.modules if m.filename == filename), None)
                            item_type, item_kind, item_extra = self.resolve_member(last_resolved_type, 'Item', current_module_name) or (None, None, None)
                            if item_type:
                                last_resolved_type = item_type
                                last_resolved_kind = item_kind or 'Unknown'
                            else:
                                # Fallback or error
                                last_resolved_kind = 'Unknown'

                    expect_member = False
                else:
                    last_resolved_type = None
                    last_resolved_kind = None
                    last_resolved_symbol = None
                    expect_member = False
                    prev_keyword = None
                    i += 1
            
            elif token.type == 'OPERATOR':
                val = token.value.lower()
                if val == '&':
                    last_resolved_type = 'String'
                    last_resolved_kind = 'Expression'
                elif val in ('=', '<>', '<', '>', '<=', '>=', 'like', 'is'):
                    last_resolved_type = 'Boolean'
                    last_resolved_kind = 'Expression'
                elif val in ('+', '-', '*', '/', '^', 'mod', '\\'):
                     # Preserve numeric type if possible, else Double
                     if last_resolved_type not in ('Integer', 'Long', 'Single', 'Double', 'Currency', 'Byte'):
                         last_resolved_type = 'Double'
                     last_resolved_kind = 'Expression'
                else:
                    # Other operators (e.g. . ! which shouldn't be here if handled above?)
                    # . is handled specifically. ! might be here.
                    if val == '!': 
                        # Bang operator (Collection access)
                        # Not fully supported, treat as Unknown/Variant
                        pass
                
                last_resolved_symbol = None
                expect_member = False
                prev_keyword = None
                i += 1

            elif token.type in ('STRING', 'INTEGER', 'FLOAT'):
                last_resolved_type = None
                last_resolved_kind = 'Expression' # Literal
                last_resolved_symbol = None
                expect_member = False
                prev_keyword = None
                i += 1
            else:
                i += 1
        
        if implied_type:
            return implied_type, 'Expression', None
            
        return last_resolved_type, last_resolved_kind, last_resolved_symbol

    def split_args(self, tokens):
        args = []
        if not tokens: return args

        current_arg = []
        depth = 0
        for t in tokens:
            if t.value == '(':
                depth += 1
                current_arg.append(t)
            elif t.value == ')':
                depth -= 1
                current_arg.append(t)
            elif t.value == ',' and depth == 0:
                args.append(current_arg)
                current_arg = []
            else:
                current_arg.append(t)
        if current_arg or (args and tokens[-1].value == ','):
             args.append(current_arg)
        return args

    @staticmethod
    def _param_name(p):
        if isinstance(p, dict):
            return (p.get('name') or '').lower() or None
        return (getattr(p, 'name', '') or '').lower() or None

    @staticmethod
    def _param_is_paramarray(p):
        if isinstance(p, dict):
            return bool(p.get('is_paramarray', False))
        return bool(getattr(p, 'is_paramarray', False))

    def _validate_named_args(self, args, param_defs, name, filename, line, context,
                             is_source_proc=False):
        """VBA140/141/142 — validate `name:=value` call arguments.

        - VBA141 (duplicate) is purely syntactic.
        - VBA140 (unknown name) only fires for scanned-source procedures
          (`is_source_proc`) whose parameter names are authoritative. Host-
          model parameter names are advisory (they can abbreviate or rename
          the real API, e.g. `Range.Find(What:=…)`), so VBA140 is suppressed
          for them to avoid false positives.
        - VBA142 — a procedure with a ParamArray cannot be called with named
          arguments.
        """
        param_names = {self._param_name(p) for p in (param_defs or [])}
        param_names.discard(None)
        has_param_array = any(self._param_is_paramarray(p) for p in (param_defs or []))
        all_named = is_source_proc and bool(param_defs) and all(
            self._param_name(p) for p in param_defs)

        seen = set()
        any_named = False
        for arg in args:
            if (len(arg) >= 2 and arg[0].type == 'IDENTIFIER'
                    and arg[1].type == 'OPERATOR' and arg[1].value == ':='):
                any_named = True
                nm = arg[0].value
                low = nm.lower()
                if low in seen:
                    self.errors.append({
                        "file": filename, "line": line,
                        "rule_id": "VBA141", "severity": "error",
                        "message": f"Named argument '{nm}' specified more than once in call to '{name}'.",
                    })
                seen.add(low)
                if all_named and low not in param_names:
                    self.errors.append({
                        "file": filename, "line": line,
                        "rule_id": "VBA140", "severity": "error",
                        "message": f"Named argument '{nm}' not found in '{name}'.",
                    })

        if any_named and has_param_array:
            self.errors.append({
                "file": filename, "line": line,
                "rule_id": "VBA142", "severity": "error",
                "message": (
                    f"'{name}' has a ParamArray and cannot be called with named arguments."
                ),
            })

    def validate_signature(self, name, symbol, arg_tokens, filename, line, context, scope=None, with_stack=None):
        extra = symbol.get('extra')
        if not extra: return

        if scope is None: scope = self.global_scope
        if with_stack is None: with_stack = []

        args = self.split_args(arg_tokens)
        arg_count = len(args)

        min_args = 0
        max_args = 999
        param_defs = []

        if isinstance(extra, ProcedureNode):
             min_args = 0
             max_args = len(extra.args)
             param_defs = extra.args
             has_param_array = False

             for arg in extra.args:
                  if getattr(arg, 'is_paramarray', False):
                       has_param_array = True

                  if not arg.is_optional and not getattr(arg, 'is_paramarray', False):
                       min_args += 1

             if has_param_array:
                  max_args = 999

             # For Property Let/Set, the last argument is the value being assigned (RHS),
             # so it doesn't appear in the argument list (parentheses).
             if 'let' in extra.proc_type.lower() or 'set' in extra.proc_type.lower():
                 if max_args > 0: max_args -= 1
                 if min_args > 0: min_args -= 1

        elif isinstance(extra, dict):
             min_args = extra.get('min_args', 0)
             max_args = extra.get('max_args', 999)
             if 'args' in extra:
                 param_defs = extra['args']

        # Phase 2.9 — named-argument validation (VBA140/141/142). VBA140
        # (unknown name) only for scanned-source procedures whose parameter
        # names are authoritative; host-model param names are advisory.
        self._validate_named_args(
            args, param_defs, name, filename, line, context,
            is_source_proc=isinstance(extra, ProcedureNode))

        if arg_count < min_args:
             self.errors.append({
                 "file": filename,
                 "line": line,
                 "message": f"Argument count mismatch for '{name}': Expected at least {min_args}, got {arg_count}."
             })
        if arg_count > max_args:
             self.errors.append({
                 "file": filename,
                 "line": line,
                 "message": f"Argument count mismatch for '{name}': Expected at most {max_args}, got {arg_count}."
             })

        # Check ByRef Type Mismatch
        if param_defs:
            for i, arg_tokens_list in enumerate(args):
                if i >= len(param_defs):
                    # Could be ParamArray. If last param is ParamArray, usage is valid.
                    # We can check param_defs[-1].is_paramarray
                    last_param = param_defs[-1]
                    is_last_pa = False
                    if isinstance(last_param, dict):
                        is_last_pa = last_param.get('is_paramarray', False)
                    else:
                        is_last_pa = getattr(last_param, 'is_paramarray', False)

                    if is_last_pa:
                        continue
                    break

                param = param_defs[i]

                # Check for ParamArray
                is_pa = False
                if isinstance(param, dict):
                    is_pa = param.get('is_paramarray', False)
                else:
                    is_pa = getattr(param, 'is_paramarray', False)

                if is_pa: continue

                # Analyze Argument
                arg_type, arg_kind, _ = self.analyze_expression_info(arg_tokens_list, scope, filename, context, with_stack, report_errors=False)

                # Get Param Info
                mech = 'ByRef'
                param_type = 'Variant'
                param_name = 'Unknown'

                if isinstance(param, dict):
                    mech = param.get('mechanism', 'ByRef')
                    param_type = param.get('type', 'Variant')
                    param_name = param.get('name', 'Unknown')
                else:
                    mech = getattr(param, 'mechanism', 'ByRef')
                    param_type = param.type_name
                    param_name = param.name

                # `As Any` (and the array form `As Any()`) is VBA's
                # universally-compatible sentinel in `Declare` statements —
                # callers may pass any concrete type. Short-circuit the
                # ByRef strict-type check for either form.
                if param_type.lower() in ('any', 'any()'):
                    continue

                if mech == 'ByRef':
                    # Only check if argument is a Variable (L-Value)
                    if arg_kind == 'Variable':
                         # Allow Variant to accept any type
                         if param_type.lower() == 'variant':
                             continue

                         # Strict Type Equality
                         # Ignore if types are Unknown
                         if param_type != 'Unknown' and arg_type != 'Unknown':
                             p_lower = param_type.lower()
                             a_lower = arg_type.lower()
                             
                             compatible = False
                             if p_lower == a_lower:
                                 compatible = True
                             elif p_lower == 'object' and a_lower not in ('string', 'integer', 'long', 'boolean', 'double', 'single', 'currency', 'date', 'byte', 'variant'):
                                 compatible = True
                             elif p_lower.endswith('.' + a_lower):
                                 compatible = True
                             elif a_lower.endswith('.' + p_lower):
                                 compatible = True
                             
                             # RELAXATION:
                             # If the argument is a generic Object, we allow it.
                             # VBA treats 'Object' as a dynamic type that can be passed to strong-typed ByRef parameters
                             # (often resolving at runtime or via implicit temp variables).
                             # This prevents false positives for code like:
                             #    Sub Fill(s As Shape) ... End Sub
                             #    Dim o As Object: Fill o
                             if not compatible:
                                 if a_lower == 'object':
                                     compatible = True

                             # RELAXATION:
                             # `Variant` can hold any object reference,
                             # so passing a Variant to an `Object`/class
                             # ByRef parameter is valid VBA — the runtime
                             # unwraps the contained reference. Symmetric
                             # to the Object→class relaxation above.
                             if not compatible and a_lower == 'variant':
                                 if p_lower == 'object' or self._is_class_type(param_type):
                                     compatible = True

                             # RELAXATION:
                             # Every VBA class implements `IUnknown`
                             # implicitly. A class-typed argument is
                             # therefore compatible with a ByRef IUnknown
                             # parameter (the canonical stdCOM-style
                             # COM-interop pattern).
                             if not compatible and p_lower == 'iunknown' and self._is_class_type(arg_type):
                                 compatible = True

                             # RELAXATION:
                             # Enums are Long under the hood. Real VBA
                             # accepts both directions:
                             #   `Dim x As MyEnum: Inner x`  → `ByRef p As Long`
                             #   `Dim x As Long:   Inner x`  → `ByRef p As MyEnum`
                             # Treat any Enum<->Long pairing as compatible.
                             if not compatible:
                                 enum_long_compat = (
                                     (p_lower == 'long' and self._is_enum_type(arg_type)) or
                                     (a_lower == 'long' and self._is_enum_type(param_type))
                                 )
                                 if enum_long_compat:
                                     compatible = True

                             if not compatible:
                                 self.errors.append({
                                     "file": filename,
                                     "line": line,
                                     "message": f"ByRef argument type mismatch. Parameter '{param_name}' expects '{param_type}', but got variable of type '{arg_type}'."
                                 })

    # Names a member-chain hop must be silent about. Either intentionally
    # permissive (Object / Variant) or carrying no resolvable member info
    # (procedure-kind literals from host models that lack a `returns`
    # field; qualified types whose owning library isn't loaded).
    _PERMISSIVE_PRIMITIVES = {
        'object', 'variant', 'unknown', 'control', 'form',
        # Procedure-kind names that some host models emit as the type
        # when the real return-type is unknown. Treating them as types
        # would produce nonsense like "Member 'X' not found in type 'Sub'".
        'sub', 'function', 'property', 'event',
    }

    def _is_permissive_chain_type(self, type_name):
        """A chain hop on this type cannot produce a useful diagnostic.

        Skips:
          * primitive bag-types we never had information about (Object,
            Variant, …),
          * procedure-kind literals smuggled in as types when host
            models lack a `returns` field (Sub / Function / Property),
          * qualified types from libraries the analyser doesn't have
            loaded (`ComctlLib.Node`, `MSForms.Control`, …) — i.e. the
            namespace prefix isn't a known Project class / module and
            isn't in the host model.
        """
        if not type_name:
            return True
        low = type_name.lower()
        if low in self._PERMISSIVE_PRIMITIVES:
            return True
        # Array-type residual like "Cell()" — the `(...)` consume step
        # should have stripped the suffix, so seeing it here means the
        # caller is checking a type it can't navigate into.
        if low.endswith('()'):
            return True
        if '.' in type_name:
            prefix = type_name.split('.', 1)[0].lower()
            # If the namespace prefix is neither a known project module
            # nor a loaded reference / host class, we can't validate
            # members on it.
            known = False
            if any(m.name.lower() == prefix for m in self.modules):
                known = True
            elif prefix in self.reference_names:
                known = True
            elif self.config.get_class(prefix):
                known = True
            elif prefix in (self.config.object_model.get("classes", {}) or {}):
                known = True
            if not known:
                return True
        return False

    def resolve_member(self, type_name, member_name, current_module_name=None):
        return self._resolve_member_internal(type_name, member_name, current_module_name)

    def _resolve_member_internal(self, type_name, member_name, current_module_name=None):
        # Handle Qualified Types
        if '.' in type_name:
             simple_name = type_name.split('.')[-1]
             res = self._resolve_member_base(type_name, member_name, current_module_name)
             if res: return res
             res = self._resolve_member_base(simple_name, member_name, current_module_name)
             if res: return res
             return None
        
        return self._resolve_member_base(type_name, member_name, current_module_name)

    def _resolve_member_base(self, type_name, member_name, current_module_name=None):
        # 1. Check UDTs (Local types)
        if type_name.lower() in self.udts:
            udt = self.udts[type_name.lower()]
            for m in udt.members:
                if m.name.lower() == member_name.lower():
                    return m.type_name, 'Variable', None
        
        # 2. Check Project Modules & Classes (Source Code)
        # PRIORITIZED: If type_name matches a Project Module/Class, search strictly within it.
        # This prevents masking "Member Not Found" errors by falling back to globals/libs.
        found_module_match = False
        for mod in self.modules:
            if mod.name.lower() == type_name.lower():
                found_module_match = True
                
                is_local = current_module_name and current_module_name.lower() == mod.name.lower()

                # Check Variables
                for v in mod.variables:
                    allowed_scopes = ('public', 'global', 'friend')
                    if is_local:
                        allowed_scopes = ('public', 'global', 'friend', 'private', 'dim')
                    if v.name.lower() == member_name.lower() and v.scope.lower() in allowed_scopes:
                        return v.type_name, 'Variable', None

                # Check Procedures
                for p in mod.procedures:
                    allowed_scopes = ('public', 'friend')
                    if is_local:
                        allowed_scopes = ('public', 'friend', 'private')
                    if p.name.lower() == member_name.lower() and p.scope.lower() in allowed_scopes:
                         return p.return_type, 'Procedure', p

                # FALLBACK for Special Project Classes
                if mod.module_type == 'Form':
                     # Check 'UserForm' base class members
                     userform_cls = self.config.get_class('UserForm')
                     if userform_cls:
                         members = userform_cls.get('members', {})
                         for m_name, m_def in members.items():
                             if m_name.lower() == member_name.lower():
                                 t = m_def.get('type', 'Variant')
                                 return t, 'Expression', m_def

                     # Implicit Controls (Form Heuristic - Keep for compatibility unless causing issues)
                     # Since we can't always parse controls perfectly from .frm, assume other members are Controls
                     return 'Object', 'Variable', None

                if mod.name.lower() == 'thisdocument':
                     doc_cls = self.config.get_class('Document') or self.config.get_class('IVDocument')
                     if doc_cls:
                         members = doc_cls.get('members', {})
                         for m_name, m_def in members.items():
                             if m_name.lower() == member_name.lower():
                                 t = m_def.get('type', 'Variant')
                                 return t, 'Expression', m_def

        if found_module_match:
             # Strict Check: If we found the module/class but not the member, STOP.
             # Do not fall back to References or Globals.
             return None

        # 3. Check Library References (Global lookups)
        if type_name.lower() in self.reference_names:
             sym = self.global_scope.resolve(member_name)
             if sym:
                  return sym['type'], sym.get('kind', 'Expression'), sym.get('extra')
        
        # 4. Check Enums
        # If type_name matches a known Enum, check its members
        enums = self.config.object_model.get("enums", {})
        if type_name.lower() in enums:
            members = enums[type_name.lower()]
            # Case insensitive lookup
            for m in members:
                if m.lower() == member_name.lower():
                    return "Long", "EnumItem", None
            
            # Fallback: Check Global Scope (e.g. VisUnitCodes.visMillimeters where visMillimeters is Global)
            sym = self.global_scope.resolve(member_name)
            if sym:
                 return sym['type'], sym.get('kind', 'Expression'), sym.get('extra')
        
        # 5. Check Config Classes (Loaded from Model)
        cls_def = self.config.get_class(type_name)
        if cls_def:
            members = cls_def.get("members", {})
            for m_name, m_def in members.items():
                if m_name.lower() == member_name.lower():
                    raw_type = m_def.get("type", "Variant")
                    # Host models commonly encode the *kind* in `type`
                    # (Sub / Function / Property) and the actual return
                    # type in `returns`. If only `type` is present and
                    # it's a procedure-kind literal, the chain after
                    # this hop is unknowable — return Variant so the
                    # downstream walker treats the chain as permissive
                    # instead of asking "Member 'X' not found in type
                    # 'Function'".
                    if raw_type in ('Sub', 'Function', 'Property', 'Event'):
                        ret = m_def.get("returns")
                        t = ret if ret else 'Variant'
                        return t, 'Procedure', m_def
                    return raw_type, 'Expression', m_def

        return None

    def is_label(self, tokens):
        if len(tokens) >= 2:
            if tokens[0].type == 'IDENTIFIER' and tokens[1].type == 'OPERATOR' and tokens[1].value == ':':
                return True
        return False

    def is_ignorable(self, tokens):
        if not tokens: return True
        for t in tokens:
            if t.type != 'COMMENT':
                return False
        return True

    def is_unconditional_jump(self, tokens):
        if not tokens: return False
        t0 = tokens[0]
        val = t0.value.lower()

        if val == 'goto':
            # Check if strictly GoTo Label (Simple GoTo is 2 tokens, plus maybe a colon if parsed that way)
            if len(tokens) >= 2 and tokens[1].type == 'IDENTIFIER':
                 # Ensure it's not On Error GoTo ...
                 if len(tokens) == 2: return True
                 if len(tokens) == 3 and tokens[2].value == ':': return True

        if val == 'exit':
            if len(tokens) >= 2:
                t1 = tokens[1].value.lower()
                if t1 in ('sub', 'function', 'property'):
                    return True

        if val == 'end':
            if len(tokens) == 1: return True
            if len(tokens) == 2 and tokens[1].value == ':': return True

        return False

    def is_control_flow_boundary(self, tokens):
        if not tokens: return False
        val = tokens[0].value.lower()

        if val in ('else', 'elseif', 'next', 'loop', 'wend', 'case'):
            return True

        if val == 'end':
            if len(tokens) >= 2:
                 val2 = tokens[1].value.lower()
                 if val2 in ('if', 'select', 'with'):
                     return True

        return False

    def resolve_enum(self, name):
        # Look up enum constants
        enums = self.config.object_model.get("enums", {})
        for enum_name, members in enums.items():
            if name.lower() in [m.lower() for m in members.keys()]:
                return members.get(name) # Or handle case insensitive lookup properly
            # Case insensitive check
            for m_key, m_val in members.items():
                if m_key.lower() == name.lower():
                    return m_val
        return None

    def _is_class_type(self, type_name):
        """True if `type_name` names a user-defined Class (scanned `.cls`)
        or a Class from a loaded host model. Used to permit `Variant`→
        class and class→`IUnknown` ByRef compatibility, mirroring VBA's
        implicit IUnknown inheritance."""
        if not type_name or not isinstance(type_name, str):
            return False
        low = type_name.lower()
        # Scanned source: any module with module_type='Class' (or 'Form')
        # registers itself with that name; we cross-check via the existing
        # `modules` collection.
        for mod in self.modules:
            if mod.name.lower() == low and mod.module_type in ("Class", "Form"):
                return True
        # Host-model class.
        if self.config.get_class(type_name):
            return True
        return False

    def _is_enum_type(self, type_name):
        """True if `type_name` names an Enum — either one declared in the
        scanned VBA source (TypeNode with is_enum=True) or one provided by
        a loaded host model (`enums` section of object_model)."""
        if not type_name or not isinstance(type_name, str):
            return False
        low = type_name.lower()
        # Host-model enums.
        if low in self.config.object_model.get("enums", {}):
            return True
        # Source-declared enums (registered as UDTs with is_enum=True).
        udt = self.udts.get(low)
        if udt is not None and getattr(udt, "is_enum", False):
            return True
        return False
