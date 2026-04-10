import csv
import json
import re
import time
from dataclasses import dataclass, asdict
from typing import List, Set, Dict, Any, Optional

from PIL import Image
import mss
import pytesseract
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


# ── Window titles – adjust if your locale differs ──────────────────────────
CATIA_TITLE = "CATIA V5"
LANG_TITLE = "Language browser"
TYPE_DIALOG_TITLE = "Select a Type"

pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# ── Timing (seconds) ────────────────────────────────────────────────────────
TYPE_OPEN_DELAY = 0.8          # wait after clicking "..." for dialog to appear
TYPE_SELECT_DELAY = 0.5        # wait after confirming OK
TYPE_DIALOG_CLOSE_DELAY = 0.3  # wait after Cancel/ESC
FUNCTION_CLICK_DELAY = 0.25    # wait after clicking a function for tooltip to populate
OCR_RETRY_DELAY = 0.15         # extra wait before OCR retry

# ── Output files ────────────────────────────────────────────────────────────
OUTPUT_CSV = "catia_language_browser_signatures.csv"
PROGRESS_JSON = "catia_language_browser_progress.json"

# ── Set True to print every control found inside the type dialog (one-time
#    diagnostic run so you can see what class names CATIA actually uses) ────
DIAGNOSTIC_MODE = False

# ── OCR capture region (absolute screen pixels) ─────────────────────────────
# Tune these with the debug script if you change monitors or resolution
OCR_BBOX = {
    "left":   3300,
    "top":    1870,
    "width":  470,
    "height": 120,
}


@dataclass
class Row:
    object_type: str
    function_name: str
    signature: str


# ───────────────────────────────────────────────────────────────────────────
# Utility helpers
# ───────────────────────────────────────────────────────────────────────────

def safe_text(ctrl) -> str:
    try:
        t = ctrl.window_text()
        return t.strip() if t else ""
    except Exception:
        return ""


def is_separator(name: str) -> bool:
    s = name.strip()
    return s.startswith("--------") and s.endswith("--------")


def separator_inner(name: str) -> str:
    """Extract the type name from inside a separator, e.g. '---- Foo ----' -> 'Foo'."""
    return name.strip("-").strip()


# ───────────────────────────────────────────────────────────────────────────
# Window / control finders
# ───────────────────────────────────────────────────────────────────────────

def connect():
    lang = Desktop(backend="win32").window(title=LANG_TITLE)
    catia = Desktop(backend="win32").window(title_re=f".*{CATIA_TITLE}.*")
    lang.wait("visible", timeout=15)
    catia.wait("visible", timeout=15)
    return lang, catia


def get_functions_list(lang):
    return lang.child_window(title="FunctionsList", class_name="ListBox").wrapper_object()


def get_type_chooser_button(lang):
    """Return the '...' button that opens the type-selection dialog."""
    for c in lang.descendants():
        try:
            if safe_text(c) == "CATKweTypeChooserEditor":
                for child in c.children():
                    try:
                        if child.friendly_class_name() == "Button" and safe_text(child) == "...":
                            return child
                    except Exception:
                        pass
        except Exception:
            pass
    for c in lang.descendants():
        try:
            if c.friendly_class_name() == "Button" and safe_text(c) == "...":
                return c
        except Exception:
            pass
    raise RuntimeError("Could not find '...' type-chooser button in Language Browser")


def wait_for_type_dialog(timeout: float = 6.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            dlg = Desktop(backend="win32").window(title=TYPE_DIALOG_TITLE)
            if dlg.exists() and dlg.is_visible():
                return dlg
        except Exception:
            pass
        time.sleep(0.1)
    return None


def open_type_dialog(button):
    button.click_input()
    time.sleep(TYPE_OPEN_DELAY)
    dlg = wait_for_type_dialog(timeout=6.0)
    if dlg is None:
        raise RuntimeError("'Select a Type' dialog did not appear after clicking '...'")
    return dlg


def dump_dialog_controls(dialog):
    """Print every descendant control – used once in DIAGNOSTIC_MODE."""
    print("\n=== Dialog control dump ===")
    for c in dialog.descendants():
        try:
            print(f"  cls={c.friendly_class_name()!r:30s}  "
                  f"title={safe_text(c)!r:40s}  "
                  f"auto_id={getattr(c, 'automation_id', lambda: '')()!r}")
        except Exception as e:
            print(f"  [error reading control: {e}]")
    print("=== End dump ===\n")


# ── Generic finders that handle both ListBox and TreeView ───────────────────

_LIST_CLASSES = {"ListBox", "SysListView32", "ListView"}
_TREE_CLASSES = {"TreeView", "SysTreeView32"}
_RAW_LIST     = {"ListBox", "SysListView32"}
_RAW_TREE     = {"SysTreeView32", "TreeView"}


def find_list_or_tree(dialog):
    # Strategy 1: CATIA's type listbox has window title 'TypeList'
    try:
        ctrl = dialog.child_window(title="TypeList").wrapper_object()
        if ctrl:
            return ctrl, "list"
    except Exception:
        pass

    for c in dialog.descendants():
        try:
            w = c.wrapper_object()
        except Exception:
            continue
        try:
            cls = w.friendly_class_name()
            if cls in _LIST_CLASSES:
                return w, "list"
            if cls in _TREE_CLASSES:
                return w, "tree"
        except Exception:
            pass
        try:
            raw = w.class_name()
            if raw in _RAW_LIST:
                return w, "list"
            if raw in _RAW_TREE:
                return w, "tree"
        except Exception:
            pass

    return None, None


def find_first_edit(dialog):
    for c in dialog.descendants():
        try:
            if c.friendly_class_name() == "Edit":
                return c.wrapper_object()
        except Exception:
            pass
    return None


def find_button(dialog, caption: str):
    for c in dialog.descendants():
        try:
            if c.friendly_class_name() == "Button" and safe_text(c).lower() == caption.lower():
                return c.wrapper_object()
        except Exception:
            pass
    return None


# ── Dialog dismissal ────────────────────────────────────────────────────────

def close_type_dialog_cancel(dialog):
    try:
        btn = find_button(dialog, "Cancel")
        if btn:
            btn.click_input()
        else:
            try:
                dialog.set_focus()
            except Exception:
                pass
            send_keys("{ESC}")
    except Exception:
        pass
    time.sleep(TYPE_DIALOG_CLOSE_DELAY)


def confirm_type_dialog_ok(dialog):
    try:
        btn = find_button(dialog, "OK")
        if btn:
            btn.click_input()
        else:
            try:
                dialog.set_focus()
            except Exception:
                pass
            send_keys("{ENTER}")
    except Exception:
        pass
    time.sleep(TYPE_SELECT_DELAY)


# ───────────────────────────────────────────────────────────────────────────
# Collecting all type names from the dialog
# ───────────────────────────────────────────────────────────────────────────

def _items_from_control(ctrl, kind: str) -> List[str]:
    names: List[str] = []
    try:
        if kind == "list":
            names = [t.strip() for t in ctrl.item_texts() if t and t.strip()]
        elif kind == "tree":
            try:
                ctrl.expand_all()
                time.sleep(0.3)
            except Exception:
                pass
            try:
                root = ctrl.roots()
                stack = list(root)
                while stack:
                    node = stack.pop(0)
                    try:
                        text = node.text().strip()
                        if text:
                            names.append(text)
                    except Exception:
                        pass
                    try:
                        stack[:0] = list(node.children())
                    except Exception:
                        pass
            except Exception:
                try:
                    names = [t.strip() for t in ctrl.item_texts() if t and t.strip()]
                except Exception:
                    pass
    except Exception:
        pass
    return names


def get_all_types_from_dialog(button) -> List[str]:
    dlg = open_type_dialog(button)

    if DIAGNOSTIC_MODE:
        dump_dialog_controls(dlg)

    ctrl, kind = find_list_or_tree(dlg)
    if ctrl is None:
        close_type_dialog_cancel(dlg)
        raise RuntimeError(
            "Could not find a ListBox or TreeView inside 'Select a Type' dialog. "
            "Run with DIAGNOSTIC_MODE=True to see available controls."
        )

    print(f"[INFO] Type dialog uses control kind={kind!r}, class={ctrl.friendly_class_name()!r}")
    names = _items_from_control(ctrl, kind)

    seen: Set[str] = set()
    unique: List[str] = []
    for n in names:
        if n not in seen:
            seen.add(n)
            unique.append(n)

    close_type_dialog_cancel(dlg)
    print(f"[INFO] Found {len(unique)} unique types in dialog")
    return unique


# ───────────────────────────────────────────────────────────────────────────
# Selecting a specific type via the dialog
# ───────────────────────────────────────────────────────────────────────────

def _select_in_list(ctrl, target: str) -> bool:
    try:
        items = ctrl.item_texts()
    except Exception:
        return False
    for item in items:
        if item.strip().lower() == target.lower():
            try:
                ctrl.select(item.strip())
                return True
            except Exception:
                pass
    for item in items:
        if target.lower() in item.strip().lower():
            try:
                ctrl.select(item.strip())
                return True
            except Exception:
                pass
    return False


def _select_in_tree(ctrl, target: str) -> bool:
    try:
        root = ctrl.roots()
    except Exception:
        return False
    stack = list(root)
    while stack:
        node = stack.pop(0)
        try:
            text = node.text().strip()
            if text.lower() == target.lower():
                try:
                    node.ensure_visible()
                    node.click_input()
                    return True
                except Exception:
                    pass
        except Exception:
            pass
        try:
            stack[:0] = list(node.children())
        except Exception:
            pass
    return False


def select_type_via_dialog(button, target_type: str) -> bool:
    dlg = open_type_dialog(button)

    edit = find_first_edit(dlg)
    if edit:
        try:
            edit.set_edit_text(target_type)
            time.sleep(0.25)
        except Exception:
            try:
                edit.click_input()
                send_keys("^a{BACKSPACE}")
                send_keys(target_type, with_spaces=True)
                time.sleep(0.25)
            except Exception:
                pass

    ctrl, kind = find_list_or_tree(dlg)
    if ctrl is None:
        close_type_dialog_cancel(dlg)
        return False

    selected = False
    if kind == "list":
        selected = _select_in_list(ctrl, target_type)
    elif kind == "tree":
        selected = _select_in_tree(ctrl, target_type)

    if not selected:
        close_type_dialog_cancel(dlg)
        return False

    confirm_type_dialog_ok(dlg)
    return True


# ───────────────────────────────────────────────────────────────────────────
# Functions list helpers
# ───────────────────────────────────────────────────────────────────────────

def get_own_functions(listbox, type_name: str) -> List[str]:
    """Return only the functions that belong to type_name itself.

    The functions list is structured like:
        -------- TypeName --------       <- section header for this type
        OwnFunction1
        OwnFunction2
        -------- ParentType --------     <- start of inherited section, stop here

    If the first separator's inner name does NOT match type_name, the type
    has no own functions and we return [].
    """
    try:
        all_items = listbox.item_texts()
    except Exception:
        return []

    own: List[str] = []
    in_own_section = False

    for item in all_items:
        s = item.strip()
        if not s:
            continue

        if is_separator(s):
            if not in_own_section:
                inner = separator_inner(s)
                if inner.lower() == type_name.lower():
                    in_own_section = True   # found our section, start collecting
                else:
                    break                   # first separator isn't our type — bail
            else:
                break                       # hit the next section — stop
        elif in_own_section:
            own.append(s)

    return own


# ───────────────────────────────────────────────────────────────────────────
# OCR – reading the signature tooltip
# ───────────────────────────────────────────────────────────────────────────

def read_status_text_ocr(catia) -> str:
    # Absolute pixel coords — tune with the debug script if you change monitors
    with mss.mss() as sct:
        shot = sct.grab(OCR_BBOX)
        img = Image.frombytes("RGB", shot.size, shot.rgb)

    img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    img = img.convert("L")
    img = img.point(lambda x: 0 if x < 175 else 255, mode="1")

    text = pytesseract.image_to_string(img, config="--psm 6")
    return text.strip()


def normalize_ocr_text(text: str) -> str:
    text = text.replace("\n", " ").replace("\r", " ")
    text = text.replace("»", "->")
    text = text.replace("{", "(").replace("}", ")")
    text = text.replace("O:", "():").replace(" 0:", "():")
    text = re.sub(r"\s+", " ", text).strip()

    replacements = {
        "Absoluteld":                    "AbsoluteId",
        "Attributelype":                 "AttributeType",
        "GetAttributelnteger":           "GetAttributeInteger",
        "SetAttributelnteger":           "SetAttributeInteger",
        "Removelnstance":                "RemoveInstance",
        "Managelnstance":                "ManageInstance",
        "ActivatelnactivateFeature":     "ActivateInactivateFeature",
        "IsIncludedin":                  "IsIncludedIn",
        "LockPatterninstance":           "LockPatternInstance",
        "Objectlype":                    "ObjectType",
        "O bject":                       "Object",
        "Typefitter":                    "TypeFilter",
        "DefineinterferenceComputation": "DefineInterferenceComputation",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    return text


def extract_signature(raw_text: str, expected_fn: str) -> str:
    text = normalize_ocr_text(raw_text)
    if not text:
        return ""

    text = re.split(r"\bPackage\s*:", text, maxsplit=1)[0].strip()
    text = re.sub(r"^[^A-Za-z]+", "", text)

    m = re.search(
        r"[A-Za-z_][A-Za-z0-9_]*\s*->\s*"
        + re.escape(expected_fn)
        + r"\s*\([^)]*\)\s*:\s*[A-Za-z_][A-Za-z0-9_]*",
        text, flags=re.IGNORECASE,
    )
    if m:
        return m.group(0).strip()

    m = re.search(
        re.escape(expected_fn) + r"\s*\([^)]*\)\s*:\s*[A-Za-z_][A-Za-z0-9_]*",
        text, flags=re.IGNORECASE,
    )
    if m:
        return m.group(0).strip()

    for pat in (
        r"[A-Za-z_][A-Za-z0-9_]*\s*->\s*[A-Za-z_][A-Za-z0-9_]*\s*\([^)]*\)\s*:\s*[A-Za-z_][A-Za-z0-9_]*",
        r"[A-Za-z_][A-Za-z0-9_]*\s*\([^)]*\)\s*:\s*[A-Za-z_][A-Za-z0-9_]*",
    ):
        m = re.search(pat, text)
        if m:
            return m.group(0).strip()

    return ""


def get_clean_signature(catia, expected_fn: str) -> str:
    raw = read_status_text_ocr(catia)
    sig = extract_signature(raw, expected_fn)

    print(f"               [OCR raw] {raw!r}")
    print(f"               [OCR sig] {sig!r}")

    if not sig or expected_fn.lower() not in sig.lower():
        time.sleep(OCR_RETRY_DELAY)
        raw2 = read_status_text_ocr(catia)
        sig2 = extract_signature(raw2, expected_fn)
        print(f"               [OCR retry sig] {sig2!r}")
        if sig2:
            return sig2

    return sig


# ───────────────────────────────────────────────────────────────────────────
# Progress / CSV persistence
# ───────────────────────────────────────────────────────────────────────────

def load_progress(path: str = PROGRESS_JSON) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {"completed_types": {}, "failed_types": []}
        data.setdefault("completed_types", {})
        data.setdefault("failed_types", [])
        return data
    except FileNotFoundError:
        return {"completed_types": {}, "failed_types": []}
    except Exception:
        return {"completed_types": {}, "failed_types": []}


def save_progress(progress: Dict[str, Any], path: str = PROGRESS_JSON):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(progress, f, indent=2, ensure_ascii=False)


def rows_from_progress(progress: Dict[str, Any]) -> List[Row]:
    rows: List[Row] = []
    for type_name, records in progress.get("completed_types", {}).items():
        if not isinstance(records, list):
            continue
        for rec in records:
            try:
                rows.append(Row(
                    object_type=rec["object_type"],
                    function_name=rec["function_name"],
                    signature=rec["signature"],
                ))
            except Exception:
                continue
    return rows


def save_csv(rows: List[Row], path: str = OUTPUT_CSV):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["object_type", "function_name", "signature"])
        writer.writeheader()
        for row in rows:
            writer.writerow(asdict(row))


# ───────────────────────────────────────────────────────────────────────────
# Main scrape loop
# ───────────────────────────────────────────────────────────────────────────

def scrape() -> List[Row]:
    lang, catia = connect()
    functions_list = get_functions_list(lang)
    type_button = get_type_chooser_button(lang)

    progress = load_progress()
    completed_types: Dict[str, List[Dict[str, str]]] = progress.get("completed_types", {})
    failed_types: List[str] = progress.get("failed_types", [])

    type_names = get_all_types_from_dialog(type_button)
    if not type_names:
        raise RuntimeError(
            "No types found from 'Select a Type' dialog. "
            "Try running with DIAGNOSTIC_MODE=True to inspect the dialog controls."
        )

    all_rows = rows_from_progress(progress)
    seen_signatures: Set[str] = {row.signature for row in all_rows}
    remaining_types = [t for t in type_names if t not in completed_types]

    print(f"\nTotal types : {len(type_names)}")
    print(f"Completed   : {len(completed_types)}")
    print(f"Remaining   : {len(remaining_types)}\n")

    for idx, type_name in enumerate(remaining_types, start=1):
        print(f"[{idx:>4}/{len(remaining_types)}] {type_name}")

        try:
            ok = select_type_via_dialog(type_button, type_name)
            if not ok:
                print(f"           [WARN] Could not select type – skipping")
                if type_name not in failed_types:
                    failed_types.append(type_name)
                progress["failed_types"] = failed_types
                save_progress(progress)
                continue

            time.sleep(0.3)

            # Only grab functions belonging to this type, not inherited ones
            function_names = get_own_functions(functions_list, type_name)

            if not function_names:
                print(f"           -> no own functions, skipping")
                completed_types[type_name] = []
                progress["completed_types"] = completed_types
                progress["failed_types"] = [t for t in failed_types if t != type_name]
                save_progress(progress)
                continue

            type_rows: List[Row] = []

            for fn in function_names:
                try:
                    functions_list.select(fn)
                except Exception:
                    continue

                time.sleep(FUNCTION_CLICK_DELAY)
                sig = get_clean_signature(catia, fn)

                if not sig:
                    continue
                if fn.lower() not in sig.lower():
                    continue
                if sig in seen_signatures:
                    continue

                seen_signatures.add(sig)
                row = Row(type_name, fn, sig)
                type_rows.append(row)
                all_rows.append(row)
                print(f"           [fn] {fn:40s}  {sig}")

            completed_types[type_name] = [asdict(r) for r in type_rows]
            progress["completed_types"] = completed_types
            progress["failed_types"] = [t for t in failed_types if t != type_name]

            save_progress(progress)
            save_csv(all_rows)

            print(f"           -> {len(type_rows)} signatures captured")

        except Exception as ex:
            print(f"           [ERROR] {ex}")
            if type_name not in failed_types:
                failed_types.append(type_name)
            progress["failed_types"] = failed_types
            save_progress(progress)

    return all_rows


# ───────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    rows = scrape()
    save_csv(rows)
    print(f"\nDone.  {len(rows)} rows saved to {OUTPUT_CSV}")
    print(f"Progress checkpoint: {PROGRESS_JSON}")