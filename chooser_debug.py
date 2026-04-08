from pywinauto import Desktop
import time

print("Script started")

try:
    desktop = Desktop(backend="win32")
    print("Desktop acquired")
except Exception as e:
    print("Failed to get Desktop:", repr(e))
    raise

try:
    wins = desktop.windows()
    print(f"Found {len(wins)} top-level windows")
except Exception as e:
    print("Failed to enumerate windows:", repr(e))
    raise

print("\n=== Matching windows ===")
matches = []
for i, w in enumerate(wins):
    try:
        title = w.window_text()
        if title and ("CATIA" in title or "Language" in title):
            print(f"[{i}] {title!r} handle={w.handle}")
            matches.append(w)
    except Exception as e:
        print(f"[{i}] error reading title:", repr(e))

print("\n=== Exact Language browser lookup ===")
try:
    lang = desktop.window(title="Language browser")
    print("Got window object")
    print("Exists:", lang.exists(timeout=3))
    print("Visible:", lang.is_visible())
    print("Handle:", lang.handle)
    print("Text:", repr(lang.window_text()))
except Exception as e:
    print("Exact lookup failed:", repr(e))
    lang = None

if lang:
    print("\n=== Children of Language browser ===")
    try:
        children = lang.children()
        print(f"Child count: {len(children)}")
        for c in children[:30]:
            try:
                print(c.friendly_class_name(), repr(c.window_text()))
            except Exception as e:
                print("Child read error:", repr(e))
    except Exception as e:
        print("Failed reading children:", repr(e))

print("\nDone")
time.sleep(1)