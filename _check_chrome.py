"""Check Chrome profile and Cookies DB location."""
import os
from pathlib import Path

localappdata = os.environ.get("LOCALAPPDATA", "")
print(f"LOCALAPPDATA = {localappdata}")

chrome_dir = Path(localappdata) / "Google" / "Chrome" / "User Data"
print(f"Chrome dir   = {chrome_dir}")
print(f"Exists       = {chrome_dir.exists()}")

if chrome_dir.exists():
    print("\nSubfolders:")
    for item in sorted(chrome_dir.iterdir()):
        if item.is_dir():
            # Check if it's a profile folder
            cookies1 = item / "Network" / "Cookies"
            cookies2 = item / "Cookies"
            has_cookies = cookies1.exists() or cookies2.exists()
            marker = " ← HAS COOKIES" if has_cookies else ""
            print(f"  {item.name}{marker}")
    
    # Check Default profile specifically
    for profile in ["Default", "Profile 1", "Profile 2", "Profile 3"]:
        pdir = chrome_dir / profile
        if pdir.exists():
            c1 = pdir / "Network" / "Cookies"
            c2 = pdir / "Cookies"
            print(f"\n{profile}:")
            print(f"  Network/Cookies: {c1} → exists={c1.exists()}")
            print(f"  Cookies:         {c2} → exists={c2.exists()}")
            if c1.exists():
                print(f"  Size: {c1.stat().st_size} bytes")
            elif c2.exists():
                print(f"  Size: {c2.stat().st_size} bytes")
else:
    # Try alternate locations
    print("\nChrome not found at default location. Trying alternates:")
    for alt in [
        Path(localappdata) / "Google" / "Chrome SxS" / "User Data",
        Path(localappdata) / "Chromium" / "User Data",
        Path(os.environ.get("APPDATA", "")) / "Google" / "Chrome" / "User Data",
        Path(localappdata) / "Microsoft" / "Edge" / "User Data",
    ]:
        print(f"  {alt} → exists={alt.exists()}")
