import tkinter as tk
from tkinter import ttk
import sys
import os

print("Step 1: Basic Tkinter Init")
try:
    root = tk.Tk()
    root.withdraw()
    print("Tkinter root created and withdrawn.")
except Exception as e:
    print(f"FAILED Step 1: {e}")
    sys.exit(1)

print("Step 2: Checking PIL (Pillow)")
try:
    from PIL import Image, ImageTk
    # Try to load the logo
    if os.path.exists("assets/logo.png"):
        img = Image.open("assets/logo.png")
        photo = ImageTk.PhotoImage(img)
        print("Pillow image loading successful.")
    else:
        print("assets/logo.png not found, skipping specific image test.")
except Exception as e:
    print(f"FAILED Step 2: {e}")

print("Step 3: Checking sv_ttk")
try:
    import sv_ttk
    sv_ttk.set_theme("dark")
    print("sv_ttk theme set successful.")
except Exception as e:
    print(f"FAILED Step 3: {e}")

print("Step 4: Checking Machine ID (Subprocess)")
try:
    from utils.license_manager import LicenseManager
    mid = LicenseManager.get_machine_id()
    print(f"Machine ID fetch successful: {mid}")
except Exception as e:
    print(f"FAILED Step 4: {e}")

print("All diagnostic steps completed. No Bus Error during sequential tests.")
root.destroy()
