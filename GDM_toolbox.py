
import os
import sys
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as ttkb

# ------------------------------------------------------------
# Resource Path (PyInstaller safe)
# ------------------------------------------------------------
def resource_path(relative_path: str) -> str:
    """
    Get absolute path to resource, works for dev and for PyInstaller.
    In PyInstaller, files are unpacked to sys._MEIPASS.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

# ------------------------------------------------------------
# UI Helpers
# ------------------------------------------------------------
def on_enter(event):
    # keep hover lightweight
    try:
        event.widget.configure(bootstyle="info")
    except Exception:
        pass

def on_leave(event):
    try:
        event.widget.configure(bootstyle="info")
    except Exception:
        pass

def switch_page(page_name: str):
    for frame in pages.values():
        frame.pack_forget()
    pages[page_name].pack(fill=tk.BOTH, expand=True)

def change_theme(theme):
    # Light/Dark themes
    if theme == "Light":
        root.style.theme_use("flatly")
    else:
        root.style.theme_use("superhero")

def change_font_size(size):
    # Keep simple: only update direct children in pages
    for page in pages.values():
        for widget in page.winfo_children():
            try:
                widget.configure(font=("Arial", size))
            except Exception:
                pass

def change_layout(layout):
    button_size = {
        "Compact": {"width": 12, "padx": 4, "pady": 6},
        "Spacious": {"width": 18, "padx": 10, "pady": 14},
    }
    if layout not in button_size:
        return

    cfg = button_size[layout]
    for page in pages.values():
        for container in page.winfo_children():
            for child in container.winfo_children():
                if isinstance(child, (tk.Button, ttkb.Button)):
                    try:
                        child.config(width=cfg["width"])
                        child.grid_configure(padx=cfg["padx"], pady=cfg["pady"])
                    except Exception:
                        pass

# ------------------------------------------------------------
# PDF Open (Lazy + Safe)
# ------------------------------------------------------------
def info_open(file_path: str) -> bool:
    """
    Open a PDF file with default viewer (Windows).
    """
    try:
        if not file_path:
            messagebox.showerror("Error", "No file path provided.")
            return False

        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"PDF not found:\n{file_path}")
            return False

        if sys.platform.startswith("win"):
            os.startfile(file_path)
            return True

        # fallback for non-windows (optional)
        import subprocess
        if sys.platform.startswith("darwin"):
            subprocess.Popen(["open", file_path])
        else:
            subprocess.Popen(["xdg-open", file_path])
        return True

    except Exception as e:
        messagebox.showerror("Error", f"Error opening info file:\n{e}")
        return False

def info_pdf_for(index: int) -> str:
    files = {
        0: resource_path(r"Resources\Files\VGRF_Macro_Help.pdf"),
        1: resource_path(r"Resources\Files\VGRF_Merge_Help.pdf"),
        2: resource_path(r"Resources\Files\TR_Consolidation_Help.pdf"),
        3: resource_path(r"Resources\Files\TR_Template_Help.pdf"),
        4: resource_path(r"Resources\Files\Link_Generator_Help.pdf"),
    }
    return files.get(index, resource_path(r"Resources\Files\General_Help.pdf"))

# ------------------------------------------------------------
# Excel Automation (Lazy Imports - improves startup time) [1](https://coderslegacy.com/python/speed-up-pyinstaller-exe-load-time/)[2](https://github.com/orgs/pyinstaller/discussions/9080)
# ------------------------------------------------------------
def _run_excel_macro(xlsm_path: str, macro_name: str, visible: bool = True):
    """
    Helper to run a macro using win32com (imported lazily).
    """
    try:
        import win32com.client  # LAZY IMPORT

        xl = win32com.client.DispatchEx("Excel.Application")
        wb = xl.Workbooks.Open(xlsm_path)
        xl.Application.Run(macro_name)
        xl.Visible = visible
        # If you want Excel to close automatically, set visible=False and close workbook.
        # wb.Close(SaveChanges=False)
        # xl.Quit()

    except Exception as e:
        messagebox.showerror("Error", f"Excel macro failed:\n{e}")

def VGRF_Macro():
    xlsm = resource_path(r"Resources\files\VGRF_Macro.xlsm")
    _run_excel_macro(xlsm, "VGRF_Macro.xlsm!TRmacro", visible=True)

def VGRF_Merge():
    xlsm = resource_path(r"Resources\files\VGRF_Macro.xlsm")
    _run_excel_macro(xlsm, "VGRF_Macro.xlsm!Merge_VGRF", visible=True)

def TR_Consolidation():
    xlsm = resource_path(r"Resources\files\TR_Consolidation.xlsm")
    _run_excel_macro(xlsm, "TR_Consolidation.xlsm!consolidating_macro", visible=True)

def TR_Template():
    try:
        import win32com.client  # LAZY IMPORT
        xl = win32com.client.DispatchEx("Excel.Application")
        wb_path = resource_path(r"Resources\files\TR_template.xlsm")
        xl.Workbooks.Open(wb_path)
        xl.Visible = True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open TR Template:\n{e}")

def Coming_soon():
    messagebox.showinfo("Coming Soon", "This feature is under development.")

def Link_Generator():
    # xlwings is heavy - import only here (lazy)
    try:
        import xlwings as xw  # LAZY IMPORT
        import win32com.client  # LAZY IMPORT
    except Exception as e:
        messagebox.showerror("Error", f"Missing dependency:\n{e}")
        return

    root2 = ttkb.Window(themename="litera")
    root2.title("Link Generator")
    root2.geometry("600x125")

    ttkb.Label(root2, text="Link for duplication:").grid(row=0, column=0, padx=10, pady=7)
    entry1 = ttkb.Entry(root2, width=60)
    entry1.grid(row=0, column=1, padx=10, pady=7)

    ttkb.Label(root2, text="Server (india):").grid(row=1, column=0, padx=10, pady=7)
    entry2 = ttkb.Entry(root2, width=60)
    entry2.grid(row=1, column=1, padx=10, pady=7)

    ttkb.Label(root2, text="Server (europe):").grid(row=2, column=0, padx=10, pady=7)
    entry3 = ttkb.Entry(root2, width=60)
    entry3.grid(row=2, column=1, padx=10, pady=7)

    def Xl_link_generator():
        try:
            messagebox.showinfo("Information", "This feature is under development")
            # wb_path = resource_path(r"Resources\Files\Link_Generator.xlsm")
            # app = xw.App(visible=False)
            # wb = app.books.open(wb_path)

            # sheet = wb.sheets["Sheet1"]
            # sheet.range("D3").value = entry1.get()
            # wb.app.calculate()

            # india_link = sheet.range("D5").value
            # europe_link = sheet.range("D6").value

            # entry2.delete(0, "end")
            # entry2.insert(0, india_link)

            # entry3.delete(0, "end")
            # entry3.insert(0, europe_link)

            # wb.save()
            # wb.close()
            # app.quit()

            # xl = win32com.client.DispatchEx("Excel.Application")
            # wb2 = xl.Workbooks.Open(wb_path)
            # xl.Application.Run("Link_Generator.xlsm!CreateFolderStructure")
            # xl.Application.Run("Link_Generator.xlsm!Copy_Folder")
            # wb2.Close(SaveChanges=False)

        except Exception as e:
            messagebox.showerror("Error", f"Error updating Excel file:\n{e}")

    ttkb.Button(root2, text="Submit", command=Xl_link_generator).grid(row=0, column=2, columnspan=2, pady=7)
    root2.mainloop()

# ------------------------------------------------------------
# Main Window (keep startup light)
# ------------------------------------------------------------
root = ttkb.Window(themename="litera")
root.title("GDM ToolBox V1.0")
root.geometry("750x300")

# App icon: use lazy PIL only if needed
try:
    from PIL import Image, ImageTk  # still relatively light but optional
    icon_path_app = resource_path(r"Resources\Icons\ico1.png")
    if os.path.exists(icon_path_app):
        img = Image.open(icon_path_app)
        icon = ImageTk.PhotoImage(img)
        root.icon_image = icon
        root.iconphoto(False, root.icon_image)
except Exception:
    # Ignore icon failures to avoid slowing startup or crashing
    pass

# Menu bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

theme_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Theme", menu=theme_menu)
theme_menu.add_command(label="Light", command=lambda: change_theme("Light"))
theme_menu.add_command(label="Dark", command=lambda: change_theme("Dark"))

font_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Font Size", menu=font_menu)
font_menu.add_command(label="Small", command=lambda: change_font_size(5))
font_menu.add_command(label="Medium", command=lambda: change_font_size(14))
font_menu.add_command(label="Large", command=lambda: change_font_size(18))

layout_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Layout", menu=layout_menu)
layout_menu.add_command(label="Compact", command=lambda: change_layout("Compact"))
layout_menu.add_command(label="Spacious", command=lambda: change_layout("Spacious"))

# Pages registry
pages = {}

# ------------------------------------------------------------
# Page 1
# ------------------------------------------------------------
page1 = ttkb.Frame(root)
pages["Page 1"] = page1

ttkb.Label(page1, text="Analysis", font=("Arial", 18)).pack(pady=10)

button_frame1 = tk.Frame(page1)
button_frame1.pack(pady=30)

ttkb.Label(page1, text="© Powered by GDMC-BLR", font=("Arial", 5)).place(
    relx=0.0, rely=1.0, anchor="sw", x=5, y=-5
)

Page1_arr1 = [
    "VGRF Macro-TR", "VGRF Merge", "TR Consolidation", "TR Template",
    "Link Generator", "Coming Soon", "Coming Soon",
    "Coming Soon", "Coming Soon", "Coming Soon"
]
Page1_arr2 = [
    VGRF_Macro, VGRF_Merge, TR_Consolidation, TR_Template,
    Link_Generator, Coming_soon, Coming_soon, Coming_soon,
    Coming_soon, Coming_soon
]

# Create 2 rows x 5 columns; each cell has [i] + [main button]
for i in range(10):
    row = i // 5
    col = i % 5

    cell = tk.Frame(button_frame1)
    cell.grid(row=row, column=col, padx=9, pady=12, sticky="w")
    button_frame1.grid_columnconfigure(col, weight=1)

    # info "i" button (single)
    info_btn = ttkb.Button(
        cell,
        text="i",
        width=2,
        bootstyle="primary",
        padding=(1, 1),
        command=lambda i=i: info_open(info_pdf_for(i)),
        cursor="hand2"
    )
    try:
        info_btn.configure(font=("Arial", 8))
    except Exception:
        pass
    info_btn.pack(side="left", padx=(0, 2))
    # info_btn.bind("<Enter>", on_enter)
    # info_btn.bind("<Leave>", on_leave)


    # main button
    main_btn = ttkb.Button(
        cell,
        text=Page1_arr1[i],
        width=18,
        bootstyle="primary",
        padding=(1, 1),
        command=Page1_arr2[i]
    )
    main_btn.pack(side="left", fill="x", padx=(0, 0))
    # main_btn.bind("<Enter>", on_enter)
    # main_btn.bind("<Leave>", on_leave)


# ------------------------------------------------------------
# Page 2 (placeholder)
# ------------------------------------------------------------
page2 = ttkb.Frame(root)
pages["Page 2"] = page2

ttkb.Label(page2, text="Project Management", font=("Arial", 18)).pack(pady=10)

button_frame2 = tk.Frame(page2)
button_frame2.pack(pady=30)

ttkb.Label(page2, text="© Powered by GDMC-BLR", font=("Arial", 5)).place(
    relx=0.0, rely=1.0, anchor="sw", x=5, y=-5
)

Page2_arr1 = ["Coming Soon"] * 10
Page2_arr2 = [Coming_soon] * 10

for i in range(10):
    r = i // 5
    c = i % 5
    btn = ttkb.Button(
        button_frame2,
        text=Page2_arr1[i],
        width=18,
        bootstyle="primary",
        command=Page2_arr2[i]
    )
    btn.grid(row=r, column=c, padx=9, pady=10, sticky="nsew")
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)

# ------------------------------------------------------------
# Top Navigation
# ------------------------------------------------------------
nav_frame = tk.Frame(root)
nav_frame.pack(side=tk.TOP, fill=tk.X)

tk.Button(nav_frame, text="Measurement", command=lambda: switch_page("Page 1")).pack(
    side=tk.LEFT, padx=5, pady=5
)

# Optional logos (lazy PIL)
try:
    from PIL import Image, ImageTk
    te_path = resource_path(r"Resources\Logo\TE_Logo.png")
    gdmc_path = resource_path(r"Resources\Logo\GDMC_Logo.png")

    if os.path.exists(te_path):
        te_img = Image.open(te_path).resize((60, 30))
        telogo = ImageTk.PhotoImage(te_img)
        nav_frame.telogo = telogo
        tk.Label(nav_frame, image=telogo).pack(side=tk.RIGHT, padx=5, pady=5)

    if os.path.exists(gdmc_path):
        gdmc_img = Image.open(gdmc_path).resize((95, 30))
        gdmclogo = ImageTk.PhotoImage(gdmc_img)
        nav_frame.gdmclogo = gdmclogo
        tk.Label(nav_frame, image=gdmclogo).pack(side=tk.RIGHT, padx=5, pady=5)

except Exception:
    pass

# ------------------------------------------------------------
# Boot
# ------------------------------------------------------------
switch_page("Page 1")
root.style.theme_use("flatly")
root.mainloop()
