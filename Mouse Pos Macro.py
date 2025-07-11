import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, threading, keyboard, time
import win32api, win32con

CONFIG_DIR = os.path.join(os.path.expanduser("~"), "Documents", "MousePosMacro")
SETTINGS_PATH = os.path.join(CONFIG_DIR, "settings.json")

class Position:
    def __init__(self, x=0, y=0, delay=100, click="left"):
        self.x = x
        self.y = y
        self.delay = delay
        self.click = tk.StringVar(value=click)
        self.label_var = tk.StringVar(value=f"{x}, {y}")
        self.delay_var = tk.StringVar(value=str(delay))

    def to_dict(self):
        return {"x": self.x, "y": self.y, "delay": int(self.delay_var.get()), "click": self.click.get()}

    @staticmethod
    def from_dict(d):
        p = Position(d["x"], d["y"], d.get("delay", 100), d.get("click", "left"))
        p.delay_var.set(str(p.delay))
        return p

class MouseMacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mouse Pos Macro")
        self.positions = []
        self.setpos_key = "f5"
        self.toggle_key = "f6"
        self.macro_running = False
        self.repeat_count = tk.StringVar(value="0")
        self.autoload_path = ""
        self.load_settings()
        self.build_ui()
        self.bind_hotkeys()
        if self.autoload_path and os.path.exists(self.autoload_path):
            self.load_config(self.autoload_path)

    def load_settings(self):
        os.makedirs(CONFIG_DIR, exist_ok=True)
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r") as f:
                data = json.load(f)
                self.autoload_path = data.get("autoload", "")

    def save_settings(self):
        with open(SETTINGS_PATH, "w") as f:
            json.dump({"autoload": self.autoload_path}, f)

    def bind_hotkeys(self):
        # Binds the hotkeys for setting positions and toggling macro
        try:
            if hasattr(self, "_setpos_hook"):
                keyboard.remove_hotkey(self._setpos_hook)
            if hasattr(self, "_toggle_hook"):
                keyboard.remove_hotkey(self._toggle_hook)
        except:
            pass
        self._setpos_hook = keyboard.add_hotkey(self.setpos_key, self.set_next_position)
        self._toggle_hook = keyboard.add_hotkey(self.toggle_key, self.toggle_macro)

    def apply_theme(self):
        # Applies a dark theme
        style = ttk.Style()
        style.theme_use("clam")
        bg = "#2e2e2e"
        fg = "#ffffff"
        self.root.configure(bg=bg)
        style.configure(".", background=bg, foreground=fg, fieldbackground=bg)
        style.configure("TEntry", fieldbackground=bg, foreground=fg)
        style.configure("TCombobox", fieldbackground=bg, foreground=fg)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("TButton", background=bg, foreground=fg)
        style.configure("TFrame", background=bg)

    def build_ui(self):
        # Builds all UI components (buttons, fields, scrollable list)
        self.apply_theme()

        cfg_frame = ttk.Frame(self.root)
        cfg_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(cfg_frame, text="Save As…", command=self.save_as).pack(side="left")
        ttk.Button(cfg_frame, text="Load…", command=self.load_config_dialog).pack(side="left", padx=(5,0))
        self.autoload_combo = ttk.Combobox(cfg_frame, values=self.get_config_files(), width=30)
        if self.autoload_path:
            self.autoload_combo.set(os.path.basename(self.autoload_path))
        self.autoload_combo.pack(side="left", padx=5)
        ttk.Button(cfg_frame, text="Set Auto-Load", command=self.set_autoload).pack(side="left")
        ttk.Button(cfg_frame, text="Refresh", command=self.refresh_config_list).pack(side="left", padx=5)

        keys_frame = ttk.Frame(self.root)
        keys_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(keys_frame, text="Click to rebind:").pack(side="top", anchor="w")
        self.setpos_button = ttk.Button(keys_frame, text=f"Set Pos ({self.setpos_key})", command=lambda: self.rebind_popup("setpos"))
        self.setpos_button.pack(side="left")
        self.toggle_button = ttk.Button(keys_frame, text=f"Start/Stop Macro ({self.toggle_key})", command=lambda: self.rebind_popup("toggle"))
        self.toggle_button.pack(side="left", padx=(5,0))

        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(control_frame, text="Repeat (0 = infinite):").pack(side="left")
        ttk.Entry(control_frame, textvariable=self.repeat_count, width=6).pack(side="left", padx=(0, 10))

        self.pos_canvas = tk.Canvas(self.root, height=200, bg="#2e2e2e", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.pos_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.pos_canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.pos_canvas.configure(scrollregion=self.pos_canvas.bbox("all")))
        self.pos_canvas.create_window((0,0), window=self.scrollable_frame, anchor="nw")
        self.pos_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.pos_canvas.pack(side="left", fill="both", expand=True, padx=(10,0))
        self.scrollbar.pack(side="right", fill="y", padx=(0,10))

        ctrl = ttk.Frame(self.root)
        ctrl.pack(fill="x", padx=10, pady=5)
        ttk.Button(ctrl, text="Add +", command=self.add_position).pack(side="left")
        ttk.Button(ctrl, text="Clear All", command=self.clear_positions).pack(side="left", padx=(5,0))

        footer_frame = ttk.Frame(self.root)
        footer_frame.pack(side="bottom", fill="x", padx=10, pady=5)
        ttk.Label(footer_frame, text="Made by").pack(side="left")
        tk.Label(footer_frame, text=" iiVee", fg="red", bg="#2e2e2e").pack(side="left")
        tk.Label(footer_frame, text="  Discord: j7f", fg="purple", bg="#2e2e2e").pack(side="left", padx=(10, 0))

        self.update_positions_ui()

    def get_config_files(self):
        return [f for f in os.listdir(CONFIG_DIR) if f.lower().endswith(".json") and f != "settings.json"]

    def refresh_config_list(self):
        self.autoload_combo["values"] = self.get_config_files()

    def set_autoload(self):
        sel = self.autoload_combo.get()
        full = os.path.join(CONFIG_DIR, sel)
        if os.path.exists(full):
            self.autoload_path = full
            self.save_settings()
            messagebox.showinfo("Auto-load", f"Will auto-load: {sel} on startup")
        else:
            messagebox.showerror("Error", f"Config file does not exist: {sel}")

    def rebind_popup(self, which):
        # Lets the user change the hotkey
        top = tk.Toplevel(self.root)
        top.title("Press a key…")
        tk.Label(top, text="Press the new hotkey…").pack(padx=20, pady=10)
        top.grab_set()
        top.focus_force()
        top.attributes('-topmost', True)

        def on_key(e):
            key = e.keysym.lower()
            if which == "setpos":
                self.setpos_key = key
                self.setpos_button.config(text=f"Set Pos ({self.setpos_key})")
            else:
                self.toggle_key = key
                self.toggle_button.config(text=f"Start/Stop Macro ({self.toggle_key})")
            top.destroy()
            self.bind_hotkeys()

        top.bind("<Key>", on_key)

    def add_position(self):
        self.positions.append(Position())
        self.update_positions_ui()

    def clear_positions(self):
        self.positions.clear()
        self.update_positions_ui()

    def set_next_position(self):
        # Sets the first unused slot to the current mouse position using win32api.GetCursorPos()
        for p in self.positions:
            if p.x == 0 and p.y == 0:
                x, y = win32api.GetCursorPos()
                p.x, p.y = x, y
                p.label_var.set(f"{x}, {y}")
                self.update_positions_ui()
                return
        self.add_position()
        x, y = win32api.GetCursorPos()
        self.positions[-1].x, self.positions[-1].y = x, y
        self.positions[-1].label_var.set(f"{x}, {y}")
        self.update_positions_ui()

    def update_positions_ui(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        for idx, p in enumerate(self.positions):
            row = ttk.Frame(self.scrollable_frame)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"{idx+1}:").pack(side="left")
            ttk.Label(row, textvariable=p.label_var, width=15).pack(side="left", padx=5)
            ttk.Label(row, text="Delay:").pack(side="left")
            ttk.Entry(row, width=6, textvariable=p.delay_var).pack(side="left", padx=5)
            ttk.OptionMenu(row, p.click, p.click.get(), "left", "right", "middle").pack(side="left", padx=5)
            ttk.Button(row, text="✖", width=3, command=lambda i=idx: self.remove_position(i)).pack(side="right", padx=5)

    def remove_position(self, index):
        if 0 <= index < len(self.positions):
            del self.positions[index]
            self.update_positions_ui()


    def save_as(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", initialdir=CONFIG_DIR, filetypes=[("JSON files","*.json")])
        if path:
            self._save_config(path)

    def load_config_dialog(self):
        path = filedialog.askopenfilename(initialdir=CONFIG_DIR, filetypes=[("JSON files","*.json")])
        if path:
            self.load_config(path)

    def load_config(self, path):
        if not os.path.exists(path):
            print(f"[WARN] Config file not found: {path}")
            return
        with open(path, "r") as f:
            data = json.load(f)
        self.positions = [Position.from_dict(d) for d in data.get("positions", [])]
        for p in self.positions:
            p.label_var = tk.StringVar(value=f"{p.x}, {p.y}")
        self.setpos_key = data.get("setpos_key", self.setpos_key)
        self.toggle_key = data.get("toggle_key", self.toggle_key)
        self.repeat_count.set(str(data.get("repeat_count", "0")))
        self.bind_hotkeys()
        self.update_positions_ui()
        self.setpos_button.config(text=f"Set Pos ({self.setpos_key})")
        self.toggle_button.config(text=f"Start/Stop Macro ({self.toggle_key})")

    def _save_config(self, path):
        os.makedirs(CONFIG_DIR, exist_ok=True)
        data = {
            "positions": [p.to_dict() for p in self.positions],
            "setpos_key": self.setpos_key,
            "toggle_key": self.toggle_key,
            "repeat_count": int(self.repeat_count.get()),
        }
        with open(path, "w") as f:
            json.dump(data, f)
        messagebox.showinfo("Saved", f"Config saved to:\n{path}")

    def toggle_macro(self):
        if self.macro_running:
            self.macro_running = False
        else:
            self.macro_running = True
            threading.Thread(target=self.run_macro, daemon=True).start()

    def run_macro(self):
        try:
            repeat = float("inf") if self.repeat_count.get().strip() == "0" else int(self.repeat_count.get())
        except:
            messagebox.showerror("Invalid input", "Repeat must be a number.")
            return

        BUTTONS = {
            "left": (win32con.MOUSEEVENTF_LEFTDOWN, win32con.MOUSEEVENTF_LEFTUP),
            "right": (win32con.MOUSEEVENTF_RIGHTDOWN, win32con.MOUSEEVENTF_RIGHTUP),
            "middle": (win32con.MOUSEEVENTF_MIDDLEDOWN, win32con.MOUSEEVENTF_MIDDLEUP),
        }

        count = 0
        while self.macro_running and (count < repeat):
            for p in self.positions:
                if not self.macro_running:
                    return
                # Move the mouse cursor
                win32api.SetCursorPos((p.x, p.y))

                # Get the down/up events for the selected button
                down, up = BUTTONS.get(p.click.get(), BUTTONS["left"])
                # Simulate the click
                win32api.mouse_event(down, 0, 0, 0, 0)
                win32api.mouse_event(up, 0, 0, 0, 0)

                # Delay between clicks
                try:
                    delay = max(0.0001, int(p.delay_var.get()) / 1000.0)
                except:
                    delay = 0.01
                time.sleep(delay)
            count += 1

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.iconbitmap("computermouse.ico")
    except Exception:
        pass
    app = MouseMacroApp(root)
    root.mainloop()
