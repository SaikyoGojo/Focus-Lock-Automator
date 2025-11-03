import time
import pyautogui
import pygetwindow as gw
import random
import win32gui
import threading
import tkinter as tk
from tkinter import ttk, messagebox
import configparser
import os

# --- GLOBAL SETTINGS ---
PROFILES_FILE = 'automator_profiles.ini'
MIN_MOVE = 5; MAX_MOVE = 15; pyautogui.FAILSAFE = False 
FAST_PAUSE = 0.05
LONG_PRESS_DURATION = 0.5 

class AfkGuiApp:
    def __init__(self, master):
        self.master = master
        master.title("Focus-Lock Automator v6.0")
        master.resizable(False, False)

        # State & Config Variables
        self.is_running = False
        self.is_paused = False
        self.afk_thread = None
        self.config = configparser.ConfigParser() # For profiles only
        
        # Configuration Variables (Defaults are non-persistent)
        self.target_window_title = tk.StringVar(value="")
        self.min_delay_var = tk.DoubleVar(value=10.0)
        self.max_delay_var = tk.DoubleVar(value=120.0)
        self.long_press_weight_var = tk.IntVar(value=60)
        self.mouse_action_enabled = tk.BooleanVar(value=True) # Jiggle & Left Click
        self.right_click_enabled = tk.BooleanVar(value=False)
        self.scroll_enabled = tk.BooleanVar(value=False)
        self.long_press_keys = tk.StringVar(value="q, e, shift") 
        self.tap_keys = tk.StringVar(value="i, o, enter, F1") 
        self.profile_name_var = tk.StringVar(value="New Profile")
        
        self._create_widgets()
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

    # --- PROFILE MANAGEMENT ---

    def save_profile(self):
        """Saves current GUI settings to a named profile in the INI file."""
        profile_name = self.profile_name_var.get().strip()
        if not profile_name:
            messagebox.showerror("Error", "Please enter a valid profile name.")
            return

        self.config.read(PROFILES_FILE)
        
        # Save all current settings to the profile section
        self.config[profile_name] = {
            'target_window': self.target_window_title.get(),
            'min_delay': str(self.min_delay_var.get()),
            'max_delay': str(self.max_delay_var.get()),
            'long_press_weight': str(self.long_press_weight_var.get()),
            'mouse_action_enabled': str(self.mouse_action_enabled.get()),
            'right_click_enabled': str(self.right_click_enabled.get()),
            'scroll_enabled': str(self.scroll_enabled.get()),
            'long_press_keys': self.long_press_keys.get(),
            'tap_keys': self.tap_keys.get()
        }
        
        try:
            with open(PROFILES_FILE, 'w') as configfile:
                self.config.write(configfile)
            self.log_message(f"Profile '{profile_name}' saved successfully.")
        except Exception as e:
            self.log_message(f"ERROR: Could not save profile: {e}")

    def load_profile(self):
        """Loads settings from a named profile into the GUI."""
        self.config.read(PROFILES_FILE)
        profile_name = self.profile_name_var.get().strip()
        
        if profile_name in self.config:
            s = self.config[profile_name]
            
            self.target_window_title.set(s.get('target_window', self.target_window_title.get()))
            self.min_delay_var.set(s.getfloat('min_delay', self.min_delay_var.get()))
            self.max_delay_var.set(s.getfloat('max_delay', self.max_delay_var.get()))
            self.long_press_weight_var.set(s.getint('long_press_weight', self.long_press_weight_var.get()))
            self.mouse_action_enabled.set(s.getboolean('mouse_action_enabled', self.mouse_action_enabled.get()))
            self.right_click_enabled.set(s.getboolean('right_click_enabled', self.right_click_enabled.get()))
            self.scroll_enabled.set(s.getboolean('scroll_enabled', self.scroll_enabled.get()))
            self.long_press_keys.set(s.get('long_press_keys', self.long_press_keys.get()))
            self.tap_keys.set(s.get('tap_keys', self.tap_keys.get()))
            
            self.log_message(f"Profile '{profile_name}' loaded.")
            self.refresh_windows() # Ensure window list is current
        else:
            messagebox.showerror("Error", f"Profile '{profile_name}' not found.")

    # --- GUI CONSTRUCTION ---

    def _create_widgets(self):
        
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # --- Frame 1: Window Selection ---
        window_frame = ttk.LabelFrame(main_frame, text="1. Target Application")
        window_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky="ew")

        ttk.Label(window_frame, text="Select Window:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.window_selector = ttk.Combobox(window_frame, textvariable=self.target_window_title, width=35, state='readonly')
        self.window_selector.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.refresh_button = ttk.Button(window_frame, text="Refresh", command=self.refresh_windows)
        self.refresh_button.grid(row=0, column=2, padx=5, pady=5)
        self.refresh_windows()

        # --- Notebook (Tabs) ---
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, columnspan=2, pady=5, sticky="ew")

        # Create Tab Frames
        self.fast_tab = ttk.Frame(self.notebook, padding="5")
        self.custom_tab = ttk.Frame(self.notebook, padding="5")

        self.notebook.add(self.fast_tab, text='🚀 Fast Mode')
        self.notebook.add(self.custom_tab, text='⚙️ Custom Mode')

        self._build_fast_tab()
        self._build_custom_tab()
        self._build_control_widgets(main_frame)

    def _build_fast_tab(self):
        # A simple, quick click mode (Fastest Execution)
        ttk.Label(self.fast_tab, text="Mode: Optimized for minimal interruption. Performs a single, fast Jiggle+Click every interval.", wraplength=400).grid(row=0, column=0, columnspan=3, pady=5, sticky="w")
        
        ttk.Label(self.fast_tab, text="Min Delay (s):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.fast_tab, textvariable=self.min_delay_var, width=10).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.fast_tab, text="Max Delay (s):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.fast_tab, textvariable=self.max_delay_var, width=10).grid(row=2, column=1, padx=5, pady=5)

    def _build_custom_tab(self):
        # --- Custom Actions ---
        action_frame = ttk.LabelFrame(self.custom_tab, text="Custom Actions (Comma separated list)")
        action_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky="ew")
        
        # Mouse Actions
        ttk.Label(action_frame, text="Mouse Actions:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Checkbutton(action_frame, text="Jiggle & LClick", variable=self.mouse_action_enabled).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Checkbutton(action_frame, text="Right Click", variable=self.right_click_enabled).grid(row=0, column=2, padx=5, pady=5, sticky="w")
        ttk.Checkbutton(action_frame, text="Scroll", variable=self.scroll_enabled).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        # Key Actions
        ttk.Label(action_frame, text="Long Press:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(action_frame, textvariable=self.long_press_keys, width=50).grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky="ew")

        ttk.Label(action_frame, text="Tap Press:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(action_frame, textvariable=self.tap_keys, width=50).grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky="ew")
        
        # --- Timing & Priority ---
        settings_frame = ttk.LabelFrame(self.custom_tab, text="Timing & Priority")
        settings_frame.grid(row=1, column=0, pady=5, sticky="nw")

        ttk.Label(settings_frame, text="Min Delay (s):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(settings_frame, textvariable=self.min_delay_var, width=10).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(settings_frame, text="Max Delay (s):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(settings_frame, textvariable=self.max_delay_var, width=10).grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(settings_frame, text="Long Press Priority (%):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Scale(settings_frame, from_=0, to=100, variable=self.long_press_weight_var, orient=tk.HORIZONTAL, length=100).grid(row=2, column=1, padx=5, pady=5)
        ttk.Label(settings_frame, textvariable=self.long_press_weight_var).grid(row=2, column=2, padx=5, pady=5)

    def _build_control_widgets(self, main_frame):
        # --- Profile Management (Row 2, Column 0) ---
        profile_frame = ttk.LabelFrame(main_frame, text="2. Profiles")
        profile_frame.grid(row=2, column=0, pady=5, padx=5, sticky="ew")
        
        ttk.Label(profile_frame, text="Name:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(profile_frame, textvariable=self.profile_name_var, width=15).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(profile_frame, text="Save", command=self.save_profile).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(profile_frame, text="Load", command=self.load_profile).grid(row=0, column=3, padx=5, pady=5)

        # --- Control & Status (Row 2, Column 1) ---
        control_frame = ttk.LabelFrame(main_frame, text="3. Control & Status")
        control_frame.grid(row=2, column=1, pady=5, sticky="ne")
        
        self.status_label = ttk.Label(control_frame, text="Status: Stopped", font=('Arial', 10, 'bold'))
        self.status_label.grid(row=0, column=0, columnspan=3, pady=5)

        self.start_button = ttk.Button(control_frame, text="Start", command=self.start_afk, style='Accent.TButton')
        self.start_button.grid(row=1, column=0, padx=5, pady=5)
        
        self.pause_button = ttk.Button(control_frame, text="Pause", command=self.pause_afk, state=tk.DISABLED)
        self.pause_button.grid(row=1, column=1, padx=5, pady=5)
        
        self.stop_button = ttk.Button(control_frame, text="Stop", command=self.stop_afk, state=tk.DISABLED)
        self.stop_button.grid(row=1, column=2, padx=5, pady=5)

        # --- Logger (Row 3) ---
        log_frame = ttk.LabelFrame(main_frame, text="Activity Log")
        log_frame.grid(row=3, column=0, columnspan=2, pady=5, sticky="ew")

        self.log_text = tk.Text(log_frame, height=5, state=tk.DISABLED, wrap='word', bg='#f0f0f0', borderwidth=1, relief="sunken")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Accent.TButton', foreground='white', background='#32CD32', borderwidth=0, relief='flat')
        style.map('Accent.TButton', background=[('active', '#228B22')])

    # --- ACTION LOGIC ---

    def log_message(self, message, is_warning=False):
        """Appends a timestamped message to the GUI log."""
        current_time = time.strftime("[%H:%M:%S]")
        log_message = f"{current_time} {message}\n"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def find_target_window(self):
        """Finds the window object."""
        title = self.target_window_title.get()
        if not title: return None
        try:
            windows = gw.getWindowsWithTitle(title)
            return windows[0] if windows else None
        except Exception:
            return None
        
    def focus_target_window(self, target_window, previous_hwnd):
        """CRITICAL FOCUS FUNCTION: Ensures the target window is active."""
        if target_window.isMinimized:
            target_window.restore()
            time.sleep(0.1) # Time to draw window

        win32gui.SetForegroundWindow(target_window._hWnd)
        time.sleep(0.1) # Guaranteed focus wait
        
        # Setup mouse position for input
        center_x = target_window.left + target_window.width // 2
        center_y = target_window.top + target_window.height // 2
        
        return center_x, center_y

    def execute_fast_action(self):
        """Optimized for extreme speed and minimal interruption."""
        previous_hwnd = win32gui.GetForegroundWindow()
        original_mouse_x, original_mouse_y = pyautogui.position() 
        target_window = self.find_target_window()
        
        if target_window is None:
            self.update_status("Waiting for Target...")
            return

        try:
            center_x, center_y = self.focus_target_window(target_window, previous_hwnd)

            # --- FAST ACTION: Minimal Jiggle & Click ---
            pyautogui.moveTo(center_x, center_y, duration=0.0) 
            pyautogui.move(random.randint(MIN_MOVE, MAX_MOVE) * random.choice([1, -1]),
                           random.randint(MIN_MOVE, MAX_MOVE) * random.choice([1, -1]), duration=0.0)
            pyautogui.click() 
            
            # Revert focus
            pyautogui.moveTo(original_mouse_x, original_mouse_y, duration=0.0) 
            win32gui.SetForegroundWindow(previous_hwnd)
            self.log_message(f"EXECUTED: 🚀 Fast Jiggle+Click.")

        except Exception:
            win32gui.SetForegroundWindow(previous_hwnd)
            self.log_message(f"WARNING: Fast action failed.", is_warning=True)

    def execute_custom_action(self):
        """The full, weighted, multi-input action sequence."""
        previous_hwnd = win32gui.GetForegroundWindow()
        original_mouse_x, original_mouse_y = pyautogui.position() 
        target_window = self.find_target_window()
        
        if target_window is None:
            self.update_status("Waiting for Target...")
            return

        action_log = []
        
        try:
            center_x, center_y = self.focus_target_window(target_window, previous_hwnd)
            
            # Parse settings for this cycle
            long_keys = self.parse_keys(self.long_press_keys.get())
            tap_keys = self.parse_keys(self.tap_keys.get())
            
            # Calculate dynamic weights
            weights = self._calculate_weights(long_keys, tap_keys)
            
            # --- MOUSE JIGGLE (If enabled) ---
            if self.mouse_action_enabled.get():
                pyautogui.moveTo(center_x, center_y, duration=0.0) 
                pyautogui.move(random.randint(MIN_MOVE, MAX_MOVE) * random.choice([1, -1]),
                               random.randint(MIN_MOVE, MAX_MOVE) * random.choice([1, -1]), duration=0.0)
                pyautogui.click() 
                action_log.append("Jiggle+LClick")
            
            # --- WEIGHTED ACTION ---
            if weights:
                chosen_action_type = random.choices(list(weights.keys()), weights=list(weights.values()), k=1)[0]
                
                if chosen_action_type == 'long_press':
                    action_key = random.choice(long_keys)
                    pyautogui.keyDown(action_key); time.sleep(LONG_PRESS_DURATION); pyautogui.keyUp(action_key)
                    action_log.append(f"Long Press: {action_key.upper()}")
                    
                elif chosen_action_type == 'tap_press':
                    action_key = random.choice(tap_keys)
                    pyautogui.press(action_key)
                    action_log.append(f"Tap: {action_key.upper()}")
                
                elif chosen_action_type == 'right_click' and self.right_click_enabled.get():
                    pyautogui.rightClick(center_x, center_y, duration=0.0)
                    action_log.append("Right Click")
                    
                elif chosen_action_type == 'scroll' and self.scroll_enabled.get():
                    scroll_amount = random.choice([-5, 5]) 
                    pyautogui.scroll(scroll_amount)
                    action_log.append(f"Scroll: {'Up' if scroll_amount > 0 else 'Down'}")

            # Revert focus
            pyautogui.moveTo(original_mouse_x, original_mouse_y, duration=0.0) 
            win32gui.SetForegroundWindow(previous_hwnd)
            
            self.log_message(f"EXECUTED: {', '.join(action_log) if action_log else 'Mouse Jiggle Only'}")
            
        except Exception:
            win32gui.SetForegroundWindow(previous_hwnd)
            self.log_message(f"CRITICAL ERROR: Action failed. Reverting focus.", is_warning=True)

    def _calculate_weights(self, long_keys, tap_keys):
        """Calculates the proportional weights for the Custom Mode."""
        weights = {}
        long_weight = self.long_press_weight_var.get()
        
        if long_keys: weights['long_press'] = long_weight
        
        remaining_weight = 100 - weights.get('long_press', 0)
        
        other_actions = []
        if tap_keys: other_actions.append('tap_press')
        if self.right_click_enabled.get(): other_actions.append('right_click')
        if self.scroll_enabled.get(): other_actions.append('scroll')

        num_other = len(other_actions)
        if num_other > 0:
            weight_per_other = remaining_weight / num_other
            for action in other_actions:
                weights[action] = weight_per_other
        
        return weights

    # --- CONTROL FLOW ---

    def afk_loop(self):
        """The main threaded loop that controls the timing."""
        
        current_tab_index = self.notebook.index(self.notebook.select())
        is_fast_mode = (current_tab_index == 0)

        while self.is_running:
            if not self.is_paused:
                self.update_status("Running")
                
                if is_fast_mode:
                    self.execute_fast_action()
                else:
                    self.execute_custom_action()
                
                wait_time = random.uniform(self.min_delay_var.get(), self.max_delay_var.get())
                self.log_message(f"Next action in: {wait_time:.2f} seconds.")
                time.sleep(wait_time)
            else:
                time.sleep(1) 
        
        self.update_status("Stopped")

    def start_afk(self):
        """Starts the AFK process."""
        if not self.target_window_title.get():
            messagebox.showerror("Error", "Please select a target application window.")
            return

        if not self.is_running:
            self.is_running = True
            self.is_paused = False
            self.afk_thread = threading.Thread(target=self.afk_loop, daemon=True)
            self.afk_thread.start()
            self.log_message("Tool STARTED. Active Mode: " + self.notebook.tab(self.notebook.select(), "text"))
        elif self.is_paused:
            self.is_paused = False
            self.log_message("Tool RESUMED.")
            
        self.update_status("Running")

    def pause_afk(self):
        self.is_paused = True
        self.log_message("Tool PAUSED.")
        self.update_status("Paused")

    def stop_afk(self):
        if self.is_running:
            self.is_running = False
            self.is_paused = False
            self.log_message("Tool STOPPED.")
            
    def on_closing(self):
        """Handles the window being closed."""
        if self.is_running:
            self.stop_afk()
            time.sleep(0.1) 
        self.master.destroy()

    def refresh_windows(self):
        """Fetches and populates the list of all unique, non-empty window titles."""
        all_titles = gw.getAllTitles()
        valid_titles = sorted(list(set(title for title in all_titles if title)))
        self.window_selector['values'] = valid_titles
        
        current_title = self.target_window_title.get()
        if current_title not in valid_titles and valid_titles:
             self.target_window_title.set(valid_titles[0])


if __name__ == '__main__':
    root = tk.Tk()
    app = AfkGuiApp(root)
    root.mainloop()