import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox, font
import pyautogui
import time
import threading
import json
import os
from datetime import datetime
import keyboard
import random
from tkinter import colorchooser


class AutoTypingSoftware:
    def __init__(self, root):
        self.root = root
        self.root.title("Chori in Exam Software")
        self.root.geometry("900x700")
        self.root.configure(bg='#2b2b2b')

        # Variables
        self.typing_active = False
        self.speed_var = tk.DoubleVar(value=300)  # WPM
        self.accuracy_var = tk.DoubleVar(value=300)  # Accuracy percentage
        self.delay_var = tk.DoubleVar(value=5)  # Start delay in seconds
        self.loop_var = tk.BooleanVar(value=False)
        self.loop_count_var = tk.IntVar(value=1)
        self.current_loop = 0
        self.pause_var = tk.BooleanVar(value=False)
        self.selected_file = None
        self.hotkey_enabled = tk.BooleanVar(value=True)
        self.start_hotkey = "F6"
        self.stop_hotkey = "F7"
        self.pause_hotkey = "F8"
        self.custom_font = None
        self.text_content = ""

        # Initialize hotkey listener
        self.setup_hotkeys()

        # Create GUI
        self.setup_gui()

    def setup_hotkeys(self):
        """Setup hotkey listeners"""
        try:
            keyboard.add_hotkey(self.start_hotkey.lower(), self.start_typing_hotkey)
            keyboard.add_hotkey(self.stop_hotkey.lower(), self.stop_typing_hotkey)
            keyboard.add_hotkey(self.pause_hotkey.lower(), self.toggle_pause_hotkey)
        except:
            pass

    def start_typing_hotkey(self):
        if self.hotkey_enabled.get() and not self.typing_active:
            self.start_typing()

    def stop_typing_hotkey(self):
        if self.hotkey_enabled.get() and self.typing_active:
            self.stop_typing()

    def toggle_pause_hotkey(self):
        if self.hotkey_enabled.get() and self.typing_active:
            self.toggle_pause()

    def setup_gui(self):
        # Main container
        main_container = tk.Frame(self.root, bg='#2b2b2b')
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Header
        header_frame = tk.Frame(main_container, bg='#2b2b2b')
        header_frame.pack(fill=tk.X, pady=(0, 10))

        title_label = tk.Label(
            header_frame,
            text="Auto Typing Software",
            font=("Arial", 20, "bold"),
            fg='#4CAF50',
            bg='#2b2b2b'
        )
        title_label.pack()

        # Main content area
        content_frame = tk.Frame(main_container, bg='#2b2b2b')
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Left panel for text input
        left_panel = tk.Frame(content_frame, bg='#3c3c3c', relief=tk.RAISED, borderwidth=1)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # Right panel for controls
        right_panel = tk.Frame(content_frame, bg='#3c3c3c', relief=tk.RAISED, borderwidth=1)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0))

        # Text input area
        text_frame = tk.Frame(left_panel, bg='#3c3c3c')
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        text_label = tk.Label(
            text_frame,
            text="Text to Type:",
            font=("Arial", 12, "bold"),
            fg='white',
            bg='#3c3c3c'
        )
        text_label.pack(anchor=tk.W)

        # Text area with scrollbar
        text_container = tk.Frame(text_frame, bg='#3c3c3c')
        text_container.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self.text_area = scrolledtext.ScrolledText(
            text_container,
            wrap=tk.WORD,
            font=("Consolas", 11),
            bg='#1e1e1e',
            fg='white',
            insertbackground='white',
            relief=tk.FLAT,
            height=15
        )
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Control buttons for text
        text_btn_frame = tk.Frame(text_frame, bg='#3c3c3c')
        text_btn_frame.pack(fill=tk.X, pady=(5, 0))

        tk.Button(
            text_btn_frame,
            text="Load Text File",
            command=self.load_text_file,
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            text_btn_frame,
            text="Clear Text",
            command=self.clear_text,
            bg='#f44336',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5
        ).pack(side=tk.LEFT)

        # Status bar
        self.status_bar = tk.Label(
            left_panel,
            text="Ready",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg='#1e1e1e',
            fg='#4CAF50',
            font=("Arial", 10)
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))

        # Controls panel
        controls_label = tk.Label(
            right_panel,
            text="Typing Controls",
            font=("Arial", 14, "bold"),
            fg='white',
            bg='#3c3c3c'
        )
        controls_label.pack(pady=(10, 20))

        # Speed control
        speed_frame = tk.Frame(right_panel, bg='#3c3c3c')
        speed_frame.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(
            speed_frame,
            text=f"Speed: {self.speed_var.get()} WPM",
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        ).pack(anchor=tk.W)

        speed_slider = ttk.Scale(
            speed_frame,
            from_=10,
            to=500,
            variable=self.speed_var,
            orient=tk.HORIZONTAL,
            command=self.update_speed_label
        )
        speed_slider.pack(fill=tk.X, pady=(5, 0))

        # Accuracy control
        accuracy_frame = tk.Frame(right_panel, bg='#3c3c3c')
        accuracy_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(
            accuracy_frame,
            text=f"Accuracy: {self.accuracy_var.get()}%",
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        ).pack(anchor=tk.W)

        accuracy_slider = ttk.Scale(
            accuracy_frame,
            from_=0,
            to=100,
            variable=self.accuracy_var,
            orient=tk.HORIZONTAL,
            command=self.update_accuracy_label
        )
        accuracy_slider.pack(fill=tk.X, pady=(5, 0))

        # Delay control
        delay_frame = tk.Frame(right_panel, bg='#3c3c3c')
        delay_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(
            delay_frame,
            text=f"Start Delay: {self.delay_var.get()} seconds",
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        ).pack(anchor=tk.W)

        delay_slider = ttk.Scale(
            delay_frame,
            from_=0,
            to=30,
            variable=self.delay_var,
            orient=tk.HORIZONTAL,
            command=self.update_delay_label
        )
        delay_slider.pack(fill=tk.X, pady=(5, 0))

        # Loop controls
        loop_frame = tk.Frame(right_panel, bg='#3c3c3c')
        loop_frame.pack(fill=tk.X, padx=20, pady=10)

        loop_check = tk.Checkbutton(
            loop_frame,
            text="Enable Looping",
            variable=self.loop_var,
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c',
            selectcolor='#2b2b2b',
            activebackground='#3c3c3c',
            activeforeground='white'
        )
        loop_check.pack(anchor=tk.W)

        loop_count_frame = tk.Frame(loop_frame, bg='#3c3c3c')
        loop_count_frame.pack(fill=tk.X, pady=(5, 0))

        tk.Label(
            loop_count_frame,
            text="Loop Count:",
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        ).pack(side=tk.LEFT)

        loop_spinbox = tk.Spinbox(
            loop_count_frame,
            from_=1,
            to=999,
            textvariable=self.loop_count_var,
            width=8,
            bg='#1e1e1e',
            fg='white',
            insertbackground='white',
            relief=tk.FLAT
        )
        loop_spinbox.pack(side=tk.LEFT, padx=(5, 0))

        # Hotkey controls
        hotkey_frame = tk.Frame(right_panel, bg='#3c3c3c')
        hotkey_frame.pack(fill=tk.X, padx=20, pady=10)

        hotkey_check = tk.Checkbutton(
            hotkey_frame,
            text="Enable Hotkeys",
            variable=self.hotkey_enabled,
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c',
            selectcolor='#2b2b2b',
            activebackground='#3c3c3c',
            activeforeground='white'
        )
        hotkey_check.pack(anchor=tk.W)

        hotkey_info = tk.Label(
            hotkey_frame,
            text=f"Start: {self.start_hotkey} | Stop: {self.stop_hotkey} | Pause: {self.pause_hotkey}",
            font=("Arial", 9),
            fg='#888',
            bg='#3c3c3c'
        )
        hotkey_info.pack(anchor=tk.W, pady=(5, 0))

        # Control buttons
        btn_frame = tk.Frame(right_panel, bg='#3c3c3c')
        btn_frame.pack(fill=tk.X, padx=20, pady=20)

        self.start_btn = tk.Button(
            btn_frame,
            text="Start Typing (F6)",
            command=self.start_typing,
            bg='#4CAF50',
            fg='white',
            font=("Arial", 12, "bold"),
            relief=tk.FLAT,
            padx=30,
            pady=10,
            width=15
        )
        self.start_btn.pack(pady=5)

        self.pause_btn = tk.Button(
            btn_frame,
            text="Pause/Resume (F8)",
            command=self.toggle_pause,
            bg='#FF9800',
            fg='white',
            font=("Arial", 12, "bold"),
            relief=tk.FLAT,
            padx=30,
            pady=10,
            width=15,
            state=tk.DISABLED
        )
        self.pause_btn.pack(pady=5)

        self.stop_btn = tk.Button(
            btn_frame,
            text="Stop Typing (F7)",
            command=self.stop_typing,
            bg='#f44336',
            fg='white',
            font=("Arial", 12, "bold"),
            relief=tk.FLAT,
            padx=30,
            pady=10,
            width=15,
            state=tk.DISABLED
        )
        self.stop_btn.pack(pady=5)

        # Progress frame
        progress_frame = tk.Frame(right_panel, bg='#3c3c3c')
        progress_frame.pack(fill=tk.X, padx=20, pady=(0, 10))

        tk.Label(
            progress_frame,
            text="Progress:",
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        ).pack(anchor=tk.W)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient=tk.HORIZONTAL,
            length=200,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

        self.progress_label = tk.Label(
            progress_frame,
            text="0%",
            font=("Arial", 9),
            fg='white',
            bg='#3c3c3c'
        )
        self.progress_label.pack()

        # Loop counter
        self.loop_label = tk.Label(
            progress_frame,
            text="Loop: 0/0",
            font=("Arial", 9),
            fg='#888',
            bg='#3c3c3c'
        )
        self.loop_label.pack()

    def update_speed_label(self, event=None):
        self.root.after(100, lambda: self.update_label("speed"))

    def update_accuracy_label(self, event=None):
        self.root.after(100, lambda: self.update_label("accuracy"))

    def update_delay_label(self, event=None):
        self.root.after(100, lambda: self.update_label("delay"))

    def update_label(self, label_type):
        if label_type == "speed":
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Frame):
                            for grandchild in child.winfo_children():
                                if isinstance(grandchild, tk.Label) and grandchild.cget("text").startswith("Speed:"):
                                    grandchild.config(text=f"Speed: {self.speed_var.get():.0f} WPM")
        elif label_type == "accuracy":
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Frame):
                            for grandchild in child.winfo_children():
                                if isinstance(grandchild, tk.Label) and grandchild.cget("text").startswith("Accuracy:"):
                                    grandchild.config(text=f"Accuracy: {self.accuracy_var.get():.0f}%")
        elif label_type == "delay":
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Frame):
                            for grandchild in child.winfo_children():
                                if isinstance(grandchild, tk.Label) and grandchild.cget("text").startswith(
                                        "Start Delay:"):
                                    grandchild.config(text=f"Start Delay: {self.delay_var.get():.0f} seconds")

    def load_text_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Text File",
            filetypes=[
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    self.text_area.delete(1.0, tk.END)
                    self.text_area.insert(1.0, content)
                    self.selected_file = file_path
                    self.update_status(f"Loaded: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")

    def clear_text(self):
        self.text_area.delete(1.0, tk.END)
        self.selected_file = None
        self.update_status("Text cleared")

    def start_typing(self):
        if self.typing_active:
            return

        text = self.text_area.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("Warning", "Please enter some text to type!")
            return

        self.text_content = text
        self.typing_active = True
        self.current_loop = 0
        self.pause_var.set(False)

        # Update UI
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)

        # Start typing thread
        self.typing_thread = threading.Thread(target=self.type_text, daemon=True)
        self.typing_thread.start()

        self.update_status("Typing started - Switch to target window")

    def toggle_pause(self):
        self.pause_var.set(not self.pause_var.get())
        if self.pause_var.get():
            self.update_status("Typing paused")
        else:
            self.update_status("Typing resumed")

    def stop_typing(self):
        if self.typing_active:
            self.typing_active = False
            self.pause_var.set(False)

            # Reset UI
            self.start_btn.config(state=tk.NORMAL)
            self.pause_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.DISABLED)
            self.progress_bar['value'] = 0
            self.progress_label.config(text="0%")

            self.update_status("Typing stopped")

    def type_text(self):
        # Wait for delay
        delay = self.delay_var.get()
        if delay > 0:
            for i in range(int(delay * 10)):
                if not self.typing_active:
                    return
                time.sleep(0.1)
                if self.pause_var.get():
                    while self.pause_var.get() and self.typing_active:
                        time.sleep(0.1)

        # Calculate delay between characters (WPM to seconds per character)
        wpm = self.speed_var.get()
        if wpm > 0:
            char_delay = 60 / (wpm * 5)  # Average 5 characters per word
        else:
            char_delay = 0

        # Get loop count
        loop_count = self.loop_count_var.get() if self.loop_var.get() else 1

        for loop in range(loop_count):
            if not self.typing_active:
                break

            self.current_loop = loop + 1
            self.root.after(0, self.update_loop_label)

            text = self.text_content
            total_chars = len(text)

            for i, char in enumerate(text):
                if not self.typing_active:
                    break

                # Handle pause
                if self.pause_var.get():
                    while self.pause_var.get() and self.typing_active:
                        time.sleep(0.1)

                # Simulate accuracy errors
                accuracy = self.accuracy_var.get() / 100.0
                if random.random() > accuracy:
                    # Type wrong character
                    wrong_char = chr(ord(char) + 1) if char != 'z' and char != 'Z' else chr(ord(char) - 1)
                    pyautogui.write(wrong_char)
                    time.sleep(char_delay * 0.5)
                    pyautogui.press('backspace')
                    time.sleep(char_delay * 0.5)

                # Type the character
                pyautogui.write(char)

                # Update progress
                progress = (i + 1) / total_chars * 100
                self.root.after(0, lambda p=progress: self.update_progress(p))

                # Delay between characters
                time.sleep(char_delay)

            # Don't add extra delay after last loop
            if loop < loop_count - 1:
                time.sleep(1)  # Small delay between loops

        # Reset when done
        self.root.after(0, self.stop_typing)

    def update_progress(self, value):
        self.progress_bar['value'] = value
        self.progress_label.config(text=f"{value:.1f}%")

    def update_loop_label(self):
        loop_count = self.loop_count_var.get() if self.loop_var.get() else 1
        self.loop_label.config(text=f"Loop: {self.current_loop}/{loop_count}")

    def update_status(self, message):
        self.status_bar.config(text=message)
        print(f"Status: {message}")


def main():
    root = tk.Tk()
    app = AutoTypingSoftware(root)

    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()


if __name__ == "__main__":
    main()