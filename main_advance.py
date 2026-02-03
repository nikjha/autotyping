import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox, font, simpledialog, colorchooser
import pyautogui
import pyperclip
import time
import threading
import json
import os
from datetime import datetime, timedelta
import keyboard
import random
import schedule
import winsound
import csv
import re
from PIL import Image, ImageTk
import sqlite3
from dataclasses import dataclass
from typing import List, Dict, Optional
import pickle
import webbrowser


@dataclass
class TypingSession:
    id: int
    name: str
    text: str
    settings: Dict
    created_at: str


@dataclass
class TypingTask:
    id: int
    name: str
    text: str
    settings: Dict
    schedule_time: str
    enabled: bool


class ExtendedAutoTypingSoftware:
    def __init__(self, root):
        self.root = root
        self.root.title("Chori Se Typing Karna Hai")
        self.root.geometry("1100x800")
        self.root.configure(bg='#1e1e1e')

        # Database setup
        self.setup_database()

        # Variables
        self.typing_active = False
        self.speed_var = tk.DoubleVar(value=200)
        self.accuracy_var = tk.DoubleVar(value=200)
        self.delay_var = tk.DoubleVar(value=5)
        self.loop_var = tk.BooleanVar(value=False)
        self.loop_count_var = tk.IntVar(value=1)
        self.current_loop = 0
        self.pause_var = tk.BooleanVar(value=False)
        self.selected_file = None
        self.hotkey_enabled = tk.BooleanVar(value=True)
        self.start_hotkey = tk.StringVar(value="F6")
        self.stop_hotkey = tk.StringVar(value="F7")
        self.pause_hotkey = tk.StringVar(value="F8")
        self.text_content = ""
        self.sessions = []
        self.scheduled_tasks = []
        self.profiles = []
        self.current_profile = None

        # Enhanced variables
        self.random_delay_var = tk.BooleanVar(value=False)
        self.min_delay_var = tk.DoubleVar(value=0.05)
        self.max_delay_var = tk.DoubleVar(value=0.2)
        self.special_chars_var = tk.BooleanVar(value=True)
        self.cursor_pos_var = tk.StringVar(value="Current")
        self.beep_on_complete_var = tk.BooleanVar(value=True)
        self.auto_start_var = tk.BooleanVar(value=False)
        self.save_history_var = tk.BooleanVar(value=True)
        self.auto_save_var = tk.BooleanVar(value=False)

        # Statistics
        self.total_chars_typed = 0
        self.total_words_typed = 0
        self.total_sessions = 0
        self.start_time = None

        # Initialize
        self.load_sessions()
        self.load_profiles()
        self.load_scheduled_tasks()
        self.setup_hotkeys()

        # Create GUI
        self.setup_gui()

        # Start scheduler thread
        self.schedule_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.schedule_thread.start()

    def setup_database(self):
        """Initialize database for storing sessions and tasks"""
        self.db_conn = sqlite3.connect('typing_data.db', check_same_thread=False)
        self.db_cursor = self.db_conn.cursor()

        # Create tables
        self.db_cursor.execute('''
            CREATE TABLE IF NOT EXISTS sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                text TEXT,
                settings TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        self.db_cursor.execute('''
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                text TEXT,
                settings TEXT,
                schedule_time TEXT,
                enabled BOOLEAN DEFAULT 1
            )
        ''')

        self.db_cursor.execute('''
            CREATE TABLE IF NOT EXISTS profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                settings TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        self.db_cursor.execute('''
            CREATE TABLE IF NOT EXISTS statistics (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE,
                chars_typed INTEGER,
                words_typed INTEGER,
                sessions_completed INTEGER
            )
        ''')

        self.db_conn.commit()

    def load_sessions(self):
        """Load saved typing sessions from database"""
        try:
            self.db_cursor.execute("SELECT * FROM sessions ORDER BY created_at DESC LIMIT 50")
            rows = self.db_cursor.fetchall()
            self.sessions = [TypingSession(*row) for row in rows]
        except:
            self.sessions = []

    def load_profiles(self):
        """Load saved profiles"""
        try:
            self.db_cursor.execute("SELECT * FROM profiles ORDER BY name")
            rows = self.db_cursor.fetchall()
            self.profiles = rows
        except:
            self.profiles = []

    def load_scheduled_tasks(self):
        """Load scheduled tasks"""
        try:
            self.db_cursor.execute("SELECT * FROM tasks WHERE enabled = 1")
            rows = self.db_cursor.fetchall()
            self.scheduled_tasks = [TypingTask(*row) for row in rows]
        except:
            self.scheduled_tasks = []

    def setup_hotkeys(self):
        """Setup hotkey listeners"""
        try:
            keyboard.unhook_all()
            keyboard.add_hotkey(self.start_hotkey.get().lower(), self.start_typing_hotkey)
            keyboard.add_hotkey(self.stop_hotkey.get().lower(), self.stop_typing_hotkey)
            keyboard.add_hotkey(self.pause_hotkey.get().lower(), self.toggle_pause_hotkey)
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
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Main typing tab
        self.setup_typing_tab()

        # Sessions tab
        self.setup_sessions_tab()

        # Scheduler tab
        self.setup_scheduler_tab()

        # Profiles tab
        self.setup_profiles_tab()

        # Statistics tab
        self.setup_statistics_tab()

        # Settings tab
        self.setup_settings_tab()

        # Status bar
        self.setup_status_bar()

    def setup_typing_tab(self):
        """Setup main typing interface"""
        typing_frame = tk.Frame(self.notebook, bg='#2b2b2b')
        self.notebook.add(typing_frame, text="Typing")

        # Top control panel
        top_panel = tk.Frame(typing_frame, bg='#2b2b2b')
        top_panel.pack(fill=tk.X, padx=10, pady=10)

        # Quick controls
        quick_frame = tk.Frame(top_panel, bg='#3c3c3c', relief=tk.RAISED, borderwidth=1)
        quick_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        tk.Label(
            quick_frame,
            text="Quick Actions",
            font=("Arial", 11, "bold"),
            fg='white',
            bg='#3c3c3c'
        ).pack(pady=5)

        quick_btn_frame = tk.Frame(quick_frame, bg='#3c3c3c')
        quick_btn_frame.pack(pady=5)

        tk.Button(
            quick_btn_frame,
            text="📋 Paste",
            command=self.paste_from_clipboard,
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=2)

        tk.Button(
            quick_btn_frame,
            text="📄 Import",
            command=self.import_text,
            bg='#673AB7',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=2)

        tk.Button(
            quick_btn_frame,
            text="💾 Save",
            command=self.save_session,
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=2)

        # Main content
        content_frame = tk.Frame(typing_frame, bg='#2b2b2b')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Left panel - Text editor
        left_panel = tk.Frame(content_frame, bg='#2b2b2b')
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Text editor with line numbers
        editor_frame = tk.Frame(left_panel, bg='#1e1e1e')
        editor_frame.pack(fill=tk.BOTH, expand=True)

        # Line numbers
        self.line_numbers = tk.Text(
            editor_frame,
            width=4,
            padx=5,
            takefocus=0,
            border=0,
            background='#252525',
            foreground='#858585',
            state='disabled'
        )
        self.line_numbers.pack(side=tk.LEFT, fill=tk.Y)

        # Text area
        self.text_area = scrolledtext.ScrolledText(
            editor_frame,
            wrap=tk.WORD,
            font=("Consolas", 11),
            bg='#1e1e1e',
            fg='#d4d4d4',
            insertbackground='white',
            relief=tk.FLAT,
            undo=True
        )
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bind events for line numbers
        self.text_area.bind('<KeyRelease>', self.update_line_numbers)
        self.text_area.bind('<MouseWheel>', self.update_line_numbers)

        # Text statistics
        text_stats_frame = tk.Frame(left_panel, bg='#2b2b2b')
        text_stats_frame.pack(fill=tk.X, pady=(5, 0))

        self.text_stats_label = tk.Label(
            text_stats_frame,
            text="Characters: 0 | Words: 0 | Lines: 0",
            font=("Arial", 9),
            fg='#888',
            bg='#2b2b2b'
        )
        self.text_stats_label.pack()

        # Right panel - Controls
        right_panel = tk.Frame(content_frame, bg='#2b2b2b', width=300)
        right_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        right_panel.pack_propagate(False)

        # Speed control
        control_frame = tk.LabelFrame(
            right_panel,
            text="Typing Controls",
            font=("Arial", 10, "bold"),
            bg='#3c3c3c',
            fg='white',
            relief=tk.RAISED,
            borderwidth=2
        )
        control_frame.pack(fill=tk.X, padx=5, pady=5)

        self.create_slider(control_frame, "Speed (WPM):", self.speed_var, 10, 500, 1)
        self.create_slider(control_frame, "Accuracy (%):", self.accuracy_var, 0, 100, 0)
        self.create_slider(control_frame, "Delay (s):", self.delay_var, 0, 30, 1)

        # Advanced options
        adv_frame = tk.LabelFrame(
            right_panel,
            text="Advanced Options",
            font=("Arial", 10, "bold"),
            bg='#3c3c3c',
            fg='white',
            relief=tk.RAISED,
            borderwidth=2
        )
        adv_frame.pack(fill=tk.X, padx=5, pady=5)

        self.create_checkbox(adv_frame, "Random delay between keystrokes", self.random_delay_var)
        self.create_checkbox(adv_frame, "Handle special characters", self.special_chars_var)
        self.create_checkbox(adv_frame, "Beep on completion", self.beep_on_complete_var)
        self.create_checkbox(adv_frame, "Auto-start after delay", self.auto_start_var)
        self.create_checkbox(adv_frame, "Save typing history", self.save_history_var)

        # Cursor position
        cursor_frame = tk.Frame(adv_frame, bg='#3c3c3c')
        cursor_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(
            cursor_frame,
            text="Cursor Start:",
            font=("Arial", 9),
            fg='white',
            bg='#3c3c3c'
        ).pack(side=tk.LEFT)

        cursor_combo = ttk.Combobox(
            cursor_frame,
            textvariable=self.cursor_pos_var,
            values=["Current", "Beginning", "End", "Custom"],
            state="readonly",
            width=12
        )
        cursor_combo.pack(side=tk.RIGHT)

        # Loop controls
        loop_frame = tk.Frame(adv_frame, bg='#3c3c3c')
        loop_frame.pack(fill=tk.X, padx=10, pady=5)

        loop_check = tk.Checkbutton(
            loop_frame,
            text="Loop",
            variable=self.loop_var,
            font=("Arial", 9),
            fg='white',
            bg='#3c3c3c',
            selectcolor='#2b2b2b'
        )
        loop_check.pack(side=tk.LEFT)

        loop_spin = tk.Spinbox(
            loop_frame,
            from_=1,
            to=999,
            textvariable=self.loop_count_var,
            width=8,
            bg='#1e1e1e',
            fg='white'
        )
        loop_spin.pack(side=tk.RIGHT)

        # Action buttons
        btn_frame = tk.Frame(right_panel, bg='#2b2b2b')
        btn_frame.pack(fill=tk.X, padx=5, pady=10)

        self.start_btn = self.create_action_button(btn_frame, "▶ Start (F6)", self.start_typing, '#4CAF50')
        self.pause_btn = self.create_action_button(btn_frame, "⏸ Pause (F8)", self.toggle_pause, '#FF9800', tk.DISABLED)
        self.stop_btn = self.create_action_button(btn_frame, "⏹ Stop (F7)", self.stop_typing, '#f44336', tk.DISABLED)

        # Progress
        progress_frame = tk.LabelFrame(
            right_panel,
            text="Progress",
            font=("Arial", 10, "bold"),
            bg='#3c3c3c',
            fg='white'
        )
        progress_frame.pack(fill=tk.X, padx=5, pady=5)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient=tk.HORIZONTAL,
            length=200,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, padx=10, pady=5)

        self.progress_label = tk.Label(
            progress_frame,
            text="0%",
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        )
        self.progress_label.pack()

        self.loop_label = tk.Label(
            progress_frame,
            text="Loop: 0/0",
            font=("Arial", 9),
            fg='#888',
            bg='#3c3c3c'
        )
        self.loop_label.pack()

        # Real-time stats
        stats_frame = tk.LabelFrame(
            right_panel,
            text="Current Session",
            font=("Arial", 10, "bold"),
            bg='#3c3c3c',
            fg='white'
        )
        stats_frame.pack(fill=tk.X, padx=5, pady=5)

        self.stats_labels = {}
        stats = [
            ("Chars Typed:", "0"),
            ("Words Typed:", "0"),
            ("Time Elapsed:", "00:00"),
            ("Speed:", "0 WPM"),
            ("Accuracy:", "100%")
        ]

        for label, value in stats:
            frame = tk.Frame(stats_frame, bg='#3c3c3c')
            frame.pack(fill=tk.X, padx=10, pady=2)

            tk.Label(
                frame,
                text=label,
                font=("Arial", 9),
                fg='#ccc',
                bg='#3c3c3c'
            ).pack(side=tk.LEFT)

            self.stats_labels[label] = tk.Label(
                frame,
                text=value,
                font=("Arial", 9, "bold"),
                fg='white',
                bg='#3c3c3c'
            )
            self.stats_labels[label].pack(side=tk.RIGHT)

    def setup_sessions_tab(self):
        """Setup saved sessions tab"""
        sessions_frame = tk.Frame(self.notebook, bg='#2b2b2b')
        self.notebook.add(sessions_frame, text="Sessions")

        # Header
        header_frame = tk.Frame(sessions_frame, bg='#2b2b2b')
        header_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(
            header_frame,
            text="Saved Typing Sessions",
            font=("Arial", 14, "bold"),
            fg='white',
            bg='#2b2b2b'
        ).pack(side=tk.LEFT)

        tk.Button(
            header_frame,
            text="Refresh",
            command=self.refresh_sessions,
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15
        ).pack(side=tk.RIGHT)

        # Sessions list
        list_frame = tk.Frame(sessions_frame, bg='#2b2b2b')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Treeview for sessions
        columns = ("ID", "Name", "Characters", "Created", "Actions")
        self.sessions_tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",
            height=15
        )

        # Configure style
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background="#1e1e1e",
                        foreground="white",
                        fieldbackground="#1e1e1e",
                        borderwidth=0)
        style.configure("Treeview.Heading",
                        background="#3c3c3c",
                        foreground="white",
                        relief="flat")

        for col in columns:
            self.sessions_tree.heading(col, text=col)
            self.sessions_tree.column(col, width=100)

        self.sessions_tree.column("Name", width=200)
        self.sessions_tree.column("Actions", width=150)

        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.sessions_tree.yview)
        self.sessions_tree.configure(yscrollcommand=scrollbar.set)
        self.sessions_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Populate sessions
        self.populate_sessions_tree()

        # Session preview
        preview_frame = tk.LabelFrame(
            sessions_frame,
            text="Session Preview",
            font=("Arial", 10, "bold"),
            bg='#3c3c3c',
            fg='white'
        )
        preview_frame.pack(fill=tk.BOTH, padx=10, pady=(0, 10), expand=False)

        self.session_preview = scrolledtext.ScrolledText(
            preview_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg='#1e1e1e',
            fg='white',
            height=6
        )
        self.session_preview.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Bind selection event
        self.sessions_tree.bind('<<TreeviewSelect>>', self.on_session_select)

    def setup_scheduler_tab(self):
        """Setup task scheduler tab"""
        scheduler_frame = tk.Frame(self.notebook, bg='#2b2b2b')
        self.notebook.add(scheduler_frame, text="Scheduler")

        # Task list
        task_list_frame = tk.Frame(scheduler_frame, bg='#2b2b2b')
        task_list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(
            task_list_frame,
            text="Scheduled Tasks",
            font=("Arial", 14, "bold"),
            fg='white',
            bg='#2b2b2b'
        ).pack(anchor=tk.W, pady=(0, 10))

        # Treeview for tasks
        columns = ("Name", "Schedule", "Status", "Actions")
        self.tasks_tree = ttk.Treeview(
            task_list_frame,
            columns=columns,
            show="headings",
            height=10
        )

        for col in columns:
            self.tasks_tree.heading(col, text=col)
            self.tasks_tree.column(col, width=120)

        self.tasks_tree.pack(fill=tk.BOTH, expand=True)

        # Task controls
        controls_frame = tk.Frame(task_list_frame, bg='#2b2b2b')
        controls_frame.pack(fill=tk.X, pady=(10, 0))

        tk.Button(
            controls_frame,
            text="Add Task",
            command=self.add_scheduled_task,
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            controls_frame,
            text="Edit Task",
            command=self.edit_scheduled_task,
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            controls_frame,
            text="Delete Task",
            command=self.delete_scheduled_task,
            bg='#f44336',
            fg='white',
            relief=tk.FLAT
        ).pack(side=tk.LEFT, padx=5)

        # Task editor
        editor_frame = tk.Frame(scheduler_frame, bg='#2b2b2b', width=400)
        editor_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(0, 10), pady=10)
        editor_frame.pack_propagate(False)

        tk.Label(
            editor_frame,
            text="Task Editor",
            font=("Arial", 14, "bold"),
            fg='white',
            bg='#2b2b2b'
        ).pack(anchor=tk.W, pady=(0, 10))

        # Form fields
        form_frame = tk.Frame(editor_frame, bg='#3c3c3c', relief=tk.RAISED, borderwidth=1)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Name field
        tk.Label(
            form_frame,
            text="Task Name:",
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        ).pack(anchor=tk.W, padx=10, pady=(10, 0))

        self.task_name_var = tk.StringVar()
        tk.Entry(
            form_frame,
            textvariable=self.task_name_var,
            bg='#1e1e1e',
            fg='white',
            insertbackground='white'
        ).pack(fill=tk.X, padx=10, pady=5)

        # Schedule time
        tk.Label(
            form_frame,
            text="Schedule Time (HH:MM):",
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        ).pack(anchor=tk.W, padx=10, pady=(10, 0))

        self.task_time_var = tk.StringVar(value="12:00")
        tk.Entry(
            form_frame,
            textvariable=self.task_time_var,
            bg='#1e1e1e',
            fg='white',
            insertbackground='white'
        ).pack(fill=tk.X, padx=10, pady=5)

        # Repeat options
        repeat_frame = tk.Frame(form_frame, bg='#3c3c3c')
        repeat_frame.pack(fill=tk.X, padx=10, pady=10)

        self.repeat_daily_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            repeat_frame,
            text="Daily",
            variable=self.repeat_daily_var,
            fg='white',
            bg='#3c3c3c'
        ).pack(side=tk.LEFT)

        self.repeat_weekly_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            repeat_frame,
            text="Weekly",
            variable=self.repeat_weekly_var,
            fg='white',
            bg='#3c3c3c'
        ).pack(side=tk.LEFT, padx=20)

        # Save button
        tk.Button(
            form_frame,
            text="Save Task",
            command=self.save_scheduled_task,
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            pady=8
        ).pack(fill=tk.X, padx=10, pady=20)

    def setup_profiles_tab(self):
        """Setup profiles tab"""
        profiles_frame = tk.Frame(self.notebook, bg='#2b2b2b')
        self.notebook.add(profiles_frame, text="Profiles")

        # Profiles list
        list_frame = tk.Frame(profiles_frame, bg='#2b2b2b')
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)

        tk.Label(
            list_frame,
            text="Saved Profiles",
            font=("Arial", 14, "bold"),
            fg='white',
            bg='#2b2b2b'
        ).pack(anchor=tk.W, pady=(0, 10))

        self.profiles_listbox = tk.Listbox(
            list_frame,
            bg='#1e1e1e',
            fg='white',
            selectbackground='#4CAF50',
            selectforeground='white',
            font=("Arial", 11),
            height=15
        )
        self.profiles_listbox.pack(fill=tk.BOTH, expand=True)

        # Profile controls
        profile_controls = tk.Frame(list_frame, bg='#2b2b2b')
        profile_controls.pack(fill=tk.X, pady=(10, 0))

        tk.Button(
            profile_controls,
            text="Load Profile",
            command=self.load_profile,
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            width=12
        ).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(
            profile_controls,
            text="Delete Profile",
            command=self.delete_profile,
            bg='#f44336',
            fg='white',
            relief=tk.FLAT,
            width=12
        ).pack(side=tk.LEFT)

        # Current profile display
        current_frame = tk.Frame(profiles_frame, bg='#2b2b2b', width=400)
        current_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(0, 10), pady=10)
        current_frame.pack_propagate(False)

        tk.Label(
            current_frame,
            text="Current Profile Settings",
            font=("Arial", 14, "bold"),
            fg='white',
            bg='#2b2b2b'
        ).pack(anchor=tk.W, pady=(0, 10))

        # Profile settings display
        settings_frame = tk.Frame(current_frame, bg='#3c3c3c', relief=tk.RAISED, borderwidth=1)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.profile_settings_text = scrolledtext.ScrolledText(
            settings_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg='#1e1e1e',
            fg='white',
            height=10
        )
        self.profile_settings_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Save current as profile
        tk.Button(
            current_frame,
            text="Save Current as Profile",
            command=self.save_current_profile,
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            pady=8
        ).pack(fill=tk.X, pady=(10, 0))

        # Populate profiles list
        self.populate_profiles_list()

    def setup_statistics_tab(self):
        """Setup statistics tab"""
        stats_frame = tk.Frame(self.notebook, bg='#2b2b2b')
        self.notebook.add(stats_frame, text="Statistics")

        # Overall statistics
        overall_frame = tk.LabelFrame(
            stats_frame,
            text="Overall Statistics",
            font=("Arial", 12, "bold"),
            bg='#3c3c3c',
            fg='white',
            padx=20,
            pady=20
        )
        overall_frame.pack(fill=tk.X, padx=10, pady=10)

        stats_grid = tk.Frame(overall_frame, bg='#3c3c3c')
        stats_grid.pack()

        stats_data = [
            ("Total Sessions:", "0"),
            ("Total Characters:", "0"),
            ("Total Words:", "0"),
            ("Average Speed:", "0 WPM"),
            ("Total Time:", "00:00:00")
        ]

        for i, (label, value) in enumerate(stats_data):
            row = i // 2
            col = (i % 2) * 2

            tk.Label(
                stats_grid,
                text=label,
                font=("Arial", 10),
                fg='#ccc',
                bg='#3c3c3c'
            ).grid(row=row, column=col, padx=10, pady=5, sticky=tk.W)

            tk.Label(
                stats_grid,
                text=value,
                font=("Arial", 10, "bold"),
                fg='white',
                bg='#3c3c3c'
            ).grid(row=row, column=col + 1, padx=10, pady=5, sticky=tk.W)

        # Recent sessions
        recent_frame = tk.LabelFrame(
            stats_frame,
            text="Recent Sessions",
            font=("Arial", 12, "bold"),
            bg='#3c3c3c',
            fg='white',
            padx=20,
            pady=20
        )
        recent_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Treeview for recent sessions
        columns = ("Date", "Characters", "Words", "Speed", "Duration")
        self.stats_tree = ttk.Treeview(
            recent_frame,
            columns=columns,
            show="headings",
            height=8
        )

        for col in columns:
            self.stats_tree.heading(col, text=col)
            self.stats_tree.column(col, width=100)

        self.stats_tree.pack(fill=tk.BOTH, expand=True)

        # Export button
        tk.Button(
            stats_frame,
            text="Export Statistics",
            command=self.export_statistics,
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            pady=8
        ).pack(side=tk.RIGHT, padx=10, pady=(0, 10))

    def setup_settings_tab(self):
        """Setup settings tab"""
        settings_frame = tk.Frame(self.notebook, bg='#2b2b2b')
        self.notebook.add(settings_frame, text="Settings")

        # Hotkey settings
        hotkey_frame = tk.LabelFrame(
            settings_frame,
            text="Hotkey Configuration",
            font=("Arial", 12, "bold"),
            bg='#3c3c3c',
            fg='white',
            padx=20,
            pady=20
        )
        hotkey_frame.pack(fill=tk.X, padx=10, pady=10)

        hotkey_check = tk.Checkbutton(
            hotkey_frame,
            text="Enable Hotkeys",
            variable=self.hotkey_enabled,
            font=("Arial", 10),
            fg='white',
            bg='#3c3c3c'
        )
        hotkey_check.pack(anchor=tk.W, pady=(0, 10))

        # Hotkey inputs
        hotkey_grid = tk.Frame(hotkey_frame, bg='#3c3c3c')
        hotkey_grid.pack()

        hotkeys = [
            ("Start Typing:", self.start_hotkey),
            ("Stop Typing:", self.stop_hotkey),
            ("Pause/Resume:", self.pause_hotkey)
        ]

        for i, (label, var) in enumerate(hotkeys):
            tk.Label(
                hotkey_grid,
                text=label,
                font=("Arial", 10),
                fg='white',
                bg='#3c3c3c'
            ).grid(row=i, column=0, padx=10, pady=5, sticky=tk.W)

            tk.Entry(
                hotkey_grid,
                textvariable=var,
                width=10,
                bg='#1e1e1e',
                fg='white'
            ).grid(row=i, column=1, padx=10, pady=5)

        # Save hotkeys button
        tk.Button(
            hotkey_grid,
            text="Save Hotkeys",
            command=self.save_hotkeys,
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=20
        ).grid(row=3, column=0, columnspan=2, pady=10)

        # Application settings
        app_frame = tk.LabelFrame(
            settings_frame,
            text="Application Settings",
            font=("Arial", 12, "bold"),
            bg='#3c3c3c',
            fg='white',
            padx=20,
            pady=20
        )
        app_frame.pack(fill=tk.X, padx=10, pady=10)

        self.create_checkbox(app_frame, "Auto-save sessions", self.auto_save_var)
        self.create_checkbox(app_frame, "Show confirmation dialogs", tk.BooleanVar(value=True))
        self.create_checkbox(app_frame, "Minimize to system tray", tk.BooleanVar(value=False))
        self.create_checkbox(app_frame, "Start with Windows", tk.BooleanVar(value=False))

        # About section
        about_frame = tk.LabelFrame(
            settings_frame,
            text="About",
            font=("Arial", 12, "bold"),
            bg='#3c3c3c',
            fg='white',
            padx=20,
            pady=20
        )
        about_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(
            about_frame,
            text="Professional Auto Typing Suite v2.0\n\n"
                 "Features:\n"
                 "• Advanced typing automation\n"
                 "• Session management\n"
                 "• Task scheduling\n"
                 "• Profiles system\n"
                 "• Statistics tracking\n\n"
                 "© 2024 Auto Typing Software",
            font=("Arial", 9),
            fg='#ccc',
            bg='#3c3c3c',
            justify=tk.LEFT
        ).pack(anchor=tk.W)

        # Help button
        tk.Button(
            about_frame,
            text="📖 User Guide",
            command=self.open_user_guide,
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT
        ).pack(side=tk.LEFT, pady=(10, 0))

        tk.Button(
            about_frame,
            text="🐛 Report Issue",
            command=self.report_issue,
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT
        ).pack(side=tk.LEFT, padx=10, pady=(10, 0))

    def setup_status_bar(self):
        """Setup status bar at bottom"""
        self.status_bar = tk.Label(
            self.root,
            text="Ready",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg='#1e1e1e',
            fg='#4CAF50',
            font=("Arial", 10)
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def create_slider(self, parent, label, variable, from_, to, resolution):
        """Create a labeled slider"""
        frame = tk.Frame(parent, bg='#3c3c3c')
        frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(
            frame,
            text=f"{label} {variable.get():.0f}",
            font=("Arial", 9),
            fg='white',
            bg='#3c3c3c'
        ).pack(anchor=tk.W)

        slider = ttk.Scale(
            frame,
            from_=from_,
            to=to,
            variable=variable,
            orient=tk.HORIZONTAL,
            command=lambda v: self.update_slider_label(label, variable)
        )
        slider.pack(fill=tk.X, pady=(5, 0))

    def create_checkbox(self, parent, text, variable):
        """Create a checkbox"""
        check = tk.Checkbutton(
            parent,
            text=text,
            variable=variable,
            font=("Arial", 9),
            fg='white',
            bg='#3c3c3c',
            selectcolor='#2b2b2b'
        )
        check.pack(anchor=tk.W, padx=10, pady=2)

    def create_action_button(self, parent, text, command, color, state=tk.NORMAL):
        """Create an action button"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg='white',
            font=("Arial", 11, "bold"),
            relief=tk.FLAT,
            pady=8,
            state=state
        )
        btn.pack(fill=tk.X, pady=2)
        return btn

    def update_slider_label(self, label, variable):
        """Update slider label with current value"""
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, tk.Frame):
                        for grandchild in child.winfo_children():
                            if isinstance(grandchild, tk.Label) and grandchild.cget("text").startswith(
                                    label.split(":")[0]):
                                grandchild.config(text=f"{label.split(':')[0]}: {variable.get():.0f}")

    def update_line_numbers(self, event=None):
        """Update line numbers in text editor"""
        lines = self.text_area.get('1.0', 'end-1c').split('\n')
        line_numbers_text = '\n'.join(str(i) for i in range(1, len(lines) + 1))

        self.line_numbers.config(state='normal')
        self.line_numbers.delete('1.0', 'end')
        self.line_numbers.insert('1.0', line_numbers_text)
        self.line_numbers.config(state='disabled')

        # Update text statistics
        text = self.text_area.get('1.0', 'end-1c')
        chars = len(text)
        words = len(text.split())
        lines = len(text.split('\n'))
        self.text_stats_label.config(text=f"Characters: {chars} | Words: {words} | Lines: {lines}")

    def paste_from_clipboard(self):
        """Paste text from clipboard"""
        try:
            clipboard_text = pyperclip.paste()
            self.text_area.insert(tk.INSERT, clipboard_text)
            self.update_line_numbers()
            self.update_status("Text pasted from clipboard")
        except:
            self.update_status("Failed to paste from clipboard")

    def import_text(self):
        """Import text from various sources"""
        file_path = filedialog.askopenfilename(
            title="Import Text",
            filetypes=[
                ("Text files", "*.txt"),
                ("CSV files", "*.csv"),
                ("JSON files", "*.json"),
                ("All files", "*.*")
            ]
        )

        if file_path:
            try:
                if file_path.endswith('.csv'):
                    with open(file_path, 'r', encoding='utf-8') as file:
                        reader = csv.reader(file)
                        text = '\n'.join([','.join(row) for row in reader])
                elif file_path.endswith('.json'):
                    with open(file_path, 'r', encoding='utf-8') as file:
                        data = json.load(file)
                        text = json.dumps(data, indent=2)
                else:
                    with open(file_path, 'r', encoding='utf-8') as file:
                        text = file.read()

                self.text_area.delete('1.0', tk.END)
                self.text_area.insert('1.0', text)
                self.update_line_numbers()
                self.update_status(f"Imported: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import file: {str(e)}")

    def save_session(self):
        """Save current typing session"""
        name = simpledialog.askstring("Save Session", "Enter session name:")
        if name:
            text = self.text_area.get('1.0', 'end-1c')
            settings = {
                'speed': self.speed_var.get(),
                'accuracy': self.accuracy_var.get(),
                'delay': self.delay_var.get(),
                'loop': self.loop_var.get(),
                'loop_count': self.loop_count_var.get()
            }

            try:
                self.db_cursor.execute(
                    "INSERT INTO sessions (name, text, settings) VALUES (?, ?, ?)",
                    (name, text, json.dumps(settings))
                )
                self.db_conn.commit()
                self.load_sessions()
                self.populate_sessions_tree()
                self.update_status(f"Session '{name}' saved")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save session: {str(e)}")

    def refresh_sessions(self):
        """Refresh sessions list"""
        self.load_sessions()
        self.populate_sessions_tree()
        self.update_status("Sessions refreshed")

    def populate_sessions_tree(self):
        """Populate sessions treeview"""
        for item in self.sessions_tree.get_children():
            self.sessions_tree.delete(item)

        for session in self.sessions:
            chars = len(session.text)
            created = session.created_at[:19] if session.created_at else "N/A"
            self.sessions_tree.insert(
                "",
                "end",
                values=(
                    session.id,
                    session.name,
                    f"{chars:,}",
                    created,
                    "Load | Delete"
                )
            )

    def on_session_select(self, event):
        """Handle session selection"""
        selection = self.sessions_tree.selection()
        if selection:
            item = self.sessions_tree.item(selection[0])
            session_id = item['values'][0]

            # Find session
            for session in self.sessions:
                if session.id == session_id:
                    self.session_preview.delete('1.0', tk.END)
                    preview = session.text[:500] + ("..." if len(session.text) > 500 else "")
                    self.session_preview.insert('1.0', preview)
                    break

    def start_typing(self):
        """Start typing with enhanced features"""
        if self.typing_active:
            return

        text = self.text_area.get('1.0', 'end-1c').strip()
        if not text:
            messagebox.showwarning("Warning", "Please enter some text to type!")
            return

        self.text_content = text
        self.typing_active = True
        self.current_loop = 0
        self.pause_var.set(False)
        self.start_time = datetime.now()

        # Reset statistics
        self.total_chars_typed = 0
        self.total_words_typed = 0

        # Update UI
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)

        # Start typing thread
        self.typing_thread = threading.Thread(target=self.enhanced_type_text, daemon=True)
        self.typing_thread.start()

        self.update_status("Typing started - Switch to target window")

    def enhanced_type_text(self):
        """Enhanced typing algorithm with more features"""
        # Wait for delay
        delay = self.delay_var.get()
        if delay > 0:
            self.update_status(f"Starting in {delay} seconds...")
            for i in range(int(delay * 10)):
                if not self.typing_active:
                    return
                time.sleep(0.1)
                if self.pause_var.get():
                    while self.pause_var.get() and self.typing_active:
                        time.sleep(0.1)

        # Get loop count
        loop_count = self.loop_count_var.get() if self.loop_var.get() else 1

        for loop in range(loop_count):
            if not self.typing_active:
                break

            self.current_loop = loop + 1
            #self.root.after(0, self.update_loop_label)

            text = self.text_content
            total_chars = len(text)

            for i, char in enumerate(text):
                if not self.typing_active:
                    break

                # Handle pause
                if self.pause_var.get():
                    while self.pause_var.get() and self.typing_active:
                        time.sleep(0.1)

                # Calculate delay
                wpm = self.speed_var.get()
                if wpm > 0:
                    base_delay = 60 / (wpm * 5)
                else:
                    base_delay = 0

                # Add random delay if enabled
                if self.random_delay_var.get():
                    delay_factor = random.uniform(self.min_delay_var.get(), self.max_delay_var.get())
                    current_delay = base_delay + delay_factor
                else:
                    current_delay = base_delay

                # Handle special characters
                if self.special_chars_var.get():
                    self.type_special_char(char)
                else:
                    pyautogui.write(char)

                # Update statistics
                self.total_chars_typed += 1
                if char == ' ':
                    self.total_words_typed += 1

                # Update progress and stats
                progress = (i + 1) / total_chars * 100
                self.root.after(0, lambda p=progress: self.update_progress_and_stats(p))

                # Delay between characters
                if current_delay > 0:
                    time.sleep(current_delay)

            # Beep on completion if enabled
            if self.beep_on_complete_var.get() and loop == loop_count - 1:
                try:
                    winsound.Beep(1000, 200)
                    #pass
                except:
                    pass

            # Don't add extra delay after last loop
            if loop < loop_count - 1:
                time.sleep(1)

        # Save statistics
        if self.save_history_var.get():
            self.save_session_statistics()

        # Reset when done
        self.root.after(0, self.stop_typing)

    def type_special_char(self, char):
        """Handle special characters"""
        special_chars = {
            '\n': 'enter',
            '\t': 'tab',
            '\b': 'backspace',
            '\r': 'enter'
        }

        if char in special_chars:
            pyautogui.press(special_chars[char])
        else:
            pyautogui.write(char)

    def update_progress_and_stats(self, progress_value):
        """Update progress bar and statistics"""
        self.progress_bar['value'] = progress_value
        self.progress_label.config(text=f"{progress_value:.1f}%")

        # Update real-time statistics
        if self.start_time:
            elapsed = datetime.now() - self.start_time
            elapsed_str = str(elapsed).split('.')[0]

            if elapsed.total_seconds() > 0:
                speed = (self.total_chars_typed / 5) / (elapsed.total_seconds() / 60)
            else:
                speed = 0

            self.stats_labels["Chars Typed:"].config(text=f"{self.total_chars_typed:,}")
            self.stats_labels["Words Typed:"].config(text=f"{self.total_words_typed:,}")
            self.stats_labels["Time Elapsed:"].config(text=elapsed_str)
            self.stats_labels["Speed:"].config(text=f"{speed:.0f} WPM")

            # Calculate accuracy (simulated)
            accuracy = self.accuracy_var.get()
            self.stats_labels["Accuracy:"].config(text=f"{accuracy:.0f}%")

    def toggle_pause(self):
        self.pause_var.set(not self.pause_var.get())
        if self.pause_var.get():
            self.update_status("Typing paused")
            self.pause_btn.config(text="▶ Resume (F8)")
        else:
            self.update_status("Typing resumed")
            self.pause_btn.config(text="⏸ Pause (F8)")

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
            self.pause_btn.config(text="⏸ Pause (F8)")

            self.update_status("Typing completed")

    def save_session_statistics(self):
        """Save session statistics to database"""
        try:
            today = datetime.now().date()
            self.db_cursor.execute(
                "INSERT INTO statistics (date, chars_typed, words_typed, sessions_completed) VALUES (?, ?, ?, ?)",
                (today, self.total_chars_typed, self.total_words_typed, 1)
            )
            self.db_conn.commit()
        except:
            pass

    def add_scheduled_task(self):
        """Add a new scheduled task"""
        self.task_name_var.set("")
        self.task_time_var.set("12:00")
        self.repeat_daily_var.set(True)
        self.repeat_weekly_var.set(False)

    def edit_scheduled_task(self):
        """Edit selected scheduled task"""
        selection = self.tasks_tree.selection()
        if selection:
            # Implement edit functionality
            pass

    def delete_scheduled_task(self):
        """Delete selected scheduled task"""
        selection = self.tasks_tree.selection()
        if selection:
            # Implement delete functionality
            pass

    def save_scheduled_task(self):
        """Save scheduled task"""
        name = self.task_name_var.get().strip()
        time_str = self.task_time_var.get().strip()

        if not name or not time_str:
            messagebox.showwarning("Warning", "Please enter task name and time!")
            return

        try:
            # Validate time format
            datetime.strptime(time_str, "%H:%M")

            settings = {
                'speed': self.speed_var.get(),
                'accuracy': self.accuracy_var.get(),
                'delay': self.delay_var.get()
            }

            self.db_cursor.execute(
                """INSERT INTO tasks (name, text, settings, schedule_time, enabled) 
                   VALUES (?, ?, ?, ?, ?)""",
                (name, self.text_area.get('1.0', 'end-1c'), json.dumps(settings), time_str, 1)
            )
            self.db_conn.commit()

            self.update_status(f"Task '{name}' scheduled for {time_str}")
            messagebox.showinfo("Success", f"Task '{name}' has been scheduled!")
        except ValueError:
            messagebox.showerror("Error", "Invalid time format! Use HH:MM")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save task: {str(e)}")

    def run_scheduler(self):
        """Run task scheduler in background"""
        while True:
            try:
                now = datetime.now().strftime("%H:%M")
                self.db_cursor.execute(
                    "SELECT * FROM tasks WHERE schedule_time = ? AND enabled = 1",
                    (now,)
                )
                tasks = self.db_cursor.fetchall()

                for task in tasks:
                    # Execute task
                    self.execute_scheduled_task(task)

                time.sleep(60)  # Check every minute
            except:
                time.sleep(60)

    def execute_scheduled_task(self, task):
        """Execute a scheduled task"""
        try:
            # Update UI
            self.root.after(0, self.update_status, f"Executing scheduled task: {task[1]}")

            # Load task settings
            settings = json.loads(task[3])
            self.speed_var.set(settings.get('speed', 150))
            self.accuracy_var.set(settings.get('accuracy', 100))
            self.delay_var.set(settings.get('delay', 3))

            # Set text
            self.text_area.delete('1.0', tk.END)
            self.text_area.insert('1.0', task[2])

            # Start typing
            if not self.typing_active:
                self.start_typing()
        except:
            pass

    def populate_profiles_list(self):
        """Populate profiles listbox"""
        self.profiles_listbox.delete(0, tk.END)
        for profile in self.profiles:
            self.profiles_listbox.insert(tk.END, profile[1])

    def load_profile(self):
        """Load selected profile"""
        selection = self.profiles_listbox.curselection()
        if selection:
            profile_name = self.profiles_listbox.get(selection[0])

            for profile in self.profiles:
                if profile[1] == profile_name:
                    settings = json.loads(profile[2])

                    # Apply settings
                    self.speed_var.set(settings.get('speed', 150))
                    self.accuracy_var.set(settings.get('accuracy', 100))
                    self.delay_var.set(settings.get('delay', 3))
                    self.loop_var.set(settings.get('loop', False))
                    self.loop_count_var.set(settings.get('loop_count', 1))

                    # Display settings
                    self.profile_settings_text.delete('1.0', tk.END)
                    self.profile_settings_text.insert('1.0', json.dumps(settings, indent=2))

                    self.update_status(f"Profile '{profile_name}' loaded")
                    break

    def delete_profile(self):
        """Delete selected profile"""
        selection = self.profiles_listbox.curselection()
        if selection:
            profile_name = self.profiles_listbox.get(selection[0])

            if messagebox.askyesno("Confirm", f"Delete profile '{profile_name}'?"):
                try:
                    self.db_cursor.execute("DELETE FROM profiles WHERE name = ?", (profile_name,))
                    self.db_conn.commit()
                    self.load_profiles()
                    self.populate_profiles_list()
                    self.update_status(f"Profile '{profile_name}' deleted")
                except:
                    messagebox.showerror("Error", "Failed to delete profile")

    def save_current_profile(self):
        """Save current settings as profile"""
        name = simpledialog.askstring("Save Profile", "Enter profile name:")
        if name:
            settings = {
                'speed': self.speed_var.get(),
                'accuracy': self.accuracy_var.get(),
                'delay': self.delay_var.get(),
                'loop': self.loop_var.get(),
                'loop_count': self.loop_count_var.get(),
                'random_delay': self.random_delay_var.get(),
                'special_chars': self.special_chars_var.get(),
                'beep_on_complete': self.beep_on_complete_var.get()
            }

            try:
                self.db_cursor.execute(
                    "INSERT INTO profiles (name, settings) VALUES (?, ?)",
                    (name, json.dumps(settings))
                )
                self.db_conn.commit()
                self.load_profiles()
                self.populate_profiles_list()
                self.update_status(f"Profile '{name}' saved")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save profile: {str(e)}")

    def export_statistics(self):
        """Export statistics to file"""
        file_path = filedialog.asksaveasfilename(
            title="Export Statistics",
            defaultextension=".csv",
            filetypes=[
                ("CSV files", "*.csv"),
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )

        if file_path:
            try:
                # Get statistics data
                self.db_cursor.execute("SELECT * FROM statistics ORDER BY date DESC")
                data = self.db_cursor.fetchall()

                with open(file_path, 'w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(['Date', 'Characters Typed', 'Words Typed', 'Sessions Completed'])
                    writer.writerows(data)

                self.update_status(f"Statistics exported to {os.path.basename(file_path)}")
                messagebox.showinfo("Success", "Statistics exported successfully!")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export: {str(e)}")

    def save_hotkeys(self):
        """Save hotkey configuration"""
        try:
            self.setup_hotkeys()
            self.update_status("Hotkeys saved and activated")
            messagebox.showinfo("Success", "Hotkeys saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save hotkeys: {str(e)}")

    def open_user_guide(self):
        """Open user guide in browser"""
        webbrowser.open("https://github.com/your-repo/auto-typing/wiki")

    def report_issue(self):
        """Open issue reporting"""
        webbrowser.open("https://github.com/your-repo/auto-typing/issues")

    def update_status(self, message):
        """Update status bar"""
        self.status_bar.config(text=message)
        print(f"Status: {message}")


def main():
    root = tk.Tk()
    app = ExtendedAutoTypingSoftware(root)

    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    # Set icon and title
    root.title("Professional Auto Typing Suite")

    # Make window resizable
    root.minsize(900, 600)

    root.mainloop()


if __name__ == "__main__":
    main()