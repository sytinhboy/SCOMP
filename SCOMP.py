import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import zipfile
import pyzipper
import argon2
import secrets
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
import shutil
import platform
import subprocess
from PIL import Image, ImageTk
import logging
import tempfile  # Add this import for the temporary file functionality
import time  # Add this import for time-based throttling
from concurrent.futures import ThreadPoolExecutor
import concurrent.futures
import datetime  # Add this import for datetime



# Import modules specific to each operating system
if platform.system() == "Windows":
    import winreg
    import ctypes

# Add this near the top of your file, after imports
_root = None  # Global reference to the main Tk instance

# Add this function to get the root window
def get_root_window():
    global _root
    logging.debug("get_root_window called, current _root: %s", _root)
    if _root is None or not _root.winfo_exists():
        logging.debug("Creating new Tk instance")
        if TKDND_AVAILABLE:
            _root = TkinterDnD.Tk()
        else:
            _root = tk.Tk()
        _root.withdraw()  # Hide by default
        logging.debug("New Tk instance created: %s", _root)
    return _root

# Replace the current set_button_style function with this one
def set_button_style():
    try:
        style = ttk.Style()
        
        # Force default theme first to reset any previous settings
        if 'default' in style.theme_names():
            style.theme_use('default')
        
        # Create a more modern button style with subtle gradient and rounded corners
        style.configure("Modern.TButton", 
                        background="#4a86e8", 
                        foreground="white",
                        padding=(10, 5),
                        relief="flat",
                        font=("Helvetica", 10, "bold"))
        
        style.map("Modern.TButton",
                  background=[("pressed", "#000000"),  # Màu đen khi click
                            ("active", "#3a76d8")],    # Màu xanh đậm khi hover
                  foreground=[("pressed", "#ffffff")], # Màu chữ trắng khi click
                  relief=[("pressed", "sunken")])      # Hiệu ứng nút nhấn
        
        # Thêm style mới cho button lớn
        style.configure("Large.Modern.TButton", 
                        background="#4a86e8", 
                        foreground="white",
                        padding=(10, 6),
                        relief="flat",
                        font=("Helvetica", 12, "bold"))
        
        style.map("Large.Modern.TButton",
                  background=[("pressed", "#000000"),  # Màu đen khi click
                            ("active", "#3a76d8")],    # Màu xanh đậm khi hover
                  foreground=[("pressed", "#ffffff")], # Màu chữ trắng khi click 
                  relief=[("pressed", "sunken")])      # Hiệu ứng nút nhấn
        
        # Create secondary button style (gray)
        style.configure("Secondary.TButton", 
                        background="#6c757d", 
                        foreground="white",
                        padding=(10, 5),
                        relief="flat",
                        font=("Helvetica", 9))
        
        style.map("Secondary.TButton",
                  background=[("active", "#5c656d"), ("pressed", "#4c555d")],
                  relief=[("pressed", "flat")])
        
        # Style for progress bar
        style.configure("TProgressbar", 
                        thickness=8,
                        background="#4a86e8")
    except Exception as e:
        print(f"Failed to set button style: {e}")

# Call this function at the very beginning of your program
set_button_style()

# Add TkinterDnD2 import with error handling for systems where it's not installed
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TKDND_AVAILABLE = True
except ImportError:
    TKDND_AVAILABLE = False

class FileCompressorApp:
    def __init__(self, root):
        self.root = root
        
        # Set background color for entire app
        self.root.configure(bg="#f8f9fa")
        
        # Initialize drag and drop if available
        self.setup_drag_and_drop()
        
        self.setup_ui()
        self.setup_context_menu()
        
        # Initialize variables
        self.files_to_compress = []
        self.current_task = None
        self.password = None
        
        # Add this new line to store file paths passed as arguments
        self.initial_files = []
        
        # Process command line arguments if any
        self.process_command_line_args()
        
    def setup_drag_and_drop(self):
        """Setup drag and drop functionality if tkinterdnd2 is available"""
        if TKDND_AVAILABLE:
            # Make sure root is a TkinterDnD.Tk or convert it
            if not isinstance(self.root, TkinterDnD.Tk):
                # Just note that drag-drop is available but root is not compatible
                self.dnd_enabled = False
            else:
                self.dnd_enabled = True
        else:
            self.dnd_enabled = False
        
    def setup_ui(self):
        # Set color scheme
        bg_color = "#f8f9fa"  # Light gray background
        accent_color = "#4a86e8"  # Blue accent
        text_color = "#212529"  # Dark text for contrast
        border_color = "#dee2e6"  # Light border color
        
        # Main frame with padding
        main_frame = tk.Frame(self.root, bg=bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # App header with logo and title
        header_frame = tk.Frame(main_frame, bg=bg_color)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Add a simple logo/icon if available
        try:
            logo_img = tk.PhotoImage(file="icon.png").subsample(8, 8)  # Scale down the image
            logo_label = tk.Label(header_frame, image=logo_img, bg=bg_color)
            logo_label.image = logo_img  # Keep a reference
            logo_label.pack(side=tk.LEFT, padx=(0, 10))
        except:
            pass  # No logo available
        
        # App title with larger, modern font
        title_label = tk.Label(header_frame, text="Secure File Compressor", 
                              font=("Helvetica", 20, "bold"), 
                              bg=bg_color, fg=accent_color)
        title_label.pack(side=tk.LEFT)
        
        # Add version indicator with tooltip
        version_label = tk.Label(header_frame, text="v1.0", 
                               font=("Helvetica", 9), 
                               bg=bg_color, fg="#6c757d",
                               cursor="hand2")
        version_label.pack(side=tk.RIGHT, padx=(0, 5), pady=(5, 0))
        
        # Bind click event to open email client
        version_label.bind("<Button-1>", self.open_email_client)
        
        # Create tooltip for version label
        self.create_tooltip(version_label, 
                           "AES-256-GCM encryption for maximum security.\n\n"
                           "Argon2 password hashing for strong protection.\n\n"
                           "The strongest ZIP file encryption in the world.\n\n"
                           f"© {datetime.datetime.now().year} SComp\n\n"
                           "Click to contact support")
        # Content area (2-column layout)
        content_frame = tk.Frame(main_frame, bg=bg_color)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left panel for file list with border and rounded corners effect
        left_panel = tk.Frame(content_frame, bg=bg_color, bd=1, relief=tk.SOLID, highlightbackground=border_color)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # File list section title
        file_list_header = tk.Frame(left_panel, bg=accent_color, height=40)
        file_list_header.pack(fill=tk.X)
        
        tk.Label(file_list_header, text="Files to Compress", 
                fg="white", bg=accent_color, 
                font=("Helvetica", 14, "bold")).pack(side=tk.LEFT, padx=15, pady=8)
        
        # Buttons for file management
        buttons_frame = tk.Frame(left_panel, bg=bg_color)
        buttons_frame.pack(fill=tk.X, padx=15, pady=10)
        
        add_files_btn = ttk.Button(buttons_frame, text="Add Files", 
                                 command=self.add_files, 
                                 style="Modern.TButton")
        add_files_btn.configure(style="Large.Modern.TButton")  # Sử dụng style mới
        add_files_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        add_folder_btn = ttk.Button(buttons_frame, text="Add Folder", 
                                  command=self.add_folder, 
                                  style="Modern.TButton")
        add_folder_btn.configure(style="Large.Modern.TButton")  # Sử dụng style mới
        add_folder_btn.pack(side=tk.LEFT, padx=5)
        
        # File list with custom styling
        self.file_list_frame = tk.Frame(left_panel, bg="white", bd=0)
        self.file_list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        self.file_listbox = tk.Listbox(self.file_list_frame, 
                                      selectmode=tk.EXTENDED, 
                                      bg="white", 
                                      fg=text_color,
                                      font=("Helvetica", 11),
                                      bd=0, 
                                      highlightthickness=1,
                                      highlightbackground=border_color)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Enable drag and drop for the file list if available
        if self.dnd_enabled:
            self.file_listbox.drop_target_register(DND_FILES)
            self.file_listbox.dnd_bind('<<Drop>>', self.handle_drop)
            # Also make the main window a drop target
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.handle_drop)
        
            # Add drag-drop indicator
            dnd_text = "Drag and drop files here"
            self.dnd_label = tk.Label(left_panel, text=dnd_text, 
                                    bg=bg_color, fg="#6c757d", 
                                    font=("Helvetica", 9, "italic"))
            self.dnd_label.pack(fill=tk.X, pady=(0, 10), padx=15)
        
        # Scrollbar for file list
        scrollbar = ttk.Scrollbar(self.file_list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        
        # Right panel for options and actions
        right_panel = tk.Frame(content_frame, bg=bg_color, bd=1, relief=tk.SOLID, highlightbackground=border_color)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # Options section title
        options_header = tk.Frame(right_panel, bg=accent_color, height=40)
        options_header.pack(fill=tk.X)
        
        tk.Label(options_header, text="Compression Options", 
                fg="white", bg=accent_color, 
                font=("Helvetica", 14, "bold")).pack(side=tk.LEFT, padx=15, pady=8)
        
        # Options container with padding
        options_container = tk.Frame(right_panel, bg=bg_color)
        options_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Compression format and level
        format_frame = tk.Frame(options_container, bg=bg_color)
        format_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(format_frame, text="Format:", bg=bg_color, fg=text_color,
                font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
        
        self.format_var = tk.StringVar(value="zip")
        tk.Label(format_frame, text="ZIP", bg=bg_color, fg=text_color,
                font=("Helvetica", 10)).pack(side=tk.LEFT, padx=(5, 0))
        
        # Define level_var as a hidden variable with default value
        self.level_var = tk.IntVar(value=5)
        
        # Output path
        path_frame = tk.Frame(options_container, bg=bg_color)
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(path_frame, text="Output Path:", bg=bg_color, fg=text_color,
                font=("Helvetica", 11, "bold")).pack(side=tk.LEFT)
        
        # Create a frame for the rounded entry with variable border color
        self.output_entry_frame = tk.Frame(path_frame, bg=border_color, bd=0, relief=tk.SOLID)
        self.output_entry_frame.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        
        # Create a frame inside with bg color for inner padding
        output_inner_frame = tk.Frame(self.output_entry_frame, bg="white", bd=0)
        output_inner_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        self.output_path_var = tk.StringVar()
        output_entry = tk.Entry(output_inner_frame, textvariable=self.output_path_var, 
                              font=("Helvetica", 13), bd=0, relief=tk.FLAT,
                              highlightthickness=0)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=3)
        
        # Add focus events for highlighting
        output_entry.bind("<FocusIn>", lambda e: self.on_entry_focus_in(self.output_entry_frame))
        output_entry.bind("<FocusOut>", lambda e: self.on_entry_focus_out(self.output_entry_frame))
        
        browse_btn = ttk.Button(path_frame, text="Browse", 
                              command=self.browse_output, 
                              style="Modern.TButton")
        browse_btn.pack(side=tk.LEFT)
        
        # Encryption options - redesigned with switch-like appearance
        encrypt_frame = tk.Frame(options_container, bg=bg_color)
        encrypt_frame.pack(fill=tk.X, pady=(5, 10))
        
        self.encrypt_var = tk.BooleanVar(value=True)
        tk.Checkbutton(encrypt_frame, text="Enable Encryption", 
                      variable=self.encrypt_var, 
                      bg=bg_color, fg=text_color,
                      font=("Helvetica", 11, "bold"),
                      command=self.toggle_password).pack(side=tk.LEFT)
        
        # Password input with improved styling and rounded corners
        password_frame = tk.Frame(options_container, bg=bg_color)
        password_frame.pack(fill=tk.X, pady=(0, 5))
        
        tk.Label(password_frame, text="Password:", bg=bg_color, fg=text_color,
                font=("Helvetica", 12)).pack(side=tk.LEFT)
        
        # Create a frame for the rounded entry with variable border color
        self.password_entry_frame = tk.Frame(password_frame, bg=border_color, bd=0, relief=tk.SOLID)
        self.password_entry_frame.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        
        # Create a frame inside with bg color for inner padding
        password_inner_frame = tk.Frame(self.password_entry_frame, bg="white", bd=0)
        password_inner_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        self.password_var = tk.StringVar()
        self.password_entry = tk.Entry(password_inner_frame, textvariable=self.password_var, 
                                     show="•", font=("Helvetica", 12), bd=0, relief=tk.FLAT,
                                     highlightthickness=0)
        self.password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=3)
        
        # Add focus events for highlighting
        self.password_entry.bind("<FocusIn>", lambda e: self.on_entry_focus_in(self.password_entry_frame))
        self.password_entry.bind("<FocusOut>", lambda e: self.on_entry_focus_out(self.password_entry_frame))
        
        # Add show password checkbox
        show_password_frame = tk.Frame(options_container, bg=bg_color)
        show_password_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.show_password_var = tk.BooleanVar(value=False)
        tk.Checkbutton(show_password_frame, text="Show password", 
                      variable=self.show_password_var, 
                      bg=bg_color, fg=text_color,
                      font=("Helvetica", 11),
                      command=self.toggle_password_visibility).pack(side=tk.LEFT, padx=(60, 0))
        
        # Progress section - MOVED HERE TO REPLACE COMPATIBILITY MODE
        progress_frame = tk.Frame(options_container, bg=bg_color)
        progress_frame.pack(fill=tk.X, pady=(5, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, 
                                          variable=self.progress_var, 
                                          maximum=100,
                                          style="TProgressbar")
        self.progress_bar.pack(fill=tk.X)
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = tk.Label(progress_frame, textvariable=self.status_var, 
                              bg=bg_color, fg="#6c757d",
                              font=("Helvetica", 9))
        status_label.pack(side=tk.RIGHT, pady=(3, 0))
        
        # Action buttons at the bottom - update style to Modern.TButton
        action_frame = tk.Frame(options_container, bg=bg_color)
        action_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.compress_btn = ttk.Button(action_frame, text="Compress Files", 
                                    command=self.compress_files, 
                                    style="Modern.TButton")
        self.compress_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.extract_btn = ttk.Button(action_frame, text="Extract Archive", 
                               command=self.extract_files, 
                               style="Modern.TButton")  # Changed from Secondary.TButton
        self.extract_btn.pack(side=tk.LEFT)
        
        # Update context menu integration buttons if they exist
        if platform.system() in ["Windows", "Darwin"]:
            # Footer for system integration
            footer_frame = tk.Frame(main_frame, bg=bg_color)
            footer_frame.pack(fill=tk.X, pady=(15, 0))
            
            tk.Label(footer_frame, text="System Integration", 
                    font=("Helvetica", 11, "bold"),
                    bg=bg_color, fg=text_color).pack(anchor="w")
            
            integration_frame = tk.Frame(footer_frame, bg=bg_color)
            integration_frame.pack(fill=tk.X, pady=(5, 0))
            
            # Create both buttons but initially only pack one of them
            self.add_context_btn = ttk.Button(integration_frame, text="Add to Right-click Menu", 
                               command=self.add_to_context_menu, 
                               style="Modern.TButton")
            
            self.remove_context_btn = ttk.Button(integration_frame, text="Remove from Right-click Menu", 
                                  command=self.remove_from_context_menu, 
                                  style="Modern.TButton")
            
            # Check if context menu is already installed and show appropriate button
            is_context_menu_installed = check_if_context_menu_installed()
            self.toggle_context_menu_buttons(is_added=is_context_menu_installed)
        
    def update_level_label(self, *args):
        self.level_label.config(text=str(self.level_var.get()))
        
    def toggle_password(self):
        if self.encrypt_var.get():
            self.password_entry.config(state=tk.NORMAL)
        else:
            self.password_entry.config(state=tk.DISABLED)
            
    def add_files(self):
        files = filedialog.askopenfilenames(title="Select files to compress")
        if files:
            for file in files:
                if file not in self.files_to_compress:
                    self.files_to_compress.append(file)
                    self.file_listbox.insert(tk.END, os.path.basename(file))
                    
    def add_folder(self):
        folder = filedialog.askdirectory(title="Select folder to compress")
        if folder:
            self.files_to_compress.append(folder)
            self.file_listbox.insert(tk.END, f"[DIR] {os.path.basename(folder)}")
            
    def remove_files(self):
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            return
            
        # Remove in reverse order to avoid index shifting
        for index in sorted(selected_indices, reverse=True):
            del self.files_to_compress[index]
            self.file_listbox.delete(index)
            
    def browse_output(self):
        # Only ZIP format is supported now
        default_ext = ".zip"
        filetypes = [("ZIP files", "*.zip")]
        
        output_path = filedialog.asksaveasfilename(
            title="Save compressed file as",
            defaultextension=default_ext,
            filetypes=filetypes
        )
        
        if output_path:
            self.output_path_var.set(output_path)
            
    def compress_files(self):
        if not self.files_to_compress:
            messagebox.showinfo("Info", "Please add files or folders to compress")
            return
            
        if not self.output_path_var.get():
            messagebox.showinfo("Info", "Please specify an output path")
            return
            
        # Make sure we have a password if encryption is enabled
        if self.encrypt_var.get():
            # Get the password directly from the entry widget to ensure we have the latest value
            password = self.password_var.get()
            
            if not password or password.isspace():
                messagebox.showinfo("Info", "Please enter a password for encryption")
                return
                
            # Store password for use in the compression thread
            self.password = password
            
        else:
            # Explicitly set to None if encryption is disabled
            self.password = None
            print("DEBUG: Encryption disabled, password set to None")
        
        # Start compression in a separate thread
        self.executor = ThreadPoolExecutor(max_workers=4)
        self.current_task = self.executor.submit(self._compress_task)
        
        # Disable buttons during compression
        self.compress_btn.config(state=tk.DISABLED)
        self.extract_btn.config(state=tk.DISABLED)
        
    def _compress_task(self):
        success = False
        error_message = None
        
        try:
            self.status_var.set("Preparing...")
            self.progress_var.set(0)
            
            output_path = self.output_path_var.get()
            compression_level = self.level_var.get()
            use_encryption = self.encrypt_var.get()
            
            # Use the class attribute for password and confirm it's valid
            password = self.password
            
            # Double-check encryption settings before proceeding
            if use_encryption:
                if not password:
                    raise Exception("Encryption is enabled but no password was provided.")
                    
                if not isinstance(password, str):
                    raise Exception(f"Password must be a string, but got {type(password).__name__}.")
                    
                if not password.strip():
                    raise Exception("Encryption requires a non-empty password.")
            
            # Validate output directory exists
            output_dir = os.path.dirname(output_path)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Determine total size for progress tracking
            total_size = self._calculate_total_size(self.files_to_compress)
            
            # Create ZIP archive with optimized method
            if use_encryption:
                self.status_var.set("Creating encrypted archive...")
                processed_size = self._create_encrypted_zip_with_embedded_salt(output_path, password, total_size)
            else:
                self.status_var.set("Creating archive...")
                processed_size = self._create_standard_zip_optimized(output_path, compression_level, total_size)
            
            # Compression completed
            self.progress_var.set(100)
            self.status_var.set("Compression completed")
            messagebox.showinfo("Success", 
                            f"Archive has been created at:\n{output_path}")
            success = True
            
        except Exception as e:
            error_message = str(e)
            print(f"Compression error: {error_message}")
        
        finally:
            # Schedule GUI updates on the main thread
            self.root.after(0, lambda: self._finish_compression(success, error_message))

    def _finish_compression(self, success, error_message=None):
        """Handle UI updates after compression (must be called on main thread)"""
        if error_message:
            self.status_var.set("Error")
            messagebox.showerror("Error", f"An error occurred during compression: {error_message}")
        
        # Re-enable buttons
        self.compress_btn.config(state=tk.NORMAL)
        self.extract_btn.config(state=tk.NORMAL)
        
        # Clear the file list only if compression was successful
        if success:
            self.clear_files()

    

    def _calculate_total_size(self, file_paths):
        """Calculate the total size of all files to be compressed"""
        total_size = 0
        for path in file_paths:
            if os.path.isfile(path):
                total_size += os.path.getsize(path)
            else:
                # Sử dụng os.scandir để liệt kê file nhanh hơn os.walk
                for root, dirs, files in os.walk(path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            total_size += os.path.getsize(file_path)
                        except (OSError, FileNotFoundError):
                            # Bỏ qua file không đọc được
                            continue
        return total_size

    def _create_standard_zip_optimized(self, output_path, compression_level, total_size):
        """Create a standard non-encrypted ZIP archive with parallel processing"""
        processed_size = 0
        update_interval = 0.5  # Giảm tần suất cập nhật UI
        last_update_time = time.time()
        
        # Thu thập tất cả các file cần nén
        files_to_add = self._collect_files_to_compress()
        
        # Chia thành các nhóm để xử lý song song
        chunk_size = max(1, len(files_to_add) // (os.cpu_count() or 4))
        file_chunks = [files_to_add[i:i + chunk_size] for i in range(0, len(files_to_add), chunk_size)]
        
        # Nếu có "store only" mode (nén không nén, chỉ đóng gói)
        # compression_mode = zipfile.ZIP_STORED if compression_level == 0 else zipfile.ZIP_DEFLATED
        compression_mode = zipfile.ZIP_DEFLATED
        compresslevel = min(9, max(1, compression_level))  # Giới hạn trong khoảng 1-9
        
        with zipfile.ZipFile(output_path, 'w', compression_mode, compresslevel=compresslevel) as zipf:
            # Sử dụng lock để tránh race condition khi nhiều thread ghi vào cùng một file zip
            zip_lock = threading.Lock()
            progress_lock = threading.Lock()
            
            def process_file_chunk(file_chunk):
                nonlocal processed_size, last_update_time
                local_processed = 0
                
                for file_path, arcname, file_size in file_chunk:
                    try:
                        with zip_lock:
                            zipf.write(file_path, arcname)
                        
                        # Cập nhật tiến trình
                        with progress_lock:
                            local_processed += file_size
                            processed_size += file_size
                            
                            current_time = time.time()
                            if current_time - last_update_time > update_interval:
                                progress = min(100, int(processed_size * 100 / total_size))
                                self.progress_var.set(progress)
                                self.status_var.set(f"Compressing: {os.path.basename(file_path)}")
                                self.root.update_idletasks()
                                last_update_time = current_time
                                
                    except Exception as e:
                        print(f"Error processing {file_path}: {str(e)}")
                
                return local_processed
            
            # Xử lý song song các nhóm file
            with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
                futures = [executor.submit(process_file_chunk, chunk) for chunk in file_chunks]
                for future in futures:
                    future.result()  # Đợi tất cả các thread hoàn thành
            
            # Đảm bảo thanh tiến trình được cập nhật lần cuối
            self.progress_var.set(100)
            self.root.update_idletasks()
        
        return processed_size

    def _collect_files_to_compress(self):
        """Thu thập tất cả các file cần nén và trả về danh sách (file_path, arcname, file_size)"""
        files_to_add = []
        
        for path in self.files_to_compress:
            if os.path.isfile(path):
                try:
                    file_size = os.path.getsize(path)
                    files_to_add.append((path, os.path.basename(path), file_size))
                except (OSError, FileNotFoundError):
                    print(f"Warning: Cannot access file {path}")
                    continue
            else:
                base_dir = os.path.basename(path)
                for root, _, files in os.walk(path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            arcname = os.path.join(base_dir, os.path.relpath(file_path, os.path.dirname(path)))
                            file_size = os.path.getsize(file_path)
                            files_to_add.append((file_path, arcname, file_size))
                        except (OSError, FileNotFoundError):
                            print(f"Warning: Cannot access file {file_path}")
                            continue
        
        # Sắp xếp theo kích thước để cân bằng khối lượng công việc giữa các thread
        files_to_add.sort(key=lambda x: x[2], reverse=True)
        
        return files_to_add

    def _create_encrypted_zip_with_embedded_salt(self, output_path, password, total_size):
        """Create an encrypted ZIP archive with embedded salt using optimized processing"""
        processed_size = 0
        buffer_size = 16 * 1024 * 1024  # Tăng buffer lên 16MB
        update_interval = 0.5  # Giảm tần suất cập nhật UI
        last_update_time = time.time()
        
        # Validate password
        if password is None:
            raise Exception("WZ_AES encryption requires a password. Password cannot be None.")
        
        if not isinstance(password, str):
            raise Exception(f"WZ_AES encryption requires a string password. Got {type(password).__name__} instead.")
        
        if not password.strip():
            raise Exception("WZ_AES encryption requires a non-empty password.")
        
        # Ensure password is properly encoded to bytes
        try:
            password_bytes = password.encode('utf-8')
        except Exception as e:
            raise Exception(f"Failed to encode password: {str(e)}. Please use valid characters in your password.")
        
        # Generate salt for key derivation
        salt = secrets.token_bytes(16)
        
        # Thu thập tất cả file cần nén
        files_to_add = self._collect_files_to_compress()
        
        # Create a temporary file for the zip content
        with tempfile.NamedTemporaryFile(delete=False) as temp_zip_file:
            temp_zip_path = temp_zip_file.name
        
        try:
            # Create the encrypted ZIP
            with pyzipper.AESZipFile(
                temp_zip_path, 
                'w', 
                compression=pyzipper.ZIP_DEFLATED,
                encryption=pyzipper.WZ_AES
            ) as zipf:
                # Set the password for encryption - BEFORE adding any files
                zipf.setpassword(password_bytes)
                
                # Add a special file with salt information inside the ZIP
                zipf.writestr("__SALT__", salt)
                
                # Calculate key using Argon2 (for future verification)
                kdf = argon2.low_level.hash_secret_raw(
                    password_bytes,
                    salt,
                    time_cost=3,
                    memory_cost=65536,
                    parallelism=4,
                    hash_len=32,
                    type=argon2.low_level.Type.ID
                )
                
                # Process all files with reduced UI updates for better performance
                batch_size = 10  # Xử lý theo batch để giảm số lần cập nhật UI
                current_batch_size = 0
                batch_size_total = 0
                
                for idx, (file_path, arcname, file_size) in enumerate(files_to_add):
                    # Update status less frequently
                    current_time = time.time()
                    if (current_time - last_update_time > update_interval or 
                        idx == 0 or idx == len(files_to_add) - 1):
                        self.status_var.set(f"Encrypting: {idx+1}/{len(files_to_add)} files")
                        self.root.update_idletasks()
                        last_update_time = current_time
                    
                    # Add file to archive
                    zipf.write(file_path, arcname)
                    
                    # Update progress
                    processed_size += file_size
                    current_batch_size += 1
                    batch_size_total += file_size
                    
                    # Update progress bar after processing a batch or on the last file
                    if current_batch_size >= batch_size or idx == len(files_to_add) - 1:
                        progress = min(100, int(processed_size * 100 / total_size))
                        self.progress_var.set(progress)
                        self.root.update_idletasks()
                        current_batch_size = 0
                        batch_size_total = 0
            
            # Copy the temporary file to the final output path more efficiently
            self.status_var.set("Finalizing archive...")
            self.root.update_idletasks()
            
            # Sử dụng buffer lớn hơn khi copy file
            with open(temp_zip_path, 'rb') as src, open(output_path, 'wb') as dst:
                while True:
                    buffer = src.read(buffer_size)
                    if not buffer:
                        break
                    dst.write(buffer)
            
        except Exception as e:
            # Create a more detailed error message
            error_msg = str(e)
            if "password" in error_msg.lower():
                raise Exception(f"Encryption error: {error_msg}. Please check that you've provided a valid password.")
            else:
                raise Exception(f"Error creating encrypted archive: {error_msg}")
        finally:
            # Clean up the temporary file
            try:
                if os.path.exists(temp_zip_path):
                    os.unlink(temp_zip_path)
            except Exception as e:
                print(f"Warning: Failed to delete temporary file: {str(e)}")

        return processed_size

    def extract_files(self):
        """Handle extraction of files from the file list or file selection"""
        # First check if there are any files in the list
        if self.files_to_compress:
            # Filter only ZIP files from the list
            zip_files = [f for f in self.files_to_compress if f.lower().endswith('.zip')]
            if not zip_files:
                messagebox.showinfo("No ZIP Files", "Please add ZIP files to extract.")
                return
                
            # Ask for extraction directory
            extract_dir = filedialog.askdirectory(title="Select extraction folder")
            if not extract_dir:
                return
                
            # Process each ZIP file
            for zip_file in zip_files:
                # Check if password is needed
                need_password = False
                password = None
                
                try:
                    with zipfile.ZipFile(zip_file, 'r') as zipf:
                        zipf.testzip()
                except RuntimeError:
                    # Most likely password-protected
                    need_password = True
                except Exception:
                    # Other errors might occur during testing, but we'll handle them in the extraction thread
                    pass
                
                if need_password:
                    password = self.ask_password()
                    if password is None:  # User canceled
                        continue
                
                # Start extraction in a separate thread with optimized extraction task
                self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count() or 4)  
                self.current_task = self.executor.submit(self._extract_task_optimized, zip_file, extract_dir, password)
                
                # Disable buttons during extraction
                self.compress_btn.config(state=tk.DISABLED)
                self.extract_btn.config(state=tk.DISABLED)
                
                # Only process one ZIP file at a time to prevent interface issues
                break
        else:
            # If no files in list, ask for archive to extract (original behavior)
            archive_path = filedialog.askopenfilename(
                title="Select archive to extract",
                filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
            )
            
            if not archive_path:
                return
                
            # Ask for extraction directory
            extract_dir = filedialog.askdirectory(title="Select extraction folder")
            if not extract_dir:
                return
                
            # Check if password is needed
            need_password = False
            password = None
            
            try:
                with zipfile.ZipFile(archive_path, 'r') as zipf:
                    zipf.testzip()
            except RuntimeError:
                # Most likely password-protected
                need_password = True
            except Exception:
                # Other errors might occur during testing, but we'll handle them in the extraction thread
                pass
            
            if need_password:
                password = self.ask_password()
                if password is None:  # User canceled
                    return
            
            # Start extraction in a separate thread with optimized extraction task
            self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count() or 4)
            self.current_task = self.executor.submit(self._extract_task_optimized, archive_path, extract_dir, password)
            
            # Disable buttons during extraction
            self.compress_btn.config(state=tk.DISABLED)
            self.extract_btn.config(state=tk.DISABLED)

    def _extract_task_optimized(self, archive_path, extract_dir, password=None):
        """Optimized extraction task with parallel processing and improved performance"""
        success = False
        password_error = False
        update_interval = 0.5  # Reduce UI updates for better performance
        last_update_time = time.time()
        
        try:
            # Update UI from main thread
            self.root.after(0, lambda: self.status_var.set("Preparing..."))
            self.root.after(0, lambda: self.progress_var.set(0))
            
            if archive_path.lower().endswith('.zip'):
                # Get the file list and prepare for extraction
                file_info = self._prepare_extraction(archive_path, extract_dir, password)
                
                if file_info is None:
                    # Password error or other issue occurred in preparation
                    return
                    
                zipf, file_list, total_files, use_pyzipper, salt = file_info
                
                # Perform parallel extraction
                self._perform_parallel_extraction(
                    zipf, file_list, total_files, extract_dir, 
                    use_pyzipper, password, salt, update_interval
                )
                
                success = True
            
            # Schedule final success message and updates on the main thread
            if success:
                def on_success():
                    self.progress_var.set(100)
                    self.status_var.set("Extraction completed")
                    messagebox.showinfo("Success", f"Archive has been extracted to:\n{extract_dir}")
                    # Only clear the file list after a successful extraction, not on password errors
                    if not password_error and self.files_to_compress:
                        self.clear_files()
                        
                self.root.after(0, on_success)
            
        except Exception as e:
            error_message = str(e)
            print(f"Extraction error: {error_message}")
            
            # Check if this is a password-related error
            if "password" in error_message.lower():
                password_error = True
            
            # Schedule error message on the main thread
            def on_error():
                self.status_var.set("Error")
                
                # Check if we need to display a password-related error
                if "password" in error_message.lower():
                    error_title = "Password Error"
                    error_message_display = error_message
                else:
                    error_title = "Extraction Error"
                    if "compression method" not in error_message.lower():
                        error_message_display = f"An error occurred during extraction:\n{error_message}"
                    else:
                        error_message_display = error_message
                
                # Ensure the root window is visible
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
                
                # Display error with extra handling for macOS Finder integration
                if platform.system() == "Darwin" and len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
                    # When launched from Finder/Services menu on macOS, use AppleScript for more reliable dialog
                    try:
                        apple_script = f"""
                        display dialog "{error_message_display}" with title "{error_title}" buttons {{"OK"}} default button 1 with icon stop
                        """
                        subprocess.run(["osascript", "-e", apple_script], check=False)
                    except:
                        # Fall back to standard messagebox if AppleScript fails
                        messagebox.showerror(error_title, error_message_display)
                else:
                    # Standard error display for normal application use
                    messagebox.showerror(error_title, error_message_display)
            
            self.root.after(0, on_error)
            
        finally:
            # Always re-enable buttons on the main thread, regardless of success/failure
            def cleanup():
                self.compress_btn.config(state=tk.NORMAL)
                self.extract_btn.config(state=tk.NORMAL)
                
            self.root.after(0, cleanup)

    def _prepare_extraction(self, archive_path, extract_dir, password=None):
        """Prepare for extraction by opening the archive and getting file list"""
        self.root.after(0, lambda: self.status_var.set("Analyzing archive..."))
        
        # Convert password to bytes if provided
        password_bytes = password.encode() if password else None
        
        # Try with pyzipper first for better encryption support
        try:
            zipf = pyzipper.AESZipFile(archive_path, 'r')
            if password_bytes:
                zipf.setpassword(password_bytes)
            
            # Check if this is our custom format with embedded salt
            salt = None
            file_list = zipf.namelist()
            
            if '__SALT__' in file_list:
                try:
                    # Extract the salt file
                    salt = zipf.read('__SALT__')
                    
                    # If we found salt, we need to re-derive the key using Argon2
                    if salt and password:
                        kdf = argon2.low_level.hash_secret_raw(
                            password_bytes,
                            salt,
                            time_cost=3,
                            memory_cost=65536,
                            parallelism=4,
                            hash_len=32,
                            type=argon2.low_level.Type.ID
                        )
                except Exception as salt_error:
                    print(f"Salt extraction error: {str(salt_error)}")
                    # If we encounter an error reading the salt, continue with normal extraction
                    self.root.after(0, lambda: self.status_var.set(
                        "Note: Could not read embedded salt. Continuing with standard extraction."))
            
            # Remove the salt file from extraction list if it exists
            if '__SALT__' in file_list:
                file_list.remove('__SALT__')
                
            # Test if we can read files to confirm password is correct
            if file_list:
                try:
                    # Test read first file
                    test_file = file_list[0]
                    zipf.read(test_file)
                except RuntimeError as e:
                    # Password error
                    if "password" in str(e).lower() or "bad password" in str(e).lower():
                        raise Exception("The password provided is incorrect. Please try again with the correct password.")
            
            return zipf, file_list, len(file_list), True, salt
        
        except Exception as e:
            # If it's a password error, re-raise it
            if "password" in str(e).lower():
                raise e
                
            print(f"PyZipper failed: {str(e)}, trying standard zipfile")
            
            # Try with standard zipfile as fallback
            try:
                zipf = zipfile.ZipFile(archive_path, 'r')
                if password_bytes:
                    zipf.setpassword(password_bytes)
                
                file_list = zipf.namelist()
                
                # Remove the salt file from extraction list if it exists (unlikely with standard zip)
                if '__SALT__' in file_list:
                    file_list.remove('__SALT__')
                
                # Test if we can read files
                if file_list:
                    try:
                        # Test read first file
                        test_file = file_list[0]
                        zipf.read(test_file)
                    except RuntimeError as e:
                        # Password error
                        if "password" in str(e).lower():
                            raise Exception("The password provided is incorrect. Please try again with the correct password.")
                
                return zipf, file_list, len(file_list), False, None
                
            except Exception as zip_err:
                # Handle specific zipfile errors
                if "password required" in str(zip_err).lower():
                    raise Exception("This archive is password-protected. Please provide a password.")
                elif "bad password" in str(zip_err).lower():
                    raise Exception("The password provided is incorrect. Please try again with the correct password.")
                elif "compression method" in str(zip_err).lower():
                    # Special handling for unsupported compression methods
                    raise Exception(f"Unable to extract this archive: {str(zip_err)}. Try using an external application.")
                else:
                    raise Exception(f"Error opening archive: {str(zip_err)}")

    def _perform_parallel_extraction(self, zipf, file_list, total_files, extract_dir, use_pyzipper, password, salt, update_interval):
        """Extract files in parallel for better performance with auto-renaming of duplicates"""
        processed_files = 0
        error_files = []
        renamed_files = []
        last_update_time = time.time()
        
        # Create extraction directory if it doesn't exist
        if not os.path.exists(extract_dir):
            os.makedirs(extract_dir)
        
        # Prepare a lock for thread-safe operations
        progress_lock = threading.Lock()
        
        # Determine chunk size for parallel processing
        max_workers = os.cpu_count() or 4
        chunk_size = max(1, min(100, len(file_list) // max_workers))
        
        # Divide files into chunks for parallel processing
        file_chunks = [file_list[i:i+chunk_size] for i in range(0, len(file_list), chunk_size)]
        
        def extract_file_chunk(file_chunk):
            """Process a chunk of files in a worker thread"""
            nonlocal processed_files, last_update_time
            chunk_errors = []
            chunk_renamed = []
            local_processed = 0
            
            for idx, file in enumerate(file_chunk):
                try:
                    # Handle directories
                    if file.endswith('/') or file.endswith('\\'):
                        dir_path = os.path.join(extract_dir, file)
                        if os.path.exists(dir_path):
                            # If directory exists, create a new name with counter
                            counter = 1
                            while True:
                                new_dir_path = os.path.join(extract_dir, f"{file.rstrip('/\\')}_{counter}")
                                if not os.path.exists(new_dir_path):
                                    os.makedirs(new_dir_path)
                                    chunk_renamed.append((file, f"{file.rstrip('/\\')}_{counter}"))
                                    break
                                counter += 1
                        else:
                            os.makedirs(dir_path)
                        
                        with progress_lock:
                            local_processed += 1
                            processed_files += 1
                        continue
                    
                    # Extract the file with auto-renaming if necessary
                    file_data = zipf.read(file)
                    
                    # Get file path components
                    target_path = os.path.join(extract_dir, file)
                    target_dir = os.path.dirname(target_path)
                    file_name = os.path.basename(target_path)
                    file_base, file_ext = os.path.splitext(file_name)
                    
                    # Ensure the directory exists for this file
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir)
                    
                    # Check if file already exists and auto-rename if needed
                    if os.path.exists(target_path):
                        counter = 1
                        while True:
                            new_file_name = f"{file_base}_{counter}{file_ext}"
                            new_target_path = os.path.join(target_dir, new_file_name)
                            if not os.path.exists(new_target_path):
                                # Found an available name
                                chunk_renamed.append((file, f"{os.path.dirname(file)}/{new_file_name}" if os.path.dirname(file) else new_file_name))
                                target_path = new_target_path
                                break
                            counter += 1
                    
                    # Write the file with large buffer for better performance
                    with open(target_path, 'wb') as f:
                        f.write(file_data)
                    
                    with progress_lock:
                        local_processed += 1
                        processed_files += 1
                        
                        # Update UI only occasionally for better performance
                        current_time = time.time()
                        if current_time - last_update_time > update_interval:
                            progress = min(100, int(processed_files * 99 / total_files))  # Use 99% as max to show activity
                            self.root.after(0, lambda p=progress: self.progress_var.set(p))
                            self.root.after(0, lambda f=file: self.status_var.set(f"Extracting: {os.path.basename(f)}"))
                            self.root.update_idletasks()
                            last_update_time = current_time
                            
                except Exception as e:
                    # Record errors but continue with other files
                    chunk_errors.append((file, str(e)))
                    print(f"Error extracting {file}: {str(e)}")
            
            return chunk_errors, chunk_renamed
        
        # Use thread pool to extract files in parallel
        self.root.after(0, lambda: self.status_var.set("Extracting files..."))
        
        # Extract files in parallel batches
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(extract_file_chunk, chunk) for chunk in file_chunks]
            
            # Process results as they complete
            for future in concurrent.futures.as_completed(futures):
                chunk_errors, chunk_renamed = future.result()
                if chunk_errors:
                    error_files.extend(chunk_errors)
                if chunk_renamed:
                    renamed_files.extend(chunk_renamed)
        
        # Final progress update
        self.root.after(0, lambda: self.progress_var.set(99))
        self.root.after(0, lambda: self.status_var.set("Finalizing extraction..."))
        
        # Close the zipfile
        zipf.close()
        
        # Report renamed files
        if renamed_files:
            renamed_count = len(renamed_files)
            self.root.after(0, lambda: self.status_var.set(f"Extraction completed. {renamed_count} files renamed."))
            
            if renamed_count <= 5:
                renamed_details = "\n".join([f"Original: {orig} → New: {new}" for orig, new in renamed_files[:5]])
                self.root.after(0, lambda: messagebox.showinfo(
                    "Files Renamed", 
                    f"{renamed_count} files were renamed to avoid overwriting:\n{renamed_details}"
                ))
            else:
                # Just show the count if there are too many renamed files
                self.root.after(0, lambda: messagebox.showinfo(
                    "Files Renamed", 
                    f"{renamed_count} files were renamed to avoid overwriting existing files."
                ))
        
        # Report any errors that occurred during extraction
        if error_files:
            error_count = len(error_files)
            if error_count <= 5:
                error_details = "\n".join([f"{file}: {error}" for file, error in error_files])
                self.root.after(0, lambda: messagebox.showwarning(
                    "Extraction Warning", 
                    f"{error_count} files could not be extracted:\n{error_details}"
                ))
            else:
                # Just show the count if there are too many errors
                self.root.after(0, lambda: messagebox.showwarning(
                    "Extraction Warning", 
                    f"{error_count} files could not be extracted."
                ))
    
    def ask_password(self):
        """Display dialog to request password"""
        password_dialog = tk.Toplevel(self.root)
        password_dialog.title("Enter Password")
        password_dialog.geometry("300x160")  # Made slightly taller to accommodate checkbox
        password_dialog.resizable(False, False)
        password_dialog.transient(self.root)
        password_dialog.grab_set()
        
        # Center on parent
        password_dialog.geometry("+%d+%d" % (
            self.root.winfo_rootx() + (self.root.winfo_width() / 2) - 150,
            self.root.winfo_rooty() + (self.root.winfo_height() / 2) - 75
        ))
        
        # Password entry
        frame = tk.Frame(password_dialog, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(frame, text="Enter password to extract:").pack(anchor=tk.W)
        
        password_var = tk.StringVar()
        password_entry = tk.Entry(frame, textvariable=password_var, show="•", width=30)
        password_entry.pack(fill=tk.X, pady=(5, 5))
        password_entry.focus_set()
        
        # Add show password checkbox
        show_password_var = tk.BooleanVar(value=False)
        
        def toggle_extract_password_visibility():
            if show_password_var.get():
                password_entry.config(show="")
            else:
                password_entry.config(show="•")
        
        tk.Checkbutton(frame, text="Show password", variable=show_password_var, 
                      command=toggle_extract_password_visibility).pack(anchor=tk.W, pady=(0, 5))
        
        # Result variable
        result = [None]  # Using list as a mutable container
        
        # Buttons
        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill=tk.X)
        
        def on_ok():
            result[0] = password_var.get()
            password_dialog.destroy()
            
        def on_cancel():
            password_dialog.destroy()
            
        # Adding gray color to dialog buttons
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10, style="Gray.TButton").pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=10, style="Gray.TButton").pack(side=tk.RIGHT)
        
        # Handle Enter key
        password_dialog.bind("<Return>", lambda e: on_ok())
        password_dialog.bind("<Escape>", lambda e: on_cancel())
        
        # Wait for dialog to close
        self.root.wait_window(password_dialog)
        return result[0]
    
    def setup_context_menu(self):
        """Setup right-click menu for file list"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Remove", command=self.remove_files)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Clear All", command=self.clear_files)
        
        # Bind right-click event
        self.file_listbox.bind("<Button-3>", self.show_context_menu)
        
    def show_context_menu(self, event):
        """Display right-click menu"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
            
    def clear_files(self):
        """Clear all files from the list"""
        self.files_to_compress.clear()
        self.file_listbox.delete(0, tk.END)

    def add_to_context_menu(self):
        """Add to system right-click menu"""
        if platform.system() == "Windows":
            if not is_admin():
                messagebox.showinfo("Admin Rights Required", "You need to run the application with administrator rights to add to Windows right-click menu.")
                return
            
            if add_to_windows_context_menu():
                messagebox.showinfo("Success", "Added to Windows right-click menu")
                # Toggle buttons after successful addition
                self.toggle_context_menu_buttons(is_added=True)
            else:
                messagebox.showerror("Error", "Could not add to Windows right-click menu")
        
        elif platform.system() == "Darwin":  # macOS
            # Try multiple methods for maximum compatibility
            success1 = add_to_mac_context_menu()
            success2 = create_mac_quick_actions()  # Fixed: use create_mac_quick_actions instead of create_mac_quick_action
            
            if success1 or success2:
                messagebox.showinfo("Success", 
                                   "Added to macOS menu. You can access in multiple ways:\n\n"
                                   "1. Right-click on files → Quick Actions\n"
                                   "2. Right-click on files → Services menu\n"
                                   "3. Select files, then click Scripts (✓) in menu bar")
                # Toggle buttons after successful addition
                self.toggle_context_menu_buttons(is_added=True)
            else:
                messagebox.showerror("Error", "Could not add to macOS right-click menu")

    def remove_from_context_menu(self):
        """Remove from system right-click menu"""
        if platform.system() == "Windows":
            if not is_admin():
                messagebox.showinfo("Admin Rights Required", "You need to run the application with administrator rights to remove from Windows right-click menu.")
                return
            
            if remove_from_windows_context_menu():
                messagebox.showinfo("Success", "Removed from Windows right-click menu")
                # Toggle buttons after successful removal
                self.toggle_context_menu_buttons(is_added=False)
            else:
                messagebox.showerror("Error", "Could not remove from Windows right-click menu")
        
        elif platform.system() == "Darwin":  # macOS
            if remove_from_mac_context_menu():
                messagebox.showinfo("Success", "Removed from macOS right-click menu")
                # Toggle buttons after successful removal
                self.toggle_context_menu_buttons(is_added=False)
            else:
                messagebox.showerror("Error", "Could not remove from macOS right-click menu")

    def toggle_password_visibility(self):
        """Toggle between showing and hiding password characters"""
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="•")

    def handle_drop(self, event):
        """Handle files dropped onto the application"""
        # Parse the data to get file paths
        data = event.data
        
        # Different systems format multiple file paths differently
        if data.startswith('{') and data.endswith('}'):
            # Tcl list format: {path1} {path2}
            file_paths = []
            # Simple parsing for braced Tcl list
            current = ""
            in_brace = False
            for char in data:
                if char == '{':
                    if not in_brace:
                        in_brace = True
                    else:
                        current += char
                elif char == '}':
                    if in_brace:
                        in_brace = False
                        if current:
                            file_paths.append(current)
                            current = ""
                    else:
                        current += char
                else:
                    current += char
        else:
            # Space-separated or newline-separated
            file_paths = data.replace('\n', ' ').replace('\r', ' ').split()
        
        # Process each dropped file
        for path in file_paths:
            # Remove quotes if present
            path = path.strip('"\'')
            
            # Check if the file exists
            if not os.path.exists(path):
                continue
                
            # Check if it's an archive file that should be extracted
            if path.lower().endswith('.zip'):
                # Ask user whether to add to compression list or extract
                result = messagebox.askyesno("Archive Detected", 
                                           f"'{os.path.basename(path)}' is an archive file.\n\nDo you want to extract it?\n\n(Click 'No' to add it to compression list)",
                                           icon='question')
                if result:
                    # User wants to extract
                    self.extract_dropped_archive(path)
                    continue
            
            # Add to compression list
            if path not in self.files_to_compress:
                self.files_to_compress.append(path)
                if os.path.isdir(path):
                    self.file_listbox.insert(tk.END, f"[DIR] {os.path.basename(path)}")
                else:
                    self.file_listbox.insert(tk.END, os.path.basename(path))
    
    def extract_dropped_archive(self, archive_path):
        """Handle extraction of a dropped archive file"""
        # Ensure main window is visible and focused first
        self.root.lift()
        self.root.focus_force()
        
        # On macOS, temporarily set window to stay on top
        if platform.system() == "Darwin":
            self.root.attributes("-topmost", True)
        
        # Ask for extraction directory with parent window specified
        extract_dir = filedialog.askdirectory(
            title="Select extraction folder",
            parent=self.root  # Explicitly set parent window
        )
        
        # Disable topmost attribute after dialog closes (macOS)
        if platform.system() == "Darwin":
            self.root.attributes("-topmost", False)
        
        if not extract_dir:
            return
        
        # Use after_idle to ensure proper window ordering
        self.root.after_idle(lambda: self._check_password_and_extract(archive_path, extract_dir))

    def _check_password_and_extract(self, archive_path, extract_dir):
        """Check if password is needed and start extraction in a separate thread"""
        need_password = False
        password = None
        
        # Check archive type to see if it requires a password
        if archive_path.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(archive_path, 'r') as zipf:
                    zipf.testzip()
            except RuntimeError:
                # Most likely password-protected
                need_password = True
            except Exception:
                # Other errors might occur during testing, but we'll handle them in the extraction thread
                pass
        
        if need_password:
            # Schedule password prompt on the main thread
            def prompt_password():
                password = self.ask_password()
                if password is None:  # User canceled
                    return
                # After getting password, start extraction in a separate thread
                self._start_extraction_thread(archive_path, extract_dir, password)
            
            self.root.after(0, prompt_password)
        else:
            # No password needed, start extraction directly
            self._start_extraction_thread(archive_path, extract_dir, None)

    
    def _start_extraction_thread(self, archive_path, extract_dir, password):
        """Start the extraction process using ThreadPoolExecutor"""
        if not hasattr(self, 'executor'):
            self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=4)
        
        # Sử dụng _extract_task_optimized thay vì _extract_task
        self.executor.submit(self._extract_task_optimized, archive_path, extract_dir, password)
        
        # Disable buttons during extraction
        self.compress_btn.config(state=tk.DISABLED)
        self.extract_btn.config(state=tk.DISABLED)
    
    def on_entry_focus_in(self, frame):
        """Highlight the frame with blue border when entry gets focus"""
        frame.config(bg="#4a86e8")  # Change to accent color when focused
    
    def on_entry_focus_out(self, frame):
        """Reset the frame border color when entry loses focus"""
        frame.config(bg="#dee2e6")  # Reset to original border color

    def create_tooltip(self, widget, text):
        """Create a tooltip for a given widget with the specified text"""
        tooltip_window = None
        
        def enter(event):
            nonlocal tooltip_window
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25
            
            # Create a toplevel window
            tooltip_window = tk.Toplevel(widget)
            tooltip_window.wm_overrideredirect(True)  # Remove window decorations
            tooltip_window.wm_geometry(f"+{x}+{y}")
            
            # Create frame with border
            frame = tk.Frame(tooltip_window, bg="#ffffcc", bd=1, relief="solid")
            frame.pack(fill="both", expand=True)
            
            # Create and pack the label
            label = tk.Label(frame, text=text, justify=tk.LEFT,
                            bg="#ffffcc", fg="#333333", padx=10, pady=8,
                            wraplength=300, font=("Helvetica", 9))
            label.pack()
        
        def leave(event):
            nonlocal tooltip_window
            if tooltip_window:
                tooltip_window.destroy()
                tooltip_window = None
        
        # Bind events to the widget
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def toggle_context_menu_buttons(self, is_added):
        """Toggle visibility of Add and Remove context menu buttons"""
        if hasattr(self, 'add_context_btn') and hasattr(self, 'remove_context_btn'):
            if is_added:
                # Hide Add button, show Remove button
                self.add_context_btn.pack_forget()
                self.remove_context_btn.pack(side=tk.LEFT)
            else:
                # Hide Remove button, show Add button
                self.remove_context_btn.pack_forget()
                self.add_context_btn.pack(side=tk.LEFT, padx=(0, 5))

    def open_email_client(self, event=None):
        """Open the default email client with a pre-defined recipient"""
        recipient = "info@sytinh.com"
        subject = "Secure File Compressor Support"
        
        try:
            # Construct the mailto URL
            mailto_url = f"mailto:{recipient}?subject={subject}"
            
            # Open the URL with the default handler
            if platform.system() == "Windows":
                os.startfile(mailto_url)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", mailto_url], check=False)
            else:  # Linux and other Unix-like systems
                subprocess.run(["xdg-open", mailto_url], check=False)
                
        except Exception as e:
            print(f"Error opening email client: {str(e)}")
            messagebox.showerror("Error", f"Could not open email client: {str(e)}")

    def process_command_line_args(self):
        """Process files passed as command line arguments."""
        import sys
        
        # Skip the first argument (script name)
        if len(sys.argv) > 1:
            file_paths = sys.argv[1:]
            self.initial_files = file_paths
            
            # Add files to the listbox after UI is set up
            # We need to use after() because the listbox might not exist yet
            self.root.after(100, lambda: self.add_files_from_args())
    
    def add_files_from_args(self):
        """Add files that were passed as command line arguments."""
        if self.initial_files:
            for file_path in self.initial_files:
                if os.path.exists(file_path):
                    # Fix: Use files_to_compress list instead of file_paths set
                    if file_path not in self.files_to_compress:
                        self.files_to_compress.append(file_path)
                        # Add to listbox directly rather than calling update_file_listbox
                        if os.path.isdir(file_path):
                            self.file_listbox.insert(tk.END, f"[DIR] {os.path.basename(file_path)}")
                        else:
                            self.file_listbox.insert(tk.END, os.path.basename(file_path))
            
            # Clear the initial files list to avoid adding them again
            self.initial_files = []
    
    def update_file_listbox(self):
        """Update the listbox with current file paths."""
        self.file_listbox.delete(0, tk.END)
        # Fix: Use files_to_compress list instead of file_paths
        sorted_paths = sorted(self.files_to_compress)
        for path in sorted_paths:
            if os.path.isdir(path):
                self.file_listbox.insert(tk.END, f"[DIR] {os.path.basename(path)}")
            else:
                self.file_listbox.insert(tk.END, os.path.basename(path))

# Add utility functions for Windows (if running on Windows)
def add_to_windows_context_menu():
    try:
        # Create registry key for right-click menu
        app_path = os.path.abspath(sys.argv[0])
        
        # If it's a .py file, use python to run it
        if app_path.endswith('.py'):
            single_file_command = f'"{sys.executable}" "{app_path}" "%1"'
            extract_command = f'"{sys.executable}" "{app_path}" --extract "%1"'
            multi_file_command = f'"{sys.executable}" "{app_path}"'
        else:
            single_file_command = f'"{app_path}" "%1"'
            extract_command = f'"{app_path}" --extract "%1"'
            multi_file_command = f'"{app_path}"'
        
        # Create right-click menu for single files
        key_path = r'*\\shell\\SecureCompress'
        key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
        winreg.SetValueEx(key, '', 0, winreg.REG_SZ, 'SComp')
        
        # Set icon (if available)
        if os.path.exists("app_icon.ico"):
            winreg.SetValueEx(key, 'Icon', 0, winreg.REG_SZ, os.path.abspath("app_icon.ico"))
        
        # Create command
        command_key = winreg.CreateKey(key, 'command')
        winreg.SetValueEx(command_key, '', 0, winreg.REG_SZ, single_file_command)
        
        # Close registry keys
        winreg.CloseKey(command_key)
        winreg.CloseKey(key)
        
        # Create right-click menu for multiple files
        multi_key_path = r'Directory\Background\shell\SecureCompressMulti'
        multi_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, multi_key_path)
        winreg.SetValueEx(multi_key, '', 0, winreg.REG_SZ, 'SComp Selected Files')
        
        # Set icon for multi-file selection
        if os.path.exists("app_icon.ico"):
            winreg.SetValueEx(multi_key, 'Icon', 0, winreg.REG_SZ, os.path.abspath("app_icon.ico"))
        
        # Create command for multi-file selection
        multi_command_key = winreg.CreateKey(multi_key, 'command')
        # For multiple files, we'll use the shell's selected files
        for_each_command = multi_file_command + ' %V'
        winreg.SetValueEx(multi_command_key, '', 0, winreg.REG_SZ, for_each_command)
        
        # Close registry keys for multi-file selection
        winreg.CloseKey(multi_command_key)
        winreg.CloseKey(multi_key)
        
        # Add entry for File Selection context menu
        selection_key_path = r'SystemFileAssociations\Directory\shellex\ContextMenuHandlers\SecureCompressMulti'
        selection_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, selection_key_path)
        winreg.SetValueEx(selection_key, '', 0, winreg.REG_SZ, '{96236A7F-9DBC-11D2-8F0C-00C04FA31009}')
        winreg.CloseKey(selection_key)
        
        # Add "Extract with File Compressor" option for ZIP files
        extract_key_path = r'.zip\\shell\\SecureExtract'
        extract_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, extract_key_path)
        winreg.SetValueEx(extract_key, '', 0, winreg.REG_SZ, 'Extract with SComp')
        
        # Set icon for extract option if available
        if os.path.exists("app_icon.ico"):
            winreg.SetValueEx(extract_key, 'Icon', 0, winreg.REG_SZ, os.path.abspath("app_icon.ico"))
        
        # Create command for extract option
        extract_command_key = winreg.CreateKey(extract_key, 'command')
        winreg.SetValueEx(extract_command_key, '', 0, winreg.REG_SZ, extract_command)
        
        # Close registry keys for extract option
        winreg.CloseKey(extract_command_key)
        winreg.CloseKey(extract_key)
        
        return True
    except Exception as e:
        print(f"Error adding to right-click menu: {str(e)}")
        return False

def remove_from_windows_context_menu():
    try:
        # Delete original single-file registry key
        key_path = r'*\\shell\\SecureCompress'
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path + '\\command')
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
        
        # Delete multi-file registry keys
        multi_key_path = r'Directory\Background\shell\SecureCompressMulti'
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, multi_key_path + '\\command')
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, multi_key_path)
        
        # Delete context menu handler
        selection_key_path = r'SystemFileAssociations\Directory\shellex\ContextMenuHandlers\SecureCompressMulti'
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, selection_key_path)
        
        # Delete extract option registry keys
        extract_key_path = r'.zip\\shell\\SecureExtract'
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, extract_key_path + '\\command')
        winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, extract_key_path)
        
        return True
    except Exception as e:
        print(f"Error removing from right-click menu: {str(e)}")
        return False

def is_admin():
    """Check if the application is running with administrator privileges"""
    if platform.system() == "Windows":
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False
    elif platform.system() == "Darwin":  # macOS
        # On macOS, we check sudo rights via id -u
        try:
            return os.geteuid() == 0
        except:
            return False
    else:  # Linux and other operating systems
        try:
            return os.geteuid() == 0
        except:
            return False

# Add utility functions for macOS
def add_to_mac_context_menu():
    try:
        # Get the correct path to the application based on whether it's bundled or not
        if getattr(sys, 'frozen', False):
            # If running as a bundled app (.app)
            app_path = os.path.dirname(os.path.dirname(os.path.dirname(sys.executable)))
            app_name = os.path.basename(app_path)
            
            # Fix: Use proper quoting and format for macOS open command
            escaped_path = app_path.replace('"', '\\"')
            compress_command = f'open "{escaped_path}" --args --compress "$itemPath"'
            extract_command = f'open "{escaped_path}" --args --extract "$itemPath"'
            multi_compress_command = f'open "{escaped_path}" --args "$@"'
        else:
            # If running as a script
            app_path = os.path.abspath(sys.argv[0])
            python_path = sys.executable
            
            # Fix: Use proper quoting for shell commands
            escaped_python = python_path.replace('"', '\\"')
            escaped_app = app_path.replace('"', '\\"')
            compress_command = f'"{escaped_python}" "{escaped_app}" --compress "$itemPath"'
            extract_command = f'"{escaped_python}" "{escaped_app}" --extract "$itemPath"'
            multi_compress_command = f'"{escaped_python}" "{escaped_app}" "$@"'
        
        # Create scripts directory
        scripts_dir = os.path.expanduser("~/Library/Scripts/Finder")
        os.makedirs(scripts_dir, exist_ok=True)
        
        # Use very simple string concatenation for AppleScript - no f-strings
        compress_script = """-- Secure Compress Script
on run
    tell application "Finder"
        try
            set theSelection to the selection
            if theSelection is {} then
                display dialog "Please select at least one file first" buttons {"OK"} default button 1 with icon caution
            else
                set itemPaths to {}
                repeat with anItem in theSelection
                    set itemPath to POSIX path of (anItem as alias)
                    set end of itemPaths to quoted form of itemPath
                end repeat
                
                -- Handle multiple files with new compress-multi command
                set pathsString to (do shell script "echo " & (items of itemPaths as string))
                do shell script """ + '"' + multi_compress_command + '"' + """
            end if
        on error errMsg
            display dialog "An error occurred: " & errMsg buttons {"OK"} default button 1 with icon stop
        end try
    end tell
end run
"""
        
        extract_script = """-- Secure Extract Script
on run
    tell application "Finder"
        try
            set theSelection to the selection
            if theSelection is {} then
                display dialog "Please select a ZIP file first" buttons {"OK"} default button 1 with icon caution
            else
                repeat with anItem in theSelection
                    try
                        set itemPath to POSIX path of (anItem as alias)
                        -- Only process ZIP files
                        if itemPath ends with ".zip" then
                            do shell script """ + '"' + extract_command + '"' + """
                        else
                            display dialog "Only ZIP files can be extracted: " & itemPath buttons {"OK"} default button 1 with icon caution
                        end if
                    on error errMsg
                        display dialog "Error processing: " & itemPath & "\\n\\n" & errMsg buttons {"OK"} default button 1 with icon stop
                    end try
                end repeat
            end if
        on error errMsg
            display dialog "An error occurred: " & errMsg buttons {"OK"} default button 1 with icon stop
        end try
    end tell
end run
"""
        
        # Save scripts as compiled AppleScript files
        compress_path = os.path.join(scripts_dir, "Secure Compress.scpt")
        extract_path = os.path.join(scripts_dir, "Secure Extract.scpt")
        
        # Write and compile compression script
        temp_script = os.path.expanduser("~/temp_script.applescript")
        with open(temp_script, 'w', encoding='utf-8') as f:
            f.write(compress_script)
        subprocess.run(["osacompile", "-o", compress_path, temp_script], check=True)
        
        # Write and compile extraction script
        with open(temp_script, 'w', encoding='utf-8') as f:
            f.write(extract_script)
        subprocess.run(["osacompile", "-o", extract_path, temp_script], check=True)
        
        # Clean up
        os.remove(temp_script)
        
        # Create a user-friendly guide
        guide_script = """
        display dialog "Secure Compress/Extract has been added to your Scripts menu!

To use:
1. In Finder, select one or more files
2. Click on the Scripts icon (✓) in the menu bar (top right corner)
   - If you don't see this icon, go to Finder Preferences > Advanced and enable 'Show Script menu in menu bar'
3. Select 'Compress With SComp.app' to compress files or 'Extract With SComp.app' for ZIP files

You can also use the Services menu by right-clicking on files in Finder." buttons {"OK"} default button 1
        """
        
        try:
            subprocess.run(["osascript", "-e", guide_script], check=False)
        except:
            pass
            
        return True
    except Exception as e:
        print(f"Error adding to macOS right-click menu: {str(e)}")
        return False

def remove_from_mac_context_menu():
    try:
        # Get possible app name variations for more thorough cleanup
        app_names = ["File Compressor", "SecureCompressor", "Secure Compressor", "SComp"]
        try:
            # Try to get the actual application name if running as a bundled app
            bundle_path = os.path.abspath(os.path.dirname(sys.executable))
            if bundle_path.endswith('.app/Contents/MacOS'):
                app_name = os.path.basename(bundle_path.split('.app')[0])
                app_names.append(app_name)
        except Exception:
            pass

        # Remove fixed-name workflows
        services_to_remove = [
            "~/Library/Services/Secure Compress.workflow",
            "~/Library/Services/Secure Extract.workflow",
            "~/Library/Services/SecureCompress.workflow",
            "~/Library/Services/SecureExtract.workflow", 
            "~/Library/Services/SecureCompressor.workflow", 
            "~/Library/Services/CompressSecure.workflow"  # Old name
        ]
        
        # Add dynamically named workflows based on app name
        for app_name in app_names:
            services_to_remove.append(f"~/Library/Services/Compress with {app_name}.workflow")
            services_to_remove.append(f"~/Library/Services/Extract with {app_name}.workflow")
        
        # Remove all workflows
        for service_path in services_to_remove:
            path = os.path.expanduser(service_path)
            if os.path.exists(path):
                print(f"Removing service: {path}")
                shutil.rmtree(path)
        
        # Remove scripts from the Scripts menu - use correct filenames
        script_paths = [
            "~/Library/Scripts/Finder/Secure Compress.scpt",
            "~/Library/Scripts/Finder/Secure Extract.scpt",
            "~/Library/Scripts/Finder/SecureCompress.scpt",
            "~/Library/Scripts/Finder/SecureCompressor.scpt",
            "~/Library/Scripts/Finder/SecureExtract.scpt"
        ]
        
        for script_path in script_paths:
            path = os.path.expanduser(script_path)
            if os.path.exists(path):
                print(f"Removing script: {path}")
                os.remove(path)
        
        # Remove helper scripts
        helper_script_dir = os.path.expanduser("~/Library/Application Support/SecureCompressor")
        if os.path.exists(helper_script_dir):
            shutil.rmtree(helper_script_dir)
        
        # Force refresh services menu with more robust error handling
        try:
            subprocess.run(["killall", "-KILL", "pbs"], stderr=subprocess.DEVNULL, 
                          stdout=subprocess.DEVNULL, check=False)
        except:
            pass
            
        try:
            subprocess.run(["killall", "-KILL", "secd"], stderr=subprocess.DEVNULL, 
                          stdout=subprocess.DEVNULL, check=False)
        except:
            pass
            
        try:
            subprocess.run(["/System/Library/CoreServices/pbs", "-flush"], 
                         stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL, check=False)
        except:
            pass
            
        try:
            subprocess.run(["/System/Library/CoreServices/pbs", "-update"], 
                         stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL, check=False)
        except:
            pass
        
        # Restart Finder to apply changes
        try:
            subprocess.run(["killall", "Finder"], stderr=subprocess.DEVNULL, 
                         stdout=subprocess.DEVNULL, check=False)
        except:
            pass
        
        return True
    except Exception as e:
        print(f"Error removing from macOS right-click menu: {str(e)}")
        return False

def create_mac_quick_actions():
    """Create Quick Action workflows for macOS right-click menu integration"""
    try:
        # Try to get the app name from bundle if running as an app
        app_name = "SComp"  # Default name
        
        # Determine the shell commands based on how the app is running
        if getattr(sys, 'frozen', False):
            # Running as a bundled app
            app_path = os.path.abspath(sys.executable)
            app_dir = os.path.dirname(os.path.dirname(os.path.dirname(app_path)))
            app_name = os.path.basename(app_dir)
            compress_shell_command = f'open "{app_dir}" --args "$@"'
            extract_shell_command = f'open "{app_dir}" --args --extract "$@"'
        else:
            # If running as a script
            python_path = sys.executable
            script_path = os.path.abspath(sys.argv[0])
            compress_shell_command = f'"{python_path}" "{script_path}" "$@"'
            extract_shell_command = f'"{python_path}" "{script_path}" --extract "$@"'
        
        # Create workflows directory
        workflows_dir = os.path.expanduser("~/Library/Services")
        os.makedirs(workflows_dir, exist_ok=True)
        
        # Create workflow paths
        compress_workflow_path = os.path.join(workflows_dir, f"Compress with {app_name}.workflow")
        extract_workflow_path = os.path.join(workflows_dir, f"Extract with {app_name}.workflow")
        
        # Ensure clean setup
        for path in [compress_workflow_path, extract_workflow_path]:
            if os.path.exists(path):
                shutil.rmtree(path)
            os.makedirs(os.path.join(path, "Contents"), exist_ok=True)
        
        # Write Info.plist files
        compress_info = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>NSServices</key>
    <array>
        <dict>
            <key>NSMenuItem</key>
            <dict>
                <key>default</key>
                <string>Compress with {app_name}</string>
            </dict>
            <key>NSMessage</key>
            <string>runWorkflowAsService</string>
            <key>NSRequiredContext</key>
            <dict>
                <key>NSApplicationIdentifier</key>
                <string>com.apple.finder</string>
            </dict>
            <key>NSSendFileTypes</key>
            <array>
                <string>public.item</string>
            </array>
        </dict>
    </array>
</dict>
</plist>'''

        extract_info = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>NSServices</key>
    <array>
        <dict>
            <key>NSMenuItem</key>
            <dict>
                <key>default</key>
                <string>Extract with {app_name}</string>
            </dict>
            <key>NSMessage</key>
            <string>runWorkflowAsService</string>
            <key>NSRequiredContext</key>
            <dict>
                <key>NSApplicationIdentifier</key>
                <string>com.apple.finder</string>
            </dict>
            <key>NSSendFileTypes</key>
            <array>
                <string>public.zip-archive</string>
            </array>
        </dict>
    </array>
</dict>
</plist>'''

        # Write workflow files
        workflow_template = '''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>AMApplicationBuild</key>
    <string>509</string>
    <key>AMApplicationVersion</key>
    <string>2.10</string>
    <key>AMDocumentVersion</key>
    <string>2</string>
    <key>actions</key>
    <array>
        <dict>
            <key>action</key>
            <dict>
                <key>AMAccepts</key>
                <dict>
                    <key>Container</key>
                    <string>List</string>
                    <key>Optional</key>
                    <false/>
                    <key>Types</key>
                    <array>
                        <string>com.apple.cocoa.path</string>
                    </array>
                </dict>
                <key>AMActionVersion</key>
                <string>2.3.1</string>
                <key>AMApplication</key>
                <array>
                    <string>Automator</string>
                </array>
                <key>AMParameterProperties</key>
                <dict/>
                <key>AMProvides</key>
                <dict>
                    <key>Container</key>
                    <string>List</string>
                    <key>Types</key>
                    <array>
                        <string>com.apple.cocoa.path</string>
                    </array>
                </dict>
                <key>ActionBundlePath</key>
                <string>/System/Library/Automator/Run Shell Script.action</string>
                <key>ActionName</key>
                <string>Run Shell Script</string>
                <key>ActionParameters</key>
                <dict>
                    <key>COMMAND_STRING</key>
                    <string>{}</string>
                    <key>CheckedForUserDefaultShell</key>
                    <true/>
                    <key>inputMethod</key>
                    <integer>1</integer>
                    <key>shell</key>
                    <string>/bin/bash</string>
                    <key>source</key>
                    <string></string>
                </dict>
                <key>BundleIdentifier</key>
                <string>com.apple.RunShellScript</string>
                <key>CFBundleVersion</key>
                <string>2.3.1</string>
                <key>CanShowSelectedItemsWhenRun</key>
                <true/>
                <key>CanShowWhenRun</key>
                <true/>
                <key>Category</key>
                <array>
                    <string>AMCategoryUtilities</string>
                </array>
                <key>Class Name</key>
                <string>RunShellScriptAction</string>
                <key>InputUUID</key>
                <string>42DC2DF4-BD6A-49F2-830D-A3A6D9D13F67</string>
                <key>Keywords</key>
                <array>
                    <string>Shell</string>
                    <string>Script</string>
                    <string>Command</string>
                    <string>Run</string>
                    <string>Unix</string>
                </array>
                <key>OutputUUID</key>
                <string>C589E4C5-B9FC-48A2-AFB1-C85B564DF222</string>
                <key>UUID</key>
                <string>964F7FDD-8B1D-48AD-834E-F35BE50D5E70</string>
                <key>UnlocalizedApplications</key>
                <array>
                    <string>Automator</string>
                </array>
            </dict>
        </dict>
    </array>
    <key>workflowMetaData</key>
    <dict>
        <key>applicationBundleIDsByPath</key>
        <dict/>
        <key>applicationPaths</key>
        <array/>
        <key>inputTypeIdentifier</key>
        <string>com.apple.Automator.fileSystemObject</string>
        <key>outputTypeIdentifier</key>
        <string>com.apple.Automator.nothing</string>
        <key>presentationMode</key>
        <integer>15</integer>
        <key>processesInput</key>
        <true/>
        <key>serviceInputTypeIdentifier</key>
        <string>com.apple.Automator.fileSystemObject</string>
        <key>serviceOutputTypeIdentifier</key>
        <string>com.apple.Automator.nothing</string>
        <key>serviceProcessesInput</key>
        <true/>
        <key>systemImageName</key>
        <string>NSActionTemplate</string>
        <key>useAutomaticInputType</key>
        <false/>
        <key>workflowTypeIdentifier</key>
        <string>com.apple.Automator.servicesMenu</string>
    </dict>
</dict>
</plist>'''

        # Write all files
        with open(os.path.join(compress_workflow_path, "Contents/Info.plist"), "w") as f:
            f.write(compress_info)
        with open(os.path.join(extract_workflow_path, "Contents/Info.plist"), "w") as f:
            f.write(extract_info)
        
        with open(os.path.join(compress_workflow_path, "Contents/document.wflow"), "w") as f:
            f.write(workflow_template.format(compress_shell_command))
        with open(os.path.join(extract_workflow_path, "Contents/document.wflow"), "w") as f:
            f.write(workflow_template.format(extract_shell_command))

        # Refresh services
        subprocess.run(["/System/Library/CoreServices/pbs", "-flush"], 
                      stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL, check=False)
        
        return True
        
    except Exception as e:
        print(f"Error creating Quick Action for macOS: {str(e)}")
        return False

def check_if_context_menu_installed():
    """Check if the context menu entry is already installed"""
    try:
        if platform.system() == "Windows":
            # Check registry for our menu entry
            key_path = r'*\\shell\\SecureCompress'
            try:
                key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path)
                winreg.CloseKey(key)
                return True
            except FileNotFoundError:
                return False
            except Exception:
                return False
                
        elif platform.system() == "Darwin":  # macOS
            # Check for presence of our scripts or workflows
            paths_to_check = [
                "~/Library/Services/SecureCompressor.workflow",
                "~/Library/Services/Secure Compress.workflow",
                "~/Library/Scripts/Finder/Secure Compress.scpt",
                "~/Library/Scripts/Finder/SecureCompress.scpt"
            ]
            
            for path in paths_to_check:
                if os.path.exists(os.path.expanduser(path)):
                    return True
            
            return False
        
        return False
    except Exception:
        # If anything fails, default to showing the Add button
        return False

# Main function
def main():
    global _root
    
    # Setup logging for debugging
    log_dir = os.path.expanduser("~/Library/Logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "secure_compressor.log")
    
    logging.basicConfig(
        filename=log_file,
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    logging.debug("Application started")
    logging.debug(f"Arguments: {sys.argv}")
    logging.debug(f"Current directory: {os.getcwd()}")
    
    # Close any existing Tk instances first
    try:
        if tk._default_root and tk._default_root != _root and tk._default_root.winfo_exists():
            tk._default_root.quit()
            tk._default_root.destroy()
    except Exception as e:
        logging.error(f"Error closing existing Tk instances: {e}")
    
    # Get the root window (creates it if needed)
    root = get_root_window()
    
    # Apply the button style immediately after getting the root
    set_button_style()
    
    # Process command line arguments
    if len(sys.argv) > 1:
        logging.debug(f"Processing command line argument: {sys.argv[1]}")
        
        # Handle special command line arguments
        if sys.argv[1] == "--add-context-menu":
            # Window already hidden by get_root_window()
            logging.debug("Adding context menu")
            
            if platform.system() == "Windows" and not is_admin():
                messagebox.showinfo("Admin Rights Required", 
                                 "You need to run the application with administrator rights to add to Windows right-click menu.")
                root.quit()
                root.destroy()
                return
                
            if platform.system() == "Windows":
                success = add_to_windows_context_menu()
                if success:
                    messagebox.showinfo("Success", "Added to Windows right-click menu")
                else:
                    messagebox.showerror("Error", "Could not add to Windows right-click menu")
            elif platform.system() == "Darwin":  # macOS
                success1 = add_to_mac_context_menu()
                success2 = create_mac_quick_actions()
                
                if success1 or success2:
                    messagebox.showinfo("Success", 
                                     "Added to macOS menu. You can access in multiple ways:\n\n"
                                     "1. Right-click on files → Quick Actions\n"
                                     "2. Right-click on files → Services menu\n"
                                     "3. Select files, then click Scripts (✓) in menu bar")
                else:
                    messagebox.showerror("Error", "Could not add to macOS right-click menu")
            
            root.quit()
            root.destroy()
            return
            
        elif sys.argv[1] == "--remove-context-menu":
            # Window already hidden by get_root_window()
            logging.debug("Removing context menu")
            
            if platform.system() == "Windows" and not is_admin():
                messagebox.showinfo("Admin Rights Required", 
                                 "You need to run the application with administrator rights to remove from Windows right-click menu.")
                root.quit()
                root.destroy()
                return
                
            if platform.system() == "Windows":
                success = remove_from_windows_context_menu()
                if success:
                    messagebox.showinfo("Success", "Removed from Windows right-click menu")
                else:
                    messagebox.showerror("Error", "Could not remove from Windows right-click menu")
            elif platform.system() == "Darwin":  # macOS
                success = remove_from_mac_context_menu()
                if success:
                    messagebox.showinfo("Success", "Removed from macOS right-click menu")
                else:
                    messagebox.showerror("Error", "Could not remove from macOS right-click menu")
            
            root.quit()
            root.destroy()
            return
        
        # Add new argument handlers for compress and extract
        elif sys.argv[1] == "--compress" and len(sys.argv) > 2 and os.path.exists(sys.argv[2]):
            file_path = sys.argv[2]
            logging.debug(f"Opening file for compression: {file_path}")
            
            # Set the theme to default to ensure proper color rendering
            if platform.system() == "Windows":
                try:
                    style = ttk.Style()
                    style.theme_use('default')
                except Exception as e:
                    logging.error(f"Error setting theme: {e}")
            
            # Show the window for normal application use
            root.deiconify()
            root.title("Secure File Compressor")
            root.geometry("650x480")
            root.configure(bg="#f0f0f0")
            
            # Make sure the window is visible
            root.lift()
            root.focus_force()
            if platform.system() == "Darwin":  # macOS specific
                root.attributes("-topmost", True)
                root.after(1000, lambda: root.attributes("-topmost", False))
                    
            # Initialize the application with the existing root
            app = FileCompressorApp(root)
            
            # Add the file to the application
            if os.path.isfile(file_path):
                app.files_to_compress.append(file_path)
                app.file_listbox.insert(tk.END, os.path.basename(file_path))
                logging.debug(f"Added file to compression list: {file_path}")
            else:
                app.files_to_compress.append(file_path)
                app.file_listbox.insert(tk.END, f"[DIR] {os.path.basename(file_path)}")
                logging.debug(f"Added directory to compression list: {file_path}")
                
            # Set default output path
            default_output = os.path.splitext(file_path)[0] + ".zip"
            app.output_path_var.set(default_output)
            logging.debug(f"Set default output path: {default_output}")
            
            # Run the application
            logging.debug("Starting main loop")
            root.mainloop()
            return
            
        elif sys.argv[1] == "--extract" and len(sys.argv) > 2 and os.path.exists(sys.argv[2]):
            file_path = sys.argv[2]
            logging.debug(f"Opening file for extraction: {file_path}")
            
            # Set the theme to default to ensure proper color rendering
            if platform.system() == "Windows":
                try:
                    style = ttk.Style()
                    style.theme_use('default')
                except Exception as e:
                    logging.error(f"Error setting theme: {e}")
            
            # Show the window for normal application use
            root.deiconify()
            root.title("Secure File Compressor")
            root.geometry("650x480")
            root.configure(bg="#f0f0f0")
            
            # Make sure the window is visible
            root.lift()
            root.focus_force()
            if platform.system() == "Darwin":  # macOS specific
                root.attributes("-topmost", True)
                root.after(1000, lambda: root.attributes("-topmost", False))
                    
            # Initialize the application with the existing root
            app = FileCompressorApp(root)
            
            # Directly call extraction function on the file
            if file_path.lower().endswith('.zip'):
                # Start extraction process
                root.after(100, lambda: app.extract_dropped_archive(file_path))
                logging.debug(f"Initiating extraction for: {file_path}")
            else:
                messagebox.showinfo("Not a ZIP File", f"{os.path.basename(file_path)} is not a ZIP file and cannot be extracted.")
                logging.warning(f"Attempted to extract non-ZIP file: {file_path}")
                
            # Run the application
            logging.debug("Starting main loop for extraction")
            root.mainloop()
            return
        
        # Handle the existing case for a file path
        elif os.path.exists(sys.argv[1]):
            file_path = sys.argv[1]
            logging.debug(f"Opening file: {file_path}")
            
            # Set the theme to default to ensure proper color rendering
            if platform.system() == "Windows":
                try:
                    style = ttk.Style()
                    style.theme_use('default')
                except Exception as e:
                    logging.error(f"Error setting theme: {e}")
            
            # Show the window for normal application use
            root.deiconify()
            root.title("Secure File Compressor")
            root.geometry("650x480")
            root.minsize(650, 480)
            root.configure(bg="#f0f0f0")
            
            # Make sure the window is visible
            root.lift()
            root.focus_force()
            if platform.system() == "Darwin":  # macOS specific
                root.attributes("-topmost", True)
                root.after(1000, lambda: root.attributes("-topmost", False))
                    
            # Initialize the application with the existing root
            app = FileCompressorApp(root)
            
            # Add the file to the application
            if os.path.isfile(file_path):
                app.files_to_compress.append(file_path)
                app.file_listbox.insert(tk.END, os.path.basename(file_path))
                logging.debug(f"Added file to compression list: {file_path}")
            else:
                app.files_to_compress.append(file_path)
                app.file_listbox.insert(tk.END, f"[DIR] {os.path.basename(file_path)}")
                logging.debug(f"Added directory to compression list: {file_path}")
                
            # Set default output path
            default_output = os.path.splitext(file_path)[0] + ".zip"
            app.output_path_var.set(default_output)
            logging.debug(f"Set default output path: {default_output}")
            
            # Run the application
            logging.debug("Starting main loop")
            root.mainloop()
            return
        else:
            logging.warning(f"Argument does not exist as a file: {sys.argv[1]}")
    
    # No command line args or unrecognized args - just run the normal application
    logging.debug("Starting application with no arguments")
    
    # Set the theme to default to ensure proper color rendering
    if platform.system() == "Windows":
        try:
            style = ttk.Style()
            style.theme_use('default')
        except Exception as e:
            logging.error(f"Error setting theme: {e}")
    
    # Show the window for normal application use
    root.deiconify()
    root.title("Secure File Compressor")
    root.geometry("650x480")
    root.minsize(650, 480)
    root.configure(bg="#f0f0f0")
    
            
    # Initialize application with the existing root
    app = FileCompressorApp(root)
    
    # Run the application
    logging.debug("Starting main loop (normal mode)")
    root.mainloop()

if __name__ == "__main__":
    main()
