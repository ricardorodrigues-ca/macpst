"""
Main GUI window for Mac PST Converter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import logging
from pathlib import Path
from typing import List, Dict, Any
from ..core.pst_parser import PSTParser, EmailMessage
from ..core.converter import Converter


logger = logging.getLogger(__name__)


class MainWindow:
    """Main application window"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Mac PST Converter v1.0")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        self.pst_files = []
        self.current_messages = []
        self.converter = None
        self.output_directory = str(Path.home() / "Desktop" / "PST_Converted")
        
        self._setup_ui()
        self._setup_logging()
    
    def _setup_ui(self):
        """Set up the user interface"""
        style = ttk.Style()
        style.theme_use('aqua' if tk.TkVersion >= 8.5 else 'default')
        
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        self._create_file_selection_section(main_frame)
        self._create_output_section(main_frame)
        self._create_format_selection_section(main_frame)
        self._create_preview_section(main_frame)
        self._create_progress_section(main_frame)
        self._create_control_buttons(main_frame)
    
    def _create_file_selection_section(self, parent):
        """Create file selection section"""
        file_frame = ttk.LabelFrame(parent, text="PST File Selection", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Button(file_frame, text="Add PST Files", 
                  command=self._add_pst_files).grid(row=0, column=0, padx=(0, 10))
        
        self.file_listbox = tk.Listbox(file_frame, height=4)
        self.file_listbox.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self.file_listbox.bind('<Double-Button-1>', self._preview_pst_file)
        
        ttk.Button(file_frame, text="Remove", 
                  command=self._remove_selected_file).grid(row=0, column=2)
    
    def _create_output_section(self, parent):
        """Create output directory section"""
        output_frame = ttk.LabelFrame(parent, text="Output Settings", padding="5")
        output_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Output Directory:").grid(row=0, column=0, sticky=tk.W)
        
        self.output_var = tk.StringVar(value=self.output_directory)
        output_entry = ttk.Entry(output_frame, textvariable=self.output_var, state='readonly')
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        
        ttk.Button(output_frame, text="Browse", 
                  command=self._select_output_directory).grid(row=0, column=2)
    
    def _create_format_selection_section(self, parent):
        """Create format selection section"""
        format_frame = ttk.LabelFrame(parent, text="Output Formats", padding="5")
        format_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.format_vars = {}
        formats = [
            ('EML', 'eml', 'Individual email files'),
            ('MBOX', 'mbox', 'Mailbox format'),
            ('PDF', 'pdf', 'Portable Document Format')
        ]
        
        for i, (display_name, format_code, description) in enumerate(formats):
            var = tk.BooleanVar()
            self.format_vars[format_code] = var
            
            cb = ttk.Checkbutton(format_frame, text=f"{display_name} - {description}", 
                               variable=var)
            cb.grid(row=i, column=0, sticky=tk.W, pady=2)
        
        self.format_vars['eml'].set(True)
    
    def _create_preview_section(self, parent):
        """Create preview section"""
        preview_frame = ttk.LabelFrame(parent, text="PST File Preview", padding="5")
        preview_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        self.preview_tree = ttk.Treeview(preview_frame, columns=('Subject', 'From', 'Date', 'Folder'), 
                                        show='tree headings', height=10)
        
        self.preview_tree.heading('#0', text='#')
        self.preview_tree.heading('Subject', text='Subject')
        self.preview_tree.heading('From', text='From')
        self.preview_tree.heading('Date', text='Date')
        self.preview_tree.heading('Folder', text='Folder')
        
        self.preview_tree.column('#0', width=50)
        self.preview_tree.column('Subject', width=250)
        self.preview_tree.column('From', width=200)
        self.preview_tree.column('Date', width=150)
        self.preview_tree.column('Folder', width=150)
        
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, 
                                        command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=preview_scrollbar.set)
        
        self.preview_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        preview_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    
    def _create_progress_section(self, parent):
        """Create progress section"""
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding="5")
        progress_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                          maximum=100, length=400)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        status_label.grid(row=1, column=0, sticky=tk.W)
    
    def _create_control_buttons(self, parent):
        """Create control buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=5, column=0, columnspan=2, pady=(10, 0))
        
        self.convert_button = ttk.Button(button_frame, text="Start Conversion", 
                                       command=self._start_conversion)
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Clear All", 
                  command=self._clear_all).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="View Logs", 
                  command=self._show_logs).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Exit", 
                  command=self.root.quit).pack(side=tk.LEFT)
    
    def _setup_logging(self):
        """Set up logging configuration"""
        # Create a log handler that can be used by the GUI
        self.log_messages = []
        
        class GUILogHandler(logging.Handler):
            def __init__(self, gui_window):
                super().__init__()
                self.gui_window = gui_window
                
            def emit(self, record):
                msg = self.format(record)
                self.gui_window.log_messages.append(msg)
                # Keep only last 100 log messages
                if len(self.gui_window.log_messages) > 100:
                    self.gui_window.log_messages.pop(0)
        
        # Set up logging
        logging.basicConfig(level=logging.INFO,
                          format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        
        # Add GUI handler
        gui_handler = GUILogHandler(self)
        gui_handler.setLevel(logging.INFO)
        logging.getLogger().addHandler(gui_handler)
    
    def _add_pst_files(self):
        """Add PST files to the list"""
        files = filedialog.askopenfilenames(
            title="Select PST Files",
            filetypes=[("PST files", "*.pst"), ("All files", "*.*")]
        )
        
        for file_path in files:
            if file_path not in self.pst_files:
                self.pst_files.append(file_path)
                self.file_listbox.insert(tk.END, Path(file_path).name)
        
        self._update_convert_button_state()
    
    def _remove_selected_file(self):
        """Remove selected file from the list"""
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.file_listbox.delete(index)
            del self.pst_files[index]
            self._clear_preview()
        
        self._update_convert_button_state()
    
    def _select_output_directory(self):
        """Select output directory"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_directory = directory
            self.output_var.set(directory)
    
    def _preview_pst_file(self, event=None):
        """Preview selected PST file"""
        selection = self.file_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        pst_file_path = self.pst_files[index]
        
        def load_preview():
            try:
                self.status_var.set(f"Loading preview for {Path(pst_file_path).name}...")
                self.root.update()
                
                with PSTParser(pst_file_path) as parser:
                    messages = list(parser.extract_messages())
                    self.current_messages = messages[:100]  # Limit preview to 100 messages
                    
                    self.root.after(0, lambda: self._update_preview(self.current_messages))
                    
            except Exception as e:
                error_msg = f"Error loading PST file: {str(e)}"
                self.root.after(0, lambda: self.status_var.set(error_msg))
                logger.error(error_msg)
        
        threading.Thread(target=load_preview, daemon=True).start()
    
    def _update_preview(self, messages: List[EmailMessage]):
        """Update preview tree with messages"""
        self._clear_preview()
        
        for i, message in enumerate(messages):
            sender = message.sender[:30] + "..." if len(message.sender) > 30 else message.sender
            subject = message.subject[:40] + "..." if len(message.subject) > 40 else message.subject
            
            date_str = ""
            if message.sent_time:
                date_str = message.sent_time.strftime("%Y-%m-%d %H:%M")
            elif message.received_time:
                date_str = message.received_time.strftime("%Y-%m-%d %H:%M")
            
            self.preview_tree.insert('', 'end', text=str(i+1),
                                   values=(subject, sender, date_str, message.folder_path))
        
        self.status_var.set(f"Loaded {len(messages)} messages for preview")
    
    def _clear_preview(self):
        """Clear preview tree"""
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
    
    def _update_convert_button_state(self):
        """Update convert button state based on file selection"""
        if self.pst_files:
            self.convert_button.config(state='normal')
        else:
            self.convert_button.config(state='disabled')
    
    def _start_conversion(self):
        """Start the conversion process"""
        if not self.pst_files:
            messagebox.showwarning("No Files", "Please select PST files to convert.")
            return
        
        selected_formats = [fmt for fmt, var in self.format_vars.items() if var.get()]
        if not selected_formats:
            messagebox.showwarning("No Format", "Please select at least one output format.")
            return
        
        self.convert_button.config(state='disabled')
        self.progress_var.set(0)
        
        def conversion_thread():
            try:
                all_messages = []
                
                for i, pst_file in enumerate(self.pst_files):
                    file_progress = (i / len(self.pst_files)) * 50
                    self.root.after(0, lambda p=file_progress: self.progress_var.set(p))
                    self.root.after(0, lambda f=Path(pst_file).name: 
                                  self.status_var.set(f"Reading {f}..."))
                    
                    with PSTParser(pst_file) as parser:
                        messages = list(parser.extract_messages())
                        all_messages.extend(messages)
                
                self.root.after(0, lambda: self.progress_var.set(50))
                self.root.after(0, lambda: self.status_var.set("Starting conversion..."))
                
                converter = Converter(self.output_directory)
                
                def progress_callback(progress, format_name=None, current=None, total=None):
                    overall_progress = 50 + (progress / 2)
                    self.root.after(0, lambda p=overall_progress: self.progress_var.set(p))
                    
                    if format_name and current and total:
                        status = f"Converting to {format_name.upper()}: {current}/{total}"
                        self.root.after(0, lambda s=status: self.status_var.set(s))
                
                results = converter.batch_convert(all_messages, selected_formats, 
                                                progress_callback)
                
                self.root.after(0, lambda: self._show_conversion_results(results))
                
            except Exception as e:
                error_msg = f"Conversion error: {str(e)}"
                self.root.after(0, lambda: self.status_var.set(error_msg))
                self.root.after(0, lambda: messagebox.showerror("Conversion Error", error_msg))
                logger.error(error_msg)
            
            finally:
                self.root.after(0, lambda: self.convert_button.config(state='normal'))
                self.root.after(0, lambda: self.progress_var.set(0))
        
        threading.Thread(target=conversion_thread, daemon=True).start()
    
    def _show_conversion_results(self, results: Dict[str, Dict[str, Any]]):
        """Show conversion results"""
        self.progress_var.set(100)
        self.status_var.set("Conversion completed!")
        
        result_text = "Conversion Results:\n\n"
        for format_name, result in results.items():
            result_text += f"{format_name.upper()}:\n"
            result_text += f"  Converted: {result['converted_count']}\n"
            result_text += f"  Errors: {result['error_count']}\n"
            result_text += f"  Output: {result['output_directory']}\n\n"
        
        messagebox.showinfo("Conversion Complete", result_text)
    
    def _clear_all(self):
        """Clear all files and preview"""
        self.pst_files.clear()
        self.file_listbox.delete(0, tk.END)
        self._clear_preview()
        self.current_messages.clear()
        self.progress_var.set(0)
        self.status_var.set("Ready")
        self._update_convert_button_state()
    
    def _show_logs(self):
        """Show log messages in a new window"""
        log_window = tk.Toplevel(self.root)
        log_window.title("Application Logs")
        log_window.geometry("800x600")
        
        # Create scrolled text widget for logs
        log_text = scrolledtext.ScrolledText(log_window, wrap=tk.WORD)
        log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add all log messages
        for log_msg in self.log_messages:
            log_text.insert(tk.END, log_msg + "\n")
        
        # Scroll to bottom
        log_text.see(tk.END)
        
        # Make read-only
        log_text.config(state=tk.DISABLED)
        
        # Add close button
        close_btn = ttk.Button(log_window, text="Close", command=log_window.destroy)
        close_btn.pack(pady=5)
    
    def run(self):
        """Start the GUI application"""
        self.root.mainloop()