"""
Advanced OMR System Desktop Application
A modern GUI application with enhanced features for OMR answer sheet processing.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
import threading
from pathlib import Path
from simple_omr_demo import SimpleOMRProcessor, AnswerSheet
import webbrowser
from datetime import datetime
import csv


class ModernOMRApp:
    """Modern OMR Desktop Application with enhanced features."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("OMR System Pro - Answer Sheet Processor")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f8f9fa')
        
        # Set modern style
        self.setup_styles()
        
        # Initialize OMR processor
        self.omr_processor = SimpleOMRProcessor()
        self.results = []
        self.answer_key_path = ""
        self.input_path = ""
        self.current_report = None
        
        # Create GUI
        self.create_widgets()
        self.setup_layout()
        
        # Load default settings
        self.load_defaults()
    
    def setup_styles(self):
        """Setup modern styles for the application."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure styles
        style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), background='#f8f9fa')
        style.configure('Heading.TLabel', font=('Segoe UI', 12, 'bold'), background='#f8f9fa')
        style.configure('Modern.TButton', font=('Segoe UI', 10, 'bold'))
        style.configure('Success.TLabel', foreground='#28a745', font=('Segoe UI', 10, 'bold'))
        style.configure('Error.TLabel', foreground='#dc3545', font=('Segoe UI', 10, 'bold'))
    
    def create_widgets(self):
        """Create all GUI widgets."""
        
        # Main container
        self.main_container = tk.Frame(self.root, bg='#f8f9fa')
        
        # Header
        self.header_frame = tk.Frame(self.main_container, bg='#343a40', height=80)
        self.header_frame.pack_propagate(False)
        
        self.title_label = tk.Label(
            self.header_frame,
            text="OMR System Pro",
            font=('Segoe UI', 24, 'bold'),
            bg='#343a40',
            fg='white'
        )
        
        self.subtitle_label = tk.Label(
            self.header_frame,
            text="Advanced Answer Sheet Processing System",
            font=('Segoe UI', 12),
            bg='#343a40',
            fg='#adb5bd'
        )
        
        # Content area with notebook
        self.notebook = ttk.Notebook(self.main_container)
        
        # Processing Tab
        self.processing_tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(self.processing_tab, text="Processing")
        
        # Results Tab
        self.results_tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(self.results_tab, text="Results")
        
        # Statistics Tab
        self.stats_tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(self.stats_tab, text="Statistics")
        
        # Settings Tab
        self.settings_tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(self.settings_tab, text="Settings")
        
        # Create tab contents
        self.create_processing_tab()
        self.create_results_tab()
        self.create_statistics_tab()
        self.create_settings_tab()
        
        # Status bar
        self.status_frame = tk.Frame(self.main_container, bg='#e9ecef', height=30)
        self.status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(
            self.status_frame,
            text="Ready to process answer sheets",
            bg='#e9ecef',
            fg='#495057',
            font=('Segoe UI', 9)
        )
        
        self.progress_bar = ttk.Progressbar(
            self.status_frame,
            mode='indeterminate',
            length=200
        )
    
    def create_processing_tab(self):
        """Create the processing tab content."""
        
        # File Selection Section
        file_section = tk.LabelFrame(
            self.processing_tab,
            text="File Selection",
            font=('Segoe UI', 12, 'bold'),
            bg='#f8f9fa',
            fg='#495057'
        )
        file_section.pack(fill='x', padx=20, pady=10)
        
        # Answer Key
        answer_key_frame = tk.Frame(file_section, bg='#f8f9fa')
        answer_key_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(answer_key_frame, text="Answer Key (JSON):", bg='#f8f9fa', font=('Segoe UI', 10)).pack(anchor='w')
        self.answer_key_frame = tk.Frame(answer_key_frame, bg='#f8f9fa')
        self.answer_key_frame.pack(fill='x', pady=2)
        
        self.answer_key_entry = tk.Entry(
            self.answer_key_frame,
            font=('Segoe UI', 10),
            relief='solid',
            bd=1
        )
        self.answer_key_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        
        self.answer_key_button = tk.Button(
            self.answer_key_frame,
            text="Browse",
            command=self.browse_answer_key,
            bg='#007bff',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            padx=20
        )
        self.answer_key_button.pack(side='right')
        
        # Input Path
        input_frame = tk.Frame(file_section, bg='#f8f9fa')
        input_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(input_frame, text="Input (File/Folder):", bg='#f8f9fa', font=('Segoe UI', 10)).pack(anchor='w')
        self.input_frame = tk.Frame(input_frame, bg='#f8f9fa')
        self.input_frame.pack(fill='x', pady=2)
        
        self.input_entry = tk.Entry(
            self.input_frame,
            font=('Segoe UI', 10),
            relief='solid',
            bd=1
        )
        self.input_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        
        self.input_button = tk.Button(
            self.input_frame,
            text="Browse",
            command=self.browse_input,
            bg='#007bff',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            padx=20
        )
        self.input_button.pack(side='right')
        
        # Processing Controls
        controls_section = tk.LabelFrame(
            self.processing_tab,
            text="Processing Controls",
            font=('Segoe UI', 12, 'bold'),
            bg='#f8f9fa',
            fg='#495057'
        )
        controls_section.pack(fill='x', padx=20, pady=10)
        
        controls_frame = tk.Frame(controls_section, bg='#f8f9fa')
        controls_frame.pack(fill='x', padx=10, pady=10)
        
        self.process_button = tk.Button(
            controls_frame,
            text="🚀 Process Answer Sheets",
            command=self.start_processing,
            bg='#28a745',
            fg='white',
            font=('Segoe UI', 14, 'bold'),
            relief='flat',
            padx=30,
            pady=10
        )
        self.process_button.pack(side='left', padx=10)
        
        self.clear_button = tk.Button(
            controls_frame,
            text="🗑️ Clear Results",
            command=self.clear_results,
            bg='#dc3545',
            fg='white',
            font=('Segoe UI', 12, 'bold'),
            relief='flat',
            padx=20,
            pady=10
        )
        self.clear_button.pack(side='left', padx=10)
        
        # Quick Actions
        quick_frame = tk.Frame(controls_section, bg='#f8f9fa')
        quick_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(quick_frame, text="Quick Actions:", bg='#f8f9fa', font=('Segoe UI', 10, 'bold')).pack(anchor='w')
        
        quick_buttons_frame = tk.Frame(quick_frame, bg='#f8f9fa')
        quick_buttons_frame.pack(fill='x', pady=5)
        
        self.create_sample_button = tk.Button(
            quick_buttons_frame,
            text="📄 Create Sample Data",
            command=self.create_sample_data,
            bg='#6c757d',
            fg='white',
            font=('Segoe UI', 10),
            relief='flat',
            padx=15
        )
        self.create_sample_button.pack(side='left', padx=5)
        
        self.load_demo_button = tk.Button(
            quick_buttons_frame,
            text="🎯 Load Demo",
            command=self.load_demo_data,
            bg='#17a2b8',
            fg='white',
            font=('Segoe UI', 10),
            relief='flat',
            padx=15
        )
        self.load_demo_button.pack(side='left', padx=5)
    
    def create_results_tab(self):
        """Create the results tab content."""
        
        # Results Display
        results_display_frame = tk.Frame(self.results_tab, bg='#f8f9fa')
        results_display_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Results Text Area
        self.results_text = scrolledtext.ScrolledText(
            results_display_frame,
            font=('Consolas', 10),
            bg='#ffffff',
            fg='#212529',
            relief='solid',
            bd=1,
            wrap=tk.WORD
        )
        self.results_text.pack(fill='both', expand=True)
        
        # Export Controls
        export_frame = tk.Frame(self.results_tab, bg='#f8f9fa')
        export_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Label(export_frame, text="Export Options:", bg='#f8f9fa', font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        
        export_buttons_frame = tk.Frame(export_frame, bg='#f8f9fa')
        export_buttons_frame.pack(fill='x', pady=5)
        
        self.export_json_button = tk.Button(
            export_buttons_frame,
            text="📄 Export JSON",
            command=self.export_json,
            bg='#e74c3c',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            padx=15
        )
        self.export_json_button.pack(side='left', padx=5)
        
        self.export_csv_button = tk.Button(
            export_buttons_frame,
            text="📊 Export CSV",
            command=self.export_csv,
            bg='#e74c3c',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            padx=15
        )
        self.export_csv_button.pack(side='left', padx=5)
        
        self.open_results_button = tk.Button(
            export_buttons_frame,
            text="📁 Open Results Folder",
            command=self.open_results_folder,
            bg='#9b59b6',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            padx=15
        )
        self.open_results_button.pack(side='left', padx=5)
        
        self.print_results_button = tk.Button(
            export_buttons_frame,
            text="🖨️ Print Results",
            command=self.print_results,
            bg='#6c757d',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            padx=15
        )
        self.print_results_button.pack(side='left', padx=5)
    
    def create_statistics_tab(self):
        """Create the statistics tab content."""
        
        # Statistics Display
        stats_display_frame = tk.Frame(self.stats_tab, bg='#f8f9fa')
        stats_display_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Summary Statistics
        summary_frame = tk.LabelFrame(
            stats_display_frame,
            text="Summary Statistics",
            font=('Segoe UI', 12, 'bold'),
            bg='#f8f9fa',
            fg='#495057'
        )
        summary_frame.pack(fill='x', pady=5)
        
        self.summary_stats_frame = tk.Frame(summary_frame, bg='#f8f9fa')
        self.summary_stats_frame.pack(fill='x', padx=10, pady=10)
        
        # Score Distribution
        distribution_frame = tk.LabelFrame(
            stats_display_frame,
            text="Score Distribution",
            font=('Segoe UI', 12, 'bold'),
            bg='#f8f9fa',
            fg='#495057'
        )
        distribution_frame.pack(fill='x', pady=5)
        
        self.distribution_frame = tk.Frame(distribution_frame, bg='#f8f9fa')
        self.distribution_frame.pack(fill='x', padx=10, pady=10)
        
        # Accuracy Metrics
        accuracy_frame = tk.LabelFrame(
            stats_display_frame,
            text="Accuracy Metrics",
            font=('Segoe UI', 12, 'bold'),
            bg='#f8f9fa',
            fg='#495057'
        )
        accuracy_frame.pack(fill='x', pady=5)
        
        self.accuracy_frame = tk.Frame(accuracy_frame, bg='#f8f9fa')
        self.accuracy_frame.pack(fill='x', padx=10, pady=10)
    
    def create_settings_tab(self):
        """Create the settings tab content."""
        
        # Settings Display
        settings_display_frame = tk.Frame(self.settings_tab, bg='#f8f9fa')
        settings_display_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Processing Settings
        processing_settings = tk.LabelFrame(
            settings_display_frame,
            text="Processing Settings",
            font=('Segoe UI', 12, 'bold'),
            bg='#f8f9fa',
            fg='#495057'
        )
        processing_settings.pack(fill='x', pady=5)
        
        # Detection Threshold
        threshold_frame = tk.Frame(processing_settings, bg='#f8f9fa')
        threshold_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(threshold_frame, text="Mark Detection Threshold:", bg='#f8f9fa', font=('Segoe UI', 10)).pack(side='left')
        self.threshold_var = tk.DoubleVar(value=0.3)
        self.threshold_scale = tk.Scale(
            threshold_frame,
            from_=0.1,
            to=0.9,
            resolution=0.1,
            orient='horizontal',
            variable=self.threshold_var,
            bg='#f8f9fa'
        )
        self.threshold_scale.pack(side='left', padx=10)
        
        # Output Settings
        output_settings = tk.LabelFrame(
            settings_display_frame,
            text="Output Settings",
            font=('Segoe UI', 12, 'bold'),
            bg='#f8f9fa',
            fg='#495057'
        )
        output_settings.pack(fill='x', pady=5)
        
        # Auto-save
        self.auto_save_var = tk.BooleanVar(value=True)
        self.auto_save_check = tk.Checkbutton(
            output_settings,
            text="Auto-save results after processing",
            variable=self.auto_save_var,
            bg='#f8f9fa',
            font=('Segoe UI', 10)
        )
        self.auto_save_check.pack(anchor='w', padx=10, pady=5)
        
        # Default output directory
        output_dir_frame = tk.Frame(output_settings, bg='#f8f9fa')
        output_dir_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(output_dir_frame, text="Default Output Directory:", bg='#f8f9fa', font=('Segoe UI', 10)).pack(anchor='w')
        self.output_dir_entry = tk.Entry(output_dir_frame, font=('Segoe UI', 10))
        self.output_dir_entry.pack(fill='x', pady=2)
        self.output_dir_entry.insert(0, "app_results")
        
        # About Section
        about_frame = tk.LabelFrame(
            settings_display_frame,
            text="About",
            font=('Segoe UI', 12, 'bold'),
            bg='#f8f9fa',
            fg='#495057'
        )
        about_frame.pack(fill='x', pady=5)
        
        about_text = tk.Text(
            about_frame,
            height=8,
            font=('Segoe UI', 9),
            bg='#ffffff',
            fg='#495057',
            relief='solid',
            bd=1,
            wrap=tk.WORD
        )
        about_text.pack(fill='x', padx=10, pady=10)
        about_text.insert('1.0', """OMR System Pro v1.0
Advanced Answer Sheet Processing System

Features:
• Automated answer detection and scoring
• Batch processing capabilities
• Comprehensive reporting and statistics
• Multiple export formats (JSON, CSV)
• Real-time processing status
• Configurable detection parameters

Developed with Python and Tkinter
© 2024 OMR System Pro""")
        about_text.config(state='disabled')
    
    def setup_layout(self):
        """Setup the layout of widgets."""
        
        # Main container
        self.main_container.pack(fill='both', expand=True)
        
        # Header
        self.header_frame.pack(fill='x')
        self.title_label.pack(pady=10)
        self.subtitle_label.pack()
        
        # Notebook
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Status bar
        self.status_frame.pack(fill='x', side='bottom')
        self.status_label.pack(side='left', padx=10, pady=5)
        self.progress_bar.pack(side='right', padx=10, pady=5)
    
    def load_defaults(self):
        """Load default settings."""
        # Set default paths if they exist
        if os.path.exists("sample_answer_key.json"):
            self.answer_key_entry.insert(0, "sample_answer_key.json")
            self.answer_key_path = "sample_answer_key.json"
        
        if os.path.exists("sample_sheets"):
            self.input_entry.insert(0, "sample_sheets")
            self.input_path = "sample_sheets"
    
    def browse_answer_key(self):
        """Browse for answer key JSON file."""
        filename = filedialog.askopenfilename(
            title="Select Answer Key",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.answer_key_entry.delete(0, tk.END)
            self.answer_key_entry.insert(0, filename)
            self.answer_key_path = filename
    
    def browse_input(self):
        """Browse for input file or folder."""
        choice = messagebox.askyesno(
            "Input Type", 
            "Yes for folder, No for single file"
        )
        
        if choice:  # Folder
            folder = filedialog.askdirectory(title="Select Input Folder")
            if folder:
                self.input_entry.delete(0, tk.END)
                self.input_entry.insert(0, folder)
                self.input_path = folder
        else:  # File
            filename = filedialog.askopenfilename(
                title="Select Input File",
                filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff"), ("All files", "*.*")]
            )
            if filename:
                self.input_entry.delete(0, tk.END)
                self.input_entry.insert(0, filename)
                self.input_path = filename
    
    def create_sample_data(self):
        """Create sample data for testing."""
        try:
            # Create sample answer key
            answer_key = {
                "exam_info": {
                    "title": "Sample Mathematics Exam",
                    "date": "2024-01-15",
                    "total_questions": 20,
                    "options_per_question": 4
                },
                "answers": ["A", "B", "C", "D", "A", "B", "C", "D", "A", "B", "C", "D", "A", "B", "C", "D", "A", "B", "C", "D"],
                "question_weights": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
            }
            
            with open("sample_answer_key.json", "w") as f:
                json.dump(answer_key, f, indent=2)
            
            # Create sample sheets directory
            os.makedirs("sample_sheets", exist_ok=True)
            
            sample_files = [
                "student_perfect_001.png",
                "student_good_002.png", 
                "student_average_003.png",
                "student_poor_004.png",
                "student_error_005.png"
            ]
            
            for filename in sample_files:
                with open(os.path.join("sample_sheets", filename), "w") as f:
                    f.write("Mock image file")
            
            # Update UI
            self.answer_key_entry.delete(0, tk.END)
            self.answer_key_entry.insert(0, "sample_answer_key.json")
            self.answer_key_path = "sample_answer_key.json"
            
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, "sample_sheets")
            self.input_path = "sample_sheets"
            
            messagebox.showinfo("Success", "Sample data created successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create sample data: {str(e)}")
    
    def load_demo_data(self):
        """Load demo data and process it."""
        self.create_sample_data()
        self.start_processing()
    
    def start_processing(self):
        """Start processing in a separate thread."""
        if not self.answer_key_path or not self.input_path:
            messagebox.showerror("Error", "Please select both answer key and input path")
            return
        
        # Start processing in a separate thread
        self.processing_thread = threading.Thread(target=self.process_files)
        self.processing_thread.daemon = True
        self.processing_thread.start()
    
    def process_files(self):
        """Process the selected files."""
        try:
            # Update UI
            self.root.after(0, self.update_status, "Loading answer key...")
            self.root.after(0, self.progress_bar.start)
            
            # Load answer key
            self.omr_processor.load_answer_key(self.answer_key_path)
            
            # Update UI
            self.root.after(0, self.update_status, "Processing answer sheets...")
            
            # Process files
            if os.path.isfile(self.input_path):
                # Single file
                result = self.omr_processor.process_answer_sheet(self.input_path)
                self.results = [result]
            else:
                # Directory
                self.results = self.omr_processor.process_batch(self.input_path)
            
            # Generate report
            self.root.after(0, self.update_status, "Generating report...")
            output_dir = self.output_dir_entry.get() or "app_results"
            os.makedirs(output_dir, exist_ok=True)
            report = self.omr_processor.generate_report(self.results, f"{output_dir}/report.json")
            self.current_report = report
            
            # Update UI with results
            self.root.after(0, self.display_results, report)
            self.root.after(0, self.update_statistics, report)
            self.root.after(0, self.update_status, f"✅ Completed! Processed {len(self.results)} sheets")
            self.root.after(0, self.progress_bar.stop)
            
            # Auto-save if enabled
            if self.auto_save_var.get():
                self.root.after(0, self.auto_save_results)
            
        except Exception as e:
            self.root.after(0, self.update_status, f"❌ Error: {str(e)}")
            self.root.after(0, self.progress_bar.stop)
            self.root.after(0, messagebox.showerror, "Processing Error", str(e))
    
    def update_status(self, message):
        """Update status label."""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def display_results(self, report):
        """Display results in the text area."""
        self.results_text.delete(1.0, tk.END)
        
        # Summary
        summary = report['summary']
        self.results_text.insert(tk.END, "="*80 + "\n")
        self.results_text.insert(tk.END, "OMR PROCESSING RESULTS\n")
        self.results_text.insert(tk.END, "="*80 + "\n")
        self.results_text.insert(tk.END, f"Processing Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        self.results_text.insert(tk.END, f"Total sheets processed: {summary['total_sheets_processed']}\n")
        self.results_text.insert(tk.END, f"Average score: {summary['average_score']}%\n")
        self.results_text.insert(tk.END, f"Success rate: {summary['success_rate']}%\n")
        self.results_text.insert(tk.END, f"Average confidence: {summary['average_confidence']:.3f}\n")
        self.results_text.insert(tk.END, f"Processing errors: {summary['processing_errors']}\n\n")
        
        # Score Distribution
        self.results_text.insert(tk.END, "SCORE DISTRIBUTION\n")
        self.results_text.insert(tk.END, "-"*50 + "\n")
        for range_name, count in report['score_distribution'].items():
            percentage = (count / summary['total_sheets_processed']) * 100 if summary['total_sheets_processed'] > 0 else 0
            self.results_text.insert(tk.END, f"{range_name}: {count} students ({percentage:.1f}%)\n")
        self.results_text.insert(tk.END, "\n")
        
        # Pipeline Accuracy
        accuracy = report['pipeline_accuracy']
        self.results_text.insert(tk.END, "PIPELINE ACCURACY\n")
        self.results_text.insert(tk.END, "-"*50 + "\n")
        self.results_text.insert(tk.END, f"Mark detection accuracy: {accuracy['mark_detection_accuracy']}%\n")
        self.results_text.insert(tk.END, f"Grid detection success rate: {accuracy['grid_detection_success_rate']}%\n\n")
        
        # Detailed Results
        self.results_text.insert(tk.END, "DETAILED RESULTS\n")
        self.results_text.insert(tk.END, "-"*50 + "\n")
        for i, result in enumerate(report['detailed_results'], 1):
            self.results_text.insert(tk.END, f"{i:2d}. Student {result['student_id']}: {result['score']}% ")
            self.results_text.insert(tk.END, f"(Correct: {result['correct_answers']}, ")
            self.results_text.insert(tk.END, f"Incorrect: {result['incorrect_answers']}, ")
            self.results_text.insert(tk.END, f"Blank: {result['blank_answers']}, ")
            self.results_text.insert(tk.END, f"Confidence: {result['confidence']:.3f})\n")
            if result['errors']:
                self.results_text.insert(tk.END, f"     Errors: {', '.join(result['errors'])}\n")
        
        # Switch to results tab
        self.notebook.select(1)
    
    def update_statistics(self, report):
        """Update statistics display."""
        summary = report['summary']
        accuracy = report['pipeline_accuracy']
        
        # Clear existing statistics
        for widget in self.summary_stats_frame.winfo_children():
            widget.destroy()
        for widget in self.distribution_frame.winfo_children():
            widget.destroy()
        for widget in self.accuracy_frame.winfo_children():
            widget.destroy()
        
        # Summary Statistics
        stats_data = [
            ("Total Sheets", str(summary['total_sheets_processed'])),
            ("Average Score", f"{summary['average_score']}%"),
            ("Success Rate", f"{summary['success_rate']}%"),
            ("Processing Errors", str(summary['processing_errors']))
        ]
        
        for i, (label, value) in enumerate(stats_data):
            row = i // 2
            col = (i % 2) * 2
            
            tk.Label(self.summary_stats_frame, text=f"{label}:", bg='#f8f9fa', font=('Segoe UI', 10, 'bold')).grid(row=row, column=col, sticky='w', padx=5, pady=2)
            tk.Label(self.summary_stats_frame, text=value, bg='#f8f9fa', font=('Segoe UI', 10)).grid(row=row, column=col+1, sticky='w', padx=5, pady=2)
        
        # Score Distribution
        for range_name, count in report['score_distribution'].items():
            percentage = (count / summary['total_sheets_processed']) * 100 if summary['total_sheets_processed'] > 0 else 0
            frame = tk.Frame(self.distribution_frame, bg='#f8f9fa')
            frame.pack(fill='x', pady=2)
            
            tk.Label(frame, text=f"{range_name}:", bg='#f8f9fa', font=('Segoe UI', 10), width=10, anchor='w').pack(side='left')
            tk.Label(frame, text=f"{count} students ({percentage:.1f}%)", bg='#f8f9fa', font=('Segoe UI', 10)).pack(side='left', padx=10)
        
        # Accuracy Metrics
        accuracy_data = [
            ("Mark Detection Accuracy", f"{accuracy['mark_detection_accuracy']}%"),
            ("Grid Detection Success", f"{accuracy['grid_detection_success_rate']}%")
        ]
        
        for label, value in accuracy_data:
            frame = tk.Frame(self.accuracy_frame, bg='#f8f9fa')
            frame.pack(fill='x', pady=2)
            
            tk.Label(frame, text=f"{label}:", bg='#f8f9fa', font=('Segoe UI', 10, 'bold'), width=20, anchor='w').pack(side='left')
            tk.Label(frame, text=value, bg='#f8f9fa', font=('Segoe UI', 10)).pack(side='left', padx=10)
    
    def clear_results(self):
        """Clear all results."""
        self.results = []
        self.current_report = None
        self.results_text.delete(1.0, tk.END)
        self.update_status("Results cleared")
        
        # Clear statistics
        for widget in self.summary_stats_frame.winfo_children():
            widget.destroy()
        for widget in self.distribution_frame.winfo_children():
            widget.destroy()
        for widget in self.accuracy_frame.winfo_children():
            widget.destroy()
    
    def auto_save_results(self):
        """Auto-save results if enabled."""
        if self.auto_save_var.get() and self.current_report:
            try:
                output_dir = self.output_dir_entry.get() or "app_results"
                os.makedirs(output_dir, exist_ok=True)
                
                # Save JSON report
                with open(f"{output_dir}/auto_save_report.json", 'w') as f:
                    json.dump(self.current_report, f, indent=2)
                
                # Save CSV
                self.save_csv(f"{output_dir}/auto_save_results.csv")
                
                self.update_status("✅ Results auto-saved")
            except Exception as e:
                self.update_status(f"⚠️ Auto-save failed: {str(e)}")
    
    def export_json(self):
        """Export results to JSON."""
        if not self.results:
            messagebox.showwarning("Warning", "No results to export")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                if self.current_report:
                    with open(filename, 'w') as f:
                        json.dump(self.current_report, f, indent=2)
                else:
                    report = self.omr_processor.generate_report(self.results)
                    with open(filename, 'w') as f:
                        json.dump(report, f, indent=2)
                messagebox.showinfo("Success", f"Results exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")
    
    def export_csv(self):
        """Export results to CSV."""
        if not self.results:
            messagebox.showwarning("Warning", "No results to export")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                self.save_csv(filename)
                messagebox.showinfo("Success", f"Results exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")
    
    def save_csv(self, filename):
        """Save results to CSV file."""
        with open(filename, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Student ID', 'Score', 'Correct', 'Incorrect', 'Blank', 'Confidence', 'Errors'])
            
            for result in self.results:
                errors_str = "; ".join(result.processing_errors) if result.processing_errors else ""
                confidence = sum(result.confidence_scores) / len(result.confidence_scores) if result.confidence_scores else 0
                writer.writerow([
                    result.student_id,
                    result.score,
                    result.correct_answers,
                    result.incorrect_answers,
                    result.blank_answers,
                    f"{confidence:.3f}",
                    errors_str
                ])
    
    def open_results_folder(self):
        """Open the results folder."""
        output_dir = self.output_dir_entry.get() or "app_results"
        if os.path.exists(output_dir):
            os.startfile(output_dir)  # Windows
        else:
            messagebox.showwarning("Warning", "No results folder found")
    
    def print_results(self):
        """Print results."""
        if not self.results:
            messagebox.showwarning("Warning", "No results to print")
            return
        
        # Create a new window for printing
        print_window = tk.Toplevel(self.root)
        print_window.title("Print Results")
        print_window.geometry("800x600")
        
        print_text = scrolledtext.ScrolledText(print_window, font=('Consolas', 10))
        print_text.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Copy results text
        print_text.insert('1.0', self.results_text.get('1.0', tk.END))
        
        # Print button
        print_button = tk.Button(
            print_window,
            text="Print",
            command=lambda: print_text.print(),
            bg='#007bff',
            fg='white',
            font=('Segoe UI', 10, 'bold')
        )
        print_button.pack(pady=10)


def main():
    """Main function to run the application."""
    # Create and run the application
    root = tk.Tk()
    app = ModernOMRApp(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    # Start the application
    root.mainloop()


if __name__ == "__main__":
    main()
