"""
OMR System Desktop Application
A user-friendly GUI application for OMR answer sheet processing.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
import threading
from pathlib import Path
from simple_omr_demo import SimpleOMRProcessor, AnswerSheet
import webbrowser


class OMRApp:
    """Main OMR Desktop Application."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("OMR System - Answer Sheet Processor")
        self.root.geometry("1000x700")
        self.root.configure(bg='#f0f0f0')
        
        # Initialize OMR processor
        self.omr_processor = SimpleOMRProcessor()
        self.results = []
        self.answer_key_path = ""
        self.input_path = ""
        
        # Create GUI
        self.create_widgets()
        self.setup_layout()
        
    def create_widgets(self):
        """Create all GUI widgets."""
        
        # Title
        self.title_label = tk.Label(
            self.root, 
            text="OMR Answer Sheet Processing System",
            font=("Arial", 16, "bold"),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        
        # File Selection Frame
        self.file_frame = tk.LabelFrame(
            self.root, 
            text="File Selection", 
            font=("Arial", 10, "bold"),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        
        # Answer Key Selection
        self.answer_key_label = tk.Label(
            self.file_frame, 
            text="Answer Key (JSON):", 
            bg='#f0f0f0'
        )
        self.answer_key_entry = tk.Entry(
            self.file_frame, 
            width=50, 
            font=("Arial", 9)
        )
        self.answer_key_button = tk.Button(
            self.file_frame, 
            text="Browse", 
            command=self.browse_answer_key,
            bg='#3498db',
            fg='white',
            font=("Arial", 9, "bold")
        )
        
        # Input Path Selection
        self.input_label = tk.Label(
            self.file_frame, 
            text="Input (File/Folder):", 
            bg='#f0f0f0'
        )
        self.input_entry = tk.Entry(
            self.file_frame, 
            width=50, 
            font=("Arial", 9)
        )
        self.input_button = tk.Button(
            self.file_frame, 
            text="Browse", 
            command=self.browse_input,
            bg='#3498db',
            fg='white',
            font=("Arial", 9, "bold")
        )
        
        # Processing Frame
        self.processing_frame = tk.LabelFrame(
            self.root, 
            text="Processing", 
            font=("Arial", 10, "bold"),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        
        # Process Button
        self.process_button = tk.Button(
            self.processing_frame, 
            text="Process Answer Sheets", 
            command=self.start_processing,
            bg='#27ae60',
            fg='white',
            font=("Arial", 12, "bold"),
            height=2
        )
        
        # Progress Bar
        self.progress = ttk.Progressbar(
            self.processing_frame, 
            mode='indeterminate'
        )
        
        # Status Label
        self.status_label = tk.Label(
            self.processing_frame, 
            text="Ready to process", 
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        
        # Results Frame
        self.results_frame = tk.LabelFrame(
            self.root, 
            text="Results", 
            font=("Arial", 10, "bold"),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        
        # Results Text Area
        self.results_text = scrolledtext.ScrolledText(
            self.results_frame, 
            height=15, 
            width=80,
            font=("Consolas", 9),
            bg='#ffffff',
            fg='#2c3e50'
        )
        
        # Export Buttons
        self.export_frame = tk.Frame(self.results_frame, bg='#f0f0f0')
        
        self.export_json_button = tk.Button(
            self.export_frame, 
            text="Export JSON", 
            command=self.export_json,
            bg='#e74c3c',
            fg='white',
            font=("Arial", 9, "bold")
        )
        
        self.export_csv_button = tk.Button(
            self.export_frame, 
            text="Export CSV", 
            command=self.export_csv,
            bg='#e74c3c',
            fg='white',
            font=("Arial", 9, "bold")
        )
        
        self.open_results_button = tk.Button(
            self.export_frame, 
            text="Open Results Folder", 
            command=self.open_results_folder,
            bg='#9b59b6',
            fg='white',
            font=("Arial", 9, "bold")
        )
        
        # Statistics Frame
        self.stats_frame = tk.LabelFrame(
            self.root, 
            text="Statistics", 
            font=("Arial", 10, "bold"),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        
        # Statistics Labels
        self.stats_labels = {}
        stats_items = [
            "Total Sheets", "Average Score", "Success Rate", 
            "Processing Errors", "Mark Detection Accuracy"
        ]
        
        for i, item in enumerate(stats_items):
            self.stats_labels[item] = tk.Label(
                self.stats_frame, 
                text=f"{item}: -", 
                bg='#f0f0f0',
                fg='#2c3e50',
                font=("Arial", 9)
            )
    
    def setup_layout(self):
        """Setup the layout of widgets."""
        
        # Title
        self.title_label.pack(pady=10)
        
        # File Selection Frame
        self.file_frame.pack(fill='x', padx=10, pady=5)
        
        # Answer Key Row
        self.answer_key_label.grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.answer_key_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.answer_key_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Input Path Row
        self.input_label.grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.input_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self.input_button.grid(row=1, column=2, padx=5, pady=5)
        
        # Configure grid weights
        self.file_frame.grid_columnconfigure(1, weight=1)
        
        # Processing Frame
        self.processing_frame.pack(fill='x', padx=10, pady=5)
        
        self.process_button.pack(pady=10)
        self.progress.pack(fill='x', padx=10, pady=5)
        self.status_label.pack(pady=5)
        
        # Results Frame
        self.results_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.results_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Export Buttons
        self.export_frame.pack(fill='x', padx=5, pady=5)
        self.export_json_button.pack(side='left', padx=5)
        self.export_csv_button.pack(side='left', padx=5)
        self.open_results_button.pack(side='left', padx=5)
        
        # Statistics Frame
        self.stats_frame.pack(fill='x', padx=10, pady=5)
        
        # Statistics Grid
        for i, (key, label) in enumerate(self.stats_labels.items()):
            label.grid(row=0, column=i, padx=10, pady=5, sticky='w')
    
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
            self.root.after(0, self.progress.start)
            
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
            os.makedirs("app_results", exist_ok=True)
            report = self.omr_processor.generate_report(self.results, "app_results/report.json")
            
            # Update UI with results
            self.root.after(0, self.display_results, report)
            self.root.after(0, self.update_status, f"Completed! Processed {len(self.results)} sheets")
            self.root.after(0, self.progress.stop)
            
        except Exception as e:
            self.root.after(0, self.update_status, f"Error: {str(e)}")
            self.root.after(0, self.progress.stop)
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
        self.results_text.insert(tk.END, "="*60 + "\n")
        self.results_text.insert(tk.END, "PROCESSING SUMMARY\n")
        self.results_text.insert(tk.END, "="*60 + "\n")
        self.results_text.insert(tk.END, f"Total sheets processed: {summary['total_sheets_processed']}\n")
        self.results_text.insert(tk.END, f"Average score: {summary['average_score']}%\n")
        self.results_text.insert(tk.END, f"Success rate: {summary['success_rate']}%\n")
        self.results_text.insert(tk.END, f"Average confidence: {summary['average_confidence']:.3f}\n")
        self.results_text.insert(tk.END, f"Processing errors: {summary['processing_errors']}\n\n")
        
        # Score Distribution
        self.results_text.insert(tk.END, "SCORE DISTRIBUTION\n")
        self.results_text.insert(tk.END, "-"*30 + "\n")
        for range_name, count in report['score_distribution'].items():
            self.results_text.insert(tk.END, f"{range_name}: {count} students\n")
        self.results_text.insert(tk.END, "\n")
        
        # Pipeline Accuracy
        accuracy = report['pipeline_accuracy']
        self.results_text.insert(tk.END, "PIPELINE ACCURACY\n")
        self.results_text.insert(tk.END, "-"*30 + "\n")
        self.results_text.insert(tk.END, f"Mark detection accuracy: {accuracy['mark_detection_accuracy']}%\n")
        self.results_text.insert(tk.END, f"Grid detection success rate: {accuracy['grid_detection_success_rate']}%\n\n")
        
        # Detailed Results
        self.results_text.insert(tk.END, "DETAILED RESULTS\n")
        self.results_text.insert(tk.END, "-"*30 + "\n")
        for result in report['detailed_results']:
            self.results_text.insert(tk.END, f"Student {result['student_id']}: {result['score']}% ")
            self.results_text.insert(tk.END, f"(Correct: {result['correct_answers']}, ")
            self.results_text.insert(tk.END, f"Incorrect: {result['incorrect_answers']}, ")
            self.results_text.insert(tk.END, f"Blank: {result['blank_answers']}, ")
            self.results_text.insert(tk.END, f"Confidence: {result['confidence']:.3f})\n")
            if result['errors']:
                self.results_text.insert(tk.END, f"  Errors: {', '.join(result['errors'])}\n")
        
        # Update statistics labels
        self.update_statistics(summary, accuracy)
    
    def update_statistics(self, summary, accuracy):
        """Update statistics labels."""
        self.stats_labels["Total Sheets"].config(text=f"Total Sheets: {summary['total_sheets_processed']}")
        self.stats_labels["Average Score"].config(text=f"Average Score: {summary['average_score']}%")
        self.stats_labels["Success Rate"].config(text=f"Success Rate: {summary['success_rate']}%")
        self.stats_labels["Processing Errors"].config(text=f"Processing Errors: {summary['processing_errors']}")
        self.stats_labels["Mark Detection Accuracy"].config(text=f"Mark Detection: {accuracy['mark_detection_accuracy']}%")
    
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
                # Create CSV content
                csv_content = "Student ID,Score,Correct,Incorrect,Blank,Confidence,Errors\n"
                for result in self.results:
                    errors_str = "; ".join(result.processing_errors) if result.processing_errors else ""
                    confidence = sum(result.confidence_scores) / len(result.confidence_scores) if result.confidence_scores else 0
                    csv_content += f"{result.student_id},{result.score},{result.correct_answers},{result.incorrect_answers},{result.blank_answers},{confidence:.3f},{errors_str}\n"
                
                with open(filename, 'w') as f:
                    f.write(csv_content)
                messagebox.showinfo("Success", f"Results exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")
    
    def open_results_folder(self):
        """Open the results folder."""
        results_path = os.path.abspath("app_results")
        if os.path.exists(results_path):
            os.startfile(results_path)  # Windows
        else:
            messagebox.showwarning("Warning", "No results folder found")
    
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


def create_sample_data():
    """Create sample data for the app."""
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
    
    print("Sample data created successfully!")


def main():
    """Main function to run the application."""
    # Create sample data
    create_sample_data()
    
    # Create and run the application
    root = tk.Tk()
    app = OMRApp(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    # Start the application
    root.mainloop()


if __name__ == "__main__":
    main()
