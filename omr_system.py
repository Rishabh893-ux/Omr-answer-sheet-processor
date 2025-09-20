"""
OMR (Optical Mark Recognition) System
A robust solution for automatically evaluating and scoring answer sheets.
"""

import cv2
import numpy as np
import json
import os
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from pathlib import Path
import matplotlib.pyplot as plt
from PIL import Image
import pandas as pd
from scipy import ndimage
from skimage import measure, morphology
import argparse
from datetime import datetime


@dataclass
class OMRConfig:
    """Configuration class for OMR system parameters."""
    # Image processing parameters
    blur_kernel_size: int = 5
    threshold_block_size: int = 11
    threshold_c: int = 2
    
    # Answer detection parameters
    min_bubble_area: int = 50
    max_bubble_area: int = 2000
    min_contour_area: int = 30
    
    # Grid detection parameters
    min_questions: int = 10
    min_options: int = 4
    grid_tolerance: float = 0.1
    
    # Scoring parameters
    correct_mark_threshold: float = 0.3
    partial_credit: bool = False


@dataclass
class AnswerSheet:
    """Data class to store answer sheet information."""
    student_id: str
    answers: List[str]
    score: float
    total_questions: int
    correct_answers: int
    incorrect_answers: int
    blank_answers: int
    confidence_scores: List[float]
    processing_errors: List[str]


class OMRProcessor:
    """Main OMR processing class."""
    
    def __init__(self, config: OMRConfig = None):
        self.config = config or OMRConfig()
        self.answer_key = []
        self.question_count = 0
        self.option_count = 4  # Default A, B, C, D
        
    def load_answer_key(self, answer_key_path: str) -> None:
        """Load answer key from JSON file."""
        try:
            with open(answer_key_path, 'r') as f:
                data = json.load(f)
                self.answer_key = data.get('answers', [])
                self.question_count = len(self.answer_key)
                print(f"Loaded answer key with {self.question_count} questions")
        except Exception as e:
            raise ValueError(f"Error loading answer key: {e}")
    
    def preprocess_image(self, image_path: str) -> np.ndarray:
        """Preprocess the input image for better mark detection."""
        # Load image
        image = cv2.imread(image_path)
        if image is None:
            raise ValueError(f"Could not load image: {image_path}")
        
        # Convert to grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Apply Gaussian blur to reduce noise
        blurred = cv2.GaussianBlur(gray, (self.config.blur_kernel_size, self.config.blur_kernel_size), 0)
        
        # Apply adaptive thresholding
        thresh = cv2.adaptiveThreshold(
            blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
            cv2.THRESH_BINARY_INV, self.config.threshold_block_size, self.config.threshold_c
        )
        
        return thresh
    
    def detect_answer_grid(self, processed_image: np.ndarray) -> Tuple[List, List]:
        """Detect the answer grid structure."""
        # Find contours
        contours, _ = cv2.findContours(processed_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        # Filter contours by area
        valid_contours = []
        for contour in contours:
            area = cv2.contourArea(contour)
            if self.config.min_contour_area < area < self.config.max_bubble_area:
                valid_contours.append(contour)
        
        if len(valid_contours) < self.config.min_questions * self.config.min_options:
            raise ValueError("Insufficient contours detected. Check image quality.")
        
        # Sort contours by position (top to bottom, left to right)
        valid_contours = sorted(valid_contours, key=lambda c: (cv2.boundingRect(c)[1], cv2.boundingRect(c)[0]))
        
        # Group contours into rows and columns
        rows = []
        current_row = []
        last_y = -1
        
        for contour in valid_contours:
            x, y, w, h = cv2.boundingRect(contour)
            if last_y == -1 or abs(y - last_y) < 20:  # Same row
                current_row.append((contour, x, y, w, h))
            else:  # New row
                if current_row:
                    current_row.sort(key=lambda x: x[1])  # Sort by x coordinate
                    rows.append(current_row)
                current_row = [(contour, x, y, w, h)]
            last_y = y
        
        if current_row:
            current_row.sort(key=lambda x: x[1])
            rows.append(current_row)
        
        return rows, valid_contours
    
    def detect_marked_answers(self, processed_image: np.ndarray, rows: List) -> List[List[bool]]:
        """Detect which answer options are marked."""
        answers = []
        
        for row_idx, row in enumerate(rows):
            row_answers = []
            
            for option_idx, (contour, x, y, w, h) in enumerate(row):
                # Extract the region of interest
                roi = processed_image[y:y+h, x:x+w]
                
                # Count non-zero pixels (marked area)
                marked_pixels = np.count_nonzero(roi)
                total_pixels = roi.shape[0] * roi.shape[1]
                
                # Calculate fill ratio
                fill_ratio = marked_pixels / total_pixels if total_pixels > 0 else 0
                
                # Determine if this option is marked
                is_marked = fill_ratio > self.config.correct_mark_threshold
                row_answers.append(is_marked)
            
            answers.append(row_answers)
        
        return answers
    
    def process_answer_sheet(self, image_path: str) -> AnswerSheet:
        """Process a single answer sheet and return results."""
        errors = []
        
        try:
            # Preprocess image
            processed_image = self.preprocess_image(image_path)
            
            # Detect answer grid
            rows, contours = self.detect_answer_grid(processed_image)
            
            # Detect marked answers
            marked_answers = self.detect_marked_answers(processed_image, rows)
            
            # Convert to answer format
            answers = []
            confidence_scores = []
            
            for row_idx, row_answers in enumerate(marked_answers):
                marked_count = sum(row_answers)
                
                if marked_count == 0:
                    answers.append('')  # Blank answer
                    confidence_scores.append(0.0)
                elif marked_count == 1:
                    # Single answer
                    option_idx = row_answers.index(True)
                    answers.append(chr(ord('A') + option_idx))
                    confidence_scores.append(0.9)  # High confidence for single mark
                else:
                    # Multiple answers or unclear
                    answers.append('MULTIPLE')  # Multiple marks
                    confidence_scores.append(0.3)  # Low confidence
            
            # Calculate score
            score, correct, incorrect, blank = self.calculate_score(answers)
            
            # Extract student ID from filename or image
            student_id = self.extract_student_id(image_path)
            
            return AnswerSheet(
                student_id=student_id,
                answers=answers,
                score=score,
                total_questions=self.question_count,
                correct_answers=correct,
                incorrect_answers=incorrect,
                blank_answers=blank,
                confidence_scores=confidence_scores,
                processing_errors=errors
            )
            
        except Exception as e:
            errors.append(f"Processing error: {str(e)}")
            return AnswerSheet(
                student_id="UNKNOWN",
                answers=[],
                score=0.0,
                total_questions=self.question_count,
                correct_answers=0,
                incorrect_answers=0,
                blank_answers=0,
                confidence_scores=[],
                processing_errors=errors
            )
    
    def calculate_score(self, answers: List[str]) -> Tuple[float, int, int, int]:
        """Calculate score based on answer key."""
        if not self.answer_key:
            return 0.0, 0, 0, len(answers)
        
        correct = 0
        incorrect = 0
        blank = 0
        
        for i, answer in enumerate(answers):
            if i >= len(self.answer_key):
                break
                
            if answer == '':
                blank += 1
            elif answer == 'MULTIPLE':
                incorrect += 1  # Multiple marks are considered incorrect
            elif answer.upper() == self.answer_key[i].upper():
                correct += 1
            else:
                incorrect += 1
        
        total_questions = len(self.answer_key)
        score = (correct / total_questions) * 100 if total_questions > 0 else 0
        
        return score, correct, incorrect, blank
    
    def extract_student_id(self, image_path: str) -> str:
        """Extract student ID from filename or image."""
        filename = os.path.basename(image_path)
        name_without_ext = os.path.splitext(filename)[0]
        
        # Try to extract ID from filename
        if name_without_ext.isdigit():
            return name_without_ext
        
        # Default to filename
        return name_without_ext
    
    def process_batch(self, image_directory: str) -> List[AnswerSheet]:
        """Process multiple answer sheets in a directory."""
        results = []
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']
        
        for file_path in Path(image_directory).iterdir():
            if file_path.suffix.lower() in image_extensions:
                print(f"Processing: {file_path.name}")
                result = self.process_answer_sheet(str(file_path))
                results.append(result)
        
        return results
    
    def generate_report(self, results: List[AnswerSheet], output_path: str = None) -> Dict:
        """Generate comprehensive report of processing results."""
        if not results:
            return {"error": "No results to report"}
        
        # Calculate statistics
        total_sheets = len(results)
        avg_score = np.mean([r.score for r in results])
        avg_confidence = np.mean([np.mean(r.confidence_scores) for r in results if r.confidence_scores])
        
        # Count processing errors
        error_count = sum(len(r.processing_errors) for r in results)
        
        # Score distribution
        score_ranges = {
            "90-100": sum(1 for r in results if 90 <= r.score <= 100),
            "80-89": sum(1 for r in results if 80 <= r.score < 90),
            "70-79": sum(1 for r in results if 70 <= r.score < 80),
            "60-69": sum(1 for r in results if 60 <= r.score < 70),
            "0-59": sum(1 for r in results if 0 <= r.score < 60)
        }
        
        report = {
            "summary": {
                "total_sheets_processed": total_sheets,
                "average_score": round(avg_score, 2),
                "average_confidence": round(avg_confidence, 3),
                "processing_errors": error_count,
                "success_rate": round((total_sheets - error_count) / total_sheets * 100, 2) if total_sheets > 0 else 0
            },
            "score_distribution": score_ranges,
            "detailed_results": [
                {
                    "student_id": r.student_id,
                    "score": round(r.score, 2),
                    "correct_answers": r.correct_answers,
                    "incorrect_answers": r.incorrect_answers,
                    "blank_answers": r.blank_answers,
                    "confidence": round(np.mean(r.confidence_scores), 3) if r.confidence_scores else 0,
                    "errors": r.processing_errors
                }
                for r in results
            ],
            "pipeline_accuracy": {
                "mark_detection_accuracy": self._calculate_mark_detection_accuracy(results),
                "grid_detection_success_rate": self._calculate_grid_detection_success(results)
            }
        }
        
        # Save report if output path provided
        if output_path:
            with open(output_path, 'w') as f:
                json.dump(report, f, indent=2)
            print(f"Report saved to: {output_path}")
        
        return report
    
    def _calculate_mark_detection_accuracy(self, results: List[AnswerSheet]) -> float:
        """Calculate mark detection accuracy based on confidence scores."""
        all_confidences = []
        for result in results:
            all_confidences.extend(result.confidence_scores)
        
        if not all_confidences:
            return 0.0
        
        # High confidence (>0.7) indicates good detection
        high_confidence_count = sum(1 for c in all_confidences if c > 0.7)
        return round(high_confidence_count / len(all_confidences) * 100, 2)
    
    def _calculate_grid_detection_success(self, results: List[AnswerSheet]) -> float:
        """Calculate grid detection success rate."""
        successful_detections = sum(1 for r in results if not r.processing_errors)
        return round(successful_detections / len(results) * 100, 2) if results else 0.0
    
    def visualize_results(self, results: List[AnswerSheet], output_dir: str = "omr_results"):
        """Create visualizations of the processing results."""
        os.makedirs(output_dir, exist_ok=True)
        
        # Score distribution histogram
        scores = [r.score for r in results]
        plt.figure(figsize=(10, 6))
        plt.hist(scores, bins=20, alpha=0.7, color='skyblue', edgecolor='black')
        plt.xlabel('Score')
        plt.ylabel('Number of Students')
        plt.title('Score Distribution')
        plt.grid(True, alpha=0.3)
        plt.savefig(os.path.join(output_dir, 'score_distribution.png'))
        plt.close()
        
        # Confidence distribution
        all_confidences = []
        for r in results:
            all_confidences.extend(r.confidence_scores)
        
        if all_confidences:
            plt.figure(figsize=(10, 6))
            plt.hist(all_confidences, bins=20, alpha=0.7, color='lightgreen', edgecolor='black')
            plt.xlabel('Confidence Score')
            plt.ylabel('Frequency')
            plt.title('Mark Detection Confidence Distribution')
            plt.grid(True, alpha=0.3)
            plt.savefig(os.path.join(output_dir, 'confidence_distribution.png'))
            plt.close()
        
        print(f"Visualizations saved to: {output_dir}")


def main():
    """Main function for command-line usage."""
    parser = argparse.ArgumentParser(description='OMR System for Answer Sheet Processing')
    parser.add_argument('--input', '-i', required=True, help='Input image or directory path')
    parser.add_argument('--answer-key', '-k', required=True, help='Path to answer key JSON file')
    parser.add_argument('--output', '-o', help='Output directory for results')
    parser.add_argument('--config', '-c', help='Path to configuration JSON file')
    
    args = parser.parse_args()
    
    # Load configuration
    config = OMRConfig()
    if args.config and os.path.exists(args.config):
        with open(args.config, 'r') as f:
            config_data = json.load(f)
            for key, value in config_data.items():
                if hasattr(config, key):
                    setattr(config, key, value)
    
    # Initialize OMR processor
    omr = OMRProcessor(config)
    
    # Load answer key
    omr.load_answer_key(args.answer_key)
    
    # Process input
    if os.path.isfile(args.input):
        # Single image
        result = omr.process_answer_sheet(args.input)
        results = [result]
    else:
        # Directory of images
        results = omr.process_batch(args.input)
    
    # Generate report
    output_dir = args.output or "omr_results"
    os.makedirs(output_dir, exist_ok=True)
    
    report = omr.generate_report(results, os.path.join(output_dir, 'report.json'))
    
    # Create visualizations
    omr.visualize_results(results, output_dir)
    
    # Print summary
    print("\n" + "="*50)
    print("OMR PROCESSING COMPLETE")
    print("="*50)
    print(f"Total sheets processed: {report['summary']['total_sheets_processed']}")
    print(f"Average score: {report['summary']['average_score']}%")
    print(f"Success rate: {report['summary']['success_rate']}%")
    print(f"Results saved to: {output_dir}")
    print("="*50)


if __name__ == "__main__":
    main()
