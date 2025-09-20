"""
Simple OMR System Demo - Minimal Dependencies
This version works with just Python standard library and numpy.
"""

import json
import os
import random
from typing import List, Dict, Tuple
from dataclasses import dataclass


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


class SimpleOMRProcessor:
    """Simple OMR processor for demonstration."""
    
    def __init__(self):
        self.answer_key = []
        self.question_count = 0
        
    def load_answer_key(self, answer_key_path: str) -> None:
        """Load answer key from JSON file."""
        try:
            with open(answer_key_path, 'r') as f:
                data = json.load(f)
                self.answer_key = data.get('answers', [])
                self.question_count = len(self.answer_key)
                print(f"✓ Loaded answer key with {self.question_count} questions")
        except Exception as e:
            raise ValueError(f"Error loading answer key: {e}")
    
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
            elif answer.upper() == self.answer_key[i].upper():
                correct += 1
            else:
                incorrect += 1
        
        total_questions = len(self.answer_key)
        score = (correct / total_questions) * 100 if total_questions > 0 else 0
        
        return score, correct, incorrect, blank
    
    def extract_student_id(self, image_path: str) -> str:
        """Extract student ID from filename."""
        filename = os.path.basename(image_path)
        name_without_ext = os.path.splitext(filename)[0]
        return name_without_ext
    
    def process_answer_sheet(self, image_path: str) -> AnswerSheet:
        """Process a single answer sheet - simulated version."""
        errors = []
        
        try:
            student_id = self.extract_student_id(image_path)
            
            # Simulate different answer patterns based on student ID
            if "perfect" in student_id.lower():
                answers = self.answer_key.copy()
                confidence_scores = [0.95] * len(answers)
            elif "good" in student_id.lower():
                answers = self.answer_key.copy()
                if len(answers) > 0:
                    answers[-1] = "X"  # Wrong last answer
                confidence_scores = [0.9] * len(answers)
            elif "average" in student_id.lower():
                answers = self.answer_key.copy()
                if len(answers) > 1:
                    answers[-1] = ""  # Blank last answer
                confidence_scores = [0.8] * len(answers)
            elif "poor" in student_id.lower():
                answers = ["X"] * len(self.answer_key)  # All wrong
                confidence_scores = [0.7] * len(answers)
            else:
                # Random simulation
                answers = []
                confidence_scores = []
                for i, correct_answer in enumerate(self.answer_key):
                    rand = random.random()
                    if rand < 0.8:  # 80% chance of correct
                        answers.append(correct_answer)
                        confidence_scores.append(0.9)
                    elif rand < 0.9:  # 10% chance of wrong
                        wrong_options = [opt for opt in 'ABCD' if opt != correct_answer]
                        answers.append(random.choice(wrong_options))
                        confidence_scores.append(0.6)
                    else:  # 10% chance of blank
                        answers.append('')
                        confidence_scores.append(0.0)
            
            # Calculate score
            score, correct, incorrect, blank = self.calculate_score(answers)
            
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
    
    def process_batch(self, image_directory: str) -> List[AnswerSheet]:
        """Process multiple answer sheets in a directory."""
        results = []
        image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']
        
        for file_path in os.listdir(image_directory):
            if any(file_path.lower().endswith(ext) for ext in image_extensions):
                full_path = os.path.join(image_directory, file_path)
                print(f"  Processing: {file_path}")
                result = self.process_answer_sheet(full_path)
                results.append(result)
        
        return results
    
    def generate_report(self, results: List[AnswerSheet], output_path: str = None) -> Dict:
        """Generate comprehensive report of processing results."""
        if not results:
            return {"error": "No results to report"}
        
        # Calculate statistics
        total_sheets = len(results)
        avg_score = sum(r.score for r in results) / total_sheets
        avg_confidence = sum(sum(r.confidence_scores) / len(r.confidence_scores) for r in results if r.confidence_scores) / total_sheets
        
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
                    "confidence": round(sum(r.confidence_scores) / len(r.confidence_scores), 3) if r.confidence_scores else 0,
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
            print(f"✓ Report saved to: {output_path}")
        
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


def create_demo_data():
    """Create demo data for testing."""
    print("1. Creating demo data...")
    
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
    
    # Create demo directory with mock files
    os.makedirs("demo_sheets", exist_ok=True)
    
    # Create mock answer sheet files
    mock_files = [
        "student_perfect_001.png",
        "student_good_002.png", 
        "student_average_003.png",
        "student_poor_004.png",
        "student_error_005.png"
    ]
    
    for filename in mock_files:
        # Create empty file to simulate image
        with open(os.path.join("demo_sheets", filename), "w") as f:
            f.write("Mock image file")
    
    print("✓ Demo data created successfully!")


def main():
    """Main function for the demo."""
    print("="*60)
    print("OMR SYSTEM - SIMPLE DEMO")
    print("="*60)
    
    try:
        # Create demo data
        create_demo_data()
        
        # Initialize OMR processor
        print("\n2. Initializing OMR processor...")
        omr = SimpleOMRProcessor()
        
        # Load answer key
        print("\n3. Loading answer key...")
        omr.load_answer_key("sample_answer_key.json")
        
        # Process demo sheets
        print("\n4. Processing demo answer sheets...")
        results = omr.process_batch("demo_sheets")
        
        # Generate report
        print("\n5. Generating comprehensive report...")
        os.makedirs("demo_results", exist_ok=True)
        report = omr.generate_report(results, "demo_results/report.json")
        
        # Display results
        print("\n6. Processing Results:")
        print("-" * 40)
        print(f"Total sheets processed: {report['summary']['total_sheets_processed']}")
        print(f"Average score: {report['summary']['average_score']}%")
        print(f"Success rate: {report['summary']['success_rate']}%")
        print(f"Average confidence: {report['summary']['average_confidence']:.3f}")
        print(f"Processing errors: {report['summary']['processing_errors']}")
        
        print("\nScore Distribution:")
        for range_name, count in report['score_distribution'].items():
            print(f"  {range_name}: {count} students")
        
        print("\nPipeline Accuracy:")
        print(f"  Mark detection accuracy: {report['pipeline_accuracy']['mark_detection_accuracy']}%")
        print(f"  Grid detection success rate: {report['pipeline_accuracy']['grid_detection_success_rate']}%")
        
        print("\nDetailed Results:")
        print("-" * 40)
        for result in report['detailed_results']:
            print(f"Student {result['student_id']}: {result['score']}% "
                  f"(Correct: {result['correct_answers']}, "
                  f"Incorrect: {result['incorrect_answers']}, "
                  f"Blank: {result['blank_answers']}, "
                  f"Confidence: {result['confidence']:.3f})")
            if result['errors']:
                print(f"  Errors: {', '.join(result['errors'])}")
        
        print("\n" + "="*60)
        print("DEMO COMPLETE - SUCCESS!")
        print("="*60)
        print("Key Features Demonstrated:")
        print("✓ Answer key loading and validation")
        print("✓ Score calculation with multiple metrics")
        print("✓ Confidence scoring for mark detection")
        print("✓ Comprehensive error handling")
        print("✓ Detailed reporting and statistics")
        print("✓ Batch processing capabilities")
        print("✓ Automated pipeline evaluation")
        
        print("\nFiles generated:")
        print("  - demo_results/report.json: Detailed processing results")
        print("  - sample_answer_key.json: Answer key configuration")
        print("  - demo_sheets/: Sample answer sheet files")
        
        print("\nTo use with real images:")
        print("1. Install full dependencies: pip install opencv-python pandas matplotlib")
        print("2. Use omr_system.py for actual image processing")
        print("3. Process individual sheets: python omr_system.py -i image.png -k answer_key.json")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("Please check the error message above and try again.")


if __name__ == "__main__":
    main()
