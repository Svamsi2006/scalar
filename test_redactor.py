import re
from typing import Tuple, List, Dict, Any
from pii_redactor.detection_engine import DetectionEngine
from pii_redactor.replacement_manager import ReplacementManager

def create_test_data() -> Tuple[str, List[Dict[str, str]]]:
    test_text = """
The annual report for Tata Consultancy Services was prepared by Rajesh Kumar Sharma (rajesh.sharma@company.com) on 15/03/1990.
Priya Anand Mehta, who can be reached at priya.mehta@gmail.com or +91 98765 43210, presented the financial summary for Infosys Limited.
Our European partner, David James Wilson, provided his input via david.wilson@outlook.co.uk on March 15, 2024.
During the review meeting, Anita Rao from Goldman Sachs Group called in from +91 20 4505 3237.
Michael Chen, managing director of TechCorp Solutions Pvt Ltd, submitted the final approval from his office at 201, Tower 2, Montreal Business Centre, Baner, Pune - 411045.

For technical support, contact support@techcorp.in or call 022-25678901. Our US representative can be reached at +1 555-867-5309.
All documents were sent to the secondary office located at 42 Baker Street, London W1U 3BW, United Kingdom and 1600 Pennsylvania Avenue NW, Washington DC 20500.

The HR department successfully processed the records for employees with identifiers 123-45-6789 and 987-65-4321 on 22-Jan-1985.
Recent payments were completed using cards 4532 1234 5678 9012 and 5425-2334-3010-9903 on 2024-06-15.
System logs indicate access from 192.168.1.1 and 10.0.0.255 during these transactions.
    """

    ground_truth = [
        # FULL_NAME
        {'entity_type': 'FULL_NAME', 'value': 'Rajesh Kumar Sharma'},
        {'entity_type': 'FULL_NAME', 'value': 'Priya Anand Mehta'},
        {'entity_type': 'FULL_NAME', 'value': 'David James Wilson'},
        {'entity_type': 'FULL_NAME', 'value': 'Anita Rao'},
        {'entity_type': 'FULL_NAME', 'value': 'Michael Chen'},
        # EMAIL
        {'entity_type': 'EMAIL', 'value': 'rajesh.sharma@company.com'},
        {'entity_type': 'EMAIL', 'value': 'priya.mehta@gmail.com'},
        {'entity_type': 'EMAIL', 'value': 'david.wilson@outlook.co.uk'},
        {'entity_type': 'EMAIL', 'value': 'support@techcorp.in'},
        # PHONE
        {'entity_type': 'PHONE', 'value': '+91 98765 43210'},
        {'entity_type': 'PHONE', 'value': '+91 20 4505 3237'},
        {'entity_type': 'PHONE', 'value': '+1 555-867-5309'},
        {'entity_type': 'PHONE', 'value': '022-25678901'},
        # COMPANY
        {'entity_type': 'COMPANY', 'value': 'Tata Consultancy Services'},
        {'entity_type': 'COMPANY', 'value': 'Infosys Limited'},
        {'entity_type': 'COMPANY', 'value': 'Goldman Sachs Group'},
        {'entity_type': 'COMPANY', 'value': 'TechCorp Solutions Pvt Ltd'},
        # ADDRESS
        {'entity_type': 'ADDRESS', 'value': '201, Tower 2, Montreal Business Centre, Baner, Pune - 411045'},
        {'entity_type': 'ADDRESS', 'value': '42 Baker Street, London W1U 3BW, United Kingdom'},
        {'entity_type': 'ADDRESS', 'value': '1600 Pennsylvania Avenue NW, Washington DC 20500'},
        # SSN
        {'entity_type': 'SSN', 'value': '123-45-6789'},
        {'entity_type': 'SSN', 'value': '987-65-4321'},
        # CREDIT_CARD
        {'entity_type': 'CREDIT_CARD', 'value': '4532 1234 5678 9012'},
        {'entity_type': 'CREDIT_CARD', 'value': '5425-2334-3010-9903'},
        # DATE
        {'entity_type': 'DATE', 'value': '15/03/1990'},
        {'entity_type': 'DATE', 'value': 'March 15, 2024'},
        {'entity_type': 'DATE', 'value': '22-Jan-1985'},
        {'entity_type': 'DATE', 'value': '2024-06-15'},
        # IP_ADDRESS
        {'entity_type': 'IP_ADDRESS', 'value': '192.168.1.1'},
        {'entity_type': 'IP_ADDRESS', 'value': '10.0.0.255'}
    ]

    return test_text.strip(), ground_truth

def _compute_iou(start1: int, end1: int, start2: int, end2: int) -> float:
    """Compute Intersection-over-Union for two character spans."""
    intersection_start = max(start1, start2)
    intersection_end = min(end1, end2)
    intersection = max(0, intersection_end - intersection_start)
    union = (end1 - start1) + (end2 - start2) - intersection
    return intersection / union if union > 0 else 0.0


def _find_ground_truth_spans(test_text: str, ground_truth: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    """Annotate ground truth items with character offsets found in the test text."""
    annotated = []
    for gt in ground_truth:
        idx = test_text.find(gt['value'])
        entry = dict(gt)
        if idx >= 0:
            entry['start_char'] = idx
            entry['end_char'] = idx + len(gt['value'])
        annotated.append(entry)
    return annotated


def evaluate_detections(detections: List[Dict[str, Any]], ground_truth: List[Dict[str, str]], test_text: str) -> Dict[str, Dict[str, float]]:
    metrics = {}
    categories = set(gt['entity_type'] for gt in ground_truth)
    for det in detections:
        categories.add(det['entity_type'])

    # Annotate ground truth with character spans
    annotated_gt = _find_ground_truth_spans(test_text, ground_truth)

    # Approximation of total words in the text
    total_words = len(test_text.split())

    for category in categories:
        gt_category = [gt for gt in annotated_gt if gt['entity_type'] == category]
        det_category = [det for det in detections if det['entity_type'] == category]

        tp = 0
        matched_dets = set()
        
        for gt in gt_category:
            gt_val = gt['value'].lower()
            matched = False
            for i, det in enumerate(det_category):
                if i in matched_dets:
                    continue
                det_val = det['value'].lower()

                # Primary: span-based IoU matching (IoU > 0.5)
                if 'start_char' in gt and 'start_char' in det:
                    iou = _compute_iou(
                        gt['start_char'], gt['end_char'],
                        det['start_char'], det['end_char']
                    )
                    if iou > 0.5:
                        tp += 1
                        matched_dets.add(i)
                        matched = True
                        break

                # Fallback: substring containment (for ground truth without spans)
                if gt_val in det_val or det_val in gt_val:
                    tp += 1
                    matched_dets.add(i)
                    matched = True
                    break
            
        fp = len(det_category) - tp
        fn = len(gt_category) - tp
        tn = max(0, total_words - tp - fp - fn)

        precision = (tp / (tp + fp) * 100) if (tp + fp) > 0 else 0.0
        recall = (tp / (tp + fn) * 100) if (tp + fn) > 0 else 0.0
        accuracy = ((tp + tn) / (tp + tn + fp + fn) * 100) if (tp + tn + fp + fn) > 0 else 0.0

        metrics[category] = {
            'TP': tp,
            'FP': fp,
            'FN': fn,
            'Precision': precision,
            'Recall': recall,
            'Accuracy': accuracy
        }

    return metrics

def print_results(metrics: Dict[str, Dict[str, float]]):
    header = f"{'PII Category':<18} | {'TP':>4} | {'FP':>4} | {'FN':>4} | {'Precision%':>11} | {'Recall%':>9} | {'Accuracy%':>10}"
    sep = "-" * len(header)
    
    print(sep)
    print(header)
    print(sep)
    
    total_tp = 0
    total_fp = 0
    total_fn = 0
    
    display_names = {
        'FULL_NAME': 'Full Names',
        'EMAIL': 'Emails',
        'PHONE': 'Phone Numbers',
        'COMPANY': 'Companies',
        'ADDRESS': 'Addresses',
        'SSN': 'SSNs',
        'CREDIT_CARD': 'Credit Cards',
        'DATE': 'Dates',
        'IP_ADDRESS': 'IP Addresses',
    }
    
    for category in ['FULL_NAME', 'EMAIL', 'PHONE', 'COMPANY', 'ADDRESS', 'SSN', 'CREDIT_CARD', 'DATE', 'IP_ADDRESS']:
        if category not in metrics:
            continue
        stats = metrics[category]
        total_tp += stats['TP']
        total_fp += stats['FP']
        total_fn += stats['FN']
        
        name = display_names.get(category, category)
        print(f"{name:<18} | {stats['TP']:>4} | {stats['FP']:>4} | {stats['FN']:>4} | {stats['Precision']:>10.2f}% | {stats['Recall']:>8.2f}% | {stats['Accuracy']:>9.2f}%")
    
    # Handle any extra categories not in the predefined list
    for category, stats in sorted(metrics.items()):
        if category in display_names:
            continue
        total_tp += stats['TP']
        total_fp += stats['FP']
        total_fn += stats['FN']
        print(f"{category:<18} | {stats['TP']:>4} | {stats['FP']:>4} | {stats['FN']:>4} | {stats['Precision']:>10.2f}% | {stats['Recall']:>8.2f}% | {stats['Accuracy']:>9.2f}%")
        
    print(sep)
    
    # Calculate overall metrics
    overall_precision = (total_tp / (total_tp + total_fp) * 100) if (total_tp + total_fp) > 0 else 0.0
    overall_recall = (total_tp / (total_tp + total_fn) * 100) if (total_tp + total_fn) > 0 else 0.0
    avg_accuracy = sum(m['Accuracy'] for m in metrics.values()) / len(metrics) if metrics else 0.0
    
    print(f"{'OVERALL':<18} | {total_tp:>4} | {total_fp:>4} | {total_fn:>4} | {overall_precision:>10.2f}% | {overall_recall:>8.2f}% | {avg_accuracy:>9.2f}%")
    print(sep)


def test_replacement_consistency():
    print("\nTesting ReplacementManager consistency...")
    rm = ReplacementManager()
    
    # Test same name returns same fake name
    name1 = "Rajesh Kumar Sharma"
    fake1_first = rm.get_replacement("FULL_NAME", name1)
    fake1_second = rm.get_replacement("FULL_NAME", name1)
    
    print(f"Original: '{name1}'")
    print(f"First request: '{fake1_first}'")
    print(f"Second request: '{fake1_second}'")
    
    if fake1_first == fake1_second:
        print("[PASS] Consistency test passed: Same input returned same replacement.")
    else:
        print("[FAIL] Consistency test failed: Different replacements for same input.")
        
    # Test different names get different fakes
    name2 = "Priya Anand Mehta"
    fake2 = rm.get_replacement("FULL_NAME", name2)
    
    print(f"\nOriginal: '{name2}'")
    print(f"Replacement: '{fake2}'")
    
    if fake1_first != fake2:
        print("[PASS] Uniqueness test passed: Different inputs returned different replacements.")
    else:
        print("[FAIL] Uniqueness test failed: Same replacement for different inputs.")

def main():
    print("Initializing PII Detection Evaluation Suite...")
    test_text, ground_truth = create_test_data()
    
    engine = DetectionEngine()
    
    print("\nRunning detection engine on test data...")
    detections = engine.detect_pii(test_text)
    
    print("\nEvaluating detections...")
    metrics = evaluate_detections(detections, ground_truth, test_text)
    
    print("\nEvaluation Results:")
    print_results(metrics)
    
    test_replacement_consistency()

if __name__ == '__main__':
    main()
