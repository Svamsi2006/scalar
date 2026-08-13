import re
import spacy
import subprocess
import sys
from typing import List, Dict, Set, Optional

class DetectionEngine:
    def __init__(self, exclude_types: Optional[Set[str]] = None):
        self.exclude_types = exclude_types or set()
        
        try:
            self.nlp = spacy.load("en_core_web_sm")
        except OSError:
            # This runs silently on Streamlit Cloud to download the model on startup
            subprocess.run([sys.executable, "-m", "spacy", "download", "en_core_web_sm"])
            try:
                self.nlp = spacy.load("en_core_web_sm")
            except Exception as e:
                print(f"Error loading downloaded spaCy model: {e}")
                self.nlp = None
            
        if self.nlp:
            self.nlp.max_length = 2_000_000

        self._patterns: Dict[str, re.Pattern] = {
            'EMAIL': re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'),
            'PHONE': re.compile(r'(?:\+\d{1,3}[\s.-])?(?:\(?\d{2,4}\)?[\s.-]?)?\d{3,4}[\s.-]?\d{3,4}[\s.-]?\d{0,4}'),
            'SSN': re.compile(r'\b\d{3}-\d{2}-\d{4}\b'),
            'CREDIT_CARD': re.compile(r'\b(?:\d{4}[\s-]?){3}\d{1,4}\b'),
            'IP_ADDRESS': re.compile(r'\b(?:(?:25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(?:25[0-5]|2[0-4]\d|[01]?\d\d?)\b'),
            'DATE': re.compile(
                r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b'
                r'|\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b'
                r'|\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December),?\s+\d{4}\b'
                r'|\b\d{1,2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2,4}\b'
                r'|\b\d{4}-\d{2}-\d{2}\b',
                re.IGNORECASE
            )
        }

    def _count_digits(self, text: str) -> int:
        return sum(1 for c in text if c.isdigit())

    def detect_pii(self, text: str) -> List[Dict]:
        detections = []

        # 1. Regex Pattern Matching
        for entity_type, pattern in self._patterns.items():
            for match in pattern.finditer(text):
                value = match.group(0)
                
                # Phone specific validation
                if entity_type == 'PHONE':
                    num_digits = self._count_digits(value)
                    if not (7 <= num_digits <= 15):
                        continue
                    if num_digits == 4 and value.isdigit():
                        continue
                    if re.search(r'[A-Za-z]', value): # Basic heuristic to avoid PAN numbers / alphanumeric false positives
                        continue
                        
                # Credit Card specific validation
                if entity_type == 'CREDIT_CARD':
                    num_digits = self._count_digits(value)
                    if not (13 <= num_digits <= 19):
                        continue

                detections.append({
                    'entity_type': entity_type,
                    'start_char': match.start(),
                    'end_char': match.end(),
                    'value': value,
                    'source': 'regex'
                })

        # 2. spaCy NER
        if self.nlp:
            doc = self.nlp(text)
            for ent in doc.ents:
                entity_type = None
                value = ent.text

                if ent.label_ == 'PERSON':
                    if len(value) >= 2 and len(value.split()) >= 2:
                        entity_type = 'FULL_NAME'
                elif ent.label_ == 'ORG':
                    if len(value) >= 2:
                        entity_type = 'COMPANY'
                elif ent.label_ in ('GPE', 'LOC', 'FAC'):
                    if len(value) >= 5:
                        entity_type = 'ADDRESS'

                if entity_type:
                    detections.append({
                        'entity_type': entity_type,
                        'start_char': ent.start_char,
                        'end_char': ent.end_char,
                        'value': value,
                        'source': 'ner'
                    })

        # 3. Conflict resolution
        resolved = self._resolve_overlaps(detections)

        # 4. Filter excluded types
        if self.exclude_types:
            resolved = [d for d in resolved if d['entity_type'] not in self.exclude_types]

        # 5. Return sorted detections by start char
        return sorted(resolved, key=lambda x: x['start_char'])

    def _resolve_overlaps(self, detections: List[Dict]) -> List[Dict]:
        if not detections:
            return []

        # Sort by start_char ascending
        sorted_detections = sorted(detections, key=lambda x: x['start_char'])
        resolved = []
        
        current = sorted_detections[0]
        
        for nxt in sorted_detections[1:]:
            # Check for overlap
            if current['end_char'] > nxt['start_char']:
                current_len = current['end_char'] - current['start_char']
                nxt_len = nxt['end_char'] - nxt['start_char']
                
                if current_len > nxt_len:
                    pass  # Keep current
                elif nxt_len > current_len:
                    current = nxt  # Keep next
                else:
                    # Same length, prefer regex source
                    if current['source'] != 'regex' and nxt['source'] == 'regex':
                        current = nxt
            else:
                # No overlap, safe to add current and update current to nxt
                resolved.append(current)
                current = nxt
                
        resolved.append(current)
        return resolved
