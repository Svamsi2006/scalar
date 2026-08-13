import re
import logging
from datetime import datetime
from typing import Dict, Optional, Tuple, Callable

from faker import Faker

logger = logging.getLogger(__name__)

class ReplacementManager:
    """
    Manages replacements for detected PII entities across documents.
    Supports multiple redaction modes:
      - 'token': Numbered entity tokens preserving relationships without random data (e.g. [NAME_1], [COMPANY_1])
      - 'mask': Standard categorical redaction tags (e.g. [REDACTED: FULL_NAME])
      - 'blackout': Solid visual blackout blocks (e.g. [REDACTED])
      - 'synthetic': Context-preserving synthetic faking via Faker (e.g. realistic names)
    """

    def __init__(self, mode: str = 'token', seed: Optional[int] = 42, locale: str = 'en_US'):
        """
        Initialize the ReplacementManager.
        
        Args:
            mode (str): 'token', 'mask', 'blackout', or 'synthetic'
            seed (Optional[int]): Seed for Faker to ensure reproducible results.
            locale (str): Locale for Faker.
        """
        self.mode = mode.lower() if mode else 'token'
        self.faker = Faker(locale)
        if seed is not None:
            Faker.seed(seed)
            
        self.pii_map: Dict[str, Dict[str, str]] = {}
        self.original_map: Dict[str, Dict[str, str]] = {}
        self._entity_counters: Dict[str, int] = {}
        
        self._generators: Dict[str, Callable[[str], str]] = {
            'FULL_NAME': self._generate_name,
            'EMAIL': self._generate_email,
            'PHONE': self._generate_phone,
            'COMPANY': self._generate_company,
            'ADDRESS': self._generate_address,
            'SSN': self._generate_ssn,
            'CREDIT_CARD': self._generate_credit_card,
            'DATE': self._generate_date,
            'IP_ADDRESS': self._generate_ip,
            'CIN': self._generate_cin,
            'PAN': self._generate_pan,
            'GSTIN': self._generate_gstin,
            'AADHAAR': self._generate_aadhaar,
        }

    def _normalize(self, value: str) -> str:
        """
        Lowercase, strip whitespace, and collapse multiple spaces to a single space.
        """
        if not isinstance(value, str):
            value = str(value)
        value = value.strip().lower()
        return re.sub(r'\s+', ' ', value)

    def get_replacement(self, entity_type: str, real_value: str) -> str:
        """
        Get or generate a replacement for a given PII entity.
        
        Args:
            entity_type (str): The type of entity (e.g., 'FULL_NAME', 'EMAIL').
            real_value (str): The original PII string.
            
        Returns:
            str: The replacement string.
        """
        normalized = self._normalize(real_value)

        if entity_type not in self.pii_map:
            self.pii_map[entity_type] = {}
            self.original_map[entity_type] = {}
            self._entity_counters[entity_type] = 0

        # Consistent mapping: if already seen, return exact same replacement
        if normalized in self.pii_map[entity_type]:
            return self.pii_map[entity_type][normalized]

        self._entity_counters[entity_type] += 1
        counter = self._entity_counters[entity_type]

        # Generate replacement according to selected mode
        if self.mode == 'token':
            # Indexed entity tokens preserving cross-references without random data
            short_type = entity_type.replace("FULL_", "").replace("_ADDRESS", "")
            replacement = f"[{short_type}_{counter}]"
        elif self.mode == 'mask':
            # Enterprise standard categorical redaction tags
            replacement = f"[REDACTED: {entity_type}]"
        elif self.mode == 'blackout':
            # Solid visual blackout tag
            replacement = f"[REDACTED_{entity_type}]"
        elif self.mode == 'synthetic':
            # Context-preserving Faker synthetic values
            if entity_type in self._generators:
                try:
                    replacement = self._generators[entity_type](real_value)
                except Exception as e:
                    logger.error(f"Error generating synthetic replacement for {entity_type}: {e}")
                    replacement = f"[{entity_type}_{counter}]"
            else:
                replacement = f"[{entity_type}_{counter}]"
        else:
            short_type = entity_type.replace("FULL_", "").replace("_ADDRESS", "")
            replacement = f"[{short_type}_{counter}]"

        self.pii_map[entity_type][normalized] = replacement
        self.original_map[entity_type][real_value] = replacement

        return replacement

    def _generate_name(self, original: str) -> str:
        return self.faker.name()

    def _generate_email(self, original: str) -> str:
        return self.faker.email()

    def _generate_phone(self, original: str) -> str:
        if original.startswith('+91'):
            random_digits = "".join([str(self.faker.random_int(0, 9)) for _ in range(10)])
            all_digits = "91" + random_digits
            
            result = []
            digit_idx = 0
            for char in original:
                if char.isdigit():
                    if digit_idx < len(all_digits):
                        result.append(all_digits[digit_idx])
                        digit_idx += 1
                elif char == '+':
                    result.append('+')
                else:
                    result.append(char)
            
            while digit_idx < len(all_digits):
                result.append(all_digits[digit_idx])
                digit_idx += 1
                
            return "".join(result)
        
        return self.faker.phone_number()

    def _generate_company(self, original: str) -> str:
        return self.faker.company()

    def _generate_address(self, original: str) -> str:
        return self.faker.address().replace('\n', ', ')

    def _generate_ssn(self, original: str) -> str:
        return self.faker.ssn()

    def _generate_credit_card(self, original: str) -> str:
        fake_cc = self.faker.credit_card_number()
        if ' ' in original:
            return ' '.join(fake_cc[i:i+4] for i in range(0, len(fake_cc), 4))
        elif '-' in original:
            return '-'.join(fake_cc[i:i+4] for i in range(0, len(fake_cc), 4))
        return fake_cc

    def _generate_date(self, original: str) -> str:
        date_obj = self.faker.date_object()
        
        if re.match(r'^\d{2}/\d{2}/\d{4}$', original):
            return date_obj.strftime("%d/%m/%Y")
        elif re.match(r'^\d{2}-\d{2}-\d{4}$', original):
            return date_obj.strftime("%d-%m-%Y")
        elif re.match(r'^\d{4}-\d{2}-\d{2}$', original):
            return date_obj.strftime("%Y-%m-%d")
        elif re.match(r'^\d{4}/\d{2}/\d{2}$', original):
            return date_obj.strftime("%Y/%m/%d")
        elif re.match(r'^[a-zA-Z]+ \d{1,2},? \d{4}$', original):
            has_comma = ',' in original
            fmt = "%B %d, %Y" if has_comma else "%B %d %Y"
            return date_obj.strftime(fmt)
        elif re.match(r'^\d{1,2}-[a-zA-Z]{3}-\d{4}$', original):
            return date_obj.strftime("%d-%b-%Y")
            
        return self.faker.date()

    def _generate_ip(self, original: str) -> str:
        return self.faker.ipv4()

    def _generate_cin(self, original: str) -> str:
        prefix = self.faker.random_element(["U", "L"])
        rest = self.faker.bothify(text="#####??####???######").upper()
        return f"{prefix}{rest}"

    def _generate_pan(self, original: str) -> str:
        return self.faker.bothify(text="?????####?").upper()

    def _generate_gstin(self, original: str) -> str:
        return self.faker.bothify(text="##?????####?#?#").upper()

    def _generate_aadhaar(self, original: str) -> str:
        return self.faker.bothify(text="#### #### ####")

    def get_report(self) -> Dict[str, Dict[str, str]]:
        """
        Returns the original map containing actual values to their redacted counterparts.
        """
        return self.original_map

    def get_stats(self) -> Dict[str, int]:
        """
        Returns stats about how many unique replacements were made per entity type.
        """
        return {entity_type: len(replacements) for entity_type, replacements in self.original_map.items()}
