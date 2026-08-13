import argparse
import logging
from typing import Optional, Set

from pii_redactor.document_parser import DocumentParser, extract_full_text
from pii_redactor.detection_engine import DetectionEngine
from pii_redactor.replacement_manager import ReplacementManager

logger = logging.getLogger(__name__)

class PiiRedactor:
    """Orchestrator class that ties together parsing, detection, and replacement."""

    def __init__(self, exclude_types: Optional[Set[str]] = None, seed: int = 42, mode: str = 'token'):
        self.exclude_types = exclude_types or set()
        self.parser = DocumentParser()
        self.engine = DetectionEngine(exclude_types=self.exclude_types)
        self.manager = ReplacementManager(mode=mode, seed=seed)

    def process(self, input_path: str, output_path: str) -> None:
        """
        Process the input document, redact PII, and save to the output path.
        """
        logger.info(f"Loading document from {input_path}")
        self.parser.load(input_path)
        paragraphs = self.parser.iter_all_paragraphs()
        total_paragraphs = len(paragraphs)
        logger.info(f"Found {total_paragraphs} paragraphs to process.")

        for i, paragraph in enumerate(paragraphs, start=1):
            text = extract_full_text(paragraph)
            if not text.strip():
                continue

            detections = self.engine.detect_pii(text)
            if detections:
                self.parser.process_paragraph(paragraph, detections, self.manager)

            if i % 50 == 0 or i == total_paragraphs:
                logger.info(f"Processing paragraph {i}/{total_paragraphs}...")

        logger.info(f"Saving redacted document to {output_path}")
        self.parser.save(output_path)

    def print_summary(self) -> None:
        """Prints statistics from all components."""
        print("\n--- Redaction Summary ---")
        parser_stats = self.parser.get_stats()
        print("Parser Stats:")
        for k, v in parser_stats.items():
            print(f"  {k}: {v}")

        manager_stats = self.manager.get_stats()
        print("\nReplacement Stats:")
        total_replacements = 0
        for pii_type, count in manager_stats.items():
            print(f"  {pii_type}: {count}")
            total_replacements += count
        print(f"Total Replacements: {total_replacements}")
        print("-------------------------\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="PII Redactor for DOCX files")
    parser.add_argument(
        "-i", "--input", required=True, help="Path to input .docx file"
    )
    parser.add_argument(
        "-o", "--output", required=True, help="Path to output .docx file"
    )
    parser.add_argument(
        "-e",
        "--exclude",
        help="Comma-separated list of PII types to exclude",
        type=str,
        default="",
    )
    parser.add_argument(
        "-s", "--seed", help="Faker seed", type=int, default=42
    )
    parser.add_argument(
        "-m",
        "--mode",
        help="Redaction mode: token, mask, blackout, synthetic",
        type=str,
        default="token",
    )
    parser.add_argument(
        "-v", "--verbose", help="Enable DEBUG logging", action="store_true"
    )

    args = parser.parse_args()

    # Set up logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )

    # Parse excluded types
    exclude_types = set()
    if args.exclude:
        valid_choices = {
            "FULL_NAME",
            "EMAIL",
            "PHONE",
            "COMPANY",
            "ADDRESS",
            "SSN",
            "CREDIT_CARD",
            "DATE",
            "IP_ADDRESS",
        }
        types = [t.strip().upper() for t in args.exclude.split(",")]
        for t in types:
            if t in valid_choices:
                exclude_types.add(t)
            else:
                logger.warning(f"Unknown PII type to exclude: {t}")

    # Initialize and run redactor
    redactor = PiiRedactor(exclude_types=exclude_types, seed=args.seed, mode=args.mode)
    redactor.process(args.input, args.output)
    redactor.print_summary()
    
    print("Redaction complete!")


if __name__ == "__main__":
    main()
