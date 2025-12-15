"""
Unit tests for image replacement functionality.

Tests the replace_image operation that allows replacing images in a document
by name, alt text, or index.
"""
import pytest
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

from src.document_processor import DocxBulkUpdater
from src.config import load_operations_from_json


class TestImageReplacement:
    """Test cases for image replacement functionality."""

    def test_replace_first_image_by_default(self, tmp_path):
        """Test replacing the first image in a document (default behavior)."""
        # Use an actual test document with images
        source_doc = Path("profile_test_templates/FTIR_Method_19_20240912.docx")
        if not source_doc.exists():
            pytest.skip("Test document not found")

        # Copy to temporary location for testing
        test_file = tmp_path / "test_doc.docx"
        import shutil
        shutil.copy2(source_doc, test_file)

        # Get original image info
        original_doc = Document(test_file)
        original_image_part = None
        for para in original_doc.paragraphs:
            for run in para.runs:
                drawings = run._element.findall(qn('w:drawing'))
                if drawings:
                    drawing = drawings[0]
                    blip = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                    if blip is not None:
                        rel_id = blip.get(qn('r:embed'))
                        original_image_part = original_doc.part.related_parts[rel_id]
                        break
            if original_image_part:
                break

        assert original_image_part is not None, "No image found in test document"
        original_size = len(original_image_part._blob)

        # Configure replacement - replace first image (default)
        replacement_image = Path("profile_test_templates/Alliance-logo_LG_cropped.png")
        if not replacement_image.exists():
            pytest.skip("Replacement image not found")

        operations = [{
            'op': 'replace_image',
            'image_path': str(replacement_image)
        }]

        # Apply replacement
        processor = DocxBulkUpdater(operations)
        result = processor.modify_docx(test_file)

        # Verify replacement was applied
        assert result is True

        # Check that the image data changed
        modified_doc = Document(test_file)
        modified_image_part = None
        for para in modified_doc.paragraphs:
            for run in para.runs:
                drawings = run._element.findall(qn('w:drawing'))
                if drawings:
                    drawing = drawings[0]
                    blip = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                    if blip is not None:
                        rel_id = blip.get(qn('r:embed'))
                        modified_image_part = modified_doc.part.related_parts[rel_id]
                        break
            if modified_image_part:
                break

        assert modified_image_part is not None
        modified_size = len(modified_image_part._blob)

        # Image sizes should be different
        assert modified_size != original_size

        # Modified image size should match replacement image size
        replacement_size = replacement_image.stat().st_size
        assert modified_size == replacement_size

    def test_replace_image_by_name(self, tmp_path):
        """Test replacing an image by its name attribute."""
        source_doc = Path("profile_test_templates/FTIR_Method_19_20240912.docx")
        if not source_doc.exists():
            pytest.skip("Test document not found")

        test_file = tmp_path / "test_doc.docx"
        import shutil
        shutil.copy2(source_doc, test_file)

        replacement_image = Path("profile_test_templates/Alliance-logo_LG_cropped.png")
        if not replacement_image.exists():
            pytest.skip("Replacement image not found")

        # Replace by name "Picture 2"
        operations = [{
            'op': 'replace_image',
            'image_path': str(replacement_image),
            'name': 'Picture 2'
        }]

        processor = DocxBulkUpdater(operations)
        result = processor.modify_docx(test_file)

        assert result is True

    def test_replace_image_by_index(self, tmp_path):
        """Test replacing an image by its index."""
        source_doc = Path("profile_test_templates/FTIR_Method_19_20240912.docx")
        if not source_doc.exists():
            pytest.skip("Test document not found")

        test_file = tmp_path / "test_doc.docx"
        import shutil
        shutil.copy2(source_doc, test_file)

        replacement_image = Path("profile_test_templates/Alliance-logo_LG_cropped.png")
        if not replacement_image.exists():
            pytest.skip("Replacement image not found")

        # Replace image at index 0 (first image)
        operations = [{
            'op': 'replace_image',
            'image_path': str(replacement_image),
            'index': 0
        }]

        processor = DocxBulkUpdater(operations)
        result = processor.modify_docx(test_file)

        assert result is True

    def test_replace_image_by_alt_text(self, tmp_path):
        """Test replacing an image by its alt text."""
        source_doc = Path("profile_test_templates/FTIR_Method_19_20240912.docx")
        if not source_doc.exists():
            pytest.skip("Test document not found")

        test_file = tmp_path / "test_doc.docx"
        import shutil
        shutil.copy2(source_doc, test_file)

        replacement_image = Path("profile_test_templates/Alliance-logo_LG_cropped.png")
        if not replacement_image.exists():
            pytest.skip("Replacement image not found")

        # The second image has alt text "R:\Diagrams\SampleTrain.jpg"
        operations = [{
            'op': 'replace_image',
            'image_path': str(replacement_image),
            'alt_text': 'R:\\Diagrams\\SampleTrain.jpg'
        }]

        processor = DocxBulkUpdater(operations)
        result = processor.modify_docx(test_file)

        assert result is True

    def test_replace_nonexistent_image(self, tmp_path):
        """Test handling of replacement when image file doesn't exist."""
        source_doc = Path("profile_test_templates/FTIR_Method_19_20240912.docx")
        if not source_doc.exists():
            pytest.skip("Test document not found")

        test_file = tmp_path / "test_doc.docx"
        import shutil
        shutil.copy2(source_doc, test_file)

        # Try to replace with non-existent image
        operations = [{
            'op': 'replace_image',
            'image_path': 'nonexistent_image.png'
        }]

        processor = DocxBulkUpdater(operations)
        doc = Document(test_file)
        result = processor.replace_image(doc, operations[0])

        assert result is False

    def test_replace_image_with_invalid_name(self, tmp_path):
        """Test handling when specified image name doesn't exist."""
        source_doc = Path("profile_test_templates/FTIR_Method_19_20240912.docx")
        if not source_doc.exists():
            pytest.skip("Test document not found")

        test_file = tmp_path / "test_doc.docx"
        import shutil
        shutil.copy2(source_doc, test_file)

        replacement_image = Path("profile_test_templates/Alliance-logo_LG_cropped.png")
        if not replacement_image.exists():
            pytest.skip("Replacement image not found")

        # Try to replace image with non-existent name
        operations = [{
            'op': 'replace_image',
            'image_path': str(replacement_image),
            'name': 'NonExistentPicture'
        }]

        processor = DocxBulkUpdater(operations)
        doc = Document(test_file)
        result = processor.replace_image(doc, operations[0])

        assert result is False

    def test_replace_image_with_invalid_index(self, tmp_path):
        """Test handling when specified index is out of range."""
        source_doc = Path("profile_test_templates/FTIR_Method_19_20240912.docx")
        if not source_doc.exists():
            pytest.skip("Test document not found")

        test_file = tmp_path / "test_doc.docx"
        import shutil
        shutil.copy2(source_doc, test_file)

        replacement_image = Path("profile_test_templates/Alliance-logo_LG_cropped.png")
        if not replacement_image.exists():
            pytest.skip("Replacement image not found")

        # Try to replace image at invalid index
        operations = [{
            'op': 'replace_image',
            'image_path': str(replacement_image),
            'index': 99
        }]

        processor = DocxBulkUpdater(operations)
        doc = Document(test_file)
        result = processor.replace_image(doc, operations[0])

        assert result is False

    def test_config_integration(self, tmp_path):
        """Test image replacement through config file integration."""
        source_doc = Path("profile_test_templates/FTIR_Method_19_20240912.docx")
        if not source_doc.exists():
            pytest.skip("Test document not found")

        test_file = tmp_path / "test_doc.docx"
        import shutil
        shutil.copy2(source_doc, test_file)

        replacement_image = Path("profile_test_templates/Alliance-logo_LG_cropped.png")
        if not replacement_image.exists():
            pytest.skip("Replacement image not found")

        # Use absolute path in config to avoid path resolution issues
        absolute_image_path = replacement_image.absolute()

        # Create config file
        config_content = f"""[
  {{
    "op": "replace_image",
    "image_path": "{str(absolute_image_path).replace(chr(92), chr(92)+chr(92))}",
    "name": "Picture 2"
  }}
]"""

        config_file = tmp_path / "test_config.json"
        config_file.write_text(config_content)

        # Load config and apply replacement
        operations = load_operations_from_json(config_file)
        processor = DocxBulkUpdater(operations)
        result = processor.modify_docx(test_file)

        assert result is True
