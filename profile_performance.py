#!/usr/bin/env python3
"""
Performance profiling script for DOCX bulk updater.
"""
import cProfile
import pstats
import sys
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from document_processor import DocxBulkUpdater
from config import load_replacements_from_json


class DocxProfiler:
    """Performance profiler for DOCX bulk updater operations."""
    
    def __init__(self, config_file: str = "replace.json", test_dir: str = "profile_test_templates"):
        self.config_file = Path(config_file)
        self.test_dir = Path(test_dir)
        self.updater = None
        self.docx_files = []
    
    def _setup_environment(self) -> bool:
        """Initialize profiler environment and validate configuration."""
        if not self.config_file.exists():
            print(f"Error: {self.config_file} not found")
            return False
        
        if not self.test_dir.exists():
            print(f"Error: {self.test_dir} directory not found")
            return False
        
        try:
            replacements = load_replacements_from_json(self.config_file)
            print(f"Loaded {len(replacements)} replacements from config")
            
            self.updater = DocxBulkUpdater(replacements)
            self.docx_files = list(self.test_dir.glob("*.docx"))
            
            if not self.docx_files:
                print(f"Error: No DOCX files found in {self.test_dir}")
                return False
            
            print(f"Found {len(self.docx_files)} DOCX files for profiling")
            return True
            
        except Exception as e:
            print(f"Setup error: {e}")
            return False
    
    def _process_documents(self) -> None:
        """Process all test documents for profiling."""
        for docx_file in self.docx_files:
            print(f"Processing {docx_file.name}...")
            try:
                changes = self.updater.get_document_changes_preview(docx_file)
                change_count = len(changes) if changes else 0
                status = f"Found changes in {change_count} sections" if changes else "No changes needed"
                print(f"  - {status}")
            except Exception as e:
                print(f"  - Error processing {docx_file.name}: {e}")
    
    def _print_section_header(self, title: str, width: int = 50) -> None:
        """Print formatted section header."""
        print(f"\n{'=' * width}")
        print(title)
        print("=" * width)
    
    def _print_subsection_header(self, title: str, width: int = 50) -> None:
        """Print formatted subsection header."""
        print(f"\n{'-' * width}")
        print(title)
        print("-" * width)
    
    def _generate_stats_report(self, profiler: cProfile.Profile) -> None:
        """Generate and print profiling statistics."""
        # Top functions by cumulative time
        print("\nTop 20 functions by cumulative time:")
        stats = pstats.Stats(profiler)
        stats.strip_dirs().sort_stats('cumulative')
        stats.print_stats(20)
        
        # Module-specific analysis
        self._print_subsection_header("Functions from our modules (document_processor, text_replacement, formatting):")
        stats.print_stats('document_processor|text_replacement|formatting')
        
        # Save detailed profile
        output_file = 'profile_results.prof'
        profiler.dump_stats(output_file)
        print(f"\nDetailed profile saved to: {output_file}")
        print(f"To view interactively: python -m pstats {output_file}")
    
    def profile_document_processing(self) -> bool:
        """Profile the document processing with current configuration."""
        if not self._setup_environment():
            return False
        
        self._print_section_header("Starting performance profiling...")
        
        profiler = cProfile.Profile()
        profiler.enable()
        
        self._process_documents()
        
        profiler.disable()
        
        self._print_section_header("Profiling Results")
        self._generate_stats_report(profiler)
        
        return True
    
    def analyze_memory_usage(self, max_files: int = 2) -> bool:
        """Analyze memory usage during document processing."""
        try:
            import tracemalloc
        except ImportError:
            print("tracemalloc not available for memory analysis")
            return False
        
        if not self.updater or not self.docx_files:
            if not self._setup_environment():
                return False
        
        try:
            self._print_section_header("Memory Usage Analysis")
            
            tracemalloc.start()
            
            # Take snapshot before processing
            snapshot_before = tracemalloc.take_snapshot()
            
            # Process limited number of documents for memory analysis
            test_files = self.docx_files[:max_files]
            print(f"Analyzing memory usage with {len(test_files)} test files")
            
            for docx_file in test_files:
                self.updater.get_document_changes_preview(docx_file)
            
            # Take snapshot after processing
            snapshot_after = tracemalloc.take_snapshot()
            
            # Analyze memory differences
            top_stats = snapshot_after.compare_to(snapshot_before, 'lineno')
            
            print(f"\nTop 10 memory allocations during processing:")
            for stat in top_stats[:10]:
                print(stat)
            
            return True
            
        except Exception as e:
            print(f"Memory analysis error: {e}")
            return False


def main():
    """Main profiling execution."""
    print("DOCX Bulk Updater - Performance Profile")
    print("=" * 50)
    
    profiler = DocxProfiler()
    
    # Check environment
    if not Path("replace.json").exists():
        print("Warning: replace.json not found - some tests may fail")
    
    # Run performance profiling
    success = profiler.profile_document_processing()
    
    # Run memory analysis if profiling succeeded
    if success:
        profiler.analyze_memory_usage()
    
    print("\n" + "=" * 50)
    print("Profiling complete!")
    print("=" * 50)


if __name__ == "__main__":
    main()