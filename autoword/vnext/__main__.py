"""
Main entry point for AutoWord vNext CLI.

This allows running the vNext pipeline with:
python -m autoword.vnext [command] [options]
"""

from .cli import main

if __name__ == "__main__":
    exit(main())