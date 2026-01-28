#!/usr/bin/env python
"""Entry point script to run Liturgie Samensteller."""

import sys
import os

# Add src to path so imports work
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.main import main

if __name__ == "__main__":
    main()
