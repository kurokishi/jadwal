#app/core/__init__.py
"""
Core logic package.

This folder contains the main backend modules:
- scheduler
- analyzer
- excel_writer
- validator
- time_parser
- cleaner
"""

# NOTE:
# DO NOT import internal classes here to avoid circular import issues.
# Import classes explicitly where needed:
#   from app.core.scheduler import Scheduler
#   from app.core.analyzer import Analyzer
"""
