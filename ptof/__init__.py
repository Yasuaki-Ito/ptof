"""
PtoF - PPTX to Figures

A tool to extract regions marked with colored rectangles from PPTX slides
and export them as figures (PDF/PNG/SVG) for academic papers.
"""

from .core import (
    process_pptx,
    parse_color,
    COLOR_NAMES,
)

__version__ = 'v20260129'
__all__ = ['process_pptx', 'parse_color', 'COLOR_NAMES']
