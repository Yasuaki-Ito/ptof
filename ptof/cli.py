"""
PtoF - PPTX to Figures - CLI Interface

Module providing command-line interface.
"""

import argparse
import glob
import os
import sys

from .core import process_pptx, parse_color


def main():
    parser = argparse.ArgumentParser(
        prog='ptof',
        description='PtoF - Extract figures (PDF/PNG/SVG) from PPTX slides for academic papers'
    )
    parser.add_argument(
        'input',
        nargs='+',
        help='Input PPTX file path(s) (multiple files and wildcards supported)'
    )
    parser.add_argument(
        '-o', '--output',
        default='output_dir',
        help='Output directory path (default: output_dir)'
    )
    parser.add_argument(
        '--embed-fonts',
        action='store_true',
        help='Force font embedding (PDF/A format)'
    )
    parser.add_argument(
        '-c', '--color',
        default='cyan',
        help='Marker rectangle color (name: cyan, red, ... or HEX: #FF0000)'
    )
    parser.add_argument(
        '--dpi',
        type=int,
        default=300,
        help='Resolution for PNG output (default: 300)'
    )
    parser.add_argument(
        '--margin',
        type=float,
        default=0,
        help='Margin in points (positive: expand, negative: shrink)'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Show detected regions without converting'
    )
    parser.add_argument(
        '-n', '--no-overwrite',
        action='store_true',
        help='Confirm before overwriting existing files'
    )
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress output'
    )

    args = parser.parse_args()

    # Expand input files (wildcard support)
    input_files = []
    for pattern in args.input:
        # Expand wildcards
        if '*' in pattern or '?' in pattern:
            expanded = glob.glob(pattern)
            if not expanded:
                print(f"Warning: No files matched pattern: {pattern}")
            input_files.extend(expanded)
        else:
            input_files.append(pattern)

    # Remove duplicates and check existence
    input_files = list(dict.fromkeys(input_files))  # Remove duplicates while preserving order
    missing_files = [f for f in input_files if not os.path.exists(f)]

    if missing_files:
        for f in missing_files:
            print(f"Error: Input file not found: {f}")
        sys.exit(1)

    if not input_files:
        print("Error: No input files specified")
        sys.exit(1)

    try:
        marker_color = parse_color(args.color)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    all_output_files = []

    for pptx_file in input_files:
        try:
            output_files = process_pptx(
                pptx_file,
                args.output,
                args.embed_fonts,
                marker_color,
                args.dpi,
                args.margin,
                args.dry_run,
                args.quiet,
                args.no_overwrite
            )
            all_output_files.extend(output_files)

        except Exception as e:
            print(f"Error processing {pptx_file}: {e}")
            if len(input_files) == 1:
                sys.exit(1)

    if not args.quiet:
        if all_output_files:
            if args.dry_run:
                print(f"\n[Dry-run] Would create {len(all_output_files)} file(s)")
            else:
                print(f"\nSuccessfully created {len(all_output_files)} file(s):")
                for f in all_output_files:
                    print(f"  - {f}")
        else:
            print("No files were created")


if __name__ == '__main__':
    main()
