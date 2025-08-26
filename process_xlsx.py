#!/usr/bin/env python3
"""
Simple command-line interface for the XLSX processor utility.

Usage:
    python process_xlsx.py input_file.xlsx output_file.xlsx
    
Example:
    python process_xlsx.py "Client Stat.xlsx" "processed_output.xlsx"
"""

import sys
import argparse
from datetime import datetime
from xlsx_processor import XLSXProcessor


def main():
    parser = argparse.ArgumentParser(description='Process XLSX files and generate statistics')
    parser.add_argument('input_file', help='Input XLSX file path')
    parser.add_argument('output_file', help='Output XLSX file path')
    parser.add_argument('--start-date', 
                       help='Start date for calculations (YYYY-MM-DD format)',
                       default='2025-02-01')
    parser.add_argument('--months', 
                       type=int, 
                       help='Number of months to calculate',
                       default=6)
    parser.add_argument('--srt-sheet', 
                       help='Name of the SRT sheet',
                       default='SRT')
    parser.add_argument('--stat-sheet', 
                       help='Name of the STAT sheet',
                       default='STAT')
    
    args = parser.parse_args()
    
    try:
        # Parse start date
        start_date = datetime.strptime(args.start_date, '%Y-%m-%d')
        
        # Initialize processor
        processor = XLSXProcessor()
        
        print(f"Reading data from: {args.input_file}")
        processor.read_srt_data(args.input_file, args.srt_sheet)
        
        print(f"Generating statistics for {args.months} months starting from {args.start_date}")
        processor.generate_stat_data(start_date, args.months)
        
        print(f"Writing output to: {args.output_file}")
        processor.write_output_file(args.output_file, args.srt_sheet, args.stat_sheet)
        
        print("Processing completed successfully!")
        
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
