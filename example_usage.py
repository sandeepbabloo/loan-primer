#!/usr/bin/env python3
"""
Example usage of the XLSX Processor utility.

This script demonstrates how to use the XLSXProcessor class programmatically.
"""

from datetime import datetime
from xlsx_processor import XLSXProcessor


def main():
    """Demonstrate the usage of XLSXProcessor."""
    
    # Initialize the processor
    processor = XLSXProcessor()
    
    # Input and output file paths
    input_file = "Client Stat.xlsx"
    output_file = "example_output.xlsx"
    
    try:
        print("🔄 Starting XLSX processing example...")
        
        # Step 1: Read the SRT data
        print(f"📖 Reading data from {input_file}...")
        srt_data = processor.read_srt_data(input_file)
        print(f"✅ Loaded {len(srt_data)} rows from SRT sheet")
        
        # Display some basic info about the data
        print("\n📊 Data Summary:")
        print(f"   Date range: {srt_data['Date'].min()} to {srt_data['Date'].max()}")
        print(f"   Groups found: {sorted(srt_data['GRP'].unique())}")
        print(f"   Total transactions: {len(srt_data)}")
        
        # Step 2: Generate statistics
        print(f"\n🧮 Generating statistics...")
        start_date = datetime(2025, 2, 1)
        num_months = 6
        
        stat_data = processor.generate_stat_data(start_date, num_months)
        print(f"✅ Generated statistics for {num_months} months starting from {start_date.strftime('%Y-%m-%d')}")
        
        # Display some calculated values
        print("\n📈 Sample Calculations:")
        if len(stat_data) > 2:
            # Get the first month's calculations
            first_month_data = stat_data.iloc[2]  # Skip header rows
            print(f"   EOD Balance (Month 1): {first_month_data['M1']:.2f}")
            
            if len(stat_data) > 3:
                credit_data = stat_data.iloc[3]
                print(f"   Credit BT (Month 1): {credit_data['M1']:.2f}")
        
        # Step 3: Write output file
        print(f"\n💾 Writing output to {output_file}...")
        processor.write_output_file(output_file)
        print(f"✅ Output file created successfully!")
        
        # Step 4: Verify output
        print(f"\n🔍 Verifying output file...")
        import pandas as pd
        
        # Read back the generated file to verify
        output_srt = pd.read_excel(output_file, sheet_name='SRT')
        output_stat = pd.read_excel(output_file, sheet_name='STAT')
        
        print(f"   SRT sheet: {len(output_srt)} rows")
        print(f"   STAT sheet: {len(output_stat)} rows")
        
        print("\n🎉 Example completed successfully!")
        print(f"\nFiles created:")
        print(f"  - {output_file} (processed output)")
        
    except Exception as e:
        print(f"❌ Error during processing: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
