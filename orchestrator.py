"""
ETF Data Extraction and Mapping Orchestrator

This script orchestrates the complete ETF data pipeline:
1. Runs main.py to extract ETF Effective Duration data from iShares websites
2. Runs mapping_code.py to map the extracted data to ISHARESIDD_DATA_.xlsx

The orchestrator ensures proper execution order and error handling.
"""

import subprocess
import sys
import os
from datetime import datetime

def print_banner(message):
    """Print a formatted banner message"""
    print("\n" + "=" * 80)
    print(message)
    print("=" * 80)

def run_script(script_name, description):
    """
    Run a Python script and handle errors
    Returns True if successful, False otherwise
    """
    print_banner(f"EXECUTING: {description}")
    print(f"Script: {script_name}")
    print(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()

    try:
        # Run the script and capture output
        result = subprocess.run(
            [sys.executable, script_name],
            capture_output=False,  # Show output in real-time
            text=True,
            check=True
        )

        print()
        print(f"[SUCCESS] {description} completed successfully")
        print(f"End time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        return True

    except subprocess.CalledProcessError as e:
        print()
        print(f"[ERROR] {description} failed with exit code {e.returncode}")
        print(f"End time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        return False
    except FileNotFoundError:
        print()
        print(f"[ERROR] Script not found: {script_name}")
        print(f"Please ensure {script_name} exists in the current directory")
        return False
    except Exception as e:
        print()
        print(f"[ERROR] Unexpected error running {description}: {str(e)}")
        return False

def check_file_exists(filename):
    """Check if a required file exists"""
    if not os.path.exists(filename):
        print(f"[WARNING] File not found: {filename}")
        return False
    return True

def main():
    """Main orchestrator function"""
    start_time = datetime.now()

    print_banner("ETF DATA EXTRACTION AND MAPPING ORCHESTRATOR")
    print(f"Start Time: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Working Directory: {os.getcwd()}")

    # Check if required scripts exist
    print("\nChecking for required scripts...")
    main_py_exists = check_file_exists("main.py")
    mapping_py_exists = check_file_exists("mapping_code.py")

    if not main_py_exists or not mapping_py_exists:
        print("\n[ERROR] Required scripts are missing. Cannot proceed.")
        sys.exit(1)

    print("[OK] All required scripts found")

    # ==========================================================================
    # STEP 1: Run main.py to extract ETF data
    # ==========================================================================
    success_step1 = run_script(
        "main.py",
        "ETF Data Extraction (main.py)"
    )

    if not success_step1:
        print("\n[FATAL] Data extraction failed. Cannot proceed to mapping.")
        print_banner("ORCHESTRATOR FAILED")
        sys.exit(1)

    # Check if extraction produced the expected output file
    if not check_file_exists("ETF_Effective_Duration.xlsx"):
        print("\n[WARNING] ETF_Effective_Duration.xlsx was not created")
        print("[WARNING] Proceeding with mapping anyway...")

    # ==========================================================================
    # STEP 2: Run mapping_code.py to map data to ISHARESIDD_DATA_.xlsx
    # ==========================================================================
    success_step2 = run_script(
        "mapping_code.py",
        "Data Mapping (mapping_code.py)"
    )

    if not success_step2:
        print("\n[ERROR] Data mapping failed")
        print_banner("ORCHESTRATOR COMPLETED WITH ERRORS")
        sys.exit(1)

    # ==========================================================================
    # SUCCESS - Both steps completed
    # ==========================================================================
    end_time = datetime.now()
    duration = end_time - start_time

    print_banner("ORCHESTRATOR COMPLETED SUCCESSFULLY")
    print(f"Start Time: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"End Time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Total Duration: {duration}")
    print()
    print("Pipeline Summary:")
    print("  ✓ Step 1: ETF Data Extraction (main.py) - SUCCESS")
    print("  ✓ Step 2: Data Mapping (mapping_code.py) - SUCCESS")
    print()
    print("Output Files:")
    if check_file_exists("ETF_Effective_Duration.xlsx"):
        print("  ✓ ETF_Effective_Duration.xlsx - Created")
    if check_file_exists("ISHARESIDD_DATA_.xlsx"):
        print("  ✓ ISHARESIDD_DATA_.xlsx - Updated")
    print("=" * 80)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[INTERRUPTED] Orchestrator stopped by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n[FATAL ERROR] Orchestrator crashed: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
