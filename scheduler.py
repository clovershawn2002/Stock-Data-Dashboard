import schedule
import time
import subprocess
import os
import sys
from datetime import datetime

# Set the directory where the scripts are located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# List of scripts to run
SCRIPTS = [
    "polygon_api_Google.py",
    "polygon_api_Amazon.py",
    "polygon_api_Apple.py",
    "polygon_api_Tesla.py",
    "polygon_api_Nvidia.py"
]

def run_all_scripts():
    """Run all stock data scripts"""
    print(f"\n{'='*60}")
    print(f"Running scheduled scripts at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}\n")
    
    for script in SCRIPTS:
        script_path = os.path.join(SCRIPT_DIR, script)
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Running: {script}")
        print("-" * 50)
        
        try:
            result = subprocess.run(
                ["python", script_path],
                cwd=SCRIPT_DIR,
                capture_output=True,
                text=True,
                timeout=300  # 5-minute timeout per script
            )
            
            if result.returncode == 0:
                print(f"✓ {script} completed successfully")
                if result.stdout:
                    print(f"Output: {result.stdout[-200:]}")  # Last 200 chars
            else:
                print(f"✗ {script} failed with error:")
                if result.stderr:
                    print(f"Error: {result.stderr[-500:]}")  # Last 500 chars
                    
        except subprocess.TimeoutExpired:
            print(f"✗ {script} timed out (exceeded 5 minutes)")
        except Exception as e:
            print(f"✗ {script} encountered an error: {str(e)}")
    
    print(f"\n{'='*60}")
    print(f"Batch completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}\n")

def main():
    """Main scheduler loop"""
    # Check if --run-once flag is provided
    run_once = "--run-once" in sys.argv
    
    if run_once:
        # Run all scripts once and exit (for Task Scheduler)
        run_all_scripts()
        print("Run-once mode: exiting after batch completion")
        return
    
    # Schedule the job to run daily at 09:00
    schedule.every().day.at("09:00").do(run_all_scripts)
    
    print(f"Scheduler started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("Scripts scheduled to run daily at 09:00 AM")
    print("Press Ctrl+C to stop the scheduler\n")
    
    # Keep the scheduler running
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every 60 seconds
    except KeyboardInterrupt:
        print("\nScheduler stopped by user")

if __name__ == "__main__":
    main()
