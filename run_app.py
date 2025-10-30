#!/usr/bin/env python3
"""
Simple script to run the Streamlit app for the Monthly Report Scraper.
"""

import subprocess
import sys
import os
from pathlib import Path

def main():
    """Run the Streamlit app."""
    # Get the directory where this script is located
    script_dir = Path(__file__).parent
    app_path = script_dir / "streamlit_app.py"
    
    if not app_path.exists():
        print("❌ Error: streamlit_app.py not found!")
        print(f"Expected location: {app_path}")
        sys.exit(1)
    
    print("🚀 Starting Monthly Report Scraper...")
    print("📱 The app will open in your default web browser")
    print("🔗 If it doesn't open automatically, go to: http://localhost:8501")
    print("⏹️  Press Ctrl+C to stop the app")
    print("-" * 50)
    
    try:
        # Run streamlit
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", 
            str(app_path),
            "--server.port", "8501",
            "--server.address", "localhost"
        ])
    except KeyboardInterrupt:
        print("\n👋 App stopped by user")
    except Exception as e:
        print(f"❌ Error running app: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
