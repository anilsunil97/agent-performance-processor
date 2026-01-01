"""
Agent Performance Data Processor - Executable Launcher
This script launches the Streamlit app for .exe packaging
"""

import sys
import os
import subprocess
import threading
import time
import webbrowser
from pathlib import Path

def find_free_port():
    """Find a free port for Streamlit"""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

def launch_streamlit():
    """Launch Streamlit app"""
    try:
        # Get the directory where the executable is located
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            app_dir = Path(sys.executable).parent
        else:
            # Running as script
            app_dir = Path(__file__).parent
        
        # Change to app directory
        os.chdir(app_dir)
        
        # Find free port
        port = find_free_port()
        
        print(f"🚀 Starting Agent Performance Data Processor...")
        print(f"📊 App will open at: http://localhost:{port}")
        print("⏳ Please wait while the application loads...")
        
        # Launch Streamlit
        cmd = [
            sys.executable, "-m", "streamlit", "run", 
            "streamlit_app.py", 
            "--server.port", str(port),
            "--server.headless", "true",
            "--browser.gatherUsageStats", "false"
        ]
        
        # Start Streamlit in background
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Wait for Streamlit to start
        time.sleep(3)
        
        # Open browser
        webbrowser.open(f"http://localhost:{port}")
        
        print("✅ Application started successfully!")
        print("🌐 Your browser should open automatically")
        print("❌ Press Ctrl+C to stop the application")
        
        # Wait for process to complete
        try:
            process.wait()
        except KeyboardInterrupt:
            print("\n🛑 Stopping application...")
            process.terminate()
            
    except Exception as e:
        print(f"❌ Error starting application: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    launch_streamlit()