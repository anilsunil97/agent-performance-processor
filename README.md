# 📊 Agent Performance Data Processor

A professional data analysis tool for processing and analyzing agent performance data from CSV files.

🎨 **NEW: Professional Logo & Branding Added!**
🚀 **[Live Demo on Streamlit Cloud](https://agent-performance-proceappr-arjunexpress.streamlit.app/)**
💻 **NEW: Windows Installer Available!**

## ✨ Features

- 📂 Upload CSV files containing agent performance data
- 🧹 Automatic data cleaning and processing
- 📊 Data preview and summary statistics with color-coded performance indicators
- 📥 Download processed data as styled Excel reports
- 🎯 Focus on key performance metrics
- 🎨 Professional logo and attractive user interface
- 💻 Available as web app, native Windows executable, AND professional installer

## 🚀 Installation Options

### Option 1: Professional Windows Installer (NEW! ⭐)
**Perfect for business use and easy deployment**

1. Download `AgentPerformanceProcessor_Setup.exe` from the [Releases](https://github.com/anilsunil97/agent-performance-processor/releases) page
2. Double-click the installer and follow the wizard
3. Choose installation options:
   - ✅ Start Menu shortcuts
   - ✅ Desktop shortcut  
   - ✅ CSV file associations
4. Launch from Start Menu or Desktop
5. Uninstall anytime from Windows Settings

**Benefits:**
- Professional Windows integration
- Start Menu and Desktop shortcuts
- Add/Remove Programs support
- File associations for CSV files
- Clean uninstallation
- No manual setup required

### Option 2: Portable Executable
**For quick use without installation**

1. Download the `AgentPerformanceProcessor_Offline` folder
2. Double-click `AgentPerformanceProcessor_Offline.exe`
3. No installation required - runs immediately

### Option 3: Use Online (No Download)
**For occasional use**

Visit the live app: **[Agent Performance Processor](https://agent-performance-proceappr-arjunexpress.streamlit.app/)**

### Option 4: Run from Source Code
**For developers**

1. Clone this repository:
```bash
git clone https://github.com/anilsunil97/agent-performance-processor.git
cd agent-performance-processor
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
streamlit run streamlit_app.py
```

4. Open your browser and navigate to `http://localhost:8501`

## 📦 Available Versions

| Version | Type | Best For | Features |
|---------|------|----------|----------|
| **Windows Installer** | Professional Setup | Business deployment, permanent installation | Start Menu integration, Desktop shortcuts, Add/Remove Programs, File associations, Clean uninstall |
| **Portable Executable** | Standalone .exe | Quick use, no installation | Runs immediately, no setup required, fully offline |
| **Web Application** | Online service | Occasional use, any device | No download needed, always up-to-date, cross-platform |
| **Source Code** | Python script | Development, customization | Full source access, customizable, requires Python |

### 🎯 Which Version Should I Choose?

- **🏢 For Business/Office Use:** Windows Installer (professional integration)
- **⚡ For Quick Tasks:** Portable Executable (instant use)
- **🌐 For Occasional Use:** Web Application (no download)
- **👨‍💻 For Development:** Source Code (full control)

## How to Use

1. Upload your agent performance CSV file using the file uploader
2. Review the processed data and summary statistics
3. Download the cleaned Excel report

## Data Processing

The application performs the following data processing steps:

- Removes unnecessary columns (CURRENT USER GROUP, MOST RECENT USER GROUP, etc.)
- Calculates total pause time (PAUSE + DEAD + DISPO)
- Sorts agents by total inbound calls
- Formats time columns properly
- Adds a remarks column for additional notes

## Requirements

- Python 3.7+
- pandas
- openpyxl
- streamlit

## File Structure

```
├── streamlit_app.py        # Main Streamlit application (for cloud deployment)
├── agent-performance-app/
│   └── app.py              # Local version
├── requirements.txt        # Python dependencies
├── README.md              # Project documentation
└── .gitignore             # Git ignore file
```

## Deployment

### Deploy to Streamlit Cloud

1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repository
4. Set main file path to `streamlit_app.py`
5. Deploy!

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

This project is open source and available under the MIT License.