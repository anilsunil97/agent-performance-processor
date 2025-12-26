# Agent Performance Data Processor

A Streamlit web application for processing and analyzing agent performance data from CSV files.

🚀 **[Live Demo on Streamlit Cloud](https://your-app-url.streamlit.app)**

## Features

- 📂 Upload CSV files containing agent performance data
- 🧹 Automatic data cleaning and processing
- 📊 Data preview and summary statistics
- 📥 Download processed data as Excel reports
- 🎯 Focus on key performance metrics

## Quick Start

### Option 1: Use Online (Recommended)
Visit the live app: **[Agent Performance Processor](https://your-app-url.streamlit.app)**

### Option 2: Run Locally

1. Clone this repository:
```bash
git clone https://github.com/YOUR_USERNAME/agent-performance-processor.git
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