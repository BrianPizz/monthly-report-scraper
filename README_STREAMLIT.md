# Monthly Report Scraper - Streamlit Frontend

A web-based interface for processing school monthly report PDFs using Streamlit.

## 🚀 Quick Start

Link to deployment [https://ohio-sped-monthly-report-scraper.streamlit.app/]

## 📱 Features

### 🎯 Core Functionality
- **File Upload**: Drag and drop multiple PDF files
- **OCR Processing**: Advanced text extraction with intelligent corrections
- **Data Extraction**: Automatically extracts school data from PDFs
- **Real-time Processing**: Live progress updates and status tracking

### 📊 Data Management
- **Interactive Table**: View and search extracted data
- **Data Visualization**: Charts showing students and teachers by school
- **Export Options**: Download as Excel (.xlsx) or CSV (.csv)
- **Data Validation**: Built-in error checking and reporting

### ⚙️ Configuration Options
- **Verbose Mode**: Detailed processing information
- **OCR Toggle**: Enable/disable OCR processing
- **File Management**: Handle multiple files simultaneously

## 📋 Supported Data Fields

The application extracts the following information from each PDF:

| Field | Description |
|-------|-------------|
| **School Name** | Automatically detected school name |
| **Students** | Number of students (K-8 and 9-12) |
| **Teachers** | Number of teachers (K-8 and 9-12) |
| **Sub** | Substitute teachers |
| **OSS** | Out-of-school suspensions |
| **EX** | Expulsions |
| **ER** | Emergency removals |
| **MDM** | Manifestation determination meetings |

## 🎨 User Interface

### Main Dashboard
- Clean, modern interface with intuitive navigation
- Real-time processing status and progress bars
- Interactive data tables with search functionality
- Download buttons for Excel and CSV export

### Sidebar Controls
- File upload area with drag-and-drop support
- Processing configuration options
- Navigation menu for different sections

### Data Visualization
- Bar charts for students and teachers by school
- Interactive filtering and search capabilities
- Export-ready data formatting

## 🔧 Technical Details

### Dependencies
- **Streamlit**: Web application framework
- **PyMuPDF**: PDF processing and text extraction
- **Tesseract**: OCR for text recognition
- **Pandas**: Data manipulation and analysis
- **OpenPyXL**: Excel file generation

### File Structure
```
monthly-report-scraper/
├── streamlit_app.py          # Main Streamlit application
├── main.py                   # Core PDF processing logic
├── individual_parser.py      # Individual PDF parser
├── requirements.txt          # Python dependencies
├── run_app.py               # App launcher script
└── README_STREAMLIT.md      # This documentation
```

## 🚨 Troubleshooting

### Common Issues

**1. Tesseract not found**
```bash
# On macOS
brew install tesseract

# On Ubuntu/Debian
sudo apt-get install tesseract-ocr

# On Windows
# Download from: https://github.com/UB-Mannheim/tesseract/wiki
```

**2. Permission errors**
```bash
# Make sure you have write permissions in the directory
chmod +x run_app.py
```

**3. Port already in use**
```bash
# Use a different port
streamlit run streamlit_app.py --server.port 8502
```

### Performance Tips

- **Large Files**: For PDFs > 10MB, processing may take longer
- **Multiple Files**: Process files in batches for better performance
- **OCR Mode**: Disable OCR for faster processing if PDFs have good text quality
- **Verbose Mode**: Enable only when debugging issues

## 📈 Usage Examples

### Basic Workflow
1. Upload PDF files using the sidebar
2. Configure processing options (OCR, verbose mode)
3. Click "Process Files" to extract data
4. Review extracted data in the table
5. Download results as Excel or CSV

### Advanced Features
- Use the search box to filter schools
- Enable verbose mode to see detailed processing steps
- Export data in multiple formats
- View data visualizations for insights

## 🤝 Support

For issues or questions:
1. Check the troubleshooting section above
2. Enable verbose mode for detailed error information
3. Review the console output for specific error messages

## 🔄 Updates

The Streamlit frontend integrates seamlessly with your existing PDF processing logic, providing a modern web interface while maintaining all the advanced OCR correction features from the original application.
