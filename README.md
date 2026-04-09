# Enterprise Data Management Suite

**Enterprise Data Management Suite** is a high-performance Python application designed for processing, matching, and cleaning large-scale Excel and CSV datasets. Optimized for stability and speed, it can handle over **900,000+ rows** with ease.

## [GO] Key Features

- **High-Speed Processing**: Uses vectorized operations (`pandas` + `numpy`) to process large datasets in seconds.
- **Ultra-Large Data Mode**: "Direct Save" bypasses slow Excel automation to output massive results directly to CSV/XLSX.
- **Multi-Condition Filtering**: Apply complex inclusive/exclusive filtering rules across multiple columns.
- **Smart Matching**: Automatically detects relationship keys and matches columns between data sources.
- **Batch Processing**: Automated folder-based batch processing with result merging capabilities.
- **Data Quality Tools**: Built-in deduplication and text standardizing tools.
- **Standalone Portable**: Can be compiled into a single `.exe` (Windows) or `.app` (Mac).

## 📊 Performance Benchmarks (Approx.)

- **900,000 Rows Matching**: < 10 seconds.
- **Direct Export to CSV**: < 30 seconds.
- **Active Excel Injection**: Optimized via chunking (5,000 rows/chunk).

---

## 🛠️ Installation & Usage

### 1. Requirements
Ensure you have Python 3.9+ installed.
```bash
python3 -m pip install -r requirements.txt
```

### 2. Run Application
```bash
python3 app.py
```

### 3. Build Executable
- **Windows**: Right-click `build_exe.ps1` -> Run with PowerShell.
- **Mac**: Run `./build_app_mac.sh` in terminal.

---

## [OPEN] Project Structure

- `app.py`: Main entry point.
- `ui/`: UI components and tab definitions.
- `utils/`: Core data processing logic and Excel I/O.
- `presets.json`: User-defined extraction rules.
- `data_suite_debug.log`: Automated error logging for production stability.

## 📄 License
This project is licensed under the MIT License.
