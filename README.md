# json-to-excel-advanced

Convert complex nested JSON (including MongoDB exports) to Excel with multiple sheets, statistics, and full data normalization.

## ✨ Features

- **🎯 100% Data Preservation** - All fields from JSON are extracted and included
- **📊 Multiple Sheets** - Organized data across 9+ Excel tabs
- **🗂️ Smart Normalization** - Handles nested objects, arrays, and MongoDB types
- **📈 Auto Statistics** - Generates aggregated stats by different dimensions
- **🔄 Multiple Formats** - Wide format (human-readable) + Long format (analysis-ready)
- **💾 Raw Data Backup** - Complete JSON preserved in dedicated sheet
- **🌐 MongoDB Support** - Handles `$oid`, `$date`, and other MongoDB types
- **⚡ Fast Processing** - Efficiently handles large JSON files (tested with 500KB+)

## 📦 What Gets Generated

When you convert a JSON file, you get an Excel workbook with:

### Core Data Sheets
1. **Main Data** - All root-level fields + expanded audit data with readable names
2. **Detailed Items** - Nested items (e.g., questions, orders) in separate rows
3. **Raw JSON** - Complete original JSON for reference

### Statistics Sheets
4. **Stats by Category 1** - Aggregated statistics
5. **Stats by Category 2** - Additional breakdowns
6. **Stats by Category 3** - More analysis dimensions
7. **Response Analysis** - Distribution analysis

### Advanced Normalization
8. **Normalized Data** - Completely flattened JSON using `json_normalize()`
9. **Expanded Items** - Items with arrays expanded into separate rows (long format)

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/matheusnorjosa/json-to-excel-advanced.git
cd json-to-excel-advanced

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage

```bash
# Convert a JSON file
python json_to_excel.py input.json

# Specify output file
python json_to_excel.py input.json -o output.xlsx

# Custom configuration
python json_to_excel.py input.json --config config.json
```

### Python API

```python
from json_to_excel import JSONToExcelConverter

# Basic conversion
converter = JSONToExcelConverter('input.json')
converter.convert('output.xlsx')

# Advanced options
converter = JSONToExcelConverter(
    'input.json',
    nested_items_key='questoes',  # Key containing nested items
    audit_key='auditoria',        # Key containing readable names
    stats_dimensions=['municipio', 'turma']  # Dimensions for stats
)
converter.convert('output.xlsx')
```

## 📋 Requirements

- Python 3.7+
- pandas >= 2.0.0
- openpyxl >= 3.0.0

### Type Safety

This project includes full type hints (PEP 484) for better code quality and IDE support:

```bash
# Optional: Install mypy for type checking
pip install mypy pandas-stubs

# Run type checker
mypy json_to_excel.py
```

## 🎯 Use Cases

Perfect for converting:

- **MongoDB Exports** - Handles `$oid`, `$date`, nested documents
- **API Responses** - Complex nested JSON from REST APIs
- **Survey Data** - Questions with multiple responses
- **E-commerce Data** - Orders with line items
- **Analytics Data** - Events with nested properties
- **Any Complex JSON** - Deeply nested structures

## 🛠️ Advanced Features

### Custom Field Mapping

Create a `config.json` to customize field extraction:

```json
{
  "id_fields": ["_id", "student", "school"],
  "date_fields": ["createdAt", "updatedAt"],
  "nested_items": "questions",
  "audit_section": "metadata",
  "readable_name_path": ["metadata", "school", "name"]
}
```

### Handling Large Files

For very large JSON files:

```python
converter = JSONToExcelConverter('huge.json', chunk_size=1000)
converter.convert('output.xlsx', memory_efficient=True)
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [pandas](https://pandas.pydata.org/) - Powerful data analysis library
- Excel export via [openpyxl](https://openpyxl.readthedocs.io/) - Excel file handling
- Inspired by real-world data conversion challenges

## 🔖 Version

Current version: 1.0.0

## 📊 Stats

- Tested with files up to 1MB+
- Handles 1000+ records efficiently
- Supports unlimited nesting levels
- Processes 100+ columns automatically

---

Made by Matheus Norjosa
