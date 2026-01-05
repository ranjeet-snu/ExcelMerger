# 📊 Excel Data Matcher & Merger Pro

<div align="center">

![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)
![Status](https://img.shields.io/badge/status-active-success.svg)

**A powerful, user-friendly desktop application for matching and merging Excel data with an intuitive modern UI**

[Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Screenshots](#-screenshots) • [Contributing](#-contributing)

</div>

---

## ✨ Features

### 🎯 Core Functionality
- **Smart Data Matching** - Automatically match rows between two Excel files based on custom column mappings
- **Flexible Column Mapping** - Configure multiple column pairs for precise matching criteria
- **Data Normalization** - Intelligent value normalization handles spaces, punctuation, and case differences
- **Batch Processing** - Process thousands of rows with real-time progress tracking
- **Merge Additional Columns** - Automatically brings in all columns from reference file to primary file

### 🎨 Modern User Interface
- **Beautiful Gradient Buttons** - Modern, rounded buttons with smooth hover effects
- **Color-Coded Sections** - Each functional area has distinct background colors for easy navigation
- **Real-Time Progress Tracking** - Visual progress bars and detailed statistics
- **Activity Logging** - Comprehensive log with timestamps and color-coded message types
- **Responsive Design** - Clean, professional layout that works seamlessly

### 🔧 Technical Features
- **Multiple File Format Support** - Works with `.xlsx` and `.xls` files
- **Error Handling** - Robust error detection and user-friendly error messages
- **Data Validation** - Validates column mappings before processing
- **Export Options** - Save merged results to new Excel files
- **Memory Efficient** - Handles large datasets without performance issues

---

## 🚀 Installation

### Prerequisites

Ensure you have Python 3.7 or higher installed on your system.

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/excel-data-matcher.git
cd excel-data-matcher
```

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

**Dependencies:**
- `pandas` - Data manipulation and analysis
- `openpyxl` - Excel file reading/writing
- `tkinter` - GUI framework (usually pre-installed with Python)

### Step 3: Run the Application

```bash
python excel_matcher.py
```

---

## 📖 Usage

### Quick Start Guide

#### 1️⃣ **Load Your Files**

<div align="center">
<img src="https://img.shields.io/badge/Step_1-Load_Files-3b82f6?style=for-the-badge" alt="Step 1"/>
</div>

- Click **"Browse Files"** under **Primary File** to select your main Excel file
- Click **"Browse Files"** under **Reference File** to select your lookup Excel file
- The application will display the columns from each file

#### 2️⃣ **Configure Column Matching**

<div align="center">
<img src="https://img.shields.io/badge/Step_2-Configure_Matching-f59e0b?style=for-the-badge" alt="Step 2"/>
</div>

- Click **"➕ Add Matching Column"** to create a new column pair
- Select the column from your **Primary File** in the left dropdown
- Select the corresponding column from your **Reference File** in the right dropdown
- Add multiple column pairs for more precise matching
- Use the **"✕"** button to remove individual mappings
- Use **"🗑️ Clear All"** to start over

#### 3️⃣ **Process & Merge**

<div align="center">
<img src="https://img.shields.io/badge/Step_3-Process_Files-10b981?style=for-the-badge" alt="Step 3"/>
</div>

- Click **"🚀 Process & Merge Files"**
- Watch real-time progress in the Progress & Statistics panel
- Review the Activity Log for detailed processing information
- Choose where to save your merged file
- Done! 🎉

---

## 🎨 Screenshots

### Main Interface
```
┌─────────────────────────────────────────────────────────────────┐
│                📊 Excel Data Matcher & Merger Pro                │
├─────────────────────────────────────────────────────────────────┤
│                                                                   │
│  📁 Step 1: Select Excel Files                                   │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ Primary File:     [invoice_data.xlsx]   [Browse Files]   │   │
│  │ Reference File:   [product_catalog.xlsx] [Browse Files]  │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                   │
│  🔗 Step 2: Configure Column Matching                           │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │ [➕ Add Matching Column]  [🗑️ Clear All]                  │   │
│  │                                                            │   │
│  │  Primary Column: [Product ID] ⟷ Ref Column: [SKU]  [✕]  │   │
│  │  Primary Column: [Date]       ⟷ Ref Column: [Date] [✕]  │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                   │
│              [🚀 Process & Merge Files]                          │
└─────────────────────────────────────────────────────────────────┘
```

### Color Scheme

| Section | Color | Purpose |
|---------|-------|---------|
| 📁 File Selection | Light Blue | File loading interface |
| 🔗 Column Mapping | Light Amber | Matching configuration |
| 📋 File Columns | Sky Blue | Column information display |
| 📊 Progress & Statistics | Light Green | Processing status |
| 📋 Activity Log | Light Pink | Detailed logging |

---

## 🎯 Use Cases

### Business Applications
- **Invoice Matching** - Match invoices with purchase orders
- **Inventory Management** - Update product information from supplier catalogs
- **Customer Data** - Merge customer information from multiple sources
- **Sales Reports** - Combine sales data with product details
- **Data Enrichment** - Add missing information to existing datasets

### Data Management
- **Database Synchronization** - Keep multiple Excel databases in sync
- **Data Validation** - Cross-reference data between systems
- **Report Generation** - Create comprehensive reports from multiple sources
- **Data Migration** - Prepare data for system migrations
- **Quality Assurance** - Verify data consistency across files

---

## 🔍 How It Works

### Matching Algorithm

The application uses a sophisticated multi-column matching algorithm:

1. **Value Normalization**
   - Converts all values to lowercase
   - Removes extra whitespace
   - Strips punctuation and special characters
   - Handles date formatting consistently

2. **Multi-Column Matching**
   - Creates boolean conditions for each column pair
   - Uses AND logic to find rows matching ALL specified columns
   - Returns the first matching row from the reference file

3. **Data Merging**
   - Preserves all columns from the primary file
   - Adds non-matched columns from the reference file
   - Fills in matched values from the reference file
   - Maintains original row order from primary file

### Example

**Primary File:**
| Order ID | Product Code | Quantity |
|----------|--------------|----------|
| 1001 | PROD-123 | 5 |
| 1002 | PROD-456 | 3 |

**Reference File:**
| SKU | Description | Price |
|-----|-------------|-------|
| PROD-123 | Widget A | $10.00 |
| PROD-456 | Widget B | $15.00 |

**Column Mapping:**
- Primary: `Product Code` ⟷ Reference: `SKU`

**Result:**
| Order ID | Product Code | Quantity | Description | Price |
|----------|--------------|----------|-------------|-------|
| 1001 | PROD-123 | 5 | Widget A | $10.00 |
| 1002 | PROD-456 | 3 | Widget B | $15.00 |

---

## ⚙️ Configuration

### File Requirements

- **Supported Formats:** `.xlsx`, `.xls`
- **File Size:** No hard limit (memory dependent)
- **Column Names:** Must be in the first row
- **Data Types:** Text, numbers, dates are all supported

### System Requirements

- **OS:** Windows 7+, macOS 10.12+, Linux (with GUI support)
- **Python:** 3.7 or higher
- **RAM:** Minimum 2GB (4GB+ recommended for large files)
- **Disk Space:** 100MB for application and dependencies

---

## 🛠️ Troubleshooting

### Common Issues

**Issue: "No rows matched"**
- ✅ Verify column mappings are correct
- ✅ Check that data formats match between files
- ✅ Ensure there are actually matching records
- ✅ Check for extra spaces or special characters

**Issue: Application won't start**
- ✅ Verify Python version (3.7+)
- ✅ Check all dependencies are installed
- ✅ Try reinstalling dependencies: `pip install -r requirements.txt --force-reinstall`

**Issue: File won't load**
- ✅ Ensure file is a valid Excel format (.xlsx or .xls)
- ✅ Check file isn't open in another program
- ✅ Verify file isn't corrupted

**Issue: Slow performance**
- ✅ Large files take longer to process (this is normal)
- ✅ Close other applications to free up memory
- ✅ Consider splitting very large files

---

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

### Ways to Contribute

1. 🐛 **Report Bugs** - Open an issue describing the bug
2. 💡 **Suggest Features** - Share your ideas for improvements
3. 📖 **Improve Documentation** - Help make the docs better
4. 🔧 **Submit Pull Requests** - Contribute code improvements

### Development Setup

```bash
# Fork and clone the repository
git clone https://github.com/yourusername/excel-data-matcher.git
cd excel-data-matcher

# Create a virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install development dependencies
pip install -r requirements-dev.txt

# Make your changes and test
python excel_matcher.py

# Submit a pull request
```

### Code Style

- Follow PEP 8 guidelines
- Use meaningful variable names
- Add comments for complex logic
- Update documentation for new features

---

## 📋 Requirements File

Create a `requirements.txt` file with:

```
pandas>=1.3.0
openpyxl>=3.0.0
```

---



## 👨‍💻 Author

**Ranjeet Kumar**


## 🙏 Acknowledgments

- Built with Python and Tkinter
- Data processing powered by Pandas
- Excel file handling via openpyxl
- Inspired by the need for simple, effective data matching tools

---

## 📞 Support

Need help? Have questions?

- 📧 Email: ranjeet.jnv41@gmail.com


---



**Made with ❤️ and Python**

[⬆ Back to Top](#-excel-data-matcher--merger-pro)

</div>
