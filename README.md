# n8n-nodes-excel

A powerful n8n community node that lets you work with Excel (.xlsx) and CSV (.csv) files as if they were databases. Perform CRUD operations, apply filters, sort data, and seamlessly integrate with AI Agent nodes for natural language processing.

[![NPM Version](https://img.shields.io/npm/v/n8n-nodes-excel.svg)](https://www.npmjs.com/package/n8n-nodes-excel)
[![License](https://img.shields.io/npm/l/n8n-nodes-excel.svg)](https://github.com/dandacompany/n8n-nodes-excel/blob/main/LICENSE)

## ✨ Features

### 📁 File Operations
- **Create**: Generate new Excel or CSV files with custom columns
- **Upload**: Import files from binary data or file paths
- **Download**: Export files as binary data for downstream processing
- **Delete**: Remove files from the data directory

### 📊 Data Operations
- **Read**: Query data with advanced filtering, sorting, and pagination
- **Add Row**: Insert new records with dynamic column mapping
- **Update Rows**: Modify multiple records using filter criteria
- **Delete Rows**: Remove records matching filter conditions
- **Clear Data**: Remove all data while preserving headers

### 🤖 AI Agent Tool Compatibility
- **Native Integration**: Works seamlessly as an AI Agent Tool (`usableAsTool: true`)
- **Natural Language**: AI can interact with spreadsheets using natural language
- **Smart Column Mapping**: Dynamic column detection for intelligent operations

### 🔍 Advanced Filtering (v1.2.4+)
- **Text Filters**: Contains, Not Contains, Equals, Not Equals, Starts/Ends With
- **Numeric Filters**: Greater Than, Less Than, Greater/Equal, Less/Equal
- **Empty Checks**: Is Empty, Is Not Empty
- **Multiple Conditions**: Apply multiple filters simultaneously (AND logic)
- **Case-Insensitive**: Smart text matching ignores case differences

### 🔄 Sorting & Pagination
- **Column Sorting**: Ascending/Descending with automatic type detection
- **Smart Limits**: Use 0 for all rows, or specify exact count
- **Dynamic Columns**: Dropdown selection with all available columns

## 📦 Installation

### Using npm

```bash
# Navigate to your n8n custom nodes folder
cd ~/.n8n/nodes

# Install the package
npm install n8n-nodes-excel
```

### Using Docker Compose

Add to your `docker-compose.yml`:

```yaml
version: '3.7'

services:
  n8n:
    image: n8nio/n8n
    environment:
      - NODE_FUNCTION_ALLOW_EXTERNAL=n8n-nodes-excel
      - N8N_CUSTOM_EXTENSIONS=/home/node/.n8n/custom
    volumes:
      - ~/.n8n/custom:/home/node/.n8n/custom
```

Then install inside the container:

```bash
docker exec -it <n8n-container> /bin/sh
npm install n8n-nodes-excel
```

## 🚀 Usage

### Basic Workflow Example

1. **Read Data with Filters**
   - Select file: `sales_data.xlsx`
   - Add filter: `Status` equals `Active`
   - Add filter: `Amount` greater than `1000`
   - Sort by: `Date` descending
   - Limit: 10 records

2. **Update Multiple Records**
   - Use filters to select rows: `Status` equals `Pending`
   - Update fields: Set `Status` to `Active`, `ProcessedDate` to today
   - All matching rows are updated in one operation

3. **Delete Old Records**
   - Add filter: `Date` less than `2023-01-01`
   - Add filter: `Status` equals `Archived`
   - Removes all matching rows while preserving data structure

4. **Export Results**
   - Download as binary data
   - Pass to other nodes for processing

### AI Agent Integration

The node automatically works with AI Agents:

```text
User: "Show me all active customers from the customers.xlsx file"
AI Agent: [Uses Excel/CSV node to filter and return results]

User: "Add a new customer John Doe to the spreadsheet"
AI Agent: [Automatically maps fields and adds the row]
```

### Filter Examples

```javascript
// Multiple conditions (AND logic)
Filters:
  - Column: "Status" | Operator: "Equals" | Value: "Active"
  - Column: "Amount" | Operator: "Greater Than" | Value: "1000"
  - Column: "City" | Operator: "Contains" | Value: "New"

// Sorting
Sort:
  - Column: "CreatedDate" | Direction: "Descending"
```

## 🗂️ File Storage

All files are stored in the `data` directory:
- Default location: `.n8n/custom-nodes-data/n8n-nodes-excel/data`
- Auto-created if not exists
- Supports both .xlsx and .csv formats

## 🛠️ Development

### Prerequisites
- Node.js (v16+)
- n8n instance for testing

### Building from Source

```bash
# Clone the repository
git clone https://github.com/dandacompany/n8n-nodes-excel.git
cd n8n-nodes-excel

# Install dependencies
npm install

# Build the node
npm run build

# Link for development
npm link
```

### Testing

```bash
# Run tests
npm test

# Run n8n with the development node
n8n start
```

## 📋 Configuration

### Node Properties

| Property | Type | Description |
|----------|------|-------------|
| Resource | Options | Choose between File or Data operations |
| Operation | Options | Select the specific action to perform |
| File Name | String/Dropdown | Target file (AI-compatible input) |
| Sheet Name | String/Dropdown | Excel sheet selection (CSV uses Sheet1) |
| Filters | Collection | Advanced filtering conditions |
| Sort | Object | Column and direction for sorting |
| Limit | Number | Row limit (0 = all rows) |

### Dynamic Options

- **File Names**: Auto-populated from data directory
- **Sheet Names**: Dynamically loaded from Excel files
- **Column Names**: Automatically detected for filters and sorting

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 Changelog

### v1.2.7 (Latest)
- Added Delete Rows operation with filter-based row deletion
- Enhanced Update operation to Update Rows with filter-based multi-row updates
- Improved filter logic with reusable helper function
- Removed key column dependency in favor of flexible filter criteria

### v1.2.6
- Fixed repository and homepage URLs
- Added developer information and contact details
- Added keywords for better npm discoverability
- Added full MIT license text

### v1.2.5
- Enhanced README with comprehensive English documentation
- Added detailed usage examples and badges

### v1.2.4
- Added dynamic column dropdown for filters and sorting
- Improved CSV file handling for column detection

### v1.2.3
- Advanced filtering system with multiple operators
- Sorting functionality with type detection
- Improved limit handling (0 for all rows)

### v1.2.2
- Simplified display name to "Excel/CSV"

### v1.2.1
- Removed unnecessary AI Tool Resource
- Streamlined AI Agent integration

### v1.2.0
- Full AI Agent Tool compatibility
- Direct AI input support for file and sheet names

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2024 Dante Labs

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 👨‍💻 Developer Information

- **Name**: Dante
- **Company**: Dante Labs
- **Email**: datapod.k@gmail.com
- **Company Homepage**: [https://dante-datalab.com](https://dante-datalab.com)
- **YouTube**: [https://youtube.com/@dante-labs](https://youtube.com/@dante-labs)

## 🙏 Acknowledgments

- Built for the [n8n](https://n8n.io) workflow automation platform
- Uses [SheetJS](https://sheetjs.com/) for Excel file processing
- Uses [fast-csv](https://c2fo.github.io/fast-csv/) for CSV operations

## 💬 Support

- **Issues**: [GitHub Issues](https://github.com/dandacompany/n8n-nodes-excel/issues)
- **Discussions**: [GitHub Discussions](https://github.com/dandacompany/n8n-nodes-excel/discussions)
- **n8n Community**: [n8n Community Forum](https://community.n8n.io)
- **Developer Contact**: datapod.k@gmail.com

## 🌟 Star History

If you find this node useful, please consider giving it a star on GitHub!

---

**Made with ❤️ by Dante Labs for the n8n community**