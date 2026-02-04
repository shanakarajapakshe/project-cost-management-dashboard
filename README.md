# Profit Dashboard - Desktop Application

A professional, offline desktop application built with Electron, Node.js, HTML, CSS, Bootstrap, and Chart.js for managing project costs, revenues, and profits with automatic Excel file generation and storage.

## Features

### ✨ Core Features
- **Monthly Data Management**: Select month/year to create and manage separate Excel files for each period
- **Automatic Excel Generation**: Creates `.xlsx` files automatically when selecting a new month/year
- **Real-time Calculations**: Automatically calculates:
  - CE Visit Cost = CE Visit Charge × Visits/Month
  - Direct Cost = Engineer Cost + CE Visit Cost + Transport Cost
  - Overhead Cost = Direct Cost × (Overhead Allocation % / 100)
  - Total Cost = Direct Cost + Overhead Cost
  - Profit = Client Payment - Total Cost
  - Engineer Cost = Engineer Salary/Month

### 📊 Dashboard & Analytics
- **Summary Statistics**: Total projects, revenue, costs, and net profit
- **Interactive Charts**:
  - Total Cost vs Client Payment (Bar Chart)
  - Profit per Project (Bar Chart)
  - Monthly Profit Trend (Line Chart)
  - Overhead % vs Profit (Scatter Chart)
- **Data Table**: View all projects with cost, payment, and profit details
- **Project Details**: View detailed breakdown of any project

### 💾 Data Management
- **Excel Storage**: All data stored in structured `.xlsx` files
- **Automatic Backups**: Creates backup copies of Excel files
- **Export Functionality**: Export Excel files to any location
- **Offline Support**: Works completely offline
- **Data Persistence**: Uses Excel files for long-term storage

### 🎨 User Interface
- **Professional Design**: Clean, modern Bootstrap-based interface
- **Responsive Layout**: Works on different screen sizes
- **Form Validation**: Ensures all inputs are valid before submission
- **Toast Notifications**: Real-time feedback for all actions
- **Instant Updates**: Dashboard updates without page reload

## Installation

### Prerequisites
- Node.js (v16 or higher)
- npm (comes with Node.js)

### Setup Steps

1. **Extract/Clone the project**
   ```bash
   cd profit-dashboard
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Run the application**
   ```bash
   npm start
   ```

## Usage Guide

### 1. Select Month & Year
- On app launch, select the desired month and year
- Click "Load Month Data" to create/load the Excel file for that period
- The system automatically creates a new Excel file if it doesn't exist

### 2. Add Projects
Fill in the data entry form with:
- **Project Name**: Unique identifier for the project
- **No. of Engineers**: Number of engineers assigned
- **Engineer Salary/Month**: Monthly salary per engineer
- **CE Visit Charge**: Charge per customer engineer visit
- **Visits/Month**: Number of visits per month
- **Transport Cost/Month**: Monthly transportation costs
- **Client Payment/Month**: Monthly payment from client
- **Overhead Allocation (%)**: Overhead percentage to apply

Click "Save Project" to add the project to the database.

### 3. View Dashboard
After adding projects, the dashboard displays:
- Summary statistics cards
- Projects table with all details
- Four interactive charts for data visualization

### 4. Manage Projects
- **View Details**: Click "View" button to see detailed project breakdown
- **Delete Project**: Click "Delete" button to remove a project
- **Export Excel**: Click "Export to Excel" to save the file to a custom location

## Project Structure

```
profit-dashboard/
├── index.html          # Main HTML file with UI structure
├── styles.css          # CSS styling and responsive design
├── app.js             # Frontend JavaScript (data management & UI)
├── main.js            # Electron main process
├── preload.js         # Electron preload script (security)
├── excelService.js    # Excel file operations with ExcelJS
├── package.json       # Project dependencies and scripts
└── README.md          # This file
```

## Excel File Structure

Each monthly Excel file contains two worksheets:

### Projects Sheet
Columns:
- Project Name
- No. of Engineers
- Engineer Salary/Month
- CE Visit Charge
- Visits/Month
- Transport Cost/Month
- Client Payment/Month
- Overhead Allocation %
- Engineer Cost (calculated)
- CE Visit Cost (calculated)
- Direct Cost (calculated)
- Overhead Cost (calculated)
- Total Cost (calculated)
- Profit (calculated)
- Timestamp

### Summary Sheet
- Total Projects
- Total Revenue
- Total Costs
- Net Profit

All calculations use Excel formulas, making the files fully functional in Excel/LibreOffice.

## File Locations

Excel files are stored in:
- **Windows**: `C:\Users\[Username]\AppData\Roaming\profit-dashboard\excel-files\`
- **macOS**: `~/Library/Application Support/profit-dashboard/excel-files/`
- **Linux**: `~/.config/profit-dashboard/excel-files/`

Backups are stored in a `backups` subfolder within the above directory.

## Building Executables

### Build for Windows
```bash
npm run build-win
```

### Build for macOS
```bash
npm run build-mac
```

### Build for Linux
```bash
npm run build-linux
```

Built applications will be in the `dist` folder.

## Technology Stack

- **Electron**: Desktop application framework
- **Node.js**: JavaScript runtime
- **ExcelJS**: Excel file generation and manipulation
- **Bootstrap 5**: UI framework
- **Chart.js**: Data visualization
- **HTML5/CSS3**: Modern web standards

## Database Migration Path

The application is designed with modular code to easily upgrade to database storage:

### Current: Excel Storage
```javascript
class DataManager {
    async addProject(projectData) {
        // Current: Excel via Electron API
        await window.electronAPI.saveProject(month, year, projectData);
    }
}
```

### Future: SQLite/MySQL
```javascript
class DataManager {
    async addProject(projectData) {
        // Future: Database query
        await db.query('INSERT INTO projects ...', projectData);
    }
}
```

The `DataManager` class abstracts all data operations, making it easy to swap storage backends without changing UI code.

## Features Roadmap

### Phase 1 (Current)
- ✅ Monthly Excel file management
- ✅ Automatic calculations
- ✅ Dashboard with charts
- ✅ Offline support
- ✅ Automatic backups

### Phase 2 (Future)
- 📋 SQLite/MySQL database integration
- 📋 Multi-user support
- 📋 Advanced reporting
- 📋 Data import/export (CSV, JSON)
- 📋 Custom chart configurations

### Phase 3 (Future)
- 📋 Cloud sync
- 📋 Mobile companion app
- 📋 PDF report generation
- 📋 Email notifications
- 📋 Budget forecasting

## Troubleshooting

### Excel files not saving
- Check write permissions in the application data folder
- Ensure ExcelJS is properly installed: `npm install exceljs`

### Charts not displaying
- Verify Chart.js is loaded from CDN
- Check browser console for JavaScript errors

### Application won't start
- Ensure all dependencies are installed: `npm install`
- Check Node.js version: `node --version` (should be v16+)

## Support

For issues, questions, or feature requests, please:
1. Check the README documentation
2. Review the code comments in each module
3. Check the browser/Electron console for error messages

## License

MIT License - Feel free to use and modify for your needs.

## Credits

Built with ❤️ using:
- [Electron](https://www.electronjs.org/)
- [ExcelJS](https://github.com/exceljs/exceljs)
- [Bootstrap](https://getbootstrap.com/)
- [Chart.js](https://www.chartjs.org/)
