# FileMaker to Excel Export Guide
## Material Consumption Report Template

This guide explains how to use the HTML/CSS template with FileMaker Pro and export styled reports to Excel.

---

## 📋 Overview

The template (`material-report-template.html` + `material-report-styles.css`) provides a professional material consumption report layout that:
- Matches your existing report format
- Exports cleanly to Excel with preserved styling
- Works with FileMaker Web Viewer or Export functions
- Supports dynamic data population

---

## 🚀 Method 1: FileMaker Web Viewer (Recommended)

### Step 1: Create a Web Viewer Layout
1. In FileMaker, create a new layout for your report
2. Add a Web Viewer object (Insert > Web Viewer)
3. Size it to fill the layout

### Step 2: Generate HTML with FileMaker Data
Create a calculation field that generates the HTML:

```filemaker
Let([
    ~html = "<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>" & 
    /* Paste entire material-report-styles.css content here */ 
    "</style>
</head>
<body>
    <div class='report-container'>
        <div class='report-header'>
            <h1>MATERIAL CONSUMPTION REPORT</h1>
            <div class='report-meta'>
                <span>Date: " & Get(CurrentDate) & "</span>
                <span>Period: " & YourPeriodField & "</span>
            </div>
        </div>
        <table class='material-table'>
            <thead>
                <tr class='header-row'>
                    <th>MATERIAL</th>
                    <th>DESCRIPTION</th>
                    <th>UNIT</th>
                    <th>OPENING</th>
                    <th>RECEIPT</th>
                    <th>CONSUMED QTY</th>
                    <th>RETURNED QTY</th>
                    <th>STORES REJECTED</th>
                    <th>STOCK</th>
                    <th>PROD stock</th>
                    <th>Recovery</th>
                    <th>TOBACCO STD</th>
                    <th>THEO. CONS</th>
                    <th>ACTUAL CONSU</th>
                    <th>DIFF OF LOSS</th>
                    <th>% LOSS</th>
                </tr>
            </thead>
            <tbody>" &
            
            /* Loop through records */
            List(
                "<tr class='data-row'>" &
                "<td class='col-material'>" & MaterialTable::MaterialCode & "</td>" &
                "<td class='col-description'>" & MaterialTable::Description & "</td>" &
                "<td class='col-unit'>" & MaterialTable::Unit & "</td>" &
                "<td class='col-number'>" & MaterialTable::Opening & "</td>" &
                "<td class='col-number'>" & MaterialTable::Receipt & "</td>" &
                "<td class='col-number'>" & MaterialTable::ConsumedQty & "</td>" &
                "<td class='col-number'>" & MaterialTable::ReturnedQty & "</td>" &
                "<td class='col-number'>" & MaterialTable::StoresRejected & "</td>" &
                "<td class='col-number'>" & MaterialTable::Stock & "</td>" &
                "<td class='col-number'>" & MaterialTable::ProdStock & "</td>" &
                "<td class='col-number'>" & MaterialTable::Recovery & "</td>" &
                "<td class='col-number'>" & MaterialTable::TobaccoSTD & "</td>" &
                "<td class='col-number'>" & MaterialTable::TheoreticalConsumption & "</td>" &
                "<td class='col-number'>" & MaterialTable::ActualConsumption & "</td>" &
                "<td class='col-number'>" & MaterialTable::DifferenceOfLoss & "</td>" &
                "<td class='col-percent'>" & MaterialTable::PercentageLoss & "</td>" &
                "</tr>"
            ) &
            
            "</tbody>
        </table>
    </div>
</body>
</html>"
];
    ~html
)
```

### Step 3: Set Web Viewer to Use Calculation
1. Select the Web Viewer
2. In Inspector > Data tab
3. Choose "Web Address" as calculation
4. Enter your HTML generation calculation field

---

## 📤 Method 2: Direct Excel Export

### Step 1: Create Export Layout
1. Create a new Table View layout
2. Add all required fields in the correct order
3. Format fields to match report structure

### Step 2: Export with Styling
```filemaker
Export Records [
    File: "MaterialReport.xlsx";
    Create folders: Off;
    Include: All records;
    Format: Excel Workbook (.xlsx)
]
```

### Step 3: Apply Excel Template
After export, apply the styling:
1. Open the exported Excel file
2. Import the HTML template as a web query
3. Or use Excel VBA to apply styles programmatically

---

## 🎨 Method 3: HTML Export with Styling

### Step 1: Create HTML Export Script

```filemaker
Set Variable [$htmlContent; Value: YourHTMLGenerationField]
Set Variable [$filePath; Value: Get(DesktopPath) & "MaterialReport.html"]

Export Field Contents [
    MaterialTable::HTMLContent;
    "$filePath";
    Automatically open file: Yes
]
```

### Step 2: Open in Browser and Print to PDF
The HTML will open in browser with full styling, then:
1. File > Print
2. Choose "Save as PDF"
3. Or use browser's built-in PDF export

---

## 🔧 FileMaker Field Mapping

### Required Fields in Your FileMaker Table:

| FileMaker Field | Type | Description |
|----------------|------|-------------|
| MaterialCode | Text | Material ID (e.g., TOB1118) |
| Description | Text | Full material description |
| Unit | Text | Unit of measurement (KGS, BB, M) |
| Opening | Number | Opening stock |
| Receipt | Number | Received quantity |
| ConsumedQty | Number | Consumed quantity |
| ReturnedQty | Number | Returned quantity |
| StoresRejected | Number | Stores rejected |
| Stock | Number | Current stock |
| ProdStock | Number | Production stock |
| Recovery | Number | Recovery amount |
| TobaccoSTD | Number | Tobacco standard |
| TheoreticalConsumption | Number | Theoretical consumption |
| ActualConsumption | Number | Actual consumption |
| DifferenceOfLoss | Number | Difference of loss |
| PercentageLoss | Text | Loss percentage |

### Calculated Fields:

```filemaker
// Percentage Loss Calculation
Let([
    ~diff = ActualConsumption - TheoreticalConsumption;
    ~percent = If(TheoreticalConsumption > 0; 
        Round(~diff / TheoreticalConsumption * 100; 2); 
        0)
];
    ~percent & "%"
)

// Stock Calculation
Opening + Receipt - ConsumedQty - ReturnedQty - StoresRejected

// Difference of Loss
ActualConsumption - TheoreticalConsumption
```

---

## 📊 Category Grouping

To add category headers (TOBACCO, CIGARETTE PAPER, etc.):

### Method 1: Summary Field
1. Create a field `MaterialCategory`
2. Sort records by `MaterialCategory`
3. Add Sub-summary part in layout
4. Place category header in sub-summary

### Method 2: Conditional Formatting
```filemaker
// In HTML generation, check for category change
If(
    MaterialTable::MaterialCategory ≠ 
    GetNthRecord(MaterialTable::MaterialCategory; Get(RecordNumber) - 1);
    
    "<tr class='category-row'><td colspan='16' class='category-header'>" & 
    MaterialTable::MaterialCategory & 
    "</td></tr>"
)
```

---

## 🎯 Advanced Features

### 1. Conditional Row Highlighting
```filemaker
// Add class based on loss percentage
Case(
    PercentageLoss > 5; "status-error";
    PercentageLoss > 2; "status-warning";
    "status-ok"
)
```

### 2. Number Formatting
```filemaker
// Format numbers with thousand separators
NumToText(
    Round(YourNumberField; 2);
    0; // decimal places
    1  // use thousand separator
)
```

### 3. Empty Cell Handling
```filemaker
If(
    IsEmpty(YourField) or YourField = 0;
    "-";
    YourField
)
```

---

## 📱 Export to Excel with Preserved Styling

### Using FileMaker's Save as Excel:
1. Go to your report layout
2. File > Save/Send Records As > Excel
3. Choose "Create email with file as attachment" or "Save to file"
4. Formatting will be preserved

### Using Script:
```filemaker
Set Variable [$path; Value: Get(DesktopPath) & "Report_" & Get(CurrentDate) & ".xlsx"]

Save Records as Excel [
    File Name: $path;
    Create email: No;
    Automatically open: Yes
]
```

---

## 🔍 Troubleshooting

### Issue: Styling not appearing in Excel
**Solution:** Ensure you're using inline CSS or embedded styles in the HTML

### Issue: Numbers not aligning properly
**Solution:** Use `text-align: right` and monospace font for number columns

### Issue: Category headers not showing
**Solution:** Check your sub-summary parts are properly configured

### Issue: Web Viewer not displaying
**Solution:** Check calculation syntax and ensure HTML is valid

---

## 📚 Additional Resources

### FileMaker Functions Used:
- `List()` - Aggregate records into HTML
- `GetNthRecord()` - Access specific record data
- `Get(CurrentDate)` - Current date
- `Get(RecordNumber)` - Current record number
- `ExecuteSQL()` - Query data for complex reports

### CSS Classes Available:
- `.category-row` - Category headers
- `.data-row` - Regular data rows
- `.highlight-row` - Alternating row colors
- `.status-ok` - Green highlight
- `.status-warning` - Yellow highlight
- `.status-error` - Red highlight

---

## 💡 Best Practices

1. **Performance**: For large datasets (>1000 records), consider pagination
2. **Formatting**: Use FileMaker's number formatting before HTML generation
3. **Testing**: Test with sample data before full export
4. **Backup**: Always keep a backup before modifying layouts
5. **Validation**: Validate data before export to avoid errors

---

## 📞 Support

For FileMaker-specific questions:
- FileMaker Community: https://community.claris.com
- FileMaker Documentation: https://help.claris.com

For template customization:
- Modify `material-report-styles.css` for styling changes
- Update `material-report-template.html` for structure changes

---

**Last Updated:** January 2026
**Version:** 1.0
**Compatible with:** FileMaker Pro 19+, Excel 2016+