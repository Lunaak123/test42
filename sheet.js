let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
        alert("Failed to load Excel file. Please check the URL and try again.");
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null ? 'NULL' : cell; // Show 'NULL' for null values
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply operations based on user input
function applyOperations() {
    const primaryCol = document.getElementById('primary-column').value.trim().toUpperCase();
    const operationCols = document.getElementById('operation-columns').value.trim().toUpperCase().split(',');
    const operationType = document.getElementById('operation-type').value;
    const nullCheck = document.getElementById('operation').value;

    if (!primaryCol || operationCols.length === 0) {
        alert("Please enter the primary column and at least one operation column.");
        return;
    }

    const primaryIndex = columnToIndex(primaryCol);
    const operationIndices = operationCols.map(col => columnToIndex(col));

    filteredData = data.filter(row => {
        const primaryValue = row[primaryCol];
        if (nullCheck === 'null' && primaryValue !== null) return false;
        if (nullCheck === 'not-null' && primaryValue === null) return false;

        return operationIndices.every(index => {
            const value = row[operationCols[index]];
            if (nullCheck === 'null') {
                return value === null;
            } else {
                return value !== null;
            }
        });
    });

    displaySheet(filteredData);
}

// Helper function to convert column letter to index (e.g., A => 0, B => 1)
function columnToIndex(column) {
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    return alphabet.indexOf(column);
}

// Download functionality (Open Modal)
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});

// Confirm download functionality
document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value || 'download';
    const format = document.getElementById('file-format').value;

    if (format === 'xlsx') {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(filteredData);
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (format === 'csv') {
        const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(filteredData));
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${filename}.csv`;
        link.click();
    } else {
        alert('Image and PDF downloads are not implemented yet.');
    }

    document.getElementById('download-modal').style.display = 'none'; // Close modal after download
});

// Close modal functionality
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Load the Excel sheet on page load (replace with your Excel file URL)
window.onload = () => loadExcelSheet('https://example.com/path/to/your/excel/file.xlsx');
