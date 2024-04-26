const ExcelJS = require('exceljs');
const fs = require('fs');

// Define the path to the Excel file
const filePath = 'D:/library management/books.xlsx';

async function updateExcel(title, author, year) {
    try {
        // Check if the file exists
        const fileExists = fs.existsSync(filePath);

        let workbook;
        if (fileExists) {
            // If the file exists, load it
            workbook = await ExcelJS.readFile(filePath);
        } else {
            // If the file doesn't exist, create a new workbook
            workbook = new ExcelJS.Workbook();
            // Add a worksheet
            workbook.addWorksheet('Books');
        }

        // Get the first worksheet
        const worksheet = workbook.getWorksheet('Books');

        // Add the new book information as a row
        worksheet.addRow([title, author, year]);

        // Save the workbook back to the file
        await workbook.xlsx.writeFile(filePath);

        console.log('Book information saved successfully!');
    } catch (error) {
        console.error('Error:', error);
    }
}


// Example usage
updateExcel('Book Title', 'Author Name', 2024);
