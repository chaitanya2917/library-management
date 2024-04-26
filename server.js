const express = require('express');
const app = express();
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const fs = require('fs');

// Define the path to the Excel file
const filePath = 'D:/library management/books.xlsx';

// Middleware to parse JSON bodies
app.use(bodyParser.json());

// Route handler for the root route
app.get('/', (req, res) => {
    res.send('Hello, World!');
});

// Route handler for adding a book
app.post('/add-book', async (req, res) => {
    const { title, author, year } = req.body;

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
            workbook.addWorksheet('Books');
        }

        // Get the first worksheet (assuming it's the only one)
        const worksheet = workbook.getWorksheet(1);

        // Add the new book information as a row
        worksheet.addRow([title, author, year]);

        // Save the workbook back to the file
        await workbook.xlsx.writeFile(filePath);

        console.log('Book information saved successfully!');
        res.status(200).send('Book information saved successfully!');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('An error occurred while saving the book information.');
    }
});

// Start the server
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
