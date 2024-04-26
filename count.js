document.addEventListener('DOMContentLoaded', function() {
    const excelFilePath = 'books.xlsx';
    
    function excelToTable(workbook) {
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        // Sorting data based on the "count" column (assuming "count" is in the 5th column)
        data.sort((a, b) => b[4] - a[4]); // Sorting in descending order based on "count"
        
        const tableBody = document.getElementById('books-body');
        tableBody.innerHTML = ''; // Clear previous data
        
        let mostReadBook = '';
        
        data.forEach((row, rowIndex) => {
            if (rowIndex === 0) return; // Skip header row
            const [title, author, year, availability, count] = row;
            const tableRow = document.createElement('tr');
            tableRow.innerHTML = `
                <td>${title}</td>
                <td>${author}</td>
                <td>${year}</td>
                <td>${availability}</td>
                <td>${count}</td>
            `;
            tableBody.appendChild(tableRow);
            
            // Update mostReadBook if it's the first iteration or if the count is higher
            if (rowIndex === 1 || count > data[1][4]) {
                mostReadBook = title;
            }
        });
        
        const mostReadElement = document.getElementById('most-read');
        mostReadElement.textContent = `Most read book is "${mostReadBook}"`;
    }
    
    fetch(excelFilePath)
        .then(response => response.arrayBuffer())
        .then(buffer => {
            const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
            excelToTable(workbook);
        })
        .catch(error => console.error('Error fetching Excel file:', error));
});
