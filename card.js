document.addEventListener('DOMContentLoaded', function() {
    const excelFilePath = 'books_in_use.xlsx';

    document.getElementById('loadDataBtn').addEventListener('click', function() {
        fetch(excelFilePath)
            .then(response => response.arrayBuffer())
            .then(buffer => {
                const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
                displayBooks(workbook);
            })
            .catch(error => console.error('Error fetching Excel file:', error));
    });

    function displayBooks(workbook) {
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const booksList = document.getElementById('booksList');
        booksList.innerHTML = ''; // Clear previous data

        data.forEach((row, rowIndex) => {
            if (rowIndex === 0) return; // Skip header row
            const [student, title, author, year] = row;
            const bookEntry = document.createElement('div');
            bookEntry.innerHTML = `
                <p><strong>Student:</strong> ${student}</p>
                <p><strong>Title:</strong> ${title}</p>
                <p><strong>Author:</strong> ${author}</p>
                <p><strong>Year:</strong> ${year}</p>
                <hr>
            `;
            booksList.appendChild(bookEntry);
        });
    }
});
