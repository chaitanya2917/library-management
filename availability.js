function checkAvailability() {
    const bookTitle = document.getElementById('book-title').value.trim();
    if (!bookTitle) {
        alert('Please enter a book title.');
        return;
    }

    const excelFilePath = 'books.xlsx';

    fetch(excelFilePath)
        .then(response => response.arrayBuffer())
        .then(buffer => {
            const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            let availability = '';
            for (let i = 1; i < data.length; i++) {
                const [title, , , available] = data[i];
                if (title.toLowerCase() === bookTitle.toLowerCase()) {
                    availability = available === 'Y' ? 'available' : 'unavailable';
                    break;
                }
            }

            const resultContainer = document.getElementById('availability-result');
            if (availability) {
                resultContainer.innerHTML = `<p>Book "${bookTitle}" is ${availability}.</p>`;
            } else {
                resultContainer.innerHTML = `<p>Book "${bookTitle}" not found.</p>`;
            }
        })
        .catch(error => console.error('Error fetching Excel file:', error));
}
