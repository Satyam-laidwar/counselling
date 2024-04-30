/*console.log("send");
const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = 3000;


// Middleware for parsing JSON and urlencoded form data
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// POST endpoint to handle form submission
app.post('/submit-form', (req, res) => {
    // Extract data from the request body
    const { name, email, mob1, mob2, marks, address, description } = req.body;

    // Read existing Excel file
    const filePath = path.join(__dirname, 'test.xlsx');
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets['Sheet1'];

    // Append data to the worksheet
    const newRow = { Name: name, Email: email, mob1, mob2, marks, address, Description: description };
    XLSX.utils.sheet_add_json(worksheet, [newRow], { header: ['Name', 'Email', 'mob1', 'mob2', 'marks', 'address', 'Description'], skipHeader: true });

    // Write the updated workbook back to the file
    XLSX.writeFile(workbook, filePath);

    // Send a response to the client
    res.send('Data added to Excel file successfully.');
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});

*/


/*
const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = 3000;

// Middleware for parsing JSON and urlencoded form data
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// POST endpoint to handle form submission
app.post('/submit-form', (req, res) => {
    try {
        // Extract data from the request body
        const { name, email, mob1, mob2, marks, address, description } = req.body;

        // Read existing Excel file
        const filePath = path.join(__dirname, 'test.xlsx');
        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets['Sheet1'];

        // Find the next available row
        const nextRow = XLSX.utils.decode_range(worksheet['!ref']).e.r + 1;

        // Append data to the next available row
        worksheet[`A${nextRow}`] = { t: 's', v: name };
        worksheet[`B${nextRow}`] = { t: 's', v: email };
        worksheet[`C${nextRow}`] = { t: 's', v: mob1 };
        worksheet[`D${nextRow}`] = { t: 's', v: mob2 };
        worksheet[`E${nextRow}`] = { t: 's', v: marks };
        worksheet[`F${nextRow}`] = { t: 's', v: address };
        worksheet[`G${nextRow}`] = { t: 's', v: description };

        // Write the updated workbook back to the file
        XLSX.writeFile(workbook, filePath);

        // Send a response to the client
        res.send('Data added to Excel file successfully.');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Internal server error');
    }
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});


/*
const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = 3000;

// Middleware for parsing JSON and urlencoded form data
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// POST endpoint to handle form submission
app.post('/submit-form', (req, res) => {
    // Extract data from the request body
    let { name, email, mob1, mob2, marks, address, description } = req.body;

    // Read existing Excel file
    const filePath = path.join(__dirname, 'test.xlsx');
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets['Sheet1'];

    // Append data to the worksheet
    let newRow = { Name: name, Email: email, mob1, mob2, marks, address, Description: description };
    XLSX.utils.sheet_add_json(worksheet, [newRow], { header: ['Name', 'Email', 'mob1', 'mob2', 'marks', 'address', 'Description'], skipHeader: true });

    // Write the updated workbook back to the file
    XLSX.writeFile(workbook, filePath);

    // Send a response to the client
    res.send('Data added to Excel file successfully.');
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
*/

/*  code with find last row and rappend
const ExcelJS = require('exceljs');
let lastFilledRowNumber;

// Load the existing workbook
const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile('test.xlsx')
    .then(() => {
        // Access the worksheet
        const worksheet = workbook.getWorksheet('Sheet1');

        // Find the last filled row
        let lastRow = worksheet.lastRow;
        while (lastRow.getCell(1).value === null) {
            lastRow = worksheet.getRow(lastRow._number - 1);
        }
        
        // Store the last filled row number in a variable
        lastFilledRowNumber = lastRow._number;
        
        console.log('Last filled row:', lastFilledRowNumber);
    })
    .catch(err => {
        console.error('Error finding last filled row:', err);
    });


const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = 3000;

// Middleware for parsing JSON and urlencoded form data
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// POST endpoint to handle form submission
app.post('/submit-form', (req, res) => {
    try {
        // Extract data from the request body
        const { name, email, mob1, mob2, marks, address, description } = req.body;

        // Read existing Excel file
        const filePath = path.join(__dirname, 'test.xlsx');
        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets['Sheet1'];

        // Find the next available row
        const nextRow = XLSX.utils.decode_range(worksheet['!ref']).e.r + (lastFilledRowNumber  -1);

        // Append data to the next available row
        worksheet[`A${nextRow}`] = { t: 's', v: name };
        worksheet[`B${nextRow}`] = { t: 's', v: email };
        worksheet[`C${nextRow}`] = { t: 's', v: mob1 };
        worksheet[`D${nextRow}`] = { t: 's', v: mob2 };
        worksheet[`E${nextRow}`] = { t: 's', v: marks };
        worksheet[`F${nextRow}`] = { t: 's', v: address };
        worksheet[`G${nextRow}`] = { t: 's', v: description };

        // Update the range to include the new row
        worksheet['!ref'] = XLSX.utils.encode_range({
            s: { c: 0, r: 0 },
            e: { c: 6, r: nextRow }
        });

        // Write the updated workbook back to the file
        XLSX.writeFile(workbook, filePath);

        // Send a response to the client
        res.send('Data added to Excel file successfully.');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Internal server error');
    }
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});

*/
const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const path = require('path');
const app = express();
const port = 3000;

app.get('/favicon.ico', (req, res) => res.status(204));

// Middleware for parsing JSON and urlencoded form data
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Function to append data to Excel file
async function appendToExcel(name, email, mob1, mob2, marks, address, description) {
    try {
        const filePath = path.join(__dirname, 'test.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.getWorksheet('Sheet1');

        // Find the next available row
        let nextRow = worksheet.lastRow ? worksheet.lastRow.number + 1 : 1;

        // Append data to the next available row
        const newRow = worksheet.getRow(nextRow);
        newRow.getCell('A').value = name;
        newRow.getCell('B').value = email;
        newRow.getCell('C').value = mob1;
        newRow.getCell('D').value = mob2;
        newRow.getCell('E').value = marks;
        newRow.getCell('F').value = address;
        newRow.getCell('G').value = description;

        // Save the changes
        await workbook.xlsx.writeFile(filePath);
        return nextRow;
    } catch (error) {
        console.error('Error:', error);
        throw error;
    }
}

// POST endpoint to handle form submission
app.post('/submit-form', async (req, res) => {
    try {
        // Extract data from the request body
        const { name, email, mob1, mob2, marks, address, description } = req.body;

        // Append data to Excel file
        const newRowNumber = await appendToExcel(name, email, mob1, mob2, marks, address, description);

        // Send a response to the client
        res.send(`Data added to Excel file successfully. Row number: ${newRowNumber}`);
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Internal server error');
    }
});



// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
