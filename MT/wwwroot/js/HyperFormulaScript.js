import HyperFormula from 'hyperformula/dist/hyperformula.full.min.js';
import sql from 'mssql/msnodesqlv8.js';

// Database configuration with Windows Authentication for local server
const config = {
    server: '(localdb)\\mssqllocaldb', // Use the actual server name if it's different
    database: 'MT',
    driver: 'msnodesqlv8',

    //user: 'fmn',
    //password: 'Test123',

    options: {
        trustedConnection: true,
        //integratedSecurity: true,
        trustServerCertificate: true,
    },
};

// Function to insert data into the database
async function insertData(data) {
    try {
        // Connect to the database
        await sql.connect(config);

        // Iterate over each row in the data array and insert into the database
        for (const row of data) {
            const { Name, Result } = row;

            // Insert data into the database
            await sql.query`UPDATE Formulas SET Result = ${Result} WHERE Name = ${Name}`;
        }

        return { success: true, message: 'Data inserted successfully' };
    } catch (error) {
        console.error(error);
        return { success: false, message: 'Error updating data in the database' };
    } finally {
        // Close the database connection
        await sql.close();
    }
}

export function CalculateFormula(callback = handleResult, data) {
    const tableData = generateTableData(data);
    insertData(tableData);
    return tableData;
}

function handleResult(result) {
    console.log(result);
}

function generateTableData(data) {

    const hf = HyperFormula.buildFromSheets(data, { useColumnIndex: true, licenseKey: "gpl-v3", maxRows: 200000 });

    const results = [];

    const sheetID = hf.getSheetId('Formulas');
    var row = 0;

    while (hf.getCellValue({ sheet: sheetID, col: 0, row: row }) !== null) {
        // Create a new row to add
        const newRow = { Name: hf.getCellValue({ sheet: sheetID, col: 1, row: row }), Result: hf.getCellValue({ sheet: sheetID, col: 0, row: row }) };

        results.push(newRow);
        row++;
    }

    return results;
}