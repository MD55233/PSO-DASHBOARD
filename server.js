const express = require('express');
const multer = require('multer');
const path = require('path');
const cors = require('cors');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
app.use(cors());

const PORT = 8000;

// Middleware to parse "fileType" from form-data
const parseFileType = (req, res, next) => {
    if (!req.query.fileType && !req.body.fileType) {
        return res.status(400).send({ message: 'fileType is required.' });
    }
    next();
};

// Dynamic storage setup
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        const fileType = req.query.fileType || req.body.fileType;

        let folder = '';
        if (fileType === 'lubricants') {
            folder = 'uploads/excel-files/lubricants';
        } else if (fileType === 'petroleum') {
            folder = 'uploads/excel-files/petroleum';
        } else {
            return cb(new Error('Invalid fileType. Allowed: lubricants, petroleum'), null);
        }

        // Ensure the folder exists
        fs.mkdirSync(folder, { recursive: true });
        cb(null, folder);
    },
    filename: function (req, file, cb) {
        const uniqueSuffix = Date.now() + '-' + file.originalname;
        cb(null, uniqueSuffix);
    }
});

const upload = multer({ storage });

// Route to handle file upload
app.post('/upload-excel', parseFileType, upload.array('files', 5), (req, res) => {
    try {
        res.status(200).send({ message: 'Files uploaded successfully!' });
    } catch (error) {
        res.status(500).send({ message: 'Failed to upload files.', error: error.message });
    }
});

{/*-------------------------------- Validate Header in sheets ------------------------------------------*/}

// Validate headers function
const validateHeaders = (headers) => {
    const expectedHeaders1 = [
        'sales grp', 'customer code', 'customer name',
        'material code', 'shipping point name', 'vehicle text',
        'billing date', 'quantity in su'
    ].map(h => h.trim().toLowerCase());

    const expectedHeaders2 = [
        'sales grp', 'customer code', 'customer name',
        'material name', 'material code', 'quantity in su',
        'billing date', 'sku qty'
    ].map(h => h.trim().toLowerCase());

    return (
        headers.every(h => expectedHeaders1.includes(h)) ||
        headers.every(h => expectedHeaders2.includes(h))
    );
};
{/*--------------------------------//Get Header Indexes------------------------------------------*/}
//Get Header Indexes
const getColumnIndices = (headers, requiredHeaders) => {
    const normalizedHeaders = headers.map(h => h.trim().toLowerCase());
    const indices = {};

    requiredHeaders.forEach(header => {
        const normalizedHeader = header.trim().toLowerCase();
        indices[normalizedHeader] = normalizedHeaders.indexOf(normalizedHeader);
    });

    return indices;
};
{/*--------------------------------process a single Excel file for home page  ------------------------------------------*/}
// Function to process a single Excel file
const processExcelFile = (filePath) => {
    try {
        console.log(`Processing file: ${filePath}`);
        const workbook = XLSX.readFile(filePath);
        const sheets = workbook.SheetNames;
        const uniqueUsers = new Set();

        let totalUsers = 0;
        let totalOrders = 0;
        let totalSales = 0;
        let salesByYear = {};

        sheets.forEach(sheetName => {
            console.log(`Processing sheet: ${sheetName}`);
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (jsonData.length === 0) return;

            const headers = jsonData[0].map(h => h.trim().toLowerCase());
            if (validateHeaders(headers)) {
                const indices = getColumnIndices(headers, ['billing date', 'quantity in su']);
                const dateIndex = indices['billing date'];
                const quantityIndex = indices['quantity in su'];

                jsonData.slice(1).forEach(row => {
                    const customerCode = row[1]; // Customer Code column
                    const quantityInSU = parseFloat(row[quantityIndex]) || 0;
                    const saleDate = row[dateIndex];
                    let year;

                    if (saleDate) {
                        if (typeof saleDate === 'number') {
                            const excelDate = new Date((saleDate - 25569) * 86400 * 1000);
                            year = excelDate.getFullYear();
                        } else {
                            const parsedDate = new Date(saleDate);
                            if (!isNaN(parsedDate)) {
                                year = parsedDate.getFullYear();
                            }
                        }
                    }

                    // Aggregate sales by year
                    if (year) {
                        if (!salesByYear[year]) salesByYear[year] = 0;
                        salesByYear[year] += quantityInSU;
                    }

                    // Aggregate overall totals
                    if (customerCode && !uniqueUsers.has(customerCode)) {
                        uniqueUsers.add(customerCode);
                        totalUsers++;
                    }
                    totalOrders++;
                    totalSales += quantityInSU;
                });
            }
        });

        return { totalUsers, totalOrders, totalSales, salesByYear };
    } catch (error) {
        console.error(`Error processing file ${filePath}:`, error.message);
        throw error;
    }
};
{/*--------------------------------process all files in a directory for home page------------------------------------------*/}
// Function to process all files in a directory
const processDirectory = (directoryPath) => {
    try {
        const files = fs.readdirSync(directoryPath);
        let totalUsers = 0;
        let totalOrders = 0;
        let totalSales = 0;
        let salesByYear = {};

        files.forEach(file => {
            const filePath = path.join(directoryPath, file);
            if (fs.lstatSync(filePath).isFile() && file.endsWith('.xlsx')) {
                const { totalUsers: users, totalOrders: orders, totalSales: sales, salesByYear: yearData } = processExcelFile(filePath);

                totalUsers += users;
                totalOrders += orders;
                totalSales += sales;

                // Merge year-based data
                for (const year in yearData) {
                    if (!salesByYear[year]) salesByYear[year] = 0;
                    salesByYear[year] += yearData[year];
                }
            }
        });

        return { totalUsers, totalOrders, totalSales, salesByYear };
    } catch (error) {
        console.error(`Error processing directory ${directoryPath}:`, error.message);
        throw error;
    }
};

{/*--------------------------------year ministates for home page ------------------------------------------*/}
// API route to fetch aggregated data
app.get('/fetch-excel-data', async (req, res) => {
    try {
        const lubricantsPath = path.join(__dirname, 'uploads/excel-files/lubricants');
        const petroleumPath = path.join(__dirname, 'uploads/excel-files/petroleum');

        const lubricantsData = processDirectory(lubricantsPath);
        const petroleumData = processDirectory(petroleumPath);

        const totalData = {
            totalUsers: lubricantsData.totalUsers + petroleumData.totalUsers,
            totalOrders: lubricantsData.totalOrders + petroleumData.totalOrders,
            totalSales: lubricantsData.totalSales + petroleumData.totalSales
        };

        res.status(200).json(totalData);
    } catch (error) {
        console.error('Error in /fetch-excel-data:', error.message);
        res.status(500).json({ message: 'Internal Server Error', error: error.message });
    }
});

{/*--------------------------------year charts for home page ------------------------------------------*/}


// API route to fetch total sales by year for charts/graphs
app.get('/fetch-sales-by-year', async (req, res) => {
    try {
        const lubricantsPath = path.join(__dirname, 'uploads/excel-files/lubricants');
        const petroleumPath = path.join(__dirname, 'uploads/excel-files/petroleum');

        const lubricantsData = processDirectory(lubricantsPath);
        const petroleumData = processDirectory(petroleumPath);

        const totalSalesByYear = {};

        // Merge year data from both directories
        [lubricantsData.salesByYear, petroleumData.salesByYear].forEach(yearData => {
            for (const year in yearData) {
                if (!totalSalesByYear[year]) totalSalesByYear[year] = 0;
                totalSalesByYear[year] += yearData[year];
            }
        });

        res.status(200).json(totalSalesByYear);
    } catch (error) {
        console.error('Error in /fetch-sales-by-year:', error.message);
        res.status(500).json({ message: 'Internal Server Error', error: error.message });
    }
});


// Process the Excel file and filter data based on year and month
const processExcelFileForTable = (filePath, yearFilter, monthFilter) => {
    try {
        console.log(`Processing file for table: ${filePath}`);
        const workbook = XLSX.readFile(filePath);
        const sheets = workbook.SheetNames;
        const tableData = [];

        sheets.forEach(sheetName => {
            console.log(`Processing sheet: ${sheetName}`);
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (jsonData.length === 0) return;

            const headers = jsonData[0].map(h => h.trim().toLowerCase());

            if (validateHeaders(headers)) {
                jsonData.slice(1).forEach(row => {
                    const rowData = {};
                    let includeRow = true; // Flag to check if the row matches the filters

                    headers.forEach((header, index) => {
                        let value = row[index];

                        // Handle date conversion
                        if (header.includes('date') && !isNaN(value)) {
                            const date = convertExcelDateToJSDate(value);
                            const rowYear = new Date(date).getFullYear();
                            const rowMonth = new Date(date).getMonth() + 1;

                            if (yearFilter && rowYear !== yearFilter || monthFilter && rowMonth !== monthFilter) {
                                includeRow = false;
                            }
                            value = date;
                        }

                        rowData[header] = value;
                    });

                    if (includeRow) {
                        tableData.push(rowData);
                    }
                });
            }
        });

        return tableData;
    } catch (error) {
        console.error(`Error processing file ${filePath}:`, error.message);
        throw error;
    }
};

// Convert Excel date serial numbers to JavaScript Date
const convertExcelDateToJSDate = (serial) => {
    const excelEpoch = new Date(1899, 11, 30);
    return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000).toISOString().split('T')[0]; // Returns 'YYYY-MM-DD'
};

// Process directory and combine all files data
const processDirectoryForTable = (directoryPath, yearFilter, monthFilter) => {
    try {
        const files = fs.readdirSync(directoryPath);
        const combinedTableData = [];

        files.forEach(file => {
            const filePath = path.join(directoryPath, file);
            if (fs.lstatSync(filePath).isFile() && file.endsWith('.xlsx')) {
                const fileTableData = processExcelFileForTable(filePath, yearFilter, monthFilter);
                combinedTableData.push(...fileTableData);
            }
        });

        return combinedTableData;
    } catch (error) {
        console.error(`Error processing directory ${directoryPath}:`, error.message);
        throw error;
    }
};

// API route to fetch aggregated data and table data by year and month
app.get('/fetch-table-data/:year/:month', async (req, res) => {
    try {
        const yearFilter = parseInt(req.params.year, 10);
        const monthFilter = req.params.month.toLowerCase();

        if (isNaN(yearFilter)) {
            return res.status(400).json({ message: 'Invalid year parameter' });
        }

        const monthMapping = {
            january: 1,
            february: 2,
            march: 3,
            april: 4,
            may: 5,
            june: 6,
            july: 7,
            august: 8,
            september: 9,
            october: 10,
            november: 11,
            december: 12
        };

        const monthNum = monthMapping[monthFilter];
        if (!monthNum) {
            return res.status(400).json({ message: 'Invalid month parameter' });
        }

        const lubricantsPath = path.join(__dirname, 'uploads/excel-files/lubricants');
        const petroleumPath = path.join(__dirname, 'uploads/excel-files/petroleum');

        // Process data for the table with year and month filter
        const lubricantsTableData = processDirectoryForTable(lubricantsPath, yearFilter, monthNum);
        const petroleumTableData = processDirectoryForTable(petroleumPath, yearFilter, monthNum);

        // Combine table data from lubricants and petroleum
        const combinedTableData = [...lubricantsTableData, ...petroleumTableData];

        res.status(200).json({
            year: yearFilter,
            month: monthNum,
            tableData: combinedTableData
        });
    } catch (error) {
        console.error('Error in /fetch-table-data/:year/:month:', error.message);
        res.status(500).json({ message: 'Internal Server Error', error: error.message });
    }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
