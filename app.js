const express = require('express');
const path = require('path');
const ExcelJS = require('exceljs');
const multer = require('multer');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Multer for file uploads
const upload = multer({ dest: 'uploads/' }); // Temp folder for uploads

// Path to the Excel file
const excelPath = path.join(__dirname, 'timetable.xlsx');

// Helper function to read Excel data
async function readExcel() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);
    const worksheet = workbook.getWorksheet('timetable');
    const data = [];
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { // Skip header
            data.push({ 
                index: rowNumber - 1, 
                dateTime: row.getCell(1).value, 
                activity: row.getCell(2).value, 
                email: row.getCell(3).value 
            });
        }
    });
    return data;
}

// Helper function to write Excel data
async function writeExcel(data) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('timetable');
    worksheet.addRow(['DateTime', 'Activity', 'Email']); // Updated headers
    data.forEach(item => worksheet.addRow([item.dateTime, item.activity, item.email]));
    await workbook.xlsx.writeFile(excelPath);
}

// GET /recent: Returns JSON of last 2 reminders
app.get('/recent', async (req, res) => {
    try {
        const data = await readExcel();
        const recent = data.slice(-2);
        res.json(recent);
    } catch (error) {
        console.error('Error fetching recent:', error);
        res.status(500).json({ error: 'Failed to fetch recent reminders' });
    }
});

// POST /add-reminder: Add new reminder (now includes email)
app.post('/add-reminder', async (req, res) => {
    try {
        const { date, time, task, email } = req.body;
        if (!date || !time || !task || !email) {
            return res.status(400).send('Error: Missing required fields (date, time, task, email)');
        }
        const dateTime = `${date} ${time}`;
        const data = await readExcel();
        data.push({ dateTime, activity: task, email });
        await writeExcel(data);
        res.redirect('/');
    } catch (error) {
        console.error('Error saving:', error);
        res.status(500).send('Error: Failed to save reminder');
    }
});

// GET /: Serve index.html
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// GET /show: Serve show.html
app.get('/show', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'show.html'));
});

// POST /delete/:index: Delete reminder
app.post('/delete/:index', async (req, res) => {
    try {
        const index = parseInt(req.params.index) - 1;
        const data = await readExcel();
        if (index < 0 || index >= data.length) {
            return res.status(400).send('Error: Invalid index');
        }
        data.splice(index, 1);
        await writeExcel(data);
        res.redirect('/show');
    } catch (error) {
        console.error('Error deleting:', error);
        res.status(500).send('Error: Failed to delete reminder');
    }
});

// GET /update/:index: Serve update.html
app.get('/update/:index', async (req, res) => {
    try {
        const index = parseInt(req.params.index) - 1;
        const data = await readExcel();
        if (index < 0 || index >= data.length) {
            return res.status(400).send('Error: Invalid index');
        }
        const item = data[index];
        res.redirect(`/update.html?index=${req.params.index}&dateTime=${encodeURIComponent(item.dateTime)}&activity=${encodeURIComponent(item.activity)}&email=${encodeURIComponent(item.email)}`);
    } catch (error) {
        console.error('Error loading update:', error);
        res.status(500).send('Error: Failed to load update form');
    }
});

// POST /update/:index: Update reminder (now includes email)
app.post('/update/:index', async (req, res) => {
    try {
        const index = parseInt(req.params.index) - 1;
        const { dateTime, activity, email } = req.body;
        const data = await readExcel();
        if (index < 0 || index >= data.length) {
            return res.status(400).send('Error: Invalid index');
        }
        data[index] = { dateTime, activity, email };
        await writeExcel(data);
        res.redirect('/show');
    } catch (error) {
        console.error('Error updating:', error);
        res.status(500).send('Error: Failed to update reminder');
    }
});

// POST /upload: Upload and append data from XLSX file
app.post('/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).send('Error: No file uploaded');
        }
        const uploadPath = path.join(__dirname, req.file.path);
        const uploadWorkbook = new ExcelJS.Workbook();
        await uploadWorkbook.xlsx.readFile(uploadPath);
        const uploadWorksheet = uploadWorkbook.getWorksheet(1); // Assume first sheet
        const uploadData = [];
        uploadWorksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Skip header
                uploadData.push({
                    dateTime: row.getCell(1).value,
                    activity: row.getCell(2).value,
                    email: row.getCell(3).value
                });
            }
        });
        // Append to existing data
        const existingData = await readExcel();
        const combinedData = existingData.concat(uploadData);
        await writeExcel(combinedData);
        // Clean up uploaded file
        require('fs').unlinkSync(uploadPath);
        res.redirect('/');
    } catch (error) {
        console.error('Error uploading:', error);
        res.status(500).send('Error: Failed to upload and append data');
    }
});

// GET /all: Returns JSON of all reminders
app.get('/all', async (req, res) => {
    try {
        const data = await readExcel();
        res.json(data);
    } catch (error) {
        console.error('Error fetching all:', error);
        res.status(500).json({ error: 'Failed to fetch all reminders' });
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});