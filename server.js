const express = require('express');
const multer = require('multer');
const path = require('path');
const { parseExcel } = require('./excelParser');

const app = express();
const PORT = 3000;

// Multer setup for memory storage (we process buffer directly)
const upload = multer({ storage: multer.memoryStorage() });

// In-memory data store
let dataStore = {
    courses: {},
    slots: [],
    enrollments: {},
    processed: false
};

// Serve static files (HTML, CSS)
app.use(express.static(path.join(__dirname, 'public')));

// API: Upload Excel
app.post('/admin/upload', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        console.log(`Processing file: ${req.file.originalname}`);

        const result = parseExcel(req.file.buffer);

        // Update in-memory store
        dataStore = {
            courses: result.courses,
            slots: result.slots,
            enrollments: result.enrollments,
            processed: true
        };

        res.json({
            courses: result.counts.courses,
            slots: result.counts.slots,
            students: result.counts.students,
            errors: result.errors
        });

    } catch (error) {
        console.error('Error processing upload:', error);
        res.status(500).json({ error: 'Internal server error processing file.' });
    }
});

// API: Debug Data
app.get('/debug/data', (req, res) => {
    res.json(dataStore);
});

// Start Server
app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});
