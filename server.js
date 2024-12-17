const express = require('express');
const multer = require('multer');
const path = require('path');
const cors = require('cors');
const fs = require('fs');

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

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
