const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const csv = require('csv-parser');

const app = express();
const PORT = process.env.PORT || 3000;

// Set up EJS
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Middleware
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Ensure uploads directory exists
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir);
}

// Configure multer for file upload
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['.xls', '.xlsx', '.csv'];
        const ext = path.extname(file.originalname).toLowerCase();
        if (allowedTypes.includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Only .xls, .xlsx, and .csv files are allowed'));
        }
    }
});

// ============ SHUFFLE FUNCTION ============
// Fisher-Yates (Knuth) Shuffle Algorithm
function shuffleArray(array) {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
}

// Shuffle rows while keeping header intact
function shuffleRows(data) {
    if (!data || data.length === 0) return data;
    
    // Separate header (first row) and data rows
    const header = data[0];
    const rows = data.slice(1);
    
    // Shuffle the data rows
    const shuffledRows = shuffleArray(rows);
    
    // Combine header with shuffled rows
    return [header, ...shuffledRows];
}
// ==========================================

// Read CSV file
function readCSV(filePath) {
    return new Promise((resolve, reject) => {
        const results = [];
        fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', (data) => results.push(data))
            .on('end', () => {
                // Shuffle the rows (excluding header)
                const shuffledData = shuffleRows(results);
                resolve(shuffledData);
            })
            .on('error', (error) => reject(error));
    });
}

// Read Excel file
function readExcel(filePath) {
    try {
        const workbook = xlsx.readFile(filePath);
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(firstSheet, { header: 1 });
        
        // Shuffle the rows (excluding header)
        return shuffleRows(data);
    } catch (error) {
        console.error('Error reading Excel file:', error);
        throw error;
    }
}

// Get all uploaded files with metadata
function getUploadedFiles() {
    try {
        const files = fs.readdirSync(uploadDir);
        return files.map(file => {
            const filePath = path.join(uploadDir, file);
            const stats = fs.statSync(filePath);
            return {
                name: file,
                size: (stats.size / 1024).toFixed(2) + ' KB',
                modified: stats.mtime,
                extension: path.extname(file).toLowerCase()
            };
        }).sort((a, b) => b.modified - a.modified);
    } catch (error) {
        console.error('Error reading uploads directory:', error);
        return [];
    }
}

// Routes
app.get('/', async (req, res) => {
    try {
        const files = getUploadedFiles();
        // Add error: null to the render
        res.render('index', { 
            files, 
            selectedFile: null, 
            data: null, 
            shuffled: false,
            error: null  // ← Add this line
        });
    } catch (error) {
        console.error(error);
        res.status(500).send('Error loading files');
    }
});

app.get('/file/:filename', async (req, res) => {
    try {
        const filename = req.params.filename;
        const filePath = path.join(uploadDir, filename);
        
        if (!fs.existsSync(filePath)) {
            return res.status(404).send('File not found');
        }

        const ext = path.extname(filename).toLowerCase();
        let data = [];
        let error = null;
        let shuffled = false;
        
        try {
            if (ext === '.csv') {
                data = await readCSV(filePath);
                shuffled = true;
            } else if (ext === '.xlsx' || ext === '.xls') {
                data = readExcel(filePath);
                shuffled = true;
            } else {
                error = 'Unsupported file format';
            }
        } catch (readError) {
            error = 'Error reading file: ' + readError.message;
            console.error(readError);
        }

        const files = getUploadedFiles();
        res.render('index', { files, selectedFile: filename, data, error, shuffled });
    } catch (error) {
        console.error(error);
        res.status(500).send('Error reading file');
    }
});

// Add a route to reshuffle without reloading the page
app.get('/reshuffle/:filename', async (req, res) => {
    try {
        const filename = req.params.filename;
        const filePath = path.join(uploadDir, filename);
        
        if (!fs.existsSync(filePath)) {
            return res.status(404).json({ error: 'File not found' });
        }

        const ext = path.extname(filename).toLowerCase();
        let data = [];
        
        try {
            if (ext === '.csv') {
                data = await readCSV(filePath);
            } else if (ext === '.xlsx' || ext === '.xls') {
                data = readExcel(filePath);
            }
        } catch (readError) {
            return res.status(500).json({ error: readError.message });
        }

        res.json({ data, shuffled: true });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.get('/upload', (req, res) => {
    res.render('upload', { message: null, error: null });
});

app.post('/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.render('upload', { 
                message: null, 
                error: 'Please select a file to upload' 
            });
        }

        res.render('upload', { 
            message: `File "${req.file.originalname}" uploaded successfully!`, 
            error: null 
        });
    } catch (error) {
        res.render('upload', { 
            message: null, 
            error: error.message || 'Error uploading file' 
        });
    }
});

// Delete file route (optional)
app.delete('/file/:filename', (req, res) => {
    try {
        const filename = req.params.filename;
        const filePath = path.join(uploadDir, filename);
        
        if (fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
            res.json({ success: true, message: 'File deleted' });
        } else {
            res.status(404).json({ success: false, message: 'File not found' });
        }
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

// Error handling for multer
app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError) {
        res.render('upload', { message: null, error: error.message });
    } else {
        next(error);
    }
});

app.listen(PORT, () => {
    console.log(`🚀 Server running at http://localhost:${PORT}`);
    console.log(`📁 Upload directory: ${uploadDir}`);
    console.log(`🔧 Node version: ${process.version}`);
});