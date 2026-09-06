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

// ============ FILE SYSTEM HELPERS ============

// Recursively get all files and folders
function getDirectoryStructure(dir, baseDir = '') {
    const items = [];
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    
    for (const entry of entries) {
        const fullPath = path.join(dir, entry.name);
        // Use forward slashes for web paths
        const relativePath = baseDir ? path.join(baseDir, entry.name).replace(/\\/g, '/') : entry.name;
        
        if (entry.isDirectory()) {
            const subItems = getDirectoryStructure(fullPath, relativePath);
            items.push({
                type: 'folder',
                name: entry.name,
                path: relativePath,
                children: subItems
            });
        } else if (entry.isFile()) {
            const ext = path.extname(entry.name).toLowerCase();
            if (['.xls', '.xlsx', '.csv'].includes(ext)) {
                const stats = fs.statSync(fullPath);
                items.push({
                    type: 'file',
                    name: entry.name,
                    path: relativePath,
                    fullPath: fullPath,
                    size: (stats.size / 1024).toFixed(2) + ' KB',
                    modified: stats.mtime,
                    extension: ext
                });
            }
        }
    }
    
    items.sort((a, b) => {
        if (a.type === 'folder' && b.type !== 'folder') return -1;
        if (a.type !== 'folder' && b.type === 'folder') return 1;
        return a.name.localeCompare(b.name);
    });
    
    return items;
}

// Get all files (flat list)
function getAllFiles(dir, baseDir = '') {
    const files = [];
    const entries = fs.readdirSync(dir, { withFileTypes: true });

    for (const entry of entries) {
        const fullPath = path.join(dir, entry.name);
        const relativePath = baseDir ? path.join(baseDir, entry.name) : entry.name;

        if (entry.isDirectory()) {
            const subFiles = getAllFiles(fullPath, relativePath);
            files.push(...subFiles);
        } else if (entry.isFile()) {
            const ext = path.extname(entry.name).toLowerCase();
            if (['.xls', '.xlsx', '.csv'].includes(ext)) {
                const stats = fs.statSync(fullPath);
                files.push({
                    name: entry.name,
                    path: relativePath,
                    fullPath: fullPath,
                    size: (stats.size / 1024).toFixed(2) + ' KB',
                    modified: stats.mtime,
                    extension: ext
                });
            }
        }
    }

    return files;
}

// ============ SHUFFLE FUNCTION ============
function shuffleArray(array) {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
}

function shuffleRows(data) {
    if (!data || data.length === 0) return data;
    const header = data[0];
    const rows = data.slice(1);
    const shuffledRows = shuffleArray(rows);
    return [header, ...shuffledRows];
}

// ============ FILE READERS ============
function readCSV(filePath) {
    return new Promise((resolve, reject) => {
        const results = [];
        fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', (data) => results.push(data))
            .on('end', () => {
                const shuffledData = shuffleRows(results);
                resolve(shuffledData);
            })
            .on('error', (error) => reject(error));
    });
}

function readExcel(filePath) {
    try {
        const workbook = xlsx.readFile(filePath);
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Read as array of arrays
        const rawData = xlsx.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
        
        if (!rawData || rawData.length === 0) {
            return [];
        }
        
        // Get headers from first row
        const headers = rawData[0];
        // Get data rows (skip header)
        const rows = rawData.slice(1);
        
        // Convert to object format like CSV
        const result = rows.map(row => {
            const obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index] || '';
            });
            return obj;
        });
        
        // Shuffle the rows
        const shuffledResult = shuffleArray(result);
        
        // Return with headers as first row in object format
        return shuffledResult;
    } catch (error) {
        console.error('Error reading Excel file:', error);
        throw error;
    }
}

// ============ MULTER CONFIG ============
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        let uploadPath = uploadDir;
        const folderPath = req.body.folderPath || '';

        if (folderPath) {
            uploadPath = path.join(uploadDir, folderPath);
            if (!fs.existsSync(uploadPath)) {
                fs.mkdirSync(uploadPath, { recursive: true });
            }
        }
        cb(null, uploadPath);
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

// ============ ROUTES ============
app.get('/', async (req, res) => {
    try {
        const structure = getDirectoryStructure(uploadDir);
        res.render('index', {
            structure,
            selectedFile: null,
            data: null,
            shuffled: false,
            error: null
        });
    } catch (error) {
        console.error(error);
        res.status(500).send('Error loading files');
    }
});

// Fixed routes - using query parameter instead of path parameter with wildcard
app.get('/file', async (req, res) => {
    try {
        let filePath = req.query.path;
        if (!filePath) {
            return res.status(400).send('No file path specified');
        }
        
        // Convert backslashes to forward slashes
        filePath = filePath.replace(/\\/g, '/');
        const fullPath = path.join(uploadDir, filePath);
        
        if (!fs.existsSync(fullPath)) {
            return res.status(404).send('File not found');
        }

        const ext = path.extname(filePath).toLowerCase();
        let data = [];
        let error = null;
        let shuffled = false;
        
        try {
            if (ext === '.csv') {
                data = await readCSV(fullPath);
                shuffled = true;
            } else if (ext === '.xlsx' || ext === '.xls') {
                data = readExcel(fullPath);
                shuffled = true;
            } else {
                error = 'Unsupported file format';
            }
            console.log('File data sample:', data.slice(0, 3));
        } catch (readError) {
            error = 'Error reading file: ' + readError.message;
            console.error(readError);
        }

        const structure = getDirectoryStructure(uploadDir);
        res.render('index', { 
            structure, 
            selectedFile: filePath, 
            data, 
            error, 
            shuffled 
        });
    } catch (error) {
        console.error(error);
        res.status(500).send('Error reading file');
    }
    
});

app.get('/reshuffle', async (req, res) => {
    try {
        let filePath = req.query.path;
        if (!filePath) {
            return res.status(400).json({ error: 'No file path specified' });
        }
        
        // Convert backslashes to forward slashes (for Windows compatibility)
        filePath = filePath.replace(/\\/g, '/');
        const fullPath = path.join(uploadDir, filePath);
        
        // Log for debugging
        console.log('Reshuffle request for:', filePath);
        console.log('Full path:', fullPath);
        
        // Check if file exists
        if (!fs.existsSync(fullPath)) {
            return res.status(404).json({ error: 'File not found: ' + filePath });
        }

        const ext = path.extname(filePath).toLowerCase();
        let data = [];
        
        try {
            if (ext === '.csv') {
                data = await readCSV(fullPath);
            } else if (ext === '.xlsx' || ext === '.xls') {
                data = readExcel(fullPath);
            } else {
                return res.status(400).json({ error: 'Unsupported file format' });
            }
        } catch (readError) {
            return res.status(500).json({ error: readError.message });
        }

        res.json({ data, shuffled: true });
    } catch (error) {
        console.error('Reshuffle error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.get('/upload', (req, res) => {
    res.render('upload', { message: null, error: null });
});

app.delete('/file', (req, res) => {
    try {
        let filePath = req.query.path;
        if (!filePath) {
            return res.status(400).json({ success: false, message: 'No file path specified' });
        }
        
        // Convert backslashes to forward slashes (for Windows compatibility)
        filePath = filePath.replace(/\\/g, '/');
        const fullPath = path.join(uploadDir, filePath);
        
        if (fs.existsSync(fullPath)) {
            fs.unlinkSync(fullPath);
            res.json({ success: true, message: 'File deleted' });
        } else {
            res.status(404).json({ success: false, message: 'File not found' });
        }
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.render('upload', {
                message: null,
                error: 'Please select a file to upload'
            });
        }

        const folderPath = req.body.folderPath || '';
        const displayPath = folderPath ? `in folder "${folderPath}"` : 'in root folder';

        res.render('upload', {
            message: `File "${req.file.originalname}" uploaded successfully ${displayPath}!`,
            error: null
        });
    } catch (error) {
        res.render('upload', {
            message: null,
            error: error.message || 'Error uploading file'
        });
    }
});

// Delete file route
app.get('/file', async (req, res) => {
    try {
        let filePath = req.query.path;
        if (!filePath) {
            return res.status(400).send('No file path specified');
        }
        
        // Convert backslashes to forward slashes (for Windows compatibility)
        filePath = filePath.replace(/\\/g, '/');
        
        const fullPath = path.join(uploadDir, filePath);
        
        if (!fs.existsSync(fullPath)) {
            return res.status(404).send('File not found');
        }

        const ext = path.extname(filePath).toLowerCase();
        let data = [];
        let error = null;
        let shuffled = false;
        
        try {
            if (ext === '.csv') {
                data = await readCSV(fullPath);
                shuffled = true;
            } else if (ext === '.xlsx' || ext === '.xls') {
                data = readExcel(fullPath);
                shuffled = true;
            } else {
                error = 'Unsupported file format';
            }
        } catch (readError) {
            error = 'Error reading file: ' + readError.message;
            console.error(readError);
        }

        const structure = getDirectoryStructure(uploadDir);
        res.render('index', { 
            structure, 
            selectedFile: filePath, 
            data, 
            error, 
            shuffled 
        });
    } catch (error) {
        console.error(error);
        res.status(500).send('Error reading file');
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