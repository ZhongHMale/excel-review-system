const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const sqlite3 = require('sqlite3').verbose();
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');
const { marked } = require('marked');
const createDOMPurify = require('dompurify');
const { JSDOM } = require('jsdom');
const moment = require('moment');

const app = express();
const PORT = process.env.PORT || 3001;

// 创建DOMPurify实例
const window = new JSDOM('').window;
const DOMPurify = createDOMPurify(window);

// 中间件
// 中间件
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// 设置默认字符编码
app.use((req, res, next) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    next();
});

// 创建上传目录
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

// 配置multer
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadsDir);
    },
    filename: (req, file, cb) => {
        try {
            // 安全的文件名处理
            let originalName = file.originalname;
            
            // 尝试不同的编码方式
            try {
                // 如果是乱码，尝试从latin1转utf8
                if (originalName.includes('�') || /[\x80-\xFF]/.test(originalName)) {
                    originalName = Buffer.from(originalName, 'latin1').toString('utf8');
                }
            } catch (e) {
                console.log('编码转换失败，使用原始文件名');
            }
            
            // 清理文件名，移除不安全的字符
            const cleanName = originalName
                .replace(/[<>:"/\\|?*]/g, '_')  // 替换Windows不支持的字符
                .replace(/[\x00-\x1f\x80-\x9f]/g, '_')  // 替换控制字符
                .replace(/^\.|\.$/, '_')  // 处理以点开头或结尾的文件名
                .substring(0, 100);  // 限制文件名长度
            
            // 确保文件扩展名正确
            const ext = path.extname(cleanName) || '.xlsx';
            const nameWithoutExt = path.basename(cleanName, ext);
            
            const uniqueName = `${Date.now()}-${uuidv4()}-${nameWithoutExt}${ext}`;
            
            //console.log('原始文件名:', file.originalname);
            //console.log('清理后文件名:', cleanName);
            //console.log('最终文件名:', uniqueName);
            
            cb(null, uniqueName);
        } catch (error) {
            console.error('文件名处理错误:', error);
            // 如果处理失败，使用安全的默认名称
            const ext = '.xlsx';
            const safeName = `${Date.now()}-${uuidv4()}${ext}`;
            cb(null, safeName);
        }
    }
});

const upload = multer({ 
    storage: storage,
    fileFilter: (req, file, cb) => {
        try {
            // 安全的文件名编码处理
            let originalName = file.originalname;
            
            // 尝试编码转换
            try {
                if (originalName.includes('�') || /[\x80-\xFF]/.test(originalName)) {
                    originalName = Buffer.from(originalName, 'latin1').toString('utf8');
                }
            } catch (e) {
                console.log('fileFilter编码转换失败');
            }
            
            // 更新文件对象的原始名称
            file.originalname = originalName;
            
            const allowedTypes = ['.xlsx', '.xls', '.csv'];
            const ext = path.extname(originalName).toLowerCase();
            
            if (allowedTypes.includes(ext)) {
                cb(null, true);
            } else {
                cb(new Error('只支持Excel和CSV文件'));
            }
        } catch (error) {
            console.error('fileFilter错误:', error);
            cb(new Error('文件处理失败'));
        }
    },
    limits: {
        fileSize: 50 * 1024 * 1024 // 50MB限制
    }
});

// 初始化数据库
// 初始化数据库
const db = new sqlite3.Database('./excel_system.db', (err) => {
    if (err) {
        console.error('数据库连接失败:', err);
    } else {
        console.log('数据库连接成功');
        // 设置数据库编码为UTF-8
        db.run("PRAGMA encoding = 'UTF-8'");
        db.run("PRAGMA journal_mode = WAL");
        db.run("PRAGMA synchronous = NORMAL");
        // 确保文本排序规则支持UTF-8
        db.run("PRAGMA case_sensitive_like = true");
    }
});

// 创建表
db.serialize(() => {
    // 文件主表 - 管理不同的文件
    db.run(`CREATE TABLE IF NOT EXISTS files (
        id TEXT PRIMARY KEY,
        original_name TEXT NOT NULL,
        file_hash TEXT,
        created_by TEXT NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(original_name)
    )`);

    // 文件版本表 - 每个文件的不同版本
    db.run(`CREATE TABLE IF NOT EXISTS file_versions (
        id TEXT PRIMARY KEY,
        file_id TEXT NOT NULL,
        version INTEGER NOT NULL,
        filename TEXT NOT NULL,
        file_path TEXT NOT NULL,
        submitter TEXT NOT NULL,
        change_description TEXT,
        file_size INTEGER,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (file_id) REFERENCES files (id),
        UNIQUE(file_id, version)
    )`);

    // 表格数据表
    db.run(`CREATE TABLE IF NOT EXISTS table_data (
        id TEXT PRIMARY KEY,
        version_id TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        column_name TEXT NOT NULL,
        cell_value TEXT,
        cell_type TEXT DEFAULT 'text',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (version_id) REFERENCES file_versions (id)
    )`);

    // 批注表
    db.run(`CREATE TABLE IF NOT EXISTS annotations (
        id TEXT PRIMARY KEY,
        version_id TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        column_name TEXT NOT NULL,
        annotation_type TEXT NOT NULL,
        annotation_data TEXT,
        created_by TEXT NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (version_id) REFERENCES file_versions (id)
    )`);

    // 自定义列表
    db.run(`CREATE TABLE IF NOT EXISTS custom_columns (
        id TEXT PRIMARY KEY,
        version_id TEXT NOT NULL,
        column_name TEXT NOT NULL,
        column_type TEXT DEFAULT 'text',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (version_id) REFERENCES file_versions (id)
    )`);
});

// 工具函数：解析Excel文件
function parseExcelFile(filePath) {
    // 添加编码配置，支持中文字符
    const workbook = XLSX.readFile(filePath, {
        codepage: 65001, // UTF-8编码，支持中文字符
        cellText: false,
        cellDates: true
    });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1, 
        defval: '',
        raw: false // 确保文本按字符串处理
    });
    
    if (jsonData.length === 0) return { headers: [], data: [] };
    
    const headers = jsonData[0];
    const data = jsonData.slice(1).map((row, index) => {
        const rowData = {};
        headers.forEach((header, colIndex) => {
            rowData[header] = row[colIndex] || '';
        });
        rowData._rowIndex = index;
        return rowData;
    });
    
    return { headers, data };
}

// 工具函数：渲染Markdown表格
function renderMarkdownTable(content) {
    if (!content || typeof content !== 'string') return content;
    
    // 检查是否包含Markdown表格
    const tableRegex = /\|.*\|/g;
    if (!tableRegex.test(content)) return content;
    
    try {
        const html = marked(content);
        return DOMPurify.sanitize(html);
    } catch (error) {
        console.error('Markdown渲染错误:', error);
        return content;
    }
}

// API路由

// 上传Excel文件
const crypto = require('crypto');

// 计算文件哈希
function calculateFileHash(filePath) {
    const fileBuffer = fs.readFileSync(filePath);
    const hashSum = crypto.createHash('md5');
    hashSum.update(fileBuffer);
    return hashSum.digest('hex');
}

// 上传Excel文件
app.post('/api/upload', upload.single('file'), (req, res) => {
    try {
        const { submitter, changeDescription } = req.body;
        const file = req.file;
        
        if (!file) {
            return res.status(400).json({ error: '没有上传文件' });
        }

        const { headers, data } = parseExcelFile(file.path);
        const fileHash = calculateFileHash(file.path);
        const originalName = file.originalname;

        // 首先检查文件是否已存在
        db.get(
            'SELECT id FROM files WHERE original_name = ?',
            [originalName],
            (err, existingFile) => {
                if (err) {
                    console.error(err);
                    return res.status(500).json({ error: '数据库错误' });
                }

                let fileId;
                
                const processFileVersion = (fileId) => {
                    // 获取该文件的最新版本号
                    db.get(
                        'SELECT MAX(version) as max_version FROM file_versions WHERE file_id = ?',
                        [fileId],
                        (err, row) => {
                            if (err) {
                                console.error(err);
                                return res.status(500).json({ error: '获取版本信息失败' });
                            }

                            const version = (row.max_version || 0) + 1;
                            const versionId = uuidv4();

                            // 插入新版本记录
                            db.run(
                                `INSERT INTO file_versions (id, file_id, version, filename, file_path, submitter, change_description, file_size)
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
                                [versionId, fileId, version, file.filename, file.path, submitter, changeDescription, file.size],
                                function(err) {
                                    if (err) {
                                        console.error(err);
                                        return res.status(500).json({ error: '保存版本信息失败' });
                                    }

                                    // 插入表格数据
                                    const stmt = db.prepare(`
                                        INSERT INTO table_data (id, version_id, row_index, column_name, cell_value, cell_type)
                                        VALUES (?, ?, ?, ?, ?, ?)
                                    `);

                                    data.forEach((row, rowIndex) => {
                                        headers.forEach(header => {
                                            const cellId = uuidv4();
                                            const cellValue = row[header];
                                            const cellType = typeof cellValue === 'number' ? 'number' : 'text';
                                            
                                            stmt.run([cellId, versionId, rowIndex, header, cellValue, cellType]);
                                        });
                                    });

                                    stmt.finalize();

                                    // 更新文件的最后修改时间
                                    db.run(
                                        'UPDATE files SET updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                                        [fileId]
                                    );

                                    res.json({
                                        success: true,
                                        fileId,
                                        versionId,
                                        version,
                                        headers,
                                        dataCount: data.length,
                                        message: `文件 "${originalName}" 的第 ${version} 版本上传成功`
                                    });
                                }
                            );
                        }
                    );
                };

                if (existingFile) {
                    // 文件已存在，添加新版本
                    fileId = existingFile.id;
                    processFileVersion(fileId);
                } else {
                    // 新文件，创建文件记录
                    fileId = uuidv4();
                    db.run(
                        `INSERT INTO files (id, original_name, file_hash, created_by)
                         VALUES (?, ?, ?, ?)`,
                        [fileId, originalName, fileHash, submitter],
                        function(err) {
                            if (err) {
                                console.error(err);
                                return res.status(500).json({ error: '创建文件记录失败' });
                            }
                            processFileVersion(fileId);
                        }
                    );
                }
            }
        );
    } catch (error) {
        console.error('上传错误:', error);
        res.status(500).json({ error: '文件处理失败' });
    }
});
// 获取所有文件列表
app.get('/api/files', (req, res) => {
    const query = `
        SELECT 
            f.id,
            f.original_name,
            f.created_by,
            f.created_at,
            f.updated_at,
            COUNT(fv.id) as version_count,
            MAX(fv.version) as latest_version,
            MAX(fv.created_at) as last_modified
        FROM files f
        LEFT JOIN file_versions fv ON f.id = fv.file_id
        GROUP BY f.id, f.original_name, f.created_by, f.created_at, f.updated_at
        ORDER BY f.updated_at DESC
    `;
    
    db.all(query, (err, rows) => {
        if (err) {
            console.error(err);
            return res.status(500).json({ error: '获取文件列表失败' });
        }
        res.json(rows);
    });
});

// 获取特定文件的版本列表
app.get('/api/files/:fileId/versions', (req, res) => {
    const { fileId } = req.params;
    
    db.all(
        `SELECT fv.*, f.original_name
         FROM file_versions fv
         JOIN files f ON fv.file_id = f.id
         WHERE fv.file_id = ?
         ORDER BY fv.version DESC`,
        [fileId],
        (err, rows) => {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '获取版本列表失败' });
            }
            res.json(rows);
        }
    );
});

// 获取特定版本的数据
app.get('/api/versions/:versionId/data', (req, res) => {
    const { versionId } = req.params;
    const { columns } = req.query;
    
    // 获取版本信息
    db.get(
        `SELECT fv.*, f.original_name
         FROM file_versions fv
         JOIN files f ON fv.file_id = f.id
         WHERE fv.id = ?`,
        [versionId],
        (err, versionInfo) => {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '获取版本信息失败' });
            }
            
            if (!versionInfo) {
                return res.status(404).json({ error: '版本不存在' });
            }

            // 获取表格数据
            let query = `
                SELECT row_index, column_name, cell_value, cell_type
                FROM table_data 
                WHERE version_id = ?
            `;
            
            if (columns) {
                const columnList = columns.split(',').map(col => `'${col}'`).join(',');
                query += ` AND column_name IN (${columnList})`;
            }
            
            query += ' ORDER BY row_index, column_name';

            db.all(query, [versionId], (err, rows) => {
                if (err) {
                    console.error(err);
                    return res.status(500).json({ error: '获取数据失败' });
                }

                // 重组数据
                const dataMap = {};
                const headers = new Set();
                
                rows.forEach(row => {
                    if (!dataMap[row.row_index]) {
                        dataMap[row.row_index] = { _rowIndex: row.row_index };
                    }
                    
                    // 处理Markdown内容
                    let cellValue = row.cell_value;
                    if (row.cell_type === 'text' && cellValue) {
                        cellValue = renderMarkdownTable(cellValue);
                    }
                    
                    dataMap[row.row_index][row.column_name] = {
                        value: cellValue,
                        type: row.cell_type
                    };
                    headers.add(row.column_name);
                });

                const data = Object.values(dataMap);
                
                // 获取批注信息
                db.all(
                    'SELECT * FROM annotations WHERE version_id = ?',
                    [versionId],
                    (err, annotations) => {
                        if (err) {
                            console.error(err);
                            return res.status(500).json({ error: '获取批注失败' });
                        }

                        res.json({
                            versionInfo,
                            headers: Array.from(headers),
                            data,
                            annotations
                        });
                    }
                );
            });
        }
    );
});

// 更新单元格数据
app.put('/api/files/:fileId/cell', (req, res) => {
    const { fileId } = req.params;
    const { rowIndex, columnName, value, cellType = 'text' } = req.body;

    db.run(
        `UPDATE table_data 
         SET cell_value = ?, cell_type = ?, updated_at = CURRENT_TIMESTAMP
         WHERE file_version_id = ? AND row_index = ? AND column_name = ?`,
        [value, cellType, fileId, rowIndex, columnName],
        function(err) {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '更新失败' });
            }
            
            if (this.changes === 0) {
                return res.status(404).json({ error: '单元格不存在' });
            }
            
            res.json({ success: true, message: '更新成功' });
        }
    );
});

// 添加批注
app.post('/api/files/:fileId/annotation', (req, res) => {
    const { fileId } = req.params;
    const { rowIndex, columnName, annotationType, annotationData, createdBy } = req.body;

    const annotationId = uuidv4();
    
    db.run(
        `INSERT INTO annotations (id, file_version_id, row_index, column_name, annotation_type, annotation_data, created_by)
         VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [annotationId, fileId, rowIndex, columnName, annotationType, JSON.stringify(annotationData), createdBy],
        function(err) {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '添加批注失败' });
            }
            
            res.json({ 
                success: true, 
                annotationId,
                message: '批注添加成功' 
            });
        }
    );
});

// 添加自定义列
app.post('/api/files/:fileId/column', (req, res) => {
    const { fileId } = req.params;
    const { columnName, columnType = 'text' } = req.body;

    const columnId = uuidv4();
    
    db.run(
        `INSERT INTO custom_columns (id, file_version_id, column_name, column_type)
         VALUES (?, ?, ?, ?)`,
        [columnId, fileId, columnName, columnType],
        function(err) {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '添加列失败' });
            }

            // 为所有现有行添加新列的空数据
            db.all(
                'SELECT DISTINCT row_index FROM table_data WHERE file_version_id = ?',
                [fileId],
                (err, rows) => {
                    if (err) {
                        console.error();
                        return res.status(500).json({ error: '获取行数据失败' });
                    }

                    const stmt = db.prepare(`
                        INSERT INTO table_data (id, file_version_id, row_index, column_name, cell_value, cell_type)
                        VALUES (?, ?, ?, ?, ?, ?)
                    `);

                    rows.forEach(row => {
                        const cellId = uuidv4();
                        stmt.run([cellId, fileId, row.row_index, columnName, '', columnType]);
                    });

                    stmt.finalize();

                    res.json({ 
                        success: true, 
                        columnId,
                        message: '列添加成功' 
                    });
                }
            );
        }
    );
});

// 导出Excel
app.get('/api/files/:fileId/export', (req, res) => {
    const { fileId } = req.params;
    
    db.all(
        `SELECT row_index, column_name, cell_value
         FROM table_data 
         WHERE file_version_id = ?
         ORDER BY row_index, column_name`,
        [fileId],
        (err, rows) => {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '导出失败' });
            }

            // 重组数据为Excel格式
            const dataMap = {};
            const headers = new Set();
            
            rows.forEach(row => {
                if (!dataMap[row.row_index]) {
                    dataMap[row.row_index] = {};
                }
                dataMap[row.row_index][row.column_name] = row.cell_value;
                headers.add(row.column_name);
            });

            const headerArray = Array.from(headers);
            const excelData = [headerArray];
            
            Object.keys(dataMap).sort((a, b) => parseInt(a) - parseInt(b)).forEach(rowIndex => {
                const row = headerArray.map(header => dataMap[rowIndex][header] || '');
                excelData.push(row);
            });

            // 创建工作簿
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(excelData);
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

            // 生成文件
            const exportPath = path.join(uploadsDir, `export-${fileId}-${Date.now()}.xlsx`);
            XLSX.writeFile(wb, exportPath);

            // 设置正确的响应头
            const filename = `export-${moment().format('YYYY-MM-DD-HH-mm-ss')}.xlsx`;
            res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            
            res.download(exportPath, filename, (err) => {
                if (err) {
                    console.error('下载错误:', err);
                }
                // 删除临时文件
                fs.unlink(exportPath, () => {});
            });
        }
    );
});


// 更新单元格数据API (使用versionId)
app.put('/api/versions/:versionId/cell', (req, res) => {
    const { versionId } = req.params;
    const { rowIndex, columnName, value, cellType = 'text' } = req.body;

    db.run(
        `UPDATE table_data 
         SET cell_value = ?, cell_type = ?, updated_at = CURRENT_TIMESTAMP
         WHERE version_id = ? AND row_index = ? AND column_name = ?`,
        [value, cellType, versionId, rowIndex, columnName],
        function(err) {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '更新失败' });
            }
            
            if (this.changes === 0) {
                return res.status(404).json({ error: '单元格不存在' });
            }
            
            res.json({ success: true, message: '更新成功' });
        }
    );
});

// 添加自定义列API (使用versionId)
app.post('/api/versions/:versionId/column', (req, res) => {
    const { versionId } = req.params;
    const { columnName, columnType = 'text' } = req.body;

    const columnId = uuidv4();
    
    db.run(
        `INSERT INTO custom_columns (id, version_id, column_name, column_type)
         VALUES (?, ?, ?, ?)`,
        [columnId, versionId, columnName, columnType],
        function(err) {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '添加列失败' });
            }

            // 为所有现有行添加新列的空数据
            db.all(
                'SELECT DISTINCT row_index FROM table_data WHERE version_id = ?',
                [versionId],
                (err, rows) => {
                    if (err) {
                        console.error(err);
                        return res.status(500).json({ error: '获取行数据失败' });
                    }

                    const stmt = db.prepare(`
                        INSERT INTO table_data (id, version_id, row_index, column_name, cell_value, cell_type)
                        VALUES (?, ?, ?, ?, ?, ?)
                    `);

                    rows.forEach(row => {
                        const cellId = uuidv4();
                        stmt.run([cellId, versionId, row.row_index, columnName, '', columnType]);
                    });

                    stmt.finalize();

                    res.json({ 
                        success: true,
                        columnId,
                        message: '列添加成功' 
                    });
                }
            );
        }
    );
});


// 保存当前版本的修改为新版本
app.post('/api/versions/:versionId/save-new-version', (req, res) => {
    const { versionId } = req.params;
    const { submitter, changeDescription } = req.body;

    // 首先获取当前版本信息
    db.get(
        `SELECT fv.*, f.id as file_id, f.original_name
         FROM file_versions fv
         JOIN files f ON fv.file_id = f.id
         WHERE fv.id = ?`,
        [versionId],
        (err, currentVersion) => {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '获取版本信息失败' });
            }

            if (!currentVersion) {
                return res.status(404).json({ error: '版本不存在' });
            }

            // 获取该文件的最新版本号
            db.get(
                'SELECT MAX(version) as max_version FROM file_versions WHERE file_id = ?',
                [currentVersion.file_id],
                (err, row) => {
                    if (err) {
                        console.error(err);
                        return res.status(500).json({ error: '获取版本信息失败' });
                    }

                    const newVersion = (row.max_version || 0) + 1;
                    const newVersionId = uuidv4();

                    // 创建新版本记录
                    db.run(
                        `INSERT INTO file_versions (id, file_id, version, filename, file_path, submitter, change_description, file_size, created_at)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`,
                        [newVersionId, currentVersion.file_id, newVersion, currentVersion.filename, currentVersion.file_path, submitter, changeDescription, currentVersion.file_size],
                        function(err) {
                            if (err) {
                                console.error(err);
                                return res.status(500).json({ error: '创建新版本失败' });
                            }

                            // 复制当前版本的所有数据到新版本
                            db.all(
                                'SELECT row_index, column_name, cell_value, cell_type FROM table_data WHERE version_id = ?',
                                [versionId],
                                (err, tableData) => {
                                    if (err) {
                                        console.error(err);
                                        return res.status(500).json({ error: '获取表格数据失败' });
                                    }

                                    // 批量插入新版本的数据
                                    const stmt = db.prepare(`
                                        INSERT INTO table_data (id, version_id, row_index, column_name, cell_value, cell_type)
                                        VALUES (?, ?, ?, ?, ?, ?)
                                    `);

                                    tableData.forEach(row => {
                                        const cellId = uuidv4();
                                        stmt.run([cellId, newVersionId, row.row_index, row.column_name, row.cell_value, row.cell_type]);
                                    });

                                    stmt.finalize();

                                    // 复制批注数据
                                    db.all(
                                        'SELECT row_index, column_name, annotation_type, annotation_data, created_by FROM annotations WHERE version_id = ?',
                                        [versionId],
                                        (err, annotations) => {
                                            if (err) {
                                                console.error(err);
                                                // 即使批注复制失败，也返回成功，因为主要数据已经保存
                                            } else {
                                                const annotationStmt = db.prepare(`
                                                    INSERT INTO annotations (id, version_id, row_index, column_name, annotation_type, annotation_data, created_by, created_at)
                                                    VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                                                `);

                                                annotations.forEach(annotation => {
                                                    const annotationId = uuidv4();
                                                    annotationStmt.run([annotationId, newVersionId, annotation.row_index, annotation.column_name, annotation.annotation_type, annotation.annotation_data, annotation.created_by]);
                                                });

                                                annotationStmt.finalize();
                                            }

                                            // 更新文件的最后修改时间
                                            db.run(
                                                'UPDATE files SET updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                                                [currentVersion.file_id]
                                            );

                                            res.json({
                                                success: true,
                                                versionId: newVersionId,
                                                version: newVersion,
                                                fileId: currentVersion.file_id,
                                                message: `已保存为新版本 v${newVersion}`
                                            });
                                        }
                                    );
                                }
                            );
                        }
                    );
                }
            );
        }
    );
});

// 启动服务器
app.listen(PORT, () => {
    console.log(`服务器运行在端口 ${PORT}`);
});


// 版本比较API
app.get('/api/versions/compare/:version1/:version2', (req, res) => {
    const { version1, version2 } = req.params;
    
    // 获取两个版本的数据
    const getVersionData = (versionId) => {
        return new Promise((resolve, reject) => {
            db.all(
                'SELECT row_index, column_name, cell_value, cell_type FROM table_data WHERE version_id = ? ORDER BY row_index, column_name',
                [versionId],
                (err, rows) => {
                    if (err) {
                        reject(err);
                    } else {
                        const dataMap = {};
                        rows.forEach(row => {
                            const key = `${row.row_index}-${row.column_name}`;
                            dataMap[key] = {
                                value: row.cell_value,
                                type: row.cell_type,
                                rowIndex: row.row_index,
                                columnName: row.column_name
                            };
                        });
                        resolve(dataMap);
                    }
                }
            );
        });
    };
    
    Promise.all([getVersionData(version1), getVersionData(version2)])
        .then(([data1, data2]) => {
            const changes = [];
            const allKeys = new Set([...Object.keys(data1), ...Object.keys(data2)]);
            
            allKeys.forEach(key => {
                const cell1 = data1[key];
                const cell2 = data2[key];
                
                if (!cell1 && cell2) {
                    // 新增的单元格
                    changes.push({
                        rowIndex: cell2.rowIndex,
                        columnName: cell2.columnName,
                        changeType: 'added',
                        oldValue: '',
                        newValue: cell2.value,
                        position: key
                    });
                } else if (cell1 && !cell2) {
                    // 删除的单元格
                    changes.push({
                        rowIndex: cell1.rowIndex,
                        columnName: cell1.columnName,
                        changeType: 'deleted',
                        oldValue: cell1.value,
                        newValue: '',
                        position: key
                    });
                } else if (cell1 && cell2 && cell1.value !== cell2.value) {
                    // 修改的单元格
                    changes.push({
                        rowIndex: cell1.rowIndex,
                        columnName: cell1.columnName,
                        changeType: 'modified',
                        oldValue: cell1.value,
                        newValue: cell2.value,
                        position: key
                    });
                }
            });
            
            res.json({
                success: true,
                changes: changes,
                version1,
                version2
            });
        })
        .catch(err => {
            console.error('版本比较失败:', err);
            res.status(500).json({ error: '版本比较失败' });
        });
});
