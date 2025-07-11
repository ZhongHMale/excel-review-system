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
    // 文件版本表
    db.run(`CREATE TABLE IF NOT EXISTS file_versions (
        id TEXT PRIMARY KEY,
        filename TEXT NOT NULL,
        original_name TEXT NOT NULL,
        file_path TEXT NOT NULL,
        version INTEGER NOT NULL,
        submitter TEXT NOT NULL,
        change_description TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);

    // 表格数据表
    db.run(`CREATE TABLE IF NOT EXISTS table_data (
        id TEXT PRIMARY KEY,
        file_version_id TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        column_name TEXT NOT NULL,
        cell_value TEXT,
        cell_type TEXT DEFAULT 'text',
        annotations TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (file_version_id) REFERENCES file_versions (id)
    )`);

    // 批注表
    db.run(`CREATE TABLE IF NOT EXISTS annotations (
        id TEXT PRIMARY KEY,
        file_version_id TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        column_name TEXT NOT NULL,
        annotation_type TEXT NOT NULL,
        annotation_data TEXT,
        created_by TEXT NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (file_version_id) REFERENCES file_versions (id)
    )`);

    // 自定义列表
    db.run(`CREATE TABLE IF NOT EXISTS custom_columns (
        id TEXT PRIMARY KEY,
        file_version_id TEXT NOT NULL,
        column_name TEXT NOT NULL,
        column_type TEXT DEFAULT 'text',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (file_version_id) REFERENCES file_versions (id)
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
app.post('/api/upload', upload.single('file'), (req, res) => {
    try {
        const { submitter, changeDescription } = req.body;
        const file = req.file;
        
        if (!file) {
            return res.status(400).json({ error: '没有上传文件' });
        }

        const fileId = uuidv4();
        const { headers, data } = parseExcelFile(file.path);

        // 获取当前文件的最新版本号
        db.get(
            'SELECT MAX(version) as max_version FROM file_versions WHERE original_name = ?',
            [file.originalname],
            (err, row) => {
                if (err) {
                    console.error(err);
                    return res.status(500).json({ error: '数据库错误' });
                }

                const version = (row.max_version || 0) + 1;

                // 插入文件版本记录
                db.run(
                    `INSERT INTO file_versions (id, filename, original_name, file_path, version, submitter, change_description)
                     VALUES (?, ?, ?, ?, ?, ?, ?)`,
                    [fileId, file.filename, file.originalname, file.path, version, submitter, changeDescription],
                    function(err) {
                        if (err) {
                            console.error(err);
                            return res.status(500).json({ error: '保存文件信息失败' });
                        }

                        // 插入表格数据
                        const stmt = db.prepare(`
                            INSERT INTO table_data (id, file_version_id, row_index, column_name, cell_value, cell_type)
                            VALUES (?, ?, ?, ?, ?, ?)
                        `);

                        data.forEach((row, rowIndex) => {
                            headers.forEach(header => {
                                const cellId = uuidv4();
                                const cellValue = row[header];
                                const cellType = typeof cellValue === 'number' ? 'number' : 'text';
                                
                                stmt.run([cellId, fileId, rowIndex, header, cellValue, cellType]);
                            });
                        });

                        stmt.finalize();

                        res.json({
                            success: true,
                            fileId,
                            version,
                            headers,
                            dataCount: data.length,
                            message: '文件上传成功'
                        });
                    }
                );
            }
        );
    } catch (error) {
        console.error('上传错误:', error);
        res.status(500).json({ error: '文件处理失败' });
    }
});

// 获取文件版本列表
app.get('/api/files', (req, res) => {
    db.all(
        `SELECT id, filename, original_name, version, submitter, change_description, created_at
         FROM file_versions 
         ORDER BY original_name, version DESC`,
        (err, rows) => {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '获取文件列表失败' });
            }
            res.json(rows);
        }
    );
});

// 获取特定版本的数据
app.get('/api/files/:fileId/data', (req, res) => {
    const { fileId } = req.params;
    const { columns } = req.query;
    
    // 获取文件信息
    db.get(
        'SELECT * FROM file_versions WHERE id = ?',
        [fileId],
        (err, fileInfo) => {
            if (err) {
                console.error(err);
                return res.status(500).json({ error: '获取文件信息失败' });
            }
            
            if (!fileInfo) {
                return res.status(404).json({ error: '文件不存在' });
            }

            // 获取表格数据
            let query = `
                SELECT row_index, column_name, cell_value, cell_type, annotations
                FROM table_data 
                WHERE file_version_id = ?
            `;
            
            if (columns) {
                const columnList = columns.split(',').map(col => `'${col}'`).join(',');
                query += ` AND column_name IN (${columnList})`;
            }
            
            query += ' ORDER BY row_index, column_name';

            db.all(query, [fileId], (err, rows) => {
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
                        type: row.cell_type,
                        annotations: row.annotations ? JSON.parse(row.annotations) : null
                    };
                    headers.add(row.column_name);
                });

                const data = Object.values(dataMap);
                
                // 获取批注信息
                db.all(
                    'SELECT * FROM annotations WHERE file_version_id = ?',
                    [fileId],
                    (err, annotations) => {
                        if (err) {
                            console.error(err);
                            return res.status(500).json({ error: '获取批注失败' });
                        }

                        res.json({
                            fileInfo,
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

// 获取文件版本对比
app.get('/api/files/compare/:fileId1/:fileId2', (req, res) => {
    const { fileId1, fileId2 } = req.params;
    
    const getData = (fileId) => {
        return new Promise((resolve, reject) => {
            db.all(
                `SELECT row_index, column_name, cell_value
                 FROM table_data 
                 WHERE file_version_id = ?
                 ORDER BY row_index, column_name`,
                [fileId],
                (err, rows) => {
                    if (err) reject(err);
                    else resolve(rows);
                }
            );
        });
    };

    Promise.all([getData(fileId1), getData(fileId2)])
        .then(([data1, data2]) => {
            // 比较逻辑
            const changes = [];
            const map1 = {};
            const map2 = {};

            data1.forEach(row => {
                const key = `${row.row_index}-${row.column_name}`;
                map1[key] = row.cell_value;
            });

            data2.forEach(row => {
                const key = `${row.row_index}-${row.column_name}`;
                map2[key] = row.cell_value;
            });

            // 找出差异
            const allKeys = new Set([...Object.keys(map1), ...Object.keys(map2)]);
            
            allKeys.forEach(key => {
                const [rowIndex, columnName] = key.split('-');
                const value1 = map1[key];
                const value2 = map2[key];
                
                if (value1 !== value2) {
                    changes.push({
                        rowIndex: parseInt(rowIndex),
                        columnName,
                        oldValue: value1 || '',
                        newValue: value2 || '',
                        changeType: !value1 ? 'added' : !value2 ? 'deleted' : 'modified'
                    });
                }
            });

            res.json({ changes });
        })
        .catch(err => {
            console.error(err);
            res.status(500).json({ error: '版本对比失败' });
        });
});

// 启动服务器
app.listen(PORT, () => {
    console.log(`服务器运行在端口 ${PORT}`);
});