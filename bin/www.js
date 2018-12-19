/**
 * 启动web服务，运行示例
 */
const Html2docx = require('../index');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const express = require('express');
const app = express();
let storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/')
    },
    filename: function (req, file, cb) {
        console.log('file',file)
        cb(null, Date.now() +'-' + file.originalname)
    }
});
let upload = multer({ storage: storage });

/**
 * 首页
 */
app.get('/', (req, res) => {
    fs.createReadStream(path.join(__dirname, './views/index.html')).pipe(res);
});

/**
 * 上传html
 */
app.post('/', upload.single('file'), (req, res) => {
    console.log(req.file)
    console.log(req.body)
    let filePath = req.file.path;
    let template = req.body.template;
    res.send('ok')
});

const port = 9000;
app.listen(port, () => {
    console.log(`启动 http://127.0.0.1:${port}`)
});
