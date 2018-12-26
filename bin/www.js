/**
 * 启动web服务，运行示例
 */
'use strict';
const Html2docx = require('../index').Html2docx;
const debug = require('debug')('docx:bin:www');
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const app = express();
let storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, path.join(__dirname, './uploads/'))
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname)
    }
});
let upload = multer({storage: storage});

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
    debug('file:', req.file)
    debug('body:', req.body)
    if(req.file.mimetype !== 'text/html'){
        return res.json({code: 4, msg: '请上传html文件'});
    }
    let filePath = req.file.path;
    let template = req.body.template;
    let pandoc = req.body.pandoc;
    let evenAndOddHeaders = Number(req.body.evenAndOddHeaders) === 1;
    debug('req.body', req.body);
    fs.readFile(filePath, 'utf8', (err, data) => {
        let destPath = path.dirname(filePath);
        let templateFile = path.resolve(__dirname, '../test/template', template + '.docx');

        let html2docx = new Html2docx({
            evenAndOddHeaders: evenAndOddHeaders,                // 区分奇偶页，默认false，不区分
            destPath: destPath,                      // 生成的word输出路径
            templateFile: templateFile,              // word模板文件,
            filename: req.file.originalname,         // 默认文件名
            pandocVersion: pandoc,                   // pandoc 版本
        });

        html2docx.parse(data).then(() => {
            res.json({code: 0, msg: 'ok', data: destPath + path.sep + req.file.originalname + '.docx'});
        }).catch((e) => {
            res.send(e)
        });

    });
});


const port = 9000;
app.listen(port, () => {
    console.log(`启动 http://127.0.0.1:${port}`)
});
