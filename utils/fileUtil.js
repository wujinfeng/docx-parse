'use strict';
const debug = require('debug')('docx:utils:fileUtil');
const fs = require('fs');

/**
 * 生成html文件
 * @param inputFile
 * @param content
 * @returns {Promise<any>}
 */
let generateHtmlFile = function (inputFile, content) {
    return new Promise((resolve, reject) => {
        fs.writeFile(inputFile, content, 'utf8', function (err) {
            if (err) {
                reject(err);
            } else {
                resolve();
            }
        });
    });
};

/**
 *  删除html文件
 * @param inputFile
 * @returns {Promise<any>}
 */
let deleteHtmlFile = function (inputFile) {
    return new Promise((resolve, reject) => {
        fs.unlink(inputFile, (err) => {
            if (err) {
                reject(err);
            } else {
                resolve();
            }
        });
    });
};

/**
 * 读取文件内容
 * @param filePath
 * @param encoding
 * @returns {Promise<any>}
 */
let getFileData = function (filePath, encoding) {
    return new Promise((resolve, reject) => {
        fs.readFile(filePath, {encoding: encoding}, function (err, data) {
            if (err) {
                reject(err);
            } else {
                resolve(data);
            }
        });
    });
};

module.exports = {
    generateHtmlFile: generateHtmlFile,
    deleteHtmlFile: deleteHtmlFile,
    getFileData: getFileData
};
