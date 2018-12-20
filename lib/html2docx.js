'use strict';
const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;
const debug = require('debug')('docx:lib');
const {exec} = require('child_process');
const uuid = require('uuid/v4');
const JSZip = require('jszip');
const merge = require('merge');
const path = require('path');
const os = require('os');
const fs = require('fs');

/**
 *  获取html
 *  使用pandoc转为word, 引用模板
 *  解压 word,
 *  处理 settings.xml document.xml
 *  在压缩成 XXX.docx
 */

const defaultConfig = {
    evenAndOddHeaders: false,         // 区分奇偶页，默认false，不区分
    content: '',                      // html内容，字符串
    destPath: '',                     // 生成的word输出路径
    templateFile: '',                 // word模板文件,
    filename: uuid(),                 // 默认文件名
};

class Html2docx {
    constructor(config = {}) {
        this.config = merge(defaultConfig, config);
        this.inputFile = path.resolve(os.tmpdir(), uuid() + '.html');
        this.outFile = path.resolve(this.config.destPath, this.config.filename + '.docx');
        this.zip = '';
        this.checkConfig();
    }

    /**
     * 检查配置
     */
    checkConfig() {
        if (!this.config.content) {
            throw new Error('内容不能为空');
        }
        if (!this.config.destPath) {
            throw new Error('目标目录不存在');
        }
        if (!this.config.templateFile) {
            throw new Error('模板文件不存在');
        }
    }

    /**
     * 生成html文件
     */
    generateHtmlFile() {
        return new Promise((resolve, reject) => {
            fs.writeFile(this.inputFile, this.config.content, 'utf8', function (err) {
                if (err) {
                    reject(err);
                } else {
                    resolve();
                }
            });
        });
    }

    /**
     * 删除html文件
     */
    deleteHtmlFile() {
        return new Promise((resolve, reject) => {
            fs.unlink(this.inputFile, (err) => {
                if (err) {
                    reject(err);
                } else {
                    resolve();
                }
            });
        });
    }

    /**
     * 读取文件内容
     * @param filePath
     * @param encoding
     * @returns {Promise<any>}
     */
    getFileData(filePath, encoding) {
        return new Promise((resolve, reject) => {
            fs.readFile(filePath, {encoding: encoding}, function (err, data) {
                if (err) {
                    reject(err);
                } else {
                    resolve(data);
                }
            });
        });
    }

    /**
     * 转成word, 使用pandoc
     * @returns {Promise<any>}
     */
    transDocx() {
        return new Promise((resolve, reject) => {
            let cmd = 'pandoc -f html+smart -t docx+smart --reference-doc ' + this.config.templateFile + ' -o ' + this.outFile + ' ' + this.inputFile;
            debug('html2doc命令：', cmd);
            exec(cmd, (error, stdout, stderr) => {
                if (error) {
                    reject(error);
                } else {
                    resolve(stdout);
                }
            });
        })
    }

    /**
     * word转为zip文件
     * @returns {Promise<void>}
     */
    async getZip() {
        debug('outFile:', this.outFile)
        let content = await this.getFileData(this.outFile)
        this.zip = await JSZip.loadAsync(content);
    }

    /**
     * 设置 table 边框属性
     * @returns {Promise<void>}
     */
    async setTableBorder() {
        let xmlstr = await this.zip.file('word/document.xml').async('string');
        //debug('xmlstr: ', xmlstr)
        let dom = new DOMParser().parseFromString(xmlstr, 'text/xml');
        let nodeList = dom.documentElement.getElementsByTagName('w:tblPr'); // 获取table元素
        // 创建border边框里到元素及元素属性
        let createBorderEleAttr = function (tagName) {
            let ele = dom.createElement(tagName);
            ele.setAttribute('w:val', 'single');
            ele.setAttribute('w:sz', '4');
            ele.setAttribute('w:space', '0');
            ele.setAttribute('w:color', 'aaaaaa');
            return ele;
        };

        // 创建table的 border边框元素
        let borderEle = function () {
            let ele = dom.createElement('w:tblBorders');
            ele.appendChild(createBorderEleAttr('w:top'));
            ele.appendChild(createBorderEleAttr('w:left'));
            ele.appendChild(createBorderEleAttr('w:bottom'));
            ele.appendChild(createBorderEleAttr('w:right'));
            ele.appendChild(createBorderEleAttr('w:insideH'));
            ele.appendChild(createBorderEleAttr('w:insideV'));
            return ele;
        };

        // border边框元素加入到table， 元素每次都要生成，用函数
        for (let i = 0; i < nodeList.length; i++) {
            nodeList.item(i).appendChild(borderEle());
        }
        let newXmlStr = new XMLSerializer().serializeToString(dom); // 转成字符串
        // debug('newXmlStr:', newXmlStr)
        await this.zip.file('word/document.xml', newXmlStr); // 写回到文件
    }

    /**
     * 设置奇偶页 页眉
     * @param isHas
     * @returns {Promise<boolean>}
     */
    async setEvenAndOddHeaders() {
        let xmlstr = await this.zip.file('word/settings.xml').async('string');
        //debug('xmlstr: ', xmlstr)
        let dom = new DOMParser().parseFromString(xmlstr, 'text/xml');
        let nodeList = dom.documentElement.getElementsByTagName('w:evenAndOddHeaders');
        let find = nodeList.length > 0; // 奇偶页头元素是否存在
        if (this.config.evenAndOddHeaders) {
            if (!find) { // 添加
                let evenAndOddHeaders = dom.createElement('w:evenAndOddHeaders');
                dom.documentElement.appendChild(evenAndOddHeaders);
            }
        } else {
            if (find) { // 删除
                for (let i = 0; i < nodeList.length; i++) {
                    dom.documentElement.removeChild(nodeList.item(i))
                }
            }
        }
        let newXmlStr = new XMLSerializer().serializeToString(dom);
        //debug('newXmlStr:', newXmlStr)
        await this.zip.file('word/settings.xml', newXmlStr); // 写回到文件里 settings.xml
    }

    /**
     * 保存文档
     * @returns {Promise<any>}
     */
    saveDocument() {
        return new Promise((resolve, reject) => {
            this.zip.generateNodeStream({streamFiles: true})
                .pipe(fs.createWriteStream(this.outFile))
                .on('finish', () => {
                    debug('生成 word over')
                    resolve()
                })
                .on('error', (err) => {
                    reject(err);
                })
        })
    }

    async parse() {
        try {
            await this.generateHtmlFile();
            await this.transDocx();
            await this.getZip();
            await this.setEvenAndOddHeaders();
            await this.setTableBorder();
            await this.saveDocument();
            await this.deleteHtmlFile();
            return true;
        } catch (e) {
            throw e
        }
    }

}

module.exports = Html2docx;
