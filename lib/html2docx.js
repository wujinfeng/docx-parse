const {exec} = require('child_process');
const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');
const DOMParser = require('xmldom').DOMParser;
const XMLSerializer = require('xmldom').XMLSerializer;

/**
 *  获取html
 *  使用pandoc转为word, 引用模板
 *  解压 word,
 *  处理 document.xml
 *  在压缩成 XXX.docx
 */

class Html2docx {
    constructor(filename) {
        this.filename = filename;
        this.inputPath = path.resolve(__dirname, 'data/input', filename);
        this.destPath = path.resolve(__dirname, 'data/docx', filename + '.docx');
        this.testPath = path.resolve(__dirname, 'data/docx', filename + 'test.docx');
        this.tempDir = path.resolve(__dirname, 'data/temp', filename);
        this.templatePath = path.resolve(__dirname, 'template', 'A3_template.docx');
        this.cmd = 'pandoc -f html+smart -t docx+smart --reference-doc ' + this.templatePath + ' -o ' + this.destPath + ' ' + this.inputPath;
        this.zip = '';
        console.log('html2doc命令：', this.cmd)
    }

    /**
     * 读取输入到html文件内容
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
            exec(this.cmd, (error, stdout, stderr) => {
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
        console.log('destPath:', this.destPath)
        let content = await this.getFileData(this.destPath)
        this.zip = await JSZip.loadAsync(content);
    }

    /**
     * 设置 table 边框属性
     * @returns {Promise<void>}
     */
    async setTableBorder() {
        let xmlstr = await this.zip.file('word/document.xml').async('string');
        console.log('xmlstr: ', xmlstr)
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
        console.log('newXmlStr:', newXmlStr)
        await this.zip.file('word/document.xml', newXmlStr); // 写回到文件
    }

    /**
     * 设置奇偶页 页眉
     * @param isHas
     * @returns {Promise<boolean>}
     */
    async setEvenAndOddHeaders(isHas) {
        let xmlstr = await this.zip.file('word/settings.xml').async('string');
        console.log('xmlstr: ', xmlstr)
        let dom = new DOMParser().parseFromString(xmlstr, 'text/xml');
        let nodeList = dom.documentElement.getElementsByTagName('w:evenAndOddHeaders');
        let isExist = nodeList.length > 0; // 奇偶页头元素是否存在
        if (isHas) {
            if (!isExist) { // 添加
                let evenAndOddHeaders = dom.createElement('w:evenAndOddHeaders');
                dom.documentElement.appendChild(evenAndOddHeaders);
            }
        } else {
            if (isExist) { // 删除
                for (let i = 0; i < nodeList.length; i++) {
                    dom.documentElement.removeChild(nodeList.item(i))
                }
            }
        }
        let newXmlStr = new XMLSerializer().serializeToString(dom);
        console.log('newXmlStr:', newXmlStr)
        await this.zip.file('word/settings.xml', newXmlStr); // 写回到文件里 settings.xml
    }

    /**
     * 保存文档
     * @returns {Promise<any>}
     */
    saveDocument() {
        return new Promise((resolve, reject) => {
            this.zip.generateNodeStream({streamFiles: true})
                .pipe(fs.createWriteStream(this.testPath))
                .on('finish', () => {
                    console.log('生成 word over')
                    resolve()
                })
                .on('error', (err) => {
                    reject(err);
                })
        })
    }


}


module.exports = Html2docx;



