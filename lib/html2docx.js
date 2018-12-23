'use strict';
const debug = require('debug')('docx:lib:html2docx');
const pandocUtil = require('../utils/pandocUtil');
const fileUtil = require('../utils/fileUtil');
const xmlUtil = require('../utils/xmlUtil');
const uuid = require('uuid/v4');
const JSZip = require('jszip');
const merge = require('merge');
const path = require('path');
const os = require('os');

/**
 *  获取html
 *  使用pandoc转为word, 引用模板
 *  解压 word,
 *  处理 settings.xml document.xml
 *  在压缩成 XXX.docx
 */

const defaultConfig = {
    evenAndOddHeaders: false,         // 区分奇偶页，默认false，不区分
    destPath: '',                     // 生成的word输出路径
    templateFile: '',                 // word模板文件,
    filename: uuid(),                 // 默认文件名,
    pandocVersion: '2.5',              // pandoc版本，默认2.5
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
        if (!this.config.destPath) {
            throw new Error('目标目录不存在');
        }
        if (!this.config.templateFile) {
            throw new Error('模板文件不存在');
        }
    }

    /**
     * word转为zip文件
     * @returns {Promise<void>}
     */
    async getZip() {
        debug('outFile:', this.outFile)
        let content = await fileUtil.getFileData(this.outFile);
        this.zip = await JSZip.loadAsync(content);
    }

    /**
     * 解析html, pandoc2.5
     * @param content
     * @returns {Promise<boolean>}
     */
    async parse25(content) {
        if (!content) {
            throw new Error('内容不能为空');
        }
        try {
            await fileUtil.generateHtmlFile(this.inputFile, content);
            await pandocUtil.transDocx(this.config.pandocVersion, this.config.templateFile, this.inputFile, this.outFile);
            await this.getZip();
            await xmlUtil.setEvenAndOddHeaders(this.zip, this.config.evenAndOddHeaders);
            await xmlUtil.setTableBorder(this.zip);
            await xmlUtil.saveDocument(this.zip, this.outFile);
            await fileUtil.deleteHtmlFile(this.inputFile);
            return true;
        } catch (e) {
            throw e
        }
    }

    /**
     * 解析html, pandoc1.19
     * @param content
     * @returns {Promise<boolean>}
     */
    async parse119(content) {
        if (!content) {
            throw new Error('内容不能为空');
        }
        try {
            await fileUtil.generateHtmlFile(this.inputFile, content);
            await pandocUtil.transDocx(this.config.pandocVersion, this.config.templateFile, this.inputFile, this.outFile);
            await this.getZip();
            await xmlUtil.setEvenAndOddHeaders(this.zip, this.config.evenAndOddHeaders);
            await xmlUtil.setTableBorder(this.zip);
            await xmlUtil.saveDocument(this.zip, this.outFile);
            await fileUtil.deleteHtmlFile(this.inputFile);
            return true;
        } catch (e) {
            throw e
        }
    }

    async parse(content) {
        let that = this
        let strategies = {
            '1.19': function (content) {
                return that.parse119(content);
            },
            '2.5': function (content) {
                return that.parse25(content);
            }
        };
        return await strategies[that.config.pandocVersion](content);
    }

}

module.exports = Html2docx;
