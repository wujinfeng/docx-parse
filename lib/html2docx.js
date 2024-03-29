'use strict';
const debug = require('debug')('docx:lib:html2docx');
const pandocUtil = require('../utils/pandocUtil');
const fileUtil = require('../utils/fileUtil');
const htmlUtil = require('../utils/htmlUtil');
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
    evenAndOddHeaders: true,          // 区分奇偶页，默认false，不区分
    destPath: '',                     // 生成的word输出路径
    templateFile: '',                 // word模板文件,
    filename: uuid(),                 // 默认文件名,
    pandocVersion: '2.7',             // pandoc版本，默认2.7,
    deleteHeader: false               // 是否删除页眉
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
        try {
            let content = await fileUtil.getFileData(this.outFile);
            this.zip = await JSZip.loadAsync(content);
        } catch (e) {
            throw e
        }
    }

    /**
     * 解析html, pandoc2.7
     * @param content
     * @returns {Promise<boolean>}
     */
    async parse27(content) {
        try {
            let {html, stemsTableWidth} = await htmlUtil.preProcessAsemble(content);
            await fileUtil.generateHtmlFile(this.inputFile, html);
            await pandocUtil.transDocx(this.config.pandocVersion, this.config.templateFile, this.inputFile, this.outFile);
            await this.getZip();
            await xmlUtil.postProcess27(this.zip, stemsTableWidth); // 修改word文件
            await xmlUtil.setEvenAndOddHeaders(this.zip, this.config.evenAndOddHeaders); // 奇偶页
            await xmlUtil.saveDocument(this.zip, this.outFile); // 保存
            await fileUtil.deleteHtmlFile(this.inputFile); // 删除html
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
        try {
            let {html, stemsTableWidth} = await htmlUtil.preProcessAsemble(content);
            await fileUtil.generateHtmlFile(this.inputFile, html);
            await pandocUtil.transDocx(this.config.pandocVersion, this.config.templateFile, this.inputFile, this.outFile);
            await this.getZip();
            await xmlUtil.postProcess119(this.zip, stemsTableWidth); // 修改word文件
            await xmlUtil.deleteHeader(this.zip, this.config.deleteHeader); // 页眉
            await xmlUtil.setEvenAndOddHeaders(this.zip, this.config.evenAndOddHeaders); // 奇偶页
            await xmlUtil.saveDocument(this.zip, this.outFile); // 保存
            await fileUtil.deleteHtmlFile(this.inputFile); // 删除html
            return true;
        } catch (e) {
            throw e
        }
    }

    async parse(content) {
        let that = this;
        if (!content) {
            throw new Error('内容不能为空');
        }
        let strategies = {
            '1.19': async function (content) {
                return await that.parse119(content);
            },
            '2.7': async function (content) {
                return await that.parse27(content);
            }
        };
        try {
            return await strategies[that.config.pandocVersion](content);
        } catch (e) {
            throw e
        }
    }

}

module.exports = Html2docx;
