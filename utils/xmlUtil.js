'use strict';
const debug = require('debug')('docx:utils:xmlUtil');
const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;
const fs = require('fs');

/**
 * 保存文档
 * @returns {Promise<any>}
 */
let saveDocument = function(zip, outFile) {
    return new Promise((resolve, reject) => {
        zip.generateNodeStream({streamFiles: true})
            .pipe(fs.createWriteStream(outFile))
            .on('finish', () => {
                debug('生成 word over')
                resolve()
            })
            .on('error', (err) => {
                reject(err);
            })
    })
};

/**
 * 设置 table 边框属性
 * @returns {Promise<void>}
 */
let setTableBorder = async function(zip) {
    let xmlstr = await zip.file('word/document.xml').async('string');
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
    await zip.file('word/document.xml', newXmlStr); // 写回到文件
};

/**
 * 设置奇偶页 页眉
 * @param isHas
 * @returns {Promise<boolean>}
 */
let setEvenAndOddHeaders = async function(zip, evenAndOddHeaders) {
    let xmlstr = await zip.file('word/settings.xml').async('string');
    //debug('xmlstr: ', xmlstr)
    let dom = new DOMParser().parseFromString(xmlstr, 'text/xml');
    let nodeList = dom.documentElement.getElementsByTagName('w:evenAndOddHeaders');
    let find = nodeList.length > 0; // 奇偶页头元素是否存在
    if (evenAndOddHeaders) {
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
    await zip.file('word/settings.xml', newXmlStr); // 写回到文件里 settings.xml
};

module.exports = {
    setTableBorder: setTableBorder,
    setEvenAndOddHeaders: setEvenAndOddHeaders,
    saveDocument: saveDocument
};
