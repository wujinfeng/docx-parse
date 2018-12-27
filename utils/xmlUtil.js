'use strict';
const debug = require('debug')('docx:utils:xmlUtil');
const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;
const _ = require('underscore');
const fs = require('fs');

/**
 * 创建元素和属性
 * @param dom
 * @param tagName
 * @param attrObj
 * @returns {*|ActiveX.IXMLDOMElement|HTMLElement}
 */
let createDomAndAttr = function (dom, tagName, attrObj) {
    let ele = dom.createElement(tagName);
    if (Object.prototype.toString.call(attrObj) === '[object Object]') {
        for (let i in attrObj) {
            if (attrObj.hasOwnProperty(i)) {
                ele.setAttribute(i, attrObj[i]);
            }
        }
    }
    return ele;
};

/**
 * 保存文档
 * @returns {Promise<any>}
 */
let saveDocument = function (zip, outFile) {
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
let setTableBorder = async function (zip) {
    let xmlstr = await zip.file('word/document.xml').async('string');
    let dom = new DOMParser().parseFromString(xmlstr, 'text/xml');
    let nodeList = dom.documentElement.getElementsByTagName('w:tblPr'); // 获取table元素
    // 创建border边框里到元素及元素属性
    let borderChild = ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV'];
    let getBorderChild = function (tagName) {
        return createDomAndAttr(dom, tagName, {'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '666666'})
    };
    // 创建table的 border边框元素
    let borderEle = function () {
        let ele = dom.createElement('w:tblBorders');
        for (let i = 0, n = borderChild.length; i < n; i++) {
            ele.appendChild(getBorderChild(borderChild[i]))
        }
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
 * @param evenAndOddHeaders
 * @returns {Promise<boolean>}
 */
let setEvenAndOddHeaders = async function (zip, evenAndOddHeaders) {
    let xmlstr = await zip.file('word/settings.xml').async('string');
    //debug('xmlstr: ', xmlstr)
    let dom = new DOMParser().parseFromString(xmlstr, 'text/xml');
    let nodeList = dom.documentElement.getElementsByTagName('w:evenAndOddHeaders');
    let mirrorMargins = dom.documentElement.getElementsByTagName('w:mirrorMargins');
    let find = nodeList.length > 0; // 奇偶页头元素是否存在
    if (evenAndOddHeaders) {
        if (!find) { // 添加
            let mirrorMargins = dom.createElement('w:mirrorMargins');
            let evenAndOddHeaders = dom.createElement('w:evenAndOddHeaders');
            dom.documentElement.appendChild(mirrorMargins);
            dom.documentElement.appendChild(evenAndOddHeaders);
        }
    } else {
        if (find) { // 删除
            for (let i = 0; i < nodeList.length; i++) {
                dom.documentElement.removeChild(nodeList.item(i))
            }
            for (let i = 0; i < mirrorMargins.length; i++) {
                dom.documentElement.removeChild(mirrorMargins.item(i))
            }
        }
    }
    let newXmlStr = new XMLSerializer().serializeToString(dom);
    //debug('newXmlStr:', newXmlStr)
    await zip.file('word/settings.xml', newXmlStr); // 写回到文件里 settings.xml
};

let _handleTableInDoc = function (stemsTableWidth, xmlStr) {
    //表格边框
    var toReplaceStr = '<w:r><w:t xml:space="preserve">this_is_a_tag_for_table</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="a3"/><w:tblW w:w="5000" w:type="pct" />';
    for (let stemTableIndex = 0; stemTableIndex < stemsTableWidth.length; stemTableIndex++) {
        let reg = new RegExp(toReplaceStr + '<w:tblLook /></w:tblPr><w:tblGrid />');
        let toInsertStr = '';
        for (let colIndex = 0; colIndex < stemsTableWidth[stemTableIndex].width.length; colIndex++) {
            let curWidth = stemsTableWidth[stemTableIndex].width[colIndex] * 13;
            toInsertStr += `<w:gridCol w:w=\"${curWidth}\"/>`;
        }

        toInsertStr = "<w:tblGrid>" + toInsertStr + "</w:tblGrid>";
        xmlStr = xmlStr.replace(reg, toReplaceStr + '<w:tblLook /></w:tblPr>' + toInsertStr);
    }

    var reg = new RegExp(toReplaceStr, 'g');
    xmlStr = xmlStr.replace(reg,
        '</w:p><w:tbl><w:tblPr><w:tblStyle w:val="a3"/><w:tblW w:w="5000" w:type="pct" /><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>');

    reg = new RegExp('<w:tblGrid />', 'g');
    xmlStr = xmlStr.replace(reg,
        '<w:tblGrid><w:gridCol w:w="1650"/><w:gridCol w:w="1650"/><w:gridCol w:w="1650"/><w:gridCol w:w="1650"/></w:tblGrid>');

    var spanTagStr = '<w:tc><w:p><w:pPr><w:textAlignment w:val="center"/><w:pStyle w:val="Compact" /><w:jc w:val="left" /></w:pPr><w:r><w:t xml:space="preserve">this_is_a_tag_for_span</w:t></w:r></w:p></w:tc>';
    //合并单元格情况处理
    while (xmlStr.search(spanTagStr) !== -1) {
        let spanTagIndex = xmlStr.search(spanTagStr);
        let spanTagNum = 2;
        let tmpStr = xmlStr.substr(spanTagIndex + spanTagStr.length);

        //计算多少个单元格需要合并
        while (tmpStr.search(spanTagStr) === 0) {
            spanTagNum++;
            tmpStr = tmpStr.substr(spanTagStr.length);
        }
        let frontStr = xmlStr.substr(0, spanTagIndex);
        let insertIndex = frontStr.lastIndexOf('<w:tc>');
        let insertStr = '<w:tcPr><w:gridSpan w:val="' + spanTagNum + '"/></w:tcPr>';
        xmlStr = frontStr.substr(0, insertIndex + '<w:tc>'.length) + insertStr +
            frontStr.substr(insertIndex + '<w:tc>'.length) + tmpStr;
    }

    var rowSpanBeginReg = new RegExp('<w:tc><w:p><w:pPr><w:textAlignment w:val="center"/><w:pStyle w:val="Compact" /><w:jc w:val="left" /></w:pPr><w:r><w:t xml:space="preserve">this_is_a_tag_for_row_span_begin', 'g');

    xmlStr = xmlStr.replace(rowSpanBeginReg,
        '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:pPr><w:textAlignment w:val="center"/><w:pStyle w:val="Compact" /><w:jc w:val="left" /></w:pPr><w:r><w:t xml:space="preserve">');

    var rowSpanTdReg = new RegExp('<w:tc><w:p><w:pPr><w:textAlignment w:val="center"/><w:pStyle w:val="Compact" /><w:jc w:val="left" /></w:pPr><w:r><w:t xml:space="preserve">this_is_a_tag_for_row_span', 'g');
    xmlStr = xmlStr.replace(rowSpanTdReg,
        '<w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p><w:pPr><w:textAlignment w:val="center"/><w:pStyle w:val="Compact" /><w:jc w:val="left" /></w:pPr><w:r><w:t xml:space="preserve">');

    return xmlStr;
};

/**
 * 设置表格属性
 * @param dom
 * @param tableDom
 */
let setTblPr = function (dom, tableDom) {
    let tblPr = dom.createElement('w:tblPr');
    let tblStyle = createDomAndAttr(dom, 'w:tblStyle', {'w:val': 'a3'});
    let tblW = createDomAndAttr(dom, 'w:tblW', {'w:w': '5000', 'w:type': 'pct'});
    let tblBorders = dom.createElement('w:tblBorders');
    let tblLook = dom.createElement('w:tblLook');
    let borderChild = ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV'];
    let getBorderChild = function (tagName) {
        return createDomAndAttr(dom, tagName, {'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '666666'})
    };
    for (let i = 0, n = borderChild.length; i < n; i++) {
        tblBorders.appendChild(getBorderChild(borderChild[i]))
    }
    tblPr.appendChild(tblStyle);
    tblPr.appendChild(tblW);
    tblPr.appendChild(tblBorders);
    tblPr.appendChild(tblLook);
    let oldTblPr = tableDom.getElementsByTagName('w:tblPr');
    if (oldTblPr.length > 0) {
        tableDom.replaceChild(tblPr, oldTblPr.item(0))
    } else {
        tableDom.insertBefore(tblPr, tableDom.firstChild)
    }
};

/**
 * 设置网格
 * @param dom
 * @param tableDom
 */
let setTblGrid = function (dom, tableDom) {
    let tblGrid = dom.createElement('w:tblGrid');
    let child = ['w:gridCol', 'w:gridCol', 'w:gridCol', 'w:gridCol'];
    let getChild = function (tagName) {
        return createDomAndAttr(dom, tagName, {'w:w': '1650'})
    };
    for (let i = 0, n = child.length; i < n; i++) {
        tblGrid.appendChild(getChild(child[i]))
    }
    let oldTblGrid = tableDom.getElementsByTagName('w:tblGrid');
    if (oldTblGrid.length > 0) {
        tableDom.replaceChild(tblGrid, oldTblGrid.item(0))
    } else {
        tableDom.insertBefore(tblGrid, tableDom.firstChild);
    }
};

/**
 * 设置单元格 跨列，跨行
 * this_is_a_tag_for_span， 跨列
 * this_is_a_tag_for_row_span_begin, 跨行开始
 * this_is_a_tag_for_row_span， 继续跨行
 * @param dom
 * @param tableDom
 */
let setSpan = function (dom, tableDom, stemsTableWidth) {
    let tcDomArr = tableDom.getElementsByTagName('w:tc');
    let deleteTcArr = [];
    for (let i = 0, n = tcDomArr.length; i < n; i++) {
        let tc = tcDomArr.item(i);
        let tArr = tc.getElementsByTagName('w:t');
        for (let j = 0, k = tArr.length; j < k; j++) {
            let t = tArr.item(j);
            if (t.firstChild.nodeValue.indexOf('this_is_a_tag_for_span') > -1) {
                let prevTc = tc.previousSibling;
                let tcPr = prevTc.getElementsByTagName('w:tcPr');
                if (tcPr.length > 0) {
                    let gridSpan = tcPr.item(0).getElementsByTagName('w:gridSpan');
                    if (gridSpan.length > 0) {
                        let num = Number(gridSpan.item(0).getAttribute('w:val')) + 1;
                        gridSpan.item(0).setAttribute('w:val', num.toString())
                    } else {
                        let newGridSpan = createDomAndAttr(dom, 'w:gridSpan', {'w:val': '2'});
                        tcPr.item(0).appendChild(newGridSpan)
                    }
                } else {
                    let newTcPr = createDomAndAttr(dom, 'w:tcPr');
                    let newGridSpan = createDomAndAttr(dom, 'w:gridSpan', {'w:val': '2'});
                    newTcPr.appendChild(newGridSpan);
                    if (prevTc.hasChildNodes()) {
                        prevTc.insertBefore(newTcPr, prevTc.firstChild)
                    } else {
                        prevTc.appendChild(newTcPr)
                    }
                }
                deleteTcArr.push(tc);
                break;
            } else if (t.firstChild.nodeValue.indexOf('this_is_a_tag_for_row_span_begin') > -1) {
                t.firstChild.textContent = t.firstChild.nodeValue.replace(/this_is_a_tag_for_row_span_begin/, '');
                let tcPr = dom.createElement('w:tcPr');
                let vMerge = createDomAndAttr(dom, 'w:vMerge', {'w:val': 'restart'});
                tcPr.appendChild(vMerge);
                tc.insertBefore(tcPr, tc.firstChild);
                break;
            } else if (t.firstChild.nodeValue.indexOf('this_is_a_tag_for_row_span') > -1) {
                t.firstChild.textContent = t.firstChild.nodeValue.replace(/this_is_a_tag_for_row_span/, '');
                let tcPr = dom.createElement('w:tcPr');
                let vMerge = createDomAndAttr(dom, 'w:vMerge', {'w:val': 'continue'});
                tcPr.appendChild(vMerge);
                tc.insertBefore(tcPr, tc.firstChild);
                break;
            }
        }
    }
    for (let i = 0, n = deleteTcArr.length; i < n; i++) { // 跨列的元素需要删除 单元格 tc
        let tc = deleteTcArr[i];
        tc.parentNode.removeChild(tc);
    }

};

/**
 * 设置指定内容是"this_is_a_tag_for_table"的表格样式
 * @param dom
 */
let setTableStyle = function (dom) {
    let tDomArr = dom.documentElement.getElementsByTagName('w:t');
    for (let i = 0, n = tDomArr.length; i < n; i++) {
        let t = tDomArr.item(i);
        if (t.firstChild.nodeValue === 'this_is_a_tag_for_table') {
            let p = t.parentNode.parentNode;
            let tableDom = p.nextSibling;
            if (tableDom.tagName === 'w:tbl') { // 找到表格标签
                setTblPr(dom, tableDom);
                setTblGrid(dom, tableDom);
                setSpan(dom, tableDom);
            }
            p.removeChild(t.parentNode) // 去掉table标志:this_is_a_tag_for_table
        }
    }
};

/**
 * 设置下划线
 * 标签：this_is_tag_underline
 * @param dom
 */
let setUnderline = function(dom){
    let tDomArr = dom.documentElement.getElementsByTagName('w:t');
    for(let i=0, n=tDomArr.length; i<n; i++){
        let t = tDomArr.item(i);
        if (t.firstChild.nodeValue.indexOf('this_is_tag_underline') > -1) {
            t.firstChild.textContent = t.firstChild.nodeValue.replace(/this_is_tag_underline/, '');
            let rPr = dom.createElement('w:rPr');
            let underline = createDomAndAttr(dom, 'w:u', {'w:val': 'single'});
            rPr.appendChild(underline);
            t.parentNode.insertBefore(rPr, t);
        }
    }
};

let postProcess = async function (zip, stemsTableWidth = []) {
    let xmlStr = await zip.file('word/document.xml').async('string');
    let dom = new DOMParser().parseFromString(xmlStr, 'text/xml');
    deleteTblW(dom);
    addPPrCenter(dom);
    setPage(dom);
    setExtent(dom);
    setTblStyleTblW(dom);
    setUnderline(dom);
    setTableStyle(dom, stemsTableWidth);

    xmlStr = new XMLSerializer().serializeToString(dom);

    xmlStr = xmlStr.replace(/cuihovah_/g, '&#x');
    xmlStr = xmlStr.replace(/_cuihovah/g, ';');
    xmlStr = xmlStr.replace(/\s*title=""\s*/g, ' ');
    //xmlStr = _handleTableInDoc(stemsTableWidth, xmlStr);
    await zip.file('word/document.xml', xmlStr); // 写回到文件
};

/**
 * 设置所有表格 tblStyle, tblW
 * @param dom
 */
let setTblStyleTblW = function (dom) {
    let tblPr = dom.documentElement.getElementsByTagName('w:tblPr');
    for (let i = 0, n = tblPr.length; i < n; i++) {
        let tblpr = tblPr.item(i);
        let tblStyle = tblpr.getElementsByTagName('w:tblStyle');
        let tblW = tblpr.getElementsByTagName('w:tblW');
        if (tblStyle.length > 0) {
            tblStyle.item(0).setAttribute('w:val', 'a3');
        } else {
            let newTblStyle = createDomAndAttr(dom, 'w:tblStyle', {'w:val': 'a3'});
            tblpr.appendChild(newTblStyle)
        }
        if (tblW.length > 0) {
            tblW.item(0).setAttribute('w:w', '5000');
            tblW.item(0).setAttribute('w:type', 'pct');
        } else {
            let newTblW = createDomAndAttr(dom, 'w:tblW', {'w:w': '5000', 'w:type': 'pct'});
            tblpr.appendChild(newTblW)
        }
    }
};

/**
 * 删除所有 w:tblW 标签
 * @param dom
 */
let deleteTblW = function (dom) {
    let tblW = dom.documentElement.getElementsByTagName('w:tblW');
    for (let i = 0, n = tblW.length; i < n; i++) {
        let tbl = tblW.item(i);
        tbl.parentNode.removeChild(tbl);
    }
};

/**
 * 添加 pPr 居中
 * @param dom
 */
let addPPrCenter = function (dom) {
    let pPrDom = dom.documentElement.getElementsByTagName('w:pPr');
    for (let i = 0, n = pPrDom.length; i < n; i++) {
        let ppr = pPrDom.item(i);
        let textAlignment = ppr.getElementsByTagName('w:textAlignment');
        if (textAlignment.length > 0) {
            textAlignment.item(0).setAttribute('w:val', 'center')
        } else {
            let newTextAlignment = createDomAndAttr(dom, 'w:textAlignment', {'w:val': 'center'});
            if (ppr.hasChildNodes()) {
                ppr.insertBefore(newTextAlignment, ppr.firstChild);
            } else {
                ppr.appendChild(newTextAlignment)
            }
        }
    }
};

/**
 * 设置图片范围
 * @param dom
 * @returns {*}
 */
let setExtent = function (dom) {
    let extentDom = dom.documentElement.getElementsByTagName('wp:extent');
    for (let i = 0, n = extentDom.length; i < n; i++) {
        let item = extentDom.item(i);
        let cx = item.getAttribute('cx');
        let cy = item.getAttribute('cy');
        item.setAttribute('cx', (Math.floor(Number(cx) / 1.4)).toString());
        item.setAttribute('cy', (Math.floor(Number(cy) / 1.4)).toString());
    }
};

/**
 * 设置分页 ，分页标志是 w:t word_docx_and_html_change_page
 * 遍历所有标签w:t,如果找到分页标志，则创建分页w:p 替换此 w:p
 * @param xmlStr
 * @returns {*|string}
 */
let setPage = function (dom) {
    let tDomArr = dom.documentElement.getElementsByTagName('w:t');

    let pagePDom = createDomAndAttr(dom, 'w:p');
    let pPr = createDomAndAttr(dom, 'w:pPr');
    let textAlignment = createDomAndAttr(dom, 'w:textAlignment', {'w:val': 'center'});
    let pStyle = createDomAndAttr(dom, 'w:pStyle', {'w:val': 'Compact'});
    let widowControl = createDomAndAttr(dom, 'w:widowControl');
    let jc = createDomAndAttr(dom, 'w:jc', {'w:val': 'left'});
    let r = createDomAndAttr(dom, 'w:r');
    let br = createDomAndAttr(dom, 'w:br', {'w:type': 'page'});
    let bookmarkStart = createDomAndAttr(dom, 'w:bookmarkStart', {'w:id': '0', 'w:name': '_GoBack'});
    let bookmarkEnd = createDomAndAttr(dom, 'w:bookmarkStart', {'w:id': '0'});

    pPr.appendChild(textAlignment);
    pPr.appendChild(pStyle);
    pPr.appendChild(widowControl);
    pPr.appendChild(jc);
    r.appendChild(br);

    pagePDom.appendChild(pPr);
    pagePDom.appendChild(r);
    pagePDom.appendChild(bookmarkStart);
    pagePDom.appendChild(bookmarkEnd);
    for (let i = 0, n = tDomArr.length; i < n; i++) {
        let tDom = tDomArr.item(i);
        if (tDom.firstChild.nodeValue === 'word_docx_and_html_change_page') {
            let pDom = tDom.parentNode.parentNode;
            if (pDom.tagName === 'w:p') {
                dom.replaceChild(pagePDom, pDom)
            }
        }
    }
};


module.exports = {
    setTableBorder: setTableBorder,
    setEvenAndOddHeaders: setEvenAndOddHeaders,
    saveDocument: saveDocument,
    postProcess: postProcess
};
