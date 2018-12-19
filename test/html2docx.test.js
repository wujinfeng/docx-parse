const Html2docxTest = require('../modules/utils/html2docx');
const assert = require('assert');
/*
// 获取html内容
new Html2docx('wjf.html').getHtml().then((data)=>{
	console.log(data)
}).catch((err)=>{
    console.log(err)
});
*/
/*

// 转换到docx
(async function () {
    try {
        let docx = await new Html2docx('wjf.html').transDocx()
        console.log(docx)
    } catch (e) {
        console.log(e)
    }

})()
*/
/*
// 获取 document.xml
(async function () {
    try {
        let xmlstr = await new Html2docx('wjf.html').getDocumentXmlDom()
        console.log(xmlstr)

    } catch (e) {
        console.log(e)
    }

})()*/

/*

// 获取 setting.xml setEvenAndOddHeaders
(async function () {
    try {
        let html2docx = new Html2docx('wjf.html');
        await html2docx.transDocx()
        await html2docx.getZip();
        await html2docx.setEvenAndOddHeaders(true)
        await html2docx.saveDocument();
    } catch (e) {
        console.log(e)
    }

})()
*/

//table 添加边框
/*
(async function () {
    try {
        let html2docx = new Html2docx('table.html');
        await html2docx.transDocx()
        await html2docx.getZip();
        await html2docx.setTableBorder()
        await html2docx.saveDocument();
    } catch (e) {
        console.log(e)
    }

})()*/

