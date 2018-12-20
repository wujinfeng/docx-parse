# docx-parse
The HTML generate word.docx

## 安装

```
npm install --save @wujinfeng/docx-parse
```

## 使用方法

html生成word.docx

```
const Html2docx = require('docx-parse').Html2docx;
let html2docx = new Html2docx({
            evenAndOddHeaders: false,          // 区分奇偶页，默认false：不区分
            content: '',                       // html内容，字符串（必填）
            destPath: '',                      // 生成的word输出路径（必填）
            templateFile: '',                  // word模板文件（必填）
            filename: '',                      // 文件名
        });

html2docx.parse(); // promise对象，执行成功返回true

```

需要安装 pandoc 2.5版本.
