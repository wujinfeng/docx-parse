# docx-parse

## 简介

The HTML 生成 word.docx

## 安装

```
npm install http://git.iyunxiao.com/wujinfeng/docx-parse.git
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

## 测试

使用mocha，assert. 运行测试
 ``` npm run test ```

test/template目录下是测试用的word模板文件
test/data/docx 生成的word文件
test/data/input 测试用到的html

## 查看示例

启动服务
 ``` npm run dev ```
 会运行 bin/wwww.js

在bin目录下web服务。

## 问题反馈

使用过程中遇到到问题，请提交上来。或者你自己也可以修改代码，修复bug，添加功能，大家一起完善功能。
