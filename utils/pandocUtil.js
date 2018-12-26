const debug = require('debug')('docx:utils:pandocUtil');
const {exec} = require('child_process');

/**
 * 转word, pandoc版本2.5, 1.9
 * @returns {Promise<any>}
 */
let pandocV25TransDocx = function (templateFile, inputFile, outFile) {
    return new Promise((resolve, reject) => {
        let cmd = 'pandoc -f html+smart -t docx+smart --reference-doc ' + templateFile + ' -o ' + outFile + ' ' + inputFile;
        debug('pandoc2.5命令：', cmd);
        exec(cmd, (error, stdout, stderr) => {
            if (error) {
                reject(error);
            } else {
                resolve(stdout);
            }
        });
    });
};

let pandocV119TransDocx = function (templateFile, inputFile, outFile) {
    return new Promise((resolve, reject) => {
        let cmd = 'pandoc -S --reference-docx ' + templateFile + ' ' + inputFile + ' -o ' + outFile;
        debug('pandoc1.19命令：', cmd);
        exec(cmd, (error, stdout, stderr) => {
            if (error) {
                reject(error);
            } else {
                resolve(stdout);
            }
        });
    });
};

let strategies = {
    '1.19': async function (templateFile, inputFile, outFile) {
        return await pandocV119TransDocx(templateFile, inputFile, outFile);
    },
    '2.5': async function (templateFile, inputFile, outFile) {
        return await pandocV25TransDocx(templateFile, inputFile, outFile);
    }
};

let transDocx = async function (version, templateFile, inputFile, outFile) {
    try {
        return await strategies[version](templateFile, inputFile, outFile);
    } catch (e) {
        throw e;
    }
};

module.exports = {
    transDocx: transDocx
};

