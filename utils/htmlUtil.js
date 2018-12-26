const debug = require('debug')('docx:utils:htmlUtil');
const _ = require('underscore');
const jsdom = require('jsdom');
const path = require('path');

let _parseTableInHtml = function (stemsTableWidth, window) {
    let $tables = window.$('.edittable');
    _.each($tables, function (table) {
        let stemTableWidth = {width: []};
        window.$(table).before('<span>this_is_a_tag_for_table</span>');
        stemTableWidth.trNum = window.$(table).find('tr').length;

        let isCombine = false;
        let colNum = 0;
        let firstTr = true;
        let curTrIndex = -1;
        _.each(window.$(table).find('tr'), function (tr) {
            ++curTrIndex;
            let curTdIndex = -1;
            _.each(window.$(tr).find('td'), function (td) {
                ++curTdIndex;
                let colspan = Number(window.$(td)[0].getAttribute('colspan'));
                if (colspan !== 0) {
                    window.$(td)[0].removeAttribute('colspan');
                    isCombine = true;
                    for (let i = 0; i < colspan - 1; ++i) {
                        window.$(td).after('<td>this_is_a_tag_for_span</td>');
                    }
                }
                if (firstTr) {
                    (colspan !== 0) ? colNum += colspan : colNum += 1;
                }
                if (colspan !== 0) {
                    curTdIndex += colspan - 1;
                }

                let rowspan = Number(window.$(td)[0].getAttribute('rowspan'));
                if (rowspan !== 0) {
                    window.$(td)[0].innerHTML = 'this_is_a_tag_for_row_span_begin' + window.$(td)[0].innerHTML;
                    window.$(td)[0].removeAttribute('rowspan');

                    let findItemRowIndex = curTrIndex;
                    let findItemColIndex = curTdIndex;
                    for (let i = 1; i < rowspan; ++i) {
                        let insertTr = window.$(table).find('tr')[findItemRowIndex + i];
                        //console.log('colnum: '+ colNum);
                        if ((findItemColIndex != 0) && (findItemColIndex == colNum - 1)) {
                            //console.log('after insert '+ (findItemRowIndex + i) +':'+ (findItemColIndex-1))
                            let insertTd = window.$(insertTr).find('td')[findItemColIndex - 1];
                            window.$(insertTd).after('<td>this_is_a_tag_for_row_span</td>');
                        } else {
                            let insertTd = window.$(insertTr).find('td')[findItemColIndex];
                            //console.log('before insert ' + (findItemRowIndex +i )+':'+ findItemColIndex);
                            window.$(insertTd).before('<td>this_is_a_tag_for_row_span</td>');
                        }
                    }
                }
            })

            firstTr = false;
        });
        if (!isCombine) {
            let trObj = window.$(table).find('tr')[0];
            _.each(window.$(trObj).find('td'), function (td) {
                stemTableWidth.width.push(Number(window.$(td)[0].getAttribute('width')));
            });
        } else {
            for (let i = 0; i < colNum; ++i) {
                stemTableWidth.width.push(0);
            }
        }
        stemsTableWidth.push(stemTableWidth);
    });
    _.each(stemsTableWidth, function (stemTable) {
        if (stemTable.width.indexOf(0) !== -1) {
            stemTable.width = _.map(stemTable.width, function (val) {
                return 8000 / 13 / stemTable.width.length;
            });
        }
    });
};

/**
 * 预先处理html，table
 * @param content
 */
let preProcessAsemble = async function (content) {
    return new Promise((resolve, reject) => {
        let stemsTableWidth = [];
        const path_ = path.resolve(__dirname, '../node_modules/jquery/dist/jquery.min.js');
        const html_ = '<body id="master">' + content + '</body>';
        jsdom.env(html_, [path_], function (err, window) {
            if(err){
                return reject(err);
            }
            let stemsTableWidth = [];
            _parseTableInHtml(stemsTableWidth, window);

            let thisHtml = window.$('#master').html();

            let $options = window.$('.question-options');
            const docx_with = 430;

            /**
             * The code is all operations are synchronized.
             */
            _.each($options, function (question_options) {
                if (window.$(question_options).find('.option').length != 0) {
                    var max_width = _.max(_.map(window.$(question_options).find('.option'), function (option) {
                        return Number(window.$(option)[0].getAttribute('data-width'));
                    }));

                    if (max_width > docx_with * 0.4) {
                        var strout = '';
                        _.each(window.$(question_options).find('.option'), function (option) {
                            strout += '<p class="option">' + window.$(option)[0].innerHTML + '</p>';
                        });
                        window.$('<p class="Enter">' + strout).insertBefore(window.$(question_options));
                    } else if (max_width > docx_with * 0.2) {

                        var tds = _.map(window.$(question_options).find('.option'), function (option) {
                            return '<td class="option">' + window.$(option)[0].innerHTML + '</td>';
                        });
                        var strout = '';
                        for (var ix in tds) {
                            var td = tds[ix];
                            if (Number(ix) % 2 == 0)
                                strout += '<tr>' + td;
                            else
                                strout += td;
                        }
                        var $tl = window.$('<table class="options">' + strout + '</table>');
                        $tl.insertBefore(window.$(question_options));
                    } else {
                        var tds = _.map(window.$(question_options).find('.option'), function (option) {
                            return '<td class="option">' + window.$(option)[0].innerHTML + '</td>';
                        });

                        var strout = '';
                        for (var ix in tds) {
                            var td = tds[ix];
                            strout += td;
                        }

                        var $tl = window.$('<table class="options"><tr>' + strout + '</tr></table>');
                        $tl.insertBefore(window.$(question_options));
                    }
                    window.$(question_options).remove();
                }
            });

            /**
             * image tag new line
             */
            _.each(window.$('img'), function (img) {
                var $p_start = window.$(window.document.createElement('p'));
                var $p_end = window.$(window.document.createElement('p'));
                $p_start.html('&nbsp;');
                $p_end.html('&nbsp;');

                $p_start.html(img.outerHTML);
                $p_start.insertBefore(window.$(img));
                window.$(img).remove();
            });

            _.each(window.$('math'), function (math) {
                var str = window.$(math)[0].outerHTML;
                str = str.replace(/[\u4E00-\u9FCB]/g,
                    function (a) {
                        return escape(a).replace(/(%u)(\w{4})/gi, 'cuihovah_$2_cuihovah');
                    });
                window.$(str).insertBefore(window.$(math));
                window.$(math).remove();
            });

            let html = window.$('#master').html();
            //debug('html: ', html);
            resolve({html, stemsTableWidth});
        });
    })
};

module.exports = {
    preProcessAsemble: preProcessAsemble
};
