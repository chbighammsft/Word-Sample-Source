
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#run-sample').click(runSample);
        });
    };

    function runSample() {

        var context = Office.context;

        Word.run(function (context) {

            return context.sync()
                .then(processHyperlinks)
                .then(processTables)
                .then(processParagraphs)
                .then(processWords);

            // return processWords();
            // return processHyperlinks();

            function processHyperlinks() {
                console.log("1) processHyperlinks called");
                var hyperlinks = context.document.body.getRange().getHyperlinkRanges();
                hyperlinks.load();

                return context.sync().then(function () {
                    console.log("2) processHyperlinks loop");
                    for (var i = 0; i < hyperlinks.items.length; i++) {
                        var link = hyperlinks.items[i];
                        var mdLink = '[' + link.text + '](' + link.hyperlink + ') ';
                        link.hyperlink = "";
                        link.insertText(mdLink, 'Replace');
                    }
                });
            }

            function processTables() {
                console.log("3) processTables called");
                var tables = context.document.body.tables;
                tables.load()

                context.sync().then(function () {
                    console.log("4) processTables loop");
                    for (var i = 0; i < tables.items.length; i++) {
                        var table = tables.items[i];

                        for (var j = 0; j < table.rowCount; j++) {
                            var row = table.values[j];

                            var rowParagraph = table.insertParagraph('| ', 'Before');
                            rowParagraph.style = 'Normal';

                            for (var k = 0; k < row.length; k++) {
                                var cell = row[k];

                                if (j < table.headerRowCount) {
                                    rowParagraph.insertText('**' + cell + '** | ', 'End');
                                }
                                else {
                                    rowParagraph.insertText(cell + ' | ', 'End');
                                }
                                console.log();
                            }
                        }
                        table.deleteRows(0, table.rowCount);
                    }
                });
            }

            function processParagraphs() {
                console.log("5) processParagraphs called");
                var paragraphs = context.document.body.paragraphs;
                paragraphs.load();

                return context.sync().then(function () {
                    var isCode = false;
                    var isList = false;


                    for (var i = 0; i < paragraphs.items.length; i++) {
                        console.log("6) processParagraphs loop");
                        var paragraph = paragraphs.items[i];
                        if (paragraph.style.indexOf('Code') === -1) {
                            if (isCode) {
                                var oldStyle = paragraph.style;
                                paragraph.style = 'Normal';
                                paragraph.insertParagraph('```', 'Before');
                                paragraph.style = oldStyle;
                                isCode = false;
                            }
                        }

                        if (paragraph.style.indexOf('List') === -1) {
                            if (isList) {
                                var oldStyle = paragraph.style;
                                paragraph.style = 'Normal';
                                paragraph.insertParagraph('', 'Before');
                                paragraph.style = oldStyle;
                                isList = false;
                            }
                        }

                        // Only process a paragraph outside of a table.
                        if (paragraph.tableNestingLevel === 0) {
                            if (paragraph.style.indexOf('Heading') >= 0) {
                                if (paragraph.style.indexOf('1') >= 0) {
                                    paragraph.insertText('# ', 'Start');
                                }
                                else if (paragraph.style.indexOf('2') >= 0) {
                                    paragraph.insertText('## ', 'Start');
                                } else if (paragraph.style.indexOf('3') >= 0) {
                                    paragraph.insertText('### ', 'Start');
                                } else if (paragraph.style.indexOf('4') >= 0) {
                                    paragraph.insertText('#### ', 'Start');
                                }
                                paragraph.style = 'Normal'
                            }

                            else if (paragraph.style.indexOf('Emphasis') >= 0) {
                                paragraph.insertText('*', 'Start');
                                paragraph.insertText('*', 'End');
                                paragraph.insertParagraph('', 'After');
                                paragraph.style = 'Normal';
                            }

                            else if (paragraph.style.indexOf('Code') >= 0) {
                                if (!isCode) {
                                    paragraph.insertParagraph('```', 'Before');
                                    isCode = true;
                                }
                            }

                            else if (paragraph.style.indexOf('List') >= 0) {
                                paragraph.insertText('* ', 'Start');
                                paragraph.style = 'Normal';
                                isList = true;
                            }

                            else if (paragraph.style.indexOf('Normal') >= 0) {
                                paragraph.style = 'Normal';
                                paragraph.insertParagraph('', 'After');
                            }

                            else {
                                paragraph.insertText('[' + paragraph.style + ' is unknown style] ', 'Start');
                            }

                        }
                    }
                });
            }

            function processWords() {
                console.log("7) processWords called");
                var paragraphs = context.document.body.paragraphs;
                paragraphs.load();

                return context.sync().then(function () {
                    console.log("8) processWords loop");
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        console.log("9) processing paragraph " + i);
                        handleWords(paragraphs.items[i]);
                    }
                });

                function handleWords(paragraph) {
                    var wordRanges = paragraph.getTextRanges([' '], true);
                    wordRanges.load("font, text");

                    context.sync().then(function () {
                        for (var i = 0; i < wordRanges.items.length; i++) {
                            var word = wordRanges.items[i];

                            var previousWord = wordRanges.items[i - 1];
                            var nextWord = wordRanges.items[i + 1];

                            if (word.font.bold) {
                                console.log('bold word found: ' + word.text);
                                if ((typeof previousWord === 'undefined') || !previousWord.font.bold) {
                                    word.insertText('**', 'Start');
                                }
                                if ((typeof nextWord === 'undefined') || !nextWord.font.bold) {
                                    word.insertText('**', 'End');
                                }
                            }

                            if (word.font.italic) {
                                console.log('italic word found: ' + word.text);
                                if ((typeof previousWord === 'undefined') || !previousWord.font.italic) {
                                    word.insertText('*', 'Start');
                                }
                                if ((typeof nextWord === 'undefined') || !nextWord.font.italic) {
                                    word.insertText('*', 'End');
                                }
                            }
                        }
                    });
                };
            }
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
                console.log("Trace info: " + JSON.stringify(error.traceMessages));
            }
        });
    }
})();