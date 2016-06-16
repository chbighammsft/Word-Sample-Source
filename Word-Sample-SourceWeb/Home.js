
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

            // Main function
            // Runs each of the processing functions in order.
            return context.sync()
                .then(processHyperlinks)
                .then(processTables)
                .then(processParagraphs)
                .then(processWords);

            // End of main function.

            // Gets a collection of all of the hyperlinks in a document and
            // converts them to Markdown style hyperlinks.
            function processHyperlinks() {
                context.trace("Called processHyperlinks")
                var hyperlinks = context.document.body.getRange().getHyperlinkRanges();
                hyperlinks.load();

                return context.sync().then(function () {
                    for (var i = 0; i < hyperlinks.items.length; i++) {
                        var link = hyperlinks.items[i];
                        var mdLink = '[' + link.text + '](' + link.hyperlink + ') ';
                        link.hyperlink = "";
                        link.insertText(mdLink, 'Replace');
                    }
                });
            }

            // Gets a collection of all of the tables in the document and converts
            // them to Markdown-style tables with bold text in the header rows.
            // The function does not currently handle justification in the table cells.
            function processTables() {
                context.trace("Called processTables");
                var tables = context.document.body.tables;
                tables.load();

                context.sync().then(function () {
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
                            }
                        }
                        table.deleteRows(0, table.rowCount);
                    }
                });
            }

            // Gets a collection of all of the paragraphs in the document and then 
            // converts paragraph styles to Markdown styles. The following styles
            // are handled by this function:
            //     * Normal
            //     * Heading levels 1 through 4
            //     * List
            //     * Emphasis
            //     * Code
            function processParagraphs() {
                context.trace("Called processParagraphs");
                var paragraphs = context.document.body.paragraphs;
                paragraphs.load();

                return context.sync().then(function () {
                    var isCode = false;
                    var isList = false;


                    for (var i = 0; i < paragraphs.items.length; i++) {
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

            // Gets a collection of all of the paragraphs in a document and
            // then for each paragraph gets a collection of all of the words
            // in the paragraph. It checks each word to see if it's formatted
            // italic or bold, and if so, outputs the appropriate Markdown
            // code ("*" for italic, "**" for bold). If several words in a 
            // row are styled the same, the Markdown code is only output at 
            // the beginning and end of the string.
            function processWords() {
                context.trace("Called processWords");
                var paragraphs = context.document.body.paragraphs;
                paragraphs.load();

                return context.sync().then(function () {
                    for (var i = 0; i < paragraphs.items.length; i++) {
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
                                if ((typeof previousWord === 'undefined') || !previousWord.font.bold) {
                                    word.insertText('**', 'Start');
                                }
                                if ((typeof nextWord === 'undefined') || !nextWord.font.bold) {
                                    word.insertText('**', 'End');
                                }
                            }

                            if (word.font.italic) {
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