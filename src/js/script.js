import * as zip from '../../node_modules/@zip.js/zip.js/lib/zip.js';


class DOCXParser {
    constructor(file) {
        this.file = file;

        // Общая структура для всех шаблонов
        this.resultData = {
            code: null,
            type: null,
            subtype: null,
            tables: []
        };
    }

    // Парсим выбранный файл
    async parseDOCX() {
        const reader = new zip.ZipReader(new zip.BlobReader(this.file));
        const entries = await reader.getEntries();
        let parsedText;

        for (let key in entries) {
            const entry = entries[key];
    
            if (entry.filename === 'word/document.xml') {
                parsedText = await entry.getData(new zip.TextWriter());

                break;
            }
        }

        const parser = new DOMParser();
        const XMLDocument = parser.parseFromString(parsedText, 'application/xml');

        this.XMLDocument = XMLDocument;

        const documentContent = this._parseDOM();
    }

    getData() {

    }

    // Парсим DOM
    _parseDOM() {
        // Массив с контентом документа
        const documentContent = Array.from(this.XMLDocument.querySelector('document').querySelector('body').childNodes);
        console.log(documentContent);

        // Проходимся по узлам контента документа и выбираем то, что необходимо
        const newResult = documentContent.map((node) => {
            return this._parseNodeContent(node);
        });

        console.log(newResult);
    }

    // Парсим контент отдельного узла из DOM
    _parseNodeContent(node) {
        switch (node.nodeName.slice(2)) {
            // Если узел - параграф
            case 'p':
                // Берем строки параграфа
                const rowsArray = Array.from(node.querySelectorAll('r'));

                // Проходимся по полученным строкам
                const parsedRows = rowsArray.map((row) => {
                    // Получаем массив с контентом из строки
                    const rowContentArray = Array.from(row.childNodes);

                    // Проходимся по массиву с контентом из строки, предварительно отфильтровав нужный тип контента
                    const parsedRowContent = rowContentArray.filter((node) => {
                        return node.nodeName.slice(2) === 't';
                    }).map((node) => {
                        // Если текущий узел - текст
                        if (node.nodeName.slice(2) === 't') {
                            // Считываем текст узла
                            const nodeText = node.textContent;

                            // Если считанный текст содержит код ЭОМа
                            if (nodeText.includes('Код ЭОМа:')) {
                                this.resultData.code = nodeText.slice(9).trim();
                            }

                            // Если считанный текст содержит тип и подтип
                            if (nodeText.includes('::')) {
                                const index = nodeText.indexOf('::');

                                this.resultData.type = nodeText.slice(0, index).trim();
                                this.resultData.subtype = nodeText.slice(index + 2).trim();
                            }

                            return {
                                type: 'text',
                                content: nodeText
                            };
                        }
                    });

                    return {
                        type: 'paragraphRow',
                        content: parsedRowContent
                    };
                });

                return {
                    type: 'paragraph',
                    rows: parsedRows
                };
            // Если узел - таблица
            case 'tbl':
                // Выбираем только строки
                const tableRows = Array.from(node.childNodes).filter((node) => {
                    return node.nodeName.slice(2) === 'tr';
                });

                // Проходимся по строкам таблицы
                const parsedTableRows = tableRows.map((row) => {
                    // В строке выбираем только ячейки
                    const rowCells = Array.from(row.childNodes).filter((node) => {
                        return node.nodeName.slice(2) === 'tc';
                    });

                    // Проходимся по ячейкам одной строки
                    const parsedTableCells = rowCells.map((cell) => {
                        // Проходимся по контенту одной ячейки
                        return Array.from(cell.childNodes).filter((node) => {
                            return (
                                node.nodeName.slice(2) === 'p' ||
                                node.nodeName.slice(2) === 'tbl'
                            );
                        }).map((node) => {
                            return this._parseNodeContent(node);
                        });
                    });

                    return {
                        type: 'tableRow',
                        cells: parsedTableCells
                    }
                });

                return {
                    type: 'table',
                    rows: parsedTableRows
                };
            default:
                break;
        }
    }
}

const init = function() {
    const fileInput = document.querySelector('#file');

    fileInput.addEventListener('change', async () => {
        if (!fileInput.files.length) return;

        const selectedFile = fileInput.files[0];

        // Создаем объект нашего парсера
        const parserObject = new DOCXParser(selectedFile);

        // Парсим выбранный файл
        await parserObject.parseDOCX();
    });
}

init();
