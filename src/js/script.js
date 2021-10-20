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

        // Парсим DOM и фильтруем необходимые узлы
        this.parsedData = this._parseDOM().filter((node, i) => {
            if (node) {
                if (node.type === 'paragraph' && i > 1) {
                    return false;
                } else {
                    return true;
                }
            } else {
                return false;
            }
        });
    }

    // Получаем данные из запарсенного документа
    getData() {
        // Определяем тип шаблона
        switch (this.resultData.type) {
            case '08. Статический видеоряд':
                // Определяем подтип шаблона
                switch (this.resultData.subtype) {
                    case 'Изображение или фото':
                        this._getStaticImagesData();
                        console.log(this.resultData);

                        break;
                    default:
                        break;
                }

                break;
            default:
                break;
        }
    }

    // Парсим DOM
    _parseDOM() {
        // Массив с контентом документа
        const documentContent = Array.from(this.XMLDocument.querySelector('document').querySelector('body').childNodes);

        // Проходимся по узлам контента документа и выбираем то, что необходимо
        const parsedData = documentContent.map((node) => {
            return this._parseNodeContent(node);
        });

        return parsedData;
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

                const resultText = this._concatStrings(parsedRows);

                return {
                    type: 'paragraph',
                    text: resultText
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
                        const cellNodes = Array.from(cell.childNodes).filter((node) => {
                            return (
                                node.nodeName.slice(2) === 'p' ||
                                node.nodeName.slice(2) === 'tbl'
                            );
                        }).map((node) => {
                            return this._parseNodeContent(node);
                        });

                        return {
                            type: 'tableCell',
                            content: cellNodes
                        };
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

    // Метод для объединения разорванных строк с текстом в одну
    _concatStrings(rows) {
        if (!rows.length) return;

        const resultText = rows.map((row) => {
            return row.content.map((content) => {
                if (content.type === 'text') {
                    return content.content;
                }
            }).join('');
        }).join('');

        return resultText;
    }

    // Методы для получения данных из шаблона "08. Статический видеоряд :: Изображение или фото"
    _getStaticImagesData() {
        console.log(this.resultData);
        console.log(this.parsedData);

        // Сначала фильтруем таблицы, после чего уже обрабатываем только их
        this.resultData.tables = this.parsedData.filter((node) => {
            return node.type === 'table';
        }).map((table, i) => {
            // Для 1-й таблицы - особый механизм выбора данных
            if (i === 0) {
                const tableData = {
                    title: table.rows[0].cells[0].content[0].text,
                    taskType: '',
                    taskStatement: table.rows[2].cells[1].content[0].text
                };

                // Считываем тип задания
                table.rows[1].cells.forEach((cell) => {
                    const taskType = cell.content[0].text.trim();

                    if (
                        (taskType.charCodeAt(0) >= 1040 && taskType.charCodeAt(0) <= 1103) ||
                        taskType.charCodeAt(0) === 1025 ||
                        taskType.charCodeAt(0) === 1105
                    ) {
                        tableData.taskType = taskType;
                    }
                });

                return {
                    id: i,
                    data: tableData
                };
            }

            // const resultTable = [];

            // // Идем по строкам
            // table.rows.forEach((row) => {

            // });
        });
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

        // Получаем необходимые данные из запарсенного документа
        parserObject.getData();
    });
}

init();
