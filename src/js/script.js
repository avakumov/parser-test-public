import * as zip from '../../node_modules/@zip.js/zip.js/lib/zip.js';


class DOCXParser {
    constructor(file) {
        this.file = file;

        // Общая структура для всех шаблонов
        this.resultData = {
            code: '',
            type: '',
            subtype: '',
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

        // Парсим DOM
        const parsedData = this._parseDOM();

        // Фильтруем данные на всех уровнях вложенности (2 раза, т.к. после 1-й фильтрации
        // некоторые параграфы удаляются в последнюю очередь, и массивы становятся пустыми)
        this.filteredParsedData = this._filterData(parsedData, true);
        this.filteredParsedData = this._filterData(this.filteredParsedData, true);
    }

    // Получаем данные из запарсенного документа
    getData() {
        // Определяем тип шаблона
        switch (this.resultData.type) {
            case '08. Статический видеоряд':
                // Определяем подтип шаблона
                switch (this.resultData.subtype) {
                    case 'Изображение или фото':
                        // Выбираем данные для шаблона
                        this._getStaticImagesData();

                        break;
                    default:
                        break;
                }

                break;
            default:
                break;
        }

        return this.resultData;
    }

    _getTypeSubtype(text) {
        // Если считанный текст содержит код ЭОМа
        if (text.includes('Код ЭОМа:')) {
            this.resultData.code = text.slice(9).trim();
        }

        // Если считанный текст содержит тип и подтип
        if (text.includes('::')) {
            const index = text.indexOf('::');

            this.resultData.type = text.slice(0, index).trim();
            this.resultData.subtype = text.slice(index + 2).trim();
        }
    }

    // Фильтруем данные на уровнях вложенности ниже 1-го
    _filterData(data, onTopLevel = false) {
        return data.filter((node, i) => {
            if (node) {
                switch (node.type) {
                    case 'paragraph':
                        if (onTopLevel ? (!node.text || i > 1) : (!node.text)) return false;

                        // Вытаскиваем тип и подтип шаблона из параграфа
                        this._getTypeSubtype(node.text);

                        return true;
                    case 'tableCell':
                        if (!node.content.length) return false;

                        node.content = this._filterData(node.content);

                        return true;
                    case 'tableRow':
                        node.cells = this._filterData(node.cells);

                        return true;
                    case 'table':
                        node.rows = this._filterData(node.rows);

                        return true;
                    default:
                        return true;
                }
            } else {
                return false;
            }
        });
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

    // Метод для получения данных из шаблона "08. Статический видеоряд :: Изображение или фото"
    _getStaticImagesData() {
        // Сначала фильтруем таблицы, после чего уже обрабатываем только их
        this.resultData.tables = this.filteredParsedData.filter((node) => {
            return node.type === 'table';
        }).map((table, i) => {
            let tableData;

            // Для 1-й таблицы - особый механизм выбора данных
            if (i === 0) {
                tableData = {
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
            }

            // 2-я таблица
            if (i === 1) {
                tableData = {
                    rows: []
                };

                table.rows.forEach((row, i) => {
                    if (i === 0) return;

                    tableData.rows.push({
                        number: row.cells[0].content[0].text ? row.cells[0].content[0].text : '',
                        title: row.cells[1].content[0].text ? row.cells[1].content[0].text : '',
                        annotation: row.cells[2].content[0].text ? row.cells[2].content[0].text : '',
                        description: row.cells[3].content[0].text ? row.cells[3].content[0].text : '',
                        text: row.cells[4].content.map((content) => {
                            return content.text ? content.text : '';
                        })
                    });
                });
            }

            // 3-я таблица
            if (i === 2) {
                tableData = {
                    teacherRecom: table.rows[1].cells[0].content.map((content) => {
                        return content.text ? content.text : '';
                    }),
                    studentRecom: table.rows[1].cells[1].content.map((content) => {
                        return content.text ? content.text : '';
                    })
                };
            }

            return {
                id: i,
                data: tableData
            };
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
        // const data = parserObject.getData();

        // console.log(data);
        console.log(parserObject.resultData);
    });
}

init();
