
// Функция загрузки файла
const loadFile = () => {
    const fileInput = document.getElementById('fileInput');
    if (!fileInput.files.length) return alert("Выберите файл!");

    // Читаем выбранный файл
    const reader = new FileReader();
    reader.onload = function(event) {
        const data = event.target.result;
        
        // Преобразуем файл в Workbook объект
        const workbook = XLSX.read(data, {type:'binary'});

        // Получаем первый лист из рабочей книги
        let first_sheet_name = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[first_sheet_name];

        // Преобразование листа в массив строк таблицы
        const excelData = XLSX.utils.sheet_to_json(worksheet, {header:1});

        renderTable(excelData);
    };
    reader.readAsBinaryString(fileInput.files[0]);
};

// Отрисовка таблицы
const renderTable = (data) => {
    const tableBody = document.querySelector('#excelTable tbody') || document.createElement('tbody');
    while(tableBody.firstChild){
        tableBody.removeChild(tableBody.firstChild); // Очистка предыдущей таблицы
    }

    for(let i=0;i<data.length;i++) {
        const row = document.createElement('tr');
        for(let j=0;j<data[i].length;j++){
            const cell = document.createElement('td');
            cell.textContent = data[i][j];
            row.appendChild(cell);
        }
        tableBody.appendChild(row);
    }

    document.getElementById('excelTable').appendChild(tableBody);
};

// Сохранение изменений обратно в файл
const saveChanges = () => {
    const rows = Array.from(document.querySelectorAll('#excelTable tr'));
    const data = [];

    rows.forEach((row, idx) => {
        const cells = Array.from(row.cells).map(cell => cell.innerText.trim());
        data.push(cells);
    });

    // Создание нового Workbook объекта
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // Генерируем бинарный поток файла .xlsx
    const binary = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    // Создаем ссылку для скачивания обновленного файла
    const blob = new Blob([s2ab(binary)], { type: 'application/octet-stream' });
    const downloadLink = document.createElement('a');
    downloadLink.href = window.URL.createObjectURL(blob);
    downloadLink.download = 'edited.xlsx';
    downloadLink.click();
};

// Helper для преобразования строки в ArrayBuffer
function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for(var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}
