// Функция для преобразования листа Excel в HTML-таблицу
function sheetToHtml(sheet) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    let out = '<thead><tr>';
    for(let C = range.s.c; C <= range.e.c; ++C){
        out += `<th>${sheet[XLSX.utils.encode_cell({c:C,r:range.s.r})]?.v || ''}</th>`; // Заголовки столбцов
    }
    out += '</tr></thead>';
    out += '<tbody>';
    for(let R = range.s.r + 1; R <= range.e.r; ++R){ // Пропускаем первую строку заголовков
        out += '<tr>';
        for(let C = range.s.c; C <= range.e.c; ++C){
            let cell_address = {c:C,r:R};
            let value = sheet[XLSX.utils.encode_cell(cell_address)]?.v || '';
            out += `<td contenteditable>${value}</td>`;
        }
        out += '</tr>';
    }
    return out + '</tbody>';
}

// Функция для загрузки файла Excel и заполнения таблицы
function loadExcel() {
    var file = document.getElementById('excelFileInput').files[0];
    if (!file) return alert("Выберите файл!");
    var reader = new FileReader();
    reader.onload = function(event) {
        var data = event.target.result;
        var workbook = XLSX.read(data, {type:'binary'});
        var first_sheet_name = workbook.SheetNames[0]; // Используем первый лист
        var worksheet = workbook.Sheets[first_sheet_name];
        
        document.getElementById('dataTable').innerHTML = sheetToHtml(worksheet);
    };
    reader.readAsBinaryString(file);   
}

// Функция сохранения изменений обратно в файл Excel
function saveExcel() {
    var table = document.querySelector('#dataTable');
    var rows = Array.from(table.rows).map(row => [...row.cells].map(td => td.textContent)); // Получаем все ячейки
    var wb = XLSX.utils.book_new(); // Новый документ
    var ws = XLSX.utils.aoa_to_sheet(rows); // Преобразуем массив строк в лист
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1'); // Добавляем лист в книгу
    XLSX.writeFile(wb, 'Автопарк готов на год.xlsx'); // Сохраняем файл
    alert("Изменённый файл доступен для скачивания.");
}
