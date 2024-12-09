// Функция для обработки файла Excel
function handleFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const binaryData = e.target.result;
            const workbook = XLSX.read(binaryData, { type: 'binary' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Преобразуем данные в массив
            let excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Обрабатываем данные
            processExcelData(excelData);
        } catch (error) {
            console.error("Ошибка обработки файла:", error);
            document.getElementById('output').innerHTML = `Ошибка обработки файла: ${error.message}`;
        }
    };

    reader.readAsBinaryString(file);
}

// Обработчик кнопки "Open File"
document.getElementById('openFile').addEventListener('click', function () {
    document.getElementById('fileInput').click();
});

document.getElementById('fileInput').addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
});

// Функция для обработки данных из Excel
function processExcelData(excelData) {
    // Функция для преобразования даты Excel
    function parseExcelDate(value) {
        if (typeof value === 'number') {
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            return new Date(excelEpoch.getTime() + value * 86400000);
        } else if (typeof value === 'string') {
            return new Date(value);
        }
        return null;
    }

    // Фильтрация по времени (с 9 утра прошлого дня до 9 утра сегодняшнего дня)
    const now = new Date();
    const today9am = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 9));
    const yesterday9am = new Date(today9am);
    yesterday9am.setUTCDate(yesterday9am.getUTCDate() - 1);

    // Фильтруем данные по времени
    const filteredData = excelData.filter((row, index) => {
        if (index === 0) return true; // Оставляем заголовки
        const date = parseExcelDate(row[0]);
        return date >= yesterday9am && date < today9am;
    });

    // Удаляем строки с "Diesel Delivery" и "AllSeeds1"
    const cleanedData = filteredData.filter(row => {
        const firm = row[3]; // Колонка "Фирма"
        return firm !== 'Diesel Delivery' && firm !== 'AllSeeds1';
    });

    // Проходим по колонке "Распределитель" и меняем "KDP-103" на "Хуторская Алма"
    cleanedData.forEach(row => {
        if (row[2] === 'KDP-103') { // Колонка "Распределитель"
            row[2] = 'Хуторская Алма';
        }

        // Если в колонке "Распределитель" пусто, копируем из "Добавил"
        if (!row[2]) { // Колонка "Распределитель"
            row[2] = row[11]; // Копируем из колонки "Добавил" (колонка L)
        }

        // Меняем распределитель на "Форс Алма", если в колонке "Фирма" определенные значения
        const firm = row[3]; // Колонка "Фирма"
        const firmsToReplace = ["Алма Вин ТОВ (1)", "Форстранс (1)", "Иванов Катена (1)", "Алма Ритейл (1)"];
        if (firmsToReplace.includes(firm)) {
            row[2] = 'Форс Алма'; // Меняем "Распределитель" на "Форс Алма"
        }
    });

    // Суммируем данные по распределителям
    let hutorskayaAlmaSum = 0;
    let vinnytsiaAlmaSum = 0;
    let kdp106Sum = 0;
    let forsalmaSum = 0;

    cleanedData.forEach(row => {
        const distributor = row[2]; // Колонка "Распределитель"
        const quantity = row[6] || 0; // Колонка "Количество"

        if (distributor === 'Хуторская Алма') {
            hutorskayaAlmaSum += quantity;
        } else if (distributor === 'Винница') {
            vinnytsiaAlmaSum += quantity;
        } else if (distributor === 'KDP106') {
            kdp106Sum += quantity;
        } else if (distributor === 'Форс Алма') {
            forsalmaSum += quantity;
        }
    });

    // Формируем итоговый вывод с переносами строк
    const output = `
Хуторская Алма: ${hutorskayaAlmaSum.toFixed(2)}<br>
Винница: ${vinnytsiaAlmaSum.toFixed(2)}<br>
KDP106: ${kdp106Sum.toFixed(2)}<br>
Форс Алма: ${forsalmaSum.toFixed(2)}<br><br>

Сумма: ${(hutorskayaAlmaSum + vinnytsiaAlmaSum + kdp106Sum + forsalmaSum).toFixed(2)}
`;

    // Отображаем результат с использованием innerHTML, чтобы вставить <br> теги
    document.getElementById('output').innerHTML = output;
}
