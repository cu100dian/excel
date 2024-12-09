// Функция для обработки Excel файла
function handleFile(file) {
    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });

            // Выбираем первый лист
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Преобразуем данные из Excel в массив
            let data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Функция обработки данных
            processExcelData(data);
        } catch (error) {
            console.error("Ошибка обработки файла:", error);
            document.getElementById('output').textContent = `Ошибка обработки файла: ${error.message}`;
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
function processExcelData(data) {
    // Преобразуем даты, фильтруем и суммируем значения
    const now = new Date();
    const today9am = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 9));
    const yesterday9am = new Date(today9am);
    yesterday9am.setUTCDate(yesterday9am.getUTCDate() - 1);

    // Фильтруем строки по времени (с 9 утра вчера до 9 утра сегодня)
    const filteredData = data.filter((row, index) => {
        if (index === 0) return true; // Оставляем заголовки
        const date = new Date(row[0]);
        return date >= yesterday9am && date < today9am;
    });

    // Суммируем "Количество" по фирмам и распределителям
    let firmsSum = 0;
    let hutorskayaAlmaSum = 0;
    let vinnytsiaAlmaSum = 0;
    let safyanySum = 0;
    const firms = ["Алма Вин ТОВ (1)", "Форстранс (1)", "Иванов Катена (1)", "Алма Ритейл (1)"];

    filteredData.forEach(row => {
        const firm = row[3]; // Колонка D "Фирма"
        const distributor = row[2]; // Колонка C "Распределитель"
        const quantity = row[6] || 0; // Колонка G "Количество"

        // Суммируем для фирм
        if (firms.includes(firm)) {
            firmsSum += quantity;
        }

        // Суммируем для распределителей
        if (distributor === "KDP-103" || distributor === "Хуторская Алма") {
            hutorskayaAlmaSum += quantity;
        } else if (distributor === "Винница") {
            vinnytsiaAlmaSum += quantity;
        } else if (distributor === "KDP106") {
            safyanySum += quantity;
        }
    });

    // Формируем итоговый вывод
    const totalSum = firmsSum + hutorskayaAlmaSum + vinnytsiaAlmaSum + safyanySum;
    const output = `
________________________
Форс/Алма - ${firmsSum.toFixed(2)}
Хуторская Алма - ${hutorskayaAlmaSum.toFixed(2)}
Винница Алма - ${vinnytsiaAlmaSum.toFixed(2)}
Сафьяны - ${safyanySum.toFixed(2)}
________________________
Сумма - ${totalSum.toFixed(2)}
________________________
`;

    // Отображаем результат
    document.getElementById('output').textContent = output;
}
