// Основная функция
function processExcel() {
    try {
        // Загружаем Excel файл
        const fileName = 'smdp_20241209_192452.xlsx'; // Укажите ваш путь к файлу
        const workbook = XLSX.readFile(fileName);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Преобразуем в массив
        let data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Копируем значения из колонки "Добавил" в "Распределитель", если "Распределитель" пустой
        data.forEach(row => {
            if (!row[2]) { // Если колонка "Распределитель" пуста
                row[2] = row[11]; // Копируем значение из колонки "Добавил" (L) в "Распределитель" (C)
            }
        });

        // Убираем строки, где "Фирма" = "AllSeeds1" или "Diesel Delivery"
        data = data.filter(row => row[3] !== 'AllSeeds1' && row[3] !== 'Diesel Delivery');

        // Получаем даты для фильтрации
        const now = new Date();
        const today9am = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 9));
        const yesterday9am = new Date(today9am);
        yesterday9am.setUTCDate(yesterday9am.getUTCDate() - 1);

        // Фильтруем данные по времени
        const filteredData = data.filter((row, index) => {
            if (index === 0) return true; // Оставляем заголовки
            const date = parseExcelDate(row[0]); // Преобразуем дату из первого столбца
            return date >= yesterday9am && date < today9am;
        });

        // 1. Суммируем "Количество" по фирмам
        const firms = ["Алма Вин ТОВ (1)", "Форстранс (1)", "Иванов Катена (1)", "Алма Ритейл (1)"];
        let firmsSum = 0;

        filteredData.forEach(row => {
            if (firms.includes(row[3])) { // Если фирма в списке
                firmsSum += row[6] || 0;  // Суммируем "Количество" (колонка G)
            }
        });

        // 2. Суммируем "Количество" по распределителям
        let hutorskayaAlmaSum = 0;
        const distributors = {
            "Винница": 0,
            "Сафьяны": 0
        };

        filteredData.forEach(row => {
            const distributor = row[2];  // Колонка C "Распределитель"
            const quantity = row[6] || 0; // Колонка G "Количество"
            if (distributor === "KDP-103" || distributor === "Хуторская Алма") {
                hutorskayaAlmaSum += quantity;
            } else if (distributor === "Винница") {
                distributors["Винница"] += quantity;
            } else if (distributor === "KDP106") {
                distributors["Сафьяны"] += quantity;
            }
        });

        // Красивый вывод
        console.log("\n________________________");
        console.log(`Форс/Алма - ${firmsSum.toFixed(2)}`);
        console.log(`Хуторская Алма - ${hutorskayaAlmaSum.toFixed(2)}`);
        console.log(`Винница Алма - ${distributors["Винница"].toFixed(2)}`);
        console.log(`Сафьяны - ${distributors["Сафьяны"].toFixed(2)}`);
        console.log("________________________");

        // 3. Подсчитываем итоговую сумму
        const totalSum = firmsSum + hutorskayaAlmaSum + distributors["Винница"] + distributors["Сафьяны"];
        console.log(`Сумма - ${totalSum.toFixed(2)}`);
        console.log("\n________________________");

    } catch (error) {
        console.error(`Ошибка: ${error.message}`);
    }
}
