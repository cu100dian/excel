document.getElementById("openFile").addEventListener("click", () => {
    document.getElementById("fileInput").click();
});

document.getElementById("fileInput").addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            let rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Фильтрация данных
            rows = filterAndProcessData(rows);

            // Форматирование результата
            const output = formatOutput(rows);
            document.getElementById("output").innerText = output;
        } catch (error) {
            document.getElementById("output").innerText = `Ошибка обработки файла: ${error.message}`;
        }
    };
    reader.readAsArrayBuffer(file);
});

function parseExcelDate(value) {
    if (typeof value === "number") {
        const excelEpoch = new Date(Date.UTC(1899, 11, 30));
        return new Date(excelEpoch.getTime() + value * 86400000);
    } else if (typeof value === "string") {
        return new Date(value);
    }
    return null;
}

function filterAndProcessData(rows) {
    const now = new Date();
    const today9am = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 9));
    const yesterday9am = new Date(today9am);
    yesterday9am.setUTCDate(yesterday9am.getUTCDate() - 1);

    // Фильтруем строки по времени
    const filteredRows = rows.filter((row, index) => {
        if (index === 0) return true; // Заголовки не фильтруем
        const date = parseExcelDate(row[0]); // Дата в первом столбце
        if (date instanceof Date && !isNaN(date)) {
            return date >= yesterday9am && date < today9am;
        }
        return false;
    });

    // Удаляем строки с фирмами "AllSeeds1" и "Diesel Delivery"
    const result = filteredRows.filter((row, index) => {
        if (index === 0) return true; // Заголовки не фильтруем
        const firm = row[3]; // Название фирмы в четвертой колонке
        return firm !== "AllSeeds1" && firm !== "Diesel Delivery";
    });

    return result;
}

function formatOutput(rows) {
    // Подсчет значений
    let firmsSum = 0;
    let hutorskayaAlmaSum = 0;
    const distributors = { "Винница": 0, "Сафьяны": 0 };

    rows.forEach((row, index) => {
        if (index === 0) return; // Пропускаем заголовок
        const distributor = row[2]; // Колонка C
        const quantity = parseFloat(row[6]) || 0; // Колонка G
        const firm = row[3]; // Колонка D

        if (["Алма Вин ТОВ (1)", "Форстранс (1)", "Иванов Катена (1)", "Алма Ритейл (1)"].includes(firm)) {
            firmsSum += quantity;
        } else if (distributor === "KDP-103" || distributor === "Хуторская Алма") {
            hutorskayaAlmaSum += quantity;
        } else if (distributor === "Винница") {
            distributors["Винница"] += quantity;
        } else if (distributor === "KDP106") {
            distributors["Сафьяны"] += quantity;
        }
    });

    // Итоговая сумма
    const totalSum = firmsSum + hutorskayaAlmaSum + distributors["Винница"] + distributors["Сафьяны"];

    // Формируем вывод
    return `
________________________
Форс/Алма - ${firmsSum.toFixed(2)}
Хуторская Алма - ${hutorskayaAlmaSum.toFixed(2)}
Винница Алма - ${distributors["Винница"].toFixed(2)}
Сафьяны - ${distributors["Сафьяны"].toFixed(2)}
________________________
Сумма - ${totalSum.toFixed(2)}
    `;
}
