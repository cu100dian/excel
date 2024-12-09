document.getElementById('openFile').addEventListener('click', () => {
    document.getElementById('fileInput').click();
});

document.getElementById('fileInput').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) {
        alert('Файл не выбран.');
        return;
    }

    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Обработка данных
            const firms = ["Алма Вин ТОВ (1)", "Форстранс (1)", "Иванов Катена (1)", "Алма Ритейл (1)"];
            let firmsSum = 0;
            let hutorskayaAlmaSum = 0;
            const distributors = {
                "Винница": 0,
                "Сафьяны": 0
            };

            jsonData.forEach((row, index) => {
                if (index === 0) return; // Пропускаем заголовки

                // Копируем значения из "Добавил" в "Распределитель", если пусто
                if (!row[2]) row[2] = row[11];

                const distributor = row[2];
                const quantity = row[6] || 0;

                if (firms.includes(row[3])) {
                    firmsSum += quantity;
                } else if (distributor === "KDP-103" || distributor === "Хуторская Алма") {
                    hutorskayaAlmaSum += quantity;
                } else if (distributor === "Винница") {
                    distributors["Винница"] += quantity;
                } else if (distributor === "KDP106") {
                    distributors["Сафьяны"] += quantity;
                }
            });

            // Рассчитаем итоговую сумму
            const totalSum =
                firmsSum +
                hutorskayaAlmaSum +
                distributors["Винница"] +
                distributors["Сафьяны"];

            // Красивый вывод
            const output = `
________________________
Форс/Алма - ${firmsSum.toFixed(2)}
Хуторская Алма - ${hutorskayaAlmaSum.toFixed(2)}
Винница Алма - ${distributors["Винница"].toFixed(2)}
Сафьяны - ${distributors["Сафьяны"].toFixed(2)}
________________________
Сумма - ${totalSum.toFixed(2)}
________________________
            `;

            document.getElementById('output').innerText = output;
        } catch (error) {
            console.error('Ошибка обработки файла:', error);
            alert('Ошибка обработки файла. Проверьте формат.');
        }
    };

    reader.onerror = (e) => {
        console.error('Ошибка чтения файла:', e);
        alert('Ошибка чтения файла.');
    };

    reader.readAsArrayBuffer(file);
});
