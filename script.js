// Скрипт для обработки Excel-файла
const XLSX = window.XLSX;

document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("fileInput");
    const openFileButton = document.getElementById("openFile");
    const outputDiv = document.getElementById("output");

    // При клике на кнопку открываем диалог выбора файла
    openFileButton.addEventListener("click", () => {
        fileInput.click();
    });

    // При выборе файла
    fileInput.addEventListener("change", async (event) => {
        const file = event.target.files[0];
        if (!file) {
            outputDiv.textContent = "Файл не выбран!";
            return;
        }

        // Читаем файл
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: "array" });

                // Берем первый лист
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                // Ваши фильтры
                const results = processExcel(rows);
                displayResults(results);
            } catch (err) {
                outputDiv.textContent = `Ошибка обработки файла: ${err.message}`;
            }
        };
        reader.readAsArrayBuffer(file);
    });

    // Ваша логика фильтров
    function processExcel(data) {
        const firms = ["Алма Вин ТОВ (1)", "Форстранс (1)", "Иванов Катена (1)", "Алма Ритейл (1)"];
        const distributors = {
            "Винница": 0,
            "Сафьяны": 0
        };
        let firmsSum = 0, hutorskayaAlmaSum = 0;

        data.forEach((row, index) => {
            if (index === 0) return; // Пропускаем заголовок
            const distributor = row[2];
            const quantity = parseFloat(row[6]) || 0;
            const firm = row[3];

            if (firms.includes(firm)) {
                firmsSum += quantity;
            } else if (distributor === "KDP-103" || distributor === "Хуторская Алма") {
                hutorskayaAlmaSum += quantity;
            } else if (distributor === "Винница") {
                distributors["Винница"] += quantity;
            } else if (distributor === "KDP106") {
                distributors["Сафьяны"] += quantity;
            }
        });

        const totalSum = firmsSum + hutorskayaAlmaSum + distributors["Винница"] + distributors["Сафьяны"];
        return {
            "Форс/Алма": firmsSum.toFixed(2),
            "Хуторская Алма": hutorskayaAlmaSum.toFixed(2),
            "Винница Алма": distributors["Винница"].toFixed(2),
            "Сафьяны": distributors["Сафьяны"].toFixed(2),
            "Сумма": totalSum.toFixed(2)
        };
    }

    // Выводим результаты в окно
    function displayResults(results) {
        outputDiv.innerHTML = `
            ________________________
            <br>
            Форс/Алма - ${results["Форс/Алма"]}
            <br>
            Хуторская Алма - ${results["Хуторская Алма"]}
            <br>
            Винница Алма - ${results["Винница Алма"]}
            <br>
            Сафьяны - ${results["Сафьяны"]}
            <br>
            ________________________
            <br>
            Сумма - ${results["Сумма"]}
        `;
    }
});
