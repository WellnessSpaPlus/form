document.getElementById('contactForm').addEventListener('submit', function(e) {
    e.preventDefault();

    // Проверка reCAPTCHA (если есть)
    if (typeof grecaptcha !== 'undefined' && grecaptcha.getResponse() === '') {
        alert('Подтвердите, что вы не робот!');
        return;
    }

    // 1. Получаем текущие данные из localStorage (или создаём новый массив)
    let allRequests = JSON.parse(localStorage.getItem('requests') || []);

    // 2. Добавляем новую заявку
    const newRequest = {
        Имя: document.getElementById('name').value,
        Email: document.getElementById('email').value,
        Телефон: document.getElementById('phone').value,
        Сообщение: document.getElementById('message').value,
        Дата: new Date().toLocaleString('ru-RU'), // Формат даты для России
    };
    allRequests.push(newRequest);

    // 3. Сохраняем обновлённый массив в localStorage
    localStorage.setItem('requests', JSON.stringify(allRequests));

    // 4. Генерируем Excel-файл и предлагаем сохранить
    generateExcel(allRequests);
    
    // 5. Очищаем форму
    this.reset();
    if (typeof grecaptcha !== 'undefined') grecaptcha.reset();
});

// Функция для создания Excel-файла
function generateExcel(data) {
    // Динамически загружаем библиотеку SheetJS
    const script = document.createElement('script');
    script.src = 'https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js';
    script.onload = () => {
        // Создаём книгу Excel
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, 'Заявки');

        // Генерируем файл и скачиваем
        XLSX.writeFile(wb, 'Все_заявки.xlsx');
        alert('Файл сохранён в папку "Загрузки"!');
    };
    document.head.appendChild(script);
}