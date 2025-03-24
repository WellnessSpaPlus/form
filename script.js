document.getElementById('contactForm').addEventListener('submit', function(e) {
    e.preventDefault();

    // Проверка reCAPTCHA (если есть)
    if (typeof grecaptcha !== 'undefined' && grecaptcha.getResponse() === '') {
        alert('Подтвердите, что вы не робот!');
        return;
    }

    // Получаем данные формы
    const formData = {
        Имя: document.getElementById('name').value,
        Email: document.getElementById('email').value,
        Телефон: document.getElementById('phone').value,
        Сообщение: document.getElementById('message').value,
        Дата: new Date().toLocaleString(),
    };

    // Загружаем библиотеку SheetJS динамически
    const script = document.createElement('script');
    script.src = 'https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js';
    script.onload = () => {
        // Создаем книгу Excel
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet([formData]);

        // Добавляем лист в книгу
        XLSX.utils.book_append_sheet(wb, ws, 'Заявки');

        // Генерируем файл и скачиваем
        XLSX.writeFile(wb, 'Заявка_' + new Date().getTime() + '.xlsx');
        
        alert('Данные сохранены в Excel!');
        this.reset();
        if (typeof grecaptcha !== 'undefined') grecaptcha.reset();
    };
    document.head.appendChild(script);
});