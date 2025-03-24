document.getElementById('contactForm').addEventListener('submit', function(e) {
    e.preventDefault();

    // Проверка reCAPTCHA
    if (typeof grecaptcha !== 'undefined' && grecaptcha.getResponse() === '') {
        alert('Подтвердите, что вы не робот!');
        return;
    }

    // Получаем текущие данные из localStorage (если есть)
    let allRequests = JSON.parse(localStorage.getItem('requests') || []);

    // Добавляем новую заявку
    const newRequest = {
        Имя: document.getElementById('name').value,
        Email: document.getElementById('email').value,
        Телефон: document.getElementById('phone').value,
        Сообщение: document.getElementById('message').value,
        Дата: new Date().toLocaleString(),
    };
    allRequests.push(newRequest);

    // Сохраняем в localStorage
    localStorage.setItem('requests', JSON.stringify(allRequests));

    // Загружаем библиотеку SheetJS
    const script = document.createElement('script');
    script.src = 'https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js';
    script.onload = () => {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(allRequests);
        XLSX.utils.book_append_sheet(wb, ws, 'Все заявки');
        XLSX.writeFile(wb, 'Все_заявки.xlsx');
        
        alert('Данные добавлены в общий файл!');
        this.reset();
        if (typeof grecaptcha !== 'undefined') grecaptcha.reset();
    };
    document.head.appendChild(script);
});