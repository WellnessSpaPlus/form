document.getElementById('contactForm').addEventListener('submit', async function(e) {
    e.preventDefault();

    // Проверка reCAPTCHA (если есть)
    if (typeof grecaptcha !== 'undefined' && grecaptcha.getResponse() === '') {
        alert('Подтвердите, что вы не робот!');
        return;
    }

    // Получаем д    анные формы
    const formData = {
        Имя: document.getElementById('name').value,
        Email: document.getElementById('email').value,
        Телефон: document.getElementById('phone').value,
        Сообщение: document.getElementById('message').value,
        Дата: new Date().toLocaleString(),
    };

    // Динамически загружаем библиотеку docx
    const script = document.createElement('script');
    script.src = 'https://unpkg.com/docx@8.0.0/build/index.js';
    script.onload = async () => {
        const { Document, Paragraph, TextRun, Packer } = docx;
        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Новая заявка", bold: true, size: 28 }),
                        ],
                    }),
                    new Paragraph({ text: `Имя: ${formData.Имя}` }),
                    new Paragraph({ text: `Email: ${formData.Email}` }),
                    new Paragraph({ text: `Телефон: ${formData.Телефон}` }),
                    new Paragraph({ text: `Сообщение: ${formData.Сообщение}` }),
                    new Paragraph({ text: `Дата: ${formData.Дата}` }),
                ],
            }],
        });

        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = `Заявка_${new Date().getTime()}.docx`;
        a.click();
        
        alert('Данные сохранены в Word!');
        this.reset();
        if (typeof grecaptcha !== 'undefined') grecaptcha.reset();
    };
    document.head.appendChild(script);
});