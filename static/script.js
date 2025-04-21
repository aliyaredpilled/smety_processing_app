// static/script.js

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('upload-form');
    const submitButton = document.getElementById('submit-button');
    const progressContainer = document.getElementById('progress-container');
    const progressBar = document.getElementById('progress-bar');
    const statusMessage = document.getElementById('status-message');
    const resultContainer = document.getElementById('result-container');
    const resultMessage = document.getElementById('result-message');
    const downloadLink = document.getElementById('download-link');
    const errorContainer = document.getElementById('error-container');
    const fileNameDisplay = document.getElementById('file-name');

    // === Новые элементы ===
    const smetaTypeSelect = document.getElementById('smeta_type');
    const fileInput = document.getElementById('file');
    const turbosmetchikVersionGroup = document.getElementById('turbosmetchik-version-group');
    const turbosmetchikVersionSelect = document.getElementById('turbosmetchik_version');
    // =====================

    let progressInterval = null; // Переменная для хранения ID интервала

    // Функция для генерации простого уникального ID сессии
    function generateClientSessionId() {
        return Date.now() + '-' + Math.random().toString(36).substring(2, 15);
    }

    // Функция для запроса прогресса
    async function pollProgress(sessionId) {
        try {
            const response = await fetch(`/progress/${sessionId}`);
            if (!response.ok) {
                // Если сервер не отвечает или ошибка, прекращаем поллинг
                console.error('Ошибка запроса прогресса:', response.status);
                stopPolling();
                return;
            }
            const data = await response.json();

            // Обновляем прогресс бар
            let percentage = 0;
            if (data.total > 0) {
                percentage = Math.round((data.processed / data.total) * 100);
            } else if (data.status && data.status !== "Ошибка" && data.status !== "Готово" && data.status !== "Не найдено") {
                 // Если total еще 0, но идет работа (например, распаковка), показываем небольшой прогресс
                 percentage = 10; // Или другое значение
            }
            progressBar.style.width = `${percentage}%`;
            // Можно добавить текст на прогресс-бар: progressBar.textContent = `${percentage}%`;

            // Обновляем текстовый статус
            if (data.status) {
                statusMessage.textContent = data.status;
            }

            // Если есть ошибка от сервера или статус "Готово" или "Ошибка", останавливаем поллинг
            if (data.error || data.status === "Готово" || data.status === "Ошибка") {
                stopPolling();
                 // Дополнительно можно проверить, если статус Готово,
                 // но основная форма еще не ответила - возможно, нужна доп. логика
            }

        } catch (error) {
            console.error('Сетевая ошибка при запросе прогресса:', error);
            stopPolling(); // Останавливаем при сетевых ошибках
        }
    }

    // Функция для остановки поллинга
    function stopPolling() {
        if (progressInterval) {
            clearInterval(progressInterval);
            progressInterval = null;
            console.log("Поллинг прогресса остановлен.");
        }
    }

    // === Новая функция: Проверка валидности формы для активации кнопки ===
    function checkFormValidity() {
        const mainTypeSelected = smetaTypeSelect.value;
        const fileSelected = fileInput.files.length > 0;
        let versionSelected = true; // По умолчанию true, станет false если Турбосметчик выбран, но версия нет

        if (mainTypeSelected === 'Турбосметчик') {
            turbosmetchikVersionGroup.style.display = 'block'; // Показываем выбор версии
            turbosmetchikVersionSelect.required = true;      // Делаем выбор версии обязательным
            versionSelected = turbosmetchikVersionSelect.value !== '';
        } else {
            turbosmetchikVersionGroup.style.display = 'none';  // Скрываем выбор версии
            turbosmetchikVersionSelect.required = false;     // Снимаем обязательность
            turbosmetchikVersionSelect.value = '';           // Сбрасываем выбор версии
        }

        // Отображаем имя файла, если он выбран
        if (fileSelected) {
            fileNameDisplay.textContent = `Выбран файл: ${fileInput.files[0].name}`;
        } else {
            fileNameDisplay.textContent = ''; // Очищаем имя файла, если ничего не выбрано
        }

        // Активируем кнопку только если все условия выполнены
        submitButton.disabled = !(mainTypeSelected && fileSelected && versionSelected);
    }
    // ===================================================================

    // === Обработка выбора основного типа сметы ===
    smetaTypeSelect.addEventListener('change', checkFormValidity); // Просто вызываем проверку
    // ========================================================

    // === Обработка выбора версии Турбосметчика ===
    turbosmetchikVersionSelect.addEventListener('change', checkFormValidity);
    // =========================================================

    // === Обработка выбора файла ===
    fileInput.addEventListener('change', checkFormValidity);
    // ============================================

    form.addEventListener('submit', async (event) => {
        event.preventDefault();
        stopPolling(); // Останавливаем предыдущий поллинг, если был

        submitButton.disabled = true;
        progressContainer.style.display = 'block';
        progressBar.style.width = '0%';
        progressBar.textContent = ''; // Очищаем текст на баре
        statusMessage.textContent = 'Загрузка файла...';
        resultContainer.style.display = 'none';
        errorContainer.style.display = 'none';
        errorContainer.textContent = '';

        const clientSessionId = generateClientSessionId();
        console.log("Новая сессия:", clientSessionId);

        const formData = new FormData();

        // --- Собираем данные формы --- 
        formData.append('file', fileInput.files[0]); // Добавляем файл
        formData.append('client_session_id', clientSessionId); // Добавляем ID сессии

        // Формируем правильное имя типа сметы
        let finalSmetaType = smetaTypeSelect.value;
        if (finalSmetaType === 'Турбосметчик') {
            finalSmetaType += '-' + turbosmetchikVersionSelect.value;
        }
        formData.append('smeta_type', finalSmetaType); // Добавляем правильный тип сметы
        console.log("Отправляемый тип сметы:", finalSmetaType); // Лог для отладки
        // -----------------------------

        // --- Запускаем поллинг прогресса ---
        progressInterval = setInterval(() => {
            pollProgress(clientSessionId);
        }, 1500); // Запрашивать каждые 1.5 секунды
        // ---------------------------------

        try {
            // Основной запрос на загрузку и обработку
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData,
            });

            // --- Останавливаем поллинг после получения ответа ---
            stopPolling();
            // ---------------------------------------------------

            const data = await response.json();

            // Небольшая задержка перед скрытием прогресс-бара, если успешно
            if (response.ok && data.success) {
                 progressBar.style.width = '100%'; // Показываем 100% перед скрытием
                 statusMessage.textContent = data.message || 'Готово';
                 await new Promise(resolve => setTimeout(resolve, 500));
            }
             progressContainer.style.display = 'none'; // Скрываем прогресс


            if (response.ok && data.success) {
                resultMessage.textContent = data.message || 'Обработка успешно завершена!';
                downloadLink.href = data.download_url;
                let buttonText = "Скачать результат";
                if (data.download_filename) {
                    downloadLink.setAttribute('download', data.download_filename);
                    if (data.download_filename.toLowerCase().endsWith('.zip')) buttonText = "Скачать результаты (ZIP)";
                    else if (data.download_filename.toLowerCase().endsWith('.xlsx') || data.download_filename.toLowerCase().endsWith('.xlsm')) buttonText = "Скачать результат (Excel)";
                    else { const parts = data.download_filename.split('.'); const extension = parts.length > 1 ? parts.pop() : 'файл'; buttonText = `Скачать (${extension.toUpperCase()})`; }
                } else downloadLink.removeAttribute('download');
                downloadLink.textContent = buttonText;
                resultContainer.style.display = 'block';
                errorContainer.style.display = 'none';
            } else {
                throw new Error(data.error || `Ошибка сервера: ${response.status}`);
            }

        } catch (error) {
             // Ловим ошибки основного запроса /upload
            stopPolling(); // Убедимся, что поллинг остановлен при ошибке
            console.error('Ошибка при отправке или обработке:', error);
            progressContainer.style.display = 'none'; // Скрываем прогресс при ошибке
            errorContainer.textContent = `Произошла ошибка: ${error.message}`;
            errorContainer.style.display = 'block';
            resultContainer.style.display = 'none';
        } finally {
            // В любом случае снова проверяем валидность (кнопка может остаться disabled, если форма сброшена)
            checkFormValidity();
            form.reset(); // Сбрасываем форму
            fileNameDisplay.textContent = ''; // Очищаем имя файла при сбросе
        }
    });
});