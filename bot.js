const { google } = require('googleapis');
const TelegramBot = require('node-telegram-bot-api');

// Настройки
const SPREADSHEET_ID = 'example_sheet_id';
const SHEET_NAME = 'Sheet1';
const TELEGRAM_TOKEN = 'example_bot_token'; // Замените на ваш токен
const CHAT_ID = 'example_used_id'; // Замените на ваш chat_id

// Инициализация бота
const bot = new TelegramBot(TELEGRAM_TOKEN, { polling: true });

// Авторизация в Google Sheets
const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json',
    scopes: 'https://www.googleapis.com/auth/spreadsheets',
});

const sheets = google.sheets({ version: 'v4', auth });

// Функция для проверки, содержит ли диапазон текст "Denis B"
async function containsDenisB(range) {
    try {
        const res = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!${range}`,
            majorDimension: 'ROWS', // Получаем данные по строкам
        });

        const values = res.data.values;

        // Если данные отсутствуют, возвращаем false
        if (!values || values.length === 0) {
            return false;
        }

        // Проверяем, есть ли текст "Denis B" в любой ячейке диапазона
        return values.some(row => row.some(cell => cell.includes('Denis B')));
    } catch (error) {
        console.error('Ошибка при проверке текста "Denis B":', error);
        return false;
    }
}

// Функция для проверки диапазона на пустые ячейки
async function checkRange(range) {
    try {
        const res = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!${range}`,
            majorDimension: 'COLUMNS', // Указываем, что хотим получить данные по столбцам
        });

        const values = res.data.values;

        // Если данные отсутствуют, считаем диапазон пустым
        if (!values || values.length === 0) {
            return true;
        }

        // Проверяем, есть ли пустые ячейки
        return values.some(column => column.some(cell => cell === '' || cell === null));
    } catch (error) {
        console.error('Ошибка при проверке диапазона:', error);
        return false;
    }
}

// Функция для проверки конкретной ячейки (B16)
async function checkCellB16() {
    try {
        const res = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!B16`, // Проверяем только ячейку B16
        });

        const values = res.data.values;

        // Если данные отсутствуют или ячейка пустая, возвращаем true
        if (!values || values.length === 0 || values[0][0] === '' || values[0][0] === null) {
            return true;
        }

        return false;
    } catch (error) {
        console.error('Ошибка при проверке ячейки B16:', error);
        return false;
    }
}

// Функция для проверки конкретной ячейки (E16)
async function checkCellE16() {
    try {
        const res = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!E16`, // Проверяем только ячейку E16
        });

        const values = res.data.values;

        // Если данные отсутствуют или ячейка пустая, возвращаем true
        if (!values || values.length === 0 || values[0][0] === '' || values[0][0] === null) {
            return true;
        }

        return false;
    } catch (error) {
        console.error('Ошибка при проверке ячейки E16:', error);
        return false;
    }
}

// Основная функция для проверки пустых ячеек
async function checkEmptyCells() {
    try {
        const ranges = ['B2:B16', 'E2:E16'];

        // Проверяем, есть ли текст "Denis B" в указанных диапазонах
        let hasDenisB = false;
        for (const range of ranges) {
            if (await containsDenisB(range)) {
                hasDenisB = true;
                break;
            }
        }

        if (hasDenisB) {
            console.log('Текст "Denis B" найден. Уведомления отключены.');
            return; // Прекращаем выполнение, если текст найден
        }

        // Если текст "Denis B" не найден, проверяем пустые ячейки
        let isEmpty = false;

        // Проверяем диапазоны
        for (const range of ranges) {
            if (await checkRange(range)) {
                isEmpty = true;
                break;
            }
        }

        // Отдельно проверяем ячейку B16
        if (!isEmpty && (await checkCellB16())) {
            isEmpty = true;
        }

        // Отдельно проверяем ячейку E16
        if (!isEmpty && (await checkCellE16())) {
            isEmpty = true;
        }

        if (isEmpty) {
            console.log('Обнаружены пустые ячейки!');
            await bot.sendMessage(CHAT_ID, 'Обнаружены пустые ячейки!');
        } else {
            console.log('Пустых ячеек не обнаружено.');
        }
    } catch (error) {
        console.error('Ошибка в checkEmptyCells:', error);
    }
}

// Проверка каждые 10 секунд (для теста)
setInterval(checkEmptyCells, 10000);

// Первая проверка при запуске
checkEmptyCells();