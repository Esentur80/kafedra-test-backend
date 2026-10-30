require('dotenv').config();
const express = require('express');
const multer = require('multer');
const mammoth = require('mammoth');
const { createClient } = require('@supabase/supabase-js');
const { GoogleGenAI } = require('@google/genai');

const app = express();
const PORT = process.env.PORT || 5000;

// ==========================================
// 1. ИНИЦИАЛИЗАЦИЯ КЛИЕНТОВ И MIDDLEWARE
// ==========================================

// Парсинг JSON в теле запросов
app.use(express.json());

// Инициализация Supabase
if (!process.env.SUPABASE_URL || !process.env.SUPABASE_ANON_KEY) {
  console.error('Ошибка: SUPABASE_URL или SUPABASE_ANON_KEY не заданы в .env');
  process.exit(1);
}
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_ANON_KEY);

// Инициализация Google Gen AI SDK
if (!process.env.GEMINI_API_KEY) {
  console.error('Ошибка: GEMINI_API_KEY не задан в .env');
  process.exit(1);
}
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Настройка Multer (хранение файлов в оперативной памяти Buffer для mammoth)
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 15 * 1024 * 1024 }, // лимит 15 МБ
  fileFilter: (req, file, cb) => {
    // Проверка расширения и MIME-типа для .docx
    if (
      file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      file.originalname.endsWith('.docx')
    ) {
      cb(null, true);
    } else {
      cb(new Error('Разрешены только файлы Word (.docx)'), false);
    }
  }
});

// ==========================================
// 2. ЭНДПОИНТЫ API
// ==========================================

/**
 * 1. Healthcheck / Ping
 * GET /ping
 */
app.get('/ping', (req, res) => {
  return res.status(200).send('pong');
});

/**
 * 2. Генерация / парсинг теста через Gemini AI
 * POST /api/tests/generate-ai
 * Принимает multipart/form-data:
 *   - file: .docx файл (опционально)
 *   - prompt: текстовое описание / тема (опционально)
 */
app.post('/api/tests/generate-ai', upload.single('file'), async (req, res) => {
  try {
    let sourceContent = '';

    // Если передан Word-документ — извлекаем сырой текст через mammoth
    if (req.file) {
      const { value: text } = await mammoth.extractRawText({ buffer: req.file.buffer });
      sourceContent = text.trim();
    }

    // Если передан текст/промпт в теле запроса — добавляем его
    if (req.body.prompt) {
      sourceContent += `\nДополнительные требования/тема: ${req.body.prompt}`;
    }

    if (!sourceContent) {
      return res.status(400).json({
        error: 'Необходимо предоставить файл .docx или текстовый промпт для генерации теста'
      });
    }

    // Системная инструкция для модели Gemini
    const systemInstruction = `
Ты — опытный методист и академический эксперт кафедры университета.
Твоя задача — составить или извлечь из предоставленного материала академический тест СТРОГО ИЗ 35 ВОПРОСОВ.
Каждый вопрос должен иметь ровно 4 варианта ответов и ровно 1 правильный ответ.
Индекс правильного ответа (correct_index) должен быть числом от 0 до 3 (где 0 — первый вариант, 3 — четвертый).

Ты ОБЯЗАН вернуть результат строго в формате JSON-массива без какого-либо дополнительного текста, markdown-блоков (без \`\`\`json) и комментариев.
Формат каждого объекта:
{
  "question": "Текст вопроса",
  "options": [
    "Вариант 1",
    "Вариант 2",
    "Вариант 3",
    "Вариант 4"
  ],
  "correct_index": 0
}
Всего объектов в массиве должно быть РОВНО 35.
    `.trim();

    // Запрос к Gemini
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: `Исходный материал:\n${sourceContent}`,
      config: {
        systemInstruction,
        responseMimeType: 'application/json',
        temperature: 0.2 // низкая температура для строгости соблюдения формата
      }
    });

    let rawJsonText = response.text;
    if (!rawJsonText) {
      throw new Error('Пустой ответ от модели Gemini');
    }

    // Парсим результат
    const questions = JSON.parse(rawJsonText);

    if (!Array.isArray(questions)) {
      throw new Error('Сгенерированный результат не является JSON-массивом');
    }

    return res.status(200).json({
      success: true,
      count: questions.length,
      questions
    });
  } catch (error) {
    console.error('Ошибка в /api/tests/generate-ai:', error);
    return res.status(500).json({
      error: 'Не удалось сгенерировать тест',
      details: error.message
    });
  }
});

/**
 * 3. Сохранение теста преподавателем в Supabase
 * POST /api/tests
 * Body JSON: {
 *   teacher_id: "uuid/string",
 *   title: "Модульный контроль 1",
 *   module_number: 1,
 *   questions_json: [ ...35 вопросов... ],
 *   open_at: "2026-10-30T10:00:00Z",
 *   close_at: "2026-10-30T11:30:00Z"
 * }
 */
app.post('/api/tests', async (req, res) => {
  try {
    const { teacher_id, title, module_number, questions_json, open_at, close_at } = req.body;

    // Валидация входных данных
    if (!teacher_id || !title || module_number === undefined || !questions_json || !open_at || !close_at) {
      return res.status(400).json({
        error: 'Все поля (teacher_id, title, module_number, questions_json, open_at, close_at) обязательны'
      });
    }

    if (!Array.isArray(questions_json) || questions_json.length !== 35) {
      return res.status(400).json({
        error: `Тест должен содержать ровно 35 вопросов. Получено: ${Array.isArray(questions_json) ? questions_json.length : 0}`
      });
    }

    // Валидация дат
    const openDate = new Date(open_at);
    const closeDate = new Date(close_at);

    if (isNaN(openDate.getTime()) || isNaN(closeDate.getTime())) {
      return res.status(400).json({ error: 'Некорректный формат дат open_at или close_at (ожидается ISO)' });
    }

    if (openDate >= closeDate) {
      return res.status(400).json({ error: 'Время открытия (open_at) должно быть раньше времени закрытия (close_at)' });
    }

    // Вставка в Supabase
    const { data, error } = await supabase
      .from('tests')
      .insert([
        {
          teacher_id,
          title,
          module_number,
          questions_json,
          open_at: openDate.toISOString(),
          close_at: closeDate.toISOString()
        }
      ])
      .select()
      .single();

    if (error) {
      throw error;
    }

    return res.status(201).json({
      success: true,
      message: 'Тест успешно сохранен в базе данных',
      test: data
    });
  } catch (error) {
    console.error('Ошибка в POST /api/tests:', error);
    return res.status(500).json({
      error: 'Ошибка при сохранении теста',
      details: error.message
    });
  }
});

/**
 * 4. Сдача теста студентом с подсчетом баллов
 * POST /api/tests/submit
 * Body JSON: {
 *   test_id: "uuid/string",
 *   student_id: "uuid/string",
 *   answers: [0, 2, 1, 3, ...] // Массив из 35 выбранных индексов ответов
 * }
 */
app.post('/api/tests/submit', async (req, res) => {
  try {
    const { test_id, student_id, answers } = req.body;

    if (!test_id || !student_id || !Array.isArray(answers)) {
      return res.status(400).json({
        error: 'Поля test_id, student_id и массив ответов answers обязательны'
      });
    }

    // Получаем тест из Supabase
    const { data: test, error: fetchError } = await supabase
      .from('tests')
      .select('*')
      .eq('id', test_id)
      .single();

    if (fetchError || !test) {
      return res.status(404).json({ error: 'Тест с указанным ID не найден' });
    }

    // Проверка серверного времени
    const now = new Date();
    const openAt = new Date(test.open_at);
    const closeAt = new Date(test.close_at);

    if (now < openAt) {
      return res.status(403).json({
        error: 'Тест еще не открыт для прохождения',
        open_at: test.open_at,
        server_time: now.toISOString()
      });
    }

    if (now > closeAt) {
      return res.status(403).json({
        error: 'Время прохождения теста истекло',
        close_at: test.close_at,
        server_time: now.toISOString()
      });
    }

    const questions = test.questions_json;
    if (!Array.isArray(questions)) {
      return res.status(500).json({ error: 'Ошибка структуры вопросов теста в базе данных' });
    }

    // Подсчет баллов (1 вопрос = 1 балл, максимум 35)
    let score = 0;
    const details = [];

    questions.forEach((q, idx) => {
      const studentAnswer = answers[idx];
      const isCorrect = studentAnswer !== undefined && studentAnswer === q.correct_index;

      if (isCorrect) {
        score += 1;
      }

      details.push({
        question_number: idx + 1,
        student_answer: studentAnswer ?? null,
        is_correct: isCorrect
      });
    });

    // Сохранение попытки в таблицу `results`
    const { data: resultRecord, error: insertError } = await supabase
      .from('results')
      .insert([
        {
          test_id,
          student_id,
          score,
          max_score: questions.length, // 35
          answers_json: answers,
          submitted_at: now.toISOString()
        }
      ])
      .select()
      .single();

    if (insertError) {
      throw insertError;
    }

    return res.status(200).json({
      success: true,
      message: 'Тест успешно сдан',
      result_id: resultRecord.id,
      score,
      max_score: questions.length,
      percentage: Math.round((score / questions.length) * 100),
      submitted_at: resultRecord.submitted_at,
      details
    });
  } catch (error) {
    console.error('Ошибка в POST /api/tests/submit:', error);
    return res.status(500).json({
      error: 'Ошибка при обработке результатов теста',
      details: error.message
    });
  }
});

// Глобальный обработчик ошибок Multer
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError || err.message.includes('Word')) {
    return res.status(400).json({ error: err.message });
  }
  return res.status(500).json({ error: 'Внутренняя ошибка сервера' });
});

// Запуск сервера
app.listen(PORT, () => {
  console.log(`Сервер кафедры успешно запущен на порту ${PORT}`);
  console.log(`Healthcheck доступен по адресу: http://localhost:${PORT}/ping`);
});