const SHEET_ID = '1FcetqNVvXNI78h0mcQdEJBEVXzkHcgaddFrCn2VOugk';

// Та же схема, что работает в Shop, но берём вкладку Активности.
const SHEET_URL =
  `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent('Активности')}`;

const tg = window.Telegram?.WebApp;

if (tg) {
  tg.ready();
  tg.expand();
  tg.setBackgroundColor?.('bg_color');
}

const user = tg?.initDataUnsafe?.user;
const hello = document.getElementById('hello');

if (user) {
  hello.textContent = `Здравствуйте, ${user.first_name || 'гость'}!`;
}

const activitiesContainer = document.getElementById('activities');
const calendar = document.getElementById('calendar');
const message = document.getElementById('message');
const activityModal = document.getElementById('activity-modal');
const activityTitle = document.getElementById('activity-title');
const activityDescription = document.getElementById('activity-description');
const activityImageWrap = document.getElementById('activity-image-wrap');
const activityImage = document.getElementById('activity-image');
const modalClose = document.getElementById('modal-close');
const modalOk = document.getElementById('modal-ok');

const fallbackActivities = [
  {
    key: 'diogen',
    title: 'Диоген',
    description: 'Можно прийти в мастерскую и провести время в своём ритме: поработать над чем-то своим, потискать кожу, попить чаю, посидеть в тишине или вообще ничего не делать.',
    image: ''
  },
  {
    key: 'masterclass',
    title: 'Мастер-класс',
    description: 'Вы выбираете изделие и приходите делать его вместе с мастером. Мастер-класс проходит в заданное время и длится столько, сколько указано в его описании.',
    image: ''
  },
  {
    key: 'order',
    title: 'Обсудить заказ',
    description: 'Если вы хотите сделать изделие на заказ, можно прийти в мастерскую лично: обсудить задумку, материалы, размеры, детали и все нюансы будущей вещи.',
    image: ''
  }
];

let activities = fallbackActivities;

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function parseGvizResponse(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1 || end <= start) {
    throw new Error('Google Таблица не вернула данные. Проверьте публикацию таблицы в интернете.');
  }
  try {
    return JSON.parse(text.slice(start, end + 1));
  } catch {
    throw new Error('Не удалось разобрать ответ Google Таблицы.');
  }
}

function getCellValue(cells, index, fallback = '') {
  const cell = cells[index];
  return cell && cell.v !== null && cell.v !== undefined ? cell.v : fallback;
}

function getActivityImageUrl(value) {
  const image = String(value || '').trim();
  if (!image) return '';
  if (/^(https?:)?\/\//i.test(image) || image.startsWith('data:')) return image;
  return image;
}

function showLoadError(error) {
  if (!activitiesContainer) return;
  const messageText = error?.message || String(error) || 'Неизвестная ошибка';
  activitiesContainer.innerHTML = `<div class="load-error"><h3>Не удалось загрузить активности</h3><p>${escapeHtml(messageText)}</p><button type="button" onclick="loadActivities()">Повторить</button></div>`;
}

async function loadActivities() {
  if (activitiesContainer) activitiesContainer.innerHTML = '<div class="loading">Загрузка активностей...</div>';
  try {
    const response = await fetch(SHEET_URL, { method: 'GET', cache: 'no-store' });
    if (!response.ok) throw new Error(`Google Sheets вернул HTTP ${response.status}`);
    const text = await response.text();
    const json = parseGvizResponse(text);
    if (!json.table || !Array.isArray(json.table.rows)) {
      throw new Error('В ответе Google Таблицы отсутствуют строки с активностями.');
    }

    activities = json.table.rows.slice(1).map((row, index) => {
      const cells = row.c || [];
      return {
        key: String(getCellValue(cells, 0, `activity-${index}`)).trim().toLowerCase(),
        title: String(getCellValue(cells, 1, 'Без названия')).trim(),
        description: String(getCellValue(cells, 2, '')).trim(),
        image: getActivityImageUrl(getCellValue(cells, 3, ''))
      };
    }).filter(activity => activity.key && activity.title);

    if (!activities.length) throw new Error('Во вкладке «Активности» нет заполненных активностей.');
    console.log('Активности загружены:', activities);
    renderActivities();
  } catch (error) {
    console.error('Ошибка загрузки активностей:', error);
    showLoadError(error);
  }
}

function renderActivities() {
  activitiesContainer.innerHTML = activities.map(activity => `
    <button class="choice" type="button" data-action="${escapeHtml(activity.key)}">
      <span class="choice-icon">${activity.key === 'diogen' ? '◌' : activity.key === 'masterclass' ? '✦' : '✎'}</span>
      <span>
        <strong>${escapeHtml(activity.title)}</strong>
        <small>${escapeHtml(activity.description)}</small>
      </span>
    </button>
  `).join('');

  activitiesContainer.querySelectorAll('.choice').forEach(button => {
    button.addEventListener('click', () => openActivity(button.dataset.action));
  });
}

function openActivity(key) {
  const activity = activities.find(item => item.key === key);
  if (!activity) return;

  activityTitle.textContent = activity.title;
  activityDescription.textContent = activity.description;

  if (activity.image) {
    activityImage.src = activity.image;
    activityImage.alt = activity.title;
    activityImageWrap.classList.remove('hidden');
    activityImage.onerror = () => {
      activityImage.removeAttribute('src');
      activityImageWrap.classList.add('hidden');
    };
  } else {
    activityImage.removeAttribute('src');
    activityImageWrap.classList.add('hidden');
  }

  activityModal.classList.remove('hidden');
}

function localDate(offset) {
  const d = new Date();
  d.setHours(12, 0, 0, 0);
  d.setDate(d.getDate() + offset);
  return d;
}

function dateLabel(date) {
  return date.toLocaleDateString('ru-RU', { day: 'numeric', month: 'long', weekday: 'long' });
}

const demoDays = [
  { offset: 1, slots: [['10:00–11:00', 'free', 'свободно'], ['11:00–12:00', 'free', 'свободно'], ['12:00–13:00', 'dio', 'Диоген · 1/2'], ['13:00–14:00', 'dio', 'Диоген · 1/2'], ['15:00–16:00', 'closed', 'закрыто'], ['16:00–17:00', 'free', 'свободно']] },
  { offset: 2, slots: [['10:00–11:00', 'free', 'свободно'], ['11:00–12:00', 'mk', 'МК · 3/5'], ['12:00–13:00', 'mk', 'МК · 3/5'], ['13:00–14:00', 'mk', 'МК · 3/5'], ['15:00–16:00', 'free', 'свободно'], ['16:00–17:00', 'free', 'свободно']] },
  { offset: 3, slots: [['10:00–11:00', 'closed', 'закрыто'], ['11:00–12:00', 'free', 'свободно'], ['12:00–13:00', 'free', 'свободно'], ['13:00–14:00', 'dio', 'Диоген · 2/2'], ['14:00–15:00', 'dio', 'Диоген · 2/2'], ['16:00–17:00', 'free', 'свободно']] }
];

function renderCalendar() {
  calendar.innerHTML = demoDays.map(day => {
    const date = localDate(day.offset);
    const slots = day.slots.map(([time, state, label]) => `
      <button class="slot ${state}" type="button" data-state="${state}" data-time="${time}">
        <strong>${time}</strong>
        <small>${label}</small>
      </button>
    `).join('');
    return `<article class="day"><div class="day-title"><span>${dateLabel(date)}</span></div><div class="slots">${slots}</div></article>`;
  }).join('');

  document.querySelectorAll('.slot').forEach(button => {
    button.addEventListener('click', () => {
      if (button.dataset.state === 'closed') {
        showMessage('Это время закрыто. Выберите другой час.');
        return;
      }
      showMessage(`Вы выбрали ${button.dataset.time}. Пока это демонстрационный календарь — запись ещё не отправляется.`);
    });
  });
}

function closeActivityModal() { activityModal.classList.add('hidden'); }

modalClose.addEventListener('click', closeActivityModal);
modalOk.addEventListener('click', closeActivityModal);
activityModal.addEventListener('click', event => { if (event.target === activityModal) closeActivityModal(); });

function showMessage(text) {
  message.textContent = text;
  message.classList.remove('hidden');
  message.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

renderCalendar();
loadActivities();
