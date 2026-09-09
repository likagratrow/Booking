const tg = window.Telegram?.WebApp;

if (tg) {
  tg.ready();
  tg.expand();
}

const user = tg?.initDataUnsafe?.user;
const hello = document.getElementById('hello');

if (user) {
  hello.textContent = `Здравствуйте, ${user.first_name || 'гость'}!`;
}

const calendar = document.getElementById('calendar');
const message = document.getElementById('message');

function localDate(offset) {
  const d = new Date();
  d.setHours(12, 0, 0, 0);
  d.setDate(d.getDate() + offset);
  return d;
}

function dateLabel(date) {
  return date.toLocaleDateString('ru-RU', {
    day: 'numeric',
    month: 'long',
    weekday: 'long'
  });
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

    return `
      <article class="day">
        <div class="day-title">
          <span>${dateLabel(date)}</span>
        </div>
        <div class="slots">${slots}</div>
      </article>
    `;
  }).join('');
}

renderCalendar();

document.querySelectorAll('.choice').forEach(button => {
  button.addEventListener('click', () => {
    const type = button.dataset.action === 'diogen' ? 'Диоген' : 'Мастер-класс';
    showMessage(`Вы выбрали «${type}». Следующим шагом здесь появятся доступные варианты записи.`);
  });
});

document.querySelectorAll('.slot').forEach(button => {
  button.addEventListener('click', () => {
    if (button.dataset.state === 'closed') {
      showMessage('Это время закрыто. Выберите другой час.');
      return;
    }
    showMessage(`Вы выбрали ${button.dataset.time}. Пока это демонстрационный календарь — запись ещё не отправляется.`);
  });
});

function showMessage(text) {
  message.textContent = text;
  message.classList.remove('hidden');
  message.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
