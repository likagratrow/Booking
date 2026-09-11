const SHEET_ID='1FcetqNVvXNI78h0mcQdEJBEVXzkHcgaddFrCn2VOugk';
const SHEET_URL=`https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent('Активности')}&headers=0&range=A2:D`;
const EVENTS_SHEET_URL=`https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent('Ивенты')}&headers=0&range=A2:I`;
const BOOKING_API_URL='https://script.google.com/macros/s/AKfycbwpYRqIunVa1uBfMqaf9HY4ICKXeeaCUbYLtoBxYT6e0_tYyBSvgiVfmeU-3SkEC2SqxQ/exec';

const tg=window.Telegram?.WebApp;
if(tg){tg.ready();tg.expand();tg.setBackgroundColor?.('bg_color');}
const user=tg?.initDataUnsafe?.user;

const activitiesContainer=document.getElementById('activities');
const calendar=document.getElementById('calendar');
const message=document.getElementById('message');
const activityModal=document.getElementById('activity-modal');
const activityTitle=document.getElementById('activity-title');
const activityDescription=document.getElementById('activity-description');
const activityImageWrap=document.getElementById('activity-image-wrap');
const activityImage=document.getElementById('activity-image');
const modalClose=document.getElementById('modal-close');
const modalOk=document.getElementById('modal-ok');
const eventsModal=document.getElementById('events-modal');
const eventsTitle=document.getElementById('events-title');
const eventsList=document.getElementById('events-list');
const eventsClose=document.getElementById('events-close');

let activities=[];
let events=[];
let bookingSlots=[];
let selectedCalendar={activity:'',name:''};
let calendarGeometry=null;
let bookingSlotByKey=new Map();
let bookingSlotById=new Map();
let bookingState=null;

const CALENDAR_START_HOUR=10;
const CALENDAR_END_HOUR=20;
const CALENDAR_DAYS=30;

const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;').replace(/'/g,'&#039;');
function parse(t){const a=t.indexOf('{'),b=t.lastIndexOf('}');if(a<0||b<=a)throw Error('Google Таблица не вернула данные.');return JSON.parse(t.slice(a,b+1));}
function cell(c,i,f=''){return c[i]?.v??f;}
function norm(v){return String(v||'').trim().toLowerCase();}
function dateKey(s){const [d,m,y]=String(s).split(' ')[0].split('.').map(Number);return new Date(y,m-1,d,12,0,0);}
function dateText(d){return d.toLocaleDateString('ru-RU',{weekday:'short',day:'numeric',month:'short'}).replace(' г.','');}
function addDays(d,n){const x=new Date(d);x.setDate(x.getDate()+n);return x;}
function pad(n){return String(n).padStart(2,'0');}
function dateString(d){return`${pad(d.getDate())}.${pad(d.getMonth()+1)}.${d.getFullYear()}`;}
function timeKey(minutes){return`${pad(Math.floor(minutes/60))}:${pad(minutes%60)}`;}
function slotKey(date,time){return`${date}|${time}`;}
function slotTime(s){return String(s?.start||'').split(' ')[1]||'';}
function slotMinutes(s){const [h,m]=slotTime(s).split(':').map(Number);return Number.isFinite(h)&&Number.isFinite(m)?h*60+m:NaN;}
function slotEndMinutes(s){const [h,m]=String(s?.end||'').split(' ')[1]?.split(':').map(Number)||[];return Number.isFinite(h)&&Number.isFinite(m)?h*60+m:NaN;}
function bookingKey(s){return slotKey(String(s?.start||'').split(' ')[0],slotTime(s));}

function activityOf(slot){
  const activity=norm(slot?.activity);
  if(activity==='диоген')return'Диоген';
  if(activity==='обсудить заказ')return'Обсудить заказ';
  if(activity==='masterclass'||activity==='мастер-класс')return'МК';
  const format=norm(slot?.format);
  if(format==='диоген')return'Диоген';
  if(format==='обсудить заказ')return'Обсудить заказ';
  if(format==='мк'||format==='masterclass'||format==='мастер-класс')return'МК';
  return String(slot?.format||'').trim();
}
function isFlexibleActivity(activity){return activity==='Диоген'||activity==='Обсудить заказ';}
function selectedMatches(slot){
  if(!selectedCalendar.activity)return false;
  if(activityOf(slot)!==selectedCalendar.activity)return false;
  return selectedCalendar.activity!=='МК'||norm(slot.name)===norm(selectedCalendar.name);
}
function slotAvailable(slot,tickets=1){return !!slot&&slot.available&&Number(slot.free||0)>=tickets;}

function weekStart(){const today=new Date();today.setHours(12,0,0,0);return today;}
function buildCalendarGeometry(){
  const first=weekStart();
  const days=Array.from({length:CALENDAR_DAYS},(_,i)=>addDays(first,i));
  const hours=Array.from({length:CALENDAR_END_HOUR-CALENDAR_START_HOUR},(_,i)=>CALENDAR_START_HOUR+i);
  const cells=new Map();
  days.forEach(d=>{
    const date=dateString(d);
    hours.forEach(hour=>{
      const start=timeKey(hour*60);
      cells.set(slotKey(date,start),{key:slotKey(date,start),date,start,end:timeKey((hour+1)*60),hour});
    });
  });
  calendarGeometry={days,hours,cells};
  return calendarGeometry;
}
function ensureCalendarGeometry(){return calendarGeometry||buildCalendarGeometry();}

async function loadActivities(){
  activitiesContainer.innerHTML='<div class="loading">Загрузка активностей...</div>';
  try{
    const r=await fetch(SHEET_URL,{cache:'no-store'});
    const j=parse(await r.text());
    activities=(j.table.rows||[]).map((x,i)=>{
      const c=x.c||[];
      return{key:String(cell(c,0,`activity-${i}`)).trim().toLowerCase(),title:String(cell(c,1,'Без названия')).trim(),description:String(cell(c,2,'')).trim(),image:String(cell(c,3,'')).trim()};
    }).filter(x=>x.key&&x.title);
    renderActivities();
  }catch(e){
    activitiesContainer.innerHTML=`<div class="load-error">Не удалось загрузить активности: ${esc(e.message)}</div>`;
  }
}

async function loadEvents(){
  try{
    const r=await fetch(EVENTS_SHEET_URL,{cache:'no-store'});
    const j=parse(await r.text());
    events=(j.table.rows||[]).map((x,i)=>{
      const c=x.c||[];
      return{activity:String(cell(c,0,'')).trim(),name:String(cell(c,1,'Без названия')).trim(),price:String(cell(c,2,'')).trim(),age:String(cell(c,3,'')).trim(),duration:String(cell(c,4,'')).trim(),complexity:String(cell(c,5,'')).trim(),image:String(cell(c,6,'')).trim(),description:String(cell(c,8,'')).trim(),index:i};
    }).filter(x=>x.activity&&x.name);
  }catch(e){
    console.error('Ошибка загрузки ивентов:',e);
    events=[];
  }
  buildBookingIndexes();
  renderActivities();
}

async function loadBookingData(){
  try{
    const r=await fetch(BOOKING_API_URL,{cache:'no-store'});
    if(!r.ok)throw Error(`Ошибка сервера: ${r.status} ${r.statusText}`);
    const text=await r.text();
    let j;
    try{j=JSON.parse(text);}catch(e){throw Error(`Сервер вернул не JSON: ${text.slice(0,200)}`);}
    if(!j.ok)throw Error(j.error||'Не удалось получить календарь.');
    if(!Array.isArray(j.slots))throw Error('Сервер вернул данные без массива слотов.');
    bookingSlots=j.slots;
    buildBookingIndexes();
    renderCalendar();
    return j;
  }catch(e){
    console.error('Ошибка загрузки календаря:',e);
    bookingSlots=[];
    buildBookingIndexes();
    renderCalendar();
    showMessage(`Не удалось загрузить календарь: ${e.message}`);
    throw e;
  }
}

async function apiBook(slot,tickets,endTime=''){
  if(!user?.id)throw Error('Не удалось определить Telegram ID. Откройте запись внутри Telegram.');
  const payload={action:'book',telegramId:String(user.id),name:String(user.first_name||user.username||'гость'),format:slot.format,eventName:slot.name,date:slot.start.split(' ')[0],time:slot.start.split(' ')[1],tickets:Number(tickets),endTime};
  const r=await fetch(BOOKING_API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload)});
  const j=await r.json();
  if(!j.ok)throw Error(j.error||'Не удалось создать бронь.');
  return j;
}

function buildBookingIndexes(){
  bookingSlotByKey=new Map();
  bookingSlotById=new Map();
  bookingSlots.forEach(slot=>{
    const start=slotMinutes(slot),end=slotEndMinutes(slot),date=String(slot?.start||'').split(' ')[0];
    if(!Number.isFinite(start)||!Number.isFinite(end)||!date||!slot?.slotId)return;
    bookingSlotById.set(String(slot.slotId),slot);
    for(let t=start;t<Math.max(start+60,end);t+=60){
      const key=slotKey(date,timeKey(t));
      if(!bookingSlotByKey.has(key))bookingSlotByKey.set(key,[]);
      bookingSlotByKey.get(key).push(slot);
    }
  });
}

function durationHours(event){
  const match=String(event?.duration||'').match(/\d+(?:[.,]\d+)?/);
  return match?Math.max(1,Number(match[0].replace(',','.'))):1;
}
function eventForSelectedMasterclass(){return events.find(e=>norm(e.activity)==='masterclass'&&norm(e.name)===norm(selectedCalendar.name));}
function flexibleSlotsAt(date,start,activity){return bookingSlotByKey.get(slotKey(date,timeKey(start)))?.filter(s=>activityOf(s)===activity)||[];}
function masterclassSlotAt(date,start){return bookingSlotByKey.get(slotKey(date,timeKey(start)))?.find(s=>activityOf(s)==='МК'&&norm(s.name)===norm(selectedCalendar.name));}

function masterclassFits(startSlot,tickets){
  const event=eventForSelectedMasterclass();
  if(!event||!slotAvailable(startSlot,tickets))return false;
  const needed=durationHours(event);
  const start=slotMinutes(startSlot);
  const date=String(startSlot.start||'').split(' ')[0];
  if(start+needed*60>CALENDAR_END_HOUR*60)return false;
  for(let i=0;i<needed;i++){
    const part=masterclassSlotAt(date,start+i*60);
    if(!slotAvailable(part,tickets))return false;
  }
  return true;
}
function flexibleFits(startSlot,tickets=1){return isFlexibleActivity(activityOf(startSlot))&&slotAvailable(startSlot,tickets);}
function candidateForCell(cell){
  const matches=bookingSlotByKey.get(cell.key)||[];
  if(!selectedCalendar.activity)return null;
  const slot=matches.find(s=>selectedMatches(s)&&slotAvailable(s,1));
  if(!slot)return null;
  if(selectedCalendar.activity==='МК')return masterclassFits(slot,Number(slot.occupied||0)>0?1:2)?slot:null;
  return flexibleFits(slot,1)?slot:null;
}
function hasFreeSlot(cell){return(bookingSlotByKey.get(cell.key)||[]).some(s=>slotAvailable(s,1));}

function renderActivities(){
  if(!activities.length){activitiesContainer.innerHTML='<div class="load-error">Не удалось найти активности.</div>';return;}
  activitiesContainer.innerHTML=activities.map(a=>`<button class="choice" type="button" data-action="${esc(a.key)}"><span class="choice-icon">${a.key==='diogen'?'◌':a.key==='masterclass'?'✦':'✎'}</span><span><strong>${esc(a.title)}</strong><small>${esc(a.description)}</small></span><span class="choice-calendar"><span class="choice-calendar-icon">▣</span><span class="choice-calendar-arrow">↓</span></span></button>`).join('');
  activitiesContainer.querySelectorAll('.choice').forEach(button=>button.addEventListener('click',e=>e.target.closest('.choice-calendar')?openEventChooser(button.dataset.action):openActivity(button.dataset.action)));
}
function openActivity(key){
  const activity=activities.find(x=>x.key===key);
  if(!activity)return;
  activityTitle.textContent=activity.title;
  activityDescription.textContent=activity.description;
  if(activity.image){activityImage.src=activity.image;activityImage.alt=activity.title;activityImageWrap.classList.remove('hidden');}else activityImageWrap.classList.add('hidden');
  activityModal.classList.remove('hidden');
}
function stars(v){const n=Math.max(0,Math.min(5,parseInt(v,10)||0));return[1,2,3,4,5].map(i=>`<span class="complexity-star${i>5-n?' filled':''}">★</span>`).join('');}
function openEventChooser(key){
  key=norm(key);
  if(key==='order'){closeEvents();goToCalendar('order');return;}
  const activity=activities.find(x=>norm(x.key)===key);
  if(!activity)return;
  const list=events.filter(x=>norm(x.activity)===key);
  eventsTitle.textContent=activity.title;
  eventsList.innerHTML=list.length?list.map((event,i)=>`<article class="event-card"><div class="event-image-wrap ${event.image?'':'empty'}">${event.image?`<img class="event-image" src="${esc(event.image)}" alt="${esc(event.name)}">`:'Фото пока нет'}<div class="event-overlay"><div class="event-text-block"><h3>${esc(event.name)}</h3><strong class="event-price">${esc(event.price||'—')} р</strong></div><div class="event-age"><strong>${esc(event.age||'—')}</strong></div><div class="event-duration"><strong>${esc(event.duration||'—')}</strong></div><div class="event-complexity"><strong class="complexity-stars">${stars(event.complexity)}</strong></div></div><div class="event-actions"><button class="event-details" type="button" data-i="${i}">Подробнее</button><button class="event-book" type="button" data-i="${i}">Записаться</button></div></div><div class="event-description hidden" data-i="${i}">${esc(event.description||'Описание пока придумываем')}</div></article>`).join(''):'<div class="events-empty">Пока нет доступных вариантов.</div>';
  eventsList.querySelectorAll('.event-details').forEach(button=>button.addEventListener('click',()=>{
    const description=eventsList.querySelector(`.event-description[data-i="${button.dataset.i}"]`);
    if(!description)return;
    const open=!description.classList.contains('hidden');
    description.classList.toggle('hidden');
    button.textContent=open?'Подробнее':'Свернуть';
  }));
  eventsList.querySelectorAll('.event-book').forEach(button=>button.addEventListener('click',()=>{
    const event=list[+button.dataset.i];
    closeEvents();
    goToCalendar(key,event.name);
  }));
  eventsModal.classList.remove('hidden');
}
function closeEvents(){eventsModal.classList.add('hidden');}

function renderCalendar(){
  const geometry=ensureCalendarGeometry();
  calendar.innerHTML=`<div class="calendar-scroll"><div class="calendar-grid"><div class="calendar-corner"></div><div class="calendar-days">${geometry.days.map(d=>`<div class="calendar-day-head">${esc(dateText(d))}</div>`).join('')}</div><div class="calendar-times"></div><div class="calendar-columns">${geometry.days.map(d=>{const date=dateString(d);return`<div class="calendar-column"><div class="calendar-hour-lines">${geometry.hours.map(hour=>{
    const cell=geometry.cells.get(slotKey(date,timeKey(hour*60)));
    const free=hasFreeSlot(cell);
    const slot=candidateForCell(cell);
    const candidate=!!slot;
    return`<div class="calendar-cell">${free?`<button class="calendar-slot free${candidate?' candidate':''}${selectedCalendar.activity&&!candidate?' selected-unavailable':''}" type="button" data-slot-id="${esc(slot?.slotId||'')}" data-generic="${selectedCalendar.activity?'0':'1'}" ${selectedCalendar.activity&&!candidate?'disabled':''}><strong>${esc(timeKey(hour*60))}–${esc(timeKey((hour+1)*60))}</strong><small>свободно</small></button>`:''}</div>`;
  }).join('')}</div></div>`;}).join('')}</div></div></div>`;
  calendar.querySelectorAll('.calendar-slot').forEach(button=>button.addEventListener('click',()=>{
    if(button.dataset.generic==='1'){flashActivityChoices();return;}
    const slot=bookingSlotById.get(button.dataset.slotId);
    if(slot)selectSlot(slot);
  }));
}
function flashActivityChoices(){
  const card=activitiesContainer.closest('.card')||activitiesContainer;
  card.classList.remove('calendar-activity-flash');void card.offsetWidth;card.classList.add('calendar-activity-flash');
  setTimeout(()=>card.classList.remove('calendar-activity-flash'),950);
}
function ensureCalendarStyles(){
  if(document.getElementById('calendar-clean-styles'))return;
  const style=document.createElement('style');style.id='calendar-clean-styles';
  style.textContent=`
    .calendar-card{min-height:0;overflow:hidden}.calendar{min-height:0;height:100%;flex:1 1 0;overflow:hidden}.calendar-scroll{height:100%;min-height:0;overflow-x:auto;overflow-y:hidden}.calendar-grid{grid-template-columns:3120px;width:3120px;min-width:3120px;height:100%;min-height:0}.calendar-days,.calendar-columns{width:3120px}.calendar-columns,.calendar-column,.calendar-hour-lines{height:100%;min-height:0}.calendar-slot:disabled{cursor:default}.calendar-slot.candidate{box-shadow:0 0 0 2px color-mix(in srgb,var(--tg-accent) 70%,transparent)}.calendar-slot.selected-unavailable{opacity:.45}.calendar-activity-flash{animation:calendarActivityFlash .9s ease-in-out 1}.booking-time-wrap{display:flex;align-items:center;justify-content:center;gap:3px}.booking-duration-inline{display:flex!important;flex-direction:column;align-items:center;justify-content:center;gap:0;margin-left:1px;line-height:1}.booking-duration-inline .booking-arrow{width:16px!important;height:13px!important;min-height:13px!important;padding:0!important;border:0!important;background:transparent!important;font-size:9px!important;line-height:10px!important;color:var(--tg-hint)!important}.booking-duration-inline .booking-arrow:disabled{opacity:.25}.booking-tickets-control{display:grid;grid-template-columns:36px minmax(70px,1fr) 36px;align-items:center;gap:6px;width:100%}.booking-tickets-control #booking-tickets-input{width:100%;text-align:center}.booking-ticket-step{width:36px;height:36px;padding:0;color:var(--tg-text);background:var(--tg-secondary-bg);border:1px solid var(--tg-separator);border-radius:8px;font-size:20px;line-height:1;cursor:pointer}.booking-max-mark{display:inline-flex;align-items:center;justify-content:center;margin-left:6px;padding:2px 4px;color:var(--tg-accent);border:1px solid var(--tg-accent);border-radius:4px;font-size:8px;font-weight:800;line-height:1;vertical-align:middle}#booking-modal .eyebrow.hidden{display:none}@keyframes calendarActivityFlash{0%,100%{box-shadow:none}15%,35%{box-shadow:0 0 0 3px color-mix(in srgb,var(--tg-accent) 75%,transparent),0 0 18px color-mix(in srgb,var(--tg-accent) 35%,transparent)}50%,70%{box-shadow:none}}
  `;
  document.head.appendChild(style);
}

function ensureBookingModal(){
  if(document.getElementById('booking-modal'))return;
  document.body.insertAdjacentHTML('beforeend',`<div id="booking-modal" class="modal hidden" role="dialog" aria-modal="true" aria-labelledby="booking-title"><div class="modal-content booking-modal-content"><button id="booking-close" class="modal-close" type="button" aria-label="Закрыть">×</button><p class="eyebrow hidden">ЗАПИСЬ</p><h2 id="booking-title">Записаться</h2><div id="booking-activity" class="booking-activity"></div><div class="booking-date-row"><div class="booking-time-wrap"><button id="booking-start" class="booking-time" type="button"></button><span class="booking-dash">–</span><button id="booking-end" class="booking-time" type="button"></button></div><div id="booking-date" class="booking-date"></div></div><div id="booking-duration-controls" class="booking-duration-controls hidden"><button id="booking-end-up" class="booking-arrow" type="button" aria-label="Увеличить длительность">↑</button><button id="booking-end-down" class="booking-arrow" type="button" aria-label="Уменьшить длительность">↓</button></div><div class="booking-tickets"><label for="booking-tickets-input">Количество мест</label><input id="booking-tickets-input" type="number" min="1" value="1"></div><button id="booking-submit" class="modal-ok" type="button">Записаться</button></div></div>`);
  const modal=document.getElementById('booking-modal');
  document.getElementById('booking-close').addEventListener('click',closeBooking);
  modal.addEventListener('click',e=>{if(e.target===modal)closeBooking();});
  document.getElementById('booking-submit').addEventListener('click',submitBooking);
  document.getElementById('booking-tickets-input').addEventListener('input',renderBookingModal);
  document.getElementById('booking-end-up').addEventListener('click',()=>changeFlexibleEnd(1));
  document.getElementById('booking-end-down').addEventListener('click',()=>changeFlexibleEnd(-1));
}
function bookingMinTickets(){if(!bookingState)return 1;return Number(bookingState.slot.occupied||0)>0?1:(activityOf(bookingState.slot)==='МК'?2:1);}
function bookingMaxTickets(){
  if(!bookingState)return 0;
  const {slot,start,end}=bookingState,activity=activityOf(slot),day=String(slot.start||'').split(' ')[0];
  let max=Infinity;
  for(let t=start;t<end;t+=60){
    const part=activity==='МК'?masterclassSlotAt(day,t):flexibleSlotsAt(day,t,activity).find(s=>slotAvailable(s,1));
    if(!part)return 0;
    max=Math.min(max,Number(part.free||0));
  }
  return Number.isFinite(max)?Math.max(0,max):0;
}
function availableFlexibleEnd(slot,tickets=1){
  const activity=activityOf(slot),day=String(slot.start||'').split(' ')[0],start=slotMinutes(slot),ends=[];
  for(let end=start+60;end<=CALENDAR_END_HOUR*60;end+=60){
    const part=flexibleSlotsAt(day,end-60,activity).find(s=>slotAvailable(s,tickets));
    if(!part)break;ends.push(end);
  }
  return ends;
}
function ensureTicketControls(){
  const input=document.getElementById('booking-tickets-input');if(!input||document.getElementById('booking-tickets-minus'))return;
  const wrap=document.createElement('div');wrap.className='booking-tickets-control';input.parentNode.insertBefore(wrap,input);wrap.appendChild(input);
  const minus=document.createElement('button');minus.id='booking-tickets-minus';minus.type='button';minus.className='booking-ticket-step';minus.textContent='−';minus.setAttribute('aria-label','Уменьшить количество');
  const plus=document.createElement('button');plus.id='booking-tickets-plus';plus.type='button';plus.className='booking-ticket-step';plus.textContent='+';plus.setAttribute('aria-label','Увеличить количество');
  wrap.insertBefore(minus,input);wrap.appendChild(plus);
  minus.addEventListener('click',()=>{const min=Number(input.min)||1;input.value=String(Math.max(min,(Number(input.value)||min)-1));input.dispatchEvent(new Event('input',{bubbles:true}));});
  plus.addEventListener('click',()=>{const max=Number(input.max)||1;input.value=String(Math.min(max,(Number(input.value)||0)+1));input.dispatchEvent(new Event('input',{bubbles:true}));});
}
function setupDurationControls(){const controls=document.getElementById('booking-duration-controls'),wrap=document.querySelector('#booking-modal .booking-time-wrap');if(!controls||!wrap)return;if(controls.parentElement!==wrap)wrap.appendChild(controls);controls.classList.add('booking-duration-inline');}
function renderBookingModal(){
  if(!bookingState)return;
  const {slot,start,end}=bookingState,min=bookingMinTickets(),max=bookingMaxTickets(),activity=activityOf(slot);
  document.getElementById('booking-start').textContent=timeKey(start);document.getElementById('booking-end').textContent=timeKey(end);document.getElementById('booking-date').textContent=formatBookingDate(slot.start);
  document.getElementById('booking-title').textContent=activity==='МК'?'Мастер-класс':activity;document.getElementById('booking-submit').textContent=activity==='МК'?'Мастер-класс':activity;
  setupDurationControls();
  const controls=document.getElementById('booking-duration-controls'),flexible=isFlexibleActivity(activity);controls.classList.toggle('hidden',!flexible);
  const up=document.getElementById('booking-end-up'),down=document.getElementById('booking-end-down');
  if(flexible){const input=document.getElementById('booking-tickets-input'),tickets=Math.max(min,Number(input?.value)||min),ends=availableFlexibleEnd(slot,tickets);up.disabled=!ends.includes(end+60);down.disabled=end<=start+60;}
  ensureTicketControls();
  const input=document.getElementById('booking-tickets-input');input.min=String(min);input.max=String(Math.max(min,max));
  let value=Number(input.value);if(!Number.isInteger(value)||value<min)value=min;if(value>max)value=max;input.value=String(value);
  const oldMark=document.getElementById('booking-max-mark');if(oldMark)oldMark.remove();
  if(max>=min&&value===max){const mark=document.createElement('span');mark.id='booking-max-mark';mark.className='booking-max-mark';mark.textContent='MAX';input.parentElement.appendChild(mark);}
}
function formatBookingDate(s){const d=dateKey(s);return d.toLocaleDateString('ru-RU',{day:'numeric',month:'long'});}
function openBooking(slot){
  ensureBookingModal();
  const activity=activityOf(slot),start=slotMinutes(slot),defaultEnd=Math.min(start+60,slotEndMinutes(slot));
  bookingState={slot,start,end:isFlexibleActivity(activity)?defaultEnd:slotEndMinutes(slot),tickets:Number(slot.occupied||0)>0?1:(activity==='МК'?2:1)};
  document.getElementById('booking-activity').textContent=activity==='МК'?(selectedCalendar.name||slot.name||'Мастер-класс'):activity;
  const input=document.getElementById('booking-tickets-input');if(input)input.value=String(bookingState.tickets);
  renderBookingModal();document.getElementById('booking-modal').classList.remove('hidden');
}
function changeFlexibleEnd(direction){
  if(!bookingState||!isFlexibleActivity(activityOf(bookingState.slot)))return;
  const min=bookingMinTickets(),input=document.getElementById('booking-tickets-input'),tickets=Math.max(min,Number(input?.value)||min),ends=availableFlexibleEnd(bookingState.slot,tickets),target=bookingState.end+direction*60;
  if(target<bookingState.start+60||!ends.includes(target))return;bookingState.end=target;renderBookingModal();
}
function closeBooking(){const modal=document.getElementById('booking-modal');if(modal)modal.classList.add('hidden');bookingState=null;}
function selectSlot(slot){
  if(!selectedCalendar.activity){flashActivityChoices();return;}
  if(!slotAvailable(slot,1)){showMessage('В этом слоте больше нет свободных мест.');return;}
  if(!selectedMatches(slot)){showMessage('Выберите слот выбранной активности.');return;}
  const min=Number(slot.occupied||0)>0?1:(activityOf(slot)==='МК'?2:1);
  if(activityOf(slot)==='МК'&&!masterclassFits(slot,min)){showMessage('Этот слот не подходит по длительности мастер-класса.');return;}
  if(isFlexibleActivity(activityOf(slot))&&!flexibleFits(slot,1)){showMessage('В этом месте нет свободного времени.');return;}
  openBooking(slot);
}
async function submitBooking(){
  if(!bookingState)return;
  const {slot,start,end}=bookingState,input=document.getElementById('booking-tickets-input'),tickets=Number(input.value)||0,min=bookingMinTickets(),max=bookingMaxTickets();
  if(!Number.isInteger(tickets)||tickets<min||tickets>max){showMessage(`Введите целое число от ${min} до ${max}.`);return;}
  const endTime=timeKey(end);closeBooking();showMessage('Создаём бронь...');
  try{const result=await apiBook({...slot,start:`${slot.start.split(' ')[0]} ${timeKey(start)}`},tickets,endTime);showMessage(`Бронь №${result.booking.id} создана: ${result.booking.date} ${result.booking.time}, ${result.booking.tickets} мест.`);await loadBookingData();}catch(e){showMessage(`Не удалось записаться: ${e.message}`);}
}
function goToCalendar(type,name=''){
  if(type==='masterclass')selectedCalendar={activity:'МК',name};
  else if(type==='diogen')selectedCalendar={activity:'Диоген',name:''};
  else if(type==='order')selectedCalendar={activity:'Обсудить заказ',name:''};
  else selectedCalendar={activity:'',name:''};
  renderCalendar();if(selectedCalendar.activity)showMessage(`Расписание: ${name||selectedCalendar.activity}.`);
  setTimeout(()=>document.querySelector('.calendar-card')?.scrollIntoView({behavior:'smooth',block:'start'}),40);
}
function showMessage(text){message.textContent=text;message.classList.remove('hidden');message.scrollIntoView({behavior:'smooth',block:'nearest'});}
function closeActivity(){activityModal.classList.add('hidden');}

modalClose.addEventListener('click',closeActivity);modalOk.addEventListener('click',closeActivity);activityModal.addEventListener('click',e=>{if(e.target===activityModal)closeActivity();});eventsClose.addEventListener('click',closeEvents);eventsModal.addEventListener('click',e=>{if(e.target===eventsModal)closeEvents();});
ensureCalendarStyles();buildCalendarGeometry();renderCalendar();window.addEventListener('resize',()=>{calendarGeometry=null;buildCalendarGeometry();renderCalendar();});Promise.all([loadActivities(),loadEvents(),loadBookingData()]).catch(()=>{});