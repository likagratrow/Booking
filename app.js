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
let bookingSlotById=new Map();
let bookingModal=null;

const CALENDAR_START_HOUR=10;
const CALENDAR_END_HOUR=20;
const CALENDAR_DAYS=30;

const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;').replace(/'/g,'&#039;');
function parse(t){const a=t.indexOf('{'),b=t.lastIndexOf('}');if(a<0||b<=a)throw Error('Google Таблица не вернула данные.');return JSON.parse(t.slice(a,b+1));}
function cell(c,i,f=''){return c[i]?.v??f;}
function norm(v){return String(v??'').trim().toLowerCase();}
function pad(n){return String(n).padStart(2,'0');}
function dateText(d){return d.toLocaleDateString('ru-RU',{weekday:'short',day:'numeric',month:'short'}).replace(' г.','');}
function addDays(d,n){const x=new Date(d);x.setDate(x.getDate()+n);return x;}
function dateString(d){return`${pad(d.getDate())}.${pad(d.getMonth()+1)}.${d.getFullYear()}`;}
function timeKey(minutes){return`${pad(Math.floor(minutes/60))}:${pad(minutes%60)}`;}
function parseDuration(value){
  const raw=norm(value).replace(',','.');
  if(!raw)return 60;
  const hours=raw.match(/(\d+(?:\.\d+)?)\s*(?:ч|час|часа|часов)/);
  if(hours)return Math.max(60,Math.round(Number(hours[1])*60));
  const minutes=raw.match(/(\d+(?:\.\d+)?)\s*(?:мин|минут|минута)/);
  if(minutes)return Math.max(60,Math.round(Number(minutes[1])));
  const number=raw.match(/\d+(?:\.\d+)?/);
  return number?Math.max(60,Math.round(Number(number[0])*60)):60;
}
function slotDateTime(value){
  const raw=String(value??'').trim();
  let m=raw.match(/^(\d{2})\.(\d{2})\.(\d{4})[ T](\d{1,2}):(\d{2})/);
  if(m)return{date:`${m[1]}.${m[2]}.${m[3]}`,minutes:Number(m[4])*60+Number(m[5])};
  m=raw.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{1,2}):(\d{2})/);
  if(m)return{date:`${m[3]}.${m[2]}.${m[1]}`,minutes:Number(m[4])*60+Number(m[5])};
  const d=new Date(raw);
  return Number.isNaN(d.getTime())?{date:'',minutes:NaN}:{date:dateString(d),minutes:d.getHours()*60+d.getMinutes()};
}
function slotDate(s){return slotDateTime(s?.start).date;}
function slotTime(s){const m=slotDateTime(s?.start).minutes;return Number.isFinite(m)?timeKey(m):'';}
function slotMinutes(s){return slotDateTime(s?.start).minutes;}
function slotEndMinutes(s){return slotDateTime(s?.end).minutes;}
function activityByKey(key){const k=norm(key);return activities.find(a=>norm(a.key)===k)||null;}
function eventBySelection(){return events.find(e=>norm(e.activity)===norm(selectedCalendar.activity)&&(!selectedCalendar.name||norm(e.name)===norm(selectedCalendar.name)))||null;}
function isDiogenActivity(key){
  const activity=activityByKey(key);
  if(!activity)return false;
  if(norm(activity.title)==='диоген'||norm(activity.key)==='диоген')return true;
  return bookingSlots.some(s=>norm(s.activity)===norm(key)&&norm(s.format)==='диоген');
}
function activityPaletteClass(key){
  const matches=bookingSlots.find(s=>norm(s.activity)===norm(key));
  if(matches?.format&&norm(matches.format)==='диоген')return'dio';
  const activity=activityByKey(key);
  if(norm(activity?.title)==='диоген'||norm(key)==='диоген')return'dio';
  if(matches?.format&&norm(matches.format)==='мк')return'mk';
  return'mk';
}
function slotHasStarted(slot){
  const p=slotDateTime(slot?.start);if(!p.date||!Number.isFinite(p.minutes))return false;
  const [d,m,y]=p.date.split('.').map(Number);
  return new Date(y,m-1,d,Math.floor(p.minutes/60),p.minutes%60,0,0).getTime()<=Date.now();
}
function slotAvailable(slot,tickets=1){return !!slot&&!slotHasStarted(slot)&&slot.available!==false&&Number(slot.free??0)>=Number(tickets||1);}
function slotCapacity(slot){const capacity=Number(slot?.capacity);if(Number.isFinite(capacity)&&capacity>0)return capacity;const free=Number(slot?.free||0),occupied=Number(slot?.occupied||0);return free+occupied;}
function slotOccupied(slot){const occupied=Number(slot?.occupied);if(Number.isFinite(occupied))return occupied;return Math.max(0,slotCapacity(slot)-Number(slot?.free||0));}

function buildCalendarGeometry(){
  const first=new Date();first.setHours(12,0,0,0);
  const days=Array.from({length:CALENDAR_DAYS},(_,i)=>addDays(first,i));
  const hours=Array.from({length:CALENDAR_END_HOUR-CALENDAR_START_HOUR},(_,i)=>CALENDAR_START_HOUR+i);
  calendarGeometry={days,hours};return calendarGeometry;
}
function ensureCalendarGeometry(){return calendarGeometry||buildCalendarGeometry();}

async function loadActivities(){
  activitiesContainer.innerHTML='<div class="loading-activities">Загрузка...</div>';
  try{const r=await fetch(SHEET_URL,{cache:'no-store'}),j=parse(await r.text());activities=(j.table.rows||[]).map((x,i)=>{const c=x.c||[];return{key:String(cell(c,0,'')).trim(),title:String(cell(c,1,'')).trim(),description:String(cell(c,2,'')).trim(),image:String(cell(c,3,'')).trim(),index:i};}).filter(x=>x.key&&x.title);renderActivities();}
  catch(e){activitiesContainer.innerHTML=`<div class="load-error">Не удалось загрузить активности: ${esc(e.message)}</div>`;}
}
async function loadEvents(){
  try{const r=await fetch(EVENTS_SHEET_URL,{cache:'no-store'}),j=parse(await r.text());events=(j.table.rows||[]).map((x,i)=>{const c=x.c||[];return{activity:String(cell(c,0,'')).trim(),name:String(cell(c,1,'')).trim(),price:String(cell(c,2,'')).trim(),age:String(cell(c,3,'')).trim(),duration:String(cell(c,4,'')).trim(),complexity:String(cell(c,5,'')).trim(),image:String(cell(c,6,'')).trim(),description:String(cell(c,8,'')).trim(),index:i};}).filter(x=>x.activity&&x.name);}
  catch(e){console.error('Ошибка загрузки ивентов:',e);events=[];}
  buildBookingIndexes();renderActivities();
}
async function loadBookingData(){
  try{const r=await fetch(BOOKING_API_URL,{cache:'no-store'});if(!r.ok)throw Error(`Ошибка сервера: ${r.status} ${r.statusText}`);const text=await r.text();let j;try{j=JSON.parse(text);}catch(e){throw Error(`Сервер вернул не JSON: ${text.slice(0,200)}`);}if(!j.ok)throw Error(j.error||'Не удалось получить календарь.');if(!Array.isArray(j.slots))throw Error('Сервер вернул данные без массива слотов.');bookingSlots=j.slots;buildBookingIndexes();renderCalendar();return j;}
  catch(e){console.error('Ошибка загрузки календаря:',e);bookingSlots=[];buildBookingIndexes();renderCalendar();showMessage(`Не удалось загрузить календарь: ${e.message}`);throw e;}
}
function buildBookingIndexes(){bookingSlotById=new Map();bookingSlots.forEach(s=>{if(s?.slotId)bookingSlotById.set(String(s.slotId),s);});}

function renderActivities(){
  if(!activities.length){activitiesContainer.innerHTML='<div class="load-error">Не удалось найти активности.</div>';return;}
  const icons=['◌','✦','✎','◆','●','◇'];
  activitiesContainer.innerHTML=activities.map((a,i)=>`<button class="choice" type="button" data-action="${esc(a.key)}"><span class="choice-icon">${icons[i%icons.length]}</span><span><strong>${esc(a.title)}</strong><small>${esc(a.description)}</small></span><span class="choice-calendar"><span class="choice-calendar-icon">▣</span><span class="choice-calendar-arrow">↓</span></span></button>`).join('');
  activitiesContainer.querySelectorAll('.choice').forEach(button=>button.addEventListener('click',e=>{const key=button.dataset.action;if(e.target.closest('.choice-calendar'))openEventChooser(key);else openActivity(key);}));
}
function openActivity(key){const activity=activityByKey(key);if(!activity)return;activityTitle.textContent=activity.title;activityDescription.textContent=activity.description;if(activity.image){activityImage.src=activity.image;activityImage.alt=activity.title;activityImageWrap.classList.remove('hidden');}else activityImageWrap.classList.add('hidden');activityModal.classList.remove('hidden');}
function stars(v){const n=Math.max(0,Math.min(5,parseInt(v,10)||0));return[1,2,3,4,5].map(i=>`<span class="complexity-star${i>5-n?' filled':''}">★</span>`).join('');}
function openEventChooser(key){
  key=String(key||'').trim();const activity=activityByKey(key);if(!activity)return;const list=events.filter(x=>norm(x.activity)===norm(key));
  if(list.length===1){closeEvents();goToCalendar(key,list[0].name);return;}
  if(!list.length){closeEvents();goToCalendar(key,'');return;}
  eventsTitle.textContent=activity.title;
  eventsList.innerHTML=list.map((event,i)=>`<article class="event-card"><div class="event-image-wrap ${event.image?'':'empty'}">${event.image?`<img class="event-image" src="${esc(event.image)}" alt="${esc(event.name)}">`:'Фото пока нет'}<div class="event-overlay"><div class="event-text-block"><h3>${esc(event.name)}</h3><strong class="event-price">${esc(event.price||'—')} р</strong></div><div class="event-age"><strong>${esc(event.age||'—')}</strong></div><div class="event-duration"><strong>${esc(event.duration||'—')}</strong></div><div class="event-complexity"><strong class="complexity-stars">${stars(event.complexity)}</strong></div></div><div class="event-actions"><button class="event-details" type="button" data-i="${i}">Подробнее</button><button class="event-book" type="button" data-i="${i}">Записаться</button></div></div><div class="event-description hidden" data-i="${i}">${esc(event.description||'Описание пока придумываем')}</div></article>`).join('');
  eventsList.querySelectorAll('.event-details').forEach(button=>button.addEventListener('click',()=>{const description=eventsList.querySelector(`.event-description[data-i="${button.dataset.i}"]`);if(!description)return;const open=!description.classList.contains('hidden');description.classList.toggle('hidden');button.textContent=open?'Подробнее':'Свернуть';}));
  eventsList.querySelectorAll('.event-book').forEach(button=>button.addEventListener('click',()=>{const event=list[Number(button.dataset.i)];closeEvents();goToCalendar(key,event?.name||'');}));eventsModal.classList.remove('hidden');
}
function closeEvents(){eventsModal.classList.add('hidden');}

function slotForCell(date,hour){const start=hour*60;return bookingSlots.find(s=>slotDate(s)===date&&Number.isFinite(slotMinutes(s))&&Number.isFinite(slotEndMinutes(s))&&slotMinutes(s)<=start&&slotEndMinutes(s)>start)||null;}
function selectedEvent(){return eventBySelection();}
function selectedDuration(){const e=selectedEvent();return isDiogenActivity(selectedCalendar.activity)?60:parseDuration(e?.duration||'60');}
function selectedFormat(){const existing=bookingSlots.find(s=>norm(s.activity)===norm(selectedCalendar.activity));return existing?.format||(isDiogenActivity(selectedCalendar.activity)?'Диоген':(activityByKey(selectedCalendar.activity)?.title||'МК'));}
function rangeIsFree(date,start,end,ignoreSlot=null){for(let m=start;m<end;m+=60){const slot=slotForCell(date,Math.floor(m/60));if(slot&&slot!==ignoreSlot)return false;}return end<=CALENDAR_END_HOUR*60;}
function selectedStartValid(date,start){if(!selectedCalendar.activity)return false;if(slotForCell(date,Math.floor(start/60)))return false;const duration=selectedDuration();return rangeIsFree(date,start,start+duration);}
function diogenEndOptions(date,start){const options=[];for(let end=start+60;end<=CALENDAR_END_HOUR*60;end+=60){if(!rangeIsFree(date,start,end))break;options.push(end);}return options;}
function syntheticSlot(date,start,end){return{slotId:`new-${date}-${start}`,activity:selectedCalendar.activity,name:selectedCalendar.name||activityByKey(selectedCalendar.activity)?.title||'Запись',format:selectedFormat(),start:`${date} ${timeKey(start)}`,end:`${date} ${timeKey(end)}`,available:true,free:1,minTickets:1,capacity:1};}
function slotVisual(slot,continuation){const cls=activityPaletteClass(slot.activity);const occupied=slotOccupied(slot),capacity=slotCapacity(slot);return`<button type="button" class="calendar-cell calendar-booked ${cls}${continuation?' continuation':''}" data-slot-id="${esc(slot.slotId)}"><span class="calendar-cell-time">${continuation?'':esc(slotTime(slot))}–${continuation?'':esc(timeKey(Math.min(CALENDAR_END_HOUR*60,slotEndMinutes(slot))))}</span>${continuation?'':`<strong>${esc(slot.name||'Без названия')}</strong><small>${occupied}/${capacity}</small>`}</button>`;}
function renderFreeCell(date,hour){const start=hour*60;const valid=selectedStartValid(date,start);return`<button type="button" class="calendar-cell calendar-free ${valid?'selected-start':''}" data-date="${esc(date)}" data-hour="${hour}"><span class="calendar-cell-time">${timeKey(start)}–${timeKey(start+60)}</span></button>`;}
function renderCalendar(){
  const geometry=ensureCalendarGeometry();
  const dayHeads=geometry.days.map(d=>`<div class="calendar-day-head">${esc(dateText(d))}</div>`).join('');
  const columns=geometry.days.map(d=>{const date=dateString(d);const cells=geometry.hours.map(hour=>{const slot=slotForCell(date,hour);if(slot){const start=slotMinutes(slot);return slotVisual(slot,start!==hour*60);}return renderFreeCell(date,hour);}).join('');return`<div class="calendar-column">${cells}</div>`;}).join('');
  calendar.innerHTML=`<div class="calendar-scroll"><div class="calendar-grid"><div class="calendar-corner"></div><div class="calendar-days">${dayHeads}</div><div class="calendar-times"></div><div class="calendar-columns">${columns}</div></div></div>`;
  calendar.querySelectorAll('.calendar-booked').forEach(el=>el.addEventListener('click',()=>{const slot=bookingSlotById.get(el.dataset.slotId);if(!slot)return;if(slotAvailable(slot,1))openBooking(slot);else showMessage('На это время свободных мест нет.');}));
  calendar.querySelectorAll('.calendar-free').forEach(el=>el.addEventListener('click',()=>handleFreeCell(el.dataset.date,Number(el.dataset.hour))));
}
function promptActivityChoice(date,hour){activitiesContainer.querySelectorAll('.choice').forEach(x=>x.classList.add('choice-highlight'));activitiesContainer.scrollIntoView({behavior:'smooth',block:'center'});showMessage(`Что выберете? ${date} · ${timeKey(hour*60)}`);setTimeout(()=>activitiesContainer.querySelectorAll('.choice').forEach(x=>x.classList.remove('choice-highlight')),2500);}
function handleFreeCell(date,hour){const start=hour*60;if(!selectedCalendar.activity){promptActivityChoice(date,hour);return;}if(!selectedStartValid(date,start)){showMessage('Этот старт не подходит: занят нужный диапазон или запись не помещается до 20:00.');return;}openNewBooking(date,start);}
function goToCalendar(key,name=''){selectedCalendar={activity:String(key||'').trim(),name:String(name||'').trim()};renderCalendar();document.querySelector('.calendar-card')?.scrollIntoView({behavior:'smooth',block:'start'});}
function showMessage(text){message.textContent=text;message.classList.remove('hidden');clearTimeout(showMessage.timer);showMessage.timer=setTimeout(()=>message.classList.add('hidden'),6000);}

function ensureBookingModal(){if(bookingModal)return bookingModal;const wrap=document.createElement('div');wrap.className='modal hidden';wrap.innerHTML='<div class="modal-content"><button class="modal-close booking-close" type="button" aria-label="Закрыть">×</button><p class="eyebrow">ЗАПИСЬ</p><h2 class="booking-title"></h2><div class="booking-body"></div><button class="modal-ok booking-submit" type="button">Записаться</button></div>';document.body.appendChild(wrap);wrap.querySelector('.booking-close').addEventListener('click',()=>wrap.classList.add('hidden'));wrap.addEventListener('click',e=>{if(e.target===wrap)wrap.classList.add('hidden');});bookingModal=wrap;return wrap;}
function bookingPayloadSlot(date,start,end){return syntheticSlot(date,start,end);}
function openNewBooking(date,start){const slot=bookingPayloadSlot(date,start,start+60);if(isDiogenActivity(selectedCalendar.activity)){const options=diogenEndOptions(date,start);openBooking(slot,{diogen:true,endOptions:options});}else openBooking(slot,{diogen:false,endOptions:[start+selectedDuration()]});}
function openBooking(slot,options={}){
  const modal=ensureBookingModal(),title=modal.querySelector('.booking-title'),body=modal.querySelector('.booking-body'),submit=modal.querySelector('.booking-submit');
  const start=slotMinutes(slot),existingEnd=slotEndMinutes(slot),diogen=options.diogen??(isDiogenActivity(slot.activity));
  let endOptions=options.endOptions?.length?options.endOptions:[existingEnd];
  if(diogen&&!options.endOptions){endOptions=[start+60];for(let end=start+120;end<=CALENDAR_END_HOUR*60;end+=60){const conflict=bookingSlots.some(other=>other!==slot&&slotDate(other)===slotDate(slot)&&slotMinutes(other)<=end-60&&slotEndMinutes(other)>end-60);if(conflict)break;endOptions.push(end);}}
  const maxTickets=Math.max(1,Number(slot.free||1)),minTickets=Math.max(1,Number(slot.minTickets||1));
  title.textContent=slot.name||'Запись';
  body.innerHTML=`<p style="margin:0 0 10px;color:var(--tg-hint);font-size:14px">${esc(slotDate(slot))} · ${esc(timeKey(start))}</p><label style="display:block;margin-bottom:10px;font-size:13px">Количество мест<select class="booking-tickets" style="display:block;width:100%;margin-top:5px;padding:9px;border:1px solid var(--tg-separator);border-radius:8px;background:var(--tg-secondary-bg);color:var(--tg-text)">${Array.from({length:Math.max(1,maxTickets-minTickets+1)},(_,i)=>{const n=minTickets+i;return`<option value="${n}">${n}</option>`}).join('')}</select></label>${endOptions.length>1?`<label style="display:block;margin-bottom:12px;font-size:13px">Время окончания<select class="booking-end" style="display:block;width:100%;margin-top:5px;padding:9px;border:1px solid var(--tg-separator);border-radius:8px;background:var(--tg-secondary-bg);color:var(--tg-text)">${endOptions.map(end=>`<option value="${timeKey(end)}">${timeKey(end)}</option>`).join('')}</select></label>`:`<p style="margin:0 0 12px;color:var(--tg-hint);font-size:13px">Окончание: ${timeKey(endOptions[0])}</p>`}<p class="booking-hint" style="margin:0;color:var(--tg-hint);font-size:12px">${diogen?'Можно занять несколько последовательных часов.':''}</p>`;
  submit.disabled=false;submit.textContent='Записаться';
  submit.onclick=async()=>{const tickets=Number(body.querySelector('.booking-tickets')?.value||1);const endTime=body.querySelector('.booking-end')?.value||timeKey(endOptions[0]);const end=slotDateTime(`${slotDate(slot)} ${endTime}`).minutes;if(end<=start||end>CALENDAR_END_HOUR*60){showMessage('Неверное время окончания.');return;}submit.disabled=true;submit.textContent='Проверяем...';try{await loadBookingData();const baseSlot=bookingSlotById.get(String(slot.slotId))||null;if(!rangeIsFree(slotDate(slot),start,end,baseSlot)){throw Error('Это время уже занято. Календарь обновлён.');}const fresh=bookingPayloadSlot(slotDate(slot),start,end);await apiBook(fresh,tickets,endTime);modal.classList.add('hidden');showMessage('Запись создана.');await loadBookingData();}catch(e){showMessage(e.message||'Не удалось создать запись.');submit.disabled=false;submit.textContent='Записаться';}};
  modal.classList.remove('hidden');
}
async function apiBook(slot,tickets,endTime=''){if(!user?.id)throw Error('Не удалось определить Telegram ID. Откройте запись внутри Telegram.');const payload={action:'book',telegramId:String(user.id),name:String(user.first_name||user.username||'гость'),format:slot.format,eventName:slot.name,date:slotDate(slot),time:slotTime(slot),tickets:Number(tickets),endTime};const r=await fetch(BOOKING_API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload)}),j=await r.json();if(!j.ok)throw Error(j.error||'Не удалось создать бронь.');return j;}

function closeActivity(){activityModal.classList.add('hidden');}
modalClose?.addEventListener('click',closeActivity);modalOk?.addEventListener('click',closeActivity);eventsClose?.addEventListener('click',closeEvents);activityModal?.addEventListener('click',e=>{if(e.target===activityModal)closeActivity();});eventsModal?.addEventListener('click',e=>{if(e.target===eventsModal)closeEvents();});
document.addEventListener('keydown',e=>{if(e.key!=='Escape')return;activityModal?.classList.add('hidden');eventsModal?.classList.add('hidden');bookingModal?.classList.add('hidden');});
(async()=>{try{await Promise.all([loadActivities(),loadEvents(),loadBookingData()]);}catch(e){console.error(e);}})();
