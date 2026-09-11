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
function slotKey(date,time){return`${date}|${time}`;}
function slotTime(s){return String(s?.start||'').split(' ')[1]||'';}
function slotMinutes(s){const [h,m]=slotTime(s).split(':').map(Number);return Number.isFinite(h)&&Number.isFinite(m)?h*60+m:NaN;}
function slotEndMinutes(s){const p=String(s?.end||'').split(' ')[1]||'';const [h,m]=p.split(':').map(Number);return Number.isFinite(h)&&Number.isFinite(m)?h*60+m:NaN;}
function bookingKey(s){return slotKey(String(s?.start||'').split(' ')[0],slotTime(s));}
function activityByKey(key){const k=norm(key);return activities.find(a=>norm(a.key)===k)||null;}
function eventByKeyAndName(key,name){const k=norm(key),n=norm(name);return events.find(e=>norm(e.activity)===k&&norm(e.name)===n)||null;}
function activityPaletteClass(key){const i=activities.findIndex(a=>norm(a.key)===norm(key));const palette=['dio','mk','discuss'];return i>=0&&i<palette.length?palette[i]:'free';}
function slotHasStarted(slot){
  const raw=String(slot?.start||'');
  const parts=raw.split(' ');
  if(parts.length<2)return false;
  const [d,m,y]=parts[0].split('.').map(Number);
  const [h,mi]=parts[1].split(':').map(Number);
  if(![d,m,y,h,mi].every(Number.isFinite))return false;
  return new Date(y,m-1,d,h,mi,0,0).getTime()<=Date.now();
}
function slotAvailable(slot,tickets=1){return !!slot&&!slotHasStarted(slot)&&slot.available&&Number(slot.free||0)>=Number(tickets||1);}

function buildCalendarGeometry(){
  const first=new Date();
  first.setHours(12,0,0,0);
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
  activitiesContainer.innerHTML='<div class="loading-activities">Загрузка...</div>';
  try{
    const r=await fetch(SHEET_URL,{cache:'no-store'});
    const j=parse(await r.text());
    activities=(j.table.rows||[]).map((x,i)=>{
      const c=x.c||[];
      const key=String(cell(c,0,'')).trim();
      const title=String(cell(c,1,'')).trim();
      return{key,title,description:String(cell(c,2,'')).trim(),image:String(cell(c,3,'')).trim(),index:i};
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
      return{activity:String(cell(c,0,'')).trim(),name:String(cell(c,1,'')).trim(),price:String(cell(c,2,'')).trim(),age:String(cell(c,3,'')).trim(),duration:String(cell(c,4,'')).trim(),complexity:String(cell(c,5,'')).trim(),image:String(cell(c,6,'')).trim(),description:String(cell(c,8,'')).trim(),index:i};
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

function slotsAtCell(cell){return(bookingSlotByKey.get(cell.key)||[]).filter(s=>slotMinutes(s)===cell.hour*60);}
function selectedMatches(slot){
  if(!selectedCalendar.activity)return true;
  if(norm(slot?.activity)!==norm(selectedCalendar.activity))return false;
  return !selectedCalendar.name||norm(slot?.name)===norm(selectedCalendar.name);
}
function candidateForCell(cell){
  return slotsAtCell(cell).find(s=>selectedMatches(s)&&slotAvailable(s,1))||null;
}
function allSlotsForCell(cell){return slotsAtCell(cell).filter(s=>selectedMatches(s));}

function renderActivities(){
  if(!activities.length){activitiesContainer.innerHTML='<div class="load-error">Не удалось найти активности.</div>';return;}
  const icons=['◌','✦','✎','◆','●','◇'];
  activitiesContainer.innerHTML=activities.map((a,i)=>`<button class="choice" type="button" data-action="${esc(a.key)}"><span class="choice-icon">${icons[i%icons.length]}</span><span><strong>${esc(a.title)}</strong><small>${esc(a.description)}</small></span><span class="choice-calendar"><span class="choice-calendar-icon">▣</span><span class="choice-calendar-arrow">↓</span></span></button>`).join('');
  activitiesContainer.querySelectorAll('.choice').forEach(button=>button.addEventListener('click',e=>{
    const key=button.dataset.action;
    if(e.target.closest('.choice-calendar'))openEventChooser(key);else openActivity(key);
  }));
}

function openActivity(key){
  const activity=activityByKey(key);
  if(!activity)return;
  activityTitle.textContent=activity.title;
  activityDescription.textContent=activity.description;
  if(activity.image){activityImage.src=activity.image;activityImage.alt=activity.title;activityImageWrap.classList.remove('hidden');}else activityImageWrap.classList.add('hidden');
  activityModal.classList.remove('hidden');
}

function stars(v){const n=Math.max(0,Math.min(5,parseInt(v,10)||0));return[1,2,3,4,5].map(i=>`<span class="complexity-star${i>5-n?' filled':''}">★</span>`).join('');}

function openEventChooser(key){
  key=String(key||'').trim();
  const activity=activityByKey(key);
  if(!activity)return;
  const list=events.filter(x=>norm(x.activity)===norm(key));
  if(list.length===1){closeEvents();goToCalendar(key,list[0].name);return;}
  if(!list.length){closeEvents();goToCalendar(key,'');return;}
  eventsTitle.textContent=activity.title;
  eventsList.innerHTML=list.map((event,i)=>`<article class="event-card"><div class="event-image-wrap ${event.image?'':'empty'}">${event.image?`<img class="event-image" src="${esc(event.image)}" alt="${esc(event.name)}">`:'Фото пока нет'}<div class="event-overlay"><div class="event-text-block"><h3>${esc(event.name)}</h3><strong class="event-price">${esc(event.price||'—')} р</strong></div><div class="event-age"><strong>${esc(event.age||'—')}</strong></div><div class="event-duration"><strong>${esc(event.duration||'—')}</strong></div><div class="event-complexity"><strong class="complexity-stars">${stars(event.complexity)}</strong></div></div><div class="event-actions"><button class="event-details" type="button" data-i="${i}">Подробнее</button><button class="event-book" type="button" data-i="${i}">Записаться</button></div></div><div class="event-description hidden" data-i="${i}">${esc(event.description||'Описание пока придумываем')}</div></article>`).join('');
  eventsList.querySelectorAll('.event-details').forEach(button=>button.addEventListener('click',()=>{
    const description=eventsList.querySelector(`.event-description[data-i="${button.dataset.i}"]`);
    if(!description)return;
    const open=!description.classList.contains('hidden');
    description.classList.toggle('hidden');
    button.textContent=open?'Подробнее':'Свернуть';
  }));
  eventsList.querySelectorAll('.event-book').forEach(button=>button.addEventListener('click',()=>{
    const event=list[Number(button.dataset.i)];
    closeEvents();
    goToCalendar(key,event?.name||'');
  }));
  eventsModal.classList.remove('hidden');
}
function closeEvents(){eventsModal.classList.add('hidden');}

function renderCalendar(){
  const geometry=ensureCalendarGeometry();
  const dayHeads=geometry.days.map(d=>`<div class="calendar-day-head">${esc(dateText(d))}</div>`).join('');
  const columns=geometry.days.map(d=>{
    const date=dateString(d);
    const lines='<div class="calendar-hour-lines">'+geometry.hours.map(()=>'<div></div>').join('')+'</div>';
    const slots=bookingSlots.filter(s=>String(s?.start||'').split(' ')[0]===date&&selectedMatches(s)&&Number.isFinite(slotMinutes(s))&&Number.isFinite(slotEndMinutes(s))&&slotMinutes(s)>=CALENDAR_START_HOUR*60&&slotMinutes(s)<CALENDAR_END_HOUR*60);
    const rendered=slots.map(slot=>renderSlot(slot)).join('');
    return`<div class="calendar-column">${lines}${rendered}</div>`;
  }).join('');
  calendar.innerHTML=`<div class="calendar-scroll"><div class="calendar-grid"><div class="calendar-corner"></div><div class="calendar-days">${dayHeads}</div><div class="calendar-times"></div><div class="calendar-columns">${columns}</div></div></div>`;
  calendar.querySelectorAll('.calendar-slot[data-slot-id]').forEach(el=>el.addEventListener('click',()=>{
    const slot=bookingSlotById.get(el.dataset.slotId);
    if(slot&&slotAvailable(slot,1))openBooking(slot);else showMessage('Этот слот уже закрыт.');
  }));
}

function renderSlot(slot){
  const start=slotMinutes(slot),end=Math.min(CALENDAR_END_HOUR*60,Math.max(start+60,slotEndMinutes(slot)));
  const top=((start-CALENDAR_START_HOUR*60)/((CALENDAR_END_HOUR-CALENDAR_START_HOUR)*60))*100;
  const height=((end-start)/((CALENDAR_END_HOUR-CALENDAR_START_HOUR)*60))*100;
  const available=slotAvailable(slot,1);
  const cls=available?activityPaletteClass(slot.activity):'closed';
  const title=selectedCalendar.activity||slot.activity?slot.name:'Свободно';
  const free=Number(slot.free||0);
  const meta=available?(free>0?`свободно: ${free}`:'закрыто'):'закрыто';
  return`<button type="button" class="calendar-slot ${cls}" data-slot-id="${esc(slot.slotId)}" style="top:${top}%;height:${Math.max(height,4)}%;" ${available?'':'disabled'}><strong>${esc(title||'Без названия')}</strong><small>${esc(meta)}</small></button>`;
}

function goToCalendar(key,name=''){
  selectedCalendar={activity:String(key||'').trim(),name:String(name||'').trim()};
  renderCalendar();
  document.querySelector('.calendar-card')?.scrollIntoView({behavior:'smooth',block:'start'});
}

function showMessage(text){
  message.textContent=text;
  message.classList.remove('hidden');
  clearTimeout(showMessage.timer);
  showMessage.timer=setTimeout(()=>message.classList.add('hidden'),6000);
}

function ensureBookingModal(){
  if(bookingModal)return bookingModal;
  const wrap=document.createElement('div');
  wrap.className='modal hidden';
  wrap.innerHTML='<div class="modal-content"><button class="modal-close booking-close" type="button" aria-label="Закрыть">×</button><p class="eyebrow">ЗАПИСЬ</p><h2 class="booking-title"></h2><div class="booking-body"></div><button class="modal-ok booking-submit" type="button">Записаться</button></div>';
  document.body.appendChild(wrap);
  wrap.querySelector('.booking-close').addEventListener('click',()=>wrap.classList.add('hidden'));
  wrap.addEventListener('click',e=>{if(e.target===wrap)wrap.classList.add('hidden');});
  bookingModal=wrap;
  return wrap;
}

function openBooking(slot){
  const modal=ensureBookingModal();
  const title=modal.querySelector('.booking-title');
  const body=modal.querySelector('.booking-body');
  const submit=modal.querySelector('.booking-submit');
  const start=slotMinutes(slot),end=slotEndMinutes(slot);
  const maxTickets=Math.max(0,Number(slot.free||0));
  const minTickets=Math.max(1,Number(slot.minTickets||1));
  const fixedDuration=Math.max(60,end-start);
  const durationOptions=[];
  for(let minutes=60;minutes<=fixedDuration;minutes+=60)durationOptions.push(minutes);
  title.textContent=slot.name||'Запись';
  const date=String(slot.start||'').split(' ')[0];
  const time=slotTime(slot);
  body.innerHTML=`<p style="margin:0 0 10px;color:var(--tg-hint);font-size:14px">${esc(date)} · ${esc(time)}${fixedDuration>60?`–${esc(timeKey(end))}`:''}</p><label style="display:block;margin-bottom:10px;font-size:13px">Количество мест<select class="booking-tickets" style="display:block;width:100%;margin-top:5px;padding:9px;border:1px solid var(--tg-separator);border-radius:8px;background:var(--tg-secondary-bg);color:var(--tg-text)">${Array.from({length:Math.max(0,maxTickets-minTickets+1)},(_,i)=>{const n=minTickets+i;return`<option value="${n}">${n}</option>`}).join('')}</select></label>${durationOptions.length>1?`<label style="display:block;margin-bottom:12px;font-size:13px">Время окончания<select class="booking-end" style="display:block;width:100%;margin-top:5px;padding:9px;border:1px solid var(--tg-separator);border-radius:8px;background:var(--tg-secondary-bg);color:var(--tg-text)">${durationOptions.map(minutes=>`<option value="${timeKey(start+minutes)}">${timeKey(start+minutes)}</option>`).join('')}</select></label>`:''}<p class="booking-hint" style="margin:0;color:var(--tg-hint);font-size:12px">Свободно мест: ${maxTickets}</p>`;
  const valid=maxTickets>=minTickets&&slotAvailable(slot,minTickets);
  submit.disabled=!valid;
  submit.textContent=valid?'Записаться':'Слот недоступен';
  submit.onclick=async()=>{
    if(!slotAvailable(slot,1)){modal.classList.add('hidden');showMessage('Этот слот уже занят или закрыт.');return;}
    const tickets=Number(body.querySelector('.booking-tickets')?.value||minTickets);
    const endTime=body.querySelector('.booking-end')?.value||'';
    if(tickets<minTickets||tickets>Number(slot.free||0)){showMessage('Недостаточно свободных мест.');return;}
    submit.disabled=true;
    submit.textContent='Записываем...';
    try{
      await apiBook(slot,tickets,endTime);
      modal.classList.add('hidden');
      showMessage('Запись создана.');
      await loadBookingData();
    }catch(e){
      showMessage(e.message||'Не удалось создать запись.');
      submit.disabled=false;
      submit.textContent='Записаться';
    }
  };
  modal.classList.remove('hidden');
}

async function apiBook(slot,tickets,endTime=''){
  if(!user?.id)throw Error('Не удалось определить Telegram ID. Откройте запись внутри Telegram.');
  const payload={action:'book',telegramId:String(user.id),name:String(user.first_name||user.username||'гость'),format:slot.format,eventName:slot.name,date:slot.start.split(' ')[0],time:slot.start.split(' ')[1],tickets:Number(tickets),endTime};
  const r=await fetch(BOOKING_API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload)});
  const j=await r.json();
  if(!j.ok)throw Error(j.error||'Не удалось создать бронь.');
  return j;
}

function closeActivity(){activityModal.classList.add('hidden');}
modalClose?.addEventListener('click',closeActivity);
modalOk?.addEventListener('click',closeActivity);
eventsClose?.addEventListener('click',closeEvents);
activityModal?.addEventListener('click',e=>{if(e.target===activityModal)closeActivity();});
eventsModal?.addEventListener('click',e=>{if(e.target===eventsModal)closeEvents();});

document.addEventListener('keydown',e=>{
  if(e.key!=='Escape')return;
  activityModal?.classList.add('hidden');
  eventsModal?.classList.add('hidden');
  bookingModal?.classList.add('hidden');
});

(async()=>{
  try{
    await Promise.all([loadActivities(),loadEvents(),loadBookingData()]);
  }catch(e){console.error(e);}
})();
