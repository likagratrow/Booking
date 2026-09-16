/**
 * Strange Things Booking — Google Apps Script
 * Business data comes from "Ивенты"; "Брони" is the booking journal;
 * Google Calendar is the occupancy source.
 */
const EVENTS_SHEET_NAME='Ивенты';
const BOOKING_SHEET_NAME='Брони';
const BOOKING_CALENDAR_NAME='Странные Вещи Букинг';
const FREE_EVENT_TITLE='Свободно';
const SLOT_STEP_MINUTES=60;
const LOOKAHEAD_DAYS=30;
const BOOKING_MARKER='ST_BOOKING_SLOT:';

function doGet(){try{return jsonResponse(getBookingData());}catch(e){return jsonResponse({ok:false,error:e.message});}}
function doPost(e){try{if(!e?.postData?.contents)throw Error('Не получены данные POST-запроса.');const data=JSON.parse(e.postData.contents);if(data.action==='book')return jsonResponse(createBooking(data));if(data.action==='cancel')return jsonResponse(cancelBooking(data));throw Error('Неизвестное действие: '+data.action);}catch(e){return jsonResponse({ok:false,error:e.message});}}
function jsonResponse(data){return ContentService.createTextOutput(JSON.stringify(data,null,2)).setMimeType(ContentService.MimeType.JSON);}

function getBookingData(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),es=ss.getSheetByName(EVENTS_SHEET_NAME),bs=ss.getSheetByName(BOOKING_SHEET_NAME);
  if(!es)throw Error('Лист "'+EVENTS_SHEET_NAME+'" не найден.');if(!bs)throw Error('Лист "'+BOOKING_SHEET_NAME+'" не найден.');
  const events=readEvents(es),bookings=readBookings(bs),calendar=readCalendar();
  return {ok:true,settings:{slotStepMinutes:SLOT_STEP_MINUTES,lookaheadDays:LOOKAHEAD_DAYS},events:events,bookings:{count:bookings.length,items:bookings},calendar:{name:BOOKING_CALENDAR_NAME,events:calendar.events,freeWindows:calendar.freeWindows,blocks:calendar.blocks}};
}

function createBooking(data){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(),es=ss.getSheetByName(EVENTS_SHEET_NAME),bs=ss.getSheetByName(BOOKING_SHEET_NAME);
    if(!es||!bs)throw Error('Не найден лист бронирований или мероприятий.');
    const events=readEvents(es),bookings=readBookings(bs),calendar=readCalendar();
    const tickets=Number(data.tickets);if(!Number.isInteger(tickets)||tickets<1)throw Error('Количество билетов должно быть целым числом от 1.');
    const telegramId=String(data.telegramId??'').trim();if(!telegramId)throw Error('Не указан Telegram ID.');
    const telegramName=resolveTelegramName(data);if(!telegramName)throw Error('Не указано имя Telegram пользователя.');
    const eventName=String(data.eventName||'').trim(),activity=String(data.activity||'').trim();
    const date=String(data.date||'').trim(),time=String(data.time||'').trim();if(!eventName||!date||!time)throw Error('Не указаны мероприятие, дата или время.');
    const event=findEventByIdentity(events,activity,eventName);if(!event)throw Error('Выбранное мероприятие не найдено.');
    const start=parseBookingDateTime(date,time);if(start.getTime()<=Date.now())throw Error('Время начала этого слота уже прошло.');
    const capacity=getCapacity(event);if(capacity<1)throw Error('Для выбранного мероприятия не задана вместимость.');
    const end=resolveBookingEnd(data,event,start);const hours=getHourlyRanges(start,end);if(!hours.length)throw Error('Некорректный интервал бронирования.');
    let hasExisting=false;
    hours.forEach(function(h){
      if(intersectsAnyBlock(h.start,h.end,calendar.blocks))throw Error('В выбранном интервале есть занятое время.');
      const occ=getCalendarBookingOccupancy(calendar.events,h.start,h.end,event);
      if(occ.exists)hasExisting=true;
      const free=Math.max(0,capacity-occ.occupied);
      if(free<1)throw Error('В выбранном интервале есть час без свободных мест.');
      if(tickets>free)throw Error('В одном из выбранных часов свободно только '+free+' мест.');
      if(!occ.exists&&!hasFreeCalendarHours(h.start,1,calendar.freeWindows,calendar.blocks))throw Error('Один из выбранных часов больше недоступен.');
    });
    if(!hasExisting){const min=Number(event.min);if(!Number.isFinite(min)||min<1)throw Error('Для выбранного мероприятия не задано минимальное количество мест.');if(tickets<min)throw Error('Для нового мероприятия нужно забронировать минимум '+min+' мест.');}
    hours.forEach(function(h){getAllBookingsAtStart(bookings,h.start).forEach(function(b){if(!sameEvent(b,event))throw Error('Это время уже занято другим мероприятием.');});});
    const id=getNextBookingId(bs),slotId=makeBookingSlotId(event,start,end);
    bs.appendRow([id,telegramId,telegramName,String(event.activity).trim(),String(event.name).trim(),date,time,tickets,'active',slotId]);
    const row=bs.getLastRow(),updated=readBookings(bs);
    try{syncCalendarBooking(event,start,end,updated);}catch(e){try{bs.deleteRow(row);}catch(x){Logger.log(x.message);}throw e;}
    const finalBookings=readBookings(bs),slotBookings=getBookingsForExactSlot(finalBookings,start,event),occupied=calculateOccupancy(slotBookings);
    return {ok:true,booking:{id:id,telegramId:telegramId,telegramName:telegramName,activity:event.activity,name:event.name,date:date,time:time,endTime:formatTime(end),tickets:tickets,status:'active',slotId:slotId},slot:{slotId:slotId,start:formatDateTime(start),end:formatDateTime(end),capacity:capacity,occupied:occupied,free:Math.max(0,capacity-occupied)}};
  }finally{lock.releaseLock();}
}

function cancelBooking(data){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(),bs=ss.getSheetByName(BOOKING_SHEET_NAME);if(!bs)throw Error('Лист "'+BOOKING_SHEET_NAME+'" не найден.');
    const map=getBookingColumnMap(bs),id=String(data.bookingId??data.id??'').trim();if(!id)throw Error('Не указан ID брони.');
    const rows=bs.getLastRow()<2?[]:bs.getRange(2,1,bs.getLastRow()-1,10).getValues();let rowNumber=null,booking=null;
    rows.forEach(function(row,i){if(rowNumber!==null)return;if(String(row[map.id]).trim()===id){rowNumber=i+2;booking=bookingFromRow(row,map);}});
    if(!booking)throw Error('Бронь не найдена.');if(booking.status==='cancel')throw Error('Эта бронь уже отменена.');
    if(data.telegramId!==undefined&&String(data.telegramId).trim()!==String(booking.telegramId).trim())throw Error('Эта бронь принадлежит другому пользователю.');
    const event=findEventForBooking(booking);if(!event)throw Error('Не удалось определить мероприятие для отменённой брони.');
    const interval=getBookingInterval(booking,event);bs.getRange(rowNumber,map.status+1).setValue('cancel');
    const bookings=readBookings(bs);
    try{syncCalendarBooking(event,interval.start,interval.end,bookings);}catch(e){bs.getRange(rowNumber,map.status+1).setValue('active');throw e;}
    return {ok:true,booking:{id:booking.id,status:'cancel'}};
  }finally{lock.releaseLock();}
}

function readEvents(sheet){
  if(sheet.getLastRow()<2)return[];const rows=sheet.getRange(2,1,sheet.getLastRow()-1,11).getValues();
  return rows.map(function(r){return{activity:r[0]||'',name:r[1]||'',price:r[2]||'',age:r[3]||'',duration:parseDuration(r[4]),complexity:r[5]||'',image:r[6]||'',description:r[8]||'',amount:parseAmount(r[9]),min:parseAmount(r[10])};}).filter(function(e){return String(e.name).trim()!=='';});
}
function getBookingColumnMap(sheet){
  const h=sheet.getRange(1,1,1,Math.min(10,sheet.getLastColumn())).getValues()[0].map(function(x){return String(x||'').trim().toLowerCase();});
  const find=function(){for(let j=0;j<arguments.length;j++){const i=h.indexOf(arguments[j].toLowerCase());if(i>=0)return i;}return -1;};
  const modern={id:find('id'),telegramId:find('telegram id'),telegramName:find('telegram name'),activity:find('activity'),name:find('name'),date:find('date'),time:find('time'),ticket:find('ticket'),status:find('status'),slotId:find('slot_id')};
  if(Object.values(modern).every(function(i){return i>=0;}))return modern;
  return {id:0,telegramId:1,telegramName:2,activity:3,name:4,date:5,time:6,ticket:7,status:8,slotId:9,legacy:true};
}
function bookingFromRow(r,m){return{id:r[m.id],telegramId:r[m.telegramId],telegramName:r[m.telegramName],activity:String(r[m.activity]||'').trim(),name:String(r[m.name]||'').trim(),date:normalizeDateValue(r[m.date]),time:normalizeTimeValue(r[m.time]),tickets:Number(r[m.ticket])||0,status:normalizeStatus(r[m.status]),slotId:r[m.slotId]||null};}
function readBookings(sheet){
  if(sheet.getLastRow()<2)return[];const m=getBookingColumnMap(sheet),rows=sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues();
  return rows.map(function(r){return bookingFromRow(r,m);}).filter(function(b){return b.id!==''&&b.id!==null;});
}
function normalizeStatus(v){return String(v||'').trim()||'active';}
function resolveTelegramName(data){const username=String(data.telegramName||data.username||'').trim();if(username)return username.charAt(0)==='@'?username:'@'+username;return String(data.name||'').trim().replace(/^@/,'');}

function findEventByIdentity(events,activity,name){const a=String(activity||'').trim(),n=String(name||'').trim(),matches=events.filter(function(e){return String(e.name).trim()===n;});if(a){const exact=matches.filter(function(e){return String(e.activity).trim()===a;});if(exact.length===1)return exact[0];}if(matches.length===1)return matches[0];return null;}
function findEventForBooking(b){const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName(EVENTS_SHEET_NAME),events=readEvents(sheet),exact=findEventByIdentity(events,b.activity,b.name);if(exact)return exact;const parsed=parseBookingSlotId(b.slotId)||parseOldSlotId(b.slotId);if(parsed)return findEventByIdentity(events,parsed.activity,parsed.eventName);return null;}
function sameEvent(b,e){return String(b.activity).trim()===String(e.activity).trim()&&String(b.name).trim()===String(e.name).trim();}
function getBookingsForExactSlot(bookings,start,event){const end=addMinutes(start,SLOT_STEP_MINUTES);return bookings.filter(function(b){if(b.status!=='active'||!sameEvent(b,event))return false;const interval=getBookingInterval(b,event);return interval.start.getTime()<end.getTime()&&interval.end.getTime()>start.getTime();});}
function getAllBookingsAtStart(bookings,start){return bookings.filter(function(b){return b.status==='active'&&b.date===formatDate(start)&&b.time===formatTime(start);});}
function calculateOccupancy(bs){return bs.reduce(function(n,b){return n+(Number(b.tickets)||0);},0);}

function readCalendar(){const c=getBookingCalendar(),now=new Date(),from=startOfDay(now),to=new Date(from);to.setDate(to.getDate()+LOOKAHEAD_DAYS+1);const all=c.getEvents(from,to),events=[],freeWindows=[],blocks=[];all.forEach(function(e){const title=String(e.getTitle()||'').trim(),item={id:e.getId(),title:title,start:e.getStartTime().toISOString(),end:e.getEndTime().toISOString(),description:String(e.getDescription()||''),isBooking:isBookingCalendarEvent(e)};events.push(item);if(title===FREE_EVENT_TITLE)freeWindows.push({id:item.id,start:item.start,end:item.end});else if(!item.isBooking)blocks.push({id:item.id,title:title,start:item.start,end:item.end});});return{events:events,freeWindows:freeWindows,blocks:blocks};}
function getCalendarBookingOccupancy(items,start,end,event){const matches=(items||[]).filter(function(i){if(!i||!i.isBooking)return false;const s=new Date(i.start),e=new Date(i.end);if(s>=end||e<=start)return false;const parsed=parseCalendarSlotId(i.description),title=String(i.title||'').trim(),expected=String(event.activity).trim()+' — '+String(event.name).trim();return(Boolean(parsed&&calendarSlotMatchesEvent(parsed,event))||title===expected);});if(!matches.length)return{exists:false,occupied:0,capacity:getCapacity(event)};let occupied=0,capacity=getCapacity(event);matches.forEach(function(i){const m=String(i.description).match(/(\d+)\s*\/\s*(\d+)\s*$/m);if(m){occupied=Math.max(occupied,Number(m[1]));capacity=Number(m[2])||capacity;}});return{exists:true,occupied:occupied,capacity:capacity};}
function parseCalendarSlotId(description){const marker=String(description||'').split('\n')[0];if(marker.indexOf(BOOKING_MARKER)!==0)return null;const id=marker.slice(BOOKING_MARKER.length),m=id.match(/^(.*)_(\d{8}-\d{4})$/);if(!m)return null;return{slotId:id,prefix:m[1],start:m[2]};}
function calendarSlotMatchesEvent(parsed,event){const prefix=String(event.activity).trim()+'_'+String(event.name).trim()+'_';return String(parsed.prefix).indexOf(prefix)===0;}
function getBookingCalendar(){const c=CalendarApp.getCalendarsByName(BOOKING_CALENDAR_NAME);if(!c||!c.length)throw Error('Календарь "'+BOOKING_CALENDAR_NAME+'" не найден.');return c[0];}
function hasFreeCalendarHours(start,hours,freeWindows,blocks){for(let i=0;i<Number(hours);i++){const s=addMinutes(start,i*SLOT_STEP_MINUTES),e=addMinutes(s,SLOT_STEP_MINUTES),free=freeWindows.some(function(w){return new Date(w.start)<=s&&new Date(w.end)>=e;});if(!free||intersectsAnyBlock(s,e,blocks))return false;}return true;}
function intersectsAnyBlock(start,end,blocks){return blocks.some(function(b){return start.getTime()<new Date(b.end).getTime()&&end.getTime()>new Date(b.start).getTime();});}
function isBookingCalendarEvent(e){return String(e.getDescription()||'').indexOf(BOOKING_MARKER)===0;}

function syncCalendarBooking(event,start,end,bookings){const calendar=getBookingCalendar(),existing=calendar.getEvents(start,end).filter(function(e){if(!isBookingCalendarEvent(e))return false;const p=parseCalendarSlotId(e.getDescription()),title=String(e.getTitle()||'').trim(),expected=String(event.activity).trim()+' — '+String(event.name).trim();return(Boolean(p&&calendarSlotMatchesEvent(p,event))||title===expected);});let ce=existing[0]||null,baseStart=start,baseEnd=end;if(ce){const p=parseCalendarSlotId(ce.getDescription()),parsedStart=p&&p.start?parseCompactDateTime(p.start):null;if(parsedStart){baseStart=parsedStart;baseEnd=ce.getEndTime();}}const active=getBookingsForExactSlot(bookings,baseStart,event),occupancy=calculateOccupancy(active);if(occupancy<=0){existing.forEach(function(e){e.deleteEvent();});restoreFreeHours(baseStart,baseEnd);synchronizeCalendar();return;}const title=String(event.activity).trim()+' — '+String(event.name).trim(),description=buildCalendarDescription(event,baseStart,bookings);existing.slice(1).forEach(function(e){e.deleteEvent();});if(!ce)ce=calendar.createEvent(title,baseStart,baseEnd,{description:description});else{ce.setTitle(title);ce.setDescription(description);}synchronizeCalendar();}
function buildCalendarDescription(event,start,bookings){const slotId=makeSlotId(event,start),active=getBookingsForExactSlot(bookings,start,event),occ=calculateOccupancy(active),cap=getCapacity(event);return [BOOKING_MARKER+slotId].concat(active.map(function(b){return String(b.telegramName||'').trim()+' ('+(Number(b.tickets)||0)+')';})).concat([occ+'/'+cap]).join('\n');}
function restoreFreeHours(start,end){const c=getBookingCalendar();for(let s=new Date(start);s<end;s=addMinutes(s,60)){const e=addMinutes(s,60),events=c.getEvents(s,e),hasBlock=events.some(function(x){const t=String(x.getTitle()||'').trim();return t!==FREE_EVENT_TITLE&&!isBookingCalendarEvent(x);}),hasFree=events.some(function(x){return String(x.getTitle()||'').trim()===FREE_EVENT_TITLE&&x.getStartTime().getTime()===s.getTime()&&x.getEndTime().getTime()===e.getTime();});if(!hasBlock&&!hasFree)c.createEvent(FREE_EVENT_TITLE,s,e);}}
function synchronizeCalendar(){const c=getBookingCalendar(),now=new Date(),from=startOfDay(now),to=new Date(from);to.setDate(to.getDate()+LOOKAHEAD_DAYS+1);const all=c.getEvents(from,to),free=all.filter(function(e){return String(e.getTitle()||'').trim()===FREE_EVENT_TITLE;}),bookings=all.filter(isBookingCalendarEvent),blocks=all.filter(function(e){const t=String(e.getTitle()||'').trim();return t!==FREE_EVENT_TITLE&&!isBookingCalendarEvent(e);});if(!free.length)return;const desired={};free.forEach(function(f){for(let s=new Date(f.getStartTime());s<f.getEndTime();s=addMinutes(s,60)){const e=addMinutes(s,60);if(e>f.getEndTime())break;if(!intersectsCalendarEvents(s,e,blocks)&&!intersectsCalendarEvents(s,e,bookings))desired[s.getTime()+'_'+e.getTime()]={s:s,e:e};}});free.forEach(function(f){const k=f.getStartTime().getTime()+'_'+f.getEndTime().getTime();if(!desired[k])f.deleteEvent();});Object.keys(desired).forEach(function(k){const d=desired[k];if(free.some(function(f){return f.getStartTime().getTime()===d.s.getTime()&&f.getEndTime().getTime()===d.e.getTime();}))return;c.createEvent(FREE_EVENT_TITLE,d.s,d.e);});}
function intersectsCalendarEvents(s,e,events){return events.some(function(x){return s.getTime()<x.getEndTime().getTime()&&e.getTime()>x.getStartTime().getTime();});}

function makeSlotId(event,start){return[String(event.activity).trim(),String(event.name).trim(),Utilities.formatDate(start,Session.getScriptTimeZone(),'yyyyMMdd-HHmm')].join('_');}
function makeBookingSlotId(event,start,end){return[String(event.activity).trim(),String(event.name).trim(),Utilities.formatDate(start,Session.getScriptTimeZone(),'yyyyMMdd-HHmm'),Utilities.formatDate(end,Session.getScriptTimeZone(),'yyyyMMdd-HHmm')].join('::');}
function parseBookingSlotId(id){const p=String(id||'').split('::');if(p.length!==4)return null;return{activity:p[0],eventName:p[1],start:p[2],end:p[3]};}
function parseOldSlotId(id){const m=String(id||'').match(/^([^_]+)_(.+)_(\d{8}-\d{4})$/);return m?{activity:m[1],eventName:m[2],start:m[3]}:null;}
function parseSlotIdentity(id){const p=String(id||'').split('_');if(p.length<3)return null;return{activity:p[0],eventName:p.slice(1,-1).join('_'),start:p[p.length-1],slotId:id};}
function getBookingInterval(b,event){const p=parseBookingSlotId(b.slotId);if(p&&p.start&&p.end){const s=parseCompactDateTime(p.start),e=parseCompactDateTime(p.end);if(s&&e)return{start:s,end:e};}const s=parseBookingDateTime(b.date,b.time);return{start:s,end:addMinutes(s,Number(event.duration)*60)};}
function resolveBookingEnd(data,event,start){const supplied=String(data.endTime||'').trim();if(supplied)return parseBookingDateTime(formatDate(start),supplied);const duration=Number(event.duration);if(!(duration>0))throw Error('Для выбранного мероприятия не задана длительность.');return addMinutes(start,duration*60);}
function getHourlyRanges(s,e){const a=[];for(let x=new Date(s);x<e;x=addMinutes(x,60))a.push({start:new Date(x),end:addMinutes(x,60)});return a;}
function parseCompactDateTime(v){const m=String(v||'').match(/^(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})$/);return m?parseBookingDateTime(m[3]+'.'+m[2]+'.'+m[1],m[4]+':'+m[5]):null;}
function parseBookingDateTime(d,t){const r=Utilities.parseDate(String(d)+' '+String(t),Session.getScriptTimeZone(),'dd.MM.yyyy HH:mm');if(!r||isNaN(r.getTime()))throw Error('Некорректная дата или время: '+d+' '+t);return r;}
function parseDuration(v){if(v===null||v===undefined||v==='')return 0;if(typeof v==='number')return v;const s=String(v).trim().toLowerCase(),m=s.match(/(\d+(?:[.,]\d+)?)/);if(!m)return 0;const n=Number(m[1].replace(',','.'));return s.indexOf('мин')>=0?n/60:n;}
function parseAmount(v){const m=String(v??'').match(/\d+/);return m?Number(m[0]):0;}
function getCapacity(e){return Math.max(0,Number(e.amount)||0);}
function getNextBookingId(sheet){if(sheet.getLastRow()<2)return 1;return Math.max(0,...sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().map(function(r){return Number(r[0])||0;}))+1;}
function normalizeDateValue(v){return v instanceof Date&&!isNaN(v)?formatDate(v):String(v??'').trim();}
function normalizeTimeValue(v){return v instanceof Date&&!isNaN(v)?formatTime(v):String(v??'').trim();}
function formatDate(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd.MM.yyyy');}
function formatTime(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'HH:mm');}
function formatDateTime(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd.MM.yyyy HH:mm');}
function startOfDay(d){const x=new Date(d);x.setHours(0,0,0,0);return x;}
function addMinutes(d,n){return new Date(d.getTime()+n*60000);}
