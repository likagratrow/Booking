/**
 * Strange Things Booking
 * Google Apps Script
 *
 * Источники:
 * 1. Лист "Ивенты" — мероприятия и их параметры.
 * 2. Лист "Брони" — существующие бронирования.
 * 3. Календарь "Странные Вещи Букинг":
 *      - "Свободно" = рабочее время;
 *      - другие обычные события = блокировки;
 *      - события с техническим маркером = созданные бронирования.
 *
 * Ивенты:
 * A activity
 * B name
 * C price
 * D age
 * E duration (часы)
 * F complexity
 * G image
 * H не используется
 * I description
 * J amount (вместимость)
 * K min (минимум для новой брони)
 *
 * Брони:
 * A ID
 * B Telegram ID
 * C Имя
 * D Формат
 * E МК
 * F Дата
 * G Время
 * H Билетов
 * I Статус
 * J slot_id
 */

const EVENTS_SHEET_NAME = 'Ивенты';
const BOOKING_SHEET_NAME = 'Брони';
const BOOKING_CALENDAR_NAME = 'Странные Вещи Букинг';
const FREE_EVENT_TITLE = 'Свободно';
const SLOT_STEP_MINUTES = 60;
const LOOKAHEAD_DAYS = 30;
const BOOKING_MARKER = 'ST_BOOKING_SLOT:';

function doGet(e) {
  try { return jsonResponse(getBookingData()); }
  catch (error) { return jsonResponse({ok:false,error:error.message}); }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Не получены данные POST-запроса.');
    const data = JSON.parse(e.postData.contents);
    if (!data.action) throw new Error('Не указано действие.');
    let result;
    if (data.action === 'book') result = createBooking(data);
    else if (data.action === 'cancel') result = cancelBooking(data);
    else throw new Error('Неизвестное действие: ' + data.action);
    return jsonResponse(result);
  } catch (error) { return jsonResponse({ok:false,error:error.message}); }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data,null,2)).setMimeType(ContentService.MimeType.JSON);
}

function getBookingData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Не удалось получить текущую таблицу.');
  const eventsSheet = spreadsheet.getSheetByName(EVENTS_SHEET_NAME);
  if (!eventsSheet) throw new Error('Лист "'+EVENTS_SHEET_NAME+'" не найден.');
  const bookingsSheet = spreadsheet.getSheetByName(BOOKING_SHEET_NAME);
  if (!bookingsSheet) throw new Error('Лист "'+BOOKING_SHEET_NAME+'" не найден.');
  const events = readEvents(eventsSheet);
  const bookings = readBookings(bookingsSheet);
  const calendar = readCalendar();
  return {
    ok:true,
    settings:{slotStepMinutes:SLOT_STEP_MINUTES,lookaheadDays:LOOKAHEAD_DAYS},
    events:events,
    bookings:{count:bookings.length,items:bookings},
    calendar:{name:BOOKING_CALENDAR_NAME,events:calendar.events,freeWindows:calendar.freeWindows,blocks:calendar.blocks}
  };
}

function createBooking(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const eventsSheet = spreadsheet.getSheetByName(EVENTS_SHEET_NAME);
    const bookingsSheet = spreadsheet.getSheetByName(BOOKING_SHEET_NAME);
    if (!eventsSheet) throw new Error('Лист "'+EVENTS_SHEET_NAME+'" не найден.');
    if (!bookingsSheet) throw new Error('Лист "'+BOOKING_SHEET_NAME+'" не найден.');

    const events = readEvents(eventsSheet);
    const bookings = readBookings(bookingsSheet);
    const calendar = readCalendar();
    const tickets = Number(data.tickets);
    if (!Number.isInteger(tickets) || tickets < 1) throw new Error('Количество билетов должно быть целым числом от 1.');

    const telegramId = data.telegramId === undefined || data.telegramId === null ? '' : String(data.telegramId).trim();
    const userName = data.name === undefined || data.name === null ? '' : String(data.name).trim();
    if (!telegramId) throw new Error('Не указан Telegram ID.');
    if (!userName) throw new Error('Не указано имя.');

    const eventName = String(data.eventName || '').trim();
    const date = String(data.date || '').trim();
    const time = String(data.time || '').trim();
    const activityHint = String(data.activity || '').trim();
    if (!eventName) throw new Error('Не указано мероприятие.');
    if (!date || !time) throw new Error('Не указаны дата и время.');

    const event = findEventByName(events,eventName,activityHint);
    if (!event) throw new Error('Выбранное мероприятие не найдено.');
    const format = getFormat(event);
    const slotStart = parseBookingDateTime(date,time);
    if (slotStart.getTime() <= new Date().getTime()) throw new Error('Время начала этого слота уже прошло.');

    const capacity = getCapacity(event);
    if (capacity <= 0) throw new Error('Для выбранного мероприятия не задана вместимость.');

    let slotEnd;
    if (getFormat(event) === 'Диоген') {
      const endTime = String(data.endTime || '').trim();
      slotEnd = endTime ? parseBookingDateTime(date,endTime) : addMinutes(slotStart,Number(event.duration)*60);
      if (slotEnd.getTime() <= slotStart.getTime()) throw new Error('Время окончания должно быть позже времени начала.');
      if (slotStart.getMinutes() !== 0 || slotEnd.getMinutes() !== 0 || slotStart.getSeconds() !== 0 || slotEnd.getSeconds() !== 0) throw new Error('Диоген можно бронировать только целыми часами.');
    } else {
      if (!(Number(event.duration) > 0)) throw new Error('Для выбранного мероприятия не задана длительность.');
      slotEnd = addMinutes(slotStart,Number(event.duration)*60);
    }

    const affectedHours = getHourlyRanges(slotStart,slotEnd);
    if (affectedHours.length === 0) throw new Error('Некорректный интервал бронирования.');
    let hasExistingOccupancy = false;

    affectedHours.forEach(function(hour) {
      if (intersectsAnyBlock(hour.start,hour.end,calendar.blocks)) throw new Error('В выбранном интервале есть занятое время.');
      const calendarOccupancy = getCalendarBookingOccupancy(calendar.events,hour.start,hour.end,event);
      const occupancy = calendarOccupancy.exists ? calendarOccupancy.occupied : 0;
      if (calendarOccupancy.exists) hasExistingOccupancy = true;
      const free = Math.max(0,capacity-occupancy);
      if (free <= 0) throw new Error('В выбранном интервале есть час без свободных мест.');
      if (tickets > free) throw new Error('В одном из выбранных часов свободно только '+free+' мест.');
      if (!calendarOccupancy.exists && !hasFreeCalendarHours(hour.start,1,calendar.freeWindows,calendar.blocks)) throw new Error('Один из выбранных часов больше недоступен.');
    });

    const configuredMin = event.min;
    if (!hasExistingOccupancy) {
      if (!Number.isFinite(Number(configuredMin)) || Number(configuredMin) < 1) throw new Error('Для выбранного мероприятия не задано минимальное количество мест.');
      const minTickets = Number(configuredMin);
      if (tickets < minTickets) throw new Error('Для нового мероприятия нужно забронировать минимум '+minTickets+' мест.');
    }

    affectedHours.forEach(function(hour) {
      const conflictingBookings = getAllBookingsAtStart(bookings,hour.start);
      conflictingBookings.forEach(function(booking) {
        if (normalizeFormat(booking.format) !== format) throw new Error('Это время уже занято другим форматом.');
        if (format === 'МК' && String(booking.activity || '').trim() !== event.name) throw new Error('Это время уже занято другим мастер-классом.');
        if (format === 'Диоген') {
          const bookingEvent = findEventForBooking(booking);
          if (bookingEvent && String(bookingEvent.name).trim() !== String(event.name).trim()) throw new Error('Это время уже занято другим форматом Диогена.');
        }
      });
    });

    const bookingId = getNextBookingId(bookingsSheet);
    const slotId = makeBookingSlotId(event,slotStart,slotEnd);
    const activityForSheet = getFormat(event) === 'МК' ? event.name : '—';
    bookingsSheet.appendRow([bookingId,telegramId,userName,getFormat(event),activityForSheet,date,time,tickets,'active',slotId]);
    const bookingRow = bookingsSheet.getLastRow();
    const updatedBookingsForCalendar = readBookings(bookingsSheet);

    try {
      if (getFormat(event) === 'Диоген') syncDiogenCalendarRange(event,slotStart,slotEnd,updatedBookingsForCalendar);
      else if (!hasExistingOccupancy) createCalendarBooking(event,slotStart,updatedBookingsForCalendar);
      else updateCalendarBooking(event,slotStart,updatedBookingsForCalendar);
    } catch (calendarError) {
      try { bookingsSheet.deleteRow(bookingRow); }
      catch (rollbackError) { Logger.log('Не удалось откатить строку брони: '+rollbackError.message); }
      throw calendarError;
    }

    const updatedBookings = readBookings(bookingsSheet);
    const updatedSlotBookings = getBookingsForExactSlot(updatedBookings,slotStart,event);
    const updatedOccupancy = calculateOccupancy(updatedSlotBookings);
    return {
      ok:true,
      booking:{id:bookingId,telegramId:telegramId,name:userName,format:getFormat(event),eventName:event.name,date:date,time:time,endTime:formatTime(slotEnd),tickets:tickets,status:'active',slotId:slotId},
      slot:{slotId:slotId,start:formatDateTime(slotStart),end:formatDateTime(slotEnd),capacity:capacity,occupied:updatedOccupancy,free:Math.max(0,capacity-updatedOccupancy)}
    };
  } finally { lock.releaseLock(); }
}

function cancelBooking(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const bookingsSheet = spreadsheet.getSheetByName(BOOKING_SHEET_NAME);
    if (!bookingsSheet) throw new Error('Лист "'+BOOKING_SHEET_NAME+'" не найден.');
    const bookingId = data.bookingId !== undefined ? data.bookingId : data.id;
    if (bookingId === undefined || bookingId === null || bookingId === '') throw new Error('Не указан ID брони.');
    const lastRow = bookingsSheet.getLastRow();
    const values = lastRow < 2 ? [] : bookingsSheet.getRange(2,1,lastRow-1,10).getValues();
    let rowNumber=null,booking=null;
    for (let i=0;i<values.length;i++) {
      const row=values[i];
      if (String(row[0]).trim()===String(bookingId).trim()) {
        rowNumber=i+2;
        booking={id:row[0],telegramId:row[1],name:row[2],format:row[3],activity:row[4],date:normalizeDateValue(row[5]),time:normalizeTimeValue(row[6]),tickets:Number(row[7])||0,status:String(row[8]||'').trim().toLowerCase(),slotId:row[9]||null};
        break;
      }
    }
    if (!booking) throw new Error('Бронь не найдена.');
    if (booking.status==='cancelled') throw new Error('Эта бронь уже отменена.');
    if (data.telegramId!==undefined && String(data.telegramId).trim()!==String(booking.telegramId).trim()) throw new Error('Эта бронь принадлежит другому пользователю.');
    const slotStart=parseBookingDateTime(booking.date,booking.time);
    const event=findEventForBooking(booking);
    if (!event) throw new Error('Не удалось определить мероприятие для отменённой брони.');
    const bookingInterval=getBookingInterval(booking,event);
    bookingsSheet.getRange(rowNumber,9).setValue('cancelled');
    const bookings=readBookings(bookingsSheet);
    try {
      if (getFormat(event)==='Диоген') syncDiogenCalendarRange(event,bookingInterval.start,bookingInterval.end,bookings);
      else {
        const remaining=getBookingsForExactSlot(bookings,slotStart,event);
        if (remaining.length===0) restoreCalendarFreeSlot(event,slotStart);
        else updateCalendarBooking(event,slotStart,bookings);
      }
    } catch (calendarError) {
      try { bookingsSheet.getRange(rowNumber,9).setValue('active'); }
      catch (rollbackError) { Logger.log('Не удалось откатить статус отменённой брони: '+rollbackError.message); }
      throw calendarError;
    }
    const updatedCalendar=readCalendar();
    return {ok:true,booking:{id:booking.id,status:'cancelled'},calendar:updatedCalendar};
  } finally { lock.releaseLock(); }
}

function findEvent(events,format,eventName) {
  return events.find(function(event){ return getFormat(event)===format && String(event.name).trim()===String(eventName).trim(); }) || null;
}

function findEventByName(events,eventName,activityHint) {
  const name=String(eventName||'').trim();
  const matches=events.filter(function(event){return String(event.name||'').trim()===name;});
  if (matches.length===0) return null;
  if (matches.length===1) return matches[0];
  const hintedActivity=String(activityHint||'').trim().toLowerCase();
  if (hintedActivity) {
    const hinted=matches.filter(function(event){return String(event.activity||'').trim().toLowerCase()===hintedActivity;});
    if (hinted.length===1) return hinted[0];
  }
  throw new Error('Найдено несколько мероприятий с названием "'+name+'". Не удалось однозначно определить активность.');
}

function findEventForBooking(booking) {
  const spreadsheet=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=spreadsheet.getSheetByName(EVENTS_SHEET_NAME);
  const events=readEvents(sheet);
  const format=normalizeFormat(booking.format);
  if (format==='МК') return findEvent(events,'МК',booking.activity);
  if (booking.slotId && String(booking.slotId).indexOf('::')!==-1) {
    const parsed=parseBookingSlotId(booking.slotId);
    if (parsed && parsed.eventName) { const found=findEvent(events,'Диоген',parsed.eventName); if(found)return found; }
  }
  if (booking.slotId) {
    const parsedOld=parseOldSlotId(booking.slotId);
    if (parsedOld && parsedOld.eventName) { const found=findEvent(events,'Диоген',parsedOld.eventName); if(found)return found; }
  }
  const candidates=events.filter(function(event){return getFormat(event)==='Диоген';});
  return candidates.length ? candidates[0] : null;
}

function synchronizeCalendar() {
  const calendar=getBookingCalendar();
  const now=new Date();
  const from=startOfDay(now);
  const to=new Date(from); to.setDate(to.getDate()+LOOKAHEAD_DAYS+1);
  const events=calendar.getEvents(from,to);
  const freeEvents=events.filter(function(event){return String(event.getTitle()||'').trim()===FREE_EVENT_TITLE;});
  const bookingEvents=events.filter(function(event){return isBookingCalendarEvent(event);});
  const blocks=events.filter(function(event){const title=String(event.getTitle()||'').trim();return title!==FREE_EVENT_TITLE&&!isBookingCalendarEvent(event);});
  if (freeEvents.length===0) return;
  const mergedWindows=mergeFreeWindows(freeEvents);
  const desiredFreeSlots={};
  mergedWindows.forEach(function(window){
    const roundedStart=floorToHour(window.start),roundedEnd=ceilToHour(window.end);
    let cursor=new Date(roundedStart);
    while(cursor.getTime()<roundedEnd.getTime()){
      const chunkStart=new Date(cursor),chunkEnd=addMinutes(chunkStart,SLOT_STEP_MINUTES);
      const conflictsWithBlock=intersectsCalendarEvents(chunkStart,chunkEnd,blocks);
      const conflictsWithBooking=intersectsCalendarEvents(chunkStart,chunkEnd,bookingEvents);
      if(!conflictsWithBlock&&!conflictsWithBooking){const key=chunkStart.getTime()+'_'+chunkEnd.getTime();desiredFreeSlots[key]={start:chunkStart,end:chunkEnd};}
      cursor=addMinutes(cursor,SLOT_STEP_MINUTES);
    }
  });
  const existingDesiredSlots={};
  freeEvents.forEach(function(event){
    const start=event.getStartTime(),end=event.getEndTime(),key=start.getTime()+'_'+end.getTime();
    if(desiredFreeSlots[key]) { if(!existingDesiredSlots[key]) existingDesiredSlots[key]=event; else event.deleteEvent(); }
    else event.deleteEvent();
  });
  Object.keys(desiredFreeSlots).forEach(function(key){if(existingDesiredSlots[key])return;const slot=desiredFreeSlots[key];calendar.createEvent(FREE_EVENT_TITLE,slot.start,slot.end);});
}

function maintainCalendarFreeWindows() {
  const calendar=getBookingCalendar();
  const now=new Date(),from=startOfDay(now),to=new Date(from);to.setDate(to.getDate()+LOOKAHEAD_DAYS+1);
  const events=calendar.getEvents(from,to);
  const needsSynchronization=events.some(function(event){const title=String(event.getTitle()||'').trim();if(title!==FREE_EVENT_TITLE)return false;const durationMinutes=(event.getEndTime().getTime()-event.getStartTime().getTime())/60000;return durationMinutes>SLOT_STEP_MINUTES;});
  if(!needsSynchronization)return;
  synchronizeCalendar();
}

function setupCalendarMaintenanceTrigger() {
  const handler='maintainCalendarFreeWindows';
  const triggers=ScriptApp.getProjectTriggers();
  const alreadyExists=triggers.some(function(trigger){return trigger.getHandlerFunction()===handler;});
  if(alreadyExists)return;
  ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(3).create();
}

function readCalendar() {
  const calendar=getBookingCalendar();
  const now=new Date(),from=startOfDay(now),to=new Date(from);to.setDate(to.getDate()+LOOKAHEAD_DAYS+1);
  const calendarEvents=calendar.getEvents(from,to);
  const events=[],freeWindows=[],blocks=[];
  calendarEvents.forEach(function(event){
    const title=String(event.getTitle()||'').trim();
    const item={id:event.getId(),title:title,start:event.getStartTime().toISOString(),end:event.getEndTime().toISOString(),description:String(event.getDescription()||''),isBooking:isBookingCalendarEvent(event)};
    events.push(item);
    if(title===FREE_EVENT_TITLE) freeWindows.push({id:item.id,start:item.start,end:item.end});
    else if(!item.isBooking) blocks.push({id:item.id,title:title||'Занято',start:item.start,end:item.end});
  });
  return {events:events,freeWindows:freeWindows,blocks:blocks};
}

function getCalendarBookingOccupancy(calendarEvents,start,end,event) {
  const capacity=getCapacity(event);
  const targetName=String(event.name||'').trim();
  const expectedTitle=getFormat(event)==='МК'?'МК — '+targetName:'Диоген — '+targetName;
  const matches=(calendarEvents||[]).filter(function(item){
    if(!item||!item.isBooking)return false;
    const itemStart=new Date(item.start),itemEnd=new Date(item.end);
    if(isNaN(itemStart.getTime())||isNaN(itemEnd.getTime()))return false;
    if(itemStart.getTime()>=end.getTime()||itemEnd.getTime()<=start.getTime())return false;
    if(String(item.title||'').trim()!==expectedTitle)return false;
    return true;
  });
  if(!matches.length)return{exists:false,occupied:0,capacity:capacity};
  const occupancyValues=matches.map(function(item){
    const text=String(item.description||'');
    const found=[...text.matchAll(/(\d+)\s*\/\s*(\d+)\s*$/gm)];
    if(!found.length)return null;
    const last=found[found.length-1];
    return{occupied:Number(last[1]),capacity:Number(last[2])};
  }).filter(Boolean);
  if(!occupancyValues.length)return{exists:true,occupied:0,capacity:capacity};
  const occupancy=occupancyValues[0];
  return{exists:true,occupied:Math.max(0,occupancy.occupied),capacity:occupancy.capacity>0?occupancy.capacity:capacity};
}

function getBookingCalendar() {
  const calendars=CalendarApp.getCalendarsByName(BOOKING_CALENDAR_NAME);
  if(!calendars||calendars.length===0) throw new Error('Календарь "'+BOOKING_CALENDAR_NAME+'" не найден.');
  return calendars[0];
}

function hasFreeCalendarHours(start,durationHours,freeWindows,blocks) {
  const hours=Number(durationHours);
  if(!hours||hours<=0)return false;
  for(let i=0;i<hours;i++){
    const hourStart=addMinutes(start,i*SLOT_STEP_MINUTES),hourEnd=addMinutes(hourStart,SLOT_STEP_MINUTES);
    const hasFree=freeWindows.some(function(window){const freeStart=new Date(window.start),freeEnd=new Date(window.end);return freeStart.getTime()<=hourStart.getTime()&&freeEnd.getTime()>=hourEnd.getTime();});
    if(!hasFree)return false;
    if(intersectsAnyBlock(hourStart,hourEnd,blocks))return false;
  }
  return true;
}

function readEvents(sheet) {
  const lastRow=sheet.getLastRow();
  if(lastRow<2)return [];
  const values=sheet.getRange(2,1,lastRow-1,11).getValues();
  return values.map(function(row){return{
    activity:row[0]||'',name:row[1]||'',price:row[2]||'',age:row[3]||'',duration:parseDuration(row[4]),complexity:row[5]||'',image:row[6]||'',description:row[8]||'',amount:parseAmount(row[9]),min:parseAmount(row[10])
  };}).filter(function(event){return event.name!=='';});
}

function readBookings(sheet) {
  const lastRow=sheet.getLastRow();
  if(lastRow<2)return [];
  const values=sheet.getRange(2,1,lastRow-1,10).getValues();
  return values.map(function(row){return{id:row[0],telegramId:row[1],name:row[2],format:row[3],activity:row[4],date:normalizeDateValue(row[5]),time:normalizeTimeValue(row[6]),tickets:Number(row[7])||0,status:String(row[8]||'').trim().toLowerCase(),slotId:row[9]||null};}).filter(function(booking){return booking.id!==''&&booking.id!==null;});
}

function getBookingsForExactSlot(bookings,slotStart,event) {
  const targetDate=formatDate(slotStart),targetTime=formatTime(slotStart),targetFormat=getFormat(event);
  return bookings.filter(function(booking){
    if(booking.status==='cancelled')return false;
    if(normalizeFormat(booking.format)!==targetFormat)return false;
    if(targetFormat==='МК') return booking.date===targetDate&&booking.time===targetTime&&String(booking.activity||'').trim()===String(event.name).trim();
    const bookingEvent=findEventForBooking(booking);if(!bookingEvent)return false;
    if(String(bookingEvent.name).trim()!==String(event.name).trim())return false;
    const interval=getBookingInterval(booking,bookingEvent),hourEnd=addMinutes(slotStart,SLOT_STEP_MINUTES);
    return interval.start.getTime()<hourEnd.getTime()&&interval.end.getTime()>slotStart.getTime();
  });
}

function getAllBookingsAtStart(bookings,slotStart) {
  const targetDate=formatDate(slotStart),targetTime=formatTime(slotStart);
  return bookings.filter(function(booking){return booking.status!=='cancelled'&&booking.date===targetDate&&booking.time===targetTime;});
}

function calculateOccupancy(bookings){return bookings.reduce(function(total,booking){return total+(Number(booking.tickets)||0);},0);}

function getFormat(event) {
  if(String(event.activity).trim().toLowerCase()==='диоген')return 'Диоген';
  return 'МК';
}

function normalizeFormat(value) {
  const text=String(value||'').trim().toLowerCase();
  if(text==='диоген')return 'Диоген';
  if(text==='мк'||text==='masterclass')return 'МК';
  return String(value||'').trim();
}

function getCapacity(event){return Math.max(0,Number(event.amount)||0);}
function calculateMinTickets(event,occupancy,capacity){if(occupancy>0)return 1;return Number(event.min)||0;}

function createCalendarBooking(event,slotStart,bookings) {
  const calendar=getBookingCalendar(),durationMinutes=Number(event.duration)*60,slotEnd=addMinutes(slotStart,durationMinutes);
  const title=getFormat(event)==='МК'?'МК — '+event.name:'Диоген — '+event.name;
  const description=buildCalendarBookingDescription(event,slotStart,bookings);
  let bookingEvent=null;
  try { bookingEvent=calendar.createEvent(title,slotStart,slotEnd,{description:description});synchronizeCalendar();return bookingEvent; }
  catch(error){
    if(bookingEvent){try{bookingEvent.deleteEvent();}catch(cleanupError){Logger.log('Не удалось удалить календарную бронь после ошибки: '+cleanupError.message);}}
    try{synchronizeCalendar();}catch(syncError){Logger.log('Не удалось повторно синхронизировать календарь: '+syncError.message);}
    throw error;
  }
}

function updateCalendarBooking(event,slotStart,bookings) {
  const calendar=getBookingCalendar(),slotEnd=addMinutes(slotStart,Number(event.duration)*60),slotId=makeSlotId(event,slotStart),events=calendar.getEvents(slotStart,slotEnd);
  let bookingEvent=null;
  events.forEach(function(calendarEvent){if(bookingEvent)return;if(!isBookingCalendarEvent(calendarEvent))return;const description=String(calendarEvent.getDescription()||'');if(description.indexOf(BOOKING_MARKER+slotId)===0)bookingEvent=calendarEvent;});
  if(!bookingEvent)return createCalendarBooking(event,slotStart,bookings);
  const title=getFormat(event)==='МК'?'МК — '+event.name:'Диоген — '+event.name,description=buildCalendarBookingDescription(event,slotStart,bookings);
  bookingEvent.setTitle(title);bookingEvent.setDescription(description);synchronizeCalendar();return bookingEvent;
}

function syncDiogenCalendarRange(event,rangeStart,rangeEnd,bookings) {
  let cursor=new Date(rangeStart);
  while(cursor.getTime()<rangeEnd.getTime()){const hourStart=new Date(cursor),hourEnd=addMinutes(hourStart,SLOT_STEP_MINUTES);syncDiogenCalendarHour(event,hourStart,hourEnd,bookings);cursor=hourEnd;}
  synchronizeCalendar();
}

function syncDiogenCalendarHour(event,hourStart,hourEnd,bookings) {
  const calendar=getBookingCalendar(),slotId=makeSlotId(event,hourStart),events=calendar.getEvents(hourStart,hourEnd),bookingEvents=[];
  events.forEach(function(calendarEvent){if(!isBookingCalendarEvent(calendarEvent))return;const description=String(calendarEvent.getDescription()||'');if(description.indexOf(BOOKING_MARKER+slotId)===0)bookingEvents.push(calendarEvent);});
  const hourBookings=getBookingsForExactSlot(bookings,hourStart,event),occupancy=calculateOccupancy(hourBookings);
  if(occupancy<=0){bookingEvents.forEach(function(calendarEvent){calendarEvent.deleteEvent();});return;}
  const title='Диоген — '+event.name,description=buildCalendarBookingDescription(event,hourStart,bookings);
  let bookingEvent=bookingEvents.length?bookingEvents[0]:null;
  bookingEvents.slice(1).forEach(function(calendarEvent){calendarEvent.deleteEvent();});
  if(!bookingEvent)bookingEvent=calendar.createEvent(title,hourStart,hourEnd,{description:description});
  else{bookingEvent.setTitle(title);bookingEvent.setDescription(description);}
}

function buildCalendarBookingDescription(event,slotStart,bookings) {
  const slotId=makeSlotId(event,slotStart),slotBookings=getBookingsForExactSlot(bookings,slotStart,event),capacity=getCapacity(event),occupied=calculateOccupancy(slotBookings),lines=[];
  lines.push(BOOKING_MARKER+slotId);
  slotBookings.forEach(function(booking){const name=String(booking.name||'').trim(),tickets=Number(booking.tickets)||0;lines.push(name+' ('+tickets+')');});
  lines.push(occupied+'/'+capacity);return lines.join('\n');
}

function restoreCalendarFreeSlot(event,slotStart) {
  const calendar=getBookingCalendar(),slotEnd=addMinutes(slotStart,Number(event.duration)*60),events=calendar.getEvents(slotStart,slotEnd);
  events.forEach(function(calendarEvent){if(!isBookingCalendarEvent(calendarEvent))return;const description=String(calendarEvent.getDescription()||'');if(description.indexOf(BOOKING_MARKER+makeSlotId(event,slotStart))===0)calendarEvent.deleteEvent();});
  let cursor=new Date(slotStart);
  while(cursor.getTime()<slotEnd.getTime()){
    const hourEnd=addMinutes(cursor,SLOT_STEP_MINUTES),existing=calendar.getEvents(cursor,hourEnd);
    const alreadyFree=existing.some(function(calendarEvent){return String(calendarEvent.getTitle()||'').trim()===FREE_EVENT_TITLE&&calendarEvent.getStartTime().getTime()===cursor.getTime()&&calendarEvent.getEndTime().getTime()===hourEnd.getTime();});
    const hasBlock=existing.some(function(calendarEvent){const title=String(calendarEvent.getTitle()||'').trim();return title!==FREE_EVENT_TITLE&&!isBookingCalendarEvent(calendarEvent);});
    if(!alreadyFree&&!hasBlock)calendar.createEvent(FREE_EVENT_TITLE,cursor,hourEnd);
    cursor=hourEnd;
  }
}

function isBookingCalendarEvent(event){const description=String(event.getDescription()||'');return description.indexOf(BOOKING_MARKER)===0;}
function intersectsCalendarEvents(start,end,events){return events.some(function(event){const eventStart=event.getStartTime(),eventEnd=event.getEndTime();return start.getTime()<eventEnd.getTime()&&end.getTime()>eventStart.getTime();});}
function intersectsAnyBlock(start,end,blocks){return blocks.some(function(block){const blockStart=new Date(block.start),blockEnd=new Date(block.end);return start.getTime()<blockEnd.getTime()&&end.getTime()>blockStart.getTime();});}

function mergeFreeWindows(events){
  const windows=events.map(function(event){return{start:event.getStartTime(),end:event.getEndTime()};}).sort(function(a,b){return a.start.getTime()-b.start.getTime();});
  const result=[];
  windows.forEach(function(window){if(result.length===0){result.push({start:window.start,end:window.end});return;}const last=result[result.length-1];if(window.start.getTime()<=last.end.getTime()){if(window.end.getTime()>last.end.getTime())last.end=window.end;}else result.push({start:window.start,end:window.end});});
  return result;
}
function floorToHour(date){const result=new Date(date);result.setMinutes(0,0,0);return result;}
function ceilToHour(date){const result=new Date(date);if(result.getMinutes()!==0||result.getSeconds()!==0||result.getMilliseconds()!==0)result.setHours(result.getHours()+1);result.setMinutes(0,0,0);return result;}

function makeSlotId(event,start){return[getFormat(event),event.name||'Диоген',Utilities.formatDate(start,Session.getScriptTimeZone(),'yyyyMMdd-HHmm')].join('_');}
function makeBookingSlotId(event,start,end){return[getFormat(event),event.name||'Диоген',Utilities.formatDate(start,Session.getScriptTimeZone(),'yyyyMMdd-HHmm'),Utilities.formatDate(end,Session.getScriptTimeZone(),'yyyyMMdd-HHmm')].join('::');}
function parseBookingSlotId(slotId){const parts=String(slotId||'').split('::');if(parts.length!==4)return null;return{format:parts[0],eventName:parts[1],start:parts[2],end:parts[3]};}
function parseOldSlotId(slotId){const text=String(slotId||''),match=text.match(/^([^_]+)_(.+)_(\d{8}-\d{4})$/);if(!match)return null;return{format:match[1],eventName:match[2],start:match[3]};}

function getBookingInterval(booking,event){
  const parsed=parseBookingSlotId(booking.slotId);
  if(parsed&&parsed.end){const start=parseCompactDateTime(parsed.start),end=parseCompactDateTime(parsed.end);if(start&&end&&end.getTime()>start.getTime())return{start:start,end:end};}
  const start=parseBookingDateTime(booking.date,booking.time),end=addMinutes(start,Number(event.duration)*60);return{start:start,end:end};
}

function getHourlyRanges(start,end){const ranges=[];let cursor=new Date(start);while(cursor.getTime()<end.getTime()){const hourEnd=addMinutes(cursor,SLOT_STEP_MINUTES);ranges.push({start:new Date(cursor),end:new Date(hourEnd)});cursor=hourEnd;}return ranges;}
function parseCompactDateTime(value){const match=String(value||'').match(/^(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})$/);if(!match)return null;return parseBookingDateTime(match[3]+'.'+match[2]+'.'+match[1],match[4]+':'+match[5]);}
function parseDuration(value){if(value===null||value===undefined||value==='')return 0;if(typeof value==='number')return value;const text=String(value).trim().toLowerCase(),numberMatch=text.match(/(\d+(?:[.,]\d+)?)/);if(!numberMatch)return 0;const number=Number(numberMatch[1].replace(',','.'));if(text.indexOf('мин')!==-1)return number/60;return number;}
function parseAmount(value){if(value===null||value===undefined||value==='')return 0;const numberMatch=String(value).match(/\d+/);if(!numberMatch)return 0;return Number(numberMatch[0]);}
function getNextBookingId(sheet){const lastRow=sheet.getLastRow();if(lastRow<2)return 1;const values=sheet.getRange(2,1,lastRow-1,1).getValues();let maxId=0;values.forEach(function(row){const number=Number(row[0]);if(Number.isFinite(number)&&number>maxId)maxId=number;});return maxId+1;}
function parseBookingDateTime(dateString,timeString){const timezone=Session.getScriptTimeZone(),result=Utilities.parseDate(String(dateString)+' '+String(timeString),timezone,'dd.MM.yyyy HH:mm');if(!result||isNaN(result.getTime()))throw new Error('Некорректная дата или время: '+dateString+' '+timeString);return result;}
function normalizeDateValue(value){if(value instanceof Date&&!isNaN(value.getTime()))return formatDate(value);if(value===null||value===undefined)return '';return String(value).trim();}
function normalizeTimeValue(value){if(value instanceof Date&&!isNaN(value.getTime()))return formatTime(value);if(value===null||value===undefined)return '';return String(value).trim();}
function formatDate(date){return Utilities.formatDate(date,Session.getScriptTimeZone(),'dd.MM.yyyy');}
function formatTime(date){return Utilities.formatDate(date,Session.getScriptTimeZone(),'HH:mm');}
function formatDateTime(date){return Utilities.formatDate(date,Session.getScriptTimeZone(),'dd.MM.yyyy HH:mm');}
function startOfDay(date){const result=new Date(date);result.setHours(0,0,0,0);return result;}
function addMinutes(date,minutes){return new Date(date.getTime()+minutes*60*1000);}

function testBooking(){
  const data={action:'book',telegramId:'99999',name:'Тест',activity:'masterclass',eventName:'Кошелёк',date:'21.09.2026',time:'11:00',tickets:2};
  const result=createBooking(data);Logger.log(JSON.stringify(result,null,2));
}