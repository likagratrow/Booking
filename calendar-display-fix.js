(function(){
  const style=document.createElement('style');
  style.textContent=`
    .calendar-grid{grid-template-columns:3120px;width:3120px;min-width:3120px}
    .calendar-days,.calendar-columns{width:3120px}
    .calendar-slot.generic{cursor:pointer;pointer-events:auto}
    .calendar-slot.selected-unavailable{opacity:.45;cursor:default}
    .calendar-slot.candidate{box-shadow:0 0 0 2px color-mix(in srgb,var(--tg-accent) 70%,transparent);}
    .booking-time-wrap{display:flex;align-items:center;justify-content:center;gap:3px}
    .booking-duration-inline{display:flex!important;flex-direction:column;align-items:center;justify-content:center;gap:0;margin-left:1px;line-height:1}
    .booking-duration-inline .booking-arrow{width:16px!important;height:13px!important;min-height:13px!important;padding:0!important;border:0!important;background:transparent!important;font-size:9px!important;line-height:10px!important;color:var(--tg-hint)!important}
    .booking-duration-inline .booking-arrow:disabled{opacity:.25}
    .booking-tickets-control{display:grid;grid-template-columns:36px minmax(70px,1fr) 36px;align-items:center;gap:6px;width:100%}
    .booking-tickets-control #booking-tickets-input{width:100%;text-align:center}
    .booking-ticket-step{width:36px;height:36px;padding:0;color:var(--tg-text);background:var(--tg-secondary-bg);border:1px solid var(--tg-separator);border-radius:8px;font-size:20px;line-height:1;cursor:pointer}
    .booking-max-mark{display:inline-flex;align-items:center;justify-content:center;margin-left:6px;padding:2px 4px;color:var(--tg-accent);border:1px solid var(--tg-accent);border-radius:4px;font-size:8px;font-weight:800;line-height:1;vertical-align:middle}
    #booking-modal .eyebrow.hidden{display:none}
  `;
  document.head.appendChild(style);

  function activityOf(s){const a=norm(s?.activity||'');if(a==='диоген')return'Диоген';if(a==='обсудить заказ')return'Обсудить заказ';if(a==='masterclass'||a==='мастер-класс')return'МК';return norm(s?.format||'');}
  function matchesSelected(s){
    if(!selectedCalendar.format)return false;
    if(selectedCalendar.format==='МК')return activityOf(s)==='МК'&&norm(s.name)===norm(selectedCalendar.name);
    return activityOf(s)===selectedCalendar.format;
  }
  function masterclassFits(slot){
    const event=events.find(e=>norm(e.activity)==='masterclass'&&norm(e.name)===norm(selectedCalendar.name));
    if(!event)return false;
    const needed=Math.max(1,durationHours(event)),start=slotMinutes(slot),date=String(slot.start||'').split(' ')[0];
    if(start+needed*60>CALENDAR_END_HOUR*60)return false;
    for(let i=0;i<needed;i++){
      const part=bookingSlotByKey.get(slotKey(date,timeKey(start+i*60)))?.find(s=>activityOf(s)==='МК'&&norm(s.name)===norm(selectedCalendar.name));
      const required=i===0?2:1;
      if(!part||!part.available||Number(part.free||0)<required)return false;
    }
    return true;
  }
  function candidateForCell(cell,matches){
    if(!selectedCalendar.format)return null;
    const selected=matches.filter(matchesSelected).find(s=>s.available&&Number(s.free||0)>0);
    if(!selected)return null;
    if(selectedCalendar.format==='МК')return masterclassFits(selected)?selected:null;
    return selected;
  }
  function displaySlot(matches){return matches.find(s=>s.available&&Number(s.free||0)>0)||matches.find(s=>s.available)||matches[0]||null;}
  function renderFixedCalendar(){
    const geometry=ensureCalendarGeometry();
    calendar.innerHTML=`<div class="calendar-scroll"><div class="calendar-grid"><div class="calendar-corner"></div><div class="calendar-days">${geometry.days.map(d=>`<div class="calendar-day-head">${esc(dateText(d))}</div>`).join('')}</div><div class="calendar-times"></div><div class="calendar-columns">${geometry.days.map(d=>{const date=dateString(d);return`<div class="calendar-column"><div class="calendar-hour-lines">${geometry.hours.map(hour=>{const cell=geometry.cells.get(slotKey(date,timeKey(hour*60))),matches=bookingSlotByKey.get(cell.key)||[],booking=displaySlot(matches),candidate=candidateForCell(cell,matches),generic=!selectedCalendar.format&&matches.some(s=>s.available&&Number(s.free||0)>0),state=booking?slotClass(booking):'';if(!booking&&!generic)return`<div class="calendar-cell"></div>`;if(!selectedCalendar.format)return`<div class="calendar-cell"><button class="calendar-slot free generic" type="button"><strong>${esc(timeKey(hour*60))}–${esc(timeKey((hour+1)*60))}</strong><small>свободно</small></button></div>`;const label=booking?`${activityOf(booking)==='МК'?'МК — '+booking.name:activityOf(booking)} · ${slotTime(booking)}–${timeKey(Math.min(slotEndMinutes(booking),slotMinutes(cell)+60))} · ${slotLabel(booking)}`:'';return`<div class="calendar-cell"><button class="calendar-slot ${state}${candidate?' candidate':' selected-unavailable'}" type="button" data-slot-id="${esc(booking?.slotId||'')}" ${candidate?'':'disabled'}><strong>${esc(timeKey(hour*60))}–${esc(timeKey((hour+1)*60))}</strong><small>${esc(label||'недоступно')}</small></button></div>`;}).join('')}</div></div>`;}).join('')}</div></div></div>`;
    calendar.querySelectorAll('.calendar-slot[data-slot-id]:not(.generic)').forEach(b=>b.addEventListener('click',()=>{const s=bookingSlotById.get(b.dataset.slotId);if(s)selectSlot(s);}));
    calendar.querySelectorAll('.calendar-slot.generic').forEach(b=>b.addEventListener('click',()=>{showMessage('Сначала выберите активность.');document.querySelector('#activities')?.scrollIntoView({behavior:'smooth',block:'center'});}));
  }
  window.renderCalendar=renderFixedCalendar;
  renderFixedCalendar();
})();
