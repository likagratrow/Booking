(function(){
  const originalRenderCalendar=window.renderCalendar;
  if(typeof originalRenderCalendar!=='function')return;

  const style=document.createElement('style');
  style.textContent='.calendar-grid{grid-template-columns:3120px;width:3120px;min-width:3120px}.calendar-days,.calendar-columns{width:3120px}.calendar-slot.generic{cursor:default;pointer-events:none}';
  document.head.appendChild(style);

  function normalizeInitialCalendar(){
    if(selectedCalendar.format)return;
    const geometry=ensureCalendarGeometry();
    document.querySelectorAll('#calendar .calendar-column').forEach((column,dayIndex)=>{
      column.querySelectorAll('.calendar-cell').forEach((cell,hourIndex)=>{
        const day=geometry.days[dayIndex],hour=geometry.hours[hourIndex];
        if(!day||hour==null)return;
        const key=slotKey(dateString(day),timeKey(hour*60));
        const matches=bookingSlotByKey.get(key)||[];
        const available=matches.some(s=>s.available&&Number(s.free||0)>0);
        if(!available){cell.innerHTML='';return;}
        cell.innerHTML=`<div class="calendar-slot free generic"><strong>${timeKey(hour*60)}–${timeKey((hour+1)*60)}</strong><small>свободно</small></div>`;
      });
    });
  }

  window.renderCalendar=function(){
    originalRenderCalendar();
    normalizeInitialCalendar();
  };

  normalizeInitialCalendar();
})();
