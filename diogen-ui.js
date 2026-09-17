(function(){
  function isDiogenKey(key){
    const raw=norm(key);
    const activity=activityByKey(raw)||activities.find(function(a){return norm(a.title)===raw;});
    return Boolean(activity&&norm(activity.key)==='diogen');
  }

  window.openExistingBooking=function(event,clickedDate='',clickedHour=NaN){
    if(!event||eventHasStarted(event))return;
    const name=bookingEventName(event),eventData=eventForContext({name:name});
    if(!eventData)return;
    const occupancy=bookingEventOccupancy(event),free=Math.max(0,Number(occupancy?.capacity||0)-Number(occupancy?.occupied||0));
    if(free<1)return;
    const context={activity:eventData.activity,name:eventData.name};
    const date=eventDate(event),existingStart=eventStartMinutes(event),existingEnd=eventEndMinutes(event);
    if(!date||!Number.isFinite(existingStart)||!Number.isFinite(existingEnd)||existingEnd<=existingStart)return;
    const diogen=isDiogenKey(context.activity);
    const start=diogen&&clickedDate===date&&Number.isFinite(clickedHour)?Math.max(existingStart,Number(clickedHour)*60):existingStart;
    if(start>=existingEnd)return;
    const endOptions=diogen?diogenEndOptions(date,start,context):[existingEnd];
    if(!endOptions.length)return;
    const slot=bookingPayload(date,start,endOptions[endOptions.length-1],context);
    slot.name=eventData.name;
    slot.free=free;
    slot.capacity=Number(occupancy.capacity);
    slot.minTickets=1;
    slot.available=true;
    openBooking(slot,{diogen:diogen,endOptions:endOptions,context:context});
  };
})();
