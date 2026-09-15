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


// ============================================================
// НАСТРОЙКИ
// ============================================================

const EVENTS_SHEET_NAME = 'Ивенты';
const BOOKING_SHEET_NAME = 'Брони';
const BOOKING_CALENDAR_NAME = 'Странные Вещи Букинг';

const FREE_EVENT_TITLE = 'Свободно';

const SLOT_STEP_MINUTES = 60;

const LOOKAHEAD_DAYS = 30;

// Техническая метка событий, созданных системой бронирования.
// Пользователь её в календаре не видит.
const BOOKING_MARKER = 'ST_BOOKING_SLOT:';


// ============================================================
// HTTP GET
// ============================================================

function doGet(e) {
  try {
    const result = getBookingData();

    return jsonResponse(result);

  } catch (error) {
    return jsonResponse({
      ok: false,
      error: error.message
    });
  }
}


// ============================================================
// HTTP POST
// ============================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Не получены данные POST-запроса.');
    }

    const data = JSON.parse(e.postData.contents);

    if (!data.action) {
      throw new Error('Не указано действие.');
    }

    let result;

    if (data.action === 'book') {
      result = createBooking(data);

    } else if (data.action === 'cancel') {
      result = cancelBooking(data);

    } else {
      throw new Error(
        'Неизвестное действие: ' + data.action
      );
    }

    return jsonResponse(result);

  } catch (error) {
    return jsonResponse({
      ok: false,
      error: error.message
    });
  }
}


// ============================================================
// JSON RESPONSE
// ============================================================

function jsonResponse(data) {
  return ContentService
    .createTextOutput(
      JSON.stringify(data, null, 2)
    )
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// ОСНОВНАЯ ФУНКЦИЯ
// ============================================================

function getBookingData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'Не удалось получить текущую таблицу.'
    );
  }

  const eventsSheet = spreadsheet.getSheetByName(
    EVENTS_SHEET_NAME
  );

  if (!eventsSheet) {
    throw new Error(
      'Лист "' +
      EVENTS_SHEET_NAME +
      '" не найден.'
    );
  }

  const bookingsSheet = spreadsheet.getSheetByName(
    BOOKING_SHEET_NAME
  );

  if (!bookingsSheet) {
    throw new Error(
      'Лист "' +
      BOOKING_SHEET_NAME +
      '" не найден.'
    );
  }

  // Сохраняем старую механику календаря:
  // большое окно "Свободно" перед чтением приводится
  // к часовым участкам с шагом SLOT_STEP_MINUTES.
  synchronizeCalendar();

  const events = readEvents(eventsSheet);
  const bookings = readBookings(bookingsSheet);
  const calendar = readCalendar();

  return {
    ok: true,

    settings: {
      slotStepMinutes: SLOT_STEP_MINUTES,
      lookaheadDays: LOOKAHEAD_DAYS
    },

    events: events,

    bookings: {
      count: bookings.length,
      items: bookings
    },

    calendar: {
      name: BOOKING_CALENDAR_NAME,
      events: calendar.events,
      freeWindows: calendar.freeWindows,
      blocks: calendar.blocks
    }
  };
}


// ============================================================
// СОЗДАНИЕ БРОНИ
// ============================================================

function createBooking(data) {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();

    const eventsSheet =
      spreadsheet.getSheetByName(
        EVENTS_SHEET_NAME
      );

    const bookingsSheet =
      spreadsheet.getSheetByName(
        BOOKING_SHEET_NAME
      );

    if (!eventsSheet) {
      throw new Error(
        'Лист "' +
        EVENTS_SHEET_NAME +
        '" не найден.'
      );
    }

    if (!bookingsSheet) {
      throw new Error(
        'Лист "' +
        BOOKING_SHEET_NAME +
        '" не найден.'
      );
    }

    /*
     * Здесь НЕ вызываем synchronizeCalendar().
     *
     * Сначала читаем реальное текущее состояние календаря
     * и проверяем, можно ли поставить бронь.
     */
    const events = readEvents(eventsSheet);
    const bookings = readBookings(bookingsSheet);
    const calendar = readCalendar();

    const tickets = Number(data.tickets);

    if (!Number.isInteger(tickets) || tickets < 1) {
      throw new Error(
        'Количество билетов должно быть целым числом от 1.'
      );
    }

    const telegramId =
      data.telegramId === undefined ||
      data.telegramId === null
        ? ''
        : String(data.telegramId).trim();

    const userName =
      data.name === undefined ||
      data.name === null
        ? ''
        : String(data.name).trim();

    if (!telegramId) {
      throw new Error(
        'Не указан Telegram ID.'
      );
    }

    if (!userName) {
      throw new Error(
        'Не указано имя.'
      );
    }

    const format =
      normalizeFormat(data.format);

    const eventName =
      String(
        data.eventName || ''
      ).trim();

    const date =
      String(
        data.date || ''
      ).trim();

    const time =
      String(
        data.time || ''
      ).trim();

    if (!date || !time) {
      throw new Error(
        'Не указаны дата и время.'
      );
    }

    const event = findEvent(
      events,
      format,
      eventName
    );

    if (!event) {
      throw new Error(
        'Выбранное мероприятие не найдено.'
      );
    }

    const slotStart =
      parseBookingDateTime(
        date,
        time
      );

    const now = new Date();

    if (
      slotStart.getTime() <= now.getTime()
    ) {
      throw new Error(
        'Время начала этого слота уже прошло.'
      );
    }

    const capacity =
      getCapacity(event);

    let slotEnd;

    /*
     * МК использует фиксированную длительность
     * из листа "Ивенты".
     *
     * Диоген получает выбранное пользователем
     * окончание. Если оно не передано, по умолчанию
     * используется один час.
     */
    if (
      getFormat(event) === 'Диоген'
    ) {

      const endTime =
        String(
          data.endTime || ''
        ).trim();

      if (endTime) {

        slotEnd =
          parseBookingDateTime(
            date,
            endTime
          );

      } else {

        slotEnd =
          addMinutes(
            slotStart,
            SLOT_STEP_MINUTES
          );
      }

      if (
        slotEnd.getTime() <=
        slotStart.getTime()
      ) {
        throw new Error(
          'Время окончания должно быть позже времени начала.'
        );
      }

      if (
        slotStart.getMinutes() !== 0 ||
        slotEnd.getMinutes() !== 0 ||
        slotStart.getSeconds() !== 0 ||
        slotEnd.getSeconds() !== 0
      ) {
        throw new Error(
          'Диоген можно бронировать только целыми часами.'
        );
      }

    } else {

      slotEnd =
        addMinutes(
          slotStart,
          Number(event.duration) * 60
        );
    }

    /*
     * Проверяем каждый час выбранного интервала.
     *
     * Для МК диапазон фиксированный.
     * Для Диогена каждый час проверяется отдельно,
     * поэтому другой человек может присоединиться
     * только к тем часам, где ещё есть вместимость.
     */
    const affectedHours =
      getHourlyRanges(
        slotStart,
        slotEnd
      );

    if (
      affectedHours.length === 0
    ) {
      throw new Error(
        'Некорректный интервал бронирования.'
      );
    }

    affectedHours.forEach(
      function(hour) {

        if (
          intersectsAnyBlock(
            hour.start,
            hour.end,
            calendar.blocks
          )
        ) {
          throw new Error(
            'В выбранном интервале есть занятое время.'
          );
        }

        const hourBookings =
          getBookingsForExactSlot(
            bookings,
            hour.start,
            event
          );

        const occupancy =
          calculateOccupancy(
            hourBookings
          );

        const free =
          Math.max(
            0,
            capacity - occupancy
          );

        if (
          free <= 0
        ) {
          throw new Error(
            'В выбранном интервале есть час без свободных мест.'
          );
        }

        if (
          tickets > free
        ) {
          throw new Error(
            'В одном из выбранных часов свободно только ' +
            free +
            ' мест.'
          );
        }

        /*
         * Если этот час сейчас полностью свободен,
         * он обязательно должен быть разрешён
         * календарём "Свободно".
         */
        if (
          occupancy === 0 &&
          !hasFreeCalendarHours(
            hour.start,
            1,
            calendar.freeWindows,
            calendar.blocks
          )
        ) {
          throw new Error(
            'Один из выбранных часов больше недоступен.'
          );
        }
      }
    );

    /*
     * Если слот полностью свободен — для нового МК
     * сохраняем старое правило минимум 2 билета.
     *
     * Для Диогена минимум всегда 1.
     */
    const totalExistingOccupancy =
      calculateOccupancy(
        getBookingsForExactSlot(
          bookings,
          slotStart,
          event
        )
      );

    if (
      totalExistingOccupancy === 0
    ) {

      const minTickets =
        getFormat(event) === 'Диоген'
          ? 1
          : 2;

      if (
        tickets < minTickets
      ) {
        throw new Error(
          'Для нового ' +
          (
            getFormat(event) === 'МК'
              ? 'МК'
              : 'Диогена'
          ) +
          ' нужно забронировать минимум ' +
          minTickets +
          ' мест.'
        );
      }
    }

    /*
     * Для МК не разрешаем смешивать разные мероприятия
     * в одном и том же времени.
     *
     * Для Диогена проверяем каждый час отдельно:
     * другой вариант Диогена не должен пересекаться
     * с выбранным интервалом.
     */
    affectedHours.forEach(
      function(hour) {

        const conflictingBookings =
          getAllBookingsAtStart(
            bookings,
            hour.start
          );

        conflictingBookings.forEach(
          function(booking) {

            if (
              normalizeFormat(
                booking.format
              ) !== format
            ) {
              throw new Error(
                'Это время уже занято другим форматом.'
              );
            }

            if (
              format === 'МК' &&
              String(
                booking.activity || ''
              ).trim() !== event.name
            ) {
              throw new Error(
                'Это время уже занято другим мастер-классом.'
              );
            }

            if (
              format === 'Диоген'
            ) {

              const bookingEvent =
                findEventForBooking(
                  booking
                );

              if (
                bookingEvent &&
                String(
                  bookingEvent.name
                ).trim() !==
                String(
                  event.name
                ).trim()
              ) {
                throw new Error(
                  'Это время уже занято другим форматом Диогена.'
                );
              }
            }
          }
        );
      }
    );

    const bookingId =
      getNextBookingId(
        bookingsSheet
      );

    /*
     * Для новой брони сохраняем диапазон
     * в slot_id. Структура листа при этом
     * не меняется.
     */
    const slotId =
      makeBookingSlotId(
        event,
        slotStart,
        slotEnd
      );

    const activityForSheet =
      getFormat(event) === 'МК'
        ? event.name
        : '—';

    bookingsSheet.appendRow([
      bookingId,
      telegramId,
      userName,
      getFormat(event),
      activityForSheet,
      date,
      time,
      tickets,
      'active',
      slotId
    ]);

    const bookingRow =
      bookingsSheet.getLastRow();

    const updatedBookingsForCalendar =
      readBookings(
        bookingsSheet
      );

    try {

      if (
        getFormat(event) === 'Диоген'
      ) {

        syncDiogenCalendarRange(
          event,
          slotStart,
          slotEnd,
          updatedBookingsForCalendar
        );

      } else {

        if (
          totalExistingOccupancy === 0
        ) {

          createCalendarBooking(
            event,
            slotStart,
            updatedBookingsForCalendar
          );

        } else {

          updateCalendarBooking(
            event,
            slotStart,
            updatedBookingsForCalendar
          );
        }
      }

    } catch (calendarError) {

      try {

        bookingsSheet.deleteRow(
          bookingRow
        );

      } catch (rollbackError) {

        Logger.log(
          'Не удалось откатить строку брони: ' +
          rollbackError.message
        );
      }

      throw calendarError;
    }

    const updatedBookings =
      readBookings(
        bookingsSheet
      );

    const updatedCalendar =
      readCalendar();

    const updatedSlotBookings =
      getBookingsForExactSlot(
        updatedBookings,
        slotStart,
        event
      );

    const updatedOccupancy =
      calculateOccupancy(
        updatedSlotBookings
      );

    return {
      ok: true,

      booking: {
        id: bookingId,
        telegramId: telegramId,
        name: userName,
        format: getFormat(event),
        eventName: event.name,
        date: date,
        time: time,
        endTime: formatTime(slotEnd),
        tickets: tickets,
        status: 'active',
        slotId: slotId
      },

      slot: {
        slotId: slotId,
        start: formatDateTime(slotStart),
        end: formatDateTime(slotEnd),
        capacity: capacity,
        occupied: updatedOccupancy,
        free: Math.max(
          0,
          capacity - updatedOccupancy
        )
      },

      calendar: updatedCalendar
    };

  } finally {

    lock.releaseLock();
  }
}


// ============================================================
// ОТМЕНА БРОНИ
// ============================================================

function cancelBooking(data) {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();

    const bookingsSheet =
      spreadsheet.getSheetByName(
        BOOKING_SHEET_NAME
      );

    if (!bookingsSheet) {
      throw new Error(
        'Лист "' +
        BOOKING_SHEET_NAME +
        '" не найден.'
      );
    }

    const bookingId =
      data.bookingId !== undefined
        ? data.bookingId
        : data.id;

    if (
      bookingId === undefined ||
      bookingId === null ||
      bookingId === ''
    ) {
      throw new Error(
        'Не указан ID брони.'
      );
    }

    const lastRow =
      bookingsSheet.getLastRow();

    const values =
      lastRow < 2
        ? []
        : bookingsSheet
          .getRange(
            2,
            1,
            lastRow - 1,
            10
          )
          .getValues();

    let rowNumber = null;
    let booking = null;

    for (
      let i = 0;
      i < values.length;
      i++
    ) {

      const row = values[i];

      if (
        String(row[0]).trim() ===
        String(bookingId).trim()
      ) {

        rowNumber = i + 2;

        booking = {
          id: row[0],
          telegramId: row[1],
          name: row[2],
          format: row[3],
          activity: row[4],
          date: normalizeDateValue(row[5]),
          time: normalizeTimeValue(row[6]),
          tickets: Number(row[7]) || 0,
          status: String(row[8] || '')
            .trim()
            .toLowerCase(),
          slotId: row[9] || null
        };

        break;
      }
    }

    if (!booking) {
      throw new Error(
        'Бронь не найдена.'
      );
    }

    if (
      booking.status === 'cancelled'
    ) {
      throw new Error(
        'Эта бронь уже отменена.'
      );
    }

    if (
      data.telegramId !== undefined &&
      String(data.telegramId).trim() !==
      String(booking.telegramId).trim()
    ) {
      throw new Error(
        'Эта бронь принадлежит другому пользователю.'
      );
    }

    const slotStart =
      parseBookingDateTime(
        booking.date,
        booking.time
      );

    const event =
      findEventForBooking(
        booking
      );

    if (!event) {
      throw new Error(
        'Не удалось определить мероприятие для отменённой брони.'
      );
    }

    const bookingInterval =
      getBookingInterval(
        booking,
        event
      );

    bookingsSheet
      .getRange(
        rowNumber,
        9
      )
      .setValue('cancelled');

    const bookings =
      readBookings(
        bookingsSheet
      );

    try {

      if (
        getFormat(event) === 'Диоген'
      ) {

        syncDiogenCalendarRange(
          event,
          bookingInterval.start,
          bookingInterval.end,
          bookings
        );

      } else {

        const remaining =
          getBookingsForExactSlot(
            bookings,
            slotStart,
            event
          );

        if (remaining.length === 0) {

          restoreCalendarFreeSlot(
            event,
            slotStart
          );

        } else {

          updateCalendarBooking(
            event,
            slotStart,
            bookings
          );
        }
      }

    } catch (calendarError) {

      try {

        bookingsSheet
          .getRange(
            rowNumber,
            9
          )
          .setValue('active');

      } catch (rollbackError) {

        Logger.log(
          'Не удалось откатить статус отменённой брони: ' +
          rollbackError.message
        );
      }

      throw calendarError;
    }

    const updatedCalendar = readCalendar();

    return {
      ok: true,

      booking: {
        id: booking.id,
        status: 'cancelled'
      },

      calendar: updatedCalendar
    };

  } finally {

    lock.releaseLock();
  }
}


// ============================================================
// ПОИСК МЕРОПРИЯТИЯ
// ============================================================

function findEvent(
  events,
  format,
  eventName
) {

  return events.find(
    function(event) {

      if (
        getFormat(event) !== format
      ) {
        return false;
      }

      return (
        String(event.name).trim() ===
        String(eventName).trim()
      );
    }
  ) || null;
}


// ============================================================
// ПОИСК МЕРОПРИЯТИЯ ДЛЯ БРОНИ
// ============================================================

function findEventForBooking(
  booking
) {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      EVENTS_SHEET_NAME
    );

  const events =
    readEvents(sheet);

  const format =
    normalizeFormat(
      booking.format
    );

  if (format === 'МК') {

    return findEvent(
      events,
      'МК',
      booking.activity
    );
  }

  /*
   * Новые Диоген-брони имеют slot_id:
   *
   * Диоген::Соло::20260911-1000::20260911-1300
   */
  if (
    booking.slotId &&
    String(booking.slotId).indexOf('::') !== -1
  ) {

    const parsed =
      parseBookingSlotId(
        booking.slotId
      );

    if (
      parsed &&
      parsed.eventName
    ) {

      const found =
        findEvent(
          events,
          'Диоген',
          parsed.eventName
        );

      if (found) {
        return found;
      }
    }
  }

  /*
   * Старые Диоген slot_id:
   *
   * Диоген_Соло_20260911-1000
   *
   * Берём название между форматом
   * и последним временным идентификатором.
   */
  if (
    booking.slotId
  ) {

    const parsedOld =
      parseOldSlotId(
        booking.slotId
      );

    if (
      parsedOld &&
      parsedOld.eventName
    ) {

      const found =
        findEvent(
          events,
          'Диоген',
          parsedOld.eventName
        );

      if (found) {
        return found;
      }
    }
  }

  const candidates =
    events.filter(
      function(event) {
        return (
          getFormat(event) ===
          'Диоген'
        );
      }
    );

  return candidates.length
    ? candidates[0]
    : null;
}


// ============================================================
// СИНХРОНИЗАЦИЯ КАЛЕНДАРЯ
// ============================================================

function synchronizeCalendar() {

  const calendar =
    getBookingCalendar();

  const now =
    new Date();

  const from =
    startOfDay(now);

  const to =
    new Date(from);

  to.setDate(
    to.getDate() +
    LOOKAHEAD_DAYS +
    1
  );

  const events =
    calendar.getEvents(
      from,
      to
    );

  const freeEvents =
    events.filter(
      function(event) {
        return (
          String(
            event.getTitle() || ''
          ).trim() ===
          FREE_EVENT_TITLE
        );
      }
    );

  const bookingEvents =
    events.filter(
      function(event) {
        return isBookingCalendarEvent(
          event
        );
      }
    );

  const blocks =
    events.filter(
      function(event) {

        const title =
          String(
            event.getTitle() || ''
          ).trim();

        if (
          title === FREE_EVENT_TITLE
        ) {
          return false;
        }

        return !isBookingCalendarEvent(
          event
        );
      }
    );

  if (
    freeEvents.length === 0
  ) {
    return;
  }

  const mergedWindows =
    mergeFreeWindows(
      freeEvents
    );

  const desiredFreeSlots = {};

  mergedWindows.forEach(
    function(window) {

      const roundedStart =
        floorToHour(
          window.start
        );

      const roundedEnd =
        ceilToHour(
          window.end
        );

      let cursor =
        new Date(
          roundedStart
        );

      while (
        cursor.getTime() <
        roundedEnd.getTime()
      ) {

        const chunkStart =
          new Date(cursor);

        const chunkEnd =
          addMinutes(
            chunkStart,
            SLOT_STEP_MINUTES
          );

        const conflictsWithBlock =
          intersectsCalendarEvents(
            chunkStart,
            chunkEnd,
            blocks
          );

        const conflictsWithBooking =
          intersectsCalendarEvents(
            chunkStart,
            chunkEnd,
            bookingEvents
          );

        if (
          !conflictsWithBlock &&
          !conflictsWithBooking
        ) {

          const key =
            chunkStart.getTime() +
            '_' +
            chunkEnd.getTime();

          desiredFreeSlots[key] = {
            start: chunkStart,
            end: chunkEnd
          };
        }

        cursor =
          addMinutes(
            cursor,
            SLOT_STEP_MINUTES
          );
      }
    }
  );

  const existingDesiredSlots = {};

  freeEvents.forEach(
    function(event) {

      const start =
        event.getStartTime();

      const end =
        event.getEndTime();

      const key =
        start.getTime() +
        '_' +
        end.getTime();

      if (
        desiredFreeSlots[key]
      ) {

        if (
          !existingDesiredSlots[key]
        ) {

          existingDesiredSlots[key] =
            event;

        } else {

          event.deleteEvent();
        }

      } else {

        event.deleteEvent();
      }
    }
  );

  Object.keys(
    desiredFreeSlots
  ).forEach(
    function(key) {

      if (
        existingDesiredSlots[key]
      ) {
        return;
      }

      const slot =
        desiredFreeSlots[key];

      calendar.createEvent(
        FREE_EVENT_TITLE,
        slot.start,
        slot.end
      );
    }
  );
}


// ============================================================
// ЧТЕНИЕ КАЛЕНДАРЯ
// ============================================================

function readCalendar() {

  const calendar =
    getBookingCalendar();

  const now =
    new Date();

  const from =
    startOfDay(now);

  const to =
    new Date(from);

  to.setDate(
    to.getDate() +
    LOOKAHEAD_DAYS +
    1
  );

  const calendarEvents =
    calendar.getEvents(
      from,
      to
    );

  const events = [];
  const freeWindows = [];
  const blocks = [];

  calendarEvents.forEach(
    function(event) {

      const title =
        String(
          event.getTitle() || ''
        ).trim();

      const item = {
        id: event.getId(),
        title: title,
        start:
          event
            .getStartTime()
            .toISOString(),
        end:
          event
            .getEndTime()
            .toISOString(),
        description:
          String(
            event.getDescription() || ''
          ),
        isBooking:
          isBookingCalendarEvent(
            event
          )
      };

      events.push(item);

      if (
        title === FREE_EVENT_TITLE
      ) {

        freeWindows.push({
          id: item.id,
          start: item.start,
          end: item.end
        });

      } else if (
        !item.isBooking
      ) {

        blocks.push({
          id: item.id,
          title: title || 'Занято',
          start: item.start,
          end: item.end
        });
      }
    }
  );

  return {
    events: events,
    freeWindows: freeWindows,
    blocks: blocks
  };
}



// ============================================================
// ПОЛУЧЕНИЕ КАЛЕНДАРЯ
// ============================================================

function getBookingCalendar() {

  const calendars =
    CalendarApp.getCalendarsByName(
      BOOKING_CALENDAR_NAME
    );

  if (
    !calendars ||
    calendars.length === 0
  ) {
    throw new Error(
      'Календарь "' +
      BOOKING_CALENDAR_NAME +
      '" не найден.'
    );
  }

  return calendars[0];
}


// ============================================================
// ПОСТРОЕНИЕ СЛОТОВ
// ============================================================
// СОЗДАНИЕ ОБЪЕКТА СЛОТА
// ============================================================
// КАНДИДАТЫ НАЧАЛА СЛОТОВ
// ============================================================
// ПРОВЕРКА СВОБОДНЫХ ЧАСОВ
// ============================================================

function hasFreeCalendarHours(
  start,
  durationHours,
  freeWindows,
  blocks
) {

  const hours =
    Number(durationHours);

  if (
    !hours ||
    hours <= 0
  ) {
    return false;
  }

  for (
    let i = 0;
    i < hours;
    i++
  ) {

    const hourStart =
      addMinutes(
        start,
        i * 60
      );

    const hourEnd =
      addMinutes(
        hourStart,
        60
      );

    const hasFree =
      freeWindows.some(
        function(window) {

          const freeStart =
            new Date(
              window.start
            );

          const freeEnd =
            new Date(
              window.end
            );

          return (
            freeStart.getTime() <=
              hourStart.getTime() &&
            freeEnd.getTime() >=
              hourEnd.getTime()
          );
        }
      );

    if (!hasFree) {
      return false;
    }

    if (
      intersectsAnyBlock(
        hourStart,
        hourEnd,
        blocks
      )
    ) {
      return false;
    }
  }

  return true;
}


// ============================================================
// ЧТЕНИЕ МЕРОПРИЯТИЙ
// ============================================================

function readEvents(sheet) {

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow < 2
  ) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        10
      )
      .getValues();

  return values
    .map(
      function(row) {

        return {

          activity:
            row[0] || '',

          name:
            row[1] || '',

          price:
            row[2] || '',

          age:
            row[3] || '',

          duration:
            parseDuration(
              row[4]
            ),

          complexity:
            row[5] || '',

          image:
            row[6] || '',

          description:
            row[8] || '',

          amount:
            parseAmount(
              row[9]
            )
        };
      }
    )
    .filter(
      function(event) {
        return (
          event.name !== ''
        );
      }
    );
}


// ============================================================
// ЧТЕНИЕ БРОНЕЙ
// ============================================================

function readBookings(sheet) {

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow < 2
  ) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        10
      )
      .getValues();

  return values
    .map(
      function(row) {

        return {

          id:
            row[0],

          telegramId:
            row[1],

          name:
            row[2],

          format:
            row[3],

          activity:
            row[4],

          date:
            normalizeDateValue(
              row[5]
            ),

          time:
            normalizeTimeValue(
              row[6]
            ),

          tickets:
            Number(
              row[7]
            ) || 0,

          status:
            String(
              row[8] || ''
            )
              .trim()
              .toLowerCase(),

          slotId:
            row[9] || null
        };
      }
    )
    .filter(
      function(booking) {

        return (
          booking.id !== '' &&
          booking.id !== null
        );
      }
    );
}


// ============================================================
// ПОИСК БРОНЕЙ ТОЧНОГО СЛОТА / ЧАСА
// ============================================================

function getBookingsForExactSlot(
  bookings,
  slotStart,
  event
) {

  const targetDate =
    formatDate(
      slotStart
    );

  const targetTime =
    formatTime(
      slotStart
    );

  const targetFormat =
    getFormat(event);

  return bookings.filter(
    function(booking) {

      if (
        booking.status ===
        'cancelled'
      ) {
        return false;
      }

      if (
        normalizeFormat(
          booking.format
        ) !==
        targetFormat
      ) {
        return false;
      }

      if (
        targetFormat === 'МК'
      ) {

        if (
          booking.date !==
          targetDate
        ) {
          return false;
        }

        if (
          booking.time !==
          targetTime
        ) {
          return false;
        }

        return (
          String(
            booking.activity || ''
          ).trim() ===
          String(
            event.name
          ).trim()
        );
      }

      /*
       * Для Диогена бронь считается относящейся
       * к часу, если её интервал пересекает
       * этот час.
       */
      const bookingEvent =
        findEventForBooking(
          booking
        );

      if (
        !bookingEvent
      ) {
        return false;
      }

      if (
        String(
          bookingEvent.name
        ).trim() !==
        String(
          event.name
        ).trim()
      ) {
        return false;
      }

      const interval =
        getBookingInterval(
          booking,
          bookingEvent
        );

      const hourEnd =
        addMinutes(
          slotStart,
          SLOT_STEP_MINUTES
        );

      return (
        interval.start.getTime() <
          hourEnd.getTime() &&
        interval.end.getTime() >
          slotStart.getTime()
      );
    }
  );
}


// ============================================================
// ВСЕ БРОНИ В ТОЙ ЖЕ ТОЧКЕ ВРЕМЕНИ
// ============================================================

function getAllBookingsAtStart(
  bookings,
  slotStart
) {

  const targetDate =
    formatDate(
      slotStart
    );

  const targetTime =
    formatTime(
      slotStart
    );

  return bookings.filter(
    function(booking) {

      return (
        booking.status !==
          'cancelled' &&
        booking.date ===
          targetDate &&
        booking.time ===
          targetTime
      );
    }
  );
}


// ============================================================
// ЗАНЯТОСТЬ
// ============================================================

function calculateOccupancy(
  bookings
) {

  return bookings.reduce(
    function(total, booking) {

      return (
        total +
        (
          Number(
            booking.tickets
          ) || 0
        )
      );
    },
    0
  );
}


// ============================================================
// ФОРМАТ
// ============================================================

function getFormat(event) {

  if (
    String(
      event.activity
    )
      .trim()
      .toLowerCase() ===
    'диоген'
  ) {
    return 'Диоген';
  }

  return 'МК';
}


function normalizeFormat(value) {

  const text =
    String(
      value || ''
    )
      .trim()
      .toLowerCase();

  if (
    text === 'диоген'
  ) {
    return 'Диоген';
  }

  if (
    text === 'мк' ||
    text === 'masterclass'
  ) {
    return 'МК';
  }

  return String(
    value || ''
  ).trim();
}


// ============================================================
// ВМЕСТИМОСТЬ
// ============================================================

function getCapacity(event) {

  return Math.max(
    0,
    Number(
      event.amount
    ) || 0
  );
}


// ============================================================
// МИНИМАЛЬНОЕ КОЛИЧЕСТВО БИЛЕТОВ
// ============================================================

function calculateMinTickets(
  event,
  occupancy,
  capacity
) {

  if (
    occupancy > 0
  ) {
    return 1;
  }

  if (
    getFormat(event) ===
    'Диоген'
  ) {
    return 1;
  }

  return 2;
}


// ============================================================
// СОЗДАНИЕ СОБЫТИЯ БРОНИ В КАЛЕНДАРЕ
// ============================================================

function createCalendarBooking(
  event,
  slotStart,
  bookings
) {

  const calendar =
    getBookingCalendar();

  const durationMinutes =
    Number(
      event.duration
    ) * 60;

  const slotEnd =
    addMinutes(
      slotStart,
      durationMinutes
    );

  const title =
    getFormat(event) === 'МК'
      ? 'МК — ' + event.name
      : 'Диоген — ' + event.name;

  const description =
    buildCalendarBookingDescription(
      event,
      slotStart,
      bookings
    );

  let bookingEvent = null;

  try {

    bookingEvent =
      calendar.createEvent(
        title,
        slotStart,
        slotEnd,
        {
          description:
            description
        }
      );

    synchronizeCalendar();

    return bookingEvent;

  } catch (error) {

    if (bookingEvent) {

      try {

        bookingEvent.deleteEvent();

      } catch (cleanupError) {

        Logger.log(
          'Не удалось удалить календарную бронь после ошибки: ' +
          cleanupError.message
        );
      }
    }

    try {

      synchronizeCalendar();

    } catch (syncError) {

      Logger.log(
        'Не удалось повторно синхронизировать календарь: ' +
        syncError.message
      );
    }

    throw error;
  }
}


// ============================================================
// ОБНОВЛЕНИЕ СУЩЕСТВУЮЩЕГО СОБЫТИЯ БРОНИ
// ============================================================

function updateCalendarBooking(
  event,
  slotStart,
  bookings
) {

  const calendar =
    getBookingCalendar();

  const slotEnd =
    addMinutes(
      slotStart,
      Number(event.duration) * 60
    );

  const slotId =
    makeSlotId(
      event,
      slotStart
    );

  const events =
    calendar.getEvents(
      slotStart,
      slotEnd
    );

  let bookingEvent = null;

  events.forEach(
    function(calendarEvent) {

      if (bookingEvent) {
        return;
      }

      if (
        !isBookingCalendarEvent(
          calendarEvent
        )
      ) {
        return;
      }

      const description =
        String(
          calendarEvent.getDescription() ||
          ''
        );

      if (
        description.indexOf(
          BOOKING_MARKER +
          slotId
        ) === 0
      ) {
        bookingEvent =
          calendarEvent;
      }
    }
  );

  if (!bookingEvent) {

    return createCalendarBooking(
      event,
      slotStart,
      bookings
    );
  }

  const title =
    getFormat(event) === 'МК'
      ? 'МК — ' + event.name
      : 'Диоген — ' + event.name;

  const description =
    buildCalendarBookingDescription(
      event,
      slotStart,
      bookings
    );

  bookingEvent.setTitle(
    title
  );

  bookingEvent.setDescription(
    description
  );

  synchronizeCalendar();

  return bookingEvent;
}


// ============================================================
// ДИОГЕН — СИНХРОНИЗАЦИЯ ЧАСОВ
// ============================================================

function syncDiogenCalendarRange(
  event,
  rangeStart,
  rangeEnd,
  bookings
) {

  const calendar =
    getBookingCalendar();

  let cursor =
    new Date(
      rangeStart
    );

  while (
    cursor.getTime() <
    rangeEnd.getTime()
  ) {

    const hourStart =
      new Date(
        cursor
      );

    const hourEnd =
      addMinutes(
        hourStart,
        SLOT_STEP_MINUTES
      );

    syncDiogenCalendarHour(
      event,
      hourStart,
      hourEnd,
      bookings
    );

    cursor =
      hourEnd;
  }

  synchronizeCalendar();
}


// ============================================================
// ДИОГЕН — СИНХРОНИЗАЦИЯ ОДНОГО ЧАСА
// ============================================================

function syncDiogenCalendarHour(
  event,
  hourStart,
  hourEnd,
  bookings
) {

  const calendar =
    getBookingCalendar();

  const slotId =
    makeSlotId(
      event,
      hourStart
    );

  const events =
    calendar.getEvents(
      hourStart,
      hourEnd
    );

  let bookingEvents = [];

  events.forEach(
    function(calendarEvent) {

      if (
        !isBookingCalendarEvent(
          calendarEvent
        )
      ) {
        return;
      }

      const description =
        String(
          calendarEvent.getDescription() ||
          ''
        );

      if (
        description.indexOf(
          BOOKING_MARKER +
          slotId
        ) === 0
      ) {
        bookingEvents.push(
          calendarEvent
        );
      }
    }
  );

  const hourBookings =
    getBookingsForExactSlot(
      bookings,
      hourStart,
      event
    );

  const occupancy =
    calculateOccupancy(
      hourBookings
    );

  if (
    occupancy <= 0
  ) {

    bookingEvents.forEach(
      function(calendarEvent) {
        calendarEvent.deleteEvent();
      }
    );

    return;
  }

  const title =
    'Диоген — ' +
    event.name;

  const description =
    buildCalendarBookingDescription(
      event,
      hourStart,
      bookings
    );

  let bookingEvent =
    bookingEvents.length
      ? bookingEvents[0]
      : null;

  /*
   * Дубликаты технических событий
   * для одного часа удаляем.
   */
  bookingEvents.slice(1)
    .forEach(
      function(calendarEvent) {
        calendarEvent.deleteEvent();
      }
    );

  if (!bookingEvent) {

    bookingEvent =
      calendar.createEvent(
        title,
        hourStart,
        hourEnd,
        {
          description:
            description
        }
      );

  } else {

    bookingEvent.setTitle(
      title
    );

    bookingEvent.setDescription(
      description
    );
  }
}


// ============================================================
// ФОРМИРОВАНИЕ ОПИСАНИЯ СОБЫТИЯ БРОНИ
// ============================================================

function buildCalendarBookingDescription(
  event,
  slotStart,
  bookings
) {

  const slotId =
    makeSlotId(
      event,
      slotStart
    );

  const slotBookings =
    getBookingsForExactSlot(
      bookings,
      slotStart,
      event
    );

  const capacity =
    getCapacity(event);

  const occupied =
    calculateOccupancy(
      slotBookings
    );

  const lines = [];

  lines.push(
    BOOKING_MARKER +
    slotId
  );

  slotBookings.forEach(
    function(booking) {

      const name =
        String(
          booking.name || ''
        ).trim();

      const tickets =
        Number(
          booking.tickets
        ) || 0;

      lines.push(
        name +
        ' (' +
        tickets +
        ')'
      );
    }
  );

  lines.push(
    occupied +
    '/' +
    capacity
  );

  return lines.join('\\n');
}


// ============================================================
// ОСВОБОЖДЕНИЕ СЛОТА
// ============================================================

function restoreCalendarFreeSlot(
  event,
  slotStart
) {

  const calendar =
    getBookingCalendar();

  const slotEnd =
    addMinutes(
      slotStart,
      Number(event.duration) * 60
    );

  const events =
    calendar.getEvents(
      slotStart,
      slotEnd
    );

  events.forEach(
    function(calendarEvent) {

      if (
        !isBookingCalendarEvent(
          calendarEvent
        )
      ) {
        return;
      }

      const description =
        String(
          calendarEvent.getDescription() ||
          ''
        );

      if (
        description.indexOf(
          BOOKING_MARKER +
          makeSlotId(
            event,
            slotStart
          )
        ) === 0
      ) {
        calendarEvent.deleteEvent();
      }
    }
  );

  let cursor =
    new Date(
      slotStart
    );

  while (
    cursor.getTime() <
    slotEnd.getTime()
  ) {

    const hourEnd =
      addMinutes(
        cursor,
        60
      );

    const existing =
      calendar.getEvents(
        cursor,
        hourEnd
      );

    const alreadyFree =
      existing.some(
        function(calendarEvent) {

          return (
            String(
              calendarEvent.getTitle() ||
              ''
            ).trim() ===
              FREE_EVENT_TITLE &&
            calendarEvent
              .getStartTime()
              .getTime() ===
              cursor.getTime() &&
            calendarEvent
              .getEndTime()
              .getTime() ===
              hourEnd.getTime()
          );
        }
      );

    const hasBlock =
      existing.some(
        function(calendarEvent) {

          const title =
            String(
              calendarEvent.getTitle() ||
              ''
            ).trim();

          if (
            title ===
            FREE_EVENT_TITLE
          ) {
            return false;
          }

          return !isBookingCalendarEvent(
            calendarEvent
          );
        }
      );

    if (
      !alreadyFree &&
      !hasBlock
    ) {

      calendar.createEvent(
        FREE_EVENT_TITLE,
        cursor,
        hourEnd
      );
    }

    cursor =
      hourEnd;
  }
}


// ============================================================
// ТЕХНИЧЕСКОЕ СОБЫТИЕ БРОНИ
// ============================================================

function isBookingCalendarEvent(
  event
) {

  const description =
    String(
      event.getDescription() || ''
    );

  return (
    description.indexOf(
      BOOKING_MARKER
    ) === 0
  );
}


// ============================================================
// ПЕРЕСЕЧЕНИЕ С КАЛЕНДАРНЫМИ СОБЫТИЯМИ
// ============================================================

function intersectsCalendarEvents(
  start,
  end,
  events
) {

  return events.some(
    function(event) {

      const eventStart =
        event.getStartTime();

      const eventEnd =
        event.getEndTime();

      return (
        start.getTime() <
          eventEnd.getTime() &&
        end.getTime() >
          eventStart.getTime()
      );
    }
  );
}


// ============================================================
// ПЕРЕСЕЧЕНИЕ С BLOCKS ИЗ API
// ============================================================

function intersectsAnyBlock(
  start,
  end,
  blocks
) {

  return blocks.some(
    function(block) {

      const blockStart =
        new Date(
          block.start
        );

      const blockEnd =
        new Date(
          block.end
        );

      return (
        start.getTime() <
        blockEnd.getTime() &&
        end.getTime() >
        blockStart.getTime()
      );
    }
  );
}


// ============================================================
// ОБЪЕДИНЕНИЕ СВОБОДНЫХ ОКОН
// ============================================================

function mergeFreeWindows(
  events
) {

  const windows =
    events
      .map(
        function(event) {

          return {
            start:
              event.getStartTime(),

            end:
              event.getEndTime()
          };
        }
      )
      .sort(
        function(a, b) {

          return (
            a.start.getTime() -
            b.start.getTime()
          );
        }
      );

  const result = [];

  windows.forEach(
    function(window) {

      if (
        result.length === 0
      ) {

        result.push({
          start:
            window.start,
          end:
            window.end
        });

        return;
      }

      const last =
        result[
          result.length - 1
        ];

      if (
        window.start.getTime() <=
        last.end.getTime()
      ) {

        if (
          window.end.getTime() >
          last.end.getTime()
        ) {
          last.end =
            window.end;
        }

      } else {

        result.push({
          start:
            window.start,
          end:
            window.end
        });
      }
    }
  );

  return result;
}


// ============================================================
// ОКРУГЛЕНИЕ ВРЕМЕНИ
// ============================================================

function floorToHour(
  date
) {

  const result =
    new Date(
      date
    );

  result.setMinutes(
    0,
    0,
    0
  );

  return result;
}


function ceilToHour(
  date
) {

  const result =
    new Date(
      date
    );

  if (
    result.getMinutes() !== 0 ||
    result.getSeconds() !== 0 ||
    result.getMilliseconds() !== 0
  ) {

    result.setHours(
      result.getHours() + 1
    );
  }

  result.setMinutes(
    0,
    0,
    0
  );

  return result;
}


// ============================================================
// SLOT ID
// ============================================================

function makeSlotId(
  event,
  start
) {

  return [
    getFormat(event),
    event.name || 'Диоген',
    Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmm'
    )
  ].join('_');
}


// ============================================================
// SLOT ID БРОНИ С ИНТЕРВАЛОМ
// ============================================================

function makeBookingSlotId(
  event,
  start,
  end
) {

  return [
    getFormat(event),
    event.name || 'Диоген',
    Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmm'
    ),
    Utilities.formatDate(
      end,
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmm'
    )
  ].join('::');
}


// ============================================================
// РАЗБОР НОВОГО SLOT ID
// ============================================================

function parseBookingSlotId(
  slotId
) {

  const parts =
    String(
      slotId || ''
    ).split('::');

  if (
    parts.length !== 4
  ) {
    return null;
  }

  return {
    format:
      parts[0],

    eventName:
      parts[1],

    start:
      parts[2],

    end:
      parts[3]
  };
}


// ============================================================
// РАЗБОР СТАРОГО SLOT ID
// ============================================================

function parseOldSlotId(
  slotId
) {

  const text =
    String(
      slotId || ''
    );

  const match =
    text.match(
      /^([^_]+)_(.+)_(\d{8}-\d{4})$/
    );

  if (!match) {
    return null;
  }

  return {
    format:
      match[1],

    eventName:
      match[2],

    start:
      match[3]
  };
}


// ============================================================
// ИНТЕРВАЛ БРОНИ
// ============================================================

function getBookingInterval(
  booking,
  event
) {

  /*
   * Новые брони содержат точное окончание
   * прямо в slot_id.
   */
  const parsed =
    parseBookingSlotId(
      booking.slotId
    );

  if (
    parsed &&
    parsed.end
  ) {

    const start =
      parseCompactDateTime(
        parsed.start
      );

    const end =
      parseCompactDateTime(
        parsed.end
      );

    if (
      start &&
      end &&
      end.getTime() >
      start.getTime()
    ) {

      return {
        start: start,
        end: end
      };
    }
  }

  /*
   * Старые брони Диогена и все МК
   * используют длительность из события.
   */
  const start =
    parseBookingDateTime(
      booking.date,
      booking.time
    );

  const end =
    addMinutes(
      start,
      Number(event.duration) * 60
    );

  return {
    start: start,
    end: end
  };
}


// ============================================================
// ПОЧАСОВЫЕ ДИАПАЗОНЫ
// ============================================================

function getHourlyRanges(
  start,
  end
) {

  const ranges = [];

  let cursor =
    new Date(
      start
    );

  while (
    cursor.getTime() <
    end.getTime()
  ) {

    const hourEnd =
      addMinutes(
        cursor,
        SLOT_STEP_MINUTES
      );

    ranges.push({
      start:
        new Date(cursor),

      end:
        new Date(hourEnd)
    });

    cursor =
      hourEnd;
  }

  return ranges;
}


// ============================================================
// ПАРСИНГ КОМПАКТНОЙ ДАТЫ
// ============================================================

function parseCompactDateTime(
  value
) {

  const match =
    String(
      value || ''
    ).match(
      /^(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})$/
    );

  if (!match) {
    return null;
  }

  return parseBookingDateTime(
    match[3] +
    '.' +
    match[2] +
    '.' +
    match[1],
    match[4] +
    ':' +
    match[5]
  );
}


// ============================================================
// ПАРСИНГ ДЛИТЕЛЬНОСТИ
// ============================================================

function parseDuration(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return 0;
  }

  if (
    typeof value === 'number'
  ) {
    return value;
  }

  const text =
    String(value)
      .trim()
      .toLowerCase();

  const numberMatch =
    text.match(
      /(\d+(?:[.,]\d+)?)/
    );

  if (
    !numberMatch
  ) {
    return 0;
  }

  const number =
    Number(
      numberMatch[1]
        .replace(',', '.')
    );

  if (
    text.indexOf('мин') !== -1
  ) {
    return number / 60;
  }

  return number;
}


// ============================================================
// ПАРСИНГ ВМЕСТИМОСТИ
// ============================================================

function parseAmount(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return 0;
  }

  const numberMatch =
    String(value).match(
      /\d+/
    );

  if (
    !numberMatch
  ) {
    return 0;
  }

  return Number(
    numberMatch[0]
  );
}


// ============================================================
// ID НОВОЙ БРОНИ
// ============================================================

function getNextBookingId(
  sheet
) {

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow < 2
  ) {
    return 1;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();

  let maxId = 0;

  values.forEach(
    function(row) {

      const number =
        Number(
          row[0]
        );

      if (
        Number.isFinite(number) &&
        number > maxId
      ) {
        maxId = number;
      }
    }
  );

  return maxId + 1;
}


// ============================================================
// ДАТА/ВРЕМЯ БРОНИ
// ============================================================

function parseBookingDateTime(
  dateString,
  timeString
) {

  const timezone =
    Session.getScriptTimeZone();

  const result =
    Utilities.parseDate(
      String(dateString) +
      ' ' +
      String(timeString),
      timezone,
      'dd.MM.yyyy HH:mm'
    );

  if (
    !result ||
    isNaN(
      result.getTime()
    )
  ) {
    throw new Error(
      'Некорректная дата или время: ' +
      dateString +
      ' ' +
      timeString
    );
  }

  return result;
}


// ============================================================
// ДАТА
// ============================================================

function normalizeDateValue(
  value
) {

  if (
    value instanceof Date &&
    !isNaN(
      value.getTime()
    )
  ) {
    return formatDate(
      value
    );
  }

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(
    value
  ).trim();
}


// ============================================================
// ВРЕМЯ
// ============================================================

function normalizeTimeValue(
  value
) {

  if (
    value instanceof Date &&
    !isNaN(
      value.getTime()
    )
  ) {
    return formatTime(
      value
    );
  }

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(
    value
  ).trim();
}


// ============================================================
// ФОРМАТИРОВАНИЕ ДАТЫ
// ============================================================

function formatDate(
  date
) {

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd.MM.yyyy'
  );
}


// ============================================================
// ФОРМАТИРОВАНИЕ ВРЕМЕНИ
// ============================================================

function formatTime(
  date
) {

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'HH:mm'
  );
}


// ============================================================
// ФОРМАТИРОВАНИЕ ДАТЫ + ВРЕМЕНИ
// ============================================================

function formatDateTime(
  date
) {

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd.MM.yyyy HH:mm'
  );
}


// ============================================================
// НАЧАЛО ДНЯ
// ============================================================

function startOfDay(
  date
) {

  const result =
    new Date(
      date
    );

  result.setHours(
    0,
    0,
    0,
    0
  );

  return result;
}


// ============================================================
// ДОБАВЛЕНИЕ МИНУТ
// ============================================================

function addMinutes(
  date,
  minutes
) {

  return new Date(
    date.getTime() +
    minutes *
    60 *
    1000
  );
}


// ============================================================
// ТЕСТ БРОНИ
// ============================================================

function testBooking() {

  const data = {

    action:
      'book',

    telegramId:
      '99999',

    name:
      'Тест',

    format:
      'МК',

    eventName:
      'Кошелёк',

    date:
      '21.09.2026',

    time:
      '11:00',

    tickets:
      2
  };

  const result =
    createBooking(
      data
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
}
