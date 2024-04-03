function createSubscriptionEvents() {
  const calendarId = 'REPLACE';
  const calendar = CalendarApp.getCalendarById(calendarId);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SUBS");
  const dataRange = sheet.getRange("A21:L" + sheet.getLastRow());
  const subscriptions = dataRange.getValues();

  // Clear all events in the calendar before adding new ones
  const events = calendar.getEvents(new Date(2000, 0, 1), new Date(2100, 0, 1)); // Broad range to capture all events
  events.forEach(event => event.deleteEvent());

  subscriptions.forEach(subscription => {
    const [
      , // Link to site, skipping for event creation, could be included in description
      name,
      signUpDateString,
      freePeriod,
      initialDateString,
      interval,
      nextRenewalDateString,
      price,
      , // Yearly rate, skipping for event creation, could be included in description
      , // Amount spent to date, skipping for event creation, could be included in description
      , // The previously unmentioned blank column
      optionalEndDateString, // Adjusted to extract from column L
    ] = subscription;

    const initialDate = new Date(initialDateString);
    const optionalEndDate = optionalEndDateString ? new Date(optionalEndDateString) : null;

    if (!isNaN(initialDate.getTime())) {
      let description = `INITIAL SUBSCRIPTION START FOR ${name} AT $${price}`;
      calendar.createAllDayEvent(`${name} SUBSCRIPTION STARTED`, initialDate, {description});
    }

    if (optionalEndDate && !isNaN(optionalEndDate.getTime())) {
      let description = `END OF SUBSCRIPTION FOR ${name}`;
      calendar.createAllDayEvent(`${name} SUBSCRIPTION ENDED`, optionalEndDate, {description});
    }

    if (interval.toLowerCase() && nextRenewalDateString) {
      let nextRenewalDate = new Date(nextRenewalDateString);
      let recurrence = CalendarApp.newRecurrence();
      let rule;

      switch (interval.toLowerCase()) {
        case 'daily':
          rule = recurrence.addDailyRule();
          break;
        case 'weekly':
          rule = recurrence.addWeeklyRule();
          break;
        case 'monthly':
          rule = recurrence.addMonthlyRule();
          break;
        case 'yearly':
          rule = recurrence.addYearlyRule();
          break;
      }

      if (optionalEndDate) {
        rule.until(optionalEndDate);
      }

      let eventSeriesDescription = `Subscription renewal for ${name}. Price: ${price}.`;
      calendar.createEventSeries(
        `${name} Subscription Renewal`,
        nextRenewalDate,
        new Date(nextRenewalDate.getTime() + 3600000), // Adding 1 hour to nextRenewalDate for the event end time
        rule,
        {description: eventSeriesDescription}
      );
    }
  });
}
