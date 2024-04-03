function createSubscriptionEvents() {
  const calendarId = 'REPLACE WITH YOUR ID';
  console.log('Starting to process subscription events...');

  const calendar = CalendarApp.getCalendarById(calendarId);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SUBS");
  const dataRange = sheet.getRange("A21:L" + sheet.getLastRow());
  const subscriptions = dataRange.getValues();

  // Clear events up to 1 year into the future
  const today = new Date();
  const endDateForClearing = new Date(today.getFullYear() + 1, today.getMonth(), today.getDate());
  console.log(`Attempting to clear events up to: ${endDateForClearing.toISOString()}`);

  const eventsToClear = calendar.getEvents(today, endDateForClearing);
  console.log(`Found ${eventsToClear.length} events to clear.`);

  eventsToClear.forEach((event, index) => {
    try {
      console.log(`Attempting to delete event: ${event.getTitle()} on ${event.getStartTime().toISOString()}`);
      event.deleteEvent();
      console.log(`Successfully deleted event #${index + 1}`);
    } catch (e) {
      console.error(`Failed to delete event #${index + 1}: ${e.message}`);
    }
  });

  console.log(`Total events cleared: ${eventsToClear.length}`);

  // Process each subscription
  subscriptions.forEach(([ , name, signUpDateString, , initialDateString, interval, nextRenewalDateString, price, , , , optionalEndDateString], index) => {
    console.log(`Processing subscription ${index + 1}: ${name}`);

    const initialDate = new Date(initialDateString);
    const optionalEndDate = optionalEndDateString ? new Date(optionalEndDateString) : null;

    // Create an initial subscription event
    if (!isNaN(initialDate.getTime())) {
      const description = `INITIAL SUBSCRIPTION START FOR ${name} AT $${price}`;
      console.log(`Creating initial event for: ${name} on ${initialDate}`);
      calendar.createAllDayEvent(`${name} SUBSCRIPTION STARTED`, initialDate, {description});
    } else {
      console.log(`Skipped creating initial event for ${name} due to invalid date.`);
    }

    // Handling optional cancellation date
    if (optionalEndDate && !isNaN(optionalEndDate.getTime())) {
      const description = `END OF SUBSCRIPTION FOR ${name}`;
      console.log(`Creating end of subscription event for: ${name} on ${optionalEndDate}`);
      calendar.createAllDayEvent(`CANCEL ${name} / SUBSCRIPTION ENDED`, optionalEndDate, {description});
    }

    // Schedule renewal events as all-day events within a 1-year limit
    if (interval.toLowerCase() && nextRenewalDateString) {
      let nextRenewalDate = new Date(nextRenewalDateString);
      const description = `SUBSCRIPTION RENEWAL FOR ${name} AT $${price}`;

      while (nextRenewalDate <= endDateForClearing && (!optionalEndDate || nextRenewalDate <= optionalEndDate)) {
        console.log(`Creating renewal event for: ${name} on ${nextRenewalDate}`);
        calendar.createAllDayEvent(`${name} SUBSCRIPTION RENEWAL`, nextRenewalDate, {description});
        
        // Increment nextRenewalDate by the specified interval
        switch (interval.toLowerCase()) {
          case 'daily':
            nextRenewalDate.setDate(nextRenewalDate.getDate() + 1);
            break;
          case 'weekly':
            nextRenewalDate.setDate(nextRenewalDate.getDate() + 7);
            break;
          case 'monthly':
            nextRenewalDate.setMonth(nextRenewalDate.getMonth() + 1);
            break;
          case 'yearly':
            nextRenewalDate.setFullYear(nextRenewalDate.getFullYear() + 1);
            break;
        }
      }
    }
  });

  console.log('Finished processing subscription events.');
}
