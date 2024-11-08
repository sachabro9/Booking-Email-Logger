function logBookingEmails() {
  const label = GmailApp.getUserLabelByName("XXXXXXXXXXXXX)");  // Ensure this label matches exactly
  const threads = label.getThreads();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("xxxxxxxx");  // Adjust sheet name to the Google Sheet
  const calendar = CalendarApp.getCalendarById("xxxxxxxxxxxxxx");  // Replace with your Calendar ID

  threads.forEach(thread => {
    const messages = thread.getMessages();
    
    messages.forEach(message => {
      const date = message.getDate();
      const content = message.getPlainBody();
      let bookingSource = "Unknown";
      let bookingDate = "Unknown";
      let price = "Unknown";
      let guestName = "Unknown";
      let contactInfo = "Unknown";

      // Detect email type and extract data accordingly
      
      // Type 1: Airbnb
      if (content.includes("Recibiste una nueva reservación") || content.includes("Código de confirmación")) {
        bookingSource = "Airbnb";
        
        // Extract Guest Name
        const nameMatch = content.match(/(?:Huéspedes nuevos|Huéspedes confirmados)\s+([\w\s]+)/i);
        guestName = nameMatch ? nameMatch[1].trim() : "Unknown";

        // Extract Booking Date
        const dateMatch = content.match(/\d{1,2} [a-zA-Z]{3}\., \d{2}:\d{2}/);
        bookingDate = dateMatch ? dateMatch[0] : "Unknown";

        // Extract Total Price (last dollar amount in USD)
        const priceMatches = content.match(/\$(\d+)\s*USD/g);
        price = priceMatches && priceMatches.length > 1 ? priceMatches[priceMatches.length - 1] : "Unknown";

        // Extract Confirmation Code
        const confirmationMatch = content.match(/Código de confirmación\s*(\w+)/);
        contactInfo = confirmationMatch ? confirmationMatch[1] : "Unknown";
      }

      // Type 2: WeTravel
      if (content.includes("WeTravel Team") || content.includes("Receipt Number")) {
        bookingSource = "WeTravel";

        // Extract Guest Name
        const nameMatch = content.match(/Participant:\s*([\w\s]+)/);
        guestName = nameMatch ? nameMatch[1].trim() : "Unknown";

        // Extract Booking Date
        const dateMatch = content.match(/Trip Date:\s*([\w\s,]+)/);
        bookingDate = dateMatch ? dateMatch[1].replace(/\s+/g, ' ').trim() : "Unknown";  // Removes line breaks

        // Extract Price after "Total"
        const allPrices = [...content.matchAll(/\$(\d+(\.\d{2})?)/g)];
        const totalIndex = content.indexOf("Total");

        price = "Unknown";
        for (const match of allPrices) {
          const index = match.index;
          if (index > totalIndex) {
            price = "$" + match[1];
            break;
          }
        }

        // Extract Contact Information (email)
        const emailMatch = content.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
        contactInfo = emailMatch ? emailMatch[0] : "Unknown";
      }

      // Type 3: Bókun Notifications for GetYourGuide or Viator
      else if (content.includes("Bókun Notifications") || content.includes("Booking ref.")) {
        // Detect booking channel by looking for "Sold by" or "Booking channel"
        const channelMatch = content.match(/(?:Sold by|Booking channel)\s*([^\n]+)/);
        bookingSource = channelMatch ? channelMatch[1].trim() : "Unknown Channel";

        // Extract Guest Name
        const nameMatch = content.match(/Customer\s*([^\n]+)/);
        guestName = nameMatch ? nameMatch[1].trim() : "Unknown";

        // Extract Booking Date
        const dateMatch = content.match(/Date\s*([^\n]+)/);
        bookingDate = dateMatch ? dateMatch[1].trim() : "Unknown";

        // Extract Price for Viator (looks for "Viator amount: USD")
        const priceMatch = content.match(/Viator amount:\s*USD\s*(\d+(\.\d{2})?)/i);
        price = priceMatch ? "$" + priceMatch[1] : "Unknown";

        // Extract Contact Information (prioritizes phone if available, otherwise email)
        const contactPhoneMatch = content.match(/Customer phone\s*([^\n]+)/);
        const contactEmailMatch = content.match(/Customer email:\s*([^\n]+)/);
        contactInfo = contactPhoneMatch ? contactPhoneMatch[1].trim() : (contactEmailMatch ? contactEmailMatch[1].trim() : "Unknown");
      }

      // Parse and insert event into Google Calendar if booking date is valid
      if (bookingDate !== "Unknown") {
        try {
          const eventDate = new Date(bookingDate);
          const event = calendar.createEvent(
            `Booking from ${bookingSource}`,
            eventDate,
            new Date(eventDate.getTime() + 3 * 60 * 60 * 1000),  // Default duration 3 hours
            {
              description: `Guest: ${guestName}\nContact: ${contactInfo}\nPrice: ${price}\nSource: ${bookingSource}`,
            }
          );
          Logger.log(`Event created: ${event.getTitle()} on ${event.getStartTime()}`);
        } catch (error) {
          Logger.log("Error creating event: " + error.message);
        }
      }

      // Log extracted data to Google Sheets
      sheet.appendRow([date, bookingSource, bookingDate, price, guestName, contactInfo]);
    });
    
    // Remove the label to avoid reprocessing
    thread.removeLabel(label);
  });
}
