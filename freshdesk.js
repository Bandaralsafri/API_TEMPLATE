const axios = require('axios');     // Required for making HTTP requests
const ExcelJS = require('exceljs'); // To generate excel file

// Replace 'YOUR_DOMAIN', 'YOUR_API_KEY', and other placeholders with your actual Freshdesk information
const freshdeskConfig = {
  domain: 'YOUR_DOMAIN',
  apiKey: 'YOUR_API_KEY',
  endpoint: 'YOUR_ENDPOINT_URL',
};

// Map of status values to corresponding status labels
const statusMapping = {
    2: 'Open',
    3: 'Pending',
    4: 'Resolved',
    5: 'Closed',
    // 6: 'Waiting on Customer',    CUSTOM STATUS
    // 7: 'Waiting on Third Party', CUSTOM STATUS
    // 8: 'To Schedule'             CUSTOM STATUS
  };
  
  async function getTickets() {
    try {
      let allTickets = [];
      let nextPage = 1; // START FROM PAGE #1 AND LOOP UNTIL PAGE 100
  
      // Fetch all pages of tickets
      while (nextPage) {
        const response = await axios.get(`${freshdeskConfig.endpoint}/tickets`, {
          params: { page: nextPage,per_page: 100, },
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Basic ${Buffer.from(`${freshdeskConfig.apiKey}:X`).toString('base64')}`,  // Make sure the data getting from
                                                                                                      // the end point is using encoding UTF-8 
          },
        });
  
        // Append tickets from the current page to the array
        allTickets = allTickets.concat(response.data);
  
        // Check if there are more pages
        nextPage = response.headers['x-page'] < response.headers['x-pages'] ? response.headers['x-page'] + 1 : null;
      }
  
      // Create a new Excel workbook and add a worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Tickets');
  
      // Define the headers for the Excel file
      worksheet.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Subject', key: 'subject', width: 30 },
        { header: 'Status', key: 'status', width: 20 },
        //{ header: 'Resolution', key: 'Resolution', width: 20 }, CREATE A COLUMN NAMED 'RESOLUSTION' IN THE GENERATED EXCEL FILE
        
      ];
  
      // Add ticket data to the worksheet
      allTickets.forEach((ticket) => {
        // Check if the status is 'Closed' and set the 'Resolution' column accordingly
        
        //const Resolution =  ticket.custom_fields.cf_resolution ; GET THE DATA FROM A CUSTOM FIELD INSIDE FRESHDESK
        console.log(ticket);
        // Map status value to status label
        const status1 = statusMapping[ticket.status] || ticket.status  ;
        


        // Add a new row to the worksheet
        worksheet.addRow({
            id: ticket.id,
            //Agency_Code: Agency_Code, ADD THE VALUES FROM FRESHDESK TO EXCEL FILE
            subject: ticket.subject,
            Resolution: Resolution,
            status: status1,
        });
      });
  
      // Save the workbook to a file with name:
      await workbook.xlsx.writeFile('tickets.xlsx');
  
      console.log('Tickets data written to tickets.xlsx');
    } catch (error) {
      // Log any errors
      console.error('Error:', error.message);
    }
  }
  
  // Call the function to get tickets and write to Excel
  getTickets();