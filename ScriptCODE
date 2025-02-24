function GET_WHOIS(domain) {
  var apiKey =; 
  var url = `https://www.whoisxmlapi.com/whoisserver/WhoisService?apiKey=${apiKey}&domainName=${domain}&outputFormat=json`;

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var cell = sheet.getActiveRange();
    var row = cell.getRow();
    
    var headers = [
      "Domain Name", "WHOIS Record Date", "Updated Date", "Creation Date", "Registrar", 
      "Registrant Name", "Registrant Organization", "Registrant Address", 
      "Registrant Phone", "Registrant Fax", "Registrant Email", 
      "Registrant Admin", "Registrant Tech", "Name Servers"
    ];

    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(response.getContentText());

    if (data.ErrorMessage) {
      return [["Error", data.ErrorMessage.msg]];
    }

    var whois = data.WhoisRecord;
    
    var values = [
      whois.domainName || "N/A",
      whois.audit?.auditUpdatedDate || "N/A",
      whois.updatedDate || "N/A",
      whois.createdDate || "N/A",
      whois.registrarName || "N/A",
      whois.registrant?.name || "N/A",
      whois.registrant?.organization || "N/A",
      (whois.registrant?.street || "N/A") + ", " + 
      (whois.registrant?.city || "N/A") + ", " + 
      (whois.registrant?.state || "N/A") + ", " + 
      (whois.registrant?.postalCode || "N/A") + ", " + 
      (whois.registrant?.country || "N/A"),
      whois.registrant?.telephone || "N/A",
      whois.registrant?.fax || "N/A",
      whois.registrant?.email || "N/A",
      whois.adminContact?.name || "N/A",
      whois.technicalContact?.name || "N/A",
      (whois.nameServers?.hostNames || ["N/A"]).join(", ")
    ];

    // If function is in the first row of its column, add headers
    return row === 1 ? [headers, values] : [values];

  } catch (e) {
    return [["Error", "Error fetching WHOIS data."]];
  }
}
