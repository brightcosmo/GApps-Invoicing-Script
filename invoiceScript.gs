function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom')
    .addItem('Generate Invoice', 'exportSelectedRowToPDF')
    .addItem('Send Latest Invoice', 'sendLatestInvoiceByEmail')
    .addToUi();
}

function exportSelectedRowToPDF() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Company Details");
  const companyInfo = getCompanyInfo(configSheet);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  if (row <= 2) {
    SpreadsheetApp.getUi().alert('Please select a row other than the header row.');
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, 20).getValues()[0];
  const invoiceData = parseRowData(rowData);

  if (invoiceData.services.length === 0) {
    SpreadsheetApp.getUi().alert('No services found for this invoice.');
    return;
  }

  const doc = DocumentApp.create(`Invoice-${invoiceData.jobID}`);
  const body = doc.getBody();
  setDocumentMargins(body);

  const logoError = insertHeader(body, companyInfo);
  insertInvoiceDetails(body, invoiceData, companyInfo);

  const { pdfUrl, version } = finalizeAndSharePDF(doc, invoiceData.jobID);

  showPdfDownloadDialog(pdfUrl, version, logoError);

  DriveApp.getFileById(doc.getId()).setTrashed(true);
}

function sendLatestInvoiceByEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  if (row <= 2) {
    SpreadsheetApp.getUi().alert('Please select a row other than the header row.');
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, 20).getValues()[0];
  const invoiceData = parseRowData(rowData);

  if (!invoiceData.services.length) {
    SpreadsheetApp.getUi().alert('No services found for this invoice.');
    return;
  }

  const pdfUrl = getLatestInvoicePDFUrl(invoiceData.jobID);
  if (!pdfUrl) {
    SpreadsheetApp.getUi().alert('No invoice found for the given job ID.');
    return;
  }

  const recipientEmail = rowData[3];
  const subject = `Invoice ${invoiceData.jobID} - Latest Version`;
  const message = `Dear ${invoiceData.billingName},\n\nPlease find attached the latest version of invoice ${invoiceData.jobID}.\n\nThank you for your business!\n\nBest regards,\n${getCompanyInfo(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Company Details")).name}`;
  
  sendEmailWithAttachment(recipientEmail, subject, message, pdfUrl);
}

function getLatestInvoicePDFUrl(jobID) {
  const folder = getOrCreateFolder("Invoices");
  const files = folder.getFiles();
  let latestVersion = 0;
  let latestFile = null;

  const pattern = new RegExp(`Invoice-${jobID}_V(\\d{2})\\.pdf`);

  while (files.hasNext()) {
    const file = files.next();
    const match = file.getName().match(pattern);
    if (match) {
      const version = parseInt(match[1], 10);
      if (version > latestVersion) {
        latestVersion = version;
        latestFile = file;
      }
    }
  }

  return latestFile ? latestFile.getUrl() : null;
}

function sendEmailWithAttachment(recipientEmail, subject, message, pdfUrl) {
  try {
    const pdfFile = UrlFetchApp.fetch(pdfUrl).getBlob();
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      body: message,
      attachments: [pdfFile]
    });
    Logger.log(`Email sent to ${recipientEmail} with attachment.`);
  } catch (e) {
    Logger.log('Error sending email: ' + e.message);
  }
}

function finalizeAndSharePDF(doc, jobID) {
  if (doc) {
    doc.saveAndClose();
    const pdfBlob = doc.getAs('application/pdf');
    const folder = getOrCreateFolder("Invoices");

    let version = 1;
    let pdfFileName = `Invoice-${jobID}_V${String(version).padStart(2, '0')}.pdf`;
    while (folder.getFilesByName(pdfFileName).hasNext()) {
      version++;
      pdfFileName = `Invoice-${jobID}_V${String(version).padStart(2, '0')}.pdf`;
    }

    const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { pdfUrl: pdfFile.getUrl(), version: version };
  } else {
    const folder = getOrCreateFolder("Invoices");
    const files = folder.getFilesByName(`Invoice-${jobID}_V01.pdf`);
    if (files.hasNext()) {
      const pdfFile = files.next();
      return { pdfUrl: pdfFile.getUrl() };
    } else {
      throw new Error('No invoice found for the given job ID.');
    }
  }
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function showPdfDownloadDialog(pdfUrl, version, logoError) {
  const htmlheight = logoError ? 200 : 150;
  const htmlContent = `
    <html>
      <body>
        <p>Invoice PDF generated successfully.</p>
        <p>Version: ${version}.</p>
        <p><a href="${pdfUrl}" target="_blank" rel="noopener noreferrer">Click here to view and download your Invoice PDF</a>.</p>
        ${logoError ? `<p>Note: The logo image could not be included due to an error with the file ID.</p>` : ''}
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
                                .setWidth(300)
                                .setHeight(htmlheight);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Invoice PDF Download');
}

// Helper functions
function setDocumentMargins(body) {
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(72);
  body.setMarginRight(72);
}

function insertHeader(body, companyInfo) {
  const headerTable = body.appendTable();
  const row = headerTable.appendTableRow();
  
  const logoCell = row.appendTableCell();
  const logoError = !insertLogo(logoCell, companyInfo.logoFileId);

  const headerCell = row.appendTableCell();
  headerCell.appendParagraph(companyInfo.name).setFontFamily('PT Serif').setFontSize(24).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  headerCell.appendParagraph(companyInfo.address).setFontFamily('PT Serif').setFontSize(12).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  headerCell.appendParagraph(companyInfo.website).setFontFamily('PT Serif').setFontSize(12).setAlignment(DocumentApp.HorizontalAlignment.LEFT);

  headerTable.setBorderColor("#FFFFFF");
  headerTable.setColumnWidth(0, 100);

  return logoError;
}

function insertLogo(cell, logoFileId) {
  try {
    const logo = DriveApp.getFileById(logoFileId).getBlob();
    const maxWidth = 100;

    cell.setWidth(maxWidth);
    const logoParagraph = cell.appendParagraph('');
    const logoImage = logoParagraph.appendInlineImage(logo);

    const aspectRatio = logoImage.getHeight() / logoImage.getWidth();
    logoImage.setWidth(maxWidth);
    logoImage.setHeight(maxWidth * aspectRatio);

    logoParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    return true;
  } catch (e) {
    Logger.log('Error fetching logo: ' + e.message);
    return false;
  }
}

function insertInvoiceDetails(body, invoiceData, companyInfo) {
  const { jobID, billingName, billingAddress, subtotal, deposit, totalDue, services, discount } = invoiceData;
  
  const today = new Date();
  const dueDate = new Date(today.getTime() + (14 * 24 * 60 * 60 * 1000)); // 14 days after current date

  body.appendParagraph(`Invoice #: ${jobID}`).setFontFamily('Helvetica Neue').setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setFontFamily('Helvetica Neue').setLineSpacing(1.15);
  body.appendParagraph(`Invoice Date: ${formatDate(today)}`).setFontFamily('Helvetica Neue').setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph(`Due Date: ${formatDate(dueDate)}`).setFontFamily('Helvetica Neue').setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph("");

  body.appendParagraph("BILL TO:").setFontFamily('Helvetica Neue').setFontSize(10).setBold(true);
  body.appendParagraph(billingName).setFontFamily('Helvetica Neue').setFontSize(10).setBold(false);
  body.appendParagraph(billingAddress).setFontFamily('Helvetica Neue').setFontSize(10);
  body.appendParagraph("");

  const table = body.appendTable();
  const headerRow = table.appendTableRow();
  const headerCellService = headerRow.appendTableCell('Service Description').setForegroundColor("#FFFFFF").setBackgroundColor('#000000').setBold(true).setFontFamily('Helvetica Neue').setFontSize(10);
  headerCellService.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const headerCellRate = headerRow.appendTableCell('Rate').setBackgroundColor('#C3C3C3').setBackgroundColor('#000000').setBold(true).setFontFamily('Helvetica Neue').setFontSize(10);
  headerCellRate.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const headerCellQuantity = headerRow.appendTableCell('Qty').setBackgroundColor('#C3C3C3').setBackgroundColor('#000000').setBold(true).setFontFamily('Helvetica Neue').setFontSize(10);
  headerCellQuantity.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const headerCellTotal = headerRow.appendTableCell('Amount').setBackgroundColor('#C3C3C3').setBackgroundColor('#000000').setBold(true).setFontFamily('Helvetica Neue').setFontSize(10);
  headerCellTotal.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  table.setColumnWidth(0, 280);
  table.setColumnWidth(1, 70);
  table.setColumnWidth(2, 40);
  table.setColumnWidth(3, 70);

  services.forEach(service => {
    const row = table.appendTableRow();
    const serviceDescription = row.appendTableCell(service.name).setFontFamily('Helvetica Neue').setFontSize(10).setBold(false).setForegroundColor("#000000");
    serviceDescription.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    const serviceRate = row.appendTableCell(formatAmount(service.rate)).setFontFamily('Helvetica Neue').setFontSize(10);
    serviceRate.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

    const serviceQuantity = row.appendTableCell(service.quantity.toString()).setFontFamily('Helvetica Neue').setFontSize(10)
    serviceQuantity.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

    const serviceHeader = row.appendTableCell(formatAmount(service.total)).setFontFamily('Helvetica Neue').setFontSize(10)
    serviceHeader.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  });

  const summaryTable = body.appendTable();
  summaryTable.setBorderWidth(0);
  const summaryData = [
    { label: 'Subtotal:', value: formatCurrency(subtotal), bold: false },
    { label: 'Discount:', value: discount > 0 ? `-${formatCurrency(discount)}` : '', bold: false },
    { label: 'Deposit:', value: `-${formatCurrency(deposit)}`, bold: false },
    { label: 'Total Due:', value: formatCurrency(totalDue), bold: true }
  ];
  summaryData.forEach(item => {
    if (item.value) {
      const row = summaryTable.appendTableRow();

      const label = row.appendTableCell(item.label).setFontFamily('Helvetica Neue').setFontSize(10).setBold(item.bold);
      label.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

      const value = row.appendTableCell(item.value).setFontFamily('Helvetica Neue').setFontSize(10).setBold(item.bold);
      value.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }
  });
  summaryTable.setColumnWidth(0, 400)
  summaryTable.setBorderColor("#FFFFFF");


  body.appendParagraph("For bank transfer, pay to:").setFontFamily('Helvetica Neue').setFontSize(10).setBold(true);
  body.appendParagraph(`Bank Name: ${companyInfo.bankName}`).setFontFamily('Helvetica Neue').setFontSize(10).setBold(false);
  body.appendParagraph(`Account Number: ${companyInfo.accountNumber}`).setFontFamily('Helvetica Neue').setFontSize(10);
  body.appendParagraph(companyInfo.bankExtraInfo).setFontFamily('Helvetica Neue').setFontSize(10);
  body.appendParagraph("");

  body.appendParagraph("To pay via cheque, please send to:").setBold(true).setFontFamily('Helvetica Neue').setFontSize(10);
  body.appendParagraph(companyInfo.name).setFontFamily('Helvetica Neue').setFontSize(10).setBold(false);
  body.appendParagraph(companyInfo.address).setFontFamily('Helvetica Neue').setFontSize(10);
  body.appendParagraph(companyInfo.chequeExtraInfo).setFontFamily('Helvetica Neue').setFontSize(10);
}

function formatCurrency(amount) {
  const num = Number(amount);
  if (isNaN(num)) return 'RM0.00';
  return Utilities.formatString('RM%.2f', num);
}

function formatAmount(amount) {
  const num = Number(amount);
  if (isNaN(num)) return '0.00';
  return Utilities.formatString('%.2f', num);
}

function parseRowData(rowData) {
  const [
    jobID, billingName, billingAddress, email, depositAmountInvoiced, remainderDue,
    service1Listed, service1Fee, service1Quantity, 
    service2Listed, service2Fee, service2Quantity, 
    service3Listed, service3Fee, service3Quantity, 
    service4Listed, service4Fee, service4Quantity, 
    service5Listed, service5Fee, service5Quantity,
  ] = rowData;

  const services = [];
  let discount = 0;
  for (let i = 0; i < 5; i++) {
    const listed = [service1Listed, service2Listed, service3Listed, service4Listed, service5Listed][i] || '';
    let fee = parseFloat([service1Fee, service2Fee, service3Fee, service4Fee, service5Fee][i]) || 0;
    let quantity = parseInt([service1Quantity, service2Quantity, service3Quantity, service4Quantity, service5Quantity][i], 10) || (listed.trim() ? 1 : 0);

    if (listed.trim() !== '') {
      const total = fee * quantity;
      if (total < 0) {
        discount += Math.abs(total);
      } else {
        services.push({ name: listed, rate: fee, quantity, total });
      }
    }
  }

  const subtotal = services.reduce((acc, service) => acc + service.total, 0);
  const deposit = parseFloat(depositAmountInvoiced) || 0;
  const totalDue = subtotal - deposit;

  return { jobID, billingName, billingAddress, subtotal, deposit, totalDue, services, discount, email };
}

function getCompanyInfo(sheet) {
  const values = sheet.getRange(1, 3, 6).getValues();
  return {
    name: values[0][0],
    address: values[1][0],
    website: values[2][0],
    logoFileId: extractID(values[3][0]),
    bankName: values[4][0],
    accountNumber: values[5][0],
    bankExtraInfo: "Please include the invoice number with your payment.",
    chequeExtraInfo: "Please include the invoice number on your check."
  };
}

function extractID(url) {
  const regex = /\/d\/([a-zA-Z0-9_-]+)(?:\/|$)/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

function formatDate(date) {
  const day = date.getDate();
  const daySuffix = getDaySuffix(day);
  const month = date.toLocaleString('default', { month: 'long' });
  const year = date.getFullYear();
  return `${day}${daySuffix} ${month} ${year}`;
}

function getDaySuffix(day) {
  if (day > 3 && day < 21) return 'th';
  switch (day % 10) {
    case 1: return 'st';
    case 2: return 'nd';
    case 3: return 'rd';
    default: return 'th';
  }
}
