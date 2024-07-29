# Google Apps Script Invoice Generator

A simple and efficient Google Apps Script solution that automates the generation of PDF invoices from Google Sheets. Designed to help small and medium-sized businesses in Malaysia comply with eInvoicing regulations.

## Features

- **Automatic PDF Generation**: Creates professional invoices based on data from a Google Sheet.
- **Google Drive Integration**: Saves generated invoices directly to the user's Google Drive.
- **Email Functionality**: Sends invoices to specified email addresses.
- **Customizable Templates**: Personalize invoice templates to match your business branding.
- **User-Friendly Interface**: Easy to set up and use, with minimal technical expertise required.

## Getting Started

### Prerequisites

- A Google account with access to Google Sheets and Google Drive.

### Setup

1. **Open Google Sheets**:
   - Download the template and upload it to your Google Drive 

2. **Access Apps Script**:
   - Go to `Extensions` > `Apps Script`.

3. **Copy the Script**:
   - Replace any existing code with the script from this repository.

4. **Authorize the Script**:
   - Run the script and authorize it to access your Google Drive and send emails on your behalf.

5. **Customize the Template**:
   - Modify the second sheet called `Company Details` based on the issuing company's details.
   - To add your company logo, include its google drive link (with public viewing access).

6. **Run the Script**:
   - Execute the script. Ensure this step is repeated any time you modify the company details.

7. **Generate and send Invoices**
   - Select the row you wish to generate an invoice for.
   - Click on the dropdown called "custom" and select `Generate Invoice` followed by `Send Latest Invoice` if required.
   - Done!

## Script Overview

The script performs the following actions:
- Reads invoice data from a specified Google Sheet.
- Formats the data into a PDF invoice.
- Saves the PDF invoice to Google Drive.
- Sends the invoice to the email address listed in the sheet.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your proposed changes.

## License

This project is licensed under the MIT License.
