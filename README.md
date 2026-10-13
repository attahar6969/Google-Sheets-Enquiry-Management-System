# Google Sheets Business enquiry tracking and Sales Management system

A business enquiry tracking and sales management system built using **Google Apps Script and Google Sheets**. It helps businesses manage customer enquiries, track follow-ups, upload bills, monitor sales performance, and automate routine sales operations.

## 🚀 Features

### 1. Enquiry Management
- Create and manage customer enquiries.
- Capture customer contact details, enquiry categories, sources, and representatives.
- Track enquiry status, quotation amount, priority, and follow-up dates.
- Validate form inputs and required fields.

### 2. Follow-Up Management
- Update existing enquiries and maintain follow-up history.
- Track enquiry status changes, remarks, and expected closing dates.
- Manage won, lost, and closed enquiries.
- Maintain customer enquiry records.

### 3. Bill Upload and Verification
- Record bill numbers and bill amounts.
- Upload bill documents to Google Drive.
- Configure file-size limits and sharing permissions.
- Optionally require bill verification before marking an enquiry as Won.

### 4. Interactive Sales Dashboard
- View key performance indicators (KPIs).
- Track total, open, won, lost, and closed enquiries.
- Monitor quotation value, won sales, and open pipeline value.
- Visualize monthly enquiries and sales trends.
- Compare representative-wise sales performance.
- Explore a sales calendar and recent enquiry summaries.
- Filter dashboard data by date, representative, status, category, and source.

### 5. Administration and Settings
- Manage dropdown lists for categories, sources, representatives, statuses, and priorities.
- Configure required form fields and status rules.
- Manage authorized administrators.
- Configure the Web App connection and spreadsheet protections.
- Refresh dashboard and system data.

## 🛠️ Technologies Used

- Google Apps Script
- Google Sheets
- Google Drive
- HTML5
- CSS3
- JavaScript
- Google Charts

## 📁 Project Structure

| File | Purpose |
|---|---|
| `Code.gs` | Main backend and enquiry management logic |
| `NewEnquiry.html` | New enquiry form |
| `FollowUp.html` | Follow-up update form |
| `UploadBill.html` | Bill upload interface |
| `Styles.html` | Shared form styling |
| `Dashboard.gs` | Dashboard data processing and sheet generation |
| `Dashboard.html` | Interactive dashboard interface |
| `SettingsCode.gs` | Settings, administration, and configuration logic |
| `Settings.html` | Administration settings interface |
| `README.md` | Project documentation |

*The files listed above describe the intended project structure; ensure each file exists in the repository before relying on this list.*

## ⚙️ Setup and Installation

1. Create a Google Sheet to store enquiry records and related data.
2. Open **Extensions → Apps Script**.
3. Add the project's `.gs` and `.html` files to the Apps Script project, keeping their names consistent with the code.
4. Configure the required sheets, headers, dropdown lists, and script properties.
5. Review authorization and administrator access settings.
6. Deploy a Google Apps Script Web App if required by the project configuration.
7. Test enquiry submission, follow-up updates, bill uploads, dashboard filters, and settings.
8. Configure spreadsheet protections and verify access permissions before sharing the system with users.

**Note:** The exact sheet structure, script properties, and deployment permissions must match the implementation in the source code.

## 🔐 Security and Privacy

- Restrict administrative functions to authorized users.
- Protect sensitive spreadsheet data and uploaded documents.
- Do not commit API keys, private Web App URLs, customer information, or other confidential data to GitHub.
- Review Google Drive sharing permissions before uploading or sharing bills.
- Test access controls before deploying the system for business use.

## 🎯 Project Objective

The objective of this project is to simplify sales enquiry management by combining Google Sheets as a data store with Google Apps Script automation and an interactive dashboard. It aims to improve record keeping, follow-up visibility, sales reporting, and administrative efficiency.

## 📌 Project Status

Developed as a Google Apps Script and Google Sheets business management project. Features and deployment requirements should be verified against the latest source code.

## 👨‍💻 Author

**Attahar Faizan**

GitHub: [@attahar6969](https://github.com/attahar6969)
