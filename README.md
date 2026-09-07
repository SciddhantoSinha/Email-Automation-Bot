# 📧 Email Automation Bot

An automated email processing solution built using **UiPath** and **Microsoft 365 Outlook** that retrieves emails from an Outlook inbox, extracts important email information, downloads email attachments, and generates a structured Excel processing report.



## 📌 Project Overview

The **Email Automation Bot** automates repetitive email-processing tasks using UiPath and Microsoft 365 Outlook integration.

The bot connects to a Microsoft 365 Outlook account, retrieves emails from the Inbox, processes each email individually, downloads available attachments into a designated project folder, and records email information in an Excel report.

This project demonstrates practical **RPA, email automation, Microsoft 365 integration, file handling, DataTable processing, and Excel automation**.



## 🎯 Objectives

- Automate retrieval of emails from Microsoft 365 Outlook.
- Process emails from the Outlook Inbox.
- Extract sender, subject, and received date/time information.
- Identify whether an email contains attachments.
- Automatically download email attachments.
- Store processed email information in a structured DataTable.
- Generate an Excel report containing the processed email information.
- Reduce repetitive manual email-processing activities.



## 🛠️ Technologies Used

- **UiPath Studio**
- **Microsoft 365 Outlook**
- **UiPath Microsoft Office 365 Activities**
- **UiPath Excel Activities**
- **DataTable**
- **VB.NET Expressions**
- **File and Folder Automation**



## ⚙️ Workflow

The automation follows the following workflow:

```text
Start
  ↓
Connect to Microsoft 365 Outlook
  ↓
Retrieve Emails from Outlook Inbox
  ↓
For Each Email
  ↓
Extract Email Information
  ├── Sender
  ├── Subject
  └── Received Date/Time
  ↓
Check for Attachments
  ↓
Download Email Attachments
  ↓
Add Email Information to DataTable
  ↓
Generate Excel Report
  ↓
End
````



## 🔄 Process Details

### 1. Connect to Microsoft 365 Outlook

The bot uses the **Microsoft 365 Outlook activities** to establish a connection with the user's Outlook account.

Authentication is handled through the Microsoft 365 connection configured in UiPath.



### 2. Retrieve Emails

The bot retrieves emails from the **Inbox** using the Microsoft 365 Outlook:

**Get Email List**

The workflow is configured to retrieve emails from the Inbox for processing.



### 3. Iterate Through Emails

The retrieved emails are processed using:

**For Each Email**

Each email is assigned to the `CurrentEmail` variable for individual processing.



### 4. Extract Email Information

For every email, the bot extracts important information including:

* Sender email address
* Email subject
* Received date/time
* Attachment availability

The extracted information is stored for reporting.



### 5. Download Email Attachments

The bot uses:

**Download Email Attachments**

Attachments associated with the current email are automatically downloaded to the project's attachment storage folder.

This eliminates the need to manually open emails and save attachments individually.



### 6. Store Results in DataTable

A DataTable named:

```text
emailResults
```

is used to temporarily store the processed email information.

The DataTable contains the following columns:

| Column         | Description                           |
| -------------- | ------------------------------------- |
| Sender         | Email sender                          |
| Subject        | Email subject                         |
| ReceivedTime   | Date/time when the email was received |
| AttachmentName | Attachment information                |
| Status         | Attachment processing status          |



### 7. Generate Excel Report

After processing the emails, the collected information is written to:

```text
Output/EmailProcessingResults.xlsx
```

The Excel report provides a structured record of the emails processed by the automation.



## 📂 Project Structure

```text
Email Automation Bot/
│
├── Attachments/
│   └── Downloaded email attachments
│
├── Attachments_Test/
│   └── Attachment testing/output folder
│
├── Failed/
│   └── Failed-processing files
│
├── Input/
│   └── Input files, if required
│
├── Output/
│   └── EmailProcessingResults.xlsx
│
├── Processed/
│   └── Processed files
│
├── Main.xaml
├── project.json
├── project.uiproj
├── entry-points.json
└── README.md
```



## 📊 Excel Output

The bot generates an Excel report named:

**EmailProcessingResults.xlsx**

The report provides structured information about the emails processed by the bot.

Example:

| Sender                                            | Subject         | ReceivedTime | AttachmentName | Status           |
| ------------------------------------------------- | --------------- | ------------ | -------------- | ---------------- |
| [sender@example.com](mailto:sender@example.com)   | Job Opportunity | 09/07/2026   |                | Attachment Found |
| [example@example.com](mailto:example@example.com) | Newsletter      | 09/06/2026   |                | No Attachment    |



## 🧩 UiPath Activities Used

The project uses several UiPath activities, including:

* **Get Email List**
* **For Each Email**
* **Download Email Attachments**
* **Add Data Row**
* **Build Data Table**
* **Write Range Workbook**

These activities work together to automate email retrieval, processing, attachment handling, and reporting.



## 📦 Dependencies

The project uses the following UiPath packages:

```text
UiPath.Excel.Activities
UiPath.Mail.Activities
UiPath.System.Activities
UiPath.MicrosoftOffice365.Activities
```

The required packages are restored automatically when the project is opened in a compatible UiPath Studio environment.



## 🚀 How to Run

### Prerequisites

Before running the project, make sure you have:

* UiPath Studio installed.
* A Microsoft 365 account with Outlook access.
* Access to the Outlook Inbox that will be processed.
* Required UiPath packages installed.
* Permission to access Microsoft 365 Outlook through UiPath.



### Steps

1. Clone or download this repository.
2. Open the project in **UiPath Studio**.
3. Allow UiPath to restore the required dependencies.
4. Configure the Microsoft 365 Outlook connection if required.
5. Verify that the Outlook Inbox is selected.
6. Ensure the attachment destination folder exists.
7. Run `Main.xaml`.
8. Check the `Output` folder for:

   ```text
   EmailProcessingResults.xlsx
   ```
9. Check the attachment output folder for downloaded email attachments.



## 🔐 Microsoft 365 Authentication

The project uses a Microsoft 365 Outlook connection rather than storing an Outlook password directly inside the workflow.

The connection requires Microsoft 365 authorization and appropriate permissions to access mail data.



## 📁 Input and Output

### Input

The primary input for this automation is the user's:

**Microsoft 365 Outlook Inbox**

The bot retrieves emails directly from Outlook.

### Output

The automation produces:

* Downloaded email attachments
* `EmailProcessingResults.xlsx`



## ✨ Key Features

### 📩 Automated Email Retrieval

Automatically retrieves emails from the Outlook Inbox.

### 🔍 Email Information Extraction

Extracts sender, subject, and received date/time information.

### 📎 Attachment Automation

Automatically downloads email attachments without requiring manual saving.

### 📊 Excel Reporting

Creates a structured Excel report containing the processed email information.

### 🔄 Automated Processing

Processes multiple emails automatically using a `For Each Email` loop.

### 🔗 Microsoft 365 Integration

Uses Microsoft 365 Outlook integration for cloud-based email access.



## 💡 Business Use Cases

This type of automation can be applied to many business processes, including:

* Recruitment email processing
* Job application attachment collection
* Invoice attachment collection
* Customer support email processing
* Vendor document collection
* Purchase order processing
* Document management
* Internal communication workflows
* Automated email reporting

For example, a recruitment team could use the bot to automatically collect resumes received through email and store them for further processing.



## 📈 Benefits

* Reduces repetitive manual email processing.
* Automates attachment downloading.
* Improves consistency in email data collection.
* Creates structured email-processing records.
* Reduces manual Excel data entry.
* Saves time when processing large numbers of emails.
* Provides a foundation for more advanced intelligent automation workflows.



## 🔮 Future Enhancements

The automation can be extended with additional capabilities such as:

* Automatic email categorization.
* Sender-based processing rules.
* Subject-based filtering.
* Attachment type filtering.
* Automatic file renaming.
* Moving processed emails into dedicated folders.
* Exception handling and logging.
* Email notification after processing.
* Integration with databases.
* AI-based email classification.
* NLP-based email understanding.
* Intelligent document processing for email attachments.
* Integration with downstream RPA workflows.



## 📸 Screenshots

Screenshots demonstrating the UiPath workflow and generated Excel output can be added here.

Example:

```text
screenshots/
├── outlook-connection.png
├── email-processing-workflow.png
├── attachment-download.png
└── excel-output.png
```



## 🎓 Learning Outcomes

This project demonstrates practical experience with:

* Robotic Process Automation (RPA)
* UiPath workflow development
* Microsoft 365 Outlook automation
* Email processing
* Attachment automation
* DataTable manipulation
* Excel automation
* File and folder handling
* VB.NET expressions
* Cloud application integration
* Automated reporting



## 👨‍💻 Author

**Sciddhanto Sinha**

 That keeps the GitHub README technically honest and consistent with your other projects.
```
