# 📧 Email Automation Bot

An automated email processing solution built using **UiPath** and **Microsoft 365 Outlook** that retrieves emails from an Outlook Inbox, extracts important email information, and generates a structured Excel processing report.



## 📌 Project Overview

The **Email Automation Bot** automates repetitive email-processing tasks using UiPath and Microsoft 365 Outlook integration.

The bot connects to a Microsoft 365 Outlook account, retrieves emails from the Inbox, processes each email individually, extracts relevant email information, stores the results in a DataTable, and generates an Excel report.

This project demonstrates practical **RPA, Microsoft 365 integration, email automation, DataTable processing, and Excel automation**.



## 🎯 Objectives

- Automate retrieval of emails from Microsoft 365 Outlook.
- Process emails from the Outlook Inbox.
- Extract sender information.
- Extract email subjects.
- Extract received date and time.
- Identify whether an email contains attachments.
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
- **Excel Automation**


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
  ├── Received Date/Time
  └── Attachment Availability
  ↓
Store Information in DataTable
  ↓
Generate Excel Report
  ↓
End
````



## 🔄 Process Details

### 1. Connect to Microsoft 365 Outlook

The bot uses **Microsoft 365 Outlook activities** to establish a connection with the user's Outlook account.

Authentication is handled through the Microsoft 365 connection configured in UiPath.



### 2. Retrieve Emails

The bot retrieves emails from the **Outlook Inbox** using the Microsoft 365 Outlook:

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

The extracted information is then added to the processing DataTable.



### 5. Determine Attachment Availability

The bot checks the attachment count of each email.

Emails are classified as:

```text
Attachment Found
```

or

```text
No Attachment
```

This information is stored in the **Status** column of the output DataTable.



### 6. Store Results in DataTable

A DataTable named:

```text
emailResults
```

is used to temporarily store the processed email information.

The DataTable contains the following columns:

| Column         | Description                                  |
| -------------- | -------------------------------------------- |
| Sender         | Email sender address                         |
| Subject        | Email subject                                |
| ReceivedTime   | Date/time associated with the email          |
| AttachmentName | Reserved column for attachment information   |
| Status         | Indicates whether an attachment was detected |

The current workflow populates the Sender, Subject, ReceivedTime, and Status information.



### 7. Generate Excel Report

After processing the emails, the collected information is written to:

```text
Output\EmailProcessingResults.xlsx
```

The Excel report provides a structured record of the emails processed by the automation.

The workflow uses a **relative project path** for the output workbook, making the project more portable between compatible environments.


## 📊 Excel Output

The bot generates an Excel report named:

**EmailProcessingResults.xlsx**

The report contains structured information about the emails processed by the bot.

Example:

| Sender                                            | Subject         | ReceivedTime | AttachmentName | Status           |
| ------------------------------------------------- | --------------- | ------------ | -------------- | ---------------- |
| [sender@example.com](mailto:sender@example.com)   | Job Opportunity | 09/07/2026   |                | Attachment Found |
| [example@example.com](mailto:example@example.com) | Newsletter      | 09/06/2026   |                | No Attachment    |



## 🧩 UiPath Activities Used

The project uses several UiPath activities, including:

* **Get Email List**
* **For Each Email**
* **Add Data Row**
* **Build Data Table**
* **Write Range Workbook**

These activities work together to retrieve emails, process email information, store results, and generate an Excel report.



## 📦 Dependencies

The project uses UiPath activity packages for:

```text
UiPath.Excel.Activities
UiPath.Mail.Activities
UiPath.System.Activities
UiPath.MicrosoftOffice365.Activities
```

The required packages can be restored when the project is opened in a compatible UiPath Studio environment.



## 🚀 How to Run

### Prerequisites

Before running the project, make sure you have:

* UiPath Studio installed.
* A Microsoft 365 account with Outlook access.
* Access to the Outlook Inbox that will be processed.
* The required UiPath packages installed.
* Permission to access Microsoft 365 Outlook through UiPath.



### Steps

1. Clone or download this repository.
2. Open the project in **UiPath Studio**.
3. Allow UiPath to restore the required dependencies.
4. Configure the Microsoft 365 Outlook connection if required.
5. Verify that the Outlook Inbox is selected.
6. Run `Main.xaml`.
7. After execution, open the `Output` folder.
8. Open:

   ```text
   EmailProcessingResults.xlsx
   ```
9. Review the generated email-processing report.



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

```text
Output/
└── EmailProcessingResults.xlsx
```

The Excel file contains the structured information collected from the processed emails.



## ✨ Key Features

### 📩 Automated Email Retrieval

Automatically retrieves emails from the Outlook Inbox.

### 🔍 Email Information Extraction

Extracts:

* Sender
* Subject
* Received date/time
* Attachment availability

### 📊 Excel Reporting

Generates a structured Excel report containing the processed email information.

### 🔄 Automated Email Processing

Processes multiple emails automatically using the `For Each Email` activity.

### 🔗 Microsoft 365 Integration

Uses Microsoft 365 Outlook integration to access and process emails.

### 🗂️ Structured Data Processing

Uses a DataTable to organize the information extracted from multiple emails before generating the final Excel report.



## 💡 Business Use Cases

This type of automation can be applied to many business processes, including:

* Recruitment email processing
* Job application email collection
* Customer support email processing
* Vendor communication processing
* Document collection
* Internal communication workflows
* Automated email reporting
* Email-based data collection

For example, a recruitment team could use the bot to collect information from job-related emails and generate a structured report for further processing.



## 📈 Benefits

* Reduces repetitive manual email processing.
* Automates email data collection.
* Improves consistency in email information recording.
* Reduces manual Excel data entry.
* Generates structured reports automatically.
* Saves time when processing multiple emails.
* Provides a foundation for more advanced email automation workflows.



## 🔮 Future Enhancements

The automation can be extended with additional capabilities such as:

* Automatic attachment downloading.
* Automatic attachment renaming.
* Subject-based email filtering.
* Sender-based processing rules.
* Attachment type filtering.
* Automatic file organization.
* Moving processed emails into dedicated Outlook folders.
* Exception handling and logging.
* Email notifications after processing.
* Database integration.
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
└── excel-output.png
```


## 🎓 Learning Outcomes

This project demonstrates practical experience with:

* Robotic Process Automation (RPA)
* UiPath workflow development
* Microsoft 365 Outlook automation
* Email processing
* DataTable manipulation
* Excel automation
* VB.NET expressions
* Automated reporting
* Cloud application integration



## 👨‍💻 Author

**Sciddhanto Sinha**

GitHub:

**[https://github.com/SciddhantoSinha](https://github.com/SciddhantoSinha)**

---
```
