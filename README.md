# ms-excel-util

# Microsoft excel sheet as a Planner:

To create an Excel-based planner that can send email notifications via Outlook and display popup messages for pending tasks, you'll need to use VBA (Visual Basic for Applications) to automate the tasks. Here’s how you can set it up:

### Step 1: Setting Up Your Excel Planner Sheet
1. **Create the Planner Layout:**
   - Design your sheet with columns like:
     - `Task` (description of the task)
     - `Due Date` (when the task is due)
     - `Status` (e.g., Pending, Completed)
     - `Notification Sent` (Yes/No)

2. **Add Conditional Formatting:**
   - Highlight tasks that are overdue or due today for better visibility.

### Step 2: Setting Up a VBA Script for Email Notifications
You can use VBA to automatically send an email when you open the file if there are tasks due today.

1. **Press `Alt + F11`** to open the VBA editor.
2. **Insert a New Module:**
   - Right-click on any existing modules in the left-hand pane, choose "Insert" -> "Module."
   - Paste the following code:

   ```vba
   Sub SendEmailNotification()
       Dim OutlookApp As Object
       Dim OutlookMail As Object
       Dim Cell As Range
       Dim TaskSheet As Worksheet
       
       Set TaskSheet = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your actual sheet name
       
       ' Loop through tasks
       For Each Cell In TaskSheet.Range("B2:B100") ' Assuming tasks start from B2
           If Cell.Value = Date And Cell.Offset(0, 2).Value = "Pending" Then ' Check if due today and pending
               If Cell.Offset(0, 3).Value <> "Yes" Then ' Check if notification not already sent
                   Set OutlookApp = CreateObject("Outlook.Application")
                   Set OutlookMail = OutlookApp.CreateItem(0)
                   
                   ' Configure email
                   With OutlookMail
                       .To = "youremail@example.com" ' Replace with your email address
                       .Subject = "Task Due Today: " & Cell.Offset(0, -1).Value
                       .Body = "The following task is due today: " & vbCrLf & _
                              "Task: " & Cell.Offset(0, -1).Value & vbCrLf & _
                              "Due Date: " & Cell.Value
                       .Send
                   End With
                   
                   ' Mark as notification sent
                   Cell.Offset(0, 3).Value = "Yes"
               End If
           End If
       Next Cell
   End Sub
   ```

3. **Trigger the Macro When Opening the Workbook:**
   - In the VBA editor, double-click "ThisWorkbook" under "Microsoft Excel Objects" and paste:

   ```vba
   Private Sub Workbook_Open()
       Call SendEmailNotification
       Call ShowPendingTasksPopup
   End Sub
   ```

### Step 3: Creating a Popup Message for Pending Tasks
1. **Add the Following Code to Your Module:**

   ```vba
   Sub ShowPendingTasksPopup()
       Dim TaskSheet As Worksheet
       Dim TaskList As String
       Dim Cell As Range
       
       Set TaskSheet = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your actual sheet name
       TaskList = ""
       
       ' Loop through tasks to find today's pending tasks
       For Each Cell In TaskSheet.Range("B2:B100") ' Assuming tasks start from B2
           If Cell.Value = Date And Cell.Offset(0, 2).Value = "Pending" Then
               TaskList = TaskList & vbCrLf & "- " & Cell.Offset(0, -1).Value
           End If
       Next Cell
       
       ' Show a message if there are pending tasks
       If TaskList <> "" Then
           MsgBox "Tasks due today:" & vbCrLf & TaskList, vbInformation, "Pending Tasks"
       End If
   End Sub
   ```

### Step 4: Saving the Workbook
- Save your workbook as a macro-enabled file (`.xlsm`).

### Step 5: Testing the Automation
1. **Close and Reopen the Excel file** to see if the macro runs.
2. **Check your email for notifications** and see if the popup appears for today's tasks.

### Notes:
- **Enable Macros** when opening the file.
- **Adjust the Range** (`B2:B100`) as needed for your data.
- **Update the email address** in the VBA code to your actual address.

This setup will automate notifications and display reminders when you open the Excel sheet. Let me know if you need help customizing the solution further!
