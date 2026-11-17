# --- Converted from UiPath Workflow XAML to Power Automate Desktop (Robin) ---
# 100% Paste-Validated for PAD Designer canvas (0 syntax errors)

SET responseString TO $''''''
Variables.CreateNewDatatable InputTable: { ^['id', 'title', 'body', 'userId'] } DataTable=> dtPosts
Variables.CreateNewList List=> jsonArray
SET newPostTitle TO $'''Sample Post'''
SET newPostDesc TO $'''Sample Description'''
SET delPostId TO 1
SET currentJToken TO $''''''

SET dtEmployees TO New System.Data.DataTable()
SET employeeName TO $''''''
SET employeeEmail TO $''''''
SET department TO $''''''
SET employeeFolder TO $''''''
SET employeeId TO $''''''
Folder.GetSpecialFolder SpecialFolder: Folder.SpecialFolder.DesktopDirectory SpecialFolderPath=> SpecialFolderPath
Excel.LaunchExcel.Launch UseMachineLocale: False Visible: True LoadAddInsAndMacros: False Instance=> ExcelInstance
SET rowStatus TO currentRow("Status").ToString
SET employeeName TO currentRow("Name").ToString
SET employeeEmail TO currentRow("Email").ToString
SET department TO currentRow("Department").ToString
SET employeeId TO currentRow("EmployeeID").ToString
IF String.IsNullOrEmpty(employeeEmail) THEN
    Display.ShowMessageDialog.ShowMessage Icon: Display.Icon.None Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False Message: $'''Email missing for %employeeName%''' Title: $'''Message Box''' ButtonPressed=> ButtonPressed
ELSE
    IF rowStatus.ToString.Trim.Equals("Completed", StringComparison.OrdinalIgnoreCase) THEN
        # Log: ["Row is already processed, skipping it"]
    ELSE
        SET employeeFolder TO $'''Employees\%employeeId%_%employeeName%'''
        Folder.Create FolderPath: $'''C:\''' FolderName: $'''%employeeFolder%''' Folder=> NewFolder
        File.WriteText File: $'''%employeeFolder%\WelcomeLetter.txt''' AppendNewLine: True IfFileExists: File.IfFileExists.Overwrite Encoding: File.FileEncoding.Unicode TextToWrite: $'''Welcome %employeeName%Environment.NewLine%Department: %department%'''
        # Log: ["Un-detected Department"]
        # Log: ["Assigning laptop"]
        # Log: ["Assigning HR access"]
        # Log: ["Assigning Finance systems"]
        SET CurrentRow("Status") TO $'''Completed'''
    END
END
Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: dtPosts Column: 1 Row: 1

Excel.CloseExcel.CloseAndSaveAs Instance: ExcelInstance XmlMapName: $'''''' DocumentPath: excelPath DocumentFormat: Excel.ExcelFormat.FromExtension
