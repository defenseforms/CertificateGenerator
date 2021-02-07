# VBA Project: CertificateGenerator

### VBA Module: [GenerateCertificates.bas](./GenerateCertificates_bas.md)

  **Public Sub CreateCertificatePPT()**
    *Creates the certificates in one .pptx file.*
  **Public Function CertDate(str As String) As String**
    *Changes the date from mm/dd/yyyy to the DDth DAY OF MMMM YYYY*
  **Public Function SpellMonth(str As String) As String**
    *Converts the month number to the full spelled out month.*
  **Public Function CardinalDay(str As String) As String**
    *Converts the day as a number to the cardinal form.*
  **Public Sub ShowProgressBar()**
    *Displays the Progress Bar Form.*
  **Public Sub Progress(pctComp As Long)**
    *Updates the task progress bar.*
  **Public Sub CloseProgressBar()**
    *Closes the progress bar and displays a MsgBox with "Done".*

---
### VBA Module: [ProgressBar.frm](./ProgressBar_frm.md)

  **Private Sub UserForm_Activate()**

---
### Excel references  
*name*|*version*|*description*
---|---|---
VBA|4.2|Visual Basic For Applications
Excel|1.9|Microsoft Excel 16.0 Object Library
stdole|2.0|OLE Automation
Office|2.8|Microsoft Office 16.0 Object Library
