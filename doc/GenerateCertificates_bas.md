# VBA Project: **CertificateGenerator**
## VBA Module: **GenerateCertificates**
### Type: StdModule  

This procedure list for repo (CertificateGenerator) was automatically created on 2/7/2021 12:04:58 AM by VBADeveloper.

Below is a section for each procedure in GenerateCertificates

---
VBA Procedure: **CreateCertificatePPT**  
Type: **Sub**  
Returns: **void**  
Return Description: ****  
Scope: **Public**  
Description: **Creates the certificates in one .pptx file.**  

*Public Sub CreateCertificatePPT()*  

**no arguments required for this procedure**


---
VBA Procedure: **CertDate**  
Type: **Function**  
Returns: **String**  
Return Description: ****  
Scope: **Public**  
Description: **Changes the date from mm/dd/yyyy to the DDth DAY OF MMMM YYYY**  

*Public Function CertDate(str As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
str|String|False||


---
VBA Procedure: **SpellMonth**  
Type: **Function**  
Returns: **String**  
Return Description: ****  
Scope: **Public**  
Description: **Converts the month number to the full spelled out month.**  

*Public Function SpellMonth(str As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
str|String|False||


---
VBA Procedure: **CardinalDay**  
Type: **Function**  
Returns: **String**  
Return Description: ****  
Scope: **Public**  
Description: **Converts the day as a number to the cardinal form.**  

*Public Function CardinalDay(str As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
str|String|False||


---
VBA Procedure: **ShowProgressBar**  
Type: **Sub**  
Returns: **void**  
Return Description: ****  
Scope: **Public**  
Description: **Displays the Progress Bar Form.**  

*Public Sub ShowProgressBar()*  

**no arguments required for this procedure**


---
VBA Procedure: **Progress**  
Type: **Sub**  
Returns: **void**  
Return Description: ****  
Scope: **Public**  
Description: **Updates the task progress bar.**  

*Public Sub Progress(pctComp As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
pctComp|Long|False||


---
VBA Procedure: **CloseProgressBar**  
Type: **Sub**  
Returns: **void**  
Return Description: ****  
Scope: **Public**  
Description: **Closes the progress bar and displays a MsgBox with "Done".**  

*Public Sub CloseProgressBar()*  

**no arguments required for this procedure**
