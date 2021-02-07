Attribute VB_Name = "GenerateCertificates"
'==============================================================================
'
' Filename: GenerateCertificates.bas
' Date Created: Jul 2018
' Author: www.defenseforms.com
'
' Purpose: This Module is used to create student certificates.
'
'==============================================================================

Option Explicit

Private Const CERT_TEMPLATE As String = "CertificateTemplate.pptx"
Private Const CERT_FOLDER As String = "GeneratedCertificates"
Private Const CLASS_NUM_CELL As String = "G4"
Private Const GRAD_DATE_CELL As String = "G2"
Private Const RANK_COL As String = "A"
Private Const LAST_NAME_COL As String = "B"
Private Const FIRST_NAME_COL As String = "C"
Private Const MID_INIT_COL As String = "D"
Private Const START_ROW As Integer = 2

Private Type student
    Rank As String
    LastName As String
    FirstName As String
    MidInitial As String
End Type

'/**
' * Creates the certificates in one .pptx file.
' */
Public Sub CreateCertificatePPT()
    Dim certTemp As String
    Dim certFold As String
    certTemp = ThisWorkbook.Path & "\" & CERT_TEMPLATE
    certFold = ThisWorkbook.Path & "\" & CERT_FOLDER & "\"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(Sheet1.Name)
    ws.Unprotect

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim slideCount As Long
    slideCount = 1

    Dim endRow As Long
    endRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

    If Dir(certFold, vbDirectory) = "" Then
        MkDir certFold
    Else
        If (fso.GetFolder(certFold).Files.count > 0) Then
            Kill certFold & "*.*"
        End If
    End If

    Dim ppt As Object
    Dim presentation As Object
    Dim newSlide As Object
    Dim nameBox As Object
    Dim classNumBox As Object
    Dim dateBox As Object

    On Error Resume Next
    Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = False
    On Error GoTo Err
    Set presentation = ppt.Presentations.Open(certTemp, , , msoFalse)
    On Error Resume Next
    Set newSlide = presentation.Slides.Item(1)
    Set nameBox = newSlide.Shapes("NameBox")
    classNumBox.TextFrame.TextRange = ws.Range(CLASS_NUM_CELL).Value
    Set dateBox = newSlide.Shapes("DateBox")
    dateBox.TextFrame.TextRange = "Given this " & CertDate(ws.Range(GRAD_DATE_CELL).Value)

    Dim i As Long
    Dim pctComp As Double
    Dim fullName As String
    Dim fullText As String
    Dim student As student

    For i = START_ROW To endRow
        student.Rank = ws.Range(RANK_COL & CStr(i)).Value
        student.LastName = ws.Range(LAST_NAME_COL & CStr(i)).Value
        student.FirstName = ws.Range(FIRST_NAME_COL & CStr(i)).Value
        student.MidInitial = ws.Range(MID_INIT_COL & CStr(i)).Value
        If (student.LastName <> "" And student.FirstName <> "") Then
            Set newSlide = presentation.Slides(slideCount).Duplicate
            slideCount = slideCount + 1
            Set newSlide = presentation.Slides(slideCount - 1)
            Set nameBox = newSlide.Shapes("NameBox")

            If (student.MidInitial = "" Or student.MidInitial = "0" _
                    Or student.MidInitial = " ") Then
                fullName = UCase(student.LastName) & ", " & _
                        UCase(student.FirstName)
            Else
                fullName = UCase(student.LastName) & ", " & _
                        UCase(student.FirstName) & " " & _
                        UCase(student.MidInitial) & "."
            End If

            fullText = student.Rank & " " & fullName
            nameBox.TextFrame.TextRange = fullText

            pctComp = Round((i - (START_ROW - 1)) / (endRow) * 100, 0.1)
            Progress CLng(pctComp)

            fullName = ""
            fullText = ""
        End If
    Next i

    Set newSlide = presentation.Slides(slideCount).Delete
    presentation.SaveAs (certFold & ws.Range(CLASS_NUM_CELL).Value & "_Certificates.pptx")

    CloseProgressBar
    presentation.Close
    ppt.Quit
    ws.Protect
    Set ws = Nothing
    Set ppt = Nothing
    Set presentation = Nothing
    Set newSlide = Nothing
    Set nameBox = Nothing
    Set classNumBox = Nothing
    Set dateBox = Nothing
    Exit Sub
Err:
    MsgBox ("Unable to find the template file at: " & Chr(13) & certTemp)
End Sub

'------------------------------------------------------------------------------
'  DATE FUNCTIONS
'------------------------------------------------------------------------------

'/**
' * Changes the date from mm/dd/yyyy to the DDth DAY OF MMMM YYYY
' * for use on the certificates.
' * @param str the date string in mm/dd/yyyy format.
' * @returns {String} the formatted date for use on certificates.
' */
Public Function CertDate(str As String) As String
    Dim strArr() As String
    Dim year As String
    Dim month As String
    Dim day As String

    strArr() = Split(str, "/")
    If (UBound(strArr()) >= 2) Then
        year = strArr(2)
        month = strArr(0)
        day = strArr(1)

        CertDate = CardinalDay(day) & " day of " & SpellMonth(month) & " " + year
    End If
End Function

'/**
' * Converts the month number to the full spelled out month.
' * @param str the month in number format.
' * @returns {String} the spelled out month.
' */
Public Function SpellMonth(str As String) As String
    If (CLng(str) = 1) Then
        SpellMonth = "January"
    ElseIf (CLng(str) = 2) Then
        SpellMonth = "February"
    ElseIf (CLng(str) = 3) Then
        SpellMonth = "March"
    ElseIf (CLng(str) = 4) Then
        SpellMonth = "April"
    ElseIf (CLng(str) = 5) Then
        SpellMonth = "May"
    ElseIf (CLng(str) = 6) Then
        SpellMonth = "June"
    ElseIf (CLng(str) = 7) Then
        SpellMonth = "July"
    ElseIf (CLng(str) = 8) Then
        SpellMonth = "August"
    ElseIf (CLng(str) = 9) Then
        SpellMonth = "September"
    ElseIf (CLng(str) = 10) Then
        SpellMonth = "October"
    ElseIf (CLng(str) = 11) Then
        SpellMonth = "November"
    ElseIf (CLng(str) = 12) Then
        SpellMonth = "December"
    End If
End Function

'/**
' * Converts the day as a number to the cardinal form.
' * @param str the day in standard number format.
' * @returns {String} the day in cardinal format.
' */
Public Function CardinalDay(str As String) As String
    If (CLng(str) >= 10 And CLng(str) <= 20) Then
        CardinalDay = str + "th"
    ElseIf (CLng(Right(str, 1)) = 1) Then
        CardinalDay = str + "st"
    ElseIf (CLng(Right(str, 1)) = 2) Then
        CardinalDay = str + "nd"
    ElseIf (CLng(Right(str, 1)) = 3) Then
        CardinalDay = str + "rd"
    Else
        CardinalDay = str + "th"
    End If
End Function

'------------------------------------------------------------------------------
'  PROGRESS BAR
'------------------------------------------------------------------------------

'/**
' * Displays the Progress Bar Form.
' */
Public Sub ShowProgressBar()
    ProgressBar.Show
End Sub

'/**
' * Updates the task progress bar.
' */
Public Sub Progress(pctComp As Long)
    ProgressBar.Text.Caption = CStr(pctComp) + "% Completed"
    ProgressBar.Bar.Width = pctComp * 2
    DoEvents
End Sub

'/**
' * Closes the progress bar and displays a MsgBox with "Done".
' */
Public Sub CloseProgressBar()
    Progress CLng(0)
    ProgressBar.Hide
    Set ProgressBar = Nothing
    MsgBox "Done"
End Sub
