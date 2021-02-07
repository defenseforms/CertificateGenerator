VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Progress Indicator"
   ClientHeight    =   840
   ClientLeft      =   -156
   ClientTop       =   -672
   ClientWidth     =   2880
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
'
' Filename: ProgressBar.bas
' Date Created: Jul 2018
' Author: www.defenseforms.com
'
' Purpose: This Module is called when the Progress Bar Form is Activated.
'
'==============================================================================

Option Explicit


Private Sub UserForm_Activate()
    If (TypeName(Application.Caller) = "String") Then
        If (Application.Caller = "GenCertsBtn") Then
            GenerateCertificates.CreateCertificatePPT
        End If
    Else
        MsgBox "Error"
    End If
End Sub
