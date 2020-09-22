VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CnC technocrats pvt. ltd., Ahmedabad (INDIA)"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "About Us"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -7
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1838
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Create and Print Word Document"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -7
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   638
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
   
   Private Sub Command1_Click()
      Dim oWord As Object
      Dim oDoc As Object
      
      'Create Word Document
      Set oWord = CreateObject("Word.application")

      'Set TRUE to watch created document , FALSE for invisible
      oWord.Visible = True
      
      'Add Data into created Word Docuemnts
      Set oDoc = oWord.Documents.Add
      oDoc.Range.Select
      
        
      'Add aline on Page
      oWord.Selection.TypeText "This is on page 1" & vbCr
      'Create 2nd Page
      oWord.Selection.InsertBreak 1
      'Add aline on 2nd Page
      oWord.Selection.TypeText "This is page 2"
      
      'Printing : Routine Called from Module
      SetPrinterDuplex Printer.DeviceName, 2
      
      'Set TRUE for Background Printing
      oDoc.PrintOut Background:=False
      
      'Printing : Routine Called from Module
      SetPrinterDuplex Printer.DeviceName, 1
      
      MsgBox "Print Done", vbMsgBoxSetForeground
      
      'Save Created Documents
      oDoc.Saved = True
      oDoc.Close
      Set oDoc = Nothing
   
      oWord.Quit
      Set oWord = Nothing
   End Sub

Private Sub Command2_Click()
    MsgBox "CnC technocrats pvt. ltd." & vbCrLf & _
           "Reg. Office : 302, Gulab Tower" & vbCrLf & _
           "Nr. Satadhar Cross Roads," & vbCrLf & _
           "Thaltej, Ahmedabad - 380054" & vbCrLf & _
           "Gujarat (INDIA)." & vbCrLf & _
           "Mobile : +91-9377702237" & vbCrLf & _
           "eMial : contact@cnctechnocrats.com" & vbCrLf & _
           "Website : http://www.cnctechnocrats.com", vbOKOnly, "About Us"
End Sub
