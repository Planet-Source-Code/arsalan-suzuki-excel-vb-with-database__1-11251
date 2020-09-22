VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Excel Demo Program"
   ClientHeight    =   2895
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Rx 
      Height          =   330
      Left            =   2370
      Top             =   1860
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BIBLIO.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BIBLIO.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Titles"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go on, Just Click the button and u wil get the result"
      ForeColor       =   &H00FF0000&
      Height          =   990
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   4050
      Begin VB.CommandButton cmdStart 
         Caption         =   "Run Microsoft Excel"
         Default         =   -1  'True
         Height          =   495
         Left            =   180
         TabIndex        =   1
         Top             =   315
         Width           =   3690
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Required for recordset. I though this was the easiest approach"
      DataSource      =   "Rx"
      Height          =   705
      Left            =   2370
      TabIndex        =   2
      Top             =   1095
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Menu mnuAuthor 
      Caption         =   "About The Author"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/-------About this program---------\
'This is second part of Excel example
'In this code the sheet is populated
'with data from database.
'Then you can save it in any format you
'like (which is supported by Excel)
'-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'I'm sure that you will learn something
'new from it
'-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'If you have problem be sure to check
'1.(Menu) Project----->References
'2.Check Miscrosoft Excel 9.0 Object Library
'If you have any further problem
'contact me
'I have tested this in Excel 2000 so I dont
'know whether it will work in prev version
'of Excel
'\----------------------------------/

'Define the required variable
Dim Excel As Excel.Application ' This is the excel program
Dim ExcelWBk As Excel.Workbook ' This is the work book
Dim ExcelWS As Excel.Worksheet ' This is the sheet

'Well I have broken this program in to several subs.
'This is the main sub from where every thing will
'be called.
Private Sub cmdStart_Click()
If cmdStart.Caption = "Run Microsoft Excel" Then
    StartExcel
    cmdStart.Caption = "Create WorkSheet"
    Exit Sub ' Otherwise it will do everything in one shot
ElseIf cmdStart.Caption = "Create WorkSheet" Then
    CreateWorkSheet
    cmdStart.Caption = "Populate WorkSheet"
    Exit Sub
ElseIf cmdStart.Caption = "Populate WorkSheet" Then
    PopulateWorkSheet
    cmdStart.Caption = "Save The WorkBook"
    Exit Sub
ElseIf cmdStart.Caption = "Save The WorkBook" Then
    SaveWorkSheet
    cmdStart.Caption = "Close The WorkBook and Excel"
    Exit Sub
ElseIf cmdStart.Caption = "Close The WorkBook and Excel" Then
    CloseWorkSheet
    cmdStart.Caption = "Bye , My demonstration is finished."
    Exit Sub
ElseIf cmdStart.Caption = "Bye , My demonstration is finished." Then
    Unload Me
End If

End Sub

Private Sub mnuAuthor_Click()
frmAuthor.Show 1
End Sub

Private Sub StartExcel()
On Error GoTo err:

Set Excel = GetObject(, "Excel.Application") ' Create Excel Object.
'Well you have to do like this.
'Above line if I used CreateObject, 1st time it would
'work fine but the second time my program would
'hang.Well I found this the easiest way to do it.
'But you can do it another way if you like.


'By default after creating the Excel it will
'not be shown on the screen.
'I you want to show it then
'Excel.Visible = True ' Show Excel

Exit Sub
err:
Set Excel = CreateObject("Excel.Application") 'Create Excel Object.

End Sub

Private Sub CreateWorkSheet()
Set ExcelWBk = Excel.Workbooks.Add 'Add this Workbook to Excel.
Set ExcelWS = ExcelWBk.Worksheets(1) ' Add this sheet to this Workbook

End Sub

Private Sub PopulateWorkSheet()
Dim row As Integer

row = 2 ' This is the row, start from 2nd row bec 1st row is header.

With Rx.Recordset
    'Add header
    ExcelWS.Cells(1, 1) = UCase(.Fields(1).Name)
    ExcelWS.Cells(1, 2) = UCase(.Fields(2).Name)
    ExcelWS.Cells(1, 3) = UCase(.Fields(3).Name)
    ExcelWS.Cells(1, 4) = UCase(.Fields(4).Name)
    ExcelWS.Cells(1, 5) = UCase(.Fields(5).Name)
    ExcelWS.Cells(1, 6) = UCase(.Fields(6).Name)
    ExcelWS.Cells(1, 7) = UCase(.Fields(7).Name)
    
    Do While Not row >= 100 ' populate with first 100 records
        'Total field is 7 so
        For i = 1 To 7
            'i is the coloumn and also fields value
            ExcelWS.Cells(row, i) = .Fields(i).Value
            'If you need explanation then remove this line and see what happens
            DoEvents
        Next
        row = row + 1 ' increment row
        Me.Caption = row & " records added"
        .MoveNext
    Loop
End With

End Sub


Private Sub SaveWorkSheet()
' Save the workbook on the desktop
'I didn't had time so I have not added export feature.
'If you want to export it into another format then just
'change this line.
'e.g
'ExcelWBk.SaveAs "c:\windows\desktop\Demo.txt", xlCSV
ExcelWBk.SaveAs "c:\windows\desktop\Demo.xls"
End Sub

Private Sub CloseWorkSheet()
' Close the WorkBook
ExcelWBk.Close
' Quit Excel app
Excel.Quit

MsgBox "You can find the saved Excel Sheet on your desktop"

End Sub
