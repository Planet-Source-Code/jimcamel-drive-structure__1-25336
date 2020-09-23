VERSION 5.00
Begin VB.Form DirectStruct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Directory Structure"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Create Output Files"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.ListBox DrivesList 
      Height          =   2790
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Drives:"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Output File:"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "DirectStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare a bunch of variables
Dim d, i
Dim strDrives   As String

'Declare the drivetype type, which will hold the type
'(ie removable, fixed, etc) and the letter

Private Type Drivetype
    dtype As Integer
    letter As String
End Type

'Dim drives array and the Number of Drives
Dim Drives() As Drivetype
Dim noDrive As Integer

Private Sub Command1_Click()
Dim temp As Boolean
For i = 0 To noDrive
'check which drives have been selected
If DrivesList.Selected(i) = True Then
'find a free file number
Filenum = FreeFile
'change the label caption
Label2.Caption = "Please Wait, Building Output File"
Label2.Visible = True
'Open the file "output-" and the drive letter, ie "output-a.log"
Open Text1.Text & "output-" & Left(Drives(i).letter, Len(Drives(i).letter) - 1) & ".log" For Output As #Filenum
'Call the function to log to the file, with the recursion number and the root directory
logtofile 0, Drives(i).letter & "\"
'close the file
Close #Filenum
Label2.Caption = "Output file created"
End If
Next i
End Sub

Private Sub Form_Load()
    Dim FSO
    'create a FileSystemObject
    Set FSO = CreateObject("scripting.filesystemobject")
    noDrive = -1
    'find out all the drives and their types
    For Each d In FSO.Drives
        noDrive = noDrive + 1
        ReDim Preserve Drives(noDrive)
        Drives(noDrive).dtype = d.Drivetype
        Drives(noDrive).letter = d
    Next
    'add the drives and their types to the listbox
    For i = 0 To noDrive
        Select Case Drives(i).dtype
            Case 0
                DrivesList.AddItem "Unknown " & Drives(i).letter
            Case 1
                DrivesList.AddItem "Removable " & Drives(i).letter
            Case 2
                DrivesList.AddItem "Fixed " & Drives(i).letter
            Case 3
                DrivesList.AddItem "Remote " & Drives(i).letter
            Case 4
                DrivesList.AddItem "Cdrom " & Drives(i).letter
            Case 5
                DrivesList.AddItem "Ramdisk " & Drives(i).letter
        End Select
    Next i
Text1.Text = App.Path & "\"
End Sub

