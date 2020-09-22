VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SetFileDate"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   510
      Left            =   3615
      TabIndex        =   7
      Top             =   1320
      Width           =   1440
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   510
      Left            =   1230
      TabIndex        =   6
      Top             =   1305
      Width           =   1440
   End
   Begin VB.TextBox txtTime 
      Height          =   315
      Left            =   3690
      TabIndex        =   3
      Top             =   720
      Width           =   1245
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   1395
      TabIndex        =   2
      Top             =   720
      Width           =   1110
   End
   Begin VB.TextBox txtFile 
      Height          =   345
      Left            =   600
      TabIndex        =   0
      Top             =   210
      Width           =   5580
   End
   Begin MSComDlg.CommonDialog cmddlgFile 
      Left            =   6150
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2955
      TabIndex        =   5
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   705
      TabIndex        =   4
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Read_File()
    
    'this reads the date and time from the file
    
    'declare variables
     Dim Result As Long
     
    'check we've selected a file
     If txtFile <> "" Then
        'yep - now open
         hFile = OpenFile(txtFile, FileInfoStruct, OF_READWRITE)
         
         If hFile = HFILE_ERROR Then
            'something went wrong
             MsgBox "Error opening file", vbCritical + vbOKOnly
            'finish
             Exit Sub
         End If
         
        'else read
         Result = GetFileTime(hFile, CreateTime, LastAccessTime, LastWriteTime)
         
        'fill boxes
         If Result <> 0 Then
            'got time - convert to local time
             Result = FileTimeToLocalFileTime(CreateTime, CreateTime)
             
            'convert to system time
             Result = FileTimeToSystemTime(CreateTime, SysTime)
             
            'fill boxes
             txtDate = SysDate_To_String(SysTime)
             txtTime = SysTime_To_String(SysTime)
             
         End If
        'close
         Result = CloseHandle(hFile)
         
     End If
End Sub

Private Sub cmdClose_Click()
    
    'close
     End
     
End Sub

Private Sub cmdSet_Click()
    
    'set the date and time on the file
    
    'open the file
     Dim Result As Long
     
    
    'check we've selected the file
     If txtFile = "" Or IsNull(txtFile) Then
        'no
         MsgBox "Please select a file.", vbInformation + vbOKOnly
         Exit Sub
     End If
     
    'check we've entered a date and time
     If txtDate = "" Or IsNull(txtDate) Or Not IsDate(txtDate) Then
        'no
         MsgBox "Please enter a date", vbInformation + vbOKOnly
         Exit Sub
     End If
     
     If txtTime = "" Or IsNull(txtTime) Or Len(txtTime) <> 12 Then
        'no
         MsgBox "Please enter a time", vbInformation + vbOKOnly
     End If
     
    'open
     hFile = OpenFile(txtFile.Text, FileInfoStruct, OF_READWRITE)
     
    'check we opened it
     If hFile = HFILE_ERROR Then
        'something went wrong
         MsgBox "Error opening file.", vbCritical + vbOKOnly
         Exit Sub
     End If
     
    'put the date and time into a systime variable
     SysTime = String_To_SysDateTime(txtDate, txtTime)
          
    'convert to a file time
     Result = SystemTimeToFileTime(SysTime, CreateTime)
     
    'convert to a UTC time
     Result = LocalFileTimeToFileTime(CreateTime, CreateTime)
     
    'set
     Result = SetFileTime(hFile, CreateTime, CreateTime, CreateTime)
     
    'close file
     Result = CloseHandle(hFile)
     
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    
    'format
     KeyAscii = FormatADateTextBox(KeyAscii, frmMain.txtDate)
     
End Sub

Private Sub txtFile_Click()
    
    'turn on error trapping
     On Error GoTo Lerr
     
    'setup common dialog
     cmddlgFile.CancelError = True
     cmddlgFile.ShowOpen
     
    'set
     txtFile = cmddlgFile.FileName
     
    'read
     Read_File
     
Lerr:
    'just exit
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    
    'format
     KeyAscii = FormatATimeTextBox(KeyAscii, frmMain.txtTime)
     
End Sub
