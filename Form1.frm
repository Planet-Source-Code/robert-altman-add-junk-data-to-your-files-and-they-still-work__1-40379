VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H0000FF00&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Byte It!"
      Height          =   330
      Left            =   855
      TabIndex        =   7
      Top             =   2835
      Width           =   2085
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3060
      TabIndex        =   5
      Top             =   2295
      Width           =   1005
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   2430
      TabIndex        =   3
      Top             =   585
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   180
      TabIndex        =   2
      Top             =   585
      Width           =   2130
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   90
      Width           =   4380
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   3285
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblsize 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   2430
      Width           =   735
   End
   Begin VB.Label pd 
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   3285
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Percent Done:"
      Height          =   330
      Left            =   3510
      TabIndex        =   8
      Top             =   3060
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Size ():"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   2115
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes to add:"
      Height          =   240
      Left            =   3060
      TabIndex        =   4
      Top             =   2115
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Strt, Bte
Private Sub Command1_Click()
Dim Slash As String
If Val(Text1.Text) <= 0 Then
    MsgBox "Please enter a valid value.", vbExclamation, "Byte Me"
    Text1.SetFocus: SendKeys "+{HOME}"
    Exit Sub
End If

pb.Max = Val(Text1.Text): pb.Min = 0: pb.Value = 0
If Len(File1.Path) = 3 Then
    Slash = ""
Else
    Slash = "\"
End If

Strt = Val(lblsize.Caption)
Open File1.Path & Slash & File1 For Binary As #1
For Bte = 1 To Val(Text1.Text)
    Put #1, Strt - 1 + Bte, 0
    pb.Value = Bte
    pd.Caption = (Bte / Val(Text1.Text)) * 100 & " %"
    pd.Refresh
Next Bte
Close

MsgBox "Complete!", vbExclamation, "Byte Me"
File1_Click
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
On Error GoTo Ehandle
Dir1.Path = Drive1
Exit Sub
Ehandle:
r = MsgBox("Device not ready!", vbCritical + vbRetryCancel, "Byte Me")
If r = vbRetry Then Resume
Exit Sub
End Sub

Private Sub File1_Click()
Label2.Caption = "Size ( " & File1 & " ):"
lblsize.Caption = ShowFileSize(File1) & " bytes."
End Sub

Function ShowFileSize(file)
        Dim fs, f, s
        Dim Slash As String
If Len(File1.Path) = 3 Then
    Slash = ""
Else
    Slash = "\"
End If
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(File1.Path & Slash & file)
    ShowFileSize = f.Size
End Function

Private Sub Timer1_Timer()
Stop
pd.Caption = (Bte / Val(Text1.Text)) * 100 & " %"
End Sub
