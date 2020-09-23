VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "change proxy"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   2025
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   2280
      Top             =   2880
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "Form1.frx":0000
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Proxy"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Text            =   """"
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Text            =   " ""DIRECT"";"
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Text            =   " ""PROXY "
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2040
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Proxy File"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Proxy Being Used"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Software Made By Manu Mehrotra On 16th July 2004
'This software can change your proxy with a click of a button or it can mask your ip
'This software is fully functional on windows xp
'You can check Manu Mehrotra's website - http://www.manu.co.nr for more codes
'Please give him a feedback at - manumehrotra@rediffmail.com
'This software is for ie users only & not for those who use netscape or firefox i would be back 4 them soon
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub Command1_Click()
filew
Form4.Show
End Sub

Private Sub Command2_Click()
Close
List1.Clear
CommonDialog1.Filter = "Text Files|*.txt*"
CommonDialog1.ShowOpen
Text5.Text = CommonDialog1.FileName
openfile
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Command4_Click()
MsgBox "load the file containing ip and port in ip:port (127.00.00.99:80) format and select an ip to be used,then open a new ie window to find a masked ip which could be found out at http://www.whatismyip.com remember ip for new window would be changed only and not for all the ie windows "
End Sub



Private Sub Form_Load()
doregistrymodifications
End Sub
Private Sub doregistrymodifications()

'modifies registry

On Error Resume Next
Dim Enabled, Shell, Key
Const MigrateProxy = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MigrateProxy"
Const AutoConfigUrl = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigUrl"
Set Shell = CreateObject("wscript.shell")
Key = AutoConfigUrl
Shell.RegWrite Key, "file://C:/windows/m.pac", "REG_SZ"
Key = MigrateProxy
Shell.RegWrite Key, 0, "REG_DWORD"


End Sub




Private Sub openfile()

'loads ip addresses from a file

Dim newline As String
Close
Open Text5.Text For Input As #1
Do Until EOF(1)
Line Input #1, newline
List1.AddItem newline
Loop
Close #1
End Sub



Private Sub filew()

'creates a autoconfigurl file

filltextbox
Open "C:\windows\m.pac" For Output As #1
Print #1, "function FindProxyForURL(url, host) "
Print #1, "{"
Print #1, "if (isInNet(myIpAddress(),myIpAddress(),myIpAddress())) "
Print #1, "return" + Text1.Text + Text2.Text + Text4.Text + "; "
Print #1, "else"
Print #1, "return" + Text3.Text
Print #1, "}"
Close #1
End Sub



Private Sub filltextbox()

'Tells filew() what is the ip to be written by it in the file

Dim loopindex
For loopindex = 0 To List1.ListCount - 1
If List1.Selected(loopindex) Then
Text2.Text = List1.List(loopindex)
End If
Next loopindex
End Sub



Private Sub Form_Unload(Cancel As Integer)

'remove the autoconfigfile



Dim Enabled, Shell, Key
Const MigrateProxy = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MigrateProxy"
Const AutoConfigUrl = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigUrl"
Set Shell = CreateObject("wscript.shell")
Key = AutoConfigUrl
Shell.Regdelete AutoConfigUrl
Key = MigrateProxy
Shell.RegWrite Key, 0, "REG_DWORD"
On Error Resume Next
Kill "c:\windows\m.pac"

End Sub

Private Sub Timer1_Timer()
doregistrymodifications
filew
End Sub
