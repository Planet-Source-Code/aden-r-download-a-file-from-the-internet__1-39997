VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Database"
   ClientHeight    =   4170
   ClientLeft      =   3675
   ClientTop       =   2130
   ClientWidth     =   3150
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "WinMX (Setup - File Sharing)"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3840
      Width           =   3135
   End
   Begin VB.CommandButton Command7 
      Caption         =   "The Offspring - Defy You"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Disturbed - Prayer"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Murder Dolls - Dead In Hollywood"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Acid MSN Toolz (Zip File)"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "QuickScroll 1.1 (Zip File)"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "M$N Addon v2.1 (Zip File)"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Source Code For M$N Addon v2.1 (VB6)"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Audio:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Programs And Source:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Note: All Files Download Will Be Saved To:  C:\Program Files\"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
"URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long
    Public Function DownloadFile(URL As String, _
LocalFilename As String) As Boolean
Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

ret = DownloadFile("http://www.geocities.com/fenix2001uk/QuickScroll11.zip", _
"c:\Program Files\QuickScroll 1.1.zip")
End Sub

Private Sub Command4_Click()
ret = DownloadFile("http://www.geocities.com/fenix2001uk/AcidMSNToolz.zip", _
"c:\Program Files\Acid MSN Toolz.zip")
End Sub

Private Sub Command5_Click()
ret = DownloadFile("http://www.geocities.com/elerrorfiles/Murderdolls_deadinhollywood.wma", _
"c:\Program Files\Murder Dolls - Dead In Hollywood.wma")
End Sub

Private Sub Command6_Click()
ret = DownloadFile("http://www.geocities.com/elerrorfiles/Disturbed_Prayer.wma", _
"c:\Program Files\Disturbed - Prayer.wma")
End Sub

Private Sub Command7_Click()
ret = DownloadFile("http://www.geocities.com/elerrorfiles2/Offspring_DefyYou.wma", _
"c:\Program Files\Offspring - Defy You.wma")
End Sub

Private Sub Command8_Click()
ret = DownloadFile("http://dld25.winmx.com/74529165/winmx330.exe", _
"c:\Program Files\WinMX Setup")
End Sub

Private Sub Form_Load()

End Sub
