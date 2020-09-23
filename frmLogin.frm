VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "App Login"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWebsite 
      Caption         =   "Patrick's website"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok!"
      Default         =   -1  'True
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "-"
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
   End
   Begin VB.ComboBox cmbUsers 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblEnterPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your password:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1560
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select your username:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label lblLogin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "App Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      Height          =   405
      Left            =   90
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   195
      Picture         =   "frmLogin.frx":038A
      Top             =   240
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogin.frx":0714
      Top             =   165
      Width           =   240
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Left            =   90
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************
'* CODE BY: PATRICK MOORE (ZELDA) *
'* Feel free to re-distribute or  *
'* Use in your own projects.      *
'* Giving credit to me would be   *
'* nice :)                        *
'*                                *
'* Please vote for me if you find *
'* this code useful :]   -Patrick *
'**********************************
'http://members.nbci.com/erx931/VB/
'
'PS: Please look for more submissions to PSC by me
'    I've recently been working on a lot of them.
'    :))  All my submissions are under author name
'    "Patrick Moore (Zelda)"

'Function related to the shell website
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub GotoURL()
Dim Success As Long, URL As String

'Define the URL of my website
URL = "http://members.nbci.com/erx931/VB/"

'Open my website in your default
'web browser
Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", 1)
End Sub

Private Sub cmdOk_Click()
If cmbUsers.ListIndex = 0 Then
    'If the New User item is highlighted,
    'do this code
    Dim NewUser As String, NewPass As String
    
    'Ask the user for their username
    NewUser = InputBox("Enter a username:", "New User")
    
    'If they click cancel or left the username
    'blank, discontinue
    If NewUser = "" Then Exit Sub
    
    'Ask the user for a password
    NewPass = InputBox("Enter a password:", "New User")
    
    'Add one to the number of registered users
    NumUsers = NumUsers + 1
    
    'Save the Username, Password to the registry
    SaveSetting "AppLogin", "User" & NumUsers, "Username", NewUser
    SaveSetting "AppLogin", "User" & NumUsers, "Password", NewPass
    
    'Save the new number of registered users
    SaveSetting "AppLogin", "NumUsers", "NumUsers", NumUsers & ""
    
    'Add the new registered username to the combobox
    cmbUsers.AddItem NewUser
    
    'Highlight it
    cmbUsers.ListIndex = cmbUsers.ListCount - 1
    
    'Set the password textbox to the password they entered
    txtPassword.Text = NewPass
    
    'Msgbox the user their account information
    MsgBox "User has been created.  Please write down the following information for future use:" & vbCrLf & "Username:" & vbTab & NewUser & vbCrLf & "Password:" & vbTab & NewPass, vbInformation + vbOKOnly
End If

Dim Password As String
'Get the password for the user
Password = LCase(GetSetting("AppLogin", "User" & cmbUsers.ListIndex, "Password", ""))

If Password <> LCase(txtPassword.Text) Then
    'If the password in the registry doesn't
    'match the password they entered,
    'notify them using msgbox, then exit the sub
    MsgBox "The password you specified is incorrect.", vbExclamation + vbOKOnly
    Exit Sub
End If

'Set the Globals

'Which number the user is
UserNum = cmbUsers.ListIndex + 1

'What the user's username is
UserName = cmbUsers.List(cmbUsers.ListIndex)

'Show the main form
frmMain.Show

'Unload this form (the login form)
Unload Me
End Sub

Private Sub cmdWebsite_Click()
GotoURL
End Sub

Private Sub Form_Load()
Dim X As Integer

'Get the number of registered users
'from the registry
NumUsers = GetSetting("AppLogin", "NumUsers", "NumUsers", "0")

'Add the "New User" item to the combo box
cmbUsers.AddItem "New User"

'If there are registered users,
'add each to the combo box
If NumUsers > 0 Then
    'Cycle through each user
    For X = 1 To NumUsers
        'Add the username to the combo box
        cmbUsers.AddItem GetSetting("AppLogin", "User" & X, "Username", "")
    Next X
End If

'Highlight the "New User" item
cmbUsers.ListIndex = 0
End Sub
