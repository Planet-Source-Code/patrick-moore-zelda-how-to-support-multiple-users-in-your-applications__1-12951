VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Welcome!"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok!"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   0
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblFavFood 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   465
   End
   Begin VB.Label lblLastLoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   465
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome, User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub cmdOk_Click()
'Unload this form
Unload Me

'End the project
End
End Sub

Private Sub Form_Load()
Dim LastLoad As String, FavFood As String

'Check the registry for when the user
'last loaded this app
LastLoad = GetSetting("AppLogin", "User" & UserNum, "LastLoad", "")

'Welcome the user specifically
lblWelcome.Caption = "Welcome, " & UserName

If LastLoad = "" Then
    'If they haven't loaded the app before,
    'tell them so, and ask for their favorite food
    lblLastLoad.Caption = "This is your first time loading this app!"
    
    'Ask for favorite food
    FavFood = InputBox("Please enter your favorite food:", "Favorite Food")
    
    'Store favorite food for this user in the registry
    SaveSetting "AppLogin", "User" & UserNum, "FavFood", FavFood
Else
    'They have loaded the app before
    
    'Have the caption reflect when they first
    'loaded this app
    lblLastLoad.Caption = "You last loaded this app on " & LastLoad & "."
    
    'Get the favorite food for this user
    'from the registry
    FavFood = GetSetting("AppLogin", "User" & UserNum, "FavFood", "")
End If

'Set the caption of the Favorite Food label
'to their favorite food, as stored in the registry
lblFavFood.Caption = "Your favorite food is " & FavFood & "."

'Update when the app was last loaded
SaveSetting "AppLogin", "User" & UserNum, "LastLoad", Date
End Sub
