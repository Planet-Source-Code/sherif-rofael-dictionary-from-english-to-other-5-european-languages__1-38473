VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABOUT THE PROGRAM !"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "MAILTO:ya3amo@hotmail.com"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "website:http://vbsherif.members.easyspace.com"
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   6120
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "                     BY : SHERIF ROFAEL"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "THANKS FOR USING MY PROGRAM."
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":030A
      ForeColor       =   &H00400040&
      Height          =   3855
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M As Integer

Private Sub Command1_Click()
Select Case M
Case 1
Form1.Visible = True
Unload Me
Case 0
Label1.Caption = " german  9600 words ,italian 5160 WORDS,french 3257 WORDS,portogue  1375 WORDS, spanish 7487 WORDS , SO I THINK IT IS A GOOD DICTIANARY FOR BOTH SPANISH AND GERMAN LANGUAGES , SORRY IF IT IS NOT SO WELL , IT'S MY FIRST TRY IN MAKING A DICTIONARY IN 5 LANGUAGES AT THE SAME TIME "
Command1.Caption = " OKAY "
M = M + 1
End Select
End Sub

