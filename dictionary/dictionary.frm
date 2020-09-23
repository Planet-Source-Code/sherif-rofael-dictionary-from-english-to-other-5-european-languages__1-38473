VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSLATOR BY SHERIF ROFAEL."
   ClientHeight    =   3780
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dictionary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox answer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   1920
      Width           =   3375
   End
   Begin VB.ComboBox lang 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "dictionary.frx":030A
      Left            =   4320
      List            =   "dictionary.frx":031D
      TabIndex        =   2
      Text            =   "CHOOSE LANGUAGE"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Left            =   6120
      Top             =   3120
   End
   Begin VB.TextBox wordd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TRANSLATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ENGLISH"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "TRANSLATE FROM ENGLISH TO :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   4335
   End
   Begin VB.Menu ns 
      Caption         =   "New Search"
   End
   Begin VB.Menu HH 
      Caption         =   "HELP"
   End
   Begin VB.Menu ex 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public K As Integer
'Public proo As String
'german  9600 words
'italian 5160
'french 3257
'portogue  1375
'spanish 7487

Private Sub Command1_Click()
On Error GoTo choosee:
If lang.ListIndex = 0 Then FileName = App.Path & "\FRENCH.txt": Label3.Caption = "FRENCH"
If lang.ListIndex = 1 Then FileName = App.Path & "\GERMAN.txt": Label3.Caption = "GERMAN"
If lang.ListIndex = 2 Then FileName = App.Path & "\ITALIAN.txt": Label3.Caption = "ITALIAN"
If lang.ListIndex = 3 Then FileName = App.Path & "\SPANISH.txt": Label3.Caption = "SPANISH"
If lang.ListIndex = 4 Then FileName = App.Path & "\PORTUGUE.txt": Label3.Caption = "PORTUGUE"
wordd = LCase(wordd)
lenn = FileLen(FileName)
Open FileName For Input As #1
allwords = Input(lenn, #1)
Close #1
wordarray = Split(allwords, vbCrLf)
M = Len(wordd)
On Error GoTo nomatch:
For index = 0 To (UBound(wordarray) + 1)
Call getword(index, word, wordarray)
germanword = LCase$(Left$(word, M))
If germanword = wordd Then
Call getanswer(word, wordd, wordprinted)
'MsgBox "the word is" & wordprinted, , "sherif"
answer.Text = wordprinted
Exit Sub
End If
Next index
nomatch:
MsgBox "NO WORD MATCHS", , "SHERIF": Exit Sub
choosee:
MsgBox "please choose language", , "TRANSLATOR BY SHERIF": Exit Sub
End Sub

Function getword(index, word, wordarray)
' GETTING WORD FROM THE FILE
word = wordarray(index)
End Function


Function getanswer(word, wordd, wordprinted)
entered = Len(wordd)
wordinfile = Len(word)
'word = StrReverse(word)
'Print word
's = InStr(word, " ")
'Print s, word
'word = StrReverse(word)
'z = wordinfile - s
'Print wordinfile, s, z
wordprinted = LCase$(Right$(word, wordinfile - entered - 1))
wordprinted = LTrim(wordprinted)
wordprinted = RTrim(wordprinted)
End Function


Private Sub ex_Click()
End
End Sub

Private Sub HH_Click()
Form2.Visible = True
Unload Me
End Sub

Private Sub lang_Click()
If lang.ListIndex = 0 Then FileName = App.Path & "\FRENCH.txt": Label3.Caption = "FRENCH": MsgBox "i'm sorry psc didn't allow me to upload all my program , please see the file important.txt to download the whole program", , "sherif rofael"
If lang.ListIndex = 1 Then FileName = App.Path & "\GERMAN.txt": Label3.Caption = "GERMAN"
If lang.ListIndex = 2 Then FileName = App.Path & "\ITALIAN.txt": Label3.Caption = "ITALIAN": MsgBox "i'm sorry psc didn't allow me to upload all my program , please see the file important.txt to download the whole program", , "sherif rofael"
If lang.ListIndex = 3 Then FileName = App.Path & "\SPANISH.txt": Label3.Caption = "SPANISH"
If lang.ListIndex = 4 Then FileName = App.Path & "\PORTUGUE.txt": Label3.Caption = "PORTUGUE": MsgBox "i'm sorry psc didn't allow me to upload all my program , please see the file important.txt to download the whole program", , "sherif rofael"
End Sub

Private Sub ns_Click()
wordd.Text = ""
answer.Text = ""
End Sub
