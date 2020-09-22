VERSION 5.00
Begin VB.Form frmCombos 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get Combinations"
      Default         =   -1  'True
      Height          =   435
      Left            =   2520
      TabIndex        =   2
      Top             =   540
      Width           =   1995
   End
   Begin VB.ListBox lstOut 
      Height          =   2595
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   2355
   End
   Begin VB.TextBox txtIn 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblCount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1995
   End
End
Attribute VB_Name = "frmCombos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGet_Click()
    If txtIn.Text = "" Then Exit Sub
    lstOut.Clear
    GetCombos txtIn.Text
    'the conversion to Hex and back is a cheap way of avoiding
    'negative numbers when the count goes over integer range
    '8 characters is 40,320 combinations, which is
    '9D80 in hex, which gets interpreted otherwise as -25,216
    'By converting to hex and specifying the "&" on the end
    'we get a long value (CLng doesn't do anything to it, I tried)
    lblCount.Caption = "Total Combinations: " & _
        Format$(Val("&H" & Hex$(lstOut.ListCount) & "&"), "#,##0")
End Sub

'recursive routine
'This is really hard to describe.  I'd reccomend setting a breakpoint
'at the very beginning, setting a watch for strIn and strFixed, and
'stepping through it with a small string like "123" to get the idea.
Private Sub GetCombos(strIn As String, Optional strFixed As String)
Dim iloop As Integer
    If Len(strIn) <> 1 Then
        'Send through the routine again, tacking each character
        '(one at a time) to the end of what comes in as 'fixed'
        'and sending the remainder of the string to be processed.
        For iloop = 1 To Len(strIn)
            GetCombos _
                Left$(strIn, iloop - 1) & Mid$(strIn, iloop + 1), _
                strFixed & Mid$(strIn, iloop, 1)
        Next iloop
    Else
        'When there's no more to be split up, the 'fixed'
        'part has all the rest.  We can add it to the list.
        lstOut.AddItem strFixed & strIn
    End If
End Sub
