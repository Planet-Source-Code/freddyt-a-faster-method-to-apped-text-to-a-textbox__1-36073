VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Perform Test Again"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   2775
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LoopMin As Long = 0
Private Const LoopMax As Long = 500
'

Private Sub AppendTextbox(txtBx As TextBox, ByVal strText2Append As String)
'This method is much faster.
    'One side effect, the cursor is automatically set to the end of the text.
    'I like this side effect, it saves other un-needed coding :)
    With txtBx
        .SelStart = Len(.Text)
        .SelText = strText2Append
    End With
End Sub
'

Private Sub Command1_Click()
    Form_Load
End Sub
'

Private Sub Form_Load()
    'Try this test with and without filling
    'the text boxes before running.
    Text1.Text = String(1000, 99)
    Text2.Text = String(1000, 99)
    
    Dim iTime(3) As String
    Dim l As Long
    
    iTime(0) = Timer
    For l = LoopMin To LoopMax
        Text1.Text = Text1.Text & String(10, 169)
    Next
    iTime(1) = Timer
    
    'Start next test.
    
    iTime(2) = Timer
    Dim m As Long
    For m = LoopMin To LoopMax
        AppendTextbox Text2, String(10, 169)
    Next
    iTime(3) = Timer
    
    MsgBox "The first loop took " & iTime(1) - iTime(0) & " seconds." & vbNewLine & _
           "The second loop took " & iTime(3) - iTime(2) & " seconds." & vbNewLine
End Sub
