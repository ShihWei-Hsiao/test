VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   5160
   StartUpPosition =   3  '系統預設值
   Begin VB.Label Label1 
      Caption         =   "P/p to ping, F12 to exit"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'123
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 123 Then               '按F12結束
    Close #1
    End
ElseIf KeyCode = 112 Or KeyCode = 80 Then
    Call Command1_Click
End If
End Sub

'Private Sub Form_KeyPress(keyascii As Integer)
'
''Open "D:\practice\keylog\" & Format(Now(), "yyyy-mm-dd") & ".txt" For Append As #1
''Print #1, Format(Now(), "yyyy-mm-dd_H:M:S") & " " & Chr(keyascii)
''Close #1
'
'End Sub


Private Sub Command1_Click()

'Shell "cmd.exe /c ipconfig -all " & "> D:\practice\keylog\log.txt" '寫ipconfig到log.txt
'Shell "cmd.exe /c start " & "D:\practice\keylog\log.txt"   打開log.txt

Open "D:\practice\keylog\ping.txt" For Append As #1
Print #1, "============================================================="
Print #1, Format(Now(), "yyyy-mm-dd_H:M:S")
Close #1

Shell "cmd.exe /c ping 127.0.0.1" & " >> D:\practice\keylog\ping.txt"

Delay (4)

Open "D:\practice\keylog\ping.txt" For Append As #2
Print #2, "============================================================="
Close #2

End Sub


Private Sub Delay(ByVal Sec As Single)
Dim sgnThisTime As Single, sgnCount As Single
sgnThisTime = Timer
Do While sgnCount < Sec
    sgnCount = Timer - sgnThisTime
    DoEvents
Loop
End Sub
