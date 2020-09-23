VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Random"
      Height          =   375
      Left            =   5145
      TabIndex        =   2
      Top             =   15
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End Line"
      Height          =   495
      Left            =   6390
      TabIndex        =   1
      Top             =   570
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   6390
      TabIndex        =   0
      Top             =   15
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpx, tmpy As Integer
Dim mainx, mainy As Integer
Private Type pos
    X As Integer
    Y As Integer
End Type
Dim shape1(200) As pos
Dim ctr As Integer
Dim MainCol As Collection
Private Sub Command1_Click()
Dim L As MLine
Dim i As Integer
Dim j As Double
Dim C As Collection
Dim L2 As Collection
Dim L3 As Collection
Dim Ctemp As Collection
Dim CMtemp As Collection
Set Ctemp = New Collection
Set CMtemp = New Collection
Set MainCol = New Collection

For i = 0 To ctr - 2
    Set L2 = New Collection
    L2.Add shape1(i).X, "X1"
    L2.Add shape1(i).Y, "Y1"
    L2.Add shape1(i + 1).X, "X2"
    L2.Add shape1(i + 1).Y, "Y2"
    MainCol.Add L2
    Set L2 = Nothing
Next i

For i = 1 To 10
    Set Ctemp = New Collection
    Set CMtemp = New Collection
    For Each L2 In MainCol
        L.X1 = L2("X1")
        L.Y1 = L2("Y1")
        L.X2 = L2("X2")
        L.Y2 = L2("Y2")
        Set Ctemp = IIf(Check1.Value = 0, CreateLine(L), CreateLine2(L))
        For Each L3 In Ctemp
            CMtemp.Add L3
        Next L3
        Set Ctemp = Nothing
    Next L2
    Set MainCol = Nothing
    Set MainCol = New Collection
    Set MainCol = CMtemp
    Set CMtemp = Nothing

Me.Cls
For Each L2 In MainCol
    Me.Line (L2("X1"), L2("Y1"))-(L2("X2"), L2("Y2"))
Next L2
For j = 1 To 100000: DoEvents: Next j

Next i


End Sub

Private Sub Command2_Click()
Line (tmpx, tmpy)-(mainx, mainy)
shape1(ctr).X = shape1(0).X
shape1(ctr).Y = shape1(0).Y
ctr = ctr + 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form1.Caption = "x = " & X & " " & "y = " & Y & " " & ctr
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (tmpx = 0) And (tmpy = 0) Then
    
        shape1(ctr).X = X
        shape1(ctr).Y = Y
       tmpx = X
    mainx = X
    mainy = Y
    tmpy = Y
End If
Form1.Circle (X, Y), 2
Form1.Line (tmpx, tmpy)-(X, Y)
tmpx = X
tmpy = Y

    shape1(ctr).X = X
    shape1(ctr).Y = Y

    
ctr = ctr + 1

End Sub
