Attribute VB_Name = "Module1"
Public Type MLine
    X1 As Double
    Y1 As Double
    X2 As Double
    Y2 As Double
End Type
Function CreateLine(L As MLine) As Collection
'Standard mode
'Breaking a line in 2 pieces ---> /\ , and adding it to lines collection
Set CreateLine = New Collection
Dim Ctemp As Collection
Set Ctemp = New Collection

Dim f As Integer
Dim D As Double
Dim X As Double
Dim Y As Double
Dim S As Double
Dim A As Double
Dim aa As Double
Dim bb As Double
Dim cc As Double
Dim xx As Double
Dim yy As Double

X = (L.X1 + L.X2) / 2
Y = (L.Y1 + L.Y2) / 2

If L.Y1 = L.Y2 Then
        xx = X
        yy = Y + Sqr(((L.Y1 - L.Y2) ^ 2) + ((L.X1 - L.X2) ^ 2)) / 4
    Else
        S = (L.X2 - L.X1) / (L.Y1 - L.Y2)
        A = (((L.X1 - L.X2) ^ 2) + ((L.Y1 - L.Y2) ^ 2)) / 16
        aa = (S ^ 2) + 1
        bb = -((2 * X) + (2 * (S ^ 2) * X))
        cc = (X ^ 2) + ((S * X) ^ 2) - A
        D = ((bb ^ 2) - (4 * aa * cc))
    
        xx = ((-bb) + Sqr(D)) / (2 * aa)
        yy = S * (xx - X) + Y
End If

Ctemp.Add L.X1, "X1": Ctemp.Add L.Y1, "Y1"
Ctemp.Add xx, "X2": Ctemp.Add yy, "Y2"
CreateLine.Add Ctemp

Set Ctemp = Nothing
Set Ctemp = New Collection

Ctemp.Add xx, "X1": Ctemp.Add yy, "Y1"
Ctemp.Add L.X2, "X2": Ctemp.Add L.Y2, "Y2"

CreateLine.Add Ctemp

Set Ctemp = Nothing
End Function
Function CreateLine2(L As MLine) As Collection
'Random mode
Set CreateLine2 = New Collection
Dim Ctemp As Collection
Set Ctemp = New Collection

Dim f As Integer
Dim D As Double
Dim X As Double
Dim Y As Double
Dim S As Double
Dim A As Double
Dim aa As Double
Dim bb As Double
Dim cc As Double
Dim xx As Double
Dim yy As Double
Randomize (Timer)

X = (L.X1 + L.X2) / 2
Y = (L.Y1 + L.Y2) / 2

If L.Y1 = L.Y2 Then
        xx = X
        yy = Y + Sqr(((L.Y1 - L.Y2) ^ 2) + ((L.X1 - L.X2) ^ 2)) / 4
    Else
        S = (L.X2 - L.X1) / (L.Y1 - L.Y2)
        A = (((L.X1 - L.X2) ^ 2) + ((L.Y1 - L.Y2) ^ 2)) / 16
        aa = (S ^ 2) + 1
        bb = -((2 * X) + (2 * (S ^ 2) * X))
        cc = (X ^ 2) + ((S * X) ^ 2) - A
        D = ((bb ^ 2) - (4 * aa * cc))
    
        f = (Rnd * 10) Mod 2
        f = IIf(f = 0, 1, -1)
        xx = ((-bb) + (f * Sqr(D))) / (2 * aa)
        yy = S * (xx - X) + Y
End If
'creating & adding new lines
Ctemp.Add L.X1, "X1": Ctemp.Add L.Y1, "Y1"
Ctemp.Add xx, "X2": Ctemp.Add yy, "Y2"
CreateLine2.Add Ctemp

Set Ctemp = Nothing
Set Ctemp = New Collection

Ctemp.Add xx, "X1": Ctemp.Add yy, "Y1"
Ctemp.Add L.X2, "X2": Ctemp.Add L.Y2, "Y2"

CreateLine2.Add Ctemp

Set Ctemp = Nothing
End Function
