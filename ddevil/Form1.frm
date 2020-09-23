VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pixie() As Boolean
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.CurrentX = X
Me.CurrentY = Y
Dim colon As Long, nochange As Boolean
colon = Me.Point(X, Y)


'deksia
c = colon
If Button = 2 Then
i = X
ReDim pixie((Me.ScaleHeight + 1) * (Me.ScaleWidth + 1))

Do Until c <> colon Or i = Me.ScaleWidth
Me.PSet (i, Y), vbRed
pixie(Y * (Me.ScaleWidth + 1) + i) = True

i = i + 1
c = Me.Point(i, Y)

Loop
'aristera (pasok)
c = colon
i = X
Do Until c <> colon Or i = -1

Me.PSet (i, Y), vbRed
pixie(Y * (Me.ScaleWidth + 1) + i) = True

i = i - 1
c = Me.Point(i, Y)

Loop

'HOTHIKAN

'y thetiko
i = Y + 1
i2 = 0
p = 1
Do Until nochange
nochange = True

For i = 1 To Me.ScaleHeight
DoEvents
For i2 = 0 To Me.ScaleWidth

    If pixie((i - 1) * (Me.ScaleWidth + 1) + i2) = True Then
    
    If Me.Point(i2, i) = colon Then
        Me.PSet (i2, i), vbRed
    '    Debug.Print i2, i
        nochange = False
        pixie(i * (Me.ScaleWidth + 1) + i2) = True
    End If
    
    End If

Next i2



For i2 = 0 To Me.ScaleWidth - 1
    If pixie((i) * (Me.ScaleWidth + 1) + i2) = True Then
  If Me.Point(i2 + 1, i) = colon Then
        Me.PSet (i2 + 1, i), vbRed
      nochange = False
       pixie(i * (Me.ScaleWidth + 1) + i2 + 1) = True
    End If
    End If

Next i2

For i2 = Me.ScaleWidth To 0 Step -1
  If pixie((i) * (Me.ScaleWidth + 1) + i2) = True Then
    If Me.Point(i2 - 1, i) = colon Then
        Me.PSet (i2 - 1, i), vbRed
        nochange = False
        pixie(i * (Me.ScaleWidth + 1) + i2 - 1) = True
    End If
    End If

Next i2

Next i

Me.Refresh

If i = Me.ScaleHeight + 1 Then
'nochange = True
i = i - 1
p = -1
'\/\/\//\/\\/\///\/\ anapodo.
For i = Me.ScaleHeight To 1 Step -1
DoEvents

For i2 = 0 To Me.ScaleWidth
    If pixie(i * (Me.ScaleWidth + 1) + i2) = True Then
    If Me.Point(i2, i - 1) = colon Then
        Me.PSet (i2, i - 1), vbRed
        nochange = False
        pixie((i - 1) * (Me.ScaleWidth + 1) + i2) = True
    End If
    End If

Next i2
For i2 = 0 To Me.ScaleWidth
    If pixie((i - 1) * (Me.ScaleWidth + 1) + i2) = True Then
    If Me.Point(i2 + 1, i - 1) = colon Then
        Me.PSet (i2 + 1, i - 1), vbRed
        nochange = False
        pixie((i - 1) * (Me.ScaleWidth + 1) + i2 + 1) = True
    End If
    End If

Next i2
For i2 = Me.ScaleWidth To 0 Step -1
    If pixie((i - 1) * (Me.ScaleWidth + 1) + i2) = True Then
    If Me.Point(i2 - 1, i - 1) = colon Then
        Me.PSet (i2 - 1, i - 1), vbRed
        nochange = False
        pixie((i - 1) * (Me.ScaleWidth + 1) + i2 - 1) = True
    End If
    End If

Next i2


Next i



End If
'anapoda





Loop



End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = X & " " & Y & " " & Me.Point(X, Y)
If Button = 1 Then
Me.Line (Me.CurrentX, Me.CurrentY)-(X, Y), vbBlack
End If
End Sub
