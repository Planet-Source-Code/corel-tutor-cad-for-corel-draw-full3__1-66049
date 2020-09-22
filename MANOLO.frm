VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MANOLO 
   Caption         =   " CorelCAD"
   ClientHeight    =   6252
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7056
   OleObjectBlob   =   "MANOLO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
   Tag             =   "                               jmmisa"
End
Attribute VB_Name = "MANOLO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Engl As Boolean
'Dim escala As Double


Private Sub cmARCO_Click()
    If OptionButton1.Value Then ActiveDocument.Unit = cdrCentimeter
    If OptionButton2.Value Then ActiveDocument.Unit = cdrMeter
    If OptionButton3.Value Then ActiveDocument.Unit = cdrInch
   ' CrearArco 0
   ' Exit Sub

    Dim s As Shape
    Dim lr As Layer
    Set lr = ActiveLayer
    ActiveDocument.DrawingOriginX = -ActivePage.SizeWidth / 2
    ActiveDocument.DrawingOriginY = -ActivePage.SizeHeight / 2
 '   lr.CreateRectangle 0, 0, 3, 2
  ' ActiveDocument.ActiveWindow.Left
   ' Set s = lr.CreateEllipse(0, 0, 4, 2, 0#, 180#, False)
    'Set s = lr.CreateEllipse2(0, 0, 5, 5, 45#, 130#, False)
   
    Dim r As Double
    Dim b As Double
    Dim p As Double
    p = Val(TXp)
    b = Val(TXb) / 2
    r = ((b * b / p) + p) / 2
    LBr = Format(r, "0.####")
    Dim angulo As Double
    Dim anguloi As Double
   'angulo = 4 * Atn(p / b) * 180 / 3.141592654
    angulo = 720 * Atn(p / b) / 3.14159265358979
    anguloi = 360 * Atn(p / b) / 3.14159265358979 - 90
   ' LBl = r * angulo * 3.141592654 / 180
    LBa = Format(angulo, "0.####")
    lbArea = Format(((3.141592654 * (r ^ 2) * angulo) / 360) - (b * (r - p)), "0.####")
    LBl = Format(r * 4 * Atn(p / b), "0.####")
' angulo = angulo / 2 - 90    'angulo inicial
 '   r = r / escala
  '  b = b / escala
   ' p = p / escala
    Set s = lr.CreateEllipse2(0, 0, r, r, 180 - anguloi, anguloi, False)
    If p > r Then
       s.SetSize r * 2, p
       s.SetPosition -r, r
    Else
       s.SetSize b * 2, p
       s.SetPosition -b, r
    End If
   ' s.PositionX = -b
 '   s.ConvertToCurves
  '  Dim ppp As SubPath
   ' Set ppp = s.Curve.Subpaths(1)
    'ppp.AppendLineSegment -b, r - p
    'ppp.AppendLineSegment b, r - p, True
   ' Dim s3 As Shape
    'Set s3 = lr.CreateEllipse2(0, 0, r, r, 180 - anguloi, anguloi, False)
    
    Dim s2 As Shape
    Set s2 = lr.CreateLineSegment(-b, r - p, b, r - p)
    Set s2 = lr.CreateLineSegment(0, r - p, 0, r)
End Sub



Private Sub CMenter_Click()
    If OptionButton1.Value Then ActiveDocument.Unit = cdrCentimeter
    If OptionButton2.Value Then ActiveDocument.Unit = cdrMeter
    If OptionButton3.Value Then ActiveDocument.Unit = cdrInch
    'Dim s As Shape
    With coor
    .SelStart = 0
    .SelLength = Len(coor)
    End With
    
    If LCase(Trim(coor)) = "u" Then
       ActiveDocument.Undo
       Exit Sub
    End If
    
    Dim relativo As Boolean
    relativo = cmr.Value
    Dim nvoTrayecto As Boolean
    nvoTrayecto = CMtra.Value
    
    Dim ppp As SubPath
    Dim crv As Curve
    Dim x As Double
    Dim y As Double
    If ActiveShape Is Nothing Then
       nvoTrayecto = True
       Set crv = CreateCurve(ActiveDocument)
       x = TXX
       y = TXY
       Set ppp = crv.CreateSubPath(x, y)
    Else
       If ActiveShape.Type <> cdrCurveShape Then Exit Sub
       'Dim s As Shape
       'Set s = ActiveShape
       'Dim ppp As SubPath
       If nvoTrayecto Then
          Dim pppp As SubPath
          Set pppp = ActiveShape.Curve.Subpaths(1)
          CMtra.Value = False
          If relativo Then
             If CMini.Value Then
                x = pppp.StartNode.PositionX
                y = pppp.StartNode.PositionY
             Else
                x = pppp.EndNode.PositionX
                y = pppp.EndNode.PositionY
             End If
          End If
          Set crv = CreateCurve(ActiveDocument)
          Set ppp = crv.CreateSubPath(x, y)
       Else
          Set ppp = ActiveShape.Curve.Subpaths(1)
          If LCase(Trim(coor)) = "c" Then
             ppp.Closed = True
             Exit Sub
          End If
          If relativo Then
             If CMini.Value Then
                x = ppp.StartNode.PositionX
                y = ppp.StartNode.PositionY
             Else
                x = ppp.EndNode.PositionX
                y = ppp.EndNode.PositionY
             End If
          End If
       End If
    End If
       'Dim sss As Segment
'       x = 0
'       y = 0
       Dim i As Integer
       i = InStr(1, coor, ",")
       If i = 0 Then
          i = InStr(1, coor, "<")
          If i = 0 Then
             MsgBox "coordenada mal escrita"
             Exit Sub
          Else
             Dim l As Double
             Dim angulo As Double
             l = Left(coor, i - 1)
             angulo = Mid(coor, i + 1)
             If angulo < 0 Then angulo = 360 + angulo
             angulo = angulo * 3.14159265358979 / 180 'angulo en radianes
             x = x + Cos(angulo) * l
             y = y + Sin(angulo) * l
             'If angulo < 90 And angulo > 0 Then angulo = -angulo
             'If angulo > 90 And angulo < 180 Then angulo = 180 - angulo
             'If angulo > 180 And angulo < 270 Then angulo = 270 - angulo
          End If
       Else
          x = x + Val(Left(coor, i - 1))
          y = y + Val(Mid(coor, i + 1))
       End If
       
 '      If relativo Then
 '         If ActiveShape Is Nothing Then
 '            x = x + TXX
 '            y = y + TXY
 '         Else
 '            If CMini.Value Then
 '               x = x + ppp.StartNode.PositionX
 '               y = y + ppp.StartNode.PositionY
 '            Else
 '               x = x + ppp.EndNode.PositionX
 '               y = y + ppp.EndNode.PositionY
 '            End If
 '        End If
 '      End If
       ppp.AppendLineSegment x, y, CMini.Value
       If nvoTrayecto Then
          Dim s As Shape
          Set s = ActiveLayer.CreateCurve(crv)
       End If
   ' End If

End Sub

Private Sub cmHIDE_Click()
    MANOLO.hide
End Sub

Private Sub CMoffset_Click()
    If OptionButton1.Value Then ActiveDocument.Unit = cdrCentimeter
    If OptionButton2.Value Then ActiveDocument.Unit = cdrMeter
    If OptionButton3.Value Then ActiveDocument.Unit = cdrInch
'CrearArco Val(TXoff)
'Exit Sub
    Dim s As Shape
    Dim lr As Layer
    Set lr = ActiveLayer
    ActiveDocument.DrawingOriginX = -ActivePage.SizeWidth / 2
    ActiveDocument.DrawingOriginY = -ActivePage.SizeHeight / 2
 '   lr.CreateRectangle 0, 0, 3, 2
  ' ActiveDocument.ActiveWindow.Left
   ' Set s = lr.CreateEllipse(0, 0, 4, 2, 0#, 180#, False)
    'Set s = lr.CreateEllipse2(0, 0, 5, 5, 45#, 130#, False)
    Dim OFFSET As Double
    OFFSET = Val(TXoff)
    Dim r As Double
    Dim b As Double
    Dim p As Double
    p = Val(TXp)
    b = Val(TXb) / 2
    r = ((b * b / p) + p) / 2
    LBr = Format(r, "0.####")
    Dim angulo As Double
   'angulo = 4 * Atn(p / b) * 180 / 3.141592654
    angulo = 720 * Atn(p / b) / 3.14159265358979
   ' LBl = r * angulo * 3.141592654 / 180
    LBa = Format(angulo, "0.####")
    lbArea = Format(((3.14159265358979 * (r ^ 2) * angulo) / 360) - (b * (r - p)), "0.####")
    LBl = Format(r * 4 * Atn(p / b), "0.####")
'If offset = 0 Then
'    angulo = angulo / 2 - 90 'angulo inicial
'    Set s = lr.CreateEllipse2(0, 0, r / 2.54, r / 2.54, 180 - angulo, angulo, False)
'Else
    p = p + OFFSET
    r = r + OFFSET
    b = (p * (2 * r - p)) ^ (1 / 2)
    If b <= 0 Then
       MsgBox "ERROR B MENOR A 0"
       Exit Sub
    End If
    LBr2 = Format(r, "0.####")
  ' angulo = 4 * Atn(p / b) * 180 / 3.141592654
    angulo = 720 * Atn(p / b) / 3.14159265358979
    LBb2 = Format(b * 2, "0.####")
    LBl2 = Format(r * 4 * Atn(p / b), "0.####")
    angulo = angulo / 2 - 90 'angulo inicial
    Set s = lr.CreateEllipse2(0, 0, r, r, 180 - angulo, angulo, False)
    If p > r Then
       s.SetSize r * 2, p
       s.SetPosition -r, r
    Else
       s.SetSize b * 2, p
       s.SetPosition -b, r
    End If

'End If
End Sub




Private Sub cmref_Click()
    Dim s As Shape
    Dim lll As Double
    If ActiveSelection.Shapes.Count = 0 Then GoTo noselection
    ActiveDocument.BeginCommandGroup ("Flip V. Sólo Textos")
    ActiveSelection.Flip (cdrFlipVertical)
    For Each s In ActiveSelection.Shapes
        If s.Type = cdrTextShape Then
           With s
            lll = .PositionY
            .Flip (cdrFlipVertical)
            .PositionY = lll
            If .RotationAngle > 0 And .RotationAngle <> 90 Then
               .RotationAngle = 360 - .RotationAngle
            End If
           End With
        End If
    Next s
    ActiveDocument.EndCommandGroup
    
    Exit Sub
noselection:
    If Engl Then
       MsgBox "Make a selection"
    Else
       MsgBox "Haga una selección"
    End If
End Sub

Private Sub CMrefh_Click()
'   If ActiveSelection.Shapes.Count = 0 Then GoTo noselection
    Dim s As Shape
    Dim lll As Double
    ActiveDocument.BeginCommandGroup ("Flip H. Sólo Textos")
    ActiveSelection.Flip (cdrFlipHorizontal)
    For Each s In ActiveSelection.Shapes
        If s.Type = cdrTextShape Then
           With s
            lll = .PositionX
            .Flip (cdrFlipHorizontal)
            .PositionX = lll
            If .RotationAngle > 0 And .RotationAngle <> 90 Then
               .RotationAngle = 360 - .RotationAngle
            End If
           End With
        End If
    Next s
    ActiveDocument.EndCommandGroup
    Exit Sub
noselection:
    If Engl Then
       MsgBox "Make a selection"
    Else
       MsgBox "Haga una selección"
    End If
End Sub

Private Sub CommandButton11_Click()
    If CommandButton11.Caption = ">>" Then
       CommandButton11.Caption = "<<"
       Width = 356.4
    Else
       CommandButton11.Caption = ">>"
       Width = 156.6
    End If
    
End Sub


Private Sub Cmcerrar_Click()
    Dim s As Shape
'    Dim lll As Double
    ActiveDocument.BeginCommandGroup ("Cerrar curvas")
    For Each s In ActiveSelection.Shapes
        If s.Type = cdrCurveShape Then
           'With s
             s.Curve.Closed = True
           'End With
        End If
    Next s
    ActiveDocument.EndCommandGroup
End Sub

Private Sub CommandButton13_Click()
Me.hide
     MANOLOplus.Show
     
'On Error GoTo oops
'    MANOLOplus.Show
'    Exit Sub
'oops:
'    MsgBox "Desarrollo de Cilindro Inclinado comming soon" & vbCr & "Inclined Cone development, comming soon", vbOKCancel, "Corel-CAD"
  '  Me.hide
End Sub

Private Sub CommandButton2_Click()
Dim s As Shape
    If ActiveSelection.Shapes.Count = 0 Then GoTo noselection
    Set s = ActiveSelection.Shapes(1)
    If s.Type = cdrCurveShape Then
        Dim st As String
        If ActiveDocument.Unit = cdrCentimeter Then
           st = " cm."
        ElseIf ActiveDocument.Unit = cdrMeter Then
           st = " m."
        ElseIf ActiveDocument.Unit = cdrInch Then
           st = " inch."
        Else
           st = " "
        End If
        If Engl Then
           MsgBox "Length of curve : " & vbCrLf & s.Curve.Length & st, , "MANOLO - Corel CAD"
        Else
           MsgBox "Largo de la curva : " & vbCrLf & s.Curve.Length & st, , "MANOLO - Corel CAD"
        End If
    Else
        If Engl Then
           MsgBox "Only curve objects allowed"
        Else
           MsgBox "Solo son válidos objetos en curvas"
        End If
    End If

'   If ActiveSelection.Shapes.Count = 0 Then GoTo noselection
    Exit Sub
noselection:
    If Engl Then
       MsgBox "Select a curve"
    Else
       MsgBox "Seleccione una curva"
    End If
End Sub


Private Sub CommandButton3_Click()
    'ActiveSelection.Combine
    Dim s As Shape
    
    Dim n As Integer
    Dim i As Integer
    Dim ii As Integer
    n = ActiveSelection.Shapes.Count
Do
    'For Each s In ActiveSelection.Shapes
    If ActiveSelection.Shapes(i).Curve.Nodes(0).PositionX = ActiveSelection.Shapes(ii).Curve.Nodes(0).PositionX And ActiveSelection.Shapes(i).Curve.Nodes(0).PositionY = ActiveSelection.Shapes(ii).Curve.Nodes(0).PositionY Then
       'si el punto inicial de i y el punto inicial de ii son iguales
       i = 0
       ii = 1
ElseIf ActiveSelection.Shapes(i).Curve.Nodes(ActiveSelection.Shapes(i).Curve.Nodes.Count).PositionX = ActiveSelection.Shapes(ii).Curve.Nodes(0).PositionX And ActiveSelection.Shapes(i).Curve.Nodes(0).PositionY = ActiveSelection.Shapes(ii).Curve.Nodes(0).PositionY Then
       'si el punto final de i y el punto inicial de ii son iguales
       i = 0
       ii = 1
ElseIf ActiveSelection.Shapes(i).Curve.Nodes(0).PositionX = ActiveSelection.Shapes(ii).Curve.Nodes(0).PositionX And ActiveSelection.Shapes(i).Curve.Nodes(0).PositionY = ActiveSelection.Shapes(ii).Curve.Nodes(0).PositionY Then
       'si el punto inicial de i y el punto final de ii son iguales
       i = 0
       ii = 1
ElseIf ActiveSelection.Shapes(i).Curve.Nodes(ActiveSelection.Shapes(i).Curve.Nodes.Count).PositionX = ActiveSelection.Shapes(ii).Curve.Nodes(0).PositionX And ActiveSelection.Shapes(i).Curve.Nodes(0).PositionY = ActiveSelection.Shapes(ii).Curve.Nodes(0).PositionY Then
       'si el punto final de i y el punto final de ii son iguales
       i = 0
       ii = 1
    End If
            
Loop While ActiveSelection.Shapes.Count > 1
    Next
End Sub

Private Sub CommandButton4_Click()
    Dim lr As Layer
    Set lr = ActiveLayer
    'Dim d As Long
    Dim r As Long
    r = Val(txd) / 2
    'd = Val(txd)
    Dim sh As Shape
    Set sh = lr.CreateEllipse(0, r * 2, r * 2, 0)
    'lr.CreateLineSegment 0, r, r * 2, r
    Dim p As Double
    p = 2 * 3.14159265358979 * r / 2
    lbcodo = Format(p * 2, "0.####")
    Dim s As Double
    lr.CreateLineSegment 0, 0, p, 0
 '   lr.CreateLineSegment p, 0, p * 2, 0
    lr.CreateLineSegment 0, -r, p, -r
    Dim angulo As Double
    angulo = 90 - Val(txa) / 2
    lr.CreateLineSegment 0, -r * Tan(angulo * 3.14159265358979 / 180), 0, -r * Tan(angulo * 3.14159265358979 / 180)
   ' lr.CreateLineSegment p, -r, p * 2, -r
    'h = 2 * r * Tan(angulo * 3.14159265358979 / 180)
    lr.CreateLineSegment 0, 2 * r, 2 * r, 2 * r + 2 * r * Tan(angulo * 3.14159265358979 / 180)
    lr.CreateLineSegment 0, 2 * r, 2 * r, 2 * r
    lr.CreateLineSegment 0, 2 * r, 2 * r * Cos(2 * angulo * 3.14159265358979 / 180), 2 * r + 2 * r * Sin(2 * angulo * 3.14159265358979 / 180)
    lr.CreateLineSegment 2 * r * Cos(2 * angulo * 3.14159265358979 / 180), 2 * r + 2 * r * Sin(2 * angulo * 3.14159265358979 / 180), 2 * r, 2 * r + 2 * r * Tan(angulo * 3.14159265358979 / 180)
    lr.CreateLineSegment 2 * r, 2 * r, 2 * r, 2 * r + 2 * r * Tan(angulo * 3.14159265358979 / 180)

    Dim a As Double
    Dim x As Double
    Dim h As Double
    Dim i As Integer
    Dim ii As Integer
    ii = Val(txpre) * 5
    If ii < 2 Then Exit Sub
    Dim ppp As SubPath
    Dim crv As Curve
    Set crv = CreateCurve(ActiveDocument)
    Dim inicio As Integer
    inicio = Val(txi)
    Dim sh2 As Shape
    
    For i = 0 To ii * 2
        s = p / ii * i
        a = inicio - 180 / ii * i
        x = r * Cos(a * 3.14159265358979 / 180)
        'h = (r * r - x * x) ^ (1 / 2)
        h = x * Tan(angulo * 3.14159265358979 / 180)
        If i = 0 Then
           Set ppp = crv.CreateSubPath(0, h)
           Set sh = ActiveLayer.CreateCurve(crv)
        Else
           ppp.AppendLineSegment s, h
        End If
        If i Mod 5 = 0 Then
           If h <> 0 Then lr.CreateLineSegment s, 0, s, h
        End If
    Next
    Set sh = ActiveLayer.CreateCurve(crv)
End Sub

Private Sub cmtxt_Click()
    Dim s As Shape
    'Dim sr As ShapeRange

    For Each s In ActiveSelection.Shapes
        If s.Type <> cdrTextShape Then
           s.RemoveFromSelection
        End If
    Next
End Sub

Private Sub Cmrepl_Click()
    If ActiveSelection.Shapes.Count = 0 Then GoTo noselection
  '  Dim d As Document
  '  Dim s As Shape
  '  Dim t As Text
 '   Set d = CreateDocument
 '   Set s = d.ActiveLayer.CreateParagraphText(3, 3, 7, 7, "This is a test.")
 '   Set t = s.Text
    
    'Exit Sub
    Dim s As Shape
    'Dim sr As ShapeRange
    Dim st As String
    If Engl Then
       st = InputBox("Input text to replace", "MANOLO - Corel CAD")
    Else
       st = InputBox("Ingrese el texto a reemplazar", "MANOLO - Corel CAD")
    End If
    If st = "" Then Exit Sub
    For Each s In ActiveSelection.Shapes
        If s.Type = cdrTextShape Then
           s.Text.Range(0, 999) = st
     '      s.Text.Replace "", "ä", False, ReplaceAll:=True
      '     s.Text.BeginEdit
       '    s.Text.Selection.Select
           
    
        End If
    Next
    
    Exit Sub
noselection:
    MsgBox "Haga una selección que contenga varios objetos de texto"
End Sub



Private Sub cmnotxt_Click()
    Dim s As Shape
    'Dim sr As ShapeRange

    For Each s In ActiveSelection.Shapes
        If s.Type = cdrTextShape Then
           s.RemoveFromSelection
        End If
    Next
End Sub

Private Sub CommandButton7_Click()

End Sub

Private Sub Cmdiv_Click()
    If ActiveSelection.Shapes.Count = 0 Then GoTo noselection

    Dim ii As String
    If Engl Then
       ii = Val(InputBox("Input number of segments to divide"))
    Else
       ii = Val(InputBox("Ingrese el número de segmentos a dividir"))
    End If
    If ii < 2 Then Exit Sub
    If ii > 1000 Then Exit Sub
    
    Dim s As Shape
    Set s = ActiveSelection.Shapes(1)
    If s.Type = cdrCurveShape Then
       Dim p As SubPath
       Set p = s.Curve.Subpaths(1)
       Dim j As Integer
       ActiveDocument.BeginCommandGroup ("Dividir curva en " & ii)
       For j = 2 To ii
          'p.BreakApartAt 1 / ii * (j - 1)
           p.AddNodeAt 1 / ii * (j - 1)
       Next
       ActiveDocument.EndCommandGroup
       'p.BreakApartAt 1 / ii
    End If
    
    Exit Sub
noselection:
    MsgBox "Seleccione una curva"
End Sub

Private Sub Cmdivx_Click()
    If OptionButton1.Value Then ActiveDocument.Unit = cdrCentimeter
    If OptionButton2.Value Then ActiveDocument.Unit = cdrMeter
    If OptionButton3.Value Then ActiveDocument.Unit = cdrInch
    Dim st As String
    If OptionButton1.Value Then st = "cm."
    If OptionButton2.Value Then st = "m."
    If OptionButton3.Value Then st = "inch."
    
    If Engl Then
       st = "Input position to divide curve (" & st & "), positive value to start from beggining node, negative value to start from lastnode"
       If Val(cmdivx.Tag) <> 0 Then st = st & "; Last value = " & cmdivx.Tag
    Else
       st = "Ingrese la posición donde desea dividir la curva (" & st & "), si el valor es positivo se mide a partir del inicio, si es negativo se mide a partir del final"
       If Val(cmdivx.Tag) <> 0 Then st = st & "; último valor = " & cmdivx.Tag
    End If

    Dim ii As String
    ii = Val(InputBox(st))
    If ii = 0 Then Exit Sub
    cmdivx.Tag = ii
    
    Dim s As Shape
    Set s = ActiveSelection.Shapes(1)
    
'    If ii >= s.Curve.Length Then Exit Sub
    If Abs(ii) >= s.Curve.Length Then Exit Sub
    If ii < 0 Then ii = s.Curve.Length + ii
    If s.Type = cdrCurveShape Then
       Dim p As SubPath
       Set p = s.Curve.Subpaths(1)
      ' Dim j As Integer
       ActiveDocument.BeginCommandGroup ("Dividir curva en " & ii)
'       For j = 2 To ii
          'p.BreakApartAt 1 / ii * (j - 1)
           p.AddNodeAt ii, cdrAbsoluteSegmentOffset
 '      Next
       ActiveDocument.EndCommandGroup
       'p.BreakApartAt 1 / ii
    End If
End Sub

Private Sub coor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim st As String
If KeyCode = 13 Then 'enter
    CMenter_Click
    KeyCode = 0
    Exit Sub
ElseIf KeyCode > 36 And KeyCode < 41 Then
    st = coor
    If Left(st, 2) = "0," Then st = Mid(st, 3)
    If KeyCode = 38 Then 'up
       coor = "0," & Abs(Val(st))
    ElseIf KeyCode = 37 Then 'left
       coor = -Abs(Val(st)) & ",0"
    ElseIf KeyCode = 40 Then 'down
       coor = "0," & -Abs(Val(st))
    Else 'If KeyCode = 39 Then 'right
       coor = Abs(Val(st)) & ",0"
    End If
'    If KeyCode > 36 And KeyCode < 41 Then
       CMenter_Click
       'With coor
       '.SelStart = 0
       '.SelLength = Len(coor)
       'End With
       KeyCode = 0
End If
    'End If
    'Debug.Print KeyCode
End Sub


Private Sub Label30_Click()

End Sub

Private Sub Label31_Click()

End Sub

Private Sub OptionButton1_Click()
    ActiveDocument.Unit = cdrCentimeter
   ' escala = 2.54
End Sub

Private Sub OptionButton2_Click()
    ActiveDocument.Unit = cdrMeter
    
   ' escala = 0.0254
End Sub

Private Sub OptionButton3_Click()
    ActiveDocument.Unit = cdrInch
   ' escala = 1
End Sub

Private Sub repetir_Click()
    'primero soldar si está en partes
    Dim i As Integer
    i = ActiveSelection.Shapes.Count
    If i > 1 Then
       'ActiveSelection.Combine
       MsgBox "No se puede repetir mas de un objeto"
       Exit Sub
    End If
    If ActiveShape.Curve.Subpaths.Count > 1 Then
       MsgBox "La curva tiene mas de un tramo"
       Exit Sub
    End If
    Dim ppp As SubPath
    ppp = ActiveShape.Curve.Subpaths(1)
    Dim sss As Segment
    Dim eslinea As Boolean
    sss = ActiveShape.Curve.Segments(1)
    If sss.Type = cdrLineSegment Then
       eslinea = True
    Else
       eslinea = False
    End If
    Dim nseg As Integer
    nseg = ppp.Segments.Count
    Dim i As Integer
    For i = 1 To nseg
    
    Next
    ppp.AppendLineSegment x, y, True
 '   ppp.AppendCurveSegment x,y,
    
    
    
    
    Set s = ActiveSelection.Shapes(1)
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TXC1_Change()
    LBH = (TXC1 * TXC1 + TXC2 * TXC2) ^ (1 / 2)
    If Val(TXC1) < Val(TXH) And Val(TXC1) > 0 And Val(TXH) > 0 Then
       LBC2 = (TXH * TXH - TXC1 * TXC1) ^ (1 / 2)
    Else
       LBC2 = 0
    End If
End Sub

Private Sub TXC2_Change()
    LBH = (TXC1 * TXC1 + TXC2 * TXC2) ^ (1 / 2)
End Sub

Private Sub TXH_Change()
    If Val(TXC1) < Val(TXH) And Val(TXC1) > 0 And Val(TXH) > 0 Then
       LBC2 = (TXH * TXH - TXC1 * TXC1) ^ (1 / 2)
    Else
       LBC2 = 0
    End If
End Sub

Private Sub TXoff_Change()

End Sub

Private Sub UserForm_Initialize()
    ActiveDocument.Unit = cdrCentimeter
    Width = 156.6
    Engl = False
                                                                                                                                                                                         Caption = LTrim(Me.Tag) & Caption
    Exit Sub 'Spanish/English
    Engl = True
    Dim cc As Object
    On Error Resume Next
    For Each cc In MANOLO.Controls
'        If TypeOf cc Is CommandButton Then
           If cc.Tag <> "" Then cc.Caption = cc.Tag
 '       End If
    Next
    cmref.ControlTipText = "In a selection, mirror all objects except the text ones"
    CMrefh.ControlTipText = cmref.ControlTipText
    cmtxt.ControlTipText = "In a selection, deselect no-text objects"
    cmnotxt.ControlTipText = "In a selection, deselect text objects"
    cmrepl.ControlTipText = "In a selection, replace all text-object-captions at the same time"
    cmdiv.ControlTipText = "Divide a curve in 'n' equal parts"
    cmdivx.ControlTipText = "Divide a curve at 'x' from the beginning/end"
    cmcerrar.ControlTipText = "Close all objects in a selection"
    coor.ControlTipText = "10,5   10<60   c   u   10[arrow]"
End Sub
