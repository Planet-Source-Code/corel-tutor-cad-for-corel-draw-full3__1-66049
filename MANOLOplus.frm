VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MANOLOplus 
   Caption         =   "CorelCAD plus"
   ClientHeight    =   3396
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4548
   OleObjectBlob   =   "MANOLOplus.frx":0000
   StartUpPosition =   1  'Centrar en propietario
   Tag             =   "                                              MANOLO - "
End
Attribute VB_Name = "MANOLOplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'Dim escala As Double


Private Sub cmARCO_Click()
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
    MANOLOplus.hide
End Sub

Private Sub CMoffset_Click()
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
    ActiveSelection.Flip (cdrFlipVertical)
    Dim s As Shape
    Dim lll As Double
    ActiveDocument.BeginCommandGroup ("Flip V. Sólo Textos")
    For Each s In ActiveSelection.Shapes
        If s.Type = cdrTextShape Then
           With s
            lll = .PositionX
            .Flip (cdrFlipVertical)
            .PositionX = lll
            If .RotationAngle > 0 And .RotationAngle <> 90 Then
               .RotationAngle = 360 - .RotationAngle
            End If
           End With
        End If
    Next s
    ActiveDocument.EndCommandGroup

End Sub

Private Sub CommandButton2_Click()
Dim s As Shape
    Set s = ActiveSelection.Shapes(1)
    If s.Type = cdrCurveShape Then
        Dim st As String
        If ActiveDocument.Unit = cdrCentimeter Then
           st = " cm."
        ElseIf ActiveDocument.Unit = cdrMeter Then
           st = " m."
        Else
           st = " inch."
        End If
        MsgBox "Length of curve : " & vbCrLf & s.Curve.Length & st, , "MANOLO - Corel CAD"
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
    ActiveDocument.BeginCommandGroup ("Desarrollo de cilindro")
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
    ActiveDocument.EndCommandGroup
End Sub

Private Sub CommandButton5_Click()
    Dim s As Shape
    'Dim sr As ShapeRange

    For Each s In ActiveSelection.Shapes
        If s.Type <> cdrTextShape Then
           s.RemoveFromSelection
        End If
    Next
End Sub

Private Sub CommandButton6_Click()
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
    st = InputBox("Ingrese el texto a reemplazar", "MANOLO - Corel CAD")
    If st = "" Then Exit Sub
    For Each s In ActiveSelection.Shapes
        If s.Type = cdrTextShape Then
           s.Text.Range(0, 999) = st
     '      s.Text.Replace "", "ä", False, ReplaceAll:=True
      '     s.Text.BeginEdit
       '    s.Text.Selection.Select
           
    
        End If
    Next
End Sub



Private Sub CommandButton7_Click()
    Dim s As Shape
    'Dim sr As ShapeRange

    For Each s In ActiveSelection.Shapes
        If s.Type = cdrTextShape Then
           s.RemoveFromSelection
        End If
    Next
End Sub

Private Sub coor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then 'enter
       CMenter_Click
       KeyCode = 0
    ElseIf KeyCode = 38 Then 'up
       coor = "0," & Abs(Val(coor))
    ElseIf KeyCode = 37 Then 'left
       coor = -Abs(Val(coor)) & ",0"
    ElseIf KeyCode = 40 Then 'down
       coor = "0," & -Abs(Val(coor))
    ElseIf KeyCode = 39 Then 'right
       coor = Abs(Val(coor)) & ",0"
    End If
    If KeyCode > 36 And KeyCode < 41 Then
       CMenter_Click
       'With coor
       '.SelStart = 0
       '.SelLength = Len(coor)
       'End With
       KeyCode = 0
    End If
    Debug.Print KeyCode
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

Private Sub UserForm_Initialize()
    ActiveDocument.Unit = cdrCentimeter
    Dim cc As Object
    Exit Sub 'English / Spanish
    On Error Resume Next
    For Each cc In MANOLOplus.Controls
'        If TypeOf cc Is CommandButton Then
           If cc.Tag <> "" Then cc.Caption = cc.Tag
 '       End If
    Next
                                                                                                        
                                                                                                        Caption = cmHIDE.Tag & Caption
                                                                                                                                                                                                                      '  Label24 = "jmmisa"
                                                                                                                                                                                                                      '  Label24.ControlTipText = "jmmisa@.hotmail.com"
End Sub
