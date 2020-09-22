Attribute VB_Name = "Math"
Option Explicit

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public StopExecution As Boolean
Public PauseExecution As Boolean
Public CreateScreenShot As Boolean

Public Function IsFunctionValid(ByVal strFunction As String) As Boolean
    On Error GoTo ErrorTrap:
    IsFunctionValid = True
    
    Dim repF As String
    Dim v As Variant
    
    repF = Math.ReplaceParameter(strFunction, "t", "0.0")
    v = Main.Scr.Eval(repF)
    repF = Math.ReplaceParameter(strFunction, "t", "0.5")
    v = Main.Scr.Eval(repF)
    repF = Math.ReplaceParameter(strFunction, "t", "1.0")
    v = Main.Scr.Eval(repF)
    repF = Math.ReplaceParameter(strFunction, "t", "10.0")
    v = Main.Scr.Eval(repF)
    
    Exit Function
ErrorTrap:
    IsFunctionValid = False
End Function

Public Function RenderLissajoux(ByRef pic As PictureBox, _
                                ByVal Fx As String, _
                                ByVal Fy As String, _
                                Optional ByVal bytDarken As Byte = 10, _
                                Optional ByVal dblStart As Double = 0, _
                                Optional ByVal dblStepSize As Double = 1) As Long
    'On Error GoTo ErrorTrap
    StopExecution = False
    PauseExecution = False
    CreateScreenShot = False
    
    Dim arrGrey() As Byte
    Dim x As Long, y As Long
    Dim v As Double, repF As String
    Dim t As Double, ti As Double, S As Long
    Dim W As Double, H As Double
    Dim strImagePath As String
    
    pic.Cls
    W = pic.ScaleWidth / 2.1
    H = pic.ScaleHeight / 2.1
    
    ReDim arrGrey(pic.ScaleWidth - 1, pic.ScaleHeight - 1)
    For x = 0 To UBound(arrGrey, 1)
    For y = 0 To UBound(arrGrey, 2)
        arrGrey(x, y) = 255
    Next
    Next
    
    t = dblStart
    S = 0
    
    Do
        If StopExecution Then Exit Do
        t = t + dblStepSize
        ti = Round(t, 7)
        S = S + 1
        
        repF = Math.ReplaceParameter(Fx, "t", CStr(ti))
        v = Main.Scr.Eval(repF)
        'v = 0.15 * (2 + Sin(t / 500)) * Cos(t) * (1 + Sin(3.01 * t) ^ 2)
        x = W * v + (W + (W * 0.05))
        
        repF = Math.ReplaceParameter(Fy, "t", CStr(ti))
        v = Main.Scr.Eval(repF)
        'v = 0.15 * (2 + Sin(t / 500)) * Sin(t) * (1 + Sin(3.01 * t) ^ 2)
        y = H * v + (H + (H * 0.05))
        
        x = Region(x, 0, UBound(arrGrey, 1))
        y = Region(y, 0, UBound(arrGrey, 2))
        
        If arrGrey(x, y) < bytDarken Then
            arrGrey(x, y) = 0
        Else
            arrGrey(x, y) = arrGrey(x, y) - bytDarken
        End If
        
        pic.PSet (x, y), RGB(arrGrey(x, y), arrGrey(x, y), arrGrey(x, y))
        If CBool(S \ 1000 = S / 1000) Then DoEvents
        If PauseExecution Then
            Load Pause
            Call Pause.DrawThumbNail(arrGrey)
            Pause.Show 1
        End If
        
        If CreateScreenShot Then
            SavePicture pic.Image, CStr(App.Path & "\Temp_Viewport.bmp")
            Main.comDialog.DialogTitle = "Save current graph image..."
            Main.comDialog.Filter = "Bitmap image|*.bmp|"
            Main.comDialog.FileName = "Lissajoux_graph"
            Main.comDialog.ShowSave
            If Len(Main.comDialog.FileName) <> 0 Then
                Call RenderText(pic, 5, 0, "Fx(t); " & Fx, 0)
                Call RenderText(pic, 5, 485, "Fy(t); " & Fy, 0)
                SavePicture pic.Image, CStr(Main.comDialog.FileName)
                DoEvents
                Set Main.Viewport = LoadPicture(App.Path & "\Temp_Viewport.bmp")
            End If
            CreateScreenShot = False
        End If
    Loop
    RenderLissajoux = S
    Exit Function
    
ErrorTrap:
    MsgBox "Error during graph:" & vbNewLine & _
            "Err.Number:" & Err.Number & vbNewLine & _
            "Err.Descr:" & Err.Description & vbNewLine & vbNewLine & _
            "The graph will be aborted...", vbOKOnly Or vbExclamation, "Lissajoux error"
    RenderLissajoux = S
End Function

Public Function ReplaceParameter(ByVal strStream As String, _
                                 ByVal strParameter As String, _
                                 ByVal strReplacement As String) As String
    Dim strNewF As String
    Dim inStrResult As Variant
    Dim i As Long
    i = 1
    strStream = Space(1) & strStream & Space(1)
    strNewF = strStream
    
    Do
        inStrResult = InStr(i, strNewF, strParameter, vbTextCompare)
        If inStrResult = 0 Then
            ReplaceParameter = Trim(strNewF)
            Exit Function
        End If
        
        If IsParameterIsolated(strNewF, inStrResult) Then
            strNewF = Left(strNewF, inStrResult - 1) & strReplacement & Mid(strNewF, inStrResult + 1)
        End If
        i = inStrResult + 1
    Loop
    
End Function

Public Function IsParameterIsolated(ByVal strStream As String, ByVal lngIndex As Long) As Boolean
    IsParameterIsolated = False
    strStream = UCase(strStream)
    Dim charPos As Variant
    
    If lngIndex > 1 Then
        charPos = Asc(Mid(strStream, lngIndex - 1))
        If charPos >= 65 And charPos <= 90 Then Exit Function
    End If
    
    If lngIndex < Len(strStream) Then
        charPos = Asc(Mid(strStream, lngIndex + 1))
        If charPos >= 65 And charPos <= 90 Then Exit Function
    End If
    
    IsParameterIsolated = True
End Function

Public Function RenderText(ByRef pic As PictureBox, _
                     ByVal x As Long, _
                     ByVal y As Long, _
                     ByVal strText As String, _
                     ByVal rgbColour As Long) As Boolean
    Call SetTextColor(pic.hDC, rgbColour)
    Call TextOut(pic.hDC, x, y, strText, Len(strText))
    RenderText = True
End Function

Private Function Random() As Double
    Random = 2 * Rnd - 1
End Function

Private Function NegSQR(ByVal dblIn) As Double
    NegSQR = Sqr(Abs(dblIn))
    If dblIn < 0 Then NegSQR = -NegSQR
End Function

Private Function Region(ByVal dblvalue As Double, ByVal dblMin As Double, ByVal dblMax As Double) As Double
    Region = dblvalue
    If dblvalue < dblMin Then Region = dblMin
    If dblvalue > dblMax Then Region = dblMax
End Function
