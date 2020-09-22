Attribute VB_Name = "Interface"
Option Explicit

Public Function HardStoreFunctions() As Boolean
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(App.Path & "\Functions.dat", True)
    file.WriteLine Main.txtXFormula.Text
    file.WriteLine Main.txtYFormula.Text
    file.Close
    Set file = Nothing
    Set fso = Nothing
    HardStoreFunctions = True
End Function

Public Function RetrieveFunctions() As Boolean
    Dim fso As Object
    Dim file As Object
    Dim Fx, Fy
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(App.Path & "\Functions.dat") Then
        Fx = ""
        Fy = ""
    Else
        Set file = fso.OpenTextFile(App.Path & "\Functions.dat", 1)
        Fx = file.ReadLine
        Fy = file.ReadLine
        file.Close
    End If
    
    Set file = Nothing
    Set fso = Nothing
    If Fx = "" Then Fx = "Sin(t)"
    If Fy = "" Then Fy = "- (2 * Cos(t) ^ 4 * Rnd ^ 0.3) + 1"
    Main.txtXFormula.Text = Fx
    Main.txtYFormula.Text = Fy
    RetrieveFunctions = True
End Function

Public Function DisplayHelp(Optional ByVal strExpression As String = "", _
                            Optional ByVal strReqArgs As String = "", _
                            Optional ByVal strOptArgs As String = "", _
                            Optional ByVal strExplanation As String = "") As Boolean
    DisplayHelp = True
    On Error GoTo ErrorTrap
    
    Main.lblExpression.Caption = "   " & strExpression & "("
    Main.lblReqArg.Caption = strReqArgs
    Main.lblOptArg.Caption = strOptArgs
    Main.lblExplanation.Caption = ")  " & strExplanation
    
    Main.lblReqArg.Left = Main.lblExpression.Left + Main.lblExpression.Width
    Main.lblOptArg.Left = Main.lblReqArg.Left + Main.lblReqArg.Width
    Main.lblExplanation.Left = Main.lblOptArg.Left + Main.lblOptArg.Width
    Main.picPopup.ZOrder (0)
    Main.picPopup.Visible = True
    
    DoEvents
    Exit Function
ErrorTrap:
    DisplayHelp = False
End Function

Public Function PositionHelp(ByVal txtParent As TextBox) As Boolean
    Main.picPopup.Left = txtParent.Left
    Main.picPopup.Top = txtParent.Top + txtParent.Height
    PositionHelp = True
End Function

Public Function KillHelp() As Boolean
    Main.picPopup.Visible = False
    KillHelp = True
End Function

Public Function UpdateHelp(ByVal txtParent As TextBox) As Boolean
    Dim strComplete As String
    Dim lngCaret As Long
    Dim strMethod As String
    
    strComplete = txtParent.Text
    lngCaret = txtParent.SelStart
    
    strMethod = FindCurrentMethod(strComplete, lngCaret)
    If strMethod = "" Then
        KillHelp
        Exit Function
    End If
    
    PositionHelp txtParent
    Select Case UCase(strMethod)
    Case "SIN": Call DisplayHelp(strMethod, "Angle(Number)", , "returns the sine of any angle in radians")
    Case "COS": Call DisplayHelp(strMethod, "Angle(Number)", , "returns the cosine of any angle in radians")
    Case "TAN": Call DisplayHelp(strMethod, "Angle(Number)", , "returns the tangent of any angle in radians")
    Case "ATN": Call DisplayHelp(strMethod, "Angle(Number)", , "returns the inverse tangent of any angle in radians")
    Case "ABS": Call DisplayHelp(strMethod, "Value(Number)", , "returns the absolute (positive) version of any number")
    Case "EXP": Call DisplayHelp(strMethod, "Exponent(Number [-INF, 709])", , "returns e raised to a power")
    Case "FIX": Call DisplayHelp(strMethod, "Value(Number)", , "removes the integer portion of a number")
    Case "INT": Call DisplayHelp(strMethod, "Value(Number)", , "removes the integer portion of a number")
    Case "RND": Call DisplayHelp(strMethod, "[Type(N<0 seed; N=0 most recent; N>0 new random number)]", , "generates a random number")
    Case "LOG": Call DisplayHelp(strMethod, "Number [0, +INF]", , "returns the natural logarithm of a number")
    Case "MID": Call DisplayHelp(strMethod, "Expression(Text), Start(Number)", "[, Length(Number)]", "returns a section of a text")
    Case "RGB": Call DisplayHelp(strMethod, "Red[0,255], Green[0,255], Blue[0,255]", , "returns a long-format colour number")
    Case "SGN": Call DisplayHelp(strMethod, "Value(Number)", , "returns the sign of any number (positive=1 negative=-1)")
    Case "SQR": Call DisplayHelp(strMethod, "Value(Number [0, +INF])", , "returns the squareroot of any positive number")
    Case Else: KillHelp
    End Select
    
End Function

Private Function SpecialTrim(ByVal strIn As String) As String
    SpecialTrim = strIn
    Dim i As Long
    Dim strOut As String
    strOut = ""
    
    For i = 1 To Len(strIn)
        If Asc(Mid(strIn, i, 1)) >= 65 And Asc(Mid(strIn, i, 1)) <= 90 Then
            strOut = strOut & Mid(strIn, i, 1)
        End If
    Next
    SpecialTrim = strOut
End Function

Private Function IsVBMethod(ByVal strMethod As String) As Boolean
    IsVBMethod = False
    strMethod = UCase(strMethod)
    strMethod = SpecialTrim(strMethod)
    
    Dim strMethodList As String
    Dim inStrResult
    
    strMethodList = "SIN_COS_TAN_ATN_ABS_EXP_FIX_INT_LOG_MID_RGB_SGN_SQR_RND"
    inStrResult = InStr(1, strMethodList, strMethod, vbTextCompare)
    If inStrResult <> 0 Then IsVBMethod = True
End Function

Private Function FindCurrentMethod(ByVal strComplete As String, ByVal lngCaret As Long) As String
    FindCurrentMethod = ""
    
    Dim intDepth As Integer
    Dim i As Long
    Dim strMethod As String
    
    If lngCaret <= 3 Then Exit Function
    If Len(strComplete) <= 3 Then Exit Function
    intDepth = 0
    
    For i = lngCaret To 4 Step -1
        If Mid(strComplete, i, 1) = "(" Then intDepth = intDepth + 1
        If Mid(strComplete, i, 1) = ")" Then intDepth = intDepth - 1
        If intDepth = 1 Then
            strMethod = Mid(strComplete, i - 3, 3)
            If IsVBMethod(strMethod) Then
                FindCurrentMethod = strMethod
                Exit Function
            End If
        End If
    Next
    
End Function

Public Function FixExpression(ByVal strIn As String) As String
    Dim strOut As String
    strOut = LCase(strIn)
    strOut = Replace(strOut, " ", "")
    strOut = Replace(strOut, " ", "")
    strOut = Replace(strOut, " ", "")
    strOut = Replace(strOut, " ", "")
    strOut = Replace(strOut, "sin", " Sin")
    strOut = Replace(strOut, "cos", " Cos")
    strOut = Replace(strOut, "tan", " Tan")
    strOut = Replace(strOut, "atn", " Atn")
    strOut = Replace(strOut, "abs", " Abs")
    strOut = Replace(strOut, "exp", " Exp")
    strOut = Replace(strOut, "fix", " Fix")
    strOut = Replace(strOut, "int", " Int")
    strOut = Replace(strOut, "log", " Log")
    strOut = Replace(strOut, "mid", " Mid")
    strOut = Replace(strOut, "rgb", " Rgb")
    strOut = Replace(strOut, "rnd", " Rnd")
    strOut = Replace(strOut, "sgn", " Sgn")
    strOut = Replace(strOut, "sqr", " Sqr")
    strOut = Replace(strOut, "cint", " CInt")
    strOut = Replace(strOut, "clng", " CLng")
    strOut = Replace(strOut, "round", " Round")
    strOut = Replace(strOut, "left", " Left")
    strOut = Replace(strOut, "len", " Len")
    strOut = Replace(strOut, "right", " Right")
    strOut = Replace(strOut, "+", " +")
    strOut = Replace(strOut, "-", " -")
    strOut = Replace(strOut, "/", " / ")
    strOut = Replace(strOut, "\", " \ ")
    strOut = Replace(strOut, "*", " * ")
    strOut = Replace(strOut, "^", " ^ ")
    strOut = Replace(strOut, "  ", " ")
    strOut = Replace(strOut, "  ", " ")
    FixExpression = Trim(strOut)
End Function
