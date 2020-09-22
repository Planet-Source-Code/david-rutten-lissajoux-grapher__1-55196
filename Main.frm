VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lissajoux"
   ClientHeight    =   8775
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   9735
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   120
   End
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picPopup 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   647
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Label lblExplanation 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   ") explanation  "
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   8640
         TabIndex        =   21
         Top             =   0
         Width           =   990
      End
      Begin VB.Label lblOptArg 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   " [Optional arguments] "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   240
         Left            =   5400
         TabIndex        =   20
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label lblReqArg 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   " Required arguments "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   19
         Top             =   0
         Width           =   2235
      End
      Begin VB.Label lblExpression 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   " Expression("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1110
      End
   End
   Begin MSScriptControlCtl.ScriptControl Scr 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   0   'False
   End
   Begin VB.TextBox txtYFormula 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   8160
      Width           =   9735
   End
   Begin VB.TextBox txtXFormula 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   7680
      Width           =   9735
   End
   Begin VB.Frame frmControls 
      Caption         =   "Controls"
      Height          =   7560
      Left            =   7680
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      Begin VB.TextBox txtStart 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Text            =   "0.0"
         ToolTipText     =   "Start value of t"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1815
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton cmdSaveImage 
         Caption         =   "Save viewport"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Save the current graph as a bitmap..."
         Top             =   7080
         Width           =   1815
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "Sinc anti-alias"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "16x anti-alias"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "8x anti-alias"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "4x anti-alias"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "Preview"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "Direct"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton cmdCapture 
         Caption         =   "Capture viewport"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Copy the graph to the clipboard"
         Top             =   6720
         Width           =   1815
      End
      Begin VB.TextBox txtStepSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0.01"
         ToolTipText     =   "Incrementation of t"
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdRender 
         Caption         =   "Render"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblIncrement 
         Alignment       =   1  'Right Justify
         Caption         =   "Increment"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   735
      End
   End
   Begin VB.PictureBox Viewport 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7560
      Left            =   0
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7560
   End
   Begin VB.Label Statusbar 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   8520
      Width           =   9735
   End
   Begin VB.Menu mnuRoot 
      Caption         =   "Root"
      Visible         =   0   'False
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSin 
         Caption         =   "Sine"
      End
      Begin VB.Menu mnuCos 
         Caption         =   "Cosine"
      End
      Begin VB.Menu mnuTan 
         Caption         =   "Tangent"
      End
      Begin VB.Menu mnuAtn 
         Caption         =   "ArcTangent"
      End
      Begin VB.Menu mnuSqr 
         Caption         =   "SquareRoot"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngCaret As Long

Private Sub cmdPause_Click()
    PauseExecution = True
End Sub

Private Sub Form_Load()
    Call Interface.RetrieveFunctions
End Sub

Private Sub cmdRender_Click()
    Dim Fx As String
    Dim Fy As String
    Dim bytOpacity As Byte
    
    Fx = txtXFormula.Text
    Fy = txtYFormula.Text
    bytOpacity = 255
    
    If Not Math.IsFunctionValid(Fx) Or Not Math.IsFunctionValid(Fy) Then
        'MsgBox "One or both functions has an invalid syntax." & vbNewLine & _
                "All invalid functions are displayed in red.", vbOKOnly Or vbInformation, "Syntax error"
        'Exit Sub
    End If
    
    If optDisplay(1).Value Then
        bytOpacity = 128
    ElseIf optDisplay(2).Value Then
        bytOpacity = 60
    ElseIf optDisplay(3).Value Then
        bytOpacity = 30
    ElseIf optDisplay(4).Value Then
        bytOpacity = 15
    ElseIf optDisplay(5).Value Then
        bytOpacity = 1
    End If
    Set Viewport = Nothing
    Call Math.RenderLissajoux(Viewport, Fx, Fy, bytOpacity, Val(txtStart.Text), Val(txtStepSize.Text))
End Sub

Private Sub cmdSaveImage_Click()
    CreateScreenShot = True
End Sub

Private Sub cmdCapture_Click()
    SavePicture Viewport.Image, CStr(App.Path & "\Temp_Viewport.bmp")
    Set Viewport = LoadPicture(App.Path & "\Temp_Viewport.bmp")
    Clipboard.SetData Viewport.Image
End Sub

Private Sub cmdStop_Click()
    StopExecution = True
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub optDisplay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Index
    Case 0
        Statusbar.Caption = "Render graph with direct black pixels..."
    Case 1
        Statusbar.Caption = "Render graph with 50% opaque pixels..."
    Case 2
        Statusbar.Caption = "Render graph using 4x anti-aliasing sampling..."
    Case 3
        Statusbar.Caption = "Render graph using 8x anti-aliasing sampling..."
    Case 4
        Statusbar.Caption = "Render graph using 16x anti-aliasing sampling..."
    Case 5
        Statusbar.Caption = "Render graph using maximum 32-bit anti-aliasing sampling..."
    End Select
End Sub

Private Sub tmrUpdate_Timer()
    If tmrUpdate.Tag = "Fx" Then
        If txtXFormula.SelStart = lngCaret Then Exit Sub
        Call Interface.UpdateHelp(txtXFormula)
        lngCaret = txtXFormula.SelStart
    ElseIf tmrUpdate.Tag = "Fy" Then
        If txtYFormula.SelStart = lngCaret Then Exit Sub
        Call Interface.UpdateHelp(txtYFormula)
        lngCaret = txtYFormula.SelStart
    Else
        Call Interface.KillHelp
    End If
End Sub

Private Sub txtStart_LostFocus()
    txtStart.Text = Val(txtStart.Text)
End Sub

Private Sub txtStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "Set the starting value for the lissajoux graph"
End Sub

Private Sub txtStepSize_LostFocus()
    txtStepSize.Text = Abs(Val(txtStepSize.Text))
End Sub

'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------

Private Sub cmdStop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "Stop lissajoux iteration"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = ""
End Sub

Private Sub frmControls_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = ""
End Sub

Private Sub lblIncrement_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = ""
End Sub

Private Sub cmdCapture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "Capture the current graph to the clipboard"
End Sub

Private Sub txtStepSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "Set the increment step for the lissajoux graph"
End Sub

Private Sub txtXFormula_GotFocus()
    tmrUpdate.Tag = "Fx"
    txtXFormula.ForeColor = vbBlack
    Call Interface.PositionHelp(txtXFormula)
    tmrUpdate.Enabled = True
End Sub

Private Sub txtXFormula_LostFocus()
    tmrUpdate.Enabled = False
    lngCaret = -1
    Call Interface.KillHelp
    If Not Math.IsFunctionValid(txtXFormula.Text) Then
        txtXFormula.ForeColor = vbRed
    End If
    txtXFormula.Text = FixExpression(txtXFormula.Text)
    Call Interface.HardStoreFunctions
End Sub

Private Sub txtXFormula_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "Specify the function for the x-deviation"
End Sub

Private Sub txtYFormula_GotFocus()
    tmrUpdate.Tag = "Fy"
    txtYFormula.ForeColor = vbBlack
    Call Interface.PositionHelp(txtYFormula)
    tmrUpdate.Enabled = True
End Sub

Private Sub txtYFormula_LostFocus()
    tmrUpdate.Enabled = False
    lngCaret = -1
    Call Interface.KillHelp
    If Not Math.IsFunctionValid(txtYFormula.Text) Then
        txtYFormula.ForeColor = vbRed
    End If
    txtYFormula.Text = FixExpression(txtYFormula.Text)
    Call Interface.HardStoreFunctions
End Sub

Private Sub txtyFormula_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "Specify the function for the y-deviation"
End Sub

Private Sub Viewport_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "The lissajoux graph will be rendered here..."
End Sub

Private Sub cmdSaveImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "Save the current graph as a bitmap image..."
End Sub

Private Sub cmdRender_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Statusbar.Caption = "Start lissajoux iteration..."
End Sub
