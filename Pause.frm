VERSION 5.00
Begin VB.Form Pause 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pause lissajoux graph"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1935
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox thumb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume graph simulation"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "Pause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdResume_Click()
    PauseExecution = False
    Unload Pause
End Sub

Public Function DrawThumbNail(ByVal arrGrey) As Boolean
    thumb.BackColor = RGB(255, 200, 0)
    thumb.Cls
    Dim i As Long, j As Long
    Dim x As Long, y As Long
    Dim FAverage As Double
    
    For i = 0 To 499 Step 4
    For j = 0 To 499 Step 4
        FAverage = 0
        For x = 0 To 3
        For y = 0 To 3
            FAverage = FAverage + arrGrey(i + x, j + y)
        Next
        Next
        FAverage = FAverage / 16
        thumb.PSet (i / 4, j / 4), GetBlendColour(FAverage)
    Next
    Next
    DrawThumbNail = True
End Function

Private Function GetBlendColour(ByVal darkness As Byte) As Long
    GetBlendColour = RGB(255, 200, 0)
    Dim Factor As Double
    Dim r As Byte, g As Byte, b As Byte
    Factor = darkness / 255
    r = 155 + Factor * 100
    g = Factor * 200
    b = 0
    GetBlendColour = RGB(r, g, b)
End Function
