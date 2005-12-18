VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.4#0"; "Imgedit.ocx"
Begin VB.Form frmVisorImag 
   Caption         =   "Visor"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _Version        =   131076
      _ExtentX        =   12938
      _ExtentY        =   10186
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit2"
      DisplayScaleAlgorithm=   2
      UndoBufferSize  =   430867968
      OcrZoneVisibility=   -3592
      AnnotationOcrType=   25801
      ForceFileLinking1x=   -1  'True
      MagnifierZoom   =   25801
      sReserved1      =   -3984
      sReserved2      =   -3984
      lReserved1      =   2117707439
      lReserved2      =   2117707439
      bReserved1      =   -1  'True
      bReserved2      =   -1  'True
   End
End
Attribute VB_Name = "frmVisorImag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
ImgEdit1.Width = Me.ScaleWidth
ImgEdit1.Height = Me.ScaleHeight
ImgEdit1.FitTo 1
End Sub

Private Sub ImgEdit1_MarkEnd(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal MarkType As Integer, ByVal GroupName As String)
ImgEdit1.ZoomToSelection
ImgEdit1.Refresh
End Sub
Private Sub ImgEdit1_DblClick()
ImgEdit1.FitTo 1
End Sub

Private Sub ImgEdit1_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
On Error Resume Next
ImgEdit1.ZoomToSelection
ImgEdit1.Refresh
End Sub
