VERSION 5.00
Object = "{34A5E085-7527-11D2-9718-000000000000}#1.1#0"; "glxCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mnfrm 
   Caption         =   "Wavefront .Obj Files Viewer"
   ClientHeight    =   4305
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   4785
   Icon            =   "mnfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgLoadModel 
      Left            =   960
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin glCtl.glxCtl glxCtl1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      BorderStyle     =   0
   End
   Begin VB.Menu mnuLoadModel 
      Caption         =   "Load Model"
   End
   Begin VB.Menu MnuColor 
      Caption         =   "Change Color"
      Begin VB.Menu mnuCBackGround 
         Caption         =   "BackGround"
         Begin VB.Menu mnuBGColor 
            Caption         =   "Beige"
            Index           =   0
         End
         Begin VB.Menu mnuBGColor 
            Caption         =   "Deep Sky Blue"
            Index           =   1
         End
         Begin VB.Menu mnuBGColor 
            Caption         =   "Gray"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnuBGColor 
            Caption         =   "Pale Golden rod"
            Index           =   3
         End
         Begin VB.Menu mnuBGColor 
            Caption         =   "Medium Blue"
            Index           =   4
         End
         Begin VB.Menu mnuBGColor 
            Caption         =   "Light Gray"
            Index           =   5
         End
         Begin VB.Menu mnuBGColor 
            Caption         =   "Olive Drab"
            Index           =   6
         End
      End
      Begin VB.Menu mnuCModel 
         Caption         =   "Model"
         Enabled         =   0   'False
         Begin VB.Menu MnuMColor 
            Caption         =   "White"
            Index           =   0
         End
         Begin VB.Menu MnuMColor 
            Caption         =   "CBlack"
            Index           =   1
         End
         Begin VB.Menu MnuMColor 
            Caption         =   "Red"
            Index           =   2
         End
         Begin VB.Menu MnuMColor 
            Caption         =   "Yellow"
            Index           =   3
         End
         Begin VB.Menu MnuMColor 
            Caption         =   "Orange"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu MnuMColor 
            Caption         =   "Brown"
            Index           =   5
         End
         Begin VB.Menu MnuMColor 
            Caption         =   "Green"
            Index           =   6
         End
         Begin VB.Menu MnuMColor 
            Caption         =   "Blue"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "mnfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''------------------------------------------------------------------------
''
'' Author      : Anas S.
'' Date        : 6 June 2003
'' Version     : 1.0
'' Description : Wavefront Object File Viewer
''
''------------------------------------------------------------------------

'This program needs you to


'Form resize
Private Sub Form_Resize()
 glxCtl1.Left = 0
 glxCtl1.Top = 0
 glxCtl1.Width = ScaleWidth
 glxCtl1.Height = ScaleHeight
End Sub

'OCX Key Down
Private Sub glxCtl1_KeyDown(KeyCode As Integer, Shift As Integer)
 cOGL.KeyDown KeyCode, Shift
End Sub

'OCX Draw
Private Sub glxCtl1_Draw()
    cOGL.Draw
End Sub

'OCX GL Initialize
Private Sub glxCtl1_InitGL()
 cOGL.InitGL
End Sub

'OCX Initialize
Private Sub glxCtl1_Init()
 cOGL.Init
End Sub

'OCX KeyPress
Private Sub glxCtl1_KeyPress(KeyAscii As Integer)
 cOGL.KeyPress KeyAscii
End Sub

'OCX Mouse Move
Private Sub glxCtl1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 cOGL.MouseMove Button, Shift, x, y
End Sub

'OCX Paint
Private Sub glxCtl1_Paint()
 gCtl.Render
End Sub

Private Sub mnuHelp_Click()
 Load frmHelp
 frmHelp.Show
End Sub

'Menu Load Files
Private Sub mnuLoadModel_Click()
 dlgLoadModel.FileName = ""
 dlgLoadModel.Filter = "Wavefront Object File (*.obj)|*.obj"
 dlgLoadModel.ShowOpen
 If dlgLoadModel.FileName = "" Then Exit Sub
 MousePointer = 11
 cOGL.FileName = dlgLoadModel.FileName
 mnuCModel.Enabled = Not (cOGL.MaterialExists)
 MousePointer = 0
End Sub

'Menu BackGround Color Click
Private Sub mnuBGColor_Click(Index As Integer)
 ClearBGColorItems
  mnuBGColor(Index).Checked = True
 cOGL.BackGroundColor = Index
End Sub

'Menu Model Color Click
Private Sub MnuMColor_Click(Index As Integer)
 ClearMatColorItems
 MnuMColor(Index).Checked = True
 cOGL.MaterialColor = Index
End Sub

'This routine clears all the check marks
'in the Background color menu
Private Sub ClearBGColorItems()
 Dim i As Integer
 For i = 0 To mnuBGColor.count - 1
  mnuBGColor(i).Checked = False
 Next
End Sub

'This routine clears all the check marks
'in the Model color menu
Private Sub ClearMatColorItems()
 Dim i As Integer
 For i = 0 To MnuMColor.count - 1
  MnuMColor(i).Checked = False
 Next
End Sub
