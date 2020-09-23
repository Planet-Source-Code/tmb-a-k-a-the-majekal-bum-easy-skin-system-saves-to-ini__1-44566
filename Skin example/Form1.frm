VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox XB 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   0
      Width           =   255
      Begin VB.Label X 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " X"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox M 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3225
      ScaleWidth      =   3585
      TabIndex        =   1
      Top             =   360
      Width           =   3615
      Begin MSComDlg.CommonDialog P 
         Left            =   3120
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer T 
         Interval        =   1
         Left            =   3120
         Top             =   1440
      End
      Begin VB.TextBox T2 
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Text            =   "Easy Skinner"
         Top             =   120
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog C 
         Left            =   3120
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "ini"
      End
      Begin VB.Label US 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Cl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clear Picture"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label C8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Form Text color:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label B 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Open"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label C7 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label C6 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label C5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label C4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label C3 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.Label C2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label C1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label S 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label O 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Open/Load"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Logo Picture:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.Image L 
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Over Button Back color:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Over Button Text color:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Button Back color:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Button Text color:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Form Back color:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Title Bar Back color:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Title Bar Text color:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label idt 
         BackStyle       =   0  'Transparent
         Caption         =   "Title Bar Text:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox TB 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.Label T1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Easy Skinner"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1815
      Left            =   1800
      TabIndex        =   30
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label AB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   29
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AB_Click()
frmAbout.Show
End Sub

'Color Squares
'select the color for the object
Private Sub C1_Click()
C.ShowColor
C1.BackColor = C.Color
update
End Sub

Private Sub C2_Click()
C.ShowColor
C2.BackColor = C.Color
update
End Sub

Private Sub C3_Click()
C.ShowColor
C3.BackColor = C.Color
update
End Sub

Private Sub C4_Click()
C.ShowColor
C4.BackColor = C.Color
update
End Sub

Private Sub C5_Click()
C.ShowColor
C5.BackColor = C.Color
update
End Sub

Private Sub C6_Click()
C.ShowColor
C6.BackColor = C.Color
update
End Sub

Private Sub C7_Click()
C.ShowColor
C7.BackColor = C.Color
update
End Sub

Private Sub C8_Click()
C.ShowColor
C8.BackColor = C.Color
update
End Sub
'Buttons
Private Sub B_Click()
'Load picture
Dim file
P.ShowOpen
file = P.FileName
L.Picture = LoadPicture(file)
End Sub

Private Sub Cl_Click()
'clear picure
L.Picture = LoadPicture()
End Sub




Private Sub Form_Load()
'Make form transparent has a small bug with making other parts besides the form invisable
'fixed by putting a lable exactly behide the "hole"
MakeTransparent frmMain
End Sub

Private Sub T1_MouseMove(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
'alow boarderless form to move
'If button is down then start move
If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub T_Timer()

If US.Caption = "0" Then
B.ForeColor = C5.BackColor
Cl.ForeColor = C5.BackColor
O.ForeColor = C5.BackColor
S.ForeColor = C5.BackColor
B.BackColor = C6.BackColor
Cl.BackColor = C6.BackColor
O.BackColor = C6.BackColor
S.BackColor = C6.BackColor
US.Caption = "1"
End If
End Sub



Private Sub X_Click()
Unload Me
'frmMain = Nothing
End Sub

Private Sub O_Click()
Dim value As Long
C.ShowOpen
T1.Caption = GetINISetting(C.FileName, _
                                "Skin", _
                                "Titlebar Text", _
                                "Skin Example")
C1.BackColor = GetINISetting(C.FileName, _
                                "Skin", _
                                "TBTC", _
                                "TMB")
C2.BackColor = GetINISetting(C.FileName, _
                                "Skin", _
                                "TBBC", _
                                "Default")
C3.BackColor = GetINISetting(C.FileName, _
                                "Skin", _
                                "MFBC", _
                                " ")
C4.BackColor = GetINISetting(C.FileName, _
                                "Skin", _
                                "MFTC", _
                                "33023")
C5.BackColor = GetINISetting(C.FileName, _
                                "Skin", _
                                "BTC", _
                                "33023")
C6.BackColor = GetINISetting(C.FileName, _
                                "Skin", _
                                "BBC", _
                                "0")
C7.BackColor = GetINISetting(C.FileName, _
                                "Skin", _
                                "MOBTC", _
                                "33023")
C8.BackColor = GetINISetting(C.FileName, _
                                "Skin", _
                                "MOBBC", _
                                "0")
StripFile C.FileName
file = GetINISetting(C.FileName, _
                                "Skin", _
                                "L", _
                                "")
LoadPicture (path & file)
update
End Sub

Private Sub S_Click()
'This will save all the colors and picture name to a INI file
C.ShowSave
'  lStatus = SaveINISetting(location of ini file, _
'                           section name, _
'                           name of object, _
'                           source of value)
'This will save all the colors and picture name to a INI file
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "Titlebar Text", _
                           T2.Text)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "TBTC", _
                           C1.BackColor)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "TBBC", _
                           C2.BackColor)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "MFBC", _
                           C3.BackColor)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "MFTC", _
                           C4.BackColor)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "BTC", _
                           C5.BackColor)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "BBC", _
                           C6.BackColor)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "MOBTC", _
                           C7.BackColor)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "MOBBC", _
                           C8.BackColor)
'remove the path to the pic...saving only the name on the pic with the extention
pic = RemovePath(P.FileName)
lStatus = SaveINISetting(C.FileName, _
                           "Skin", _
                           "L", _
                           pic)
End Sub

'update the skin with nonmouseover colors
Function update()
T1.Caption = T2.Text
AB.ForeColor = C1.BackColor
AB.BackColor = C2.BackColor
T1.ForeColor = C1.BackColor
X.ForeColor = C1.BackColor
TB.BackColor = C2.BackColor
XB.BackColor = C2.BackColor
M.BackColor = C3.BackColor
idt(0).ForeColor = C4.BackColor
idt(1).ForeColor = C4.BackColor
idt(2).ForeColor = C4.BackColor
idt(3).ForeColor = C4.BackColor
idt(4).ForeColor = C4.BackColor
idt(5).ForeColor = C4.BackColor
idt(6).ForeColor = C4.BackColor
idt(7).ForeColor = C4.BackColor
idt(8).ForeColor = C4.BackColor
idt(9).ForeColor = C4.BackColor
B.ForeColor = C5.BackColor
Cl.ForeColor = C5.BackColor
O.ForeColor = C5.BackColor
S.ForeColor = C5.BackColor
B.BackColor = C6.BackColor
Cl.BackColor = C6.BackColor
O.BackColor = C6.BackColor
S.BackColor = C6.BackColor
End Function
'update the buttions with the nonmouse over colors
Private Sub M_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
US.Caption = "0"
End Sub
'update buttions with mouse over colors and prevent a bug
Private Sub B_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
B.BackColor = C8.BackColor
B.ForeColor = C7.BackColor
Cl.ForeColor = C5.BackColor
O.ForeColor = C5.BackColor
S.ForeColor = C5.BackColor
Cl.BackColor = C6.BackColor
O.BackColor = C6.BackColor
S.BackColor = C6.BackColor
End Sub
Private Sub Cl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cl.BackColor = C8.BackColor
Cl.ForeColor = C7.BackColor
B.ForeColor = C5.BackColor
O.ForeColor = C5.BackColor
S.ForeColor = C5.BackColor
B.BackColor = C6.BackColor
O.BackColor = C6.BackColor
S.BackColor = C6.BackColor
End Sub
Private Sub O_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
O.BackColor = C8.BackColor
O.ForeColor = C7.BackColor
B.ForeColor = C5.BackColor
Cl.ForeColor = C5.BackColor
S.ForeColor = C5.BackColor
B.BackColor = C6.BackColor
Cl.BackColor = C6.BackColor
S.BackColor = C6.BackColor
End Sub
Private Sub S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
S.BackColor = C8.BackColor
S.ForeColor = C7.BackColor
B.ForeColor = C5.BackColor
Cl.ForeColor = C5.BackColor
O.ForeColor = C5.BackColor
B.BackColor = C6.BackColor
Cl.BackColor = C6.BackColor
O.BackColor = C6.BackColor
End Sub


