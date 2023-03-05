VERSION 5.00
Object = "{B4D68092-804A-11D6-B7C7-8DD44F9CF15B}#1.0#0"; "aUfoButtonEx.ocx"
Begin VB.Form frmTest 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug ClsButton"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   3  'Windows Default
   Begin aUfoButtonEx.ButtonEx B1 
      Height          =   510
      Left            =   2085
      TabIndex        =   0
      Top             =   300
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   900
      ButtonStyle     =   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "THIS IS A BIG"
      CaptionPlacement=   8
      CaptionForcedPos=   0   'False
      IconCellLeftX   =   36
      IconCellWidth   =   30
      IconCellHeight  =   16
      ButtonBackStyle00=   2
      ButtonForeColor00=   0
      ButtonIconStyle00=   0
      ButtonIconImage00=   "frmTest.frx":08CA
      ButtonMaskColor00=   0
      ButtonMaskImage00=   -1  'True
      ButtonMaskImage01=   -1  'True
      ButtonWithImage01=   -1  'True
      ButtonBordColor00=   0
      ButtonBordColor01=   0
      ButtonBordColor02=   0
      ButtonBordColor03=   0
      ButtonBackStyle10=   2
      ButtonForeColor10=   255
      ButtonIconStyle10=   0
      ButtonIconImage10=   "frmTest.frx":26CC
      ButtonMaskColor10=   0
      ButtonMaskImage10=   -1  'True
      ButtonMaskImage11=   -1  'True
      ButtonWithImage11=   -1  'True
      ButtonBordColor10=   255
      ButtonBordColor11=   255
      ButtonBordColor12=   255
      ButtonBordColor13=   255
      ButtonBackStyle20=   2
      ButtonForeColor20=   10485760
      ButtonIconStyle20=   2
      ButtonIconImage20=   "frmTest.frx":44CE
      ButtonForeShadow20=   16777215
      ButtonTextShadow20=   -1  'True
      ButtonMaskColor20=   0
      ButtonMaskImage20=   -1  'True
      ButtonMaskImage21=   -1  'True
      ButtonWithImage21=   -1  'True
      ButtonBackStyle30=   2
      ButtonIconStyle30=   2
      ButtonIconImage30=   "frmTest.frx":62D0
      ButtonTextShadow30=   -1  'True
      ButtonMaskColor30=   0
      ButtonMaskImage30=   -1  'True
      ButtonMaskImage31=   -1  'True
      ButtonWithImage31=   -1  'True
      ButtonForeColor40=   33023
      ButtonIconStyle40=   0
      ButtonIconImage40=   "frmTest.frx":80D2
      ButtonForeShadow40=   12632256
      ButtonTextShadow40=   -1  'True
      ButtonMaskColor40=   0
      ButtonMaskImage40=   -1  'True
      ButtonMaskImage41=   -1  'True
      ButtonWithImage41=   -1  'True
      ButtonForeColor50=   16744448
      ButtonIconStyle50=   0
      ButtonIconImage50=   "frmTest.frx":9ED4
      ButtonForeShadow50=   14145495
      ButtonTextShadow50=   -1  'True
      ButtonMaskColor50=   0
      ButtonMaskImage50=   -1  'True
      ButtonMaskImage51=   -1  'True
      ButtonWithImage51=   -1  'True
      ButtonIconStyle60=   0
      ButtonIconImage60=   "frmTest.frx":BCD6
      ButtonTextShadow60=   -1  'True
      ButtonMaskColor60=   0
      ButtonMaskImage60=   -1  'True
      ButtonMaskImage61=   -1  'True
      ButtonWithImage61=   -1  'True
      ButtonBordColor60=   16777215
      ButtonBordColor61=   0
      ButtonBordColor62=   0
      ButtonBordColor63=   16777215
      ButtonIconStyle70=   0
      ButtonIconImage70=   "frmTest.frx":DAD8
      ButtonTextShadow70=   -1  'True
      ButtonMaskColor70=   0
      ButtonMaskImage70=   -1  'True
      ButtonMaskImage71=   -1  'True
      ButtonWithImage71=   -1  'True
      ButtonBordColor71=   0
      ButtonBordColor72=   0
      DropButtons     =   10
      DropCellColor   =   12626338
      DropCellColorOver=   0
      DropBackColor   =   12626338
      DropBackColorOver=   0
      DropGapAround   =   1
      BeginProperty DropCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropCaptionAlign=   0
      DropWidth       =   120
      DropCellHeight  =   16
      DropWithImage00 =   -1  'True
      DropIconImage00 =   "frmTest.frx":F8DA
      DropWithImage10 =   -1  'True
      DropIconImage10 =   "frmTest.frx":116DC
      DropWithImage20 =   -1  'True
      DropIconImage20 =   "frmTest.frx":134DE
      DropWithImage30 =   -1  'True
      DropIconImage30 =   "frmTest.frx":152E0
      DropWithImage40 =   -1  'True
      DropIconImage40 =   "frmTest.frx":170E2
      DropWithImage50 =   -1  'True
      DropIconImage50 =   "frmTest.frx":18EE4
      DropWithImage60 =   -1  'True
      DropIconImage60 =   "frmTest.frx":1ACE6
      DropWithImage70 =   -1  'True
      DropIconImage70 =   "frmTest.frx":1CAE8
      DropWithImage80 =   -1  'True
      DropIconImage80 =   "frmTest.frx":1E8EA
      DropWithImage90 =   -1  'True
      DropIconImage90 =   "frmTest.frx":206EC
   End
   Begin aUfoButtonEx.ButtonEx B2 
      Height          =   360
      Left            =   4665
      TabIndex        =   1
      Top             =   375
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
      ButtonStyle     =   1
      ButtonSeparator =   -1  'True
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCellWidth   =   30
      IconCellHeight  =   18
      ButtonBackStyle00=   2
      ButtonCellColor00=   255
      ButtonForeColor00=   0
      ButtonIconStyle00=   0
      ButtonIconImage00=   "frmTest.frx":224EE
      ButtonMaskColor00=   0
      ButtonMaskImage00=   -1  'True
      ButtonMaskImage01=   -1  'True
      ButtonWithImage01=   -1  'True
      ButtonBordColor00=   0
      ButtonBordColor01=   0
      ButtonBordColor02=   0
      ButtonBordColor03=   0
      ButtonBackStyle10=   2
      ButtonForeColor10=   255
      ButtonIconStyle10=   0
      ButtonIconImage10=   "frmTest.frx":242F0
      ButtonMaskColor10=   0
      ButtonMaskImage10=   -1  'True
      ButtonMaskImage11=   -1  'True
      ButtonWithImage11=   -1  'True
      ButtonBordColor10=   255
      ButtonBordColor11=   255
      ButtonBordColor12=   255
      ButtonBordColor13=   255
      ButtonBackStyle20=   2
      ButtonForeColor20=   10485760
      ButtonIconStyle20=   2
      ButtonIconImage20=   "frmTest.frx":260F2
      ButtonForeShadow20=   16777215
      ButtonTextShadow20=   -1  'True
      ButtonMaskColor20=   0
      ButtonMaskImage20=   -1  'True
      ButtonMaskImage21=   -1  'True
      ButtonWithImage21=   -1  'True
      ButtonBackColor30=   12632256
      ButtonIconStyle30=   2
      ButtonIconImage30=   "frmTest.frx":27EF4
      ButtonTextShadow30=   -1  'True
      ButtonMaskColor30=   0
      ButtonMaskImage30=   -1  'True
      ButtonMaskImage31=   -1  'True
      ButtonWithImage31=   -1  'True
      ButtonForeColor40=   33023
      ButtonIconStyle40=   0
      ButtonIconImage40=   "frmTest.frx":29CF6
      ButtonForeShadow40=   12632256
      ButtonTextShadow40=   -1  'True
      ButtonMaskColor40=   0
      ButtonMaskImage40=   -1  'True
      ButtonMaskImage41=   -1  'True
      ButtonWithImage41=   -1  'True
      ButtonForeColor50=   16744448
      ButtonIconStyle50=   0
      ButtonIconImage50=   "frmTest.frx":2BAF8
      ButtonForeShadow50=   14145495
      ButtonTextShadow50=   -1  'True
      ButtonMaskColor50=   0
      ButtonMaskImage50=   -1  'True
      ButtonMaskImage51=   -1  'True
      ButtonWithImage51=   -1  'True
      ButtonIconStyle60=   0
      ButtonIconImage60=   "frmTest.frx":2D8FA
      ButtonTextShadow60=   -1  'True
      ButtonMaskColor60=   0
      ButtonMaskImage60=   -1  'True
      ButtonMaskImage61=   -1  'True
      ButtonWithImage61=   -1  'True
      ButtonBordColor60=   16777215
      ButtonBordColor61=   0
      ButtonBordColor62=   0
      ButtonBordColor63=   16777215
      ButtonIconStyle70=   0
      ButtonIconImage70=   "frmTest.frx":2F6FC
      ButtonTextShadow70=   -1  'True
      ButtonMaskColor70=   0
      ButtonMaskImage70=   -1  'True
      ButtonMaskImage71=   -1  'True
      ButtonWithImage71=   -1  'True
      ButtonBordColor71=   0
      ButtonBordColor72=   0
      BeginProperty DropCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin aUfoButtonEx.ButtonEx B3 
      Height          =   360
      Left            =   4785
      TabIndex        =   2
      Top             =   1185
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
      ButtonStyle     =   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconCellWidth   =   30
      IconCellHeight  =   18
      ButtonBackStyle00=   2
      ButtonForeColor00=   0
      ButtonIconStyle00=   0
      ButtonIconImage00=   "frmTest.frx":314FE
      ButtonForeShadow00=   16777215
      ButtonTextShadow00=   -1  'True
      ButtonMaskColor00=   0
      ButtonMaskImage00=   -1  'True
      ButtonMaskImage01=   -1  'True
      ButtonWithImage01=   -1  'True
      ButtonBordColor00=   0
      ButtonBordColor01=   0
      ButtonBordColor02=   0
      ButtonBordColor03=   0
      ButtonBackStyle10=   2
      ButtonForeColor10=   255
      ButtonIconStyle10=   0
      ButtonIconImage10=   "frmTest.frx":33300
      ButtonMaskColor10=   0
      ButtonMaskImage10=   -1  'True
      ButtonMaskImage11=   -1  'True
      ButtonWithImage11=   -1  'True
      ButtonBordColor10=   255
      ButtonBordColor11=   255
      ButtonBordColor12=   255
      ButtonBordColor13=   255
      ButtonBackStyle20=   2
      ButtonForeColor20=   10485760
      ButtonIconStyle20=   2
      ButtonIconImage20=   "frmTest.frx":35102
      ButtonForeShadow20=   16777215
      ButtonTextShadow20=   -1  'True
      ButtonMaskColor20=   0
      ButtonMaskImage20=   -1  'True
      ButtonMaskImage21=   -1  'True
      ButtonWithImage21=   -1  'True
      ButtonBackStyle30=   2
      ButtonIconStyle30=   2
      ButtonIconImage30=   "frmTest.frx":36F04
      ButtonTextShadow30=   -1  'True
      ButtonMaskColor30=   0
      ButtonMaskImage30=   -1  'True
      ButtonMaskImage31=   -1  'True
      ButtonWithImage31=   -1  'True
      ButtonForeColor40=   33023
      ButtonIconStyle40=   0
      ButtonIconImage40=   "frmTest.frx":38D06
      ButtonForeShadow40=   12632256
      ButtonTextShadow40=   -1  'True
      ButtonMaskColor40=   0
      ButtonMaskImage40=   -1  'True
      ButtonMaskImage41=   -1  'True
      ButtonWithImage41=   -1  'True
      ButtonForeColor50=   16744448
      ButtonIconStyle50=   0
      ButtonIconImage50=   "frmTest.frx":3AB08
      ButtonForeShadow50=   14145495
      ButtonTextShadow50=   -1  'True
      ButtonMaskColor50=   0
      ButtonMaskImage50=   -1  'True
      ButtonMaskImage51=   -1  'True
      ButtonWithImage51=   -1  'True
      ButtonIconStyle60=   0
      ButtonIconImage60=   "frmTest.frx":3C90A
      ButtonTextShadow60=   -1  'True
      ButtonMaskColor60=   0
      ButtonMaskImage60=   -1  'True
      ButtonMaskImage61=   -1  'True
      ButtonWithImage61=   -1  'True
      ButtonBordColor60=   16777215
      ButtonBordColor61=   0
      ButtonBordColor62=   0
      ButtonBordColor63=   16777215
      ButtonIconStyle70=   0
      ButtonIconImage70=   "frmTest.frx":3E70C
      ButtonTextShadow70=   -1  'True
      ButtonMaskColor70=   0
      ButtonMaskImage70=   -1  'True
      ButtonMaskImage71=   -1  'True
      ButtonWithImage71=   -1  'True
      ButtonBordColor71=   0
      ButtonBordColor72=   0
      BeginProperty DropCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin aUfoButtonEx.ButtonEx B4 
      Height          =   300
      Left            =   915
      TabIndex        =   3
      Top             =   1605
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      ButtonStyle     =   1
      ButtonArrowY    =   4
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Receive"
      Icon            =   "frmTest.frx":4050E
      IconCellWidth   =   24
      IconCellHeight  =   16
      ButtonInternalGap=   "1,1,5,1"
      ButtonBordStyle10=   0
      ButtonForeShadow10=   16777215
      ButtonTextShadow10=   -1  'True
      ButtonIconStyle20=   0
      ButtonTextShadow20=   -1  'True
      DropButtons     =   5
      DropCellColorOver=   13811126
      DropBackColorOver=   13811126
      DropForeColor   =   4194432
      DropGapAround   =   0
      BeginProperty DropCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropBordStyle   =   7
      DropWidth       =   100
      DropCellHeight  =   18
      DropWithImage00 =   -1  'True
      DropIconImage00 =   "frmTest.frx":40AA8
      DropWithImage10 =   -1  'True
      DropIconImage10 =   "frmTest.frx":428AA
      DropWithImage20 =   -1  'True
      DropIconImage20 =   "frmTest.frx":446AC
      DropWithImage30 =   -1  'True
      DropIconImage30 =   "frmTest.frx":464AE
      DropWithImage40 =   -1  'True
      DropIconImage40 =   "frmTest.frx":482B0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iMove            As Integer
Private Sub B3_ButtonClick(Button As Integer)
   B1.Enabled = Not B1.Enabled
   B1.Top = B1.Top + iMove
   iMove = IIf(iMove > 0, -50, 50)
End Sub
Private Sub B2_ButtonClick(Button As Integer)
   B1.Enabled = Not B1.Enabled
   B1.Top = B1.Top - iMove
   iMove = IIf(iMove > 0, -50, 50)
End Sub
Private Sub Form_Load()
   B1.ZOrder 0
   B2.ZOrder 0
   B3.ZOrder 0
End Sub
Private Sub B1_MouseGone()
'  Debug.Print "B1 - gone"
End Sub
Private Sub B1_MouseOver()
'  Debug.Print "B1 - over"
End Sub
Private Sub B1_MousePush(Button As Integer, Shift As Integer, x As Single, y As Single)
'  Debug.Print "B1 - Push"
   B2.Enabled = Not B2.Enabled
   B3.Enabled = Not B3.Enabled
End Sub
Private Sub B1_Click(Button As Integer, ptrValue As Integer)
   B1.Caption = B1.DropCaption(Button - 1)
End Sub


