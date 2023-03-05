VERSION 5.00
Object = "{E63844D0-CCAA-43B9-BBC0-78BB7CCD6AC2}#1.0#0"; "aUfoButton.ocx"
Object = "{1EA1EF7D-6C2F-11D6-B7C7-FDD8DB077135}#1.0#0"; "AUFOOPTION.OCX"
Object = "{836F994E-6C2F-11D6-B7C7-FDD8DB077135}#1.0#0"; "AUFOINPUT.OCX"
Object = "{91E36AC5-5DE1-41D5-8818-713A48E8DF3F}#1.1#0"; "aUfoComboEx.ocx"
Object = "{F8B9583D-6C2E-11D6-B7C7-FDD8DB077135}#1.0#0"; "AUFOCHECK.OCX"
Begin VB.Form ocxForm 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   Picture         =   "ocxForm.frx":0000
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   StartUpPosition =   2  'CenterScreen
   Begin aUfoComboEx.ComboEx CB 
      Height          =   285
      Left            =   225
      TabIndex        =   5
      Top             =   2040
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ColumnField     =   "0,1,2,3,4,0,0,0,0,0"
      ColumnVType     =   "0,0,0,0,3,0,0,0,0,0"
      ColumnCSize     =   "10,10,10,50,8,0,0,0,0,0"
      ColumnWidth     =   "60,60,60,220,60,0,0,0,0,0"
      ColumnSetup     =   ",,,,0 ;(0);""Zero"";""Miss"",,,,,"
      ColumnAlign     =   "0,0,0,0,2,0,0,0,0,0"
      ColumnTitle     =   "Col.0,Col.1,Col.2,Col.3,Col.4,Col.5,Col.6,Col.7,Col.8,Col.9"
      ListColumns     =   5
      ListBoundTo     =   2
      ListCaption     =   2
      ListWidth       =   500
      ListItems       =   15
      ListCellMargin  =   2
      ListHeight      =   12
      ListScroll      =   3
      ListScrollStyle =   2
      ListHeader      =   3
      ListIntegral    =   0   'False
      ListForeColor   =   8388608
      ListGridColor   =   15124413
      ListHeaderColor =   16761024
      ListTitlesColor =   255
      ListBackColorOver=   192
      ComboMargin     =   "5,0,0,3"
      LabelBackStyle  =   0
      LabelMargin     =   "0,0,5,2"
      LabelFlashingInterval=   500
      StatusOver      =   2
      StatusAuto      =   -1  'True
      LabelCaption    =   "COMBO LABEL:"
      ComboCaption    =   "COMBOEX"
      BeginProperty FontCombo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StatusLabelForeColor0=   255
      StatusLabelForeColor4=   255
      StatusLabelBordColor41=   4227200
      ODBCoption      =   -1  'True
      ODBCsource      =   "Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=sample.mdb;User Id=;Password=;"
      ODBCselect      =   "SELECT * FROM tabUNIT;"
   End
   Begin aUfoButton.Button IE 
      Height          =   300
      Left            =   2475
      TabIndex        =   4
      Top             =   1650
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "ocxForm.frx":1FB96
      ImageColor      =   12632256
      ImageColorOver  =   12632256
      ImageCellHeight =   16
      ImageCellWidth  =   24
      BorderColorTL   =   12632256
      BorderColorTLover=   16777215
      BorderColorBR   =   12632256
      BorderColorBRover=   0
      BackColorOver   =   12632256
      ForeColor       =   0
      ForeColorOver   =   16711680
      ForeColorShadow =   14737632
      ForeColorOverShadow=   16777215
      Caption         =   "IE Button - Test Combo"
      CaptionShadow   =   -1  'True
      ShowFocus       =   0   'False
   End
   Begin aUfoButton.Button XP 
      Height          =   300
      Left            =   225
      TabIndex        =   3
      Top             =   1650
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "ocxForm.frx":20130
      ImageCellHeight =   18
      Caption         =   "XP Button"
   End
   Begin aUfoInput.TextBox TB 
      Height          =   300
      Left            =   225
      TabIndex        =   2
      Top             =   1305
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   529
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InputFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InputData_Mask  =   ""
      InputData_Format=   ""
      InputData_Size  =   20
   End
   Begin aUfoOption.OptionGroup OG 
      Height          =   900
      Left            =   225
      TabIndex        =   1
      Top             =   375
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   159
      LabelCaption    =   "CAPTION: this can be more than one line"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelTextAlignment=   0
      LabelMargin_Top =   5
      LabelMargin_Left=   5
      LabelMargin_Right=   0
      BeginProperty InputFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Buttons         =   3
      Captions        =   "Option0, Option1,Option2,,,,,,"
      InputHeight     =   18
   End
   Begin aUfoCheck.Check CK 
      Height          =   285
      Left            =   225
      TabIndex        =   0
      Top             =   75
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      LabelCaption    =   "Combo WorkMode:"
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InputFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InputTextChecked=   "OCBC Option"
      InputTextUncheck=   "Standard Combo"
   End
   Begin aUfoComboEx.ComboEx EX 
      Height          =   285
      Left            =   225
      TabIndex        =   6
      Top             =   2355
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ColumnField     =   "0,1,2,3,4,5,6,7,8,0"
      ColumnVType     =   "0,0,0,0,3,0,3,0,1,0"
      ColumnCSize     =   "10,10,10,50,8,10,1,10,16,0"
      ColumnWidth     =   "80,60,60,240,60,100,30,80,120,0"
      ColumnSetup     =   ",,,,0,,""Si"";"""";""No"",,dd-mm-yyyy hh:mm,"
      ColumnAlign     =   "0,0,0,0,2,0,1,1,1,0"
      ColumnTitle     =   "Modulo,Tipo,Risorsa,Descrizione,Tag#,Config,Use,Status,Update,C9"
      ListColumns     =   9
      ListBoundTo     =   2
      ListCaption     =   2
      ListWidth       =   400
      ListItems       =   15
      ListScroll      =   3
      ListHeader      =   3
      ListSBarColor   =   14737632
      ListTitlesColor =   16711680
      ListBackColorOver=   255
      ComboMargin     =   "5,0,0,2"
      LabelBackStyle  =   0
      LabelMargin     =   "0,0,0,2"
      LabelFlashing   =   -1  'True
      LabelFlashingInterval=   500
      Status          =   4
      StatusOver      =   2
      StatusAuto      =   -1  'True
      BeginProperty FontCombo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StatusLabelForeColor0=   16777215
      StatusLabelForeColor4=   255
      StatusLabelBordColor41=   4227200
      ODBCoption      =   -1  'True
      ODBCsource      =   "Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=sample.mdb;User Id=;Password=;"
      ODBCselect      =   "SELECT * FROM tabUNIT;"
   End
End
Attribute VB_Name = "ocxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private x
Private docPAth, defPath
Function CB_Cberror(code, text)
   MsgBox code & " - " & text
End Function
Private Sub IE_Click()
   CB.Clear
   EX.Clear
   If CK.Check Then
      CB.ODBCoption = True
      CB.Requery
      EX.ODBCoption = True
      EX.Requery
   Else
      CB.ODBCoption = False
      EX.ODBCoption = False
      For x = 0 To 1000
         CB.AddItem "row: " & Format(x, "0000"), 1000 - x
         EX.AddItem "cbo: " & Format(x, "0000"), 1000 - x
      Next
   End If
End Sub
