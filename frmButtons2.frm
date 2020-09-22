VERSION 5.00
Object = "{A7C75093-2765-11D3-A0E4-FAFD20CEB591}#5.0#0"; "CButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmButtons2 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   6960
   ClientTop       =   3015
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   3795
   Begin VB.Frame Frame1 
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   0
      TabIndex        =   29
      Top             =   6360
      Width           =   3735
      Begin VB.CheckBox chkUndSelected 
         Caption         =   "Check1"
         Height          =   195
         Left            =   3120
         TabIndex        =   44
         Top             =   1080
         Width           =   195
      End
      Begin VB.CheckBox chkUndHover 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2400
         TabIndex        =   43
         Top             =   1080
         Width           =   195
      End
      Begin VB.CheckBox chkUndNormal 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1620
         TabIndex        =   42
         Top             =   1080
         Width           =   195
      End
      Begin VB.CheckBox chkItalicSelected 
         Caption         =   "Check1"
         Height          =   195
         Left            =   3120
         TabIndex        =   41
         Top             =   840
         Width           =   195
      End
      Begin VB.CheckBox chkItalicHover 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2400
         TabIndex        =   40
         Top             =   840
         Width           =   195
      End
      Begin VB.CheckBox chkItalicNormal 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1620
         TabIndex        =   39
         Top             =   840
         Width           =   195
      End
      Begin VB.CheckBox chkBoldSelected 
         Caption         =   "Check1"
         Height          =   195
         Left            =   3120
         TabIndex        =   38
         Top             =   600
         Width           =   195
      End
      Begin VB.CheckBox chkBoldHover 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2400
         TabIndex        =   37
         Top             =   600
         Width           =   195
      End
      Begin VB.CheckBox chkBoldNormal 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1620
         TabIndex        =   36
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label12 
         Caption         =   "Underline"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label11 
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label10 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label9 
         Caption         =   "Selected"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Hover"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   31
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblFontNormal 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1380
         TabIndex        =   30
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Frame frmForeColor 
      Caption         =   "ForeColor"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   0
      TabIndex        =   22
      Top             =   5520
      Width           =   3735
      Begin VB.CommandButton cmdFCNormal 
         BackColor       =   &H80000012&
         Height          =   315
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdFCHover 
         BackColor       =   &H80000012&
         Height          =   315
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdFCSelected 
         BackColor       =   &H80000012&
         Height          =   315
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Normal:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   28
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Hover:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   27
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Selected:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2460
         TabIndex        =   26
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Frame frmBackColor 
      Caption         =   "BackColor"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   0
      TabIndex        =   15
      Top             =   4680
      Width           =   3735
      Begin VB.CommandButton cmdBCSelected 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdBCHover 
         Height          =   315
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdBCNormal 
         Height          =   315
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Selected:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2460
         TabIndex        =   21
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Hover:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   19
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Normal:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.CheckBox chkUnselectable 
      Caption         =   "Unselectable"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox chkPlaySounds 
      Caption         =   "Play Sounds"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   13
      Top             =   4140
      Width           =   1695
   End
   Begin VB.CheckBox chkCanGetFocus 
      Caption         =   "Can Get Focus"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   12
      Top             =   3840
      Width           =   1695
   End
   Begin CButton.Button cmdClear 
      Height          =   315
      Left            =   1500
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2580
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      BackColor       =   -2147483633
      SelectedBackColor=   -2147483633
      HoverBackColor  =   -2147483633
      HoverColor      =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AmbientColor    =   0   'False
      Enabled         =   -1  'True
      MaskColor       =   16777215
      UseMaskColor    =   -1  'True
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Style           =   2
      Caption         =   "Clear Events"
      Alignment       =   0
      GroupNumber     =   0
   End
   Begin VB.ListBox lstEvents 
      Height          =   2595
      Left            =   1500
      TabIndex        =   10
      Top             =   0
      Width           =   2235
   End
   Begin VB.ComboBox cbAlignment 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmButtons2.frx":0000
      Left            =   1320
      List            =   "frmButtons2.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3420
      Width           =   2415
   End
   Begin CButton.Button Button 
      Height          =   675
      Index           =   3
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Button #3"
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1191
      Picture         =   "frmButtons2.frx":0037
      PictureSelected =   "frmButtons2.frx":0351
      BackColor       =   -2147483633
      SelectedBackColor=   -2147483628
      HoverBackColor  =   -2147483633
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483630
      SelectedForeColor=   -2147483630
      HoverColor      =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AmbientColor    =   0   'False
      Enabled         =   -1  'True
      MaskColor       =   16777215
      UseMaskColor    =   -1  'True
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Caption         =   "Bu&tton"
      Alignment       =   3
      AccessKey       =   "t"
      GroupNumber     =   1
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3300
      Top             =   7020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkAmbient 
      Caption         =   "Ambient Color"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
   End
   Begin VB.ComboBox cbStyle 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmButtons2.frx":066B
      Left            =   1320
      List            =   "frmButtons2.frx":0681
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin CButton.Button Button 
      Cancel          =   -1  'True
      Height          =   675
      Index           =   1
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Button #1"
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1191
      Picture         =   "frmButtons2.frx":06B8
      PictureHover    =   "frmButtons2.frx":09D2
      PictureSelected =   "frmButtons2.frx":0CEC
      BackColor       =   -2147483633
      SelectedBackColor=   -2147483628
      HoverBackColor  =   -2147483633
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483630
      SelectedForeColor=   -2147483630
      HoverColor      =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AmbientColor    =   0   'False
      Enabled         =   -1  'True
      MaskColor       =   16777215
      UseMaskColor    =   -1  'True
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Caption         =   "B&utton"
      Alignment       =   3
      AccessKey       =   "u"
      GroupNumber     =   1
   End
   Begin CButton.Button Button 
      Default         =   -1  'True
      Height          =   675
      Index           =   0
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Button #0"
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1191
      Picture         =   "frmButtons2.frx":113E
      PictureHover    =   "frmButtons2.frx":1458
      PictureSelected =   "frmButtons2.frx":1772
      BackColor       =   -2147483633
      SelectedBackColor=   -2147483628
      HoverBackColor  =   -2147483633
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483630
      SelectedForeColor=   -2147483630
      HoverColor      =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AmbientColor    =   0   'False
      Enabled         =   -1  'True
      MaskColor       =   16777215
      UseMaskColor    =   -1  'True
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Caption         =   "&Button"
      Alignment       =   3
      AccessKey       =   "B"
      GroupNumber     =   1
   End
   Begin CButton.Button Button 
      Height          =   675
      Index           =   2
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Button #2"
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1191
      Picture         =   "frmButtons2.frx":1A8C
      PictureHover    =   "frmButtons2.frx":1DA6
      PictureSelected =   "frmButtons2.frx":20C0
      BackColor       =   -2147483633
      SelectedBackColor=   -2147483628
      HoverBackColor  =   -2147483633
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483630
      SelectedForeColor=   -2147483630
      HoverColor      =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AmbientColor    =   0   'False
      Enabled         =   -1  'True
      MaskColor       =   16777215
      UseMaskColor    =   -1  'True
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Caption         =   "Butt&on"
      Alignment       =   3
      AccessKey       =   "o"
      GroupNumber     =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Alignment:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblStyle 
      Caption         =   "Button Style:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3060
      Width           =   1335
   End
End
Attribute VB_Name = "frmButtons2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_ButtonSelected(Index As Integer, Cancel As Boolean)
    lstEvents.AddItem "ButtonSelected" & vbTab & "Button " & Index
End Sub

Private Sub Button_ButtonUnSelected(Index As Integer, Cancel As Boolean)
    lstEvents.AddItem "ButtonUnSelected" & vbTab & "Button " & Index
End Sub

Private Sub Button_Click(Index As Integer)
    lstEvents.AddItem "Click" & vbTab & vbTab & "Button " & Index
End Sub

Private Sub Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lstEvents.AddItem "MouseDown" & vbTab & "Button " & Index
End Sub

Private Sub Button_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lstEvents.AddItem "MouseUp" & vbTab & "Button " & Index
End Sub

Private Sub cbAlignment_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).Alignment = cbAlignment.ListIndex
    Next lCount
    
End Sub

Private Sub cbStyle_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).Style = cbStyle.ListIndex
    Next lCount
    
End Sub

Private Sub chkAmbient_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).AmbientColor = CBool(chkAmbient.Value)
    Next lCount

End Sub

Private Sub chkBoldHover_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).FontHover.Bold = CBool(chkBoldHover.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True

End Sub

Private Sub chkBoldNormal_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).Font.Bold = CBool(chkBoldNormal.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True
    
End Sub

Private Sub chkBoldSelected_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).FontSelected.Bold = CBool(chkBoldSelected.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True

End Sub

Private Sub chkCanGetFocus_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).CanGetFocus = CBool(chkCanGetFocus.Value)
    Next lCount
End Sub

Private Sub chkEnabled_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).Enabled = CBool(chkEnabled.Value)
    Next lCount
    
End Sub

Private Sub chkItalicHover_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).FontHover.Italic = CBool(chkItalicHover.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True
End Sub

Private Sub chkItalicNormal_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).Font.Italic = CBool(chkItalicNormal.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True
End Sub

Private Sub chkItalicSelected_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).FontSelected.Italic = CBool(chkItalicSelected.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True
End Sub

Private Sub chkPlaySounds_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).PlaySounds = CBool(chkPlaySounds.Value)
    Next lCount

End Sub

Private Sub chkUndHover_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).FontHover.Underline = CBool(chkUndHover.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True
End Sub

Private Sub chkUndNormal_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).Font.Underline = CBool(chkUndNormal.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True
End Sub

Private Sub chkUndSelected_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).FontSelected.Underline = CBool(chkUndSelected.Value)
    Next lCount

    'Only works after the button has been repainted.  For some reason
    'the Refresh method of the form doesn't work...
    Me.Visible = False
    Me.Visible = True
End Sub

Private Sub chkUnselectable_Click()
    Dim lCount As Long
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).Unselectable = CBool(chkUnselectable.Value)
    Next lCount

End Sub

Private Sub cmdBCHover_Click()
    Dim cColor As OLE_COLOR
    Dim lCount As Long
    
    On Error GoTo CancelError
    CD.CancelError = True
    CD.ShowColor
    
    cColor = CD.Color
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).BackColorHover = cColor
    Next lCount
    cmdBCHover.BackColor = cColor
    
    Exit Sub
CancelError:
End Sub

Private Sub cmdBCNormal_Click()
    Dim cColor As OLE_COLOR
    Dim lCount As Long
    
    On Error GoTo CancelError
    CD.CancelError = True
    CD.ShowColor
    
    cColor = CD.Color
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).BackColor = cColor
    Next lCount
    
    If cmdBCHover.BackColor = cmdBCNormal.BackColor Then
        'Change both the Hover and the Normal backcolors since they are the same
        cmdBCHover.BackColor = cColor
    End If
    cmdBCNormal.BackColor = cColor
    
    Exit Sub
CancelError:
End Sub

Private Sub cmdBCSelected_Click()
    Dim cColor As OLE_COLOR
    Dim lCount As Long
    
    On Error GoTo CancelError
    CD.CancelError = True
    CD.ShowColor
    
    cColor = CD.Color
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).BackColorSelected = cColor
    Next lCount
    cmdBCSelected.BackColor = cColor
    
    Exit Sub
CancelError:

End Sub

Private Sub cmdClear_Click()
    lstEvents.Clear
End Sub

Private Sub cmdFCHover_Click()
    Dim cColor As OLE_COLOR
    Dim lCount As Long
    
    On Error GoTo CancelError
    CD.CancelError = True
    CD.ShowColor
    
    cColor = CD.Color
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).ForeColorHover = cColor
    Next lCount
    cmdFCHover.BackColor = cColor
    
    Exit Sub
CancelError:

End Sub

Private Sub cmdFCNormal_Click()
    Dim cColor As OLE_COLOR
    Dim lCount As Long
    
    On Error GoTo CancelError
    CD.CancelError = True
    CD.ShowColor
    
    cColor = CD.Color
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).ForeColor = cColor
    Next lCount
    If cmdFCNormal.BackColor = cmdFCHover.BackColor Then
        cmdFCHover.BackColor = cColor
    End If
    cmdFCNormal.BackColor = cColor
    
    Exit Sub
CancelError:

End Sub

Private Sub cmdFCSelected_Click()
    Dim cColor As OLE_COLOR
    Dim lCount As Long
    
    On Error GoTo CancelError
    CD.CancelError = True
    CD.ShowColor
    
    cColor = CD.Color
    
    For lCount = Button.LBound To Button.UBound
        Button(lCount).ForeColorSelected = cColor
    Next lCount
    cmdFCSelected.BackColor = cColor
    
    Exit Sub
CancelError:

End Sub

Private Sub Form_Load()
    cbStyle.ListIndex = 0
    cbAlignment.ListIndex = 3
    
End Sub
