VERSION 5.00
Object = "{A7C75093-2765-11D3-A0E4-FAFD20CEB591}#5.0#0"; "CButton.ocx"
Begin VB.Form frmButtons 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   12495
   ClientTop       =   1845
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   1680
   Begin CButton.Button cmdClearEvents 
      Height          =   375
      Left            =   60
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
   Begin VB.ListBox List 
      Height          =   2205
      Left            =   60
      TabIndex        =   8
      Top             =   3120
      Width           =   1515
   End
   Begin VB.ComboBox cbStyle 
      Height          =   315
      ItemData        =   "frmButtons.frx":0000
      Left            =   60
      List            =   "frmButtons.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2400
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   1275
      Left            =   60
      ScaleHeight     =   1215
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   1080
      Width           =   1515
      Begin CButton.Button Button2 
         Height          =   495
         Index           =   0
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Button 2"
         Top             =   60
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   873
         Picture         =   "frmButtons.frx":004D
         PictureHover    =   "frmButtons.frx":0367
         PictureSelected =   "frmButtons.frx":0681
         BackColor       =   0
         SelectedBackColor=   255
         HoverBackColor  =   12632256
         ForeColor       =   255
         SelectedForeColor=   12632256
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         AmbientColor    =   0   'False
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Style           =   4
         Caption         =   "Button"
         Alignment       =   3
         GroupNumber     =   1
      End
      Begin CButton.Button Button2 
         Height          =   495
         Index           =   1
         Left            =   60
         TabIndex        =   6
         ToolTipText     =   "Button 2"
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         Unselectable    =   0   'False
         AmbientColor    =   -1  'True
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Style           =   5
         Caption         =   "SoftButton"
         Alignment       =   0
         GroupNumber     =   1
      End
   End
   Begin CButton.Button Button1 
      Default         =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      ToolTipText     =   "Button 1"
      Top             =   60
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
      Picture         =   "frmButtons.frx":099B
      PictureSelected =   "frmButtons.frx":0DED
      BackColor       =   8438015
      SelectedBackColor=   12648384
      HoverBackColor  =   16777152
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
      Style           =   4
      Caption         =   ""
      GroupNumber     =   0
   End
   Begin CButton.Button Button1 
      Height          =   435
      Index           =   0
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Button 1"
      Top             =   60
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
      Picture         =   "frmButtons.frx":1107
      PictureHover    =   "frmButtons.frx":1421
      PictureSelected =   "frmButtons.frx":173B
      BackColor       =   8438015
      SelectedBackColor=   12648447
      HoverBackColor  =   16777152
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
      Style           =   4
      Caption         =   ""
      CanGetFocus     =   -1  'True
      GroupNumber     =   0
   End
   Begin CButton.Button Button1 
      Height          =   435
      Index           =   3
      Left            =   1020
      TabIndex        =   2
      ToolTipText     =   "Button 1"
      Top             =   540
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
      Picture         =   "frmButtons.frx":1A55
      PictureSelected =   "frmButtons.frx":1D6F
      BackColor       =   8438015
      SelectedBackColor=   16761087
      HoverBackColor  =   16777152
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
      Style           =   4
      Caption         =   ""
      GroupNumber     =   0
   End
   Begin CButton.Button Button1 
      Height          =   435
      Index           =   2
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "Button 1"
      Top             =   540
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   767
      Picture         =   "frmButtons.frx":2089
      PictureSelected =   "frmButtons.frx":23A3
      BackColor       =   8438015
      SelectedBackColor=   12640511
      HoverBackColor  =   16777152
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
      Style           =   4
      Caption         =   ""
      CanGetFocus     =   -1  'True
      GroupNumber     =   0
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button1_ButtonSelected(Index As Integer, Cancel As Boolean)
    List.AddItem "Button " & Index & " Selected"
End Sub

Private Sub Button1_ButtonUnSelected(Index As Integer, Cancel As Boolean)
    List.AddItem "Button " & Index & " Unselected"
End Sub

Private Sub Button1_Click(Index As Integer)
    List.AddItem "Button " & Index & " Clicked"
End Sub

Private Sub cbStyle_Click()
    Dim iCount As Integer
    
    On Error Resume Next
    
    For iCount = 0 To 3
        Button1(iCount).Style = cbStyle.ListIndex
    Next
    
    For iCount = 0 To 1
        Button2(iCount).Style = cbStyle.ListIndex
    Next
    
End Sub

Private Sub cmdClearEvents_Click()
    List.Clear
End Sub

Private Sub Form_Load()
    cbStyle.ListIndex = 4
End Sub

Private Sub List_DblClick()
    List.Clear
End Sub
