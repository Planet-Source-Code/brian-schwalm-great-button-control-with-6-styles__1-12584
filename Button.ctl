VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   975
   DefaultCancel   =   -1  'True
   PropertyPages   =   "Button.ctx":0000
   ScaleHeight     =   810
   ScaleWidth      =   975
   ToolboxBitmap   =   "Button.ctx":0025
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'The available button styles
Public Enum Styles
    Cool
    Flat
    ThreeD
    Outline
    Options_1
    Options_2
End Enum

'Property Name constants
Private Const msBACK_COLOR_NAME = "BackColor"
Private Const msBACK_COLOR_SELECTED = "SelectedBackColor"
Private Const msBACK_COLOR_HOVER = "HoverBackColor"

Private Const msFORE_COLOR_NAME = "ForeColor"
Private Const msFORE_COLOR_HOVER = "HoverForeColor"
Private Const msFORE_COLOR_SELECTED = "SelectedForeColor"

Private Const msPICTURE_NAME = "Picture"
Private Const msPICTURE_HOVER = "PictureHover"
Private Const msPICTURE_SELECTED = "PictureSelected"

Private Const msLINE_COLOR = "NormalColor"
Private Const msLINE_COLOR_HOVER = "HoverColor"
Private Const msLINE_COLOR_SELECTED = "SelectedColor"
Private Const msSHOW_BORDER = "ShowBorder"

Private Const msFONT = "Font"
Private Const msFONT_HOVER = "FontHover"
Private Const msFONT_SELECTED = "FontSelected"

Private Const msENABLED_NAME = "Enabled"
Private Const msMASK_COLOR_NAME = "MaskColor"
Private Const msUSE_MASK_COLOR_NAME = "UseMaskColor"

Private Const msUNSELECTABLE = "Unselectable"
Private Const msPLAY_SOUNDS_NAME = "PlaySounds"
Private Const msTOOL_TIP_TEXT_NAME = "ToolTipText"
Private Const msAMBIENT_COLOR = "AmbientColor"
Private Const msSTYLE = "Style"
Private Const msCAPTION = "Caption"
Private Const msALIGNMENT = "Alignment"
Private Const msACCESS_KEY = "AccessKey"
Private Const msGROUP_NUMBER = "GroupNumber"
Private Const msCANGETFOCUS = "CanGetFocus"

'Property Values
Private m_bEnabled As Boolean

Private m_clrMaskColor As OLE_COLOR
Private m_bUseMaskColor As Boolean

Private m_bPlaySounds As Boolean

Private m_picPicture As Picture
Private m_picPictureHover As Picture
Private m_picPictureSelected As Picture

Private m_cBackColor As OLE_COLOR
Private m_cBackColorHover As OLE_COLOR
Private m_cBackColorSelected As OLE_COLOR

Private m_cForeColor As OLE_COLOR
Private m_cForeColorSelected As OLE_COLOR
Private m_cForeColorHover As OLE_COLOR

Private m_cLineColor As OLE_COLOR
Private m_cLineColorHover As OLE_COLOR
Private m_cLineColorSelected As OLE_COLOR
Private m_bShowBorder As Boolean

Private m_eFont As StdFont
Private m_eFontHover As StdFont
Private m_eFontSelected As StdFont

Private m_bUnselectable As Boolean
Private m_sToolTipText As String
Private m_bAmbientColor As Boolean
Private m_eStyle As Styles
Private m_bSelected As Boolean
Private m_sCaption As String
Private m_eAlignment As AlignConstants
Private m_sAccessKey As String

Private m_sParentName As String
Private m_iGroupNumber As Integer

'Class level variables
Private msToolTipBuffer As String         'Tool tip text; This string must have module or global level scope, because a pointer to it is copied into a ToolTipText structure
'Private mbClearPictureOnly As Boolean
Private mbToolTipNotInExtender As Boolean
Private moDrawTool As clsDrawPictures
Private mbGotFocus As Boolean
Private m_bCanGetFocus As Boolean
Private mbMouseOver As Boolean
Private miCurrentState As Integer
Private mWndProcNext As Long            'The address entry point for the subclassed window
Private mHWndSubClassed As Long         'hWnd of the subclassed window
Private mbLeftMouseDown As Boolean
Private mbLeftWasDown As Boolean
Private mbKeyDown As Boolean

Private mudtButtonRect As RECT
Private mudtPictureRect As RECT

Private mudtPicturePoint As POINTAPI
Private mbPropertiesLoaded As Boolean
Private mbEnterOnce As Boolean
Private mbMouseDownFired As Boolean
Private mlhHalftonePal As Long
Private mbFromAmbient As Boolean

#If DEBUGSUBCLASS Then                      'Tool that checks if in break mode and then
    Private moProcHook As Object            'Sends messages to mWndPRocNext instead of
#End If                                     'Address of my function

Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Fires when the user releases the mouse over the button."
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Fires when the user presses the mouse down on the button."
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Fires when the mouse moves over the button."
Public Event PopUp()
Public Event ButtonSelected(ByRef Cancel As Boolean)
Attribute ButtonSelected.VB_Description = "Fires when an Option style button is selected."
Public Event ButtonUnSelected(ByRef Cancel As Boolean)
Attribute ButtonUnSelected.VB_Description = "Fires when an Options style button is unselected."

'The will be fired when the user presses the hotkey (Alt-X, where X is the hotkey)
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
        RaiseEvent Click
        UserControl_Click
        DrawButtonState miCurrentState
End Sub

'If the Ambient color changes, and the button is set to be the same as the parent, then change
Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "BackColor" And m_bAmbientColor And Not mbFromAmbient Then
        AmbientColor = True
    End If
End Sub

'****************************
'UserControl event procedures
'****************************
Public Sub Click()
Attribute Click.VB_Description = "Fires when the button is clicked."
    UserControl_Click
End Sub

Private Sub UserControl_Click()
    Dim bCancel As Boolean
    Dim iCount As Integer
    Dim sIdentification As String
    
    mbMouseOver = False     'Set mbMouseOver = False so that SetCapture will be called again.  For some reason, capture is released during MouseClick
    
    If m_eStyle = Options_1 Or m_eStyle = Options_2 Or m_eStyle = Outline Then      'If this is an Options Style button, look for others to deselect
        If Not m_bSelected Then             'Button not selected.  Select it.
            RaiseEvent ButtonSelected(bCancel)
            If bCancel Then Exit Sub                'Tell the parent, and let them cancel if they want
            
            'Loop through all the other controls on the parent, and see if they are CButtons.  If so, and if they are the same option style, and in the same
            'Option group, then they should be deselected.
            On Error Resume Next
            Do
               sIdentification = UserControl.ParentControls.Item(iCount).Identification
                If Err.Number = 0 Then
                    If sIdentification = "CButton" Then     'Check to see if it is the correct type of control
                        If UserControl.ParentControls.Item(iCount).Style = m_eStyle And UserControl.ParentControls.Item(iCount).Selected And Not UserControl.ParentControls.Item(iCount).Index = UserControl.Extender.Index Then
                            If UserControl.ParentControls.Item(iCount).GroupNumber = m_iGroupNumber Then
                                UserControl.ParentControls.Item(iCount).Selected = False        'Found one -- deselect it
                            End If
                        End If
                    End If
                Else
                    Err.Clear
                End If
                iCount = iCount + 1
            Loop While iCount < UserControl.ParentControls.Count
            
            m_bSelected = True
        
        Else        'Deselect the Button
            If Not m_bUnselectable Then Exit Sub        'Check to see if this button is UnSelectable.  If not, don't let them unselect this button
            '(Note: this will only prevent the user from Clicking on the button to Unselect it -- it can still be unselected programatically, or
            '            when another button is selected).
            
            RaiseEvent ButtonUnSelected(bCancel)
            If bCancel Then Exit Sub
            
            m_bSelected = False
            miCurrentState = giRAISED
            
        End If
    End If
    
    DrawButtonState miCurrentState
    MouseOver

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '-------------------------------------------------------------------------
    'Purpose:   If the mouse is over the button and the left button is down show that the button is sunken and set a flag that the button
    '           is down
    '-------------------------------------------------------------------------
    With UserControl
        If Button = vbLeftButton And (x >= 0 And x <= .ScaleWidth) And (y >= 0 And y <= .ScaleHeight) Then
            mbLeftMouseDown = True
            DrawButtonState giSUNKEN
        End If
    End With
    
    RaiseEvent MouseDown(Button, Shift, x, y)
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
            mbKeyDown = True
            DrawButtonState giSUNKEN
        End If
    End If
        
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If mbKeyDown Then
        DrawButtonState giRAISED
        mbKeyDown = False
        UserControl_Click
        DrawButtonState miCurrentState
        RaiseEvent Click
    End If
    
End Sub

Private Sub UserControl_EnterFocus()
    '-------------------------------------------------------------------------
    'Purpose:   If tabstop property is true, show button raised so that user
    '           can see that button received focus.
    '-------------------------------------------------------------------------
    On Error GoTo UserControl_EnterFocusError
    If Not m_bCanGetFocus Then Exit Sub
    
    'Error may occur if TabStop property is not available
    If UserControl.Extender.TabStop Then
        On Error Resume Next
        mbGotFocus = True
        If Not miCurrentState = giRAISED Then DrawButtonState giRAISED
    End If
UserControl_EnterFocusError:
    Exit Sub
End Sub

Private Sub UserControl_ExitFocus()
    '-------------------------------------------------------------------------
    'Purpose:   Flatten button if the mouse is not over it
    '-------------------------------------------------------------------------
    If Not m_bCanGetFocus Then Exit Sub
    
    mbGotFocus = False
    DrawButtonState giFLATTENED
    
End Sub

Private Sub UserControl_Initialize()
    Dim iCount As Integer
    
    mbPropertiesLoaded = False
    UserControl.ScaleMode = vbPixels
    UserControl.PaletteMode = vbPaletteModeContainer
    Set moDrawTool = New clsDrawPictures
    mlhHalftonePal = CreateHalftonePalette(UserControl.hdc)
    
End Sub

Private Sub UserControl_InitProperties()
    '-------------------------------------------------------------------------
    'Purpose:   Set the default properties to be displayed the first time this control is placed on a container
    '-------------------------------------------------------------------------
    On Error Resume Next    'Error may occur if TabStop property is not available
    AmbientColor = False
    
    BackColor = vbButtonFace
    BackColorSelected = vbButtonFace
    BackColorHover = vbButtonFace
    
    ForeColor = vbBlack
    ForeColorHover = vbBlack
    ForeColorSelected = vbBlack
    
    Enabled = True
    UseMaskColor = True
    MaskColor = vbWhite
    UserControl.ScaleMode = vbPixels
    UserControl.Extender.TabStop = False
    ToolTipText = ""
    Style = Cool
    ShowBorder = False
    LineColorNormal = vbBlack
    LineColorHover = vbBlue
    LineColorSelected = vbWhite
    Caption = UserControl.Name ' "SoftButton"
    Alignment = vbAlignNone
    CanGetFocus = False
    
    Set Font = UserControl.Font
    Set FontHover = UserControl.Font
    Set FontSelected = UserControl.Font
    
    mbPropertiesLoaded = True
    m_sAccessKey = ""
    Unselectable = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '-------------------------------------------------------------------------
    'Purpose:   If the mouse is over show the button raised.  If the mouse is over and the left mouse is down and was down when the mouse left
    '           the button show button sunken.  If mouse is off button, flatten, unless left button is down show the mouse raised until the button
    '           is released.
    '-------------------------------------------------------------------------
    With UserControl
        If (x <= .ScaleWidth And x >= 0) And (y <= .ScaleHeight And y >= 0) Then
            If mbLeftWasDown Then
                mbLeftMouseDown = True
                mbLeftWasDown = False
                DrawButtonState giSUNKEN
            Else
                If Button <> vbLeftButton Then MouseOver
            End If
            RaiseEvent MouseMove(Button, Shift, x, y)
        Else
            If mbLeftMouseDown Then
                mbLeftWasDown = True
                mbLeftMouseDown = False
                DrawButtonState giRAISED
                RaiseEvent MouseMove(Button, Shift, x, y)
            ElseIf Not mbLeftWasDown Then
                Flatten
            Else
                RaiseEvent MouseMove(Button, Shift, x, y)
            End If
        End If
    End With
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '-------------------------------------------------------------------------
    'Purpose:   If the left mouse was down and the left button is up and the mouse is over the button raise a Click event.  If left button was down
    '           and mouse is off button flatten button.
    '-------------------------------------------------------------------------
    With UserControl
        If (x >= 0 And x <= .ScaleWidth) And (y >= 0 And y <= .ScaleHeight) Then
            If (mbLeftMouseDown Or mbLeftWasDown) And Button = vbLeftButton Then
                mbLeftMouseDown = False
                If (m_eStyle <> Options_1 And m_eStyle <> Options_2) Then DrawButtonState giRAISED
                MakeClick
            End If
        ElseIf mbLeftWasDown And Button = vbLeftButton Then
            mbLeftWasDown = False
            Flatten
        Else
            mbMouseOver = False     'Set mbMouseOver = False so that SetCapture will be called again.  For some reason, capture is released during MouseUp
        End If
    End With
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
    DrawButtonState miCurrentState
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim picMine As Picture
    Dim picMine2 As Picture
    Dim picMine3 As Picture
    
    On Error Resume Next
    ' Read in the properties that have been saved into the PropertyBag...
    With PropBag
        Set picMine = PropBag.ReadProperty(msPICTURE_NAME, UserControl.Picture) ' Read Picture property value
        Set picMine3 = PropBag.ReadProperty(msPICTURE_HOVER, Nothing)
        Set picMine2 = PropBag.ReadProperty(msPICTURE_SELECTED, Nothing) ' Read Picture property value
        
        If Not (picMine Is Nothing) Then Set Picture = picMine                              ' Use existing picture (This is used only if URL is empty)
        If Not (picMine3 Is Nothing) Then Set PictureHover = picMine3
        If Not (picMine2 Is Nothing) Then Set PictureSelected = picMine2
    End With
    
    BackColor = PropBag.ReadProperty(msBACK_COLOR_NAME, vbButtonFace)
    BackColorHover = PropBag.ReadProperty(msBACK_COLOR_HOVER, vbButtonFace)
    BackColorSelected = PropBag.ReadProperty(msBACK_COLOR_SELECTED, vbButtonFace)
    
    ForeColor = PropBag.ReadProperty(msFORE_COLOR_NAME, vbBlack)
    ForeColorHover = PropBag.ReadProperty(msFORE_COLOR_HOVER, vbBlack)
    ForeColorSelected = PropBag.ReadProperty(msFORE_COLOR_SELECTED, vbBlack)
    
    Set Font = PropBag.ReadProperty(msFONT, UserControl.Font)
    Set FontHover = PropBag.ReadProperty(msFONT_HOVER, UserControl.Font)
    Set FontSelected = PropBag.ReadProperty(msFONT_SELECTED, UserControl.Font)
    
    Unselectable = PropBag.ReadProperty(msUNSELECTABLE, True)
    
    AmbientColor = PropBag.ReadProperty(msAMBIENT_COLOR, False)
    Enabled = PropBag.ReadProperty(msENABLED_NAME, True)
    MaskColor = PropBag.ReadProperty(msMASK_COLOR_NAME, vbWhite)
    UseMaskColor = PropBag.ReadProperty(msUSE_MASK_COLOR_NAME, False)
    PlaySounds = PropBag.ReadProperty(msPLAY_SOUNDS_NAME, False)
    ToolTipText = PropBag.ReadProperty(msTOOL_TIP_TEXT_NAME, "")
    Style = PropBag.ReadProperty(msSTYLE, Cool)
    ShowBorder = PropBag.ReadProperty(msSHOW_BORDER, False)
    LineColorNormal = PropBag.ReadProperty(msLINE_COLOR, vbBlack)
    LineColorHover = PropBag.ReadProperty(msLINE_COLOR_HOVER, vbYellow)
    LineColorSelected = PropBag.ReadProperty(msLINE_COLOR_SELECTED, vbWhite)
    Caption = PropBag.ReadProperty(msCAPTION, "Button 1")
    Alignment = PropBag.ReadProperty(msALIGNMENT, vbAlignBottom)
    AccessKey = PropBag.ReadProperty(msACCESS_KEY, "")
    m_iGroupNumber = PropBag.ReadProperty(msGROUP_NUMBER, 0)
    CanGetFocus = PropBag.ReadProperty(msCANGETFOCUS, False)
    
    InstanciateToolTipsWindow
    mbPropertiesLoaded = True
End Sub

Private Sub UserControl_Terminate()
    Set moDrawTool = Nothing
    DeleteObject mlhHalftonePal
    glToolsCount = glToolsCount - 1
    UnSubClass
    If gbToolTipsInstanciated And glToolsCount = 0 Then
        DestroyWindow gHWndToolTip
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    'Write the properties of the button to the property bag.
    PropBag.WriteProperty msPICTURE_NAME, m_picPicture
    PropBag.WriteProperty msPICTURE_HOVER, m_picPictureHover
    PropBag.WriteProperty msPICTURE_SELECTED, m_picPictureSelected
    
    PropBag.WriteProperty msBACK_COLOR_NAME, m_cBackColor
    PropBag.WriteProperty msBACK_COLOR_SELECTED, m_cBackColorSelected
    PropBag.WriteProperty msBACK_COLOR_HOVER, m_cBackColorHover
    
    PropBag.WriteProperty msFORE_COLOR_NAME, m_cForeColor, vbBlack
    PropBag.WriteProperty msFORE_COLOR_HOVER, m_cForeColorHover, vbBlack
    PropBag.WriteProperty msFORE_COLOR_SELECTED, m_cForeColorSelected, vbBlack
    
    PropBag.WriteProperty msLINE_COLOR, m_cLineColor, vbBlack
    PropBag.WriteProperty msLINE_COLOR_HOVER, m_cLineColorHover, vbBlue
    PropBag.WriteProperty msLINE_COLOR_SELECTED, m_cLineColorSelected, vbWhite
    PropBag.WriteProperty msSHOW_BORDER, m_bShowBorder, False
    
    PropBag.WriteProperty msFONT, m_eFont
    PropBag.WriteProperty msFONT_HOVER, m_eFontHover
    PropBag.WriteProperty msFONT_SELECTED, m_eFontSelected
    
    PropBag.WriteProperty msUNSELECTABLE, m_bUnselectable, True
    
    PropBag.WriteProperty msAMBIENT_COLOR, m_bAmbientColor
    PropBag.WriteProperty msENABLED_NAME, m_bEnabled
    PropBag.WriteProperty msMASK_COLOR_NAME, m_clrMaskColor
    PropBag.WriteProperty msUSE_MASK_COLOR_NAME, m_bUseMaskColor
    PropBag.WriteProperty msPLAY_SOUNDS_NAME, m_bPlaySounds
    PropBag.WriteProperty msTOOL_TIP_TEXT_NAME, m_sToolTipText
    PropBag.WriteProperty msSTYLE, m_eStyle, Cool
    PropBag.WriteProperty msCAPTION, m_sCaption, "Button 1"
    PropBag.WriteProperty msALIGNMENT, m_eAlignment, vbAlignBottom
    PropBag.WriteProperty msACCESS_KEY, m_sAccessKey, ""
    PropBag.WriteProperty msCANGETFOCUS, m_bCanGetFocus, False
    
    PropBag.WriteProperty msGROUP_NUMBER, m_iGroupNumber
    
End Sub

Private Sub UserControl_Resize()
    'Reevaluate and Repaint coordinates
    PositionChanged
    DrawButtonState miCurrentState
End Sub

'**********************
'Public Properties
'**********************
'This property is used to identify other CButtons.  This will be used with the Options_1 and Options_2 buttons to figure out if there are any other selected
'Option buttons on the container.
Public Property Get Identification() As String
Attribute Identification.VB_Description = "For use in identifying other buttons of the same type."
Attribute Identification.VB_MemberFlags = "40"
    Identification = "CButton"
End Property

'This property is used to group the option button controls together. By adding a group number, the user can identify a group of buttons that
'will interact with one another
Public Property Let GroupNumber(iGroup As Integer)
    m_iGroupNumber = iGroup
    PropertyChanged msGROUP_NUMBER
End Property
Public Property Get GroupNumber() As Integer
Attribute GroupNumber.VB_Description = "For grouping options buttons together.  This will have the button automatically unselect other buttons in the same group when in Options style"
Attribute GroupNumber.VB_ProcData.VB_Invoke_Property = ";Behavior"
    GroupNumber = m_iGroupNumber
End Property

'This property lets the user create an access key even if the button doesn't have a caption
Public Property Let AccessKey(ByVal sKey As String)
Attribute AccessKey.VB_Description = "Enter a hot-key here if you have no caption."
Attribute AccessKey.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    If InStr(m_sCaption, "&") > 0 Then
        sKey = Mid(m_sCaption, InStr(m_sCaption, "&") + 1, 1)
    End If
    m_sAccessKey = Left(sKey, 1)
    UserControl.AccessKeys = m_sAccessKey
    PropertyChanged (msACCESS_KEY)
End Property
Public Property Get AccessKey() As String
    AccessKey = m_sAccessKey
End Property

Public Property Let ToolTipText(ByVal sToolTip As String)
Attribute ToolTipText.VB_Description = "ToolTip to display when the mouse is over the button."
Attribute ToolTipText.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    m_sToolTipText = sToolTip
    'If this property gets called with more than an empty string, I know for sure that there is not a ToolTipText extender property
    If Len(sToolTip) <> 0 Then mbToolTipNotInExtender = True
    PropertyChanged (msTOOL_TIP_TEXT_NAME)
End Property
Public Property Get ToolTipText() As String
    ToolTipText = m_sToolTipText
End Property

Public Property Let PlaySounds(ByVal bPlaySounds As Boolean)
    'The following line of code ensures that the integer value of the boolean parameter is either
    '0 or -1.  It is known that Access 97 will set the boolean's value to 255 for true. 'In this case a P-Code compiled VB5 built
    'OCX will return True for the expression (Not [boolean variable that ='s 255]).  This line ensures the reliability of boolean operations
    If CBool(bPlaySounds) Then bPlaySounds = True Else bPlaySounds = False
    m_bPlaySounds = bPlaySounds
    PropertyChanged (msPLAY_SOUNDS_NAME)
End Property
Public Property Get PlaySounds() As Boolean
Attribute PlaySounds.VB_Description = "Will play a sound when button is hovered over, or clicked."
Attribute PlaySounds.VB_ProcData.VB_Invoke_Property = ";Behavior"
    PlaySounds = m_bPlaySounds
End Property

Public Property Let MaskColor(ByVal clrMaskColor As OLE_COLOR)
    'If there is a valid picture, repaint control
    m_clrMaskColor = clrMaskColor
    If m_bUseMaskColor And Not m_picPicture Is Nothing Then DrawButtonState miCurrentState
    PropertyChanged (msMASK_COLOR_NAME)
End Property
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Mask color to use on the image."
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
    MaskColor = m_clrMaskColor
End Property

Public Property Let UseMaskColor(ByVal bUseMaskColor As Boolean)
    'If true, use the mask color.  Mask color only applies to bitmaps not icons. Repaint control Validate whether correct picture type is provided
    If CBool(bUseMaskColor) Then bUseMaskColor = True Else bUseMaskColor = False
    m_bUseMaskColor = bUseMaskColor
    If Not m_picPicture Is Nothing Then
        If m_picPicture.Type = vbPicTypeBitmap Then DrawButtonState miCurrentState
    End If
    PropertyChanged (msUSE_MASK_COLOR_NAME)
End Property
Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Whether or not to use the MaskColor on the picture."
Attribute UseMaskColor.VB_ProcData.VB_Invoke_Property = ";Behavior"
    UseMaskColor = m_bUseMaskColor
End Property

Public Property Let BackColor(ByVal clrBackColor As OLE_COLOR)
    'Control will be repainted because VB will fire paint event
    If m_bAmbientColor And Not mbFromAmbient And clrBackColor <> UserControl.Ambient.BackColor Then m_bAmbientColor = False
    If m_cBackColorHover = UserControl.BackColor Then m_cBackColorHover = clrBackColor
    UserControl.BackColor = clrBackColor
    m_cBackColor = clrBackColor
    DrawButtonState miCurrentState
    PropertyChanged (msBACK_COLOR_NAME)
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Background Color of the button"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_cBackColor
End Property

Public Property Let BackColorHover(ByVal clrBackColor As OLE_COLOR)
    'Control will be repainted because VB will fire paint event
    m_cBackColorHover = clrBackColor
    DrawButtonState miCurrentState
    PropertyChanged (msBACK_COLOR_HOVER)
End Property
Public Property Get BackColorHover() As OLE_COLOR
Attribute BackColorHover.VB_Description = "Background color of the button when the mouse is over it."
Attribute BackColorHover.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorHover = m_cBackColorHover
End Property

Public Property Let BackColorSelected(ByVal clrBackColor As OLE_COLOR)
    'Control will be repainted because VB will fire paint event
    m_cBackColorSelected = clrBackColor
    If (m_eStyle = Outline Or m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected Then
        DrawButtonState miCurrentState
    End If
    PropertyChanged (msBACK_COLOR_SELECTED)
End Property
Public Property Get BackColorSelected() As OLE_COLOR
Attribute BackColorSelected.VB_Description = "Background color of the button when it is selected.  Only applies in Options_1 or Options_2 styles."
Attribute BackColorSelected.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorSelected = m_cBackColorSelected
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Picture to display on the button."
Attribute Picture.VB_ProcData.VB_Invoke_Property = "StandardPicture;Appearance"
    Set Picture = m_picPicture
End Property
Public Property Set Picture(ByVal picPicture As Picture)
    'Validate what kind of picture is passed Only allow bitmaps and icons If not in runtime display message that UseMaskColor can't be
    'used with icons, if picture is icon. If picture is icon, make sure UseMaskColor is false Paint Control
    If Not picPicture Is Nothing And Not picPicture = 0 Then
        With picPicture
            If (.Type <> vbPicTypeBitmap) And (.Type <> vbPicTypeNone) And (.Type <> vbPicTypeIcon) Then
                If Not UserControl.Ambient.UserMode Then
                    MsgBox LoadResString(giINVALID_PIC_TYPE), vbOKOnly, UserControl.Name
                End If
                Exit Property
            End If
            
            'Now make sure that the different pictures are all of the same type.
            If Not picPicture Is Nothing And Not m_picPictureHover Is Nothing And Not m_picPictureSelected Is Nothing Then
                If (.Type <> m_picPictureHover.Type) Or (.Type <> m_picPictureSelected.Type) Then
                    If Not UserControl.Ambient.UserMode Then MsgBox LoadResString(giMISMATCH_PIC_TYPE), vbOKOnly, UserControl.Name
                    Exit Property
                End If
            End If
            
        End With
    End If
    
    'Make sure the handle to the picture is valid
    If (Not picPicture Is Nothing) Then If (picPicture.Handle = 0) Then Set picPicture = Nothing
    Set m_picPicture = picPicture
    
    PositionChanged
    DrawButtonState miCurrentState
    
    PropertyChanged (msPICTURE_NAME)
    
End Property

Public Property Get PictureSelected() As Picture
Attribute PictureSelected.VB_Description = "Picture to display when the button is selected.  Only applies for Outline style."
Attribute PictureSelected.VB_ProcData.VB_Invoke_Property = "StandardPicture;Appearance"
    Set PictureSelected = m_picPictureSelected
End Property
Public Property Set PictureSelected(ByVal picPicture As Picture)
    'Validate what kind of picture is passed Only allow bitmaps and icons If not in runtime display message that UseMaskColor can't be
    'used with icons, if picture is icon. If picture is icon, make sure UseMaskColor is false Paint Control
    If Not picPicture Is Nothing And Not picPicture = 0 Then
        With picPicture
            If (.Type <> vbPicTypeBitmap) And (.Type <> vbPicTypeNone) And (.Type <> vbPicTypeIcon) Then
                If Not UserControl.Ambient.UserMode Then
                    MsgBox LoadResString(giINVALID_PIC_TYPE), vbOKOnly, UserControl.Name
                End If
                Exit Property
            End If
            
            If Not picPicture Is Nothing And Not m_picPicture Is Nothing And Not m_picPictureHover Is Nothing Then
                If (.Type <> m_picPicture.Type) Or (.Type <> m_picPictureHover.Type) Then
                    If Not UserControl.Ambient.UserMode Then
                        MsgBox LoadResString(giMISMATCH_PIC_TYPE), vbOKOnly, UserControl.Name
                    End If
                    Exit Property
                End If
            End If
            
        End With
    End If
    
    If (Not picPicture Is Nothing) Then If (picPicture.Handle = 0) Then Set picPicture = Nothing
    Set m_picPictureSelected = picPicture
    
    PositionChanged
    DrawButtonState miCurrentState
    PropertyChanged (msPICTURE_NAME)
    
End Property

Public Property Get PictureHover() As Picture
Attribute PictureHover.VB_Description = "Picture to display on the button when the mouse is over it."
Attribute PictureHover.VB_ProcData.VB_Invoke_Property = "StandardPicture;Appearance"
    Set PictureHover = m_picPictureHover
End Property
Public Property Set PictureHover(ByVal picPicture As Picture)
    'Validate what kind of picture is passed Only allow bitmaps and icons If not in runtime display message that UseMaskColor can't be
    'used with icons, if picture is icon. If picture is icon, make sure UseMaskColor is false Paint Control
    If Not picPicture Is Nothing And Not picPicture = 0 Then
        With picPicture
            If (.Type <> vbPicTypeBitmap) And (.Type <> vbPicTypeNone) And (.Type <> vbPicTypeIcon) Then
                If Not UserControl.Ambient.UserMode Then MsgBox LoadResString(giINVALID_PIC_TYPE), vbOKOnly, UserControl.Name
                Exit Property
            End If
            
            If Not picPicture Is Nothing And Not m_picPictureSelected Is Nothing And Not m_picPicture Is Nothing Then
                If (.Type <> m_picPicture.Type) Or (.Type <> m_picPictureSelected.Type) Then
                    If Not UserControl.Ambient.UserMode Then MsgBox LoadResString(giMISMATCH_PIC_TYPE), vbOKOnly, UserControl.Name
                    Exit Property
                End If
            End If
        
        End With
    End If
    
    If (Not picPicture Is Nothing) Then If (picPicture.Handle = 0) Then Set picPicture = Nothing
    Set m_picPictureHover = picPicture
    
    PositionChanged
    DrawButtonState miCurrentState
    PropertyChanged (msPICTURE_NAME)
    
End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)
    'If button is raised, flatten it
    'Draw disabled appearance of picture
    Dim lresult As Long
    If CBool(bEnabled) Then bEnabled = True Else bEnabled = False
    UserControl.Enabled = bEnabled
    m_bEnabled = bEnabled
    If bEnabled Then
        DrawButtonState giFLATTENED
    Else
        If miCurrentState = giRAISED Then
            'Call flatten as if button does not have focus
            If mbGotFocus Then
                'Get rid of focus
                mbGotFocus = False
            End If
            Flatten
        End If
        DrawButtonState giDISABLED
    End If
    PropertyChanged (msENABLED_NAME)
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Indicates whether the button is enabled or not."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property

'This property will tell the button to stay the same color as the ambient colors.  Unfortunately,
'this property doesn't work when the parent is a picture box -- it still sticks with the backcolor of the
'form or usercontrol the button is on.  This probably could be fixed by using the .ContainerHwnd property
'of the Usercontrol...
Public Property Let AmbientColor(bAColor As Boolean)
Attribute AmbientColor.VB_Description = "Tells the button to stay use the same back color as the parent form/control.  (This does not work in a picture box)."
Attribute AmbientColor.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    
    m_bAmbientColor = bAColor
    
    If bAColor Then
        mbFromAmbient = True
        
        'Try to get the backcolor from the parent
        On Error Resume Next
        BackColor = UserControl.Ambient.BackColor
        BackColorSelected = m_cBackColor
        mbFromAmbient = False
    Else
        mbFromAmbient = True
        BackColor = m_cBackColor
        mbFromAmbient = False
    End If
    
    PropertyChanged (msAMBIENT_COLOR)
End Property
Public Property Get AmbientColor() As Boolean
    AmbientColor = m_bAmbientColor
End Property

'This property is for the Style of the button.  The Available Styles are: Cool, Flat, ThreeD, Outline, Options_1, Options_2
Public Property Let Style(eNewStyle As Styles)
    m_eStyle = eNewStyle
    PositionChanged
    DrawButtonState miCurrentState
    PropertyChanged (msSTYLE)
End Property
Public Property Get Style() As Styles
Attribute Style.VB_Description = "The style of the button."
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Style = m_eStyle
End Property

'This property will allow the programmer to prevent the user from Unselecting the Currently Selected button.
'This only applies to one of the Options styles, and it does not prevent the button from being unselected via
'other means -- for instance programmatically, or by selecting a different button.
Public Property Let Unselectable(bUnselect As Boolean)
    m_bUnselectable = bUnselect
    PropertyChanged msUNSELECTABLE
End Property
Public Property Get Unselectable() As Boolean
Attribute Unselectable.VB_Description = "Indicates whether or not the user can unselect the button when it Options style.  This will not prevent the button from being unselected, only prevent the user from clicking on a selected button to unselect it."
Attribute Unselectable.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Unselectable = m_bUnselectable
End Property

'This property is only applicable for Outline Style.  It will set the line color of the Outline when the button is
'in its normal state (not hovered over, or clicked, etc).  It will also only apply if the ShowBorder property is true
Public Property Let LineColorNormal(cNewColor As OLE_COLOR)
    m_cLineColor = cNewColor
    PropertyChanged (msLINE_COLOR)
End Property
Public Property Get LineColorNormal() As OLE_COLOR
Attribute LineColorNormal.VB_Description = "Color of the outline.  Only applies for Outline style."
Attribute LineColorNormal.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
    LineColorNormal = m_cLineColor
End Property

'This property only applies to the Outline style.  It will determine whether or not the button has a Border around it.
Public Property Let ShowBorder(bShow As Boolean)
    m_bShowBorder = bShow
    PropertyChanged (msSHOW_BORDER)
    DrawButtonState miCurrentState
End Property
Public Property Get ShowBorder() As Boolean
Attribute ShowBorder.VB_Description = "Show an outline around the button when it is not selected.  Only applies for Outline style."
Attribute ShowBorder.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowBorder = m_bShowBorder
End Property

'This property only applies to the Outline style.
Public Property Let LineColorHover(cNewColor As OLE_COLOR)
    m_cLineColorHover = cNewColor
    PropertyChanged (msLINE_COLOR_HOVER)
End Property
Public Property Get LineColorHover() As OLE_COLOR
Attribute LineColorHover.VB_Description = "Color of the outline when the mouse is over the button.  Only applies for Outline style"
Attribute LineColorHover.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
    LineColorHover = m_cLineColorHover
End Property

'This property only applies to the Outline style
Public Property Let LineColorSelected(cNewColor As OLE_COLOR)
    m_cLineColorSelected = cNewColor
    PropertyChanged (msLINE_COLOR_SELECTED)
End Property
Public Property Get LineColorSelected() As OLE_COLOR
Attribute LineColorSelected.VB_Description = "Outline color when button is selected.  Only applies for Outline style."
Attribute LineColorSelected.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
    LineColorSelected = m_cLineColorSelected
End Property

'The property to determine if the button is selected or not.  This property only applies to the Options styles.
Public Property Let Selected(bSelected As Boolean)
    Dim bCancel As Boolean
    
    If bSelected = m_bSelected Then Exit Property
    If Not (m_eStyle = Options_1 Or m_eStyle = Options_2 Or m_eStyle = Outline) Then Exit Property
    
    If bSelected Then
        UserControl_Click
    Else
        RaiseEvent ButtonUnSelected(bCancel)
        If bCancel Then Exit Property
    
        m_bSelected = bSelected
        miCurrentState = giFLATTENED
        DrawButtonState miCurrentState
    End If
    
End Property
Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Indicates whether or not the button is selected."
Attribute Selected.VB_MemberFlags = "400"
    Selected = m_bSelected
End Property

'The color of the caption in the Normal state (not selected or hovered)
Public Property Let ForeColor(cNewColor As OLE_COLOR)
    m_cForeColor = cNewColor
    If m_cForeColorHover = UserControl.ForeColor Then m_cForeColorHover = cNewColor
    UserControl.ForeColor = cNewColor
    PropertyChanged (msFORE_COLOR_NAME)
    DrawButtonState miCurrentState
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Color of the button caption"
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = m_cForeColor
End Property

'The color of the caption in the Hover state
Public Property Let ForeColorHover(cNewColor As OLE_COLOR)
Attribute ForeColorHover.VB_Description = "Color of the button caption when the mouse is over it"
Attribute ForeColorHover.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_cForeColorHover = cNewColor
    PropertyChanged (msFORE_COLOR_HOVER)
    DrawButtonState miCurrentState
End Property
Public Property Get ForeColorHover() As OLE_COLOR
    ForeColorHover = m_cForeColorHover
End Property

'The color of the caption in the Selected state
Public Property Let ForeColorSelected(cNewColor As OLE_COLOR)
    m_cForeColorSelected = cNewColor
    If m_bSelected Then UserControl.ForeColor = cNewColor
    PropertyChanged (msFORE_COLOR_SELECTED)
    DrawButtonState miCurrentState
End Property
Public Property Get ForeColorSelected() As OLE_COLOR
Attribute ForeColorSelected.VB_Description = "Color of the caption when the button is selected.  Only applies in Options style mode."
Attribute ForeColorSelected.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorSelected = m_cForeColorSelected
End Property

Public Property Let Caption(sCap As String)
    Dim HotKey As String
    Dim sTemp As String
    
    m_sCaption = sCap
    If InStr(m_sCaption, "&") > 0 Then
        HotKey = Mid(m_sCaption, InStr(m_sCaption, "&") + 1, 1)
        AccessKey = HotKey
    ElseIf Trim(m_sCaption) = "" Then
        HotKey = ""
        AccessKey = ""
    End If
    
    PropertyChanged (msCAPTION)
    DrawButtonState miCurrentState
End Property
Public Property Get Caption() As String
Attribute Caption.VB_Description = "The Caption of the button control"
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = m_sCaption
End Property

'The alignment of the Caption.
Public Property Let Alignment(eAlign As AlignConstants)
Attribute Alignment.VB_Description = "Alignment of the caption in relation to the picture."
Attribute Alignment.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_eAlignment = eAlign
    PropertyChanged (msALIGNMENT)
    DrawButtonState miCurrentState
End Property
Public Property Get Alignment() As AlignConstants
    Alignment = m_eAlignment
End Property

'The font of the caption in Normal mode (not selected or hovered)
Public Property Set Font(eNew As StdFont)
    Set m_eFont = eNew
    PropertyChanged (msFONT)
    DrawButtonState miCurrentState
End Property
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "The font of the button"
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set Font = m_eFont
End Property
    
'The font of the caption when the button is Hovered over
Public Property Get FontHover() As StdFont
Attribute FontHover.VB_Description = "The font of the button when the mouse is over it"
Attribute FontHover.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set FontHover = m_eFontHover
End Property
Public Property Set FontHover(eNew As StdFont)
    Set m_eFontHover = eNew
    PropertyChanged (msFONT_HOVER)
    DrawButtonState miCurrentState
End Property

'The font of the caption when the button is Selected (only applies to Options styles)
Public Property Get FontSelected() As StdFont
Attribute FontSelected.VB_Description = "The font of the button when it is selected.  Only applies when in Options style"
Attribute FontSelected.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set FontSelected = m_eFontSelected
End Property
Public Property Set FontSelected(eNew As StdFont)
    Set m_eFontSelected = eNew
    PropertyChanged (msFONT_SELECTED)
    DrawButtonState miCurrentState
End Property

'This will prevent the button from getting the focus (and therefore, prevent it from painting the FocusRect every time it is clicked)
Public Property Let CanGetFocus(bCanGetFocus As Boolean)
Attribute CanGetFocus.VB_Description = "Indicates whether or not te button can get focus.  This will control whether or not a Focus Rect is drawn on the button."
Attribute CanGetFocus.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    m_bCanGetFocus = bCanGetFocus
    PropertyChanged msCANGETFOCUS
    DrawButtonState miCurrentState
End Property
Public Property Get CanGetFocus() As Boolean
    CanGetFocus = m_bCanGetFocus
End Property
    

'*************************
'Private Procedures
'*************************

Private Sub MakeClick()
    '-------------------------------------------------------------------------
    'Purpose:   Raise a Click event to container, play sound
    '-------------------------------------------------------------------------
    '-----------------------------------------
    '- Added for sound support
    '-----------------------------------------
    If m_bPlaySounds Then PlaySound EVENT_MENU_COMMAND, 0, SND_SYNC
    '-----------------------------------------
    RaiseEvent Click
End Sub

Private Sub MouseOver()
    '-------------------------------------------------------------------------
    'Purpose:   Call whenever the mouse is over the button and button needs raised appearance and capture set
    '-------------------------------------------------------------------------
    If miCurrentState <> giRAISED Then DrawButtonState giRAISED
    
    If Not mbMouseOver Then
        Capture True
        mbMouseOver = True
        '-----------------------------------------
        '- Added for sound support
        '-----------------------------------------
        If Not mbEnterOnce Then
            RaiseEvent PopUp
            If m_bPlaySounds Then PlaySound EVENT_MENU_POPUP, 0, SND_SYNC
            mbEnterOnce = True
        End If
        '-----------------------------------------
    End If
End Sub

Private Sub Flatten()
    '-------------------------------------------------------------------------
    'Purpose:   Call whenever the mouse is off the control and capture needs released and button needs flattened appearance
    '-------------------------------------------------------------------------
    If mbMouseOver Then Capture False
    mbMouseOver = False
    If (Not mbGotFocus) And miCurrentState <> giFLATTENED Then DrawButtonState giFLATTENED
    '-----------------------------------------
    '- Added for sound support
    '-----------------------------------------
    '   PlaySound EVENT_MENU_POPUP, 0, SND_SYNC
    mbEnterOnce = False
    '-----------------------------------------
End Sub

Private Sub AddTool(hwnd As Long)
    '-------------------------------------------------------------------------
    'Purpose:   Add a tool to the ToolTips object
    'In:        [hWnd] hWnd of Tool being added
    '-------------------------------------------------------------------------
                   
    Dim ti As TOOLINFO
    
    With ti
        .cbSize = Len(ti)
        .uId = hwnd
        .hwnd = hwnd
        .hinst = App.hInstance
        .uFlags = TTF_IDISHWND
        .lpszText = LPSTR_TEXTCALLBACK
    End With
    
    SendMessage gHWndToolTip, TTM_ADDTOOL, 0, ti
    SendMessage gHWndToolTip, TTM_ACTIVATE, 1, ByVal hwnd
    Exit Sub
End Sub

Private Sub InstanciateToolTipsWindow()
    '-------------------------------------------------------------------------
    'Purpose:   Instanciate needed collections.
    '           Create ToolTips Class window
    '-------------------------------------------------------------------------
    glToolsCount = glToolsCount + 1
    If UserControl.Ambient.UserMode Then
        If Not gbToolTipsInstanciated Then
            gbToolTipsInstanciated = True
            InitCommonControls
            gHWndToolTip = CreateWindowEX(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, 0, _
                                                                              CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, _
                                                                              App.hInstance, ByVal 0)
            SendMessage gHWndToolTip, TTM_ACTIVATE, 1, ByVal 0
            
            #If DEBUGSUBCLASS Then
                If goWindowProcHookCreator Is Nothing Then Set goWindowProcHookCreator = CreateObject("DbgWindowProc.WindowProcHookCreator")
            #End If
        End If
        
        'Sub class this code module to receive window messages for the Usercontrol
        SubClass UserControl.hwnd
        'Add Register Usercontrol with ToolTip window
        AddTool UserControl.hwnd
    End If
    
End Sub

Private Sub SubClass(hwnd)
    '-------------------------------------------------------------------------
    'Purpose:   Subclass control so that the ToolTip Need text message can be handled.  Store handle of class as UserData of control window
    '-------------------------------------------------------------------------
    Dim lresult As Long
    
    UnSubClass
    
    #If DEBUGSUBCLASS Then
        'If in debug, SubClass window using address of WindowProcHook Let WindowProcHook CallWindowProc at address of my function
        'if in run mode but call the previous address if in break mode this prevents crashes in break mode
        Set moProcHook = goWindowProcHookCreator.CreateWindowProcHook
        With moProcHook
            .SetMainProc AddressOf SubWndProc
            mWndProcNext = SetWindowLong(hwnd, GWL_WNDPROC, CLng(.ProcAddress))
            .SetDebugProc mWndProcNext
        End With
    #Else
        mWndProcNext = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubWndProc)
    #End If
    
    If mWndProcNext Then
        mHWndSubClassed = hwnd
        lresult = SetWindowLong(hwnd, GWL_USERDATA, ObjPtr(Me))
    End If
End Sub

Private Sub UnSubClass()
    '-------------------------------------------------------------------------
    'Purpose:   Unsubclass control
    '-------------------------------------------------------------------------
    If mWndProcNext Then
        SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcNext
        mWndProcNext = 0
        
        #If DEBUGSUBCLASS Then
            Set moProcHook = Nothing
        #End If
        
    End If
End Sub

Private Sub Capture(bCapture As Boolean)
    '-------------------------------------------------------------------------
    'Purpose:   Is the only place where setcapture and releasecapture are called setcapture may be called after mouse clicks because VB seems to
    '           release capture on my behalf.
    '-------------------------------------------------------------------------
    If bCapture Then
        SetCapture UserControl.hwnd
    Else
        ReleaseCapture
    End If
End Sub

Private Sub PositionChanged()
    '-------------------------------------------------------------------------
    'Purpose:   Calculate needed coordinates for painting the control
    '-------------------------------------------------------------------------
    On Error GoTo PositionChangedError
    With mudtButtonRect
        .Bottom = UserControl.ScaleHeight
        .Right = UserControl.ScaleWidth
    End With
    If Not m_picPicture Is Nothing Then
        With mudtPicturePoint
            .x = CLng(UserControl.ScaleX(m_picPicture.Width, vbHimetric, vbPixels))
            .y = CLng(UserControl.ScaleY(m_picPicture.Height, vbHimetric, vbPixels))
        End With
        
        With mudtPictureRect
            .Left = CLng((mudtButtonRect.Right - mudtPicturePoint.x) / 2)
            .Top = CLng((mudtButtonRect.Bottom - mudtPicturePoint.y) / 2)
            .Right = .Left + mudtPicturePoint.x
            .Bottom = .Top + mudtPicturePoint.y
        End With
    End If
    Exit Sub
PositionChangedError:
    Exit Sub
End Sub

'-------------------------------------------------------------------------
'Purpose:   Draw the button in whatever state it needs to be in.
'-------------------------------------------------------------------------
Private Sub DrawButtonState(iState As Integer)
    Dim lhbmMemory As Long
    Dim lhbmMemoryOld As Long
    Dim lhdcMem As Long
    Dim lBackColor As Long
    Dim lForeColor As Long
    Dim udtPictureRect As RECT
    Dim bUseMask As Boolean
    Dim lhPal As Long
    Dim lhPalOld As Long
    Dim lhbrBack As Long
    Dim bHaveAmbientPalette As Boolean
    Dim lresult As Long
    Dim HeightUsed As Long
    Dim TextRect As RECT
    Dim lTextAlign As Long
    
    On Error GoTo DrawButtonState_Error
    If Not mbPropertiesLoaded Then Exit Sub
    
     If (m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected Then iState = giSUNKEN
     miCurrentState = iState
     udtPictureRect = mudtPictureRect
     
     On Error Resume Next        'Error will occur if the Ambient.Palette is not supported
     bHaveAmbientPalette = (Not UserControl.Ambient.Palette Is Nothing)
     If Err.Number <> 0 Then bHaveAmbientPalette = False
     Err.Clear
     If bHaveAmbientPalette Then
         'If the Palette or hPal property fails resume next and use the halftone palette
         lhPal = UserControl.Ambient.Palette.hPal
         If lhPal = 0 Then lhPal = mlhHalftonePal
         Err.Clear
     Else
         lhPal = mlhHalftonePal    'If there is no specified palette use the halftone palette.
     End If
     
     On Error GoTo DrawButtonState_Error         'If button is sunken offset the picture coordinates so that the picture looks like it is in sunken perspective
     If iState = giSUNKEN And m_eStyle <> Flat Then
         With udtPictureRect
             .Right = .Right + glSUNKEN_OFFSET
             .Left = .Left + glSUNKEN_OFFSET
             .Top = .Top + glSUNKEN_OFFSET
             .Bottom = .Bottom + glSUNKEN_OFFSET
         End With
     End If
     
     'Create memory DC and bitmap to do all of the painting work
     lhdcMem = CreateCompatibleDC(UserControl.hdc)
     lhbmMemory = CreateCompatibleBitmap(UserControl.hdc, mudtButtonRect.Right, mudtButtonRect.Bottom)
     lhbmMemoryOld = SelectObject(lhdcMem, lhbmMemory)
     lhPalOld = SelectPalette(lhdcMem, lhPal, True)
     RealizePalette lhdcMem
     
     'fill the memory DC with the appropriate background color
     If (m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected Then
         OleTranslateColor m_cBackColorSelected, 0, lBackColor
     ElseIf iState = giRAISED Or mbLeftMouseDown And Not m_bSelected Then
         OleTranslateColor m_cBackColorHover, 0, lBackColor
     Else
         OleTranslateColor m_cBackColor, 0, lBackColor
     End If
     
     SetBkColor lhdcMem, lBackColor
     lhbrBack = CreateSolidBrush(lBackColor)
     FillRect lhdcMem, mudtButtonRect, lhbrBack
     
     'First have to set up the font and forecolor (font must be established before getting TextWidth and TextHeight)
    If m_sCaption <> "" Then
        With UserControl
            If m_bEnabled Then
                If (m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected Then
                    .ForeColor = m_cForeColorSelected
                    Set .Font = m_eFontSelected
                ElseIf (iState = giRAISED Or mbLeftMouseDown And Not m_bSelected) And Not Ambient.UserMode = False Then
                    .ForeColor = m_cForeColorHover
                    Set .Font = m_eFontHover
                Else
                    .ForeColor = m_cForeColor
                    Set .Font = m_eFont
                End If
            Else
                .ForeColor = vbGrayText
            End If
        End With
    End If
    
     'Get the Picture and text dimensions
     TextRect = mudtButtonRect
     TextRect.Left = TextRect.Left + 2
     TextRect.Right = TextRect.Right - 2
     
     'Align the Rects for the text and the picture
     Select Case m_eAlignment
         Case vbAlignBottom
             TextRect.Top = TextRect.Bottom - UserControl.TextHeight(m_sCaption) - 2
             udtPictureRect.Bottom = TextRect.Top
             If (iState = giSUNKEN And Not m_eStyle = Flat) Or ((m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected) Then
                 TextRect.Left = TextRect.Left + glSUNKEN_OFFSET
                 TextRect.Right = TextRect.Right + glSUNKEN_OFFSET
                 TextRect.Top = TextRect.Top + glSUNKEN_OFFSET
                 TextRect.Bottom = TextRect.Bottom + glSUNKEN_OFFSET
             End If
            lTextAlign = DT_CENTER
            
         Case vbAlignTop
             TextRect.Top = 2
             udtPictureRect.Top = UserControl.TextHeight(m_sCaption) + 2
             If (iState = giSUNKEN And Not m_eStyle = Flat) Or ((m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected) Then
                 TextRect.Left = TextRect.Left + glSUNKEN_OFFSET
                 TextRect.Right = TextRect.Right + glSUNKEN_OFFSET
                 TextRect.Top = TextRect.Top + glSUNKEN_OFFSET
                 TextRect.Bottom = TextRect.Bottom + glSUNKEN_OFFSET
             End If
            lTextAlign = DT_CENTER
         
         Case vbAlignRight
             udtPictureRect.Left = 0
             udtPictureRect.Right = 30
             TextRect.Left = udtPictureRect.Right + 4
             TextRect.Right = UserControl.ScaleWidth - 6
             TextRect.Top = (TextRect.Bottom / 2) - (UserControl.TextHeight(m_sCaption) / 2)
             If (iState = giSUNKEN And Not m_eStyle = Flat) Or ((m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected) Then
                 TextRect.Left = TextRect.Left + glSUNKEN_OFFSET
                 TextRect.Right = TextRect.Right + glSUNKEN_OFFSET
                 TextRect.Top = TextRect.Top + glSUNKEN_OFFSET
                 TextRect.Bottom = TextRect.Bottom + glSUNKEN_OFFSET
             End If
            lTextAlign = DT_LEFT
             
         Case vbAlignLeft
             TextRect.Left = 4
             TextRect.Right = UserControl.TextWidth(m_sCaption) + 4
             TextRect.Top = (TextRect.Bottom / 2) - (UserControl.TextHeight(m_sCaption) / 2)
             udtPictureRect.Left = UserControl.ScaleWidth - (UserControl.ScaleX(m_picPicture.Width, vbTwips, vbPixels)) + 25 'TextRect.Right + 4
             udtPictureRect.Right = UserControl.ScaleWidth
             If (iState = giSUNKEN And Not m_eStyle = Flat) Or ((m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected) Then
                 TextRect.Left = TextRect.Left + glSUNKEN_OFFSET
                 TextRect.Right = TextRect.Right + glSUNKEN_OFFSET
                 TextRect.Top = TextRect.Top + glSUNKEN_OFFSET
                 TextRect.Bottom = TextRect.Bottom + glSUNKEN_OFFSET
             End If
            lTextAlign = DT_LEFT
         
         Case Else           'Center the text
             TextRect.Left = 4
             TextRect.Right = UserControl.ScaleWidth - 6
             TextRect.Top = (TextRect.Bottom / 2) - (UserControl.TextHeight(m_sCaption) / 2)
             If (iState = giSUNKEN And Not m_eStyle = Flat) Or ((m_eStyle = Options_1 Or m_eStyle = Options_2) And m_bSelected) Then
                 TextRect.Left = TextRect.Left + glSUNKEN_OFFSET
                 TextRect.Right = TextRect.Right + glSUNKEN_OFFSET
                 TextRect.Top = TextRect.Top + glSUNKEN_OFFSET
                 TextRect.Bottom = TextRect.Bottom + glSUNKEN_OFFSET
             End If
            lTextAlign = DT_CENTER
             
     End Select

     If Not m_picPicture Is Nothing Then
         If m_picPicture.Type = vbPicTypeBitmap Then
             If m_bUseMaskColor Then bUseMask = True
         End If
         
         If Not m_bEnabled Then
             'If button is disabled draw disabled picture on memory dc
             moDrawTool.DrawDisabledPicture lhdcMem, m_picPicture, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, lBackColor, bUseMask, m_clrMaskColor, lhPal
         
         ElseIf bUseMask Then
             'if using mask color draw transparent bitmap on memory dc
                 If (iState = giSUNKEN And Not m_picPictureHover Is Nothing And Not mbLeftMouseDown) Or (m_bSelected And Not m_picPictureSelected Is Nothing) Then
                 moDrawTool.DrawTransparentBitmap lhdcMem, m_picPictureSelected, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, m_clrMaskColor, lhPal
             ElseIf (iState = giRAISED And Not m_picPictureHover Is Nothing) Or (mbLeftMouseDown And Not m_bSelected And Not m_picPictureHover Is Nothing) And Not Ambient.UserMode = False Then
                 moDrawTool.DrawTransparentBitmap lhdcMem, m_picPictureHover, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, m_clrMaskColor, lhPal
             Else
                 moDrawTool.DrawTransparentBitmap lhdcMem, m_picPicture, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, m_clrMaskColor, lhPal
             End If
         
         Else
             'otherwise draw picture with no effects on button
             If m_picPicture.Type = vbPicTypeBitmap Then
                 If (iState = giSUNKEN And Not m_picPictureHover Is Nothing And Not mbLeftMouseDown) Or (m_bSelected And Not m_picPictureSelected Is Nothing) Then
                     moDrawTool.DrawBitmapToHDC lhdcMem, m_picPictureSelected, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, lhPal
                 ElseIf (iState = giRAISED And Not m_picPictureHover Is Nothing) Or (mbLeftMouseDown And Not m_bSelected And Not m_picPictureHover Is Nothing) And Not Ambient.UserMode = False Then
                     moDrawTool.DrawBitmapToHDC lhdcMem, m_picPictureHover, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, lhPal
                 Else
                     moDrawTool.DrawBitmapToHDC lhdcMem, m_picPicture, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, lhPal
                 End If
                 
             ElseIf m_picPicture.Type = vbPicTypeIcon Then
                 If (iState = giSUNKEN And Not m_picPictureHover Is Nothing And Not mbLeftMouseDown) Or (m_bSelected And Not m_picPictureSelected Is Nothing) Then
                     DrawIcon lhdcMem, udtPictureRect.Left, udtPictureRect.Top, m_picPictureSelected.Handle
                 ElseIf (iState = giRAISED And Not m_picPictureHover Is Nothing) Or (mbLeftMouseDown And Not m_bSelected And Not m_picPictureHover Is Nothing) And Not Ambient.UserMode = False Then
                     DrawIcon lhdcMem, udtPictureRect.Left, udtPictureRect.Top, m_picPictureHover.Handle
                 Else
                     DrawIcon lhdcMem, udtPictureRect.Left, udtPictureRect.Top, m_picPicture.Handle
                 End If
                 
             End If
         End If
     End If
     
     Dim x1, y1, x2, y2 As Long
     x1 = mudtButtonRect.Left
     x2 = mudtButtonRect.Right
     y1 = mudtButtonRect.Top
     y2 = mudtButtonRect.Bottom
             

'Draw Frame Needed
DrawButtonState_DrawFrame:
     
     Select Case iState
         Case giFLATTENED, giDISABLED
             If m_eStyle = ThreeD Or (m_eStyle = Options_2 And Not m_bSelected) Then
                 DrawEdge lhdcMem, mudtButtonRect, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
                 
             ElseIf m_eStyle = Outline And m_bSelected Then
                 Call DrawLine(lhdcMem, x1, y1, x1, y2, , m_cLineColorSelected)
                 Call DrawLine(lhdcMem, x2, y1, x2, y2, 2, m_cLineColorSelected)
                 Call DrawLine(lhdcMem, x1, y1, x2, y1, , m_cLineColorSelected)
                 Call DrawLine(lhdcMem, x1, y2, x2, y2, 2, m_cLineColorSelected)
             
             ElseIf m_eStyle = Outline And m_bShowBorder Then
                 Call DrawLine(lhdcMem, x1, y1, x1, y2, , m_cLineColor)
                 Call DrawLine(lhdcMem, x2, y1, x2, y2, 2, m_cLineColor)
                 Call DrawLine(lhdcMem, x1, y1, x2, y1, , m_cLineColor)
                 Call DrawLine(lhdcMem, x1, y2, x2, y2, 2, m_cLineColor)
             
             Else
                 If Not UserControl.Ambient.UserMode Then
                     DrawEdge lhdcMem, mudtButtonRect, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
                 End If
                 
             End If
             
         Case giRAISED
             If m_eStyle = Cool Or m_eStyle = ThreeD Or m_eStyle = Options_1 Or m_eStyle = Options_2 Then
                 DrawEdge lhdcMem, mudtButtonRect, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
                 
             ElseIf m_eStyle = Outline Then
                 Dim cLineColor As OLE_COLOR
                 x1 = mudtButtonRect.Left
                 x2 = mudtButtonRect.Right
                 y1 = mudtButtonRect.Top
                 y2 = mudtButtonRect.Bottom
                 If m_bSelected Then
                     cLineColor = m_cLineColorSelected
                 Else
                     cLineColor = m_cLineColorHover
                 End If
                 Call DrawLine(lhdcMem, x1, y1, x1, y2, , cLineColor)
                 Call DrawLine(lhdcMem, x2, y1, x2, y2, 2, cLineColor)
                 Call DrawLine(lhdcMem, x1, y1, x2, y1, , cLineColor)
                 Call DrawLine(lhdcMem, x1, y2, x2, y2, 2, cLineColor)
                 
             End If
             
         Case giSUNKEN
             If m_eStyle <> Flat Then
                 DrawEdge lhdcMem, mudtButtonRect, BDR_SUNKENOUTER, BF_RECT Or BF_SOFT
             End If
             
     End Select
     
     'Draw the button image
     BitBlt UserControl.hdc, 0, 0, mudtButtonRect.Right, mudtButtonRect.Bottom, lhdcMem, 0, 0, vbSrcCopy
     'Draw the text
     HeightUsed = DrawText(UserControl.hdc, m_sCaption, -1, TextRect, lTextAlign)
     'Check for focus
     If mbGotFocus Then
        Dim rectFocus As RECT
        rectFocus.Bottom = mudtButtonRect.Bottom - 5
        rectFocus.Left = mudtButtonRect.Left + 5
        rectFocus.Right = mudtButtonRect.Right - 5
        rectFocus.Top = mudtButtonRect.Top + 5
        DrawFocusRect UserControl.hdc, rectFocus
     End If
    
DrawButtonStateCleanUp:
     DeleteObject lhbrBack
     SelectPalette lhdcMem, lhPalOld, True
     RealizePalette (lhdcMem)
     DeleteObject SelectObject(lhdcMem, lhbmMemoryOld)
     DeleteDC lhdcMem
    
Exit Sub
    
DrawButtonState_Error:
    Select Case Err.Number
        Case giOBJECT_VARIABLE_NOT_SET
            Resume DrawButtonState_DrawFrame
        Case giINVALID_PICTURE
            Resume DrawButtonState_DrawFrame
        Case Else
            Resume DrawButtonStateCleanUp
    End Select
End Sub

Private Sub DrawLine(hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional LineWidth = 1, Optional ByVal LineColor As OLE_COLOR = vbBlack)
    Dim lpPoint As POINTAPI
    Dim lRetVal As Long
    Dim Brush As LOGBRUSH
    Dim hBrush As Long
    Dim hPen As Long
    Dim hPenOld As Long
    Dim lpRect As RECT
    Dim rgbLineColor As Long
    
    If hdc <> 0 Then
        'Change the color to RGB
        Call OleTranslateColor(LineColor, 0, rgbLineColor)
        'Create the pen for the lines
        hPen = CreatePen(vbSolid, LineWidth, rgbLineColor)
        'Associate the pen with the object
        hPenOld = SelectObject(hdc, hPen)
        'Draw the line
        Call MoveToEx(hdc, x1, y1, lpPoint)
        Call LineTo(hdc, x2, y2)
    
        hPen = SelectObject(hdc, hPenOld)
        Call DeleteObject(hPen)
        
    End If
        
    
End Sub

'*************************
'Friend Methods
'*************************
Friend Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '-------------------------------------------------------------------------
    'Purpose:   Handles window messages specific to subclassed window associated
    '           with this object.  Is called by SubWndProc in standard module.
    '           Relays all mouse messages to ToolTip window, and returns a value
    '           for ToolTip NeedText message.
    '-------------------------------------------------------------------------
    Dim msgStruct As MSG
    Dim hdr As NMHDR
    Dim ttt As ToolTipText
    On Error GoTo WindowProc_Error
    Select Case uMsg
        Case WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP
            With msgStruct
                .lParam = lParam
                .wParam = wParam
                .message = uMsg
                .hwnd = hwnd
            End With
            SendMessage gHWndToolTip, TTM_RELAYEVENT, 0, msgStruct
        Case WM_NOTIFY
            CopyMemory hdr, ByVal lParam, Len(hdr)
            If hdr.code = TTN_NEEDTEXT And hdr.hwndFrom = gHWndToolTip Then
                'Get the tooltip text from the UserControl class object If the host for this control provides a ToolTipText property
                'on the extender object (as in VB5).  The ToolTipText property declares will not be hit.  Therefore, the user's ToolTipText
                'may be found either in the Extender.ToolTipText property or in my own member variable m_sToolTipText
                'Error may occur if ToolTipText property is not available
                On Error Resume Next
                If mbToolTipNotInExtender Then
                    msToolTipBuffer = StrConv(m_sToolTipText, vbFromUnicode)
                Else
                    msToolTipBuffer = StrConv(UserControl.Extender.ToolTipText, vbFromUnicode)
                End If
                If Err.Number = 0 Then
                    CopyMemory ttt, ByVal lParam, Len(ttt)
                    ttt.lpszText = StrPtr(msToolTipBuffer)
                    CopyMemory ByVal lParam, ttt, Len(ttt)
                End If
            End If
        Case WM_CANCELMODE
            'A window has been put over this one
            'flatten the button
            Flatten
            mbGotFocus = False
            mbLeftMouseDown = False
            mbLeftWasDown = False
            mbMouseDownFired = False
    End Select
WindowProc_Resume:
    WindowProc = CallWindowProc(mWndProcNext, hwnd, uMsg, wParam, ByVal lParam)
    Exit Function
WindowProc_Error:
    Resume WindowProc_Resume
End Function



