VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------
'This class provides methods needed for drawing the masked bitmaps and
'disabled or embossed bitmaps and icons
'-------------------------------------------------------------------------
Private mbNoDrawStateFunction As Boolean    'If true, it has been determined that the DrawState API is not supported on this machine
Private mlhdcDeskTop As Long                'Desktop HDC to be used for all create compatibles

Public Sub DrawDisabledPicture(lhdcDest As Long, picSource As Picture, lLeft As Long, lTop As Long, lWidth As Long, _
                                                           lHeight As Long, lBackColor As Long, bUseMaskColor As Boolean, lMaskColor As Long, lhPal As Long)
    '-------------------------------------------------------------------------
    'Purpose:   Draws a bitmap or icon to a HDC, applying a disabled or embossed
    '           look to the picture.  If the DrawState API is available it will be used else, the effect will be mimicked.  If the passed picture
    '           is a bitmap a mask color can be provided to make the areas of the picture that have that color transparent
    'In:
    '   [lHDCDest] HDC of the device context to paint the picture on
    '   [picSource] Picture to paint
    '   [lLeft] X coordinate of the upper left corner of the area that the picture is to be painted on. (in pixels)
    '   [lTop] Y coordinate of the upper left corner of the area that the picture is to be painted on. (in pixels)
    '   [lWidth] Width of picture area to paint in pixels
    '   [lHeight] Height of picture area to paint in pixels
    '   [lBackColor] Is the backcolor of the HDC that the image will be painted on
    '   [bUseMaskColor] If true, use the color mask
    '   [lMaskColor] Color to mask, must be a valid HCOLORREF
    '   [lhPal] Must be a valid HPALETTE
    '-------------------------------------------------------------------------
    Dim lbmTemp As Long
    Dim lbmTempOld As Long
    Dim lhdcTemp As Long
    Dim lhPalOld As Long
    Dim udtTempRect As RECT
    Dim lhbrWhite As Long
    
DrawDisabledPictureTop:     'Return here if DrawState API fails
    On Error Resume Next
    If picSource.Type = vbPicTypeIcon Then        'Icon was passed in
        If mbNoDrawStateFunction Then
            'Don't use drawstate
            DrawDisabledPictureWithoutDrawState lhdcDest, lLeft, lTop, lWidth, lHeight, lBackColor, True, picSource
        Else
            'Use Drawstate
            Err.Clear
            DrawState lhdcDest, 0, 0, picSource.Handle, 0, lLeft, lTop, 0, 0, (DST_ICON Or DSS_DISABLED)
            If Err.LastDllError = ERROR_CALL_NOT_IMPLEMENTED Or Err.Number = giDLL_FUNCTION_NOT_FOUND Then
                mbNoDrawStateFunction = True
                GoTo DrawDisabledPictureTop
            End If
        End If
    Else        'A bitmap was passed in
        If bUseMaskColor Then            'Create a new hDC for creating Bitmap with mask color turned white
            lhdcTemp = CreateCompatibleDC(mlhdcDeskTop)     'Create a temporary HDC compatible to the Destination HDC
            lbmTemp = CreateCompatibleBitmap(mlhdcDeskTop, lWidth, lHeight)
            lbmTempOld = SelectObject(lhdcTemp, lbmTemp)
            lhPalOld = SelectPalette(lhdcTemp, lhPal, True)
            RealizePalette lhdcTemp
            
            With udtTempRect
                .Top = 0
                .Left = 0
                .Bottom = lHeight
                .Right = lWidth
            End With
            
            'Fill the hDC with white
            SetBkColor lhdcTemp, vbWhite
            lhbrWhite = CreateSolidBrush(vbWhite)
            FillRect lhdcTemp, udtTempRect, lhbrWhite
            
            'Draw transparent bitmap on white background
            DrawTransparentBitmap lhdcTemp, picSource, 0, 0, lWidth, lHeight, lMaskColor, lhPal
            SelectObject lhdcTemp, lbmTempOld
            
            If mbNoDrawStateFunction Then                'Don't use DrawState
                DrawDisabledPictureWithoutDrawState lhdcDest, lLeft, lTop, lWidth, lHeight, lBackColor, True, , lbmTemp
            Else                'Use DrawState
                Dim lresult As Long
                Err.Clear
                lresult = DrawState(lhdcDest, 0, 0, lbmTemp, 0, lLeft, lTop, 0, 0, (DST_BITMAP Or DSS_DISABLED))
                If Err.LastDllError = ERROR_CALL_NOT_IMPLEMENTED Or Err.Number = giDLL_FUNCTION_NOT_FOUND Then
                    mbNoDrawStateFunction = True
                    GoTo DrawDisabledPictureTop
                End If
            End If
            
            'Cleanup
            DeleteObject lhbrWhite
            SelectPalette lhPalOld, lhPalOld, True
            RealizePalette lhdcTemp
            DeleteObject SelectObject(lhdcTemp, lbmTempOld)
            DeleteDC lhdcTemp
            
        Else            'Use Bitmap without masking
            If mbNoDrawStateFunction Then                'Don't use drawstate
                DrawDisabledPictureWithoutDrawState lhdcDest, lLeft, lTop, lWidth, lHeight, lBackColor, False, picSource
            Else                'Use DrawState
                Err.Clear
                DrawState lhdcDest, 0, 0, picSource.Handle, 0, lLeft, lTop, 0, 0, (DST_BITMAP Or DSS_DISABLED)
                If Err.LastDllError = ERROR_CALL_NOT_IMPLEMENTED Or Err.Number = giDLL_FUNCTION_NOT_FOUND Then
                    mbNoDrawStateFunction = True
                    GoTo DrawDisabledPictureTop
                End If
            End If
        End If
    End If
    
End Sub

Private Sub DrawDisabledPictureWithoutDrawState(lhdcDest As Long, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, lBackColor As Long, _
                                    bUseMaskColor As Boolean, Optional picSource As Picture, Optional lPicHandle As Long = 0)
    '-------------------------------------------------------------------------
    'Purpose:   Draws a bitmap or icon to a HDC, applying a disabled or embossed look to the picture.  This mimics the effect of the DrawState API
    '           without using it (It is not available in NT 3.51).  If the passed picture is a bitmap a mask color can be provided to make the areas of the
    '           picture that have that color transparent
    'In:
    '   [lhdcDest] HDC of the device context to paint the picture on
    '   [lLeft] X coordinate of the upper left corner of the area that the picture is to be painted on. (in pixels)
    '   [lTop] Y coordinate of the upper left corner of the area that the picture is to be painted on. (in pixels)
    '   [lWidth] Width of picture area to paint in pixels
    '   [lHeight] Height of picture area to paint in pixels
    '   [lBackColor] Is the backcolor of the HDC that the image will be painted on
    '   [bUseMaskColor] If true, the background of the picture is white
    '   [picSource] OLE Picture to be painted.  If not provided the next parameter must contain a valid HBITMAP
    '   [lPicHandle] A valid HBITMAP
    '-------------------------------------------------------------------------
    Dim lhbmFirst As Long
    Dim lhbmOld As Long
    Dim lhbmShadow As Long
    Dim lhbmHighlight As Long
    Dim lhbmDisable As Long
    Dim lhbrPat As Long
    Dim lhbrWhite As Long
    Dim lhdcColor As Long
    Dim lhdcMono As Long
    Dim lBtnShadowColor As Long
    Dim lBtnHighlightColor As Long
    Dim udtTempRect As RECT
    Dim bIconIn As Boolean
     
    lBtnShadowColor = GetSysColor(COLOR_BTNSHADOW)
    lBtnHighlightColor = GetSysColor(COLOR_BTNHIGHLIGHT)
     
    ' We essentially need to create 2 monochrome bitmaps:  one defining the "highlight" area on the bitmap, and one defining the "shadow"
    ' area, which is how we make the resulting bitmap look 3-D.  Both the highlight and shadow bitmaps are created by offsetting
    ' calls to BitBlt() using various ROP codes.  lhbmDisable is our result bitmap.
    lhbmDisable = 0&
    lhdcMono = CreateCompatibleDC(mlhdcDeskTop)
    lhdcColor = CreateCompatibleDC(mlhdcDeskTop)
    
    'create the monochrome and color bitmaps and necessary DCs
    lhbmShadow = CreateBitmap(lWidth, lHeight, 1, 1, ByVal 0&)
    lhbmHighlight = CreateBitmap(lWidth, lHeight, 1, 1, ByVal 0&)
    lhbmDisable = CreateCompatibleBitmap(mlhdcDeskTop, lWidth, lHeight)
    If lPicHandle <> 0 Then
        lhbmOld = SelectObject(lhdcColor, lPicHandle)
    Else
        If picSource.Type = vbPicTypeBitmap Then
            lhbmOld = SelectObject(lhdcColor, picSource.Handle)
        ElseIf picSource.Type = vbPicTypeIcon Then
            bIconIn = True
            lhbmFirst = CreateCompatibleBitmap(lhdcDest, lWidth, lHeight)
            lhbmOld = SelectObject(lhdcColor, lhbmFirst)
            udtTempRect.Bottom = lHeight
            udtTempRect.Right = lWidth
            SetBkColor lhdcColor, vbWhite
            lhbrWhite = CreateSolidBrush(vbWhite)
            FillRect lhdcColor, udtTempRect, lhbrWhite
            DeleteObject lhbrWhite
            DrawIcon lhdcColor, 0, 0, picSource.Handle
        End If
    End If
   
    'set background color of bitmap for mono conversion
    If bUseMaskColor Or bIconIn Then
        SetBkColor lhdcColor, vbWhite
    Else
        SetBkColor lhdcColor, GetPixel(lhdcColor, 0, 0)
    End If
    
    'Create the shadow bitmap
    lhbmShadow = SelectObject(lhdcMono, lhbmShadow)
    PatBlt lhdcMono, 0, 0, lWidth, lHeight, vbWhiteness
    BitBlt lhdcMono, 0, 0, lWidth - 1, lHeight - 1, lhdcColor, 1, 1, vbSrcCopy
    BitBlt lhdcMono, 0, 0, lWidth, lHeight, lhdcColor, 0, 0, vbMergePaint
    lhbmShadow = SelectObject(lhdcMono, lhbmShadow)
    
    'Create the highlight bitmap
    lhbmHighlight = SelectObject(lhdcMono, lhbmHighlight)
    BitBlt lhdcMono, 0, 0, lWidth, lHeight, lhdcColor, 0, 0, vbSrcCopy
    BitBlt lhdcMono, 0, 0, lWidth - 1, lHeight - 1, lhdcColor, 1, 1, vbMergePaint
    lhbmHighlight = SelectObject(lhdcMono, lhbmHighlight)
    
    'Select old bitmap into color DC
    SelectObject lhdcColor, lhbmOld
    
    'Clear the background for the disabled bitmap
    
    SelectObject lhdcColor, lhbmDisable
    
    lhbrPat = CreateSolidBrush(lBackColor)
    lhbrPat = SelectObject(lhdcColor, lhbrPat)
    PatBlt lhdcColor, 0, 0, lWidth, lHeight, vbPatCopy
    DeleteObject SelectObject(lhdcColor, lhbrPat)
    SetBkColor lhdcColor, vbWhite
    SetTextColor lhdcColor, vbBlack
    
    
    ' blt the highlight edge
    
    lhbrPat = CreateSolidBrush(lBtnHighlightColor)
    lhbrPat = SelectObject(lhdcColor, lhbrPat)
    lhbmHighlight = SelectObject(lhdcMono, lhbmHighlight)
    BitBlt lhdcColor, 0, 0, lWidth, lHeight, lhdcMono, 0, 0, PSDPxax
    DeleteObject SelectObject(lhdcColor, lhbrPat)
    lhbmHighlight = SelectObject(lhdcMono, lhbmHighlight)
     
    'Bitblt the shadow edge
    lhbrPat = CreateSolidBrush(lBtnShadowColor)
    lhbrPat = SelectObject(lhdcColor, lhbrPat)
    lhbmShadow = SelectObject(lhdcMono, lhbmShadow)
    BitBlt lhdcColor, 0, 0, lWidth, lHeight, lhdcMono, 0, 0, PSDPxax
    BitBlt lhdcDest, lLeft, lTop, lWidth - 1, lHeight - 1, lhdcColor, 0, 0, vbSrcCopy
    DeleteObject SelectObject(lhdcColor, lhbrPat)
    lhbmShadow = SelectObject(lhdcMono, lhbmShadow)
    
    'Clean up
    DeleteObject lhbmDisable
    DeleteObject lhbmShadow
    DeleteObject lhbmHighlight
    DeleteObject lhbmFirst
    DeleteObject lhbmOld
    
    DeleteDC lhdcMono
    DeleteDC lhdcColor
End Sub

Public Sub DrawTransparentBitmap(lhdcDest As Long, picSource As Picture, lLeft As Long, lTop As Long, lWidth As Long, _
                                                                 lHeight As Long, lMaskColor As Long, lhPal As Long)
    '-------------------------------------------------------------------------
    'Purpose:   Draws a transparent bitmap to a memory HDC.  If you pass a screen HDC into this function there will be flicker.  The pixels of the passed
    '           bitmap that match the passed mask color will not be painted to the destination HDC
    'In:
    '   [lhdcDest] HDC of the memory device context to paint the picture on
    '   [picSource] Picture to paint
    '   [lLeft] X coordinate of the upper left corner of the area that the picture is to be painted on. (in pixels)
    '   [lTop] Y coordinate of the upper left corner of the area that the picture is to be painted on. (in pixels)
    '   [lWidth] Width of picture area to paint in pixels
    '   [lHeight] Height of picture area to paint in pixels
    '   [lMaskColor] Color of pixels to be masked out
    '   [lhPal] Must be a valid HPALETTE
    '-------------------------------------------------------------------------
    Dim lhdcMask As Long        'HDC of the created mask image
    Dim lhdcColor As Long       'HDC of the created color image
    Dim lhdcSrc As Long         'HDC that the source bitmap is selected into
    Dim lhbmMask As Long        'Bitmap handle to the mask image
    Dim lhbmcolor As Long       'Bitmap handle to the color image
    Dim lhbmMemSrcOld As Long
    Dim lhbmColorOld As Long
    Dim lhbmMaskOld As Long
    Dim lhPalOld As Long
    
    'Verify that the passed picture is a Bitmap
    If Not picSource.Type = vbPicTypeBitmap Then Error.Raise giINVALID_PICTURE
    'Select passed picture into an HDC
    lhdcSrc = CreateCompatibleDC(mlhdcDeskTop)
    lhbmMemSrcOld = SelectObject(lhdcSrc, picSource.Handle)
    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with lhdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    lhbmcolor = CreateCompatibleBitmap(mlhdcDeskTop, lWidth, lHeight)
    'Now create a monochrome bitmap for the mask
    lhbmMask = CreateBitmap(lWidth, lHeight, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first and then use it instead of the source bitmap
    'because the source bitmap may be a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    lhdcColor = CreateCompatibleDC(mlhdcDeskTop)
    lhbmColorOld = SelectObject(lhdcColor, lhbmcolor)
    lhPalOld = SelectPalette(lhdcColor, lhPal, True)
    RealizePalette lhdcColor
    'In case lhdcSrc contains a monochrome bitmap, we must set the destination foreground/background colors according to those currently set in lhdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor lhdcColor, GetBkColor(lhdcSrc)
    SetTextColor lhdcColor, GetTextColor(lhdcSrc)

    BitBlt lhdcColor, 0, 0, lWidth, lHeight, lhdcSrc, 0, 0, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color from the source, and black everywhere else.
    lhdcMask = CreateCompatibleDC(mlhdcDeskTop)
    lhbmMaskOld = SelectObject(lhdcMask, lhbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1 all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor lhdcColor, lMaskColor
    SetTextColor lhdcColor, vbWhite
    BitBlt lhdcMask, 0, 0, lWidth, lHeight, lhdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white bits (1) to the background color of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor lhdcColor, vbBlack
    SetBkColor lhdcColor, vbWhite
    BitBlt lhdcColor, 0, 0, lWidth, lHeight, lhdcMask, 0, 0, DSna
    'Paint the Mask to the destination
    BitBlt lhdcDest, lLeft, lTop, lWidth, lHeight, lhdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the destination
    BitBlt lhdcDest, lLeft, lTop, lWidth, lHeight, lhdcColor, 0, 0, vbSrcPaint
    'All done!
    DeleteObject SelectObject(lhdcColor, lhbmColorOld)
    SelectObject lhdcSrc, lhbmMemSrcOld
    SelectPalette lhdcColor, lhPalOld, True
    RealizePalette (lhdcColor)
    DeleteDC (lhdcSrc)
    DeleteDC (lhdcColor)
    DeleteObject SelectObject(lhdcMask, lhbmMaskOld)
    DeleteDC (lhdcMask)
    Exit Sub
End Sub

Public Sub DrawBitmapToHDC(lhdcDest As Long, picSource As Picture, lLeft As Long, lTop As Long, _
                                                         lWidth As Long, lHeight As Long, lhPal As Long)
    '-------------------------------------------------------------------------
    'Purpose:   Draws a Bitmap to an HDC without transparency
    'In:
    '   [lhdcDest] HDC of the memory device context to paint the picture on
    '   [picSource] Picture to paint
    '   [lLeft] X coordinate of the upper left corner of the area that the picture is to be painted on. (in pixels)
    '   [lTop] Y coordinate of the upper left corner of the area that the picture is to be painted on. (in pixels)
    '   [lWidth] Width of picture area to paint in pixels
    '   [lHeight] Height of picture area to paint in pixels
    '   [lhPal] Must be a valid HPALETTE
    '-------------------------------------------------------------------------
    Dim lhdcTemp As Long
    Dim lhPalOld As Long
    Dim hbmOld As Long
    
    'Validate that a bitmap was passed in
    If picSource.Type <> vbPicTypeBitmap Then Error.Raise giINVALID_PICTURE
    'Create a DC to select bitmap into
    lhdcTemp = CreateCompatibleDC(mlhdcDeskTop)
    lhPalOld = SelectPalette(lhdcTemp, lhPal, True)
    RealizePalette lhdcTemp
    'Select bitmap into DC
    hbmOld = SelectObject(lhdcTemp, picSource.Handle)
    'Copy to destination DC
    BitBlt lhdcDest, lLeft, lTop, lWidth, lHeight, lhdcTemp, 0, 0, vbSrcCopy
    'Cleanup
    SelectObject lhdcTemp, hbmOld
    SelectPalette lhdcTemp, lhPalOld, True
    RealizePalette (lhdcTemp)
    DeleteDC lhdcTemp

End Sub

Private Sub Class_Initialize()
    'Get Desktop DC and create halftone palette to be used
    mlhdcDeskTop = GetDC(ByVal 0&)
    #If TEST_NT351_MODE Then
        mbNoDrawStateFunction = True
    #End If
End Sub


Private Sub Class_Terminate()
    'Cleanup
    ReleaseDC ByVal 0&, mlhdcDeskTop
End Sub
