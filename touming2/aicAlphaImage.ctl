VERSION 5.00
Begin VB.UserControl aicAlphaImage 
   BackStyle       =   0  'Í¸Ã÷
   CanGetFocus     =   0   'False
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   ClipBehavior    =   0  'ÎÞ
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'ÎÞ
   MaskColor       =   &H80000014&
   PropertyPages   =   "aicAlphaImage.ctx":0000
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   72
   ToolboxBitmap   =   "aicAlphaImage.ctx":0015
   Windowless      =   -1  'True
End
Attribute VB_Name = "aicAlphaImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


Public Event Click(ByVal Button As Integer)
Attribute Click.VB_Description = "Occurs when the mouse is pressed and released over the control"
Attribute Click.VB_MemberFlags = "200"
Public Event DblClick(ByVal Button As Integer)
Attribute DblClick.VB_Description = "Occurs when the mouse is double clicked over the control"
Public Event MouseExit()
Attribute MouseExit.VB_Description = "Occurs when the user first moves the mouse cursor out of the control"
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the user first moves the mouse cursor into the control"
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus"
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus"
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus"
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key"
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus"
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual"
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual"

' Custom callbacks
Public Event FadeTerminated(ByVal CurrentOpacity As Long)
Attribute FadeTerminated.VB_Description = "Callback event that occurs when a fade request completes"

Private z_CbMem   As Long    'Callback allocated memory address
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
'-------------------------------------------------------------------------------------------------

' Timer and HitTest related APIs
Private Declare Function SetTimer Lib "USER32.DLL" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "USER32.DLL" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function WindowFromPoint Lib "USER32.DLL" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "USER32.DLL" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function PtInRegion Lib "GDI32.DLL" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "USER32.DLL" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "USER32.DLL" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "USER32.DLL" (ByVal hWnd As Long) As Long

' Drawing related APIs
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32.DLL" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "GDI32.DLL" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "GDI32.DLL" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32.DLL" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetClipBox Lib "GDI32.DLL" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetRgnBox Lib "GDI32.DLL" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRect Lib "USER32.DLL" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function IntersectRect Lib "USER32.DLL" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long
Private Declare Function OffsetRect Lib "USER32.DLL" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRgn Lib "GDI32.DLL" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type MASKUSAGE
    Color As Long           ' current mask color
    Applied As Boolean      ' mask has been applied
    AppliedColor As Long    ' color used to create mask; may not be same as Color
    Source As aiMaskSource  ' mask option: see aiMaskSource enum
End Type
Private Type FADERCONTROL
    tmrAddr As Long         ' AddressOf timer call back procedure
    fStep As Long           ' percent of opacity to change between steps
    fDelay As Long          ' length to delay before next step occurs
    fOpacity As Long        ' final opacity that also terminates the fader
End Type
Private Type SCALEDCOORD
    Left As Long            ' position of image within usercontrol
    Top As Long
    Width As Long           ' size of image within usercontrol
    Height As Long
    RotatedSize As Long     ' when rotated, the size needed to completely rotate image 360 degrees
    xOffset As Long         ' runtime offsets to render image....
    yOffset As Long         '   could be used to shift image when user clicks on it to create a down effect
    OneToOne As Boolean     ' flag used for painting; when image is actual size, faster renders
End Type

Public Enum aiMaskSource
    aiNoMask = 0
    aiUseMaskColor = 1
    aiUseTopLeft = 2
    aiUseTopRight = 3
    aiUseBottomLeft = 4
    aiUseBottomRight = 5
End Enum
Public Enum aiMirrorEnum
    aiMirrorNone = 0
    aiMirrorHorizontal = 1
    aiMirrorVertical = 2
    aiMirrorAll = 3
End Enum
Public Enum aiScaleMethod
    aiScaled = 0
    aiStretch = 1
    aiScaleDownOnly = 2
    aiActualSize = 3
    aiLockScale = 4
End Enum
Public Enum aiGrayScales
    aiNTSCPAL = 1     ' R=R*.299, G=G*.587, B=B*.114 - Default
    aiCCIR709 = 2     ' R=R*.213, G=G*.715, B=B*.072
    aiSimpleAvg = 3   ' R,G, and B = (R+G+B)/3
    aiRedMask = 4     ' uses only the Red sample value: RGB = Red / 3
    aiGreenMask = 5   ' uses only the Green sample value: RGB = Green / 3
    aiBlueMask = 6    ' uses only the Blue sample value: RGB = Blue / 3
    aiRedGreenMask = 7 ' uses Red & Green sample value: RGB = (Red+Green) / 2
    aiBlueGreenMask = 8 ' uses Blue & Green sample value: RGB = (Blue+Green) / 2
    aiNoGrayScale = 0
End Enum
Public Enum aiHitTestStyle  ' see HitTest property
    aiBoundingRgn = 1
    aiEnclosedRgn = 2
    aiShapedRgn = 3
    aiEntireControl = 0
End Enum
Public Enum aiOLEDropMode
    aiDropNone = vbOLEDropNone
    aiDropManual = vbOLEDropManual
End Enum
Public Enum eImageFormatUC    ' source image format
    img_Error = -1  ' no DIB has been initialized
    img_None = 0    ' no image loaded
    img_Bitmap = 1  ' standard bitmap or jpg
    img_Icon = 3    ' standard icon
    img_WMF = 2     ' windows meta file
    img_EMF = 4     ' enhanced WMF
    img_Cursor = 5  ' standard cursor
    img_BmpARGB = 6  ' 32bpp bitmap where RGB is not pre-multiplied
    img_BmpPARGB = 7 ' 32bpp bitmap where RGB is pre-multiplied
    img_IconARGB = 8 ' XP-type icon; 32bpp ARGB
    img_GIF = 9      ' gif; if class.Alpha=True, then transparent GIF
    img_PNG = 10     ' PNG image
    img_PNGicon = 11 ' PNG in icon file (Vista)
    img_CursorARGB = 12 ' alpha blended cursors? do they exist yet?
    img_CheckerBoard = 64 ' image is displaying own checkerboard pattern; no true image
End Enum

Private Enum eProps
    pHiQualty = 1       'HighQuality interpolation
    pAutoStretch = 2    'Auto Stretch
    pAutoSize = 4       'Auto Size
    pAutoRedraw = 8     'Auto Redraw
    pKeepBytes = 16     'Cache original bytes
    pOffscreen = 32     'Offscreen image maintained
    pMasked = 64        'Image is masked for transparency
    pShadowed = 128     'Image has shadow class
    pRotated = 256      'Image can be rotated
End Enum
Private cKeyProps As eProps
Private cBtnClickTracker As Integer    ' tracks button used to trigger click/dblclick events


Private cAngle As Long
Private cHitTest As aiHitTestStyle
Private cRegion As Long     ' used when cHitTest is aiShapedRgn, aiEnclosedRgn
Private cRgnBox As RECT     ' used when cHitTest is aiEntireControl, aiBoundingRgn

Private cGrayScale As aiGrayScales
Private cScaleMode As ScaleModeConstants    ' parent container's scalemode; used for public events
Private cScaleMethod As aiScaleMethod
Private cMirror As aiMirrorEnum
Private cOpacity As Long
Private cIntesity As Single
Private cMask As MASKUSAGE
Private cScaledCoords As SCALEDCOORD

Private cImage As c32bppDIB
Private cOffscreen As c32bppDIB             ' used when AutoRedraw=True

'//Timer & mouse enter/exit related variables
Private cTmrAddrOf As Long
Private cTmrHwnd As Long
Private cTopLeftPos As POINTAPI
Private cFader As FADERCONTROL

'//Shadow options
Private cShadowClass As c32bppDIB           ' used when creating shadows on non-actual size images
Private cShadowColor As Long
Private cShadowOffset As POINTAPI
Private cShadowDepth As Long
Private cShadowOpacity As Long

Private cCrossTalker As cCrossComm
Implements ICrossComm


' This module is grouped/organized into the following sections:
'   Public Properties & Functions
'   Property Page Routines
'   Inter-Control Communication
'   Support/Helper Routines
'   Usercontrol Properties
'   Callback Thunking Routines
'=========================================================================================================
' /////////////////////// P U B L I C   P R O P E R T I E S  &  F U N C T I O N S \\\\\\\\\\\\\\\\\\\\\\\\
'=========================================================================================================

Public Sub ClearImage()
Attribute ClearImage.VB_Description = "Removes image from control"
    ' Purpose: Permanently remove the control's image source
    If ((cKeyProps And pOffscreen) = pOffscreen) Then Set cOffscreen = Nothing
    cImage.DestroyDIB
    Set cShadowClass = Nothing
    If Ambient.UserMode = True Then
        UserControl.Refresh
    Else ' can be called by the property page during design time
        cImage.InitializeDIB cScaledCoords.Width, cScaledCoords.Height
        cImage.CreateCheckerBoard
        sptResize
    End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object"
    On Error Resume Next
    ' attempt to get actual scalemode. Errors occur when container has no scalemode (i.e.,frame)
    cScaleMode = UserControl.Parent.ScaleMode
    If Err Then
        Err.Clear
        cScaleMode = vbContainerPosition
        '^^ the above can be used for any scalemode; however, when container scalemode is
        ' vbTwips or vbPixels, the returned X,Y coordinates passed back are decimal not integer
    End If
    On Error GoTo 0
    ' Purpose: Simply refresh the control
    UserControl.Refresh
   
End Sub

Public Function LoadImage_FromArray(inStream() As Byte, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromArray.VB_Description = "Option to load an image from a stream of data"
    ' Purpose. Load an image from array
    If desiredIconHeight < 1 Then desiredIconHeight = UserControl.ScaleHeight
    If desiredIconWidth < 1 Then desiredIconWidth = UserControl.ScaleWidth
    If desiredIconBitDepth < 1 Then desiredIconBitDepth = 32&
    If cImage.LoadPicture_Stream(inStream, desiredIconWidth, desiredIconHeight, , , ((cKeyProps And pKeepBytes) = pKeepBytes), desiredIconBitDepth) Then
        sptReplaceImage
        LoadImage_FromArray = True
    End If
End Function
Public Function LoadImage_FromFile(FileName As String, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromFile.VB_Description = "Option to load an image from a file"
    ' Purpose. Load an image from a file
    If desiredIconHeight < 1 Then desiredIconHeight = UserControl.ScaleHeight
    If desiredIconWidth < 1 Then desiredIconWidth = UserControl.ScaleWidth
    If desiredIconBitDepth < 1 Then desiredIconBitDepth = 32&
    If cImage.LoadPicture_File(FileName, desiredIconWidth, desiredIconHeight, ((cKeyProps And pKeepBytes) = pKeepBytes), desiredIconBitDepth) Then
        sptReplaceImage
        LoadImage_FromFile = True
    End If
End Function
Public Function LoadImage_FromStdPicture(stdPic As StdPicture) As Boolean
Attribute LoadImage_FromStdPicture.VB_Description = "Option to load an image from a standard picture object"
    ' Purpose. Load an image from stdPicture object
    If cImage.LoadPicture_StdPicture(stdPic) = True Then
        If ((cKeyProps And pKeepBytes) = pKeepBytes) Then
            Dim srcData() As Byte
            cImage.SaveToStream srcData
            cImage.SetOriginalFormat srcData
        End If
        sptReplaceImage
        LoadImage_FromStdPicture = True
    End If
End Function
Public Function LoadImage_FromClipboard() As Boolean
Attribute LoadImage_FromClipboard.VB_Description = "Option to load an image from the clipboard"
    ' Purpose. Load an image from the clipboard
    If cImage.LoadPicture_ClipBoard() Then
        If ((cKeyProps And pKeepBytes) = pKeepBytes) Then
            Dim srcData() As Byte
            cImage.SaveToStream srcData
            cImage.SetOriginalFormat srcData
            Erase srcData()
        End If
        sptReplaceImage
        LoadImage_FromClipboard = True
    End If
End Function
Public Function LoadImage_FromHandle(Handle As Long) As Boolean
Attribute LoadImage_FromHandle.VB_Description = "Option to load an image from an existing memory handle"
    ' Purpose. Load an image from an image handle
    If cImage.LoadPicture_ByHandle(Handle) Then
        If ((cKeyProps And pKeepBytes) = pKeepBytes) Then
            Dim srcData() As Byte
            cImage.SaveToStream srcData
            cImage.SetOriginalFormat srcData
        End If
        sptReplaceImage
        LoadImage_FromHandle = True
    End If
End Function
Public Function LoadImage_FromResource(VBglobal As IUnknown, ResSection As Variant, ResID As Variant, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromResource.VB_Description = "Option to load an image from a resource file"
    ' Purpose. Load an image from a resource file
    
    ' Pass VBglobal like so:  VB.Global
    ' Pass ResSection as: vbResBitmap, vbResIcon, "Custom", etc
    ' Pass ResID as:  101, "MyLogo", etc
    If desiredIconHeight < 1 Then desiredIconHeight = UserControl.ScaleHeight
    If desiredIconWidth < 1 Then desiredIconWidth = UserControl.ScaleWidth
    If desiredIconBitDepth < 1 Then desiredIconBitDepth = 32&
    If cImage.LoadPicture_Resource(ResID, ResSection, VBglobal, desiredIconWidth, desiredIconHeight, , , desiredIconBitDepth) Then
        sptReplaceImage
        LoadImage_FromResource = True
    End If
End Function
Public Function LoadImage_FromOrignalBytes(Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromOrignalBytes.VB_Description = "Option to load image from bytes maintained by setting KeepOriginal format to True"
    ' Purpose. Load an image from a file
    If desiredIconHeight < 1 Then desiredIconHeight = UserControl.ScaleHeight
    If desiredIconWidth < 1 Then desiredIconWidth = UserControl.ScaleWidth
    If desiredIconBitDepth < 1 Then desiredIconBitDepth = 32&
    If cImage.LoadPicture_FromOrignalFormat(desiredIconWidth, desiredIconHeight, desiredIconBitDepth) Then
        sptReplaceImage
        LoadImage_FromOrignalBytes = True
    End If
End Function

Public Function GetImageBytes(imgBytes() As Byte, ByRef scanWidth As Long, _
                                Optional ByVal asArray2D As Boolean = False, _
                                Optional ByVal asBGRformat As Boolean = True, _
                                Optional ByVal asBottomUp As Boolean = False, _
                                Optional ByVal asPreMultiplied As Boolean = False, _
                                Optional ByVal asFileFormat As Boolean = False) As Boolean
Attribute GetImageBytes.VB_Description = "Function returns the pixels of a source image as a byte array. Parameters allow bytes to be returned in file format."
    
    ' Purpose: Return the DIB bytes in a 1 or 2 dimensional array
    ' Return value is false if no image exists
    
    ' [Parameters].
    ' imgBytes(): a dynamic byte array; will be redimensioned as needed
    ' scanWidth: placeholder for the scan width of the bytes returned
    ' asArray2D: if True, arrays is returned as (x,y) else array is 1 dimensional
    '   see asFileFormat for an exception
    ' asBGRformat: if True, bytes are BGRA else bytes are RGBA
    ' asBottomUp: if True, first byte is bottom,left of image else is top,left of image
    ' asPremultiplied: if True, bytes are in premultiplied format else they are not
    ' asFileFormat: if True, the array is zero-bound and one dimensional only
    '   no other parameters are used. The array can be saved to disk as a file
    '   Call ImageType property first to determine what image format the bytes
    '   will be returned as
    
    If Not (cImage.ImageType = imgError Or cImage.ImageType = imgNone) Then
    
        If Not cImage.ImageType = imgCheckerBoard Then
            If asFileFormat = True Then
                If cImage.GetOrginalFormat(imgBytes) = False Then
                    If cImage.isGDIplusEnabled Or cImage.isZlibEnabled Then
                        GetImageBytes = cImage.SaveToStream_PNG(imgBytes)
                    Else
                        GetImageBytes = cImage.SaveToStream(imgBytes)
                    End If
                End If
            Else
                scanWidth = cImage.scanWidth
                GetImageBytes = cImage.GetDIBbits(imgBytes, asArray2D, asBGRformat, True, True, asBottomUp, , , , , asPreMultiplied)
            End If
        End If
        
    End If
End Function
Public Function SetImageBytes(imgBytes() As Byte, _
                                Optional ByVal isBGRformat As Boolean = True, _
                                Optional ByVal isBottomUp As Boolean = False) As Boolean
Attribute SetImageBytes.VB_Description = "Function sets the pixels of a source image from a passed  byte array"
    
    ' Purpose: Sets the DIB bytes with the passed modified bytes and refresh image
    ' Return value is True if the bytes were applied
    
    ' [Parameters].
    ' imgBytes(): a dynamic byte array; must be 1 or 2 dimensional
    ' isBGRformat: if True, bytes are BGRA else bytes are RGBA
    ' isBottomUp: if True, first byte is bottom,left of image else is top,left of image
    
    If Not (cImage.ImageType = imgError) Then
        If cImage.SetDIBbits(imgBytes, isBGRformat, , , isBottomUp) = True Then
            If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
            If Me.ShadowEnabled Then sptRefreshShadow
            UserControl.Refresh
            SetImageBytes = True
        End If
    End If
End Function

Public Sub GetImageScales(ByRef Width As Long, ByRef Height As Long, _
            Optional ByVal ScaleMethod As aiScaleMethod = -1, _
            Optional ByVal desiredWidth As Long, Optional ByVal desiredHeight As Long, _
            Optional ByVal asRotatedImage As Boolean = False)
Attribute GetImageScales.VB_Description = "Function returns scaled image width and height relative to passed target width and height"
            
    ' Purpose: Return the size of the image that will fit within the passed
    '       DC dimensions using the passed scale method
    ' Notes:
    '   1. Parameters are pixels
    '   2. To resize your control to allow full rotation, use Sqr(Width*Width+Height*Height)
    
    ' [Parameters]
    '   Width. Variable to hold the returned scaled width
    '   Height. Variable to hold the returned scaled height
    '   ScaleMethod. If -1 then the current scale method will be used else the passed method
    '   desiredWidth. The desired width. If not passed, the current usercontrol width is used
    '   desiredHeight.  The desired height. If not passed, the current usercontrol height is used
    '   asRotatedImage. Indicates the desired width/height allows full 360 degree rotation
    
    If cImage.Handle = 0& Then Exit Sub
            
    If ScaleMethod < aiScaled Then
        ScaleMethod = cScaleMethod
    ElseIf ScaleMethod > aiLockScale Then
        ScaleMethod = aiLockScale
    End If
    
    If ScaleMethod = aiLockScale Then
        Width = cScaledCoords.Width
        Height = cScaledCoords.Height
        Exit Sub
    End If
    
    If desiredWidth < 1 Then desiredWidth = UserControl.ScaleWidth
    If desiredHeight < 1 Then desiredHeight = UserControl.ScaleHeight
    
    Dim xRatio As Single, yRatio As Single
    Dim rotSize As Long
    
    If asRotatedImage Then
        If ScaleMethod = aiActualSize Then
            Width = cImage.Width
            Height = cImage.Height
        Else
            rotSize = Sqr(cImage.Width * cImage.Width + cImage.Height * cImage.Height)
            xRatio = cImage.Width / rotSize
            yRatio = cImage.Height / rotSize
            Select Case ScaleMethod
            Case aiStretch
                Width = desiredWidth * xRatio
                Height = desiredHeight * yRatio
            Case Else
                If desiredWidth > desiredHeight Then
                    rotSize = desiredHeight
                Else
                    rotSize = desiredWidth
                End If
                Width = rotSize * xRatio
                Height = rotSize * yRatio
                If ScaleMethod = aiScaleDownOnly Then
                    If Width > cImage.Width Or Height > cImage.Height Then
                        Width = cImage.Width
                        Height = cImage.Height
                    End If
                End If
            End Select
        End If
    Else
        Select Case ScaleMethod
        Case aiActualSize
            Width = cImage.Width
            Height = cImage.Height
        Case aiStretch
            Width = desiredWidth
            Height = desiredHeight
        Case Else
            xRatio = desiredWidth / cImage.Width
            yRatio = desiredHeight / cImage.Height
            If yRatio < xRatio Then xRatio = yRatio
            If ScaleMethod = aiScaleDownOnly Then
                If xRatio > 1 Then xRatio = 1
            End If
            Width = cImage.Width * xRatio
            Height = cImage.Height * xRatio
        End Select
    End If
    
End Sub
            
Public Property Get ImageType() As eImageFormatUC
Attribute ImageType.VB_Description = "Returns/sets the image format contained by the control."
    ' Useful if you want to call GetImageBytes to return the image in file format for saving.
    ' Call this property first to determine the format to be returned by GetImageBytes
    If cImage.ImageType = imgError Or cImage.ImageType = imgNone Or cImage.ImageType = imgCheckerBoard Then
        ImageType = imgNone
    ElseIf (cKeyProps And pKeepBytes) = pKeepBytes Then ' original format kept
        ImageType = cImage.ImageType
    ElseIf cImage.isGDIplusEnabled = True Then
        ImageType = imgPNG
    ElseIf cImage.isZlibEnabled = True Then
        ImageType = imgPNG
    ElseIf cImage.Alpha = True Then
        ImageType = imgBmpPARGB
    Else
        ImageType = imgBitmap
    End If
End Property
Public Sub FadeInOut(ByVal FinalOpacity As Long, Optional ByVal FadeStepPercent As Long = 10, Optional ByVal FadeDelayInterval As Long = 30)
Attribute FadeInOut.VB_Description = "Automatically fades an image to a requested opacity"
    
    ' Purpose: Fade an image from one opacity to another at a stepped percentage while delaying n milliseconds between steps
    ' Note: all values should be positive numbers. Routines will double check and adjust as needed
    
    ' [Parameters]
    ' FinalOpacity: The opacity of the image when the fader terminates
    ' FadeStepPercent: What percentage of opaqueness image should change between intervals (valid values are 1 to 100)
    ' FadeDelayInterval: How long image should remain before next fade step occurs (valid values are 10 to max long value)
    
    ' Kill current fader if it exists
    If Not cFader.fStep = 0& Then
        KillTimer cTmrHwnd, -ObjPtr(Me)
        cFader.fStep = 0&
    End If
    
    ' validate final opacity parameters
    If FinalOpacity < 0& Then
        FinalOpacity = 0&
    ElseIf FinalOpacity > 100& Then
        FinalOpacity = 100&
    End If
    ' validate step value
    If FadeStepPercent = 0& Then
        Exit Sub
    Else
        ' set up fader control
        If Not cOpacity = FinalOpacity Then
            If FinalOpacity < cOpacity Then
                FadeStepPercent = -Abs(FadeStepPercent) ' must be negative
            Else
                FadeStepPercent = Abs(FadeStepPercent)  ' else must be positive
            End If
            With cFader
                ' get AddressOf for our fader timer
                If .tmrAddr = 0& Then .tmrAddr = zb_AddressOf(2, 4, 1)
                .fOpacity = FinalOpacity        ' set final opacity value
                .fStep = FadeStepPercent        ' set fader step value
                If FadeDelayInterval < 10 Then  ' set fader interval
                    .fDelay = 10
                Else
                    .fDelay = FadeDelayInterval
                End If
            End With
            ' fire first fader event
            Call Timer_Fader(cTmrHwnd, 0&, -ObjPtr(Me), 0&)
        End If
    End If
    
End Sub

Public Property Let AutoRedraw(Enable As Boolean)
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap"
    ' Purpose: create/maintain a scaled/drawn image in its own DC.
    ' Although it uses more resources, images are rendered much faster since they
    ' to not have to be resized, rotated, mirrored, etc, each time the control must be updated
    ' Additionally, since the cached image is now at a 1:1 ratio, rendering only occurs to
    ' the invalidated portion of the control, not the entire control each time
    If Not (((cKeyProps And pAutoRedraw) = pAutoRedraw) = Enable) Then
        cKeyProps = cKeyProps Xor pAutoRedraw
        If Ambient.UserMode = True Then
            cKeyProps = cKeyProps Xor pOffscreen
            If Enable = False Then Set cOffscreen = Nothing
        End If
        PropertyChanged "AutoRedraw"
    End If
End Property
Public Property Get AutoRedraw() As Boolean
    AutoRedraw = ((cKeyProps And pAutoRedraw) = pAutoRedraw)
End Property

Public Property Let HitTest(Style As aiHitTestStyle)
Attribute HitTest.VB_Description = "Returns/Sets method used to determine whether control responds to mouse events"
    ' Determines whether the mouse in the control registers as within the image
    ' Possible values are:
    ' - aiEntireControl. Mouse anywhere within the control is allowed
    ' - aiShapedRgn. Mouse on any non-transparent pixel is allowed
    ' - aiBoundingRgn. Mouse within the tightest rectangle around image is valid
    ' - aiEnclosedRgn. Same as aiShapedRgn but any transparent pixels that fall
    '       between non-transparent pixels is considered non-transparent
    If Style >= aiEntireControl And Style <= aiShapedRgn Then
        If Not Style = cHitTest Then
            cHitTest = Style
            sptRefreshRegion
            PropertyChanged "HitTest"
        End If
    End If
End Property
Public Property Get HitTest() As aiHitTestStyle
    HitTest = cHitTest
End Property

Public Property Let MaskColor(Color As OLE_COLOR)
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the image"
    ' Allows making a color within a non-transparent bitmap transparent
    If cMask.Applied Then ' previous mask already applied, remove it
        If cMask.Source = aiUseMaskColor Then
            sptUndoMask
            If (cKeyProps And pMasked) = pMasked Then
                cImage.MakeTransparent sptConvertSysColor(Color) Xor vbWhite
            Else
                cImage.MakeTransparent sptConvertSysColor(Color)
            End If
            cMask.AppliedColor = Color
            If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
            If Me.ShadowEnabled Then sptRefreshShadow
            UserControl.Refresh
        End If
    End If
    cMask.Color = Color
    PropertyChanged "MaskColor"
End Property
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = cMask.Color
End Property

Public Property Let MaskUsed(Style As aiMaskSource)
Attribute MaskUsed.VB_Description = "Returns/Sets whether the mask is to be applied to the image"
    ' Enables the mask to take effect.
    ' Only applies to non-transparent images
    If cMask.Applied = False And cImage.Alpha = True Then Exit Property
    If Style >= aiNoMask And Style <= aiUseBottomRight Then
        If Not cMask.Source = Style Then
            If Not cImage.ImageType = imgCheckerBoard Then
                Dim mColor As Long
                If cMask.Applied Then sptUndoMask
                If Style = aiNoMask Then
                    cMask.Applied = False
                Else
                    Select Case Style
                    Case aiUseMaskColor
                        cMask.AppliedColor = cMask.Color
                    Case aiUseBottomLeft
                        cMask.AppliedColor = cImage.GetPixel(0, cImage.Height - 1, , False)
                    Case aiUseBottomRight
                        cMask.AppliedColor = cImage.GetPixel(cImage.Width - 1, cImage.Height - 1, , False)
                    Case aiUseTopLeft
                        cMask.AppliedColor = cImage.GetPixel(0, 0, , False)
                    Case aiUseTopRight
                        cMask.AppliedColor = cImage.GetPixel(cImage.Width - 1, 0, , False)
                    End Select
                    cMask.Applied = True
                    If (cKeyProps And pMasked) = pMasked Then
                        If Style = aiUseMaskColor Then
                            mColor = sptConvertSysColor(cMask.AppliedColor) Xor vbWhite
                        Else
                            mColor = cMask.AppliedColor
                            cMask.AppliedColor = mColor Xor vbWhite
                        End If
                    Else
                        mColor = sptConvertSysColor(cMask.AppliedColor)
                    End If
                    cImage.MakeTransparent mColor
                End If
                If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
                If Me.ShadowEnabled Then sptRefreshShadow
                UserControl.Refresh
                cMask.Source = Style
                PropertyChanged "MaskUsed"
            End If
        End If
    End If
End Property
Public Property Get MaskUsed() As aiMaskSource
    MaskUsed = cMask.Source
End Property

Public Property Let InversedImage(Inverse As Boolean)
Attribute InversedImage.VB_Description = "Returns/Sets whether the image colors are inverted"
    ' Purpose: Inverts the colors of an image. Can be toggled
    If Not (((cKeyProps And pMasked) = pMasked) = Inverse) Then
        cKeyProps = cKeyProps Xor pMasked
        cImage.MakeImageInverse
        PropertyChanged "InversedImage"
        sptUpdateOffscreen False, False
        UserControl.Refresh
    End If
End Property
Public Property Get InversedImage() As Boolean
    InversedImage = ((cKeyProps And pMasked) = pMasked)
End Property

Public Property Let AutoSize(Value As Boolean)
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents"
    ' Purpose: Forces control to resize to the scaled image size
    If Not (((cKeyProps And pAutoSize) = pAutoSize) = Value) Then
        cKeyProps = cKeyProps Xor pAutoSize
        sptResize
        PropertyChanged "AutoSize"
    End If
End Property
Public Property Get AutoSize() As Boolean
    AutoSize = ((cKeyProps And pAutoSize) = pAutoSize)
End Property

Public Property Let ScaleMethod(Method As aiScaleMethod)
Attribute ScaleMethod.VB_Description = "Returns/sets a value that determines how a graphic resizes to fit the size of an Image control"
    ' Purpose: Determine whether image is scaled proportionally or not
    If Not Method = cScaleMethod Then
        cScaleMethod = Method
        Call sptResize
        PropertyChanged "Stretch"
    End If
End Property
Public Property Get ScaleMethod() As aiScaleMethod
    ScaleMethod = cScaleMethod
End Property

Public Property Let StretchQuality(highQuality As Boolean)
Attribute StretchQuality.VB_Description = "Returns/sets whether a graphic will be resized using the best sizing algorithms"
    ' Determines algoritm to use when stretching an image
    If Not (((cKeyProps And pHiQualty) = pHiQualty) = highQuality) Then
        cKeyProps = cKeyProps Xor pHiQualty
        cImage.HighQualityInterpolation = highQuality
        UserControl.Refresh
        PropertyChanged "StretchQuality"
    End If
End Property
Public Property Get StretchQuality() As Boolean
    StretchQuality = ((cKeyProps And pHiQualty) = pHiQualty)
End Property

Public Property Let Opacity(ByVal Opaqueness As Long)
Attribute Opacity.VB_Description = "Returns/Sets the level of translucency for the control. 100 is fully opaque and 0 is transparent"
    ' Purpose: Set how opaque the image will be rendered at; 100=fully opaque, 0=fully transparent
    If Not Opaqueness = 0 Then
        Opaqueness = Abs(Opaqueness) Mod 100
        If Opaqueness = 0 Then Opaqueness = 100
    End If
    cOpacity = Opaqueness
    PropertyChanged "Opacity"
    UserControl.Refresh
End Property
Public Property Get Opacity() As Long
    Opacity = cOpacity
End Property

Public Property Let Mirror(MirrorType As aiMirrorEnum)
Attribute Mirror.VB_Description = "Returns/Sets the current mirroring effect for the image"
    ' Purpose. Mirror an image horizontally or vertically or both
    If MirrorType >= aiMirrorNone And MirrorType <= aiMirrorAll Then
        sptMirrorImage MirrorType
        cMirror = MirrorType
        If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
        If Me.ShadowEnabled Then sptRefreshShadow
        UserControl.Refresh
        PropertyChanged "Mirror"
    End If
End Property
Public Property Get Mirror() As aiMirrorEnum
    Mirror = cMirror
End Property

Public Property Let KeepOriginalFormat(Keep As Boolean)
Attribute KeepOriginalFormat.VB_Description = "Returns/Sets whether control will maintain original image data"
    ' Forces routines to keep a copy of the image in its original format
    ' This will be used more in upcoming versions
    If Not (((cKeyProps And pKeepBytes) = pKeepBytes) = Keep) Then
        cKeyProps = cKeyProps Xor pKeepBytes
        PropertyChanged "KeepOriginalFormat"
    End If
End Property
Public Property Get KeepOriginalFormat() As Boolean
    KeepOriginalFormat = ((cKeyProps And pKeepBytes) = pKeepBytes)
End Property

Public Property Let grayScale(Style As aiGrayScales)
Attribute grayScale.VB_Description = "Returns/Sets gray scale formula used when rendering image"
    ' Option to toggle grayscale effect
    ' Note: offsetting -1/+1 to align values with the c32bppDIB enumeration
    ' That enumeration does not have a "No Grayscale" option
    If Style >= aiNoGrayScale And Style <= aiBlueGreenMask Then
        If Not cGrayScale = Style - 1 Then
            cGrayScale = Style - 1
            UserControl.Refresh
            PropertyChanged "GrayScale"
        End If
    End If
End Property
Public Property Get grayScale() As aiGrayScales
    grayScale = cGrayScale + 1
End Property

Public Property Let RotateEnabled(canRotate As Boolean)
Attribute RotateEnabled.VB_Description = "Enables/disables the Rotation property."
    If Not (((cKeyProps And pRotated) = pRotated) = Abs(canRotate)) Then
        cKeyProps = cKeyProps Xor pRotated
        sptResize
        PropertyChanged "RotateEnabled"
    End If
End Property
Public Property Get RotateEnabled() As Boolean
    RotateEnabled = ((cKeyProps And pRotated) = pRotated)
End Property

Public Property Let Rotation(ByVal newAngle As Long)
Attribute Rotation.VB_Description = "Returns/sets angle of rotation for the image from 0 to 360 degrees"
    ' Purpose: Rotate an image by set a degree (-360 to 360)
    newAngle = newAngle Mod 360
    If Not cAngle = newAngle Then
        cAngle = newAngle
        If (cKeyProps And pRotated) = pRotated Then
            If sptUpdateOffscreen(False, True) = False Then sptRefreshRegion
            If Me.ShadowEnabled Then sptRefreshShadow
            UserControl.Refresh
            PropertyChanged "Rotation"
        End If
    End If
End Property
Public Property Get Rotation() As Long
Attribute Rotation.VB_Description = "The angle of rotation for the image when RotateEnabled is True"
    Rotation = cAngle
End Property

Public Property Let Enabled(Enable As Boolean)
    ' Purpose: Enable/Disable mouse events
    If Not UserControl.Enabled = Enable Then
        UserControl.Enabled = Enable
        If Ambient.UserMode = True Then
            If Enable = True Then
                cCrossTalker.AddRef ObjPtr(Me)
            Else
                If UserControl.Tag = "Timer" Then
                    KillTimer cTmrHwnd, ObjPtr(Me)
                    UserControl.Tag = vbNullString
                End If
                cCrossTalker.RemoveRef ObjPtr(Me)
            End If
            sptRefreshRegion
        End If
        PropertyChanged "Enabled"
    End If
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let ShadowEnabled(Enable As Boolean)
Attribute ShadowEnabled.VB_Description = "Returns/sets whether shadow will be displayed"
    If Not (((cKeyProps And pShadowed) = pShadowed) = Enable) Then
        cKeyProps = cKeyProps Xor pShadowed
        sptRefreshShadow
        PropertyChanged "ShadowEnabled"
        UserControl.Refresh
    End If
End Property
Public Property Get ShadowEnabled() As Boolean
    ShadowEnabled = Abs((cKeyProps And pShadowed) = pShadowed)
End Property

Public Property Let ShadowColor(Color As OLE_COLOR)
Attribute ShadowColor.VB_Description = "Color to be applied to shadow when ShadowEnabled is True"
    If Not ShadowColor = Color Then
        cShadowColor = Color
        sptRefreshShadow
        PropertyChanged "ShadowColor"
        UserControl.Refresh
    End If
End Property
Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = cShadowColor
End Property

Public Property Let ShadowOffsetX(Offset As Long)
Attribute ShadowOffsetX.VB_Description = "Sets/Returns horizontal offset to be applied to shadow when ShadowEnabled is True"
    If Not cShadowOffset.X = Offset Then
        cShadowOffset.X = Offset
        If Me.ShadowEnabled Then UserControl.Refresh
        PropertyChanged "ShadowOffsetX"
    End If
End Property
Public Property Get ShadowOffsetX() As Long
    ShadowOffsetX = cShadowOffset.X
End Property
Public Property Let ShadowOffsetY(Offset As Long)
Attribute ShadowOffsetY.VB_Description = "Sets/returns vertical offset to be applied to shadow when ShadowEnabled is True"
    If Not cShadowOffset.Y = Offset Then
        cShadowOffset.Y = Offset
        If Me.ShadowEnabled Then UserControl.Refresh
        PropertyChanged "ShadowOffsetY"
    End If
End Property
Public Property Get ShadowOffsetY() As Long
    ShadowOffsetY = cShadowOffset.Y
End Property

Public Property Let ShadowDepth(Depth As Long)
Attribute ShadowDepth.VB_Description = "Blur depth to be applied to shadow when ShadowEnabled is True"
    If Not cShadowDepth = Depth Then
        If Depth > -1 And Depth < 11 Then
            cShadowDepth = Depth
            sptRefreshShadow
            PropertyChanged "ShadowDepth"
            UserControl.Refresh
        End If
    End If
End Property
Public Property Get ShadowDepth() As Long
    ShadowDepth = cShadowDepth
End Property

Public Property Let ShadowOpacity(Opaqueness As Long)
Attribute ShadowOpacity.VB_Description = "Sets/returns the opacity to be applied to shadow when ShadowEnabled is True"
    If Not cShadowOpacity = Opaqueness Then
        If Opaqueness > -1 And Opaqueness < 101 Then
            cShadowOpacity = Opaqueness
            If Me.ShadowEnabled Then UserControl.Refresh
            PropertyChanged "ShadowOpacity"
        End If
    End If
End Property
Public Property Get ShadowOpacity() As Long
    ShadowOpacity = cShadowOpacity
End Property
    
Public Property Let MousePointer(Pointer As MousePointerConstants)
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object"
    ' same as VB's MousePointer property
    On Error Resume Next
    UserControl.MousePointer = Pointer
    If Err Then Err.Clear
    PropertyChanged "MousePointer"
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MouseIcon(Icon As StdPicture)
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon"
    ' same as VB's MouseIcon property
    Set Me.MouseIcon = Icon
End Property
Public Property Set MouseIcon(Icon As StdPicture)
    On Error Resume Next
    Set UserControl.MouseIcon = Icon
    If Err Then Err.Clear
    PropertyChanged "MouseIcon"
End Property
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let OLEDropMode(Value As aiOLEDropMode)
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target"
    ' same as VB's OLEDropMode property
    ' Will forward the OLEDragOver, OLEDragDrop when set
    ' You can use this to allow the control to load an image dragged onto it
    On Error Resume Next
    UserControl.OLEDropMode = Value
    If Err Then Err.Clear
    PropertyChanged "OLEDropMode"
End Property
Public Property Get OLEDropMode() As aiOLEDropMode
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Sub OffsetShadow(X As Long, Y As Long)
Attribute OffsetShadow.VB_Description = "Places the shadow to a desired X,Y coordinate"
    ' sets the shadow offsets in one call preventing multiple refreshes
    cShadowOffset.X = X
    cShadowOffset.Y = Y
    If Me.ShadowEnabled = True Then UserControl.Refresh
End Sub

Public Sub OffsetImage(ByVal SetOffsets As Boolean, ByRef X As Long, ByRef Y As Long, Optional ByRef ShadowX As Long = 0&, Optional ByRef ShadowY As Long = 0&)
Attribute OffsetImage.VB_Description = "Places the image and shadow to a position other than 0,0"
    
    ' sets/returns the current offsets of the image and optionally the shadow
    ' When SetOffsets=True then the image/shadow offsets are modified
    ' When SetOffset=False then the image/shadow offsets are returned
    
    ' The offsets are from the normal position. So if you supply 1,1 then the image will
    ' be shifted one pixel right and down from its normal position, not its current position.
    ' And if you supply 0,0, then the image is shifted back to its normal position.
    
    If SetOffsets Then
        cScaledCoords.xOffset = X
        cScaledCoords.yOffset = Y
        If Not ShadowX = 0 Then cShadowOffset.X = ShadowX
        If Not ShadowY = 0 Then cShadowOffset.Y = ShadowY
        If cRegion = 0& Then
            OffsetRect cRgnBox, -cRgnBox.Left + X, -cRgnBox.Top + Y
        Else
            GetRgnBox cRegion, cRgnBox
            OffsetRgn cRegion, -cRgnBox.Left + X, -cRgnBox.Top + Y
        End If
        UserControl.Refresh
    Else
        If cRegion = 0& Then GetRgnBox cRegion, cRgnBox
        X = cScaledCoords.xOffset
        Y = cScaledCoords.yOffset
        ShadowX = cShadowOffset.X
        ShadowY = cShadowOffset.Y
    End If
    
End Sub

Public Function isMouseOver() As Boolean
Attribute isMouseOver.VB_Description = "Returns whether or not the mouse is currently over the control"
    ' test to see if mouse is currently over the image.
    ' Remember that disabled controls get no mouse events; therefore, this
    ' function will always return false for disabled controls.
    
    ' Also note that the HitTest property determines this return value also
    
    If UserControl.Enabled Then
        isMouseOver = (cCrossTalker.CapturedControl = ObjPtr(Me))
    End If
End Function

Public Property Get GDIplusInitialized() As Boolean
Attribute GDIplusInitialized.VB_Description = "Returns true if control initialized GDI+"
    ' returns whether or not GDI+ has been initialized
    GDIplusInitialized = Not (cCrossTalker.gdiToken = 0)
End Property

Public Property Get GDIplusEnabled() As Boolean
Attribute GDIplusEnabled.VB_Description = "Returns true if GDI+ installed on the system"
    ' read only property to inform user if GDI+ is installed on the system
    GDIplusEnabled = cImage.isGDIplusEnabled
End Property

Public Property Let IntensityOffset(Adjustment As Single)
Attribute IntensityOffset.VB_Description = "Adds/Subtracts lightness. Negative values make image darker, positive values make image lighter"
    If Not cIntesity = Adjustment Then
        If Adjustment >= -100 And Adjustment <= 100 Then
            cIntesity = Adjustment
            PropertyChanged "Intensity"
            UserControl.Refresh
        End If
    End If
End Property
Public Property Get IntensityOffset() As Single
    IntensityOffset = cIntesity
End Property

'=============================================================================================
' /////////////////////// P R O P E R T Y   P A G E   R O U T I N E S \\\\\\\\\\\\\\\\\\\\\\\\
'=============================================================================================

Friend Function ppgGetStream(outStream() As Byte) As Boolean
    ' PROPERTY PAGE USE ONLY. DO NOT MAKE PUBLIC
    ' Allows the property page to retrieve this control's image remotely
    If Not cImage.ImageType = imgCheckerBoard Then ppgGetStream = cImage.GetOrginalFormat(outStream)
End Function
Friend Sub ppgSetStream(inStream() As Byte, cX As Long, cY As Long, bitDepth As Long)
    ' PROPERTY PAGE USE ONLY. DO NOT MAKE PUBLIC
    ' Allows the property page to set this control's new iamge remotely & triggers an activation of the WriteProperties event
    Dim curScale As aiScaleMethod, bAutoSize As Boolean
    cImage.LoadPicture_Stream inStream(), cX, cY, , , True, bitDepth
    curScale = Me.ScaleMethod
    bAutoSize = Me.AutoSize
    cKeyProps = (cKeyProps Or pAutoSize) ' turn autosize on
    cScaleMethod = aiActualSize      ' set scalemode to actual size
    sptReplaceImage
    cScaleMethod = curScale
    If Not bAutoSize Then cKeyProps = cKeyProps Xor pAutoSize
    PropertyChanged "ScaleMethod"
End Sub
Friend Property Get ppgDIBclass() As c32bppDIB
    ' PROPERTY PAGE USE ONLY. DO NOT MAKE PUBLIC
    ' Allows property page to access this controls DIB classes
    Set ppgDIBclass = cImage
End Property

'===============================================================================================
' /////////////////////// S U P P O R T / H E L P E R   R O U T I N E S \\\\\\\\\\\\\\\\\\\\\\\\
'===============================================================================================


Private Sub sptReplaceImage()
    ' Function replaces an image and resizes the control
    ' But it also ensures current settings are applied if applicable and
    ' resets those that cannot apply
    sptMirrorImage 0&       ' mirror the image using current settings
    If cMask.Applied Then   ' did we apply a mask?
        If cImage.Alpha = True Then ' if the image is alpha, then masking is not applicable
            cMask.Applied = False   ' reset key masking values
            cMask.Source = aiNoMask
            PropertyChanged "MaskUsed"
        Else
            Select Case cMask.Source    ' apply the mask to this image
                Case aiUseBottomLeft
                    cMask.AppliedColor = cImage.GetPixel(0, cImage.Height - 1, , False)
                Case aiUseBottomRight
                    cMask.AppliedColor = cImage.GetPixel(cImage.Width - 1, cImage.Height - 1, , False)
                Case aiUseTopLeft
                    cMask.AppliedColor = cImage.GetPixel(0, 0, , False)
                Case aiUseTopRight
                    cMask.AppliedColor = cImage.GetPixel(cImage.Width - 1, 0, , False)
            End Select
            cImage.MakeTransparent sptConvertSysColor(cMask.AppliedColor)
        End If
    End If
    If (cKeyProps And pMasked) = pMasked Then cImage.MakeImageInverse
    sptResize

End Sub

Private Sub sptRefreshRegion()
    ' Routine creates a new hit test region; can be called for many reasons:
    ' - image is mirrored, changing shape of image
    ' - image is resized
    ' - image toggles mask property
    ' - image is rotated
    ' - hit test property changed
    ' - enabled property toggled
    If UserControl.Enabled = True Then
        If Ambient.UserMode = True Then
            If cHitTest = aiEntireControl Or cImage.Handle = 0& Then ' no region needed
                If Not cRegion = 0& Then
                    DeleteObject cRegion
                    cRegion = 0&
                End If
                SetRect cRgnBox, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
            Else
                If Not cRegion = 0& Then DeleteObject cRegion
                If Not cOffscreen Is Nothing Then
                    ' we have offscreen image, use it to create region
                    cRegion = cOffscreen.CreateRegion(cHitTest - 1)
                Else ' we are going to force offscreen to create the region
                    cKeyProps = cKeyProps Or pOffscreen
                    sptUpdateOffscreen True, False
                    cRegion = cOffscreen.CreateRegion(cHitTest - 1)
                    cKeyProps = cKeyProps Xor pOffscreen
                    Set cOffscreen = Nothing
                End If
                If cHitTest = aiBoundingRgn Then
                    ' we don't need to use a GDI resource, so don't
                    GetRgnBox cRegion, cRgnBox
                    DeleteObject cRegion
                    cRegion = 0&
                End If
            End If
        Else    ' design mode, use complete control
            SetRect cRgnBox, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
    Else    ' not used if in runtime, but used if in design time
        SetRect cRgnBox, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
End Sub

Private Sub sptRefreshShadow()

    Dim bUseOwnClass As Boolean
    
    If Me.ShadowEnabled = False Then
        Set cShadowClass = Nothing
    Else
        If cScaledCoords.OneToOne = False Then
            bUseOwnClass = True
        ElseIf (cKeyProps And pRotated) = pRotated Then
            bUseOwnClass = True
        ElseIf (cImage.Width > 128 Or cImage.Height > 128) Then
            bUseOwnClass = True
        End If
        If bUseOwnClass Then
            Dim tImage As c32bppDIB
            If Not cOffscreen Is Nothing Then
                Set cShadowClass = cOffscreen.CreateDropShadow(cShadowDepth, sptConvertSysColor(cShadowColor))
            ElseIf (cKeyProps And pRotated) = pRotated Then
                Set tImage = New c32bppDIB
                cImage.CopyImageTo tImage, cScaledCoords.Width, cScaledCoords.Height
                Set cShadowClass = tImage.CreateDropShadow(cShadowDepth, sptConvertSysColor(cShadowColor))
                Set tImage = Nothing
            Else
                Set tImage = New c32bppDIB
                tImage.InitializeDIB cScaledCoords.Width, cScaledCoords.Height
                cImage.Render 0, 0, 0, cScaledCoords.Width, cScaledCoords.Height, , , , , , , , tImage
                Set cShadowClass = tImage.CreateDropShadow(cShadowDepth, sptConvertSysColor(cShadowColor))
            End If
        ElseIf Not cShadowClass Is Nothing Then
            Set cShadowClass = Nothing
        End If
        If Not cShadowClass Is Nothing Then cShadowClass.gdiToken = cImage.gdiToken
    End If
    
End Sub

Private Function sptConvertSysColor(Color As Long) As Long

    ' Converts VB color constants to real color values
    If Color < 0 Then
        sptConvertSysColor = GetSysColor(Color And &HFF&)
    Else
        sptConvertSysColor = Color
    End If
    
End Function

Private Function sptUpdateOffscreen(bResize As Boolean, bUpdateRegion As Boolean) As Boolean
    ' Purpose: Maintain an offscreen image when the user has set the HasDC property to true
    
    If (cKeyProps And pOffscreen) = pOffscreen Then
        
        Dim hDC As Long
        If bResize = True Or cOffscreen Is Nothing Then
            If cOffscreen Is Nothing Then
                Set cOffscreen = New c32bppDIB
                cOffscreen.gdiToken = cImage.gdiToken
                cOffscreen.ManageOwnDC = True
            End If
            If cScaledCoords.Width = 0 Then Exit Function
            If cOffscreen.Width = UserControl.ScaleWidth And cOffscreen.Height = UserControl.ScaleHeight Then
                cOffscreen.EraseDIB
            Else
                cOffscreen.InitializeDIB UserControl.ScaleWidth, UserControl.ScaleHeight
            End If
        Else
            cOffscreen.EraseDIB
        End If
        
        hDC = cOffscreen.LoadDIBinDC(True)
        If Not hDC = 0 Then
            If (cKeyProps And pRotated) = pRotated Then
                cOffscreen.Alpha = True
                cImage.Render hDC, cOffscreen.Width \ 2, cOffscreen.Height \ 2, cScaledCoords.Width, cScaledCoords.Height, , , , , , , , cOffscreen, , , cAngle, True
            Else
                cOffscreen.Alpha = cImage.Alpha
                cImage.Render hDC, cScaledCoords.Left, cScaledCoords.Top, cScaledCoords.Width, cScaledCoords.Height, , , , , , , False, cOffscreen
            End If
        End If
        If bUpdateRegion Then sptRefreshRegion
        sptUpdateOffscreen = True
    End If
    If Me.ShadowEnabled Then sptRefreshShadow

End Function

Private Sub sptResize()
    ' Purpose: Resize a control but monitor whether or not resizing actually occurred
    ' When we call the resize event which we want to do so the cScaleWidth/Height
    ' variables are recalculated, no repainting will occur if the image is already the
    ' correct size. So monitor and refresh in that case
    Dim cX As Long, cY As Long
    cX = UserControl.ScaleWidth
    cY = UserControl.ScaleHeight
    Call UserControl_Resize
    If UserControl.ScaleWidth = cX Then
        If UserControl.ScaleHeight = cY Then UserControl.Refresh
    End If
End Sub

Private Sub sptUndoMask()
    ' Routine replaces the applied mask color with the original color when the
    ' user toggled the MaskUsed property to false
    Dim X As Long, lColor As Long, srcData() As Byte
    
    ' get the masked color we applied to the image
    If (cKeyProps And pMasked) = pMasked Then
        lColor = sptConvertSysColor(cMask.AppliedColor) Xor vbWhite
    Else
        lColor = sptConvertSysColor(cMask.AppliedColor)
    End If
    ' convert it to BGR and add an alpha value of 255
    lColor = (lColor And &HFF) * &H10000 Or ((lColor \ &H100) And &HFF) * &H100 _
        Or ((lColor \ &H10000) And &HFF) Or &HFF000000
    ' return the RGBA in a 1D array
    cImage.GetDIBbits srcData, False
    ' loop thru. Only those that are fully transparent were changed; reset them
    For X = 3 To UBound(srcData) Step 4&
        If srcData(X) = 0 Then
            CopyMemory srcData(X - 3), lColor, 4&
        End If
    Next
    ' apply the changed RGBA bytes
    cImage.SetDIBbits srcData
End Sub

Private Sub sptMirrorImage(newMirrorValue As aiMirrorEnum)
    ' Purpose: Permanently mirror an image.
    ' Mirroring this way improves rendering speed vs mirroing during rendering.
    ' Unrendering is just as simple and mirroring does not destroy any pixel information
    Dim curMirror As aiMirrorEnum
    curMirror = (cMirror Xor newMirrorValue)
    cImage.MirrorImage ((curMirror And aiMirrorHorizontal) = aiMirrorHorizontal), ((curMirror And aiMirrorVertical) = aiMirrorVertical)
End Sub


Private Sub sptSetCrossTalker()

    Dim tUC As aicAlphaImage
    Dim tIC As ICrossComm
    Dim Index As Long
    Dim myType As String
    
    myType = TypeName(Me)
    With UserControl.ParentControls
        For Index = 0 To .Count - 1
            If TypeName(.Item(Index)) = myType Then ' find control like this one
                Set tUC = .Item(Index)              ' create hard reference
                Set tIC = tUC                       ' create Implementation reference
                Set cCrossTalker = tIC.CrossTalkClass ' get that UC's cUCcrossTalk class
                If cCrossTalker Is Nothing Then
                    Set cCrossTalker = New cCrossComm ' first time, new VB/Form instance
                    If cImage.isGDIplusEnabled Then
                        ' ensure GDI+ is managed for the entire project
                        cCrossTalker.ManageGDIToken containerHwnd
                    End If
                End If
                cImage.gdiToken = cCrossTalker.gdiToken
                If Not cOffscreen Is Nothing Then cOffscreen.gdiToken = cCrossTalker.gdiToken
                If Not cShadowClass Is Nothing Then cShadowClass.gdiToken = cCrossTalker.gdiToken
                cCrossTalker.AddRef ObjPtr(Me) ' have cUCcrossTalk cache reference to Me
                Exit For                      ' done
            End If
        Next
    End With

    If Ambient.UserMode = True Then
        ' timers are only active during runtime
        If cTmrAddrOf = 0& Then cTmrAddrOf = zb_AddressOf(1, 4)
        cTmrHwnd = UserControl.containerHwnd
    End If

End Sub


'============================================================================================
' /////////////////////// U S E R C O N T R O L  P R O P E R T I E S \\\\\\\\\\\\\\\\\\\\\\\\
'============================================================================================

Private Sub UserControl_Initialize()
    Set cImage = New c32bppDIB
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Terminate()
    Set cOffscreen = Nothing            ' clean up any offscreen image
    Set cShadowClass = Nothing          ' clean up a shadow image if used
    If Not cRegion = 0& Then DeleteObject cRegion
    If Not cCrossTalker Is Nothing Then
        cCrossTalker.RemoveRef ObjPtr(Me)   ' should have been handled in UserControl_Hide
        Set cCrossTalker = Nothing
    End If
    zTerminate                          ' release virtual memory
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    cImage.GetDroppedFileNames Data
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(Y, vbPixels, cScaleMode))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(Y, vbPixels, cScaleMode), State)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cBtnClickTracker = (cBtnClickTracker Or Button)   ' track which button(s) are currently down
    RaiseEvent MouseDown(Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(Y, vbPixels, cScaleMode))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(Y, vbPixels, cScaleMode))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cBtnClickTracker = (cBtnClickTracker And &HF) ' pop off any previous last MouseUp flags
    ' track this MouseUp ony if we have a MouseDown for the same button.
    ' When a mouse is dragged from one window and mouse is over our window but the
    ' mouse capture was terminated and mouse is released, we can get a MouseUp being
    ' triggered due to that loss of capture. Prevent that
    If (cBtnClickTracker And Button) = Button Then cBtnClickTracker = (cBtnClickTracker Or Button * &H10)
    RaiseEvent MouseUp(Button, Shift, ScaleX(X, vbPixels, cScaleMode), ScaleY(Y, vbPixels, cScaleMode))
End Sub

Private Sub UserControl_Click()
    ' Usercontrols respond to left, right & middle button clicks
    ' It would be nice if we had a way to distinguish which button caused the click,
    ' but as you can see, there is no Button parameter provided; so we do it ourselves.
    ' see UserControl_MouseDown and MouseUp events that track this variable
    RaiseEvent Click(cBtnClickTracker \ &H10)
End Sub

Private Sub UserControl_DblClick()
    ' See Usercontrol_Click event comments
    RaiseEvent DblClick(cBtnClickTracker \ &H10)
End Sub

Private Sub UserControl_Hide()
    ' control is made invisible by user or is about to be destroyed
    If Not cFader.fStep = 0& Then
        ' stop fader timer
        KillTimer cTmrHwnd, -ObjPtr(Me)
        cFader.fStep = 0&
    End If
    If UserControl.Tag = "Timer" Then
        KillTimer cTmrHwnd, ObjPtr(Me)
        UserControl.Tag = vbNullString
    End If
    ' stop mouse_exit timer & clean up references as needed
    If UserControl.Enabled = True Then cCrossTalker.RemoveRef ObjPtr(Me)
End Sub

Private Sub UserControl_Show()

    If UserControl.Enabled = True Then cCrossTalker.AddRef ObjPtr(Me)
    If cScaledCoords.Width = 0& Then sptResize
    ' Thanx to Soorya for helping trouble shoot above issue.
    ' A usercontrol while in IDE will not always fire a Resize event before it is
    ' first displayed. Since scaling ratios are calculated based off the usercontrol's
    ' size and Stretch property, we need it fired, so fire it here if not previously
    ' fired; otherwise images are not drawn correctly or not drawn at all!
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)

    ' For windowless, DC-less controls, this is the only way to let VB know
    ' whether or not we want the mouse or let it go to next object in zOrder.
    ' Note: This is used both during design time and run time
    
    If cRegion = 0& Then
        If Not PtInRect(cRgnBox, X, Y) = 0& Then HitResult = vbHitResultHit
    ElseIf Not PtInRegion(cRegion, X, Y) = 0& Then
        HitResult = vbHitResultHit
    End If
    
    If Ambient.UserMode = True Then
        ' testing for mouse enter/mouse leave
        ' See sptValidateSession and sptInvalidateSession regarding these window properties
        
        Dim myPointer As Long, mPoint As POINTAPI
        
        ' get object pointer of control currently having mouse enter
        myPointer = ObjPtr(Me)
        
        Select Case cCrossTalker.CapturedControl
        Case myPointer ' same control; no change?
            If HitResult = 0& Then  ' we are no longer in the hit region
                ' trigger a mouse leave event
                KillTimer cTmrHwnd, myPointer
                UserControl.Tag = vbNullString
                cCrossTalker.CapturedControl = 0&
                RaiseEvent MouseExit
            End If
        Case 0& ' no alphaImage control has the mouse, but we may now
            If HitResult = vbHitResultHit Then
                ' convert screen coordinates to relative container coords
                GetCursorPos cTopLeftPos
                ClientToScreen cTmrHwnd, mPoint
                cTopLeftPos.X = cTopLeftPos.X - mPoint.X - X
                cTopLeftPos.Y = cTopLeftPos.Y - mPoint.Y - Y
                ' trigger a mouse enter event
                cCrossTalker.CapturedControl = myPointer
                RaiseEvent MouseEnter
                SetTimer cTmrHwnd, myPointer, 100, cTmrAddrOf
                UserControl.Tag = "Timer"
            End If
        Case Else   ' some other control has the mouse, but no longer, cause we do
            
            cCrossTalker.CapturedControl = 0&   ' forces a MouseExit from that control if needed
            ' now update the window property and call mouse enter if needed
            If HitResult = vbHitResultHit Then
                ' convert screen coordinates to relative container coords
                GetCursorPos cTopLeftPos
                ClientToScreen cTmrHwnd, mPoint
                cTopLeftPos.X = cTopLeftPos.X - mPoint.X - X
                cTopLeftPos.Y = cTopLeftPos.Y - mPoint.Y - Y
                ' firre the mouse enter event
                cCrossTalker.CapturedControl = myPointer
                RaiseEvent MouseEnter
                SetTimer cTmrHwnd, myPointer, 100, cTmrAddrOf
                UserControl.Tag = "Timer"
            End If
        End Select
    End If
End Sub

Private Sub UserControl_Paint()
    
    
    If cScaledCoords.Width < 0 Then Exit Sub
    
    Dim uRect As RECT, iRect As RECT, clipRect As RECT
    
    GetClipBox UserControl.hDC, uRect
    
    ' see if our image is within the invalidated region
    If (cKeyProps And pRotated) = pRotated Then
        SetRect iRect, cScaledCoords.Left, cScaledCoords.Top, cScaledCoords.Left + cScaledCoords.RotatedSize, cScaledCoords.Top + cScaledCoords.RotatedSize
    Else
        SetRect iRect, cScaledCoords.Left, cScaledCoords.Top, cScaledCoords.Left + cScaledCoords.Width, cScaledCoords.Top + cScaledCoords.Height
    End If
    If Not IntersectRect(clipRect, uRect, iRect) = 0 Then
        ' paint the shadow first if applicable
        If Me.ShadowEnabled Then
            If cShadowClass Is Nothing Then
                cImage.RenderDropShadow_JIT UserControl.hDC, (UserControl.ScaleWidth - cImage.Width) \ 2 + cShadowOffset.X, (UserControl.ScaleHeight - cImage.Height) \ 2 + cShadowOffset.Y, _
                    cShadowDepth, sptConvertSysColor(cShadowColor), (cShadowOpacity * cOpacity) \ 100
            
            ElseIf (cKeyProps And pRotated) = pRotated Then
                cShadowClass.Render UserControl.hDC, UserControl.ScaleWidth \ 2 + cShadowOffset.X, UserControl.ScaleHeight \ 2 + cShadowOffset.Y, cScaledCoords.Width, cScaledCoords.Height, , , , , (cShadowOpacity * cOpacity) \ 100, , , , cGrayScale, , cAngle, True
            Else
                cShadowClass.Render UserControl.hDC, (UserControl.ScaleWidth - cShadowClass.Width) \ 2 + cShadowOffset.X, _
                        (UserControl.ScaleHeight - cShadowClass.Height) \ 2 + cShadowOffset.Y, , , , , , , (cShadowOpacity * cOpacity) \ 100, , False
            End If
        End If
            
        ' now paint the image
        If Not cOffscreen Is Nothing Then       ' AutoRedraw=True, we have a sized copy, use it
            clipRect.Right = clipRect.Right - clipRect.Left
            clipRect.Bottom = clipRect.Bottom - clipRect.Top
            
            cOffscreen.Render UserControl.hDC, clipRect.Left + cScaledCoords.xOffset, clipRect.Top + cScaledCoords.yOffset, clipRect.Right, clipRect.Bottom, _
                clipRect.Left, clipRect.Top, clipRect.Right, clipRect.Bottom, cOpacity, , , , cGrayScale, cIntesity
            
        ElseIf (cKeyProps And pRotated) = pRotated Then ' rotating; slowest rendering method
            cImage.Render UserControl.hDC, UserControl.ScaleWidth \ 2 + cScaledCoords.xOffset, UserControl.ScaleHeight \ 2 + cScaledCoords.yOffset, cScaledCoords.Width, cScaledCoords.Height, , , , , cOpacity, , , , cGrayScale, cIntesity, cAngle, True
        
        ElseIf cScaledCoords.OneToOne Then ' fast rendering
            clipRect.Right = clipRect.Right - clipRect.Left
            clipRect.Bottom = clipRect.Bottom - clipRect.Top
        
            cImage.Render UserControl.hDC, clipRect.Left + cScaledCoords.xOffset, clipRect.Top + cScaledCoords.yOffset, clipRect.Right, clipRect.Bottom, _
                clipRect.Left - cScaledCoords.Left, clipRect.Top - cScaledCoords.Top, clipRect.Right, clipRect.Bottom, cOpacity, , , , cGrayScale, cIntesity
        
        Else    ' resized; besides rotation; next slowest rendering method
            cImage.Render UserControl.hDC, cScaledCoords.Left + cScaledCoords.xOffset, cScaledCoords.Top + cScaledCoords.yOffset, cScaledCoords.Width, cScaledCoords.Height, , , , , cOpacity, , , , cGrayScale, cIntesity
            
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    ' default properties for new controls
    sptSetCrossTalker
    cOpacity = 100&
    cGrayScale = aiNoGrayScale - 1& ' no grayscale
    cScaleMethod = aiScaled         ' default scale method
    cMask.Color = vbButtonFace      ' default mask color
    cKeyProps = pHiQualty Or pAutoSize ' set high quality to true & AutoSize to True by default
    cShadowColor = &H404040         ' default shadow color (Dark Gray)
    cShadowDepth = 2                ' default blur depth
    cShadowOffset.X = 2             ' default shadow offsets (towards bottom right of image)
    cShadowOffset.Y = 2
    cShadowOpacity = 50             ' default shadow opacity
    cImage.HighQualityInterpolation = True
    On Error Resume Next
    ' attempt to get actual scalemode. Errors occur when container has no scalemode (i.e.,frame)
    cScaleMode = UserControl.Parent.ScaleMode
    If Err Then
        Err.Clear
        cScaleMode = vbContainerPosition
        '^^ the above can be used for any scalemode; however, when container scalemode is
        ' vbTwips or vbPixels, the returned X,Y coordinates passed back are decimal not integer
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim srcData() As Byte
    sptSetCrossTalker
    With PropBag
        srcData = .ReadProperty("Image", srcData)
        cScaleMethod = .ReadProperty("Scaler", aiScaled)
        cOpacity = .ReadProperty("Opacity", 100)
        cMirror = .ReadProperty("Mirror", aiMirrorNone)
        cAngle = .ReadProperty("Angle", 0&)
        cGrayScale = .ReadProperty("GrayScale", aiNoGrayScale - 1)
        cHitTest = .ReadProperty("HitTest", aiEntireControl)
        cKeyProps = .ReadProperty("Props", pHiQualty)
        cMask.Applied = .ReadProperty("MaskUsed", False)
        cMask.Color = .ReadProperty("MaskColor", vbButtonFace)
        cMask.Source = .ReadProperty("MaskSource", aiNoMask)
        If cMask.Applied Then cMask.AppliedColor = .ReadProperty("Mask", vbButtonFace)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set UserControl.MouseIcon = .ReadProperty("MIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MPointer", vbDefault)
        UserControl.OLEDropMode = .ReadProperty("OLEdrop", vbOLEDropNone)
        If cScaleMethod = aiLockScale Then
            cScaledCoords.Width = .ReadProperty("ScaleCx", 0&)
            cScaledCoords.Height = .ReadProperty("ScaleCy", 0&)
        End If
        cShadowColor = .ReadProperty("ShadowColor", &H404040)
        cShadowDepth = .ReadProperty("ShadowDepth", 2)
        cShadowOffset.X = .ReadProperty("ShadowX", 2)
        cShadowOffset.Y = .ReadProperty("ShadowY", 2)
        cShadowOpacity = .ReadProperty("ShadowOpacity", 50)
        cIntesity = .ReadProperty("Intensity", 0!)
    End With
    
    On Error Resume Next
    ' attempt to get actual scalemode. Errors occur when container has no scalemode (i.e.,frame)
    cScaleMode = UserControl.Parent.ScaleMode
    If Err Then
        Err.Clear
        cScaleMode = vbContainerPosition
        '^^ the above can be used for any scalemode; however, when container scalemode is
        ' vbTwips or vbPixels, the returned X,Y coordinates passed back are decimal not integer
    End If
    
    ' load/validate the image, set it up iniitally or fall back to checkboard state
    cImage.HighQualityInterpolation = ((cKeyProps And pHiQualty) = pHiQualty)
    If cImage.LoadPicture_Stream(srcData, UserControl.ScaleWidth, UserControl.ScaleHeight, , , (Ambient.UserMode = False Or ((cKeyProps And pKeepBytes) = pKeepBytes))) = True Then
        sptMirrorImage 0&
        If cMask.Applied Then cImage.MakeTransparent sptConvertSysColor(cMask.AppliedColor)
        If (cKeyProps And pMasked) = pMasked Then cImage.MakeImageInverse
        If ((cKeyProps And pAutoRedraw) = pAutoRedraw) Then
            If Ambient.UserMode = True Then
                cKeyProps = cKeyProps Or pOffscreen
                If sptUpdateOffscreen(True, True) = False Then sptRefreshRegion
            End If
        End If
        If Me.ShadowEnabled Then sptRefreshShadow
    Else
        If cScaleMethod = aiLockScale Then cScaleMethod = aiScaled
        cKeyProps = (cKeyProps And (pHiQualty Or pKeepBytes))
        cAngle = 0
        cOpacity = 100
        cMirror = aiMirrorNone
        cScaledCoords.Height = 0: cScaledCoords.Width = 0
    End If
    Call UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim srcData() As Byte
    If Not cImage.ImageType = imgCheckerBoard Then
        Select Case cImage.ImageType
            Case imgBitmap, imgBmpARGB, imgBmpPARGB
                If cImage.isGDIplusEnabled Or cImage.isZlibEnabled Then
                    sptMirrorImage 0&
                    If cMask.Applied = True Then sptUndoMask
                    If (cKeyProps And pMasked) = pMasked Then cImage.MakeImageInverse
                    cImage.SaveToStream_PNG srcData
                End If
            Case Else
                cImage.GetOrginalFormat srcData
        End Select
    End If
        
    With PropBag
        .WriteProperty "Image", srcData
        .WriteProperty "Scaler", cScaleMethod, aiScaled
        .WriteProperty "Opacity", cOpacity, 100
        .WriteProperty "Mirror", cMirror, aiMirrorNone
        .WriteProperty "Angle", cAngle, 0&
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "MPointer", UserControl.MousePointer, vbDefault
        .WriteProperty "OLEdrop", UserControl.OLEDropMode, vbOLEDropNone
        .WriteProperty "MIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "GrayScale", cGrayScale, aiNoGrayScale - 1
        .WriteProperty "HitTest", cHitTest, aiEntireControl
        .WriteProperty "Props", (cKeyProps And Not pOffscreen), pHiQualty
        .WriteProperty "MaskUsed", cMask.Applied, False
        .WriteProperty "MaskColor", cMask.Color, vbButtonFace
        .WriteProperty "MaskSource", cMask.Source, aiNoMask
        If cMask.Applied Then .WriteProperty "Mask", cMask.AppliedColor, vbButtonFace
        .WriteProperty "ShadowColor", cShadowColor, &H404040
        .WriteProperty "ShadowDepth", cShadowDepth, 2
        .WriteProperty "ShadowX", cShadowOffset.X, 2
        .WriteProperty "ShadowY", cShadowOffset.Y, 2
        .WriteProperty "ShadowOpacity", cShadowOpacity, 50
        .WriteProperty "Intensity", cIntesity, 0!
        
        If cScaleMethod = aiLockScale Then
            .WriteProperty "ScaleCx", cScaledCoords.Width, 0&
            .WriteProperty "ScaleCy", cScaledCoords.Height, 0&
        End If
    End With
End Sub

Private Sub UserControl_Resize()
    
    If cScaledCoords.Width < 0 Then Exit Sub
    
    If cImage.Handle = 0& Then
        If Ambient.UserMode = False Then
            cImage.InitializeDIB UserControl.ScaleWidth, UserControl.ScaleHeight
            cImage.CreateCheckerBoard
        End If
    End If
    
    GetImageScales cScaledCoords.Width, cScaledCoords.Height, cScaleMethod, _
                    UserControl.ScaleWidth, UserControl.ScaleHeight, ((cKeyProps And pRotated) = pRotated)
    
    With cScaledCoords
        If (cKeyProps And pRotated) = pRotated Then
            .RotatedSize = Sqr(.Width * .Width + .Height * .Height)
            If cScaleMethod = aiLockScale Then
                .Left = (UserControl.ScaleWidth - .Width) \ 2
                If .Left < 0 Then .Left = 0
                .Top = (UserControl.ScaleHeight - .Height) \ 2
                If .Top < 0 Then .Top = 0
            Else
                .Left = (UserControl.ScaleWidth - .RotatedSize) \ 2
                .Top = (UserControl.ScaleHeight - .RotatedSize) \ 2
            End If
            .OneToOne = False
        Else
            .Left = (UserControl.ScaleWidth - .Width) \ 2
            If .Left < 0 Then .Left = 0
            .Top = (UserControl.ScaleHeight - .Height) \ 2
            If .Top < 0 Then .Top = 0
            If .Height = cImage.Height Then
                .OneToOne = (.Width = cImage.Width)
            Else
                .OneToOne = False
            End If
        End If
    End With
    
    If ((cKeyProps And pAutoSize) = pAutoSize) Then
        Dim sizeCx As Long, sizeCy As Long
        If (cKeyProps And pRotated) = pRotated Then
            sizeCx = cScaledCoords.RotatedSize
            sizeCy = sizeCx
        Else
            sizeCx = cScaledCoords.Width
            sizeCy = cScaledCoords.Height
        End If
        cScaledCoords.Left = 0: cScaledCoords.Top = 0
        cScaledCoords.Width = -cScaledCoords.Width
        UserControl.Size ScaleX(sizeCx, vbPixels, vbTwips), ScaleY(sizeCy, vbPixels, vbTwips)
        cScaledCoords.Width = -cScaledCoords.Width
    End If
    If sptUpdateOffscreen(True, True) = False Then sptRefreshRegion
    
End Sub

'=============================================================================================
' ////////////////// I N T E R - C O N T R O L   C O M M U N I C A T I O N \\\\\\\\\\\\\\\\\\\
'=============================================================================================

Private Property Get ICrossComm_CrossTalkClass() As Object
    Set ICrossComm_CrossTalkClass = cCrossTalker
End Property

Private Sub ICrossComm_ReleaseMouseCapture()
    ' When this control thinks it has the mouse over it, and another
    ' control receives a mouse event before the timer fires, that
    ' control will forward a mouse exit via this routine.
    If UserControl.Tag = "Timer" Then
        KillTimer cTmrHwnd, ObjPtr(Me)
        UserControl.Tag = vbNullString
    End If
    RaiseEvent MouseExit
End Sub


'============================================================================================
' /////////////////// C A L L B A C K   T H U N K I N G   R O U T I N E S \\\\\\\\\\\\\\\\\\\
'============================================================================================

'*************************************************************************************************
'* cCallback - Class generic callback template
'*
'* Note:
'*  The callback declarations and code are exactly the same for a Class, Form or UserControl.
'*  The callback declarations and code can co-exist with subclassing declarations and code.
'*    With both types of code in a single file,..
'*      delete the duplicated declarations and code, Ctrl+F5 will find them for you
'*      pay careful attention to the nOrdinal parameter to zAddressOf
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 The original..................................................................... 20060408
'* v1.1 Added multi-thunk support........................................................ 20060409
'* v1.2 Added optional IDE protection.................................................... 20060411
'* v1.3 Added an optional callback target object......................................... 20060413
'*************************************************************************************************

'-Callback code-----------------------------------------------------------------------------------
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
'*************************************************************************************************
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* nParamCount  - The number of parameters that will callback
'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
    Const MAX_FUNKS   As Long = 2                                               'Number of simultaneous thunks, adjust to taste
    Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
    Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
    Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  Dim nOffset     As Long
  Dim z_Cb()      As Long    'Callback thunk array
  
  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
    MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
  If nAddr = 0 Then
    MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
  
    If bIdeSafety Then                                                        'If the user wants IDE protection
      z_Cb(2, 0) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode") 'EbMode Address
    End If
    z_Cb(3, 0) = GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
    z_Cb(4, 0) = &HBB60E089
    z_Cb(6, 0) = &H73FFC589: z_Cb(7, 0) = &HC53FF04: z_Cb(8, 0) = &H7B831F75: z_Cb(9, 0) = &H20750008: z_Cb(10, 0) = &HE883E889: z_Cb(11, 0) = &HB9905004: z_Cb(13, 0) = &H74FF06E3: z_Cb(14, 0) = &HFAE2008D: z_Cb(15, 0) = &H53FF33FF: z_Cb(16, 0) = &HC2906104: z_Cb(18, 0) = &H830853FF: z_Cb(19, 0) = &HD87401F8: z_Cb(20, 0) = &H4589C031: z_Cb(21, 0) = &HEAEBFC
    
    For nOffset = 1 To MAX_FUNKS - 1                   ' for each thunk, copy the base thunk
        CopyMemory z_Cb(0, nOffset), z_Cb(0, 0), FUNK_LEN
    Next
    CopyMemory ByVal z_CbMem, z_Cb(0, 0), MEM_LEN      'Copy thunk code to executable memory
  End If
    
  nOffset = z_CbMem + nThunkNo * FUNK_LEN
  CopyMemory ByVal nOffset, ObjPtr(oCallback), 4&      'copy to vMem the objPtr for this thunk
  CopyMemory ByVal nOffset + 4, nAddr, 4&              'copy to vMem the callback addr
  CopyMemory ByVal nOffset + 20, nOffset, 4&           'copy to vMem the start of this thunk
  CopyMemory ByVal nOffset + 48, nParamCount, 4&       'copy to vMem the param count
  CopyMemory ByVal nOffset + 68, nParamCount * 4, 4&   'copy to vMem the total length of parameters
  zb_AddressOf = nOffset + 16                          'return where in vMem this thunk can be called
  
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  CopyMemory nAddr, ByVal ObjPtr(oCallback), 4                              'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    CopyMemory nAddr, ByVal i, 4                                            'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      CopyMemory zAddressOf, ByVal i - (nOrdinal * 4), 4                    'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    CopyMemory bVal, ByVal nAddr, 1                                         'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      CopyMemory zAddressOf, ByVal i - (nOrdinal * 4), 4                    'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    CopyMemory nEntry, ByVal nAddr, 4                                       'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      CopyMemory bVal, ByVal nEntry, 1                                      'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Sub zTerminate()
    
    Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag
    If Not z_CbMem = 0 Then                                            'If memory allocated
        VirtualFree z_CbMem, 0, MEM_RELEASE
        z_CbMem = 0  'Release; Indicate memory released
    End If
End Sub

'*************************************************************************************************
'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************
'Callback ordinal 2
Private Function Timer_Fader(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long

    KillTimer hWnd, TimerID    ' stop current timer
    
    ' ensure new opacity does not exceed final opacity
    If cFader.fStep < 0 Then    ' are we fading out? else we are fading in
        If cFader.fStep + cOpacity <= cFader.fOpacity Then cFader.fStep = 0&
    Else
        If cFader.fStep + cOpacity >= cFader.fOpacity Then cFader.fStep = 0&
    End If
    
    If cFader.fStep = 0& Then   ' fade to next step
        Me.Opacity = cFader.fOpacity
        RaiseEvent FadeTerminated(cOpacity)
    Else
        Me.Opacity = cOpacity + cFader.fStep
        SetTimer hWnd, TimerID, cFader.fDelay, cFader.tmrAddr
    End If
    
End Function

'Callback ordinal 1
Private Function Timer_MouseExit(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    
    KillTimer hWnd, TimerID    ' stop current timer
    
    Dim mHwnd As Long
    Dim tPoint As POINTAPI, mPoint As POINTAPI
    Dim bReset As Boolean
    
    ' validate that our control still thinks it has the mouse
    If cCrossTalker.CapturedControl = ObjPtr(Me) Then
    
        GetCursorPos mPoint                             ' get current mouse position (screen coords)
        mHwnd = WindowFromPoint(mPoint.X, mPoint.Y)     ' see if mouse is over control's container
        If mHwnd = hWnd Then
            ClientToScreen hWnd, tPoint
            ' adjust the points to control coordinates vs screen coordinates
            tPoint.X = mPoint.X - tPoint.X - cTopLeftPos.X
            tPoint.Y = mPoint.Y - tPoint.Y - cTopLeftPos.Y
            
            If cRegion = 0& Then
                ' we don't have a region, so the entire control is valid for hit testing
                If Not PtInRect(cRgnBox, tPoint.X, tPoint.Y) = 0 Then bReset = True ' restart the timer, no change
            Else
                ' we do have a region, see if the point is still within that region
                If Not PtInRegion(cRegion, tPoint.X, tPoint.Y) = 0 Then bReset = True
            End If
        End If
        If bReset Then
            ' set timer for next check
            SetTimer hWnd, TimerID, 100, cTmrAddrOf
        Else
            ' mouse no longer over the control's region/shape; fire MouseExit
            cCrossTalker.CapturedControl = 0&
            UserControl.Tag = vbNullString
            RaiseEvent MouseExit
        End If
    Else    ' control lost mouse before timer fired; do nothing
        UserControl.Tag = vbNullString
    End If
    
eh:
' CAUTION: DO NOT ADD ANY ADDITIONAL CODE OR COMMENTS PAST THE "END FUNCTION"
'          STATEMENT BELOW. Paul Caton's zProbe routine will read it as a start
'          of a new function/sub and the class timer's will fail every time.
End Function
