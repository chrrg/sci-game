VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCrossComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private m_Pointers() As Long
Private m_Count As Long
' Note. I have opted for CopyMemory & a binary sort routine since Longs are cached.
' There is no reason why a collection object cannot be used instead.

Private m_ICrossTalk As ICrossComm
Private m_GDIplusToken As Long      ' shared GDI+ token. All aicAlphaImage controls on a single form share it
Private m_CapturedObject As Long    ' which control currently is tracking mouse enter/exit


' GDI+ startup
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
' ^^ shut down occurs in the thunk. See ManageGDIToken

' ***** THUNK APIs and VARIABLES
Private Declare Function CreateWindowExA Lib "USER32.DLL" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowLong Lib "USER32.DLL" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "USER32.DLL" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "USER32.DLL" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const GWL_WNDPROC As Long = -4
Private Const GW_OWNER As Long = 4
Private Const WS_CHILD As Long = &H40000000

Public Property Let CapturedControl(ByVal ucPointer As Long)
    ' called by usercontrol to set which control is tracking mouse
    Dim Index As Long, tObject As Object
    If Not m_CapturedObject = 0& Then                               ' do we have a object tracking the mouse?
        If Not ucPointer = m_CapturedObject Then                    ' not called by the same object?
            Set m_ICrossTalk = ObjectFromPointer(m_CapturedObject)  ' call local routine to create Implementation
            'm_ICrossTalk.ReleaseMouseCapture                        ' call implementation
            Set m_ICrossTalk = Nothing                              ' release any hard refs
        End If
    End If
    m_CapturedObject = ucPointer                                    ' set current control pointer
    
    
    
 
    
    
    
    
    
End Sub
Public Property Get CapturedControl() As Long
    CapturedControl = m_CapturedObject
End Property

Public Property Get gdiToken() As Long  ' Read Only
    ' previous version returned an actual token.
    ' This version simply returns non-zero if GDI+ running
    gdiToken = m_GDIplusToken
End Property

Public Sub AddRef(ByVal ucPointer As Long)
    ' add pointer to a collection/array of pointers
    Dim bIsNew As Boolean
    Dim Index As Long
    If Not ucPointer = 0& Then          ' cannot accept zero as a pointer!
        If m_Count = 0& Then            ' first one
            ReDim m_Pointers(1 To 6)    ' add a buffer of 5 for additional pointers as needed
            m_Count = 1
            m_Pointers(m_Count) = ucPointer
        Else
            Index = FindPointer(ucPointer, bIsNew)  ' see if we have this one already
            If bIsNew Then                          ' nope, let's add it
                m_Count = m_Count + 1               ' increment & resize array as needed
                If m_Count > UBound(m_Pointers) Then ReDim Preserve m_Pointers(1 To m_Count + 5)
                If Index < m_Count Then             ' shift pointers
                    CopyMemory m_Pointers(Index + 1), m_Pointers(Index), (m_Count - Index) * 4&
                End If
                m_Pointers(Index) = ucPointer       ' insert new pointer
            End If
        End If
    End If
End Sub

Public Sub RemoveRef(ByVal ucPointer As Long)
    ' delete pointer from a collection/array of pointers
    Dim bIsNew As Boolean
    Dim Index As Long
    If Not m_Count = 0& Then
        Index = FindPointer(ucPointer, bIsNew)  ' see if we have this pointer
        If bIsNew = False Then                  ' yes, remove it
            If m_Count = 1 Then                 ' last one?
                m_Count = 0&                    ' if so, erase array & reset counter
                Erase m_Pointers
            Else
                If Index < m_Count Then         ' if not last in the array, shift array
                    CopyMemory m_Pointers(Index), m_Pointers(Index + 1), (m_Count - Index) * 4&
                End If
                m_Count = m_Count - 1           ' decrement counter & remove pointer
                If UBound(m_Pointers) > m_Count + 5 Then
                    ' resize array as needed
                    ReDim Preserve m_Pointers(1 To m_Count)
                End If
            End If
        End If
    End If
    If m_CapturedObject = ucPointer Then m_CapturedObject = 0&
End Sub

Private Function ObjectFromPointer(Index As Long) As Object
    ' creates a hard reference from a object pointer
    Dim tObject As Object
    CopyMemory tObject, Index, 4&   ' create soft ref to UC
    Set ObjectFromPointer = tObject ' create hard ref :: reference count incremented
    CopyMemory tObject, 0&, 4&      ' release soft ref
End Function

Private Function FindPointer(ByVal Criteria As Long, ByRef isNew As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between 1 and 20 iterations
    
    ' [in] Criteria. A value to search for. Order is always ascending
    ' [out] isNew. If Criteria not found, isNew is True else False
    ' [out] Return value: The Index where Criteria was found
    '                     or where the new Criteria should be inserted

    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    UB = m_Count
    LB = 1&
    
    Do Until LB > UB
        newIndex = LB + ((UB - LB) \ 2&)
        If m_Pointers(newIndex) = Criteria Then
            Exit Do
        ElseIf m_Pointers(newIndex) > Criteria Then ' new criteria is lower in sort order
            UB = newIndex - 1&
        Else ' new criteria is higher in sort order
            LB = newIndex + 1&
        End If
    Loop

    If LB > UB Then  ' criteria was not found
            
        If Criteria > m_Pointers(newIndex) Then newIndex = newIndex + 1&
        isNew = True
        
    Else
        isNew = False
    End If
    
    FindPointer = newIndex

End Function

Public Sub ManageGDIToken(ByVal containerHwnd As Long)
        
    ' On XP-pro (maybe not all builds), starting GDI+ and not shutting it down
    ' can cause IDE to crash. GDI+ shutdown occurs during normal termination code.
    ' But executing and END command or IDE toolbar's STOP button prevents termination
    ' code from executing.
    
    ' The fix. Create one window on the top most VB window. All VB apps have this.
    ' The window will terminate GDI+ when VB closes.  Now, all this is not really
    ' needed. Why? If we call GDIplusStartup, render image, and then call
    ' GdiplusShutdown no crash can occur. However, staring/stopping GDI+
    ' everytime an image needs to be repainted takes a performance hit. This can be
    ' measured on images that repaint on a short timer, say a clock. It is more
    ' efficient to start GDI+ once and terminate it when app closes.
    
    ' So this function will create (if needed) only one API window for the entire
    ' VB instance and remains as long as VB is opened. The API window is a child
    ' of a hidden VB top-level window and when that window closes, so will the
    ' API window which then will release GDI+. By keeping its own reference to
    ' the GDI+ library, there is no way for that library to be unloaded without
    ' the API window closing.
        
    On Error Resume Next
    If containerHwnd = 0& Then Exit Sub
    
    Dim cHwnd As Long
    
    ' find the hidden VB owner window. All VB apps have this (run-time & design-time).
    Do
        cHwnd = GetParent(containerHwnd)
        If Not cHwnd = 0& Then containerHwnd = cHwnd
    Loop Until cHwnd = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        cHwnd = GetWindow(containerHwnd, GW_OWNER)
        If Not cHwnd = 0& Then containerHwnd = cHwnd
    Loop Until cHwnd = 0&
    
    If FindWindowEx(containerHwnd, 0&, "Static", "GDI+Safe Patch") Then
        m_GDIplusToken = -1 ' we already have a manager running for this VB instance
        Exit Sub            ' can abort
    End If
    
    Dim gdiSI           As GdiplusStartupInput  'GDI+ startup info
    
    gdiSI.GdiplusVersion = 1                    ' attempt to start
    GdiplusStartup m_GDIplusToken, gdiSI
    If m_GDIplusToken = 0& Then                 ' failed to start
        If Err Then Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, containerHwnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = zFnAddr("user32", "CallWindowProcA")                ' function pointer CallWindowProc
            z_Code(3) = zFnAddr("kernel32", "VirtualFree")                  ' function pointer VirtualFree
            z_Code(4) = zFnAddr("kernel32", "FreeLibrary")                  ' function pointer FreeLibrary
            z_Code(5) = m_GDIplusToken                                      ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = zFnAddr("user32", "SetWindowLongA")                 ' function pointer SetWindowLong
            z_Code(8) = zFnAddr("user32", "SetTimer")                       ' function pointer SetTimer
            z_Code(9) = zFnAddr("user32", "KillTimer")                      ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
        
        Else
        
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            'MsgBox "Failed to create monitoring window", vbExclamation + vbOKOnly, "Error"
            
        End If
    
    Else
        
      VirtualFree z_ScMem, 0, MEM_RELEASE           ' failure - release memory
      'MsgBox "Failed to initialize thunk memory", vbExclamation + vbOKOnly, "Error"
        
    End If
    
End Sub

Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
'Return the address of the specified DLL/procedure

  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
  Debug.Assert zFnAddr                                     'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")

End Function


