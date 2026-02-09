Option Explicit

#If VBA7 Then
    # ruleid: vba-native-api-declaration
    Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
        ByVal lpAddress As LongPtr, _
        ByVal dwSize As LongPtr, _
        ByVal flAllocationType As Long, _
        ByVal flProtect As Long) As LongPtr
    # ruleid: vba-native-api-declaration
    Private Declare PtrSafe Function CreateThread Lib "kernel32" ( _
        ByVal lpThreadAttributes As LongPtr, _
        ByVal dwStackSize As LongPtr, _
        ByVal lpStartAddress As LongPtr, _
        ByVal lpParameter As LongPtr, _
        ByVal dwCreationFlags As Long, _
        ByRef lpThreadId As Long) As LongPtr
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    # ruleid: vba-native-api-declaration
    Private Declare Function VirtualAlloc Lib "kernel32" ( _
        ByVal lpAddress As Long, _
        ByVal dwSize As Long, _
        ByVal flAllocationType As Long, _
        ByVal flProtect As Long) As Long

    # ruleid: vba-native-api-declaration
    Private Declare Function CreateThread Lib "kernel32" ( _
        ByVal lpThreadAttributes As Long, _
        ByVal dwStackSize As Long, _
        ByVal lpStartAddress As Long, _
        ByVal lpParameter As Long, _
        ByVal dwCreationFlags As Long, _
        ByRef lpThreadId As Long) As Long

    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const PAGE_READWRITE As Long = &H4

Sub Demo_NativeApi_Pattern_NoShellcode()
    Dim addr As LongPtr
    Dim tid As Long
    Dim th As LongPtr

    ' Allocate some RW memory (this alone can be suspicious in VBA malware chains)
    # ruleid: vba-native-api-declaration
    addr = VirtualAlloc(0, 4096, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
    If addr = 0 Then
        MsgBox "VirtualAlloc failed", vbExclamation
        Exit Sub
    End If

    ' Create a thread starting at a *legitimate* API function (GetTickCount).
    ' This does NOT execute injected bytes; it's just illustrating the primitive.
    # ruleid: vba-native-api-declaration
    th = CreateThread(0, 0, AddressOf GetTickCount, 0, 0, tid)

    MsgBox "Allocated memory at: 0x" & Hex$(addr) & vbCrLf & _
           "Thread handle: 0x" & Hex$(th) & vbCrLf & _
           "Thread ID: " & tid, vbInformation
End Sub
