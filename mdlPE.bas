Attribute VB_Name = "mdlPE"
Option Explicit
Public Type IMAGE_DOS_HEADER
    Magic As Integer
    cblp As Integer
    cp As Integer
    crlc As Integer
    cparhdr As Integer
    minalloc As Integer
    maxalloc As Integer
    Ss As Integer
    sp As Integer
    csum As Integer
    ip As Integer
    cs As Integer
    lfarlc As Integer
    ovno As Integer
    res(3) As Integer
    oemid As Integer
    oeminfo As Integer
    res2(9) As Integer
    lfanew As Long
End Type
Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOtionalHeader As Integer
    Characteristics As Integer
End Type
Public Type IMAGE_DATA_DIRECTORY
    DataRVA As Long
    DataSize As Long
End Type
Public Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkVer As Byte
    MinorLinkVer As Byte
    CodeSize As Long
    InitDataSize As Long
    unInitDataSize As Long
    EntryPoint As Long
    CodeBase As Long
    DataBase As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOSVer As Integer
    MinorOSVer As Integer
    MajorImageVer As Integer
    MinorImageVer As Integer
    MajorSSVer As Integer
    MinorSSVer As Integer
    Win32Ver As Long
    ImageSize As Long
    HeaderSize As Long
    Checksum As Long
    Subsystem As Integer
    DLLChars As Integer
    StackRes As Long
    StackCommit As Long
    HeapReserve As Long
    HeapCommit As Long
    LoaderFlags As Long
    RVAsAndSizes As Long
    DataEntries(15) As IMAGE_DATA_DIRECTORY
End Type
Public Type IMAGE_SECTION_HEADER
    SectionName(7) As Byte
    Address As Long
    VirtualAddress As Long
    SizeOfData As Long
    PData As Long
    PReloc As Long
    PLineNums As Long
    RelocCount As Integer
    LineCount As Integer
    Characteristics As Long
End Type
Type IMAGE_RESOURCE_DIR
    Characteristics As Long
    TimeStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    NamedEntries As Integer
    IDEntries As Integer
End Type
Type RESOURCE_DIR_ENTRY
    Name As Long
    Offset As Long
End Type
Type RESOURCE_DATA_ENTRY
    Offset As Long
    Size As Long
    CodePage As Long
    Reserved As Long
End Type
Public Type IconDescriptor
    ID As Long
    Offset As Long
    Size As Long
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private SectionAlignment As Long
Private FileAlignment As Long
Private ResSectionRVA As Long
Private ResSectionOffset As Long

Public Function Valid_PE(hFile As Long) As Boolean
On Local Error GoTo ErrHandler
Dim Buffer(12) As Byte, lngBytesRead As Long, tDosHeader As IMAGE_DOS_HEADER
If (hFile > 0) Then
    ReadFile hFile, tDosHeader, ByVal Len(tDosHeader), lngBytesRead, ByVal 0&
    CopyMemory Buffer(0), tDosHeader.Magic, 2
    If (Chr(Buffer(0)) & Chr(Buffer(1)) = "MZ") Then
        SetFilePointer hFile, tDosHeader.lfanew, 0, 0
        ReadFile hFile, Buffer(0), 4, lngBytesRead, ByVal 0&
        If (Chr(Buffer(0)) = "P") And (Chr(Buffer(1)) = "E") And (Buffer(2) = 0) And (Buffer(3) = 0) Then
            Valid_PE = True
            Exit Function
        End If
    End If
End If
Valid_PE = False
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function Valid_PE(hFile As Long) As Boolean"
End Function

Public Function GetResTreeOffset(lFile As Long) As Long
On Error GoTo ErrHandler
Dim lDosHeader As IMAGE_DOS_HEADER, lImageHeader As IMAGE_FILE_HEADER, lOptional As IMAGE_OPTIONAL_HEADER, lSections() As IMAGE_SECTION_HEADER, l As Long, i As Integer, b As Boolean
b = False
If (lFile > 0) Then
    SetFilePointer lFile, 0, 0, 0
    ReadFile lFile, lDosHeader, Len(lDosHeader), l, ByVal 0&
    SetFilePointer lFile, ByVal lDosHeader.lfanew + 4, 0, 0
    ReadFile lFile, lImageHeader, Len(lImageHeader), l, ByVal 0&
    ReadFile lFile, lOptional, Len(lOptional), l, ByVal 0&
    ReDim lSections(lImageHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER
    ReadFile lFile, lSections(0), Len(lSections(0)) * lImageHeader.NumberOfSections, l, ByVal 0&
    If (lOptional.DataEntries(2).DataSize) Then
        SectionAlignment = lOptional.SectionAlignment
        FileAlignment = lOptional.FileAlignment
        For i = 0 To UBound(lSections)
            If (lSections(i).VirtualAddress <= lOptional.DataEntries(2).DataRVA) And ((lSections(i).VirtualAddress + lSections(i).SizeOfData) > lOptional.DataEntries(2).DataRVA) Then
                b = True
                ResSectionRVA = lSections(i).VirtualAddress
                ResSectionOffset = lSections(i).PData
                GetResTreeOffset = lSections(i).PData + (lOptional.DataEntries(2).DataRVA - lSections(i).VirtualAddress)
                Exit For
            End If
        Next i
        If Not b Then
            GetResTreeOffset = -1
        End If
    Else
        GetResTreeOffset = -1
    End If
Else
    GetResTreeOffset = -1
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function GetResTreeOffset(lFile As Long) As Long"
End Function

Public Function GetIconOffsets(hFile As Long, TreeOffset As Long, Icons() As IconDescriptor) As Long
On Local Error GoTo ErrHandler
Dim lImageResDir As IMAGE_RESOURCE_DIR, L1Entries() As RESOURCE_DIR_ENTRY, L2Root() As IMAGE_RESOURCE_DIR, L2Entries() As RESOURCE_DIR_ENTRY, L3Root() As IMAGE_RESOURCE_DIR, L3Entries() As RESOURCE_DIR_ENTRY, DataEntries() As RESOURCE_DATA_ENTRY, DIB As DIB_HEADER, iLvl1 As Integer, iLvl2 As Integer, iLvl3 As Integer, Cursor As Long, BytesRead As Long, Count As Integer
If (hFile > 0) Then
    Count = 0
    SetFilePointer hFile, ByVal TreeOffset, 0, 0
    ReadFile hFile, lImageResDir, Len(lImageResDir), BytesRead, ByVal 0
    ReDim L2Root(lImageResDir.NamedEntries + lImageResDir.IDEntries) As IMAGE_RESOURCE_DIR
    ReDim L1Entries(lImageResDir.NamedEntries + lImageResDir.IDEntries) As RESOURCE_DIR_ENTRY
    For iLvl1 = 1 To (lImageResDir.NamedEntries + lImageResDir.IDEntries)
        SetFilePointer hFile, TreeOffset + 8 + (iLvl1 * 8), 0, 0
        ReadFile hFile, L1Entries(iLvl1), 8, BytesRead, ByVal 0&
        If L1Entries(iLvl1).Name = 3 Then
            CopyMemory Cursor, L1Entries(iLvl1).Offset, 3
            Cursor = Cursor + TreeOffset
            SetFilePointer hFile, ByVal Cursor, 0, 0
            ReadFile hFile, L2Root(iLvl1), 16, BytesRead, ByVal 0&
            ReDim L3Root(L2Root(iLvl1).NamedEntries + L2Root(iLvl1).IDEntries) As IMAGE_RESOURCE_DIR
            ReDim L2Entries(L2Root(iLvl1).IDEntries + L2Root(iLvl1).NamedEntries) As RESOURCE_DIR_ENTRY
            For iLvl2 = 1 To (L2Root(iLvl1).IDEntries + L2Root(iLvl1).NamedEntries)
                CopyMemory Cursor, L1Entries(iLvl1).Offset, 3
                Cursor = Cursor + TreeOffset
                SetFilePointer hFile, Cursor + 8 + (iLvl2 * 8), 0, 0
                ReadFile hFile, L2Entries(iLvl2), 8, BytesRead, ByVal 0&
                CopyMemory Cursor, L2Entries(iLvl2).Offset, 3
                Cursor = Cursor + TreeOffset
                SetFilePointer hFile, ByVal Cursor, 0, 0
                ReadFile hFile, L3Root(iLvl2), 16, BytesRead, ByVal 0&
                ReDim L3Entries(L3Root(iLvl2).NamedEntries + L3Root(iLvl2).IDEntries) As RESOURCE_DIR_ENTRY
                ReDim DataEntries(L3Root(iLvl2).NamedEntries + L3Root(iLvl2).IDEntries) As RESOURCE_DATA_ENTRY
                For iLvl3 = 1 To (L3Root(iLvl2).NamedEntries + L3Root(iLvl2).IDEntries)
                    CopyMemory Cursor, L2Entries(iLvl2).Offset, 3
                    Cursor = Cursor + TreeOffset
                    SetFilePointer hFile, (Cursor + 8 + (iLvl3 * 8)), 0, 0
                    ReadFile hFile, L3Entries(iLvl3), 8, BytesRead, ByVal 0&
                    SetFilePointer hFile, TreeOffset + (L3Entries(iLvl3).Offset), 0, 0
                    ReadFile hFile, DataEntries(iLvl3), 16, BytesRead, ByVal 0&
                    Count = Count + 1
                    ReDim Preserve Icons(Count) As IconDescriptor
                    Icons(Count).Offset = RVAToOffset(DataEntries(iLvl3).Offset)
                    Icons(Count).ID = L2Entries(iLvl2).Name
                    SetFilePointer hFile, Icons(Count).Offset, 0, 0
                    ReadFile hFile, DIB, ByVal Len(DIB), BytesRead, ByVal 0&
                    Icons(Count).Size = DIB.ImageSize + 40
                Next iLvl3
            Next iLvl2
        End If
    Next iLvl1
Else
    Count = 0
End If
GetIconOffsets = Count
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function GetIconOffsets(hFile As Long, TreeOffset As Long, Icons() As IconDescriptor) As Long"
End Function

Public Function HackDirectories(hFile As Long, ResTree As Long, DIBOffset As Long, DIBAttrib As ICON_DIR_ENTRY) As Boolean
On Local Error GoTo ErrHandler
Dim Cursor As Long, Root As IMAGE_RESOURCE_DIR, L1Entries() As RESOURCE_DIR_ENTRY, L2Root() As IMAGE_RESOURCE_DIR, L2Entries() As RESOURCE_DIR_ENTRY, L3Root() As IMAGE_RESOURCE_DIR, L3Entries() As RESOURCE_DIR_ENTRY, DataEntries() As RESOURCE_DATA_ENTRY, IcoDir As ICON_DIR, iLvl1 As Integer, iLvl2 As Integer, iLvl3 As Integer, intC As Integer, BytesRead As Long
If (hFile >= 0) Then
    DIBOffset = OffsetToRVA(DIBOffset)
    SetFilePointer hFile, ByVal ResTree, 0, 0
    ReadFile hFile, Root, Len(Root), BytesRead, ByVal 0&
    ReDim L1Entries(Root.NamedEntries + Root.IDEntries) As RESOURCE_DIR_ENTRY
    ReDim L2Root(Root.NamedEntries + Root.IDEntries) As IMAGE_RESOURCE_DIR
    For iLvl1 = 1 To (Root.NamedEntries + Root.IDEntries)
        SetFilePointer hFile, ResTree + 8 + (iLvl1 * 8), 0, 0
        ReadFile hFile, L1Entries(iLvl1), 8, BytesRead, ByVal 0&
        If L1Entries(iLvl1).Name = &HE Then
            CopyMemory Cursor, L1Entries(iLvl1).Offset, 3
            Cursor = Cursor + ResTree
            SetFilePointer hFile, Cursor, 0, 0
            ReadFile hFile, L2Root(iLvl1), 16, BytesRead, ByVal 0&
            ReDim L2Entries(L2Root(iLvl1).NamedEntries + L2Root(iLvl1).IDEntries) As RESOURCE_DIR_ENTRY
            ReDim L3Root(L2Root(iLvl1).NamedEntries + L2Root(iLvl1).IDEntries) As IMAGE_RESOURCE_DIR
            For iLvl2 = 1 To (L2Root(iLvl1).NamedEntries + L2Root(iLvl1).IDEntries)
                CopyMemory Cursor, L1Entries(iLvl1).Offset, 3
                Cursor = Cursor + ResTree
                SetFilePointer hFile, Cursor + 8 + (iLvl2 * 8), 0, 0
                ReadFile hFile, L2Entries(iLvl2), 8, BytesRead, ByVal 0&
                CopyMemory Cursor, L2Entries(iLvl2).Offset, 3
                Cursor = Cursor + ResTree
                SetFilePointer hFile, Cursor, 0, 0
                ReadFile hFile, L3Root(iLvl2), 16, BytesRead, ByVal 0&
                ReDim L3Entries(L3Root(iLvl2).NamedEntries + L3Root(iLvl2).IDEntries) As RESOURCE_DIR_ENTRY
                For iLvl3 = 1 To (L3Root(iLvl2).NamedEntries + L3Root(iLvl2).IDEntries)
                    CopyMemory Cursor, L2Entries(iLvl2).Offset, 3
                    Cursor = Cursor + ResTree + 8 + (iLvl3 * 8)
                    SetFilePointer hFile, Cursor, 0, 0
                    ReadFile hFile, L3Entries(iLvl3), 8, BytesRead, ByVal 0&
                    CopyMemory Cursor, L3Entries(iLvl3).Offset, 3
                    Cursor = Cursor + ResTree
                    SetFilePointer hFile, Cursor, 0, 0
                    ReDim Preserve DataEntries(iLvl3) As RESOURCE_DATA_ENTRY
                    ReadFile hFile, DataEntries(iLvl3), 16, BytesRead, ByVal 0&
                    Cursor = RVAToOffset(DataEntries(iLvl3).Offset)
                    SetFilePointer hFile, Cursor, 0, 0
                    ReadFile hFile, IcoDir, 6, BytesRead, ByVal 0&
                    For intC = 1 To IcoDir.Count
                    WriteFile hFile, DIBAttrib, Len(DIBAttrib) - 4, BytesRead, ByVal 0&
                    SetFilePointer hFile, 2, 0, 1
                    Next intC
                Next iLvl3
            Next iLvl2
        ElseIf L1Entries(iLvl1).Name = 3 Then
            CopyMemory Cursor, L1Entries(iLvl1).Offset, 3
            Cursor = Cursor + ResTree
            SetFilePointer hFile, ByVal Cursor, 0, 0
            ReadFile hFile, L2Root(iLvl1), 16, BytesRead, ByVal 0&
            ReDim L2Entries(L2Root(iLvl1).NamedEntries + L2Root(iLvl1).IDEntries) As RESOURCE_DIR_ENTRY
            ReDim L3Root(L2Root(iLvl1).NamedEntries + L2Root(iLvl1).IDEntries) As IMAGE_RESOURCE_DIR
            For iLvl2 = 1 To (L2Root(iLvl1).NamedEntries + L2Root(iLvl1).IDEntries)
                CopyMemory Cursor, L1Entries(iLvl1).Offset, 3
                Cursor = Cursor + ResTree
                SetFilePointer hFile, Cursor + 8 + (iLvl2 * 8), 0, 0
                ReadFile hFile, L2Entries(iLvl2), 8, BytesRead, ByVal 0&
                CopyMemory Cursor, L2Entries(iLvl2).Offset, 3
                Cursor = Cursor + ResTree
                SetFilePointer hFile, Cursor, 0, 0
                ReadFile hFile, L3Root(iLvl2), 16, BytesRead, ByVal 0&
                ReDim L3Entries(L3Root(iLvl2).NamedEntries + L3Root(iLvl2).IDEntries) As RESOURCE_DIR_ENTRY
                For iLvl3 = 1 To (L3Root(iLvl2).NamedEntries + L3Root(iLvl2).IDEntries)
                    CopyMemory Cursor, L2Entries(iLvl2).Offset, 3
                    Cursor = Cursor + ResTree + 8 + (iLvl3 * 8)
                    SetFilePointer hFile, Cursor, 0, 0
                    ReadFile hFile, L3Entries(iLvl3), 8, BytesRead, ByVal 0&
                    Cursor = L3Entries(iLvl3).Offset + ResTree
                    SetFilePointer hFile, Cursor, 0, 0
                    WriteFile hFile, DIBOffset, 4, BytesRead, ByVal 0&
                    WriteFile hFile, CLng(DIBAttrib.dwBytesInRes + 40), 4, BytesRead, ByVal 0&
                Next iLvl3
            Next iLvl2
        End If
    Next iLvl1
Else
    HackDirectories = False
    Exit Function
End If
HackDirectories = True
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function HackDirectories(hFile As Long, ResTree As Long, DIBOffset As Long, DIBAttrib As ICON_DIR_ENTRY) As Boolean"
End Function

Private Function RVAToOffset(lRVA As Long) As Long
On Local Error GoTo ErrHandler
Dim l As Long
l = lRVA - ResSectionRVA
If (l >= 0) Then
    RVAToOffset = ResSectionOffset + l
Else
    RVAToOffset = -1
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function HackDirectories(hFile As Long, ResTree As Long, DIBOffset As Long, DIBAttrib As ICON_DIR_ENTRY) As Boolean"
End Function

Private Function OffsetToRVA(lOffset As Long) As Long
On Local Error GoTo ErrHandler
Dim l As Long
l = lOffset - ResSectionOffset
If l >= 0 Then
    OffsetToRVA = ResSectionRVA + l
Else
    OffsetToRVA = -1
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Private Function OffsetToRVA(lOffset As Long) As Long"
End Function

