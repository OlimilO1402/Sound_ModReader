Attribute VB_Name = "MUDTPointer"
Option Explicit '2007_05_17 Zeilen: 96
Public Type TUDTPtr
  pSA        As Long    '4
  cDims      As Integer '2
  fFeatures  As Integer '2
  cbElements As Long    '4
  cLocks     As Long    '4
  pvData     As Long    '4
  cElements  As Long    '4
  lLBound    As Long    '4
End Type              ' 28
Public Type TByteBuffer
  pudt As TUDTPtr
  p() As Byte
End Type
Public Const FADF_RECORD As Integer = &H20
Public Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Long, pSrc As Long, ByVal BLen As Long)
Public Declare Sub RtlZeroMemory Lib "kernel32" (pDst As Long, ByVal BLen As Long)
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (pArr() As Any) As Long

Public Property Let SAPtr(pArr As Long, pSA As Long)
  Call RtlMoveMemory(ByVal pArr, pSA, 4)
End Property
Public Property Get SAPtr(pArr As Long) As Long
  Call RtlMoveMemory(SAPtr, ByVal pArr, 4)
End Property
Public Sub ZeroSAPtr(pArr As Long)
  Call RtlZeroMemory(ByVal pArr, 4)
End Sub

Public Sub InitByteBuffer(aBB As TByteBuffer, Optional dataptr As Long, Optional byteLen As Long)
  With aBB.pudt: .pSA = VarPtr(.cDims): .cDims = 1
    .cbElements = 1: .pvData = dataptr: .cElements = byteLen
    SAPtr(ArrPtr(aBB.p)) = .pSA
  End With
End Sub
Public Sub DeleteByteBuffer(aBB As TByteBuffer)
  Call ZeroSAPtr(ArrPtr(aBB.p))
End Sub

Public Function ByteBufferToString(aBB As TByteBuffer) As String
  With aBB.pudt
    ByteBufferToString = _
        "pSA       : " & CStr(.pSA) & vbCrLf & _
        "cDims     : " & CStr(.cDims) & vbCrLf & _
        "fFeatures : " & CStr(.fFeatures) & vbCrLf & _
        "cbElements: " & CStr(.cbElements) & vbCrLf & _
        "cLocks    : " & CStr(.cLocks) & vbCrLf & _
        "pvData    : " & CStr(.pvData) & vbCrLf & _
        "cElements : " & CStr(.cElements) & vbCrLf & _
        "lLBound   : " & CStr(.lLBound) & vbCrLf & _
        "ArrPtr(p) : " & CStr(ArrPtr(aBB.p)) & vbCrLf & _
        "pSA       : " & CStr(SAPtr(ArrPtr(aBB.p)))
  End With
End Function



