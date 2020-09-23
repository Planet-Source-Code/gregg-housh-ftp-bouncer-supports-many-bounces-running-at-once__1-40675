VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' a vector is just an array, that auto resizes itself...

Public Enum EErrorVector
    eeBaseVector = 13270    ' CVector
End Enum

Private av() As String
Private iLast As Long
Private cChunk As Long

Property Get Chunk() As Long
    Chunk = cChunk
End Property

Property Let Chunk(cChunkA As Long)
    If cChunkA > 0 Then
        cChunk = IIf(cChunkA < 100, cChunkA, 100)
    End If
End Property

' Item is the default property
Property Get item(ByVal i As Long) As String
Attribute item.VB_UserMemId = 0
    If i > 0 Then
        ' If index is out-of-range, return default (Empty)
        On Error Resume Next
        If IsObject(av(i)) Then
            'Set item = av(i)
        Else
            item = av(i)
        End If
    Else
        item = Empty
    End If
End Property

Property Let item(ByVal i As Long, ByVal vItemA As String)
    If i > 0 Then
        On Error GoTo FailLetItem
        av(i) = vItemA
        If i > iLast Then
            iLast = i
        End If
    End If
    Exit Property
FailLetItem:
    If i > UBound(av) Then
        ReDim Preserve av(1 To i + cChunk) As String
        Resume              ' Try again
    End If
    Err.Raise Err.Number, Err.Source, Err.Description   ' Other VB error for client
End Property

Property Get Last() As Long
    Last = iLast
End Property

Property Let Last(iLastA As Long)
    If iLastA > 0 Then
        ReDim Preserve av(1 To iLastA) As String
        iLast = iLastA
    End If
End Property

' Friend properties to make data structure accessible to walker
Friend Property Get Vector(ByVal i As Long) As String
If i > 0 And i <= iLast Then
    If IsObject(av(i)) Then
        'Set Vector = av(i)
    Else
        Vector = av(i)
    End If
Else
    Vector = Empty
End If
End Property

Private Sub Class_Initialize()
    cChunk = 30     ' Default size can be overridden
    ReDim Preserve av(1 To cChunk) As String
    iLast = 1
End Sub

