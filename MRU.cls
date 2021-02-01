VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

' This class is for saving in Windows registry list of most recently used items. Items could be:
' last used files (paths), last used words, last used pictures or anything else...

' Top registry key is: "HKEY_CURRENT_USER\.DEFAULT\SOFTWARE\VB and VBA Program Settings"
' under this registry key will be created subkey called according to Parent

Public MaxNum% ' max number of MRU's
Public Parent$ ' registry parent key (i.e. name of your software company)
Public Child$ ' registry value (i.e. name of your application)
Public Prefix$ ' registry prefix (i.e. "File", "Word", "Picture", etc.

' class initialization
Private Sub Class_Initialize()
    MaxNum = 10
    Parent = "MRU"
    Child = "MRU"
    Prefix = "MRU"
End Sub

' returns the number of currently saved MRU's in registry
Public Property Get Num() As Integer
    Num = Val(GetSetting(Parent, Child, Prefix & "Num"))
End Property

' returns last MRU
Public Property Get Last() As String
    Last = GetSetting(Parent, Child, Prefix & "1")
End Property

' sets last MRU (replaces previous last)
Public Property Let Last(varMRU As String)
    Call SaveSetting(Parent, Child, Prefix & "1", varMRU)
End Property

' returns array fullfilled with MRU's
' array will not preserve previous values!
Public Sub GetArray(ByRef varArray() As String)
Dim i%, Num%
    Num = Val(GetSetting(Parent, Child, Prefix & "Num", 0))
    If Num > 0 Then
        ReDim varArray(Num)
        For i = 1 To Num
            varArray(i) = GetSetting(Parent, Child, Prefix & i)
        Next i
    End If
End Sub

' saves item in list of MRU's
Public Sub Save(ByVal varMRU As String)
Dim i%, Num%, p%, MRU$()
ReDim MRU(MaxNum)
    Num = Val(GetSetting(Parent, Child, Prefix & "Num", 0))
    If Num > 0 Then ' used before
        '
        For i = 1 To Num
            MRU(i) = GetSetting(Parent, Child, Prefix & i)
            If MRU(i) = varMRU Then p = i
        Next i
        '
        If p = 0 Then
            If Num < MaxNum Then
                For i = Num To 1 Step -1
                    MRU(i + 1) = MRU(i)
                Next i
                Num = Num + 1
            Else
                For i = Num - 1 To 1 Step -1
                    MRU(i + 1) = MRU(i)
                Next i
            End If
        ElseIf p > 1 Then
            For i = p - 1 To 1 Step -1
                MRU(i + 1) = MRU(i)
            Next i
        End If
        MRU(1) = varMRU
        '
        Call SaveSetting(Parent, Child, Prefix & "Num", Num)
        For i = 1 To Num
            Call SaveSetting(Parent, Child, Prefix & i, MRU(i))
        Next i
        '
    Else 'never used before
        Call SaveSetting(Parent, Child, Prefix & "Num", 1)
        Call SaveSetting(Parent, Child, Prefix & "1", varMRU)
    End If
End Sub

' clears the child key from registry
Public Sub Clear()
On Error Resume Next
    Call DeleteSetting(Parent, Child)
On Error GoTo 0
End Sub

' clears the parent key from registry
Public Sub ClearParent()
On Error Resume Next
    Call DeleteSetting(Parent)
On Error GoTo 0
End Sub
