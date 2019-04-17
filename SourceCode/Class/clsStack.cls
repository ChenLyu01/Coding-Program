VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'Date: 2019/02/01
'Describe:This is a function about map generation.
'Author:  Chenlyu and Cuiwei
'E-mail:  plarn@foxmail.com
'**************************************************************************


'====================================================================Function description====================================================

' Module declaration in the code section

'====================================================================================================================================


Option Explicit

Private s As New Collection

'=============================================================
'Describe: Processing function of stack
'Author:   Chen Lyu
'Parameter:
'=============================================================
Public Sub Push(ByVal Item As Variant)
    If s.Count = 0 Then
        s.Add Item
    Else
        s.Add Item, , 1
    End If
End Sub

Public Function Pop()
    Dim t
    t = s.Item(1)
    s.Remove 1
    Pop = t
End Function

Public Function Clear()
    While s.Count > 0
        s.Remove 1
    Wend
End Function

Public Property Get Count() As Variant
    Count = s.Count
End Property

Public Property Let Count(ByVal vNewValue As Variant)
    
End Property

Public Property Get Item(ByVal i As Integer) As Variant
    ''Start from 0, not from 1, pay attention.~~~
    Item = s.Item(i + 1)
End Property

Public Property Let Item(ByVal i As Integer, ByVal vNewValue As Variant)
End Property
