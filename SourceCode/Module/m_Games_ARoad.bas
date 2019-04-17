Attribute VB_Name = "m_Games_ARoad"
'**************************************************************************
'Date: 2019/02/01
'Describe:
'Author:  Chenlyu
'E-mail: plarn@foxmail.com
'**************************************************************************

'====================================================================Function description====================================================

' This is about automatic routing algorithms. I have improved some of the standard automatic routing algorithms.

'====================================================================================================================================

 


Option Explicit

Public Length As Integer

Private Type nude_type
    x As Integer
    y As Integer
    Father As Integer
    D1 As Integer
    'D2 As Integer
    '**********
    D2 As Single
    '**********
    Next As Integer
    ID As Integer
End Type

 

Type Closed_map
    NuDenum As Integer
    Mapval As Integer
End Type

'Public Map() As Integer
Private Opened As Integer


'=============================================================
'Describe:Find a path based on map information
'Author:  Chenlyu
'Parameter:
'=============================================================
Public Function FindPath(x As Integer, y As Integer, Dx As Integer, Dy As Integer, ByRef M As Integer, ByRef TT() As m_Position, ByRef Desc_Passed_Enable As Boolean) As Boolean
'This is the main routing procedure.

'The default map name of the program is map, so when my game calls this module, it also calls the map a two-dimensional array of maps ().

'x, y denotes the starting point coordinates, dx, dy denotes the target coordinates, and TT () the coordinates of the path to be taken m denotes the pointer to the array TT () to store the next path not to pass.

'Data of m and TT are automatically generated by this function

Dim Nude(5000) As nude_type
Dim Tmpe(-1 To 500, -1 To 500) As Closed_map
Dim ISt As Integer
Dim X1 As Integer, Y1 As Integer
Dim QQ As Integer
    X1 = x
    Y1 = y
    Nude(0).x = x
    Nude(0).y = y
    Nude(0).Father = -1
    Nude(0).D1 = 0
        
 
'***************************
 'Dim tmpx As Integer, tmpy As Integer
 'tmpx = Abs(x1 - dx): tmpy = Abs(y1 - dy)
 'Nude(0).D2 = IIf(tmpx > tmpy, tmpx, tmpy)
     Nude(0).D2 = Sqr((X1 - Dx) ^ 2 + (Y1 - Dy) ^ 2)
'***********************


    Nude(0).Next = -1
    Nude(0).ID = 0: Opened = 0
    Tmpe(X1, Y1).NuDenum = 0
    Tmpe(X1, Y1).Mapval = 1
    
Dim Maxnum As Integer
Dim MaxCounts As Integer '�������������Χ,�ӿ������ٶ�
Maxnum = 0 '������ɲ��ܸ�
'Select Case Nude(0).d2
'*************
    Select Case Fix(Nude(0).D2)
        '*************
        Case 0 To 5
        MaxCounts = 150
        Case 6 To 10
        MaxCounts = 300
        Case 11 To 20
        MaxCounts = 600
        Case 21 To 30
        MaxCounts = 1000
        Case 31 To 40
        MaxCounts = 2000
        Case Is > 40
        MaxCounts = 5000
    End Select '��� Ŀ�����ڱ���Χ�У����ܵ���)����С������Χ��Լ�ٶ�
    Do
        QQ = Getopenednude(Nude())
        ISt = FindPath_Sub1(Nude(QQ).x, Nude(QQ).y, Dx, Dy, Tmpe(), Nude(), Maxnum, Desc_Passed_Enable)
        If ISt > 0 Then GoTo FINDs '�ҵ�Ŀ��
        If Maxnum >= MaxCounts Then
            Exit Do
        End If
    Loop Until Opened = -1
    If Maxnum = 0 Then
        FindPath = False: Exit Function
    End If
   
    Dim aa As Integer, Lengh As Single
    Dim nn As Integer, mm As Integer
    Dim iii As Integer
    Dim i As Integer
    aa = Maxnum
    Lengh = Nude(1).D2
    iii = 1
    nn = Nude(1).x: mm = Nude(1).y
      
    For i = 1 To aa
        If Length > Nude(i).D2 Then
        Length = Nude(i).D2
        nn = Nude(i).x: mm = Nude(i).y
        iii = i
        End If
    
    Next
        Dx = nn: Dy = mm
FINDs:
Dim l As Integer
aa = iii
l = 0
If ISt > 0 Then aa = ISt
Do
    TT(l).x = Nude(aa).x
    TT(l).y = Nude(aa).y
    aa = Nude(aa).Father
    l = l + 1
Loop Until aa = -1
M = l - 2
FindPath = True

End Function

'=============================================================
'Describe:Find a path based on map information
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Function FindPath_Sub1(ByRef x As Integer, ByRef y As Integer, ByRef Dx As Integer, ByRef Dy As Integer, ByRef tmp() As Closed_map, ByRef Nude() As nude_type, ByRef n As Integer, ByRef Desc_Passed_Enable As Boolean) As Integer
'This is a path-finding sub-module, which generates eight-directional sub-contacts based on the parent contacts and inserts them into the open table.
Dim i As Integer, j As Integer, Fatnum As Integer, Mme As Integer, M As Integer
Dim aaa As Integer, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
Dim t As Integer
t = 0
For i = -1 To 1
    For j = -1 To 1
        If (i <> 0 Or j <> 0) And Testmap(x, y, i, j) Then  'Testmap () is used to record generated map nodes
                                                            ' If the node to be generated is already in the test map, compare the location of their parent d1. If the D1 of the target to be generated is smaller than the value of the parent D1 of the parent node, modify it to the target node.
                                                            'D1 denotes the distance from the starting point
            If tmp(x + i, y + j).Mapval = 1 Then            'Mapval = 1 represents the node to be generated in the testmap () table
                Mme = tmp(x + i, y + j).NuDenum
                aaa = tmp(x, y).NuDenum
                Fatnum = Nude(aaa).Father
                If Fatnum = -1 Then 'It's already the starting point
                
                Else
                    If Nude(Fatnum).D1 > Nude(Mme).D1 Then Nude(aaa).Father = Mme
                End If
            End If
            
            If tmp(x + i, y + j).Mapval = 0 Then 'If the node to be generated is not in the testmap () table, the node is generated
                M = tmp(x, y).NuDenum
                n = n + 1
                X1 = x + i: Y1 = y + j
                tmp(X1, Y1).Mapval = 1
                tmp(X1, Y1).NuDenum = n
                Nude(n).x = X1: Nude(n).y = Y1
                Nude(n).Father = M
                Nude(n).D1 = Nude(M).D1 + 1
           
                '***************************
                 'Dim tmpx As Integer, tmpy As Integer
                 'tmpx = Abs(x1 - dx): tmpy = Abs(y1 - dy)
                 'Nude(0).D2 = IIf(tmpx > tmpy, tmpx, tmpy)
                     Nude(n).D2 = Sqr((X1 - Dx) ^ 2 + (Y1 - Dy) ^ 2)
                '***********************
    
             
                Nude(n).ID = n
                If Nude(n).D2 = 0 Then 'Find the target
                    FindPath_Sub1 = n
                    Exit Function
                End If
                Call InstOPened(Nude(n), Nude())
            End If
        End If
        '*********Used only for monsters looking for players
        If Desc_Passed_Enable = False Then 'testmap() Map nodes used to record generated maps
            If i <> 0 Or j <> 0 Then
                If x + i = Dx And y + j = Dy Then 'Find the target
                    Dx = x:  Dy = y
                    FindPath_Sub1 = tmp(x, y).NuDenum
                    Exit Function
                End If
            End If
        End If
        '*************
    Next
Next
   
End Function

Private Sub InstOPened(Mnud As nude_type, Nude() As nude_type)
Dim Temp2 As Integer
'Dim f As Integer
'************
Dim f As Single
'**********
Dim dd As Integer
If Opened = -1 Then
    Mnud.Next = -1
    Opened = Mnud.ID
    Exit Sub
End If
f = Mnud.D2
Temp2 = Opened
Do
    If f < Nude(Temp2).D2 Then GoTo K1
    dd = Temp2
    Temp2 = Nude(Temp2).Next
Loop Until Temp2 = -1
Mnud.Next = -1
Nude(dd).Next = Mnud.ID
Exit Sub
K1:

Mnud.Next = Temp2
If Opened <> Temp2 Then
    Nude(dd).Next = Mnud.ID
Else
    Opened = Mnud.ID
End If
End Sub
Private Function Getopenednude(Nude() As nude_type) As Integer
Dim tmp3 As Integer
If Opened = -1 Then
    ' error
Else
    tmp3 = Opened
    Opened = Nude(tmp3).Next
    Getopenednude = tmp3
End If
End Function

'=============================================================
'Describe:Detecting the state of a map
'Author:  Chenlyu
'Parameter:
'=============================================================
Private Function Testmap(ByVal x As Integer, ByVal y As Integer, i As Integer, j As Integer)
'  Check whether the map is accessible
On Error GoTo Errors 'You can't go beyond the border.
'You can add other restrictions here, such as no other ants on a walking map.
If i * j = 0 Then  'Quartet Judgment
    If this_Graphic.Map.Block(x + i, y + j) = 0 Then
        Testmap = True
        Exit Function
    End If
Else   'Diagonal Quadrangle Judgment
    If this_Graphic.Map.Block(x + i, y + j) = 0 And this_Graphic.Map.Block(x + i, y) = 0 And this_Graphic.Map.Block(x, y + j) = 0 Then
        Testmap = True
        Exit Function
    End If
End If
Errors:
    Testmap = False
End Function


