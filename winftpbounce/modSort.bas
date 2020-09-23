Attribute VB_Name = "modSort"
Option Explicit

Private Sorted As Variant

Public Function QuickSort(List As Variant) As Variant
    Dim Leftpoint
    Dim Rightpoint As Double
    Leftpoint = LBound(List)
    Rightpoint = UBound(List)
    Sorted = List
    Call Quick(Leftpoint, Rightpoint)
    QuickSort = Sorted
End Function

Private Sub Quick(Leftpoint, Rightpoint)
    Dim Passedright
    Dim Passedleft As Double
    Dim Ref As Boolean
    Dim Temp As Variant
    Passedleft = Leftpoint
    Passedright = Rightpoint
    Ref = False
    
    Do Until Leftpoint = Rightpoint
        
        If Sorted(Rightpoint) < Sorted(Leftpoint) Then
            Temp = Sorted(Rightpoint)
            Sorted(Rightpoint) = Sorted(Leftpoint)
            Sorted(Leftpoint) = Temp
            
            If Ref = False Then
                Ref = True
            Else
                Ref = False
            End If
        End If
        
        If Ref = False Then
            Rightpoint = Rightpoint - 1
        Else
            Leftpoint = Leftpoint + 1
        End If
    Loop
    
    If Leftpoint - Passedleft > 1 Then
        Call Quick(Passedleft, Leftpoint - 1)
    End If
    
    If Passedright - Rightpoint > 1 Then
        Call Quick(Leftpoint + 1, Passedright)
    End If
End Sub

