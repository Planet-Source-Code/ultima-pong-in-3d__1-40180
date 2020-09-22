Attribute VB_Name = "HelperFunctions"
Public Function MakeVertex(X As Single, Y As Single, Z As Single, U As Single, V As Single) As PONGVERTEX
    With MakeVertex
        .postion.X = X
        .postion.Y = Y
        .postion.Z = Z
        .tu = U
        .tv = V
    End With
End Function

Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    With MakeVector
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function
