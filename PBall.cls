VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private VB As Direct3DVertexBuffer8
Private Position As D3DVECTOR
Private VertexSizeInBytes As Long
Private Texture As Direct3DTexture8
Private Vertices(1 To 30) As PONGVERTEX
Private MVB As Direct3DVertexBuffer8

Public Function GetPX() As Single
    GetPX = Position.X
End Function

Public Function GetPZ() As Single
    GetPZ = Position.Z
End Function

Public Function MoveBall(X As Single, Z As Single)
    For k = 1 To 30
        With Vertices(k).postion
            .X = .X + X
            .Z = .Z + Z
        End With
    Next k
    Position.X = Position.X + X
    Position.Z = Position.Z + Z
    ReloadVB
End Function

Public Function Render()
    D3DDevice.SetStreamSource 0, MVB, VertexSizeInBytes
    D3DDevice.SetVertexShader D3DFVF_PONGVERTEX
    D3DDevice.SetTexture 0, Texture
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 10
End Function

Public Function Init()
    VertexSizeInBytes = Len(Vertices(1))
    'Top
    Vertices(1) = MakeVertex(-BallRadius, BallRadius, BallRadius, 0, 1)
    Vertices(2) = MakeVertex(-BallRadius, BallRadius, -BallRadius, 0, 0)
    Vertices(3) = MakeVertex(BallRadius, BallRadius, BallRadius, 1, 1)
    Vertices(4) = MakeVertex(BallRadius, BallRadius, BallRadius, 1, 1)
    Vertices(5) = MakeVertex(-BallRadius, BallRadius, -BallRadius, 0, 0)
    Vertices(6) = MakeVertex(BallRadius, BallRadius, -BallRadius, 1, 0)
    'Back Side
    Vertices(7) = MakeVertex(-BallRadius, 0, BallRadius, 0, 0)
    Vertices(8) = MakeVertex(BallRadius, 0, BallRadius, 1, 0)
    Vertices(9) = MakeVertex(BallRadius, BallRadius, BallRadius, 1, 1)
    Vertices(10) = MakeVertex(-BallRadius, 0, BallRadius, 0, 0)
    Vertices(11) = MakeVertex(BallRadius, BallRadius, BallRadius, 1, 1)
    Vertices(12) = MakeVertex(-BallRadius, BallRadius, BallRadius, 0, 1)
    'Front Side
    Vertices(13) = MakeVertex(-BallRadius, 0, -BallRadius, 0, 0)
    Vertices(14) = MakeVertex(BallRadius, 0, -BallRadius, 1, 0)
    Vertices(15) = MakeVertex(BallRadius, BallRadius, -BallRadius, 1, 1)
    Vertices(16) = MakeVertex(-BallRadius, 0, -BallRadius, 0, 0)
    Vertices(17) = MakeVertex(BallRadius, BallRadius, -BallRadius, 1, 1)
    Vertices(18) = MakeVertex(-BallRadius, BallRadius, -BallRadius, 0, 1)
    'Left Side
    Vertices(19) = MakeVertex(-BallRadius, 0, BallRadius, 1, 0)
    Vertices(20) = MakeVertex(-BallRadius, 0, -BallRadius, 0, 0)
    Vertices(21) = MakeVertex(-BallRadius, BallRadius, BallRadius, 1, 1)
    Vertices(22) = MakeVertex(-BallRadius, 0, -BallRadius, 0, 0)
    Vertices(23) = MakeVertex(-BallRadius, BallRadius, -BallRadius, 0, 1)
    Vertices(24) = MakeVertex(-BallRadius, BallRadius, BallRadius, 1, 1)
    'Right Side
    Vertices(25) = MakeVertex(BallRadius, 0, BallRadius, 1, 0)
    Vertices(26) = MakeVertex(BallRadius, 0, -BallRadius, 0, 0)
    Vertices(27) = MakeVertex(BallRadius, BallRadius, BallRadius, 1, 1)
    Vertices(28) = MakeVertex(BallRadius, 0, -BallRadius, 0, 0)
    Vertices(29) = MakeVertex(BallRadius, BallRadius, -BallRadius, 0, 1)
    Vertices(30) = MakeVertex(BallRadius, BallRadius, BallRadius, 1, 1)
    
    Set Texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\ball.bmp")
    
    For k = 1 To 30
        Vertices(k).color = RGB(200, 200, 200)
    Next k
    
    Set MVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 30, 0, D3DFVF_PONGVERTEX, D3DPOOL_DEFAULT)
    ReloadVB
End Function

Private Function ReloadVB()
    D3DVertexBuffer8SetData MVB, 0, VertexSizeInBytes * 30, 0, Vertices(1)
End Function
