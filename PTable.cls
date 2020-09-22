VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private SideTexture As Direct3DTexture8
Private FloorTexture As Direct3DTexture8
Private VertexSizeInBytes As Long
Private Vertices(1 To 30) As PONGVERTEX
Private MVB As Direct3DVertexBuffer8

Private Const FHei = 25

Public Function Render()
    D3DDevice.SetStreamSource 0, MVB, VertexSizeInBytes
    D3DDevice.SetVertexShader D3DFVF_PONGVERTEX
    D3DDevice.SetTexture 0, FloorTexture
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    D3DDevice.SetTexture 0, SideTexture
    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 6, 8
End Function

Public Function Init()
    VertexSizeInBytes = Len(Vertices(1))
    'Floor
    Vertices(1) = MakeVertex(-100, 0, FDepth, 0, 1)
    Vertices(2) = MakeVertex(-100, 0, -FDepth, 0, 0)
    Vertices(3) = MakeVertex(100, 0, FDepth, 1, 1)
    Vertices(4) = MakeVertex(100, 0, FDepth, 1, 1)
    Vertices(5) = MakeVertex(-100, 0, -FDepth, 0, 0)
    Vertices(6) = MakeVertex(100, 0, -FDepth, 1, 0)
    'Back Side
    Vertices(7) = MakeVertex(-100, 0, FDepth, 0, 0)
    Vertices(8) = MakeVertex(100, 0, FDepth, 1, 0)
    Vertices(9) = MakeVertex(100, FHei, FDepth, 1, 1)
    Vertices(10) = MakeVertex(-100, 0, FDepth, 0, 0)
    Vertices(11) = MakeVertex(100, FHei, FDepth, 1, 1)
    Vertices(12) = MakeVertex(-100, FHei, FDepth, 0, 1)
    'Front Side
    Vertices(13) = MakeVertex(-100, 0, -FDepth, 0, 0)
    Vertices(14) = MakeVertex(100, 0, -FDepth, 1, 0)
    Vertices(15) = MakeVertex(100, FHei, -FDepth, 1, 1)
    Vertices(16) = MakeVertex(-100, 0, -FDepth, 0, 0)
    Vertices(17) = MakeVertex(100, FHei, -FDepth, 1, 1)
    Vertices(18) = MakeVertex(-100, FHei, -FDepth, 0, 1)
    'Left Side
    Vertices(19) = MakeVertex(-100, 0, FDepth, 1, 0)
    Vertices(20) = MakeVertex(-100, 0, -FDepth, 0, 0)
    Vertices(21) = MakeVertex(-100, FHei, FDepth, 1, 1)
    Vertices(22) = MakeVertex(-100, 0, -FDepth, 0, 0)
    Vertices(23) = MakeVertex(-100, FHei, -FDepth, 0, 1)
    Vertices(24) = MakeVertex(-100, FHei, FDepth, 1, 1)
    'Right Side
    Vertices(25) = MakeVertex(100, 0, FDepth, 1, 0)
    Vertices(26) = MakeVertex(100, 0, -FDepth, 0, 0)
    Vertices(27) = MakeVertex(100, FHei, FDepth, 1, 1)
    Vertices(28) = MakeVertex(100, 0, -FDepth, 0, 0)
    Vertices(29) = MakeVertex(100, FHei, -FDepth, 0, 1)
    Vertices(30) = MakeVertex(100, FHei, FDepth, 1, 1)

    Set FloorTexture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Floor.bmp")
    Set SideTexture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Side.bmp")

    For k = 1 To 30
        Vertices(k).color = RGB(200, 200, 200)
    Next k
    
    Set MVB = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 30, 0, D3DFVF_PONGVERTEX, D3DPOOL_DEFAULT)
    ReloadVB
End Function

Private Function ReloadVB()
    D3DVertexBuffer8SetData MVB, 0, VertexSizeInBytes * 30, 0, Vertices(1)
End Function
