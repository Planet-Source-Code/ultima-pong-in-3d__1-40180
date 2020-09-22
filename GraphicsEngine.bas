Attribute VB_Name = "GraphicsEngine"
Public DX As New DirectX8
Public D3DX As New D3DX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public VB As Direct3DVertexBuffer8
Public CurrTriCount As Integer

Public Type PONGVERTEX
    postion As D3DVECTOR
    color As Long
    tu As Single
    tv As Single
End Type

Public Const D3DFVF_PONGVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Const Pi As Single = 3.14159275180032

Public Function Init(HWND As Long) As Boolean

    Set D3D = DX.Direct3DCreate()
    If D3D Is Nothing Then Exit Function
    
    Dim DispMode As D3DDISPLAYMODE
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    Dim D3DPP As D3DPRESENT_PARAMETERS
    
    With D3DPP
        .Windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        .BackBufferFormat = DispMode.Format
        .hDeviceWindow = HWND
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
    End With
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, HWND, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
    If D3DDevice Is Nothing Then Exit Function
    
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    D3DDevice.SetRenderState D3DRS_LIGHTING, 0

    InitDI
    SetupGame
    SetupMatrices
    
    Init = True
End Function

Public Function SetupMatrices()
    Dim matView As D3DMATRIX
    
    D3DXMatrixLookAtLH matView, EyePosition, EyeLookAt, MakeVector(0, 1, 0)
    
    D3DDevice.SetTransform D3DTS_VIEW, matView

    Dim matProj As D3DMATRIX
    
    D3DXMatrixPerspectiveFovLH matProj, Pi / 3, 1, 1, 1000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Function

Public Function RenderAll()
    UpdateBall
    
    If P2Con = False Then AI
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    
    D3DDevice.BeginScene
        
        D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
        D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
        D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
        
        Table.Render
        Player1.Render
        Player2.Render
        Ball.Render
    
    D3DDevice.EndScene
    
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function

Public Function KillApp()
    Set DVB = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
    DIDevice.Unacquire
    Finish = True
    End
End Function
