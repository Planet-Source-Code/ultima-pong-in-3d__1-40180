Attribute VB_Name = "GameEngine"
Public Ball As PBall
Public Player1 As PPlayer
Public Player2 As PPlayer
Public Table As PTable

Public Finish As Boolean
Public InGame As Boolean
Public P2Con As Boolean

Public Rot As Single

Public P1Points As Integer
Public P2Points As Integer

Public DI As DirectInput8
Public DIDevice As DirectInputDevice8
Public DIState As DIKEYBOARDSTATE

Public EyePosition As D3DVECTOR
Public EyeLookAt As D3DVECTOR
Public Dir As D3DVECTOR

Public BallSpeed As Single

Public Const PlayerDepth = 5
Public Const PlayerWidth = 20
Public Const PlayerHeight = 5
Public Const FDepth = 150
Public Const BallRadius = 5
Public Const PlayerSpeed = 2

Public Function InitDI()
    Set DI = DX.DirectInputCreate()
    Set DIDevice = DI.CreateDevice("GUID_SysKeyboard")
    
    DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDevice.SetCooperativeLevel FrmMain.HWND, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
    DIDevice.Acquire
End Function

Public Function SetupGame()
    Set Player1 = New PPlayer
    Set Player2 = New PPlayer
    Set Table = New PTable
    Set Ball = New PBall
    
    Player1.Init RGB(255, 0, 0)
    Player1.MovePlayer 0, -120
    Player2.Init RGB(0, 0, 255)
    Player2.MovePlayer 0, 120
    Table.Init
    Ball.Init
    
    Rot = 1
    
    EyePosition = MakeVector(0, 200, -230)
    EyeLookAt = MakeVector(0, 0, 0)
End Function

Public Function StartBall()
    Dir = MakeVector(-BallSpeed, 0, -BallSpeed)
    InGame = True
End Function

Public Function StopBall()
    Dir = MakeVector(0, 0, 0)
    InGame = False
End Function

Public Function ResetBall()
    Ball.MoveBall -Ball.GetPX, -Ball.GetPZ
    StopBall
End Function

Public Function UpdateBall()
    Dim PInter As Integer
    If Abs(Ball.GetPZ) + BallRadius >= FDepth Then
        If Ball.GetPZ > 0 Then
            P1Points = P1Points + 1
        Else
            P2Points = P2Points + 1
        End If
        FrmMain.P1Score = P1Points
        FrmMain.P2Score = P2Points
        ResetBall
    Else
        PInter = CheckCollision
        If PInter > 0 Then
            Dir.Z = -Dir.Z
        Else
            If Ball.GetPX + BallRadius >= 100 Then
                Dir.X = -Dir.X
            ElseIf Ball.GetPX - BallRadius <= -100 Then
                Dir.X = -Dir.X
            End If
        End If
    End If
    Ball.MoveBall Dir.X * BallSpeed, Dir.Z * BallSpeed
End Function

Public Function CheckCollision() As Integer
    If Ball.GetPZ + BallRadius <= 120 + PlayerDepth And Ball.GetPZ + BallRadius >= 120 And Ball.GetPZ > 0 Then
        If Ball.GetPX >= Player2.GetPX - PlayerWidth Then
            If Ball.GetPX <= Player2.GetPX + PlayerWidth Then
                CheckCollision = 2
            End If
        End If
    ElseIf Ball.GetPZ - BallRadius <= -120 + PlayerDepth And Ball.GetPZ - BallRadius >= -120 And Ball.GetPZ < 0 Then
        If Ball.GetPX >= Player1.GetPX - PlayerWidth Then
            If Ball.GetPX <= Player1.GetPX + PlayerWidth Then
                CheckCollision = 1
            End If
        End If
    End If
End Function

Public Function CheckKey()
    DIDevice.GetDeviceStateKeyboard DIState
    
    With DIState
        If .Key(DIK_LEFT) <> 0 Then
            MovePlayer1 -PlayerSpeed * Rot
        End If
        If .Key(DIK_RIGHT) <> 0 Then
            MovePlayer1 PlayerSpeed * Rot
        End If
        If P2Con = True Then
            If .Key(DIK_A) <> 0 Then
                MovePlayer2 -PlayerSpeed * Rot
            End If
            If .Key(DIK_D) <> 0 Then
                MovePlayer2 PlayerSpeed * Rot
            End If
        End If
        If .Key(DIK_RETURN) <> 0 Then
            If InGame = False Then
                StartBall
            End If
        End If
        If .Key(DIK_ESCAPE) <> 0 Then
            KillApp
        End If
    End With
    
End Function

Public Function MovePlayer2(Amount As Single)
    If Amount > 0 Then
        If Player2.GetPX + PlayerWidth + Amount <= 100 Then
            Player2.MovePlayer Amount
        End If
    Else
        If Player2.GetPX - PlayerWidth + Amount >= -100 Then
            Player2.MovePlayer Amount
        End If
    End If
End Function

Public Function MovePlayer1(Amount As Single)
    If Amount > 0 Then
        If Player1.GetPX + PlayerWidth + Amount <= 100 Then
            Player1.MovePlayer Amount
        End If
    Else
        If Player1.GetPX - PlayerWidth + Amount >= -100 Then
            Player1.MovePlayer Amount
        End If
    End If
End Function

Public Function AI()
    If Player2.GetPX > Ball.GetPX Then
        MovePlayer2 -PlayerSpeed
    Else
        MovePlayer2 PlayerSpeed
    End If
End Function
