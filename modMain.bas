Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'// Main DirectX Declarations
Public mDX As New DirectX7
Public mDDraw As DirectDraw7
Public mD3D As Direct3D7
Public mD3DDevice As Direct3DDevice7

'// Vertex Declarations
Public mtlSprite(3) As D3DTLVERTEX
Public mtlOverlay(3) As D3DTLVERTEX
Public mtlBehindText(3) As D3DTLVERTEX
Public mtlLight(3) As D3DTLVERTEX

'// Surfaces / Textures Declarations
Public msFront As DirectDrawSurface7
Public msBack As DirectDrawSurface7
Public msFrame1 As DirectDrawSurface7
Public msFrame2 As DirectDrawSurface7
Public msBackground As DirectDrawSurface7
Public msOverlay As DirectDrawSurface7
Public msLight As DirectDrawSurface7

'// Screen Declarations
Public SCREEN_WIDTH As Long
Public SCREEN_HEIGHT As Long
Public SCREEN_DEPTH As Long
Public SCREEN_BACKCOLOR As Long

'// Other Declarations
Public mbRunning As Boolean
Public bShowInfo As Boolean
Public bDrawBG As Boolean
Public Sub InitDX(Width As Long, Height As Long, Depth As Long, DeviceGUID As String)
    Dim ddsd As DDSURFACEDESC2
    Dim caps As DDSCAPS2
    
    '// Create DirectDraw object
    Set mDDraw = mDX.DirectDrawCreate("")
    
    '// Set Cooperative Level (fullscreen, exclusive access)
    mDDraw.SetCooperativeLevel frmMain.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT
    mDDraw.SetDisplayMode Width, Height, Depth, 0, DDSDM_DEFAULT
    
    '// Create primary surface
    ddsd.lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_3DDEVICE Or DDSCAPS_PRIMARYSURFACE
    ddsd.lBackBufferCount = 1
    
    Set msFront = mDDraw.CreateSurface(ddsd)
    
    '// Create the backbuffer (used for 3D drawing)
    caps.lCaps = DDSCAPS_BACKBUFFER Or DDSCAPS_3DDEVICE
    Set msBack = msFront.GetAttachedSurface(caps)
    
    '// Create Direct3D
    Set mD3D = mDDraw.GetDirect3D
    
    '// Set the Device
    Set mD3DDevice = mD3D.CreateDevice(DeviceGUID, msBack)
End Sub

Public Sub Start(DeviceGUID As String)
    Dim lColor As Long
    
    '// Show the Form
    frmMain.Show
    DoEvents
    
    '// Screen Settings
    SCREEN_WIDTH = 640
    SCREEN_HEIGHT = 480
    SCREEN_DEPTH = 16
    SCREEN_BACKCOLOR = RGB2DX(0, 96, 184)
    
    '// Initialize DirectX at 640x480x16
    Call InitDX(SCREEN_WIDTH, SCREEN_HEIGHT, SCREEN_DEPTH, DeviceGUID)
    
    '// Load Textures and Surfaces
    Set msFrame1 = CreateTexture(App.Path & "\Frame1.bmp", 64, 64, Magenta)
    Set msFrame2 = CreateTexture(App.Path & "\Frame2.bmp", 64, 64, Magenta)
    Set msBackground = CreateTexture(App.Path & "\Background.bmp", 640, 480, None)
    Set msOverlay = CreateTexture(App.Path & "\Overlay.bmp", 256, 256, Black)
    Set msLight = CreateTexture(App.Path & "\Light.bmp", 256, 256, None)
    
    '// Create 'Behind Text' vertices
    lColor = mDX.CreateColorRGBA(0, 0, 0, 0.5)
    Call mDX.CreateD3DTLVertex(0, 0, 0, 1, lColor, 0, 0, 0, mtlBehindText(0))
    Call mDX.CreateD3DTLVertex(350, 0, 0, 1, lColor, 0, 0, 0, mtlBehindText(1))
    Call mDX.CreateD3DTLVertex(0, 55, 0, 1, lColor, 0, 0, 0, mtlBehindText(2))
    Call mDX.CreateD3DTLVertex(350, 55, 0, 1, lColor, 0, 0, 0, mtlBehindText(3))
    
    '// Enable color key
    mD3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
    
    '// Hide info and Show Background
    bShowInfo = False
    bDrawBG = True
    
    '// Hide the cursor
    Call ShowCursor(0)

    '// Start Running
    Call MainLoop
    
    '// Show the cursor
    Call ShowCursor(1)
    
    '// End it all
    Call Terminate
End Sub


Public Sub ClearDevice()
    '// Clear the device for drawing operations
    Dim rClear(0) As D3DRECT
    
    rClear(0).X2 = SCREEN_WIDTH
    rClear(0).Y2 = SCREEN_HEIGHT
    mD3DDevice.Clear 1, rClear, D3DCLEAR_TARGET, SCREEN_BACKCOLOR, 0, 0
End Sub

Public Sub MainLoop()
    Dim hBackDC As Long
    Dim hBGDC As Long
    Dim rBG As RECT
    
    Dim iOffsetAngle As Integer
    Dim sAlpha As Single
    Dim iAngle As Integer
    Dim iFrame As Integer
    Dim lXOffset As Long
    Dim lYOffset As Long
    Dim lColor As Long
    
    Dim FramesDone As Integer
    Dim LastFrame As Long
    Dim FrameRate As Integer
    
    mbRunning = True
    iAngle = 180
    iOffsetAngle = -45
    
    '// Draw until the program should stop running
    Do While mbRunning
        '// Rotate sprite
        iAngle = iAngle + 2
        If iAngle > 360 Then iAngle = iAngle - 360
        
        '// Move frame
        If iAngle / 5 = CInt(iAngle / 5) Then
            iFrame = iFrame + 1
            If iFrame = 2 Then iFrame = 0
        End If
        
        '// Calculate Offset
        iOffsetAngle = iOffsetAngle + 2
        If iOffsetAngle > 360 Then iOffsetAngle = iOffsetAngle - 360
        
        lXOffset = CalcCoordX(SCREEN_WIDTH / 2, 75, iOffsetAngle)
        lYOffset = CalcCoordY(SCREEN_HEIGHT / 2, 75, iOffsetAngle)
        
        '// Create the four vertices which make the sprite
        lColor = mDX.CreateColorRGBA(1, 1, 1, sAlpha)
        Call mDX.CreateD3DTLVertex(CalcCoordX(lXOffset, 32, iAngle), CalcCoordY(lYOffset, 32, iAngle), 0, 1, lColor, 0, 0, 0, mtlSprite(0))
        Call mDX.CreateD3DTLVertex(CalcCoordX(lXOffset, 32, iAngle - 90), CalcCoordY(lYOffset, 32, iAngle - 90), 0, 1, lColor, 0, 1, 0, mtlSprite(1))
        Call mDX.CreateD3DTLVertex(CalcCoordX(lXOffset, 32, iAngle + 90), CalcCoordY(lYOffset, 32, iAngle + 90), 0, 1, lColor, 0, 0, 1, mtlSprite(2))
        Call mDX.CreateD3DTLVertex(CalcCoordX(lXOffset, 32, iAngle + 180), CalcCoordY(lYOffset, 32, iAngle + 180), 0, 1, lColor, 0, 1, 1, mtlSprite(3))
        
        '// Create overlay vertices
        If iAngle < 180 Then
            sAlpha = iAngle / 180
        Else
            sAlpha = (360 - iAngle) / 180
        End If
        
        lColor = mDX.CreateColorRGBA(1, 1, 1, sAlpha)
        Call mDX.CreateD3DTLVertex(SCREEN_WIDTH - 256, SCREEN_HEIGHT - 256, 0, 1, lColor, 0, 0, 0, mtlOverlay(0))
        Call mDX.CreateD3DTLVertex(SCREEN_WIDTH - 256 + 256, SCREEN_HEIGHT - 256, 0, 1, lColor, 0, 1, 0, mtlOverlay(1))
        Call mDX.CreateD3DTLVertex(SCREEN_WIDTH - 256, SCREEN_HEIGHT - 256 + 256, 0, 1, lColor, 0, 0, 1, mtlOverlay(2))
        Call mDX.CreateD3DTLVertex(SCREEN_WIDTH - 256 + 256, SCREEN_HEIGHT - 256 + 256, 0, 1, lColor, 0, 1, 1, mtlOverlay(3))
        
        '// Create light vertices
        lColor = mDX.CreateColorRGBA(1, 1, 1, 0.5)
        Call mDX.CreateD3DTLVertex(lXOffset - 128, lYOffset - 128, 0, 1, lColor, 0, 0, 0, mtlLight(0))
        Call mDX.CreateD3DTLVertex(lXOffset + 128, lYOffset - 128, 0, 1, lColor, 0, 1, 0, mtlLight(1))
        Call mDX.CreateD3DTLVertex(lXOffset - 128, lYOffset + 128, 0, 1, lColor, 0, 0, 1, mtlLight(2))
        Call mDX.CreateD3DTLVertex(lXOffset + 128, lYOffset + 128, 0, 1, lColor, 0, 1, 1, mtlLight(3))
        
        '// Clear the device
        Call ClearDevice
        
        '// Draw the background
        If bDrawBG Then
            With rBG
                .Top = 0
                .Left = 0
                .Bottom = 480
                .Right = 640
            End With
            
            msBack.BltFast 0, 0, msBackground, rBG, DDBLTFAST_WAIT
        End If
        
        '// Start the scene
        mD3DDevice.BeginScene
        
            '// Enable Alpha Blending
            mD3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
            
            '// Set Alpha Blending parameters
            mD3DDevice.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_ONE
            mD3DDevice.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_ONE
            
            '// Draw Light
            mD3DDevice.SetTexture 0, msLight
            Call mD3DDevice.DrawPrimitive(D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlLight(0), 4, D3DDP_DEFAULT)
            
            '// Set Alpha Blending parameters
            mD3DDevice.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_SRCALPHA
            mD3DDevice.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_ONE
            
            '// Draw Overlay
            mD3DDevice.SetTexture 0, msOverlay
            Call mD3DDevice.DrawPrimitive(D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlOverlay(0), 4, D3DDP_DEFAULT)
            
            '// Draw 'Behind Text'
            mD3DDevice.SetTexture 0, Nothing
            Call mD3DDevice.DrawPrimitive(D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlBehindText(0), 4, D3DDP_DEFAULT)
            
            '// Disable Alpha Blending
            mD3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
            
            '// Set texture
            If iFrame = 0 Then
                mD3DDevice.SetTexture 0, msFrame1
            Else
                mD3DDevice.SetTexture 0, msFrame2
            End If
            
            '// Draw sprite
            Call mD3DDevice.DrawPrimitive(D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlSprite(0), 4, D3DDP_DEFAULT)
            
        '// End the scene
        mD3DDevice.EndScene
        
        '// Calculate FPS (Frames Per Seconds)
        FramesDone = FramesDone + 1
        If mDX.TickCount - LastFrame >= 1000 Then
            FrameRate = FramesDone
            FramesDone = 0
            LastFrame = mDX.TickCount
        End If
        
        '// Draw Text
        msBack.SetForeColor RGB(255, 255, 255)
        msBack.DrawText 5, 5, "Press 'Esc' to exit... (FPS: " & CStr(FrameRate) & ")", False
        msBack.DrawText 5, 20, "Press 'I' to turn information on/off (On: " & CStr(bShowInfo) & ")", False
        msBack.DrawText 5, 35, "Press 'B' to turn the background on/off (On: " & CStr(bDrawBG) & ")", False
        
        '// Draw some more text
        If bShowInfo Then
            msBack.DrawText 320, 450, "Alpha-Blended ->", False
            msBack.DrawText 310, 5, "<- Alpha-Blended", False
            msBack.DrawText lXOffset + 75, lYOffset, "<- Light=Alpha-Blended", False
            msBack.DrawText lXOffset - 150, lYOffset, "Mario=Rotated ->", False
        End If
        
        '// Draw Circles
        msBack.SetForeColor RGB(33, 148, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 120
        msBack.SetForeColor RGB(65, 170, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 115
        msBack.SetForeColor RGB(135, 197, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 110
        msBack.SetForeColor RGB(165, 225, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 105
        msBack.SetForeColor RGB(255, 255, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 100
        
        '// Draw lines around 'Behind Text'
        msBack.SetForeColor RGB(0, 0, 192)
        msBack.DrawLine 0, 55, 351, 55
        msBack.DrawLine 350, 0, 350, 55
        msBack.SetForeColor RGB(0, 0, 128)
        msBack.DrawLine 0, 56, 352, 56
        msBack.DrawLine 351, 0, 351, 56
        
        '// Flip
        msFront.Flip Nothing, DDFLIP_WAIT
        DoEvents
    Loop
End Sub

Public Sub Terminate()
    '// Clean up DirectX
    Call mDDraw.RestoreDisplayMode
    Call mDDraw.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    
    Set mD3DDevice = Nothing
    Set mD3D = Nothing
    Set msBack = Nothing
    Set msFront = Nothing
    Set mDDraw = Nothing
    Set mDX = Nothing
    
    Unload frmMain
End Sub


