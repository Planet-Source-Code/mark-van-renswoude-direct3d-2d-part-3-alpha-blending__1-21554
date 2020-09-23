Attribute VB_Name = "modDXFunctions"
Option Explicit

Public Const PI = 3.14
Public Const Rad = PI / 180

Public Enum COLORKEYOPTIONS
    None
    Black
    White
    Magenta
End Enum
Private Function ExclusiveMode() As Boolean
    Dim lngTestExMode As Long
    
    '// This function tests if we're still in exclusive mode
    lngTestExMode = mDDraw.TestCooperativeLevel
    
    If (lngTestExMode = DD_OK) Then
        ExclusiveMode = True
    Else
        ExclusiveMode = False
    End If
End Function
Public Function LostSurfaces() As Boolean
    '// This function will tell if we should reload our bitmaps or not
    LostSurfaces = False
    Do Until ExclusiveMode
        DoEvents
        LostSurfaces = True
    Loop
    
    '// If we did lose our bitmaps, restore the surfaces and return 'true'
    DoEvents
    If LostSurfaces Then
        mDDraw.RestoreAllSurfaces
    End If
End Function
Public Function CalcCoordX(Offset As Long, Length As Long, Angle As Integer)
    '// Calculate X coordinate
    CalcCoordX = Offset + Sin(Angle * Rad) * Length
End Function
Public Function CalcCoordY(Offset As Long, Length As Long, Angle As Integer)
    '// Calculate X coordinate
    CalcCoordY = Offset + Cos(Angle * Rad) * Length
End Function


Public Function CreateSurface(File As String, Width As Long, Height As Long, Optional ColKey As COLORKEYOPTIONS = None)
    Dim msSurface As DirectDrawSurface7
    Dim ddsd As DDSURFACEDESC2
    
    '// Set surface description
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN

    '// Load surfaces
    ddsd.lHeight = Width
    ddsd.lWidth = Height
    
    '// If no filename was passed, create a blank surface
    If File = "" Then
        Set msSurface = mDDraw.CreateSurface(ddsd)
    Else
        Set msSurface = mDDraw.CreateSurfaceFromFile(File, ddsd)
    End If
    
    '// Set colour key
    Dim cKey As DDCOLORKEY
    Dim ddpf As DDPIXELFORMAT
    
    Select Case ColKey
        Case COLORKEYOPTIONS.Black
            '// Black colorkey
            cKey.low = 0
            cKey.high = 0
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
        Case COLORKEYOPTIONS.White
            '// White colorkey
            msSurface.GetPixelFormat ddpf
            cKey.low = ddpf.lRBitMask + ddpf.lGBitMask + ddpf.lBBitMask
            cKey.high = cKey.low
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
        Case COLORKEYOPTIONS.Magenta
            '// Magenta colorkey
            msSurface.GetPixelFormat ddpf
            cKey.low = ddpf.lRBitMask + ddpf.lBBitMask
            cKey.high = cKey.low
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
    End Select
    
    '// Return surface
    Set CreateSurface = msSurface
End Function

Public Function RGB2DX(R As Long, G As Long, B As Long) As Long
    '// Convert RGB code to DX code
    RGB2DX = mDX.CreateColorRGBA(CSng((1 / 255) * R), CSng((1 / 255) * G), CSng((1 / 255) * B), 0)
End Function


Public Function CreateTexture(File As String, Width As Long, Height As Long, Optional ColKey As COLORKEYOPTIONS = None) As DirectDrawSurface7
    Dim enumTex As Direct3DEnumPixelFormats
    Dim msSurface As DirectDrawSurface7
    Dim ddsd As DDSURFACEDESC2
    Dim bOK As Boolean
    Dim lK As Long
    
    '// Set flags to indicate it's a texture
    ddsd.lFlags = DDSD_CAPS Or DDSD_TEXTURESTAGE Or DDSD_PIXELFORMAT
    
    '// If width and height were specified, make it that size,
    '// otherwise it will be it's normal size
    If Height <> 0 And Width <> 0 Then
        ddsd.lFlags = ddsd.lFlags Or DDSD_HEIGHT Or DDSD_WIDTH
        ddsd.lHeight = Height
        ddsd.lWidth = Width
    End If
    
    '// Check if device supports 16 bit surface
    Set enumTex = mD3DDevice.GetTextureFormatsEnum()
    
    For lK = 1 To enumTex.GetCount()
        bOK = True
        Call enumTex.GetItem(lK, ddsd.ddpfPixelFormat)

        With ddsd.ddpfPixelFormat
            If .lRGBBitCount <> 16 Then bOK = False
        End With
        
        If bOK = True Then Exit For
    Next

    If bOK = False Then
        '// No support for 16 bit textures, raise error
        Err.Raise 8001, , "No support for 16 bit textures found!"
        Exit Function
    End If
    
    '// Set texture flags
    If mD3DDevice.GetDeviceGuid() = "IID_IDirect3DHALDevice" Then
        ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
        ddsd.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        ddsd.lTextureStage = 0
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
        ddsd.ddsCaps.lCaps2 = 0
        ddsd.lTextureStage = 0
    End If
    
    '// If no filename was passed, create a blank surface
    If File = "" Then
        Set msSurface = mDDraw.CreateSurface(ddsd)
    Else
        Set msSurface = mDDraw.CreateSurfaceFromFile(File, ddsd)
    End If

    '// Set colour key
    Dim cKey As DDCOLORKEY
    Dim ddpf As DDPIXELFORMAT
    
    Select Case ColKey
        Case COLORKEYOPTIONS.Black
            '// Black colorkey
            cKey.low = 0
            cKey.high = 0
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
        Case COLORKEYOPTIONS.White
            '// White colorkey
            msSurface.GetPixelFormat ddpf
            cKey.low = ddpf.lRBitMask + ddpf.lGBitMask + ddpf.lBBitMask
            cKey.high = cKey.low
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
        Case COLORKEYOPTIONS.Magenta
            '// Magenta colorkey
            msSurface.GetPixelFormat ddpf
            cKey.low = ddpf.lRBitMask + ddpf.lBBitMask
            cKey.high = cKey.low
            msSurface.SetColorKey DDCKEY_SRCBLT, cKey
    End Select
    
    '// Return surface
    Set CreateTexture = msSurface
End Function
