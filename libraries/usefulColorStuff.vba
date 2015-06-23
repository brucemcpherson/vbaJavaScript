'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 3:54:06 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414615/raw
' this is all about colors
Option Explicit
' v2.7   3414615

Public Type colorProps
    ' this is a single type to hold everything i know how to calculate about a color
    rgb As Long
    red As Long
    green As Long
    blue As Long
    htmlHex As String
    textColor As Long
    luminance As Double
    contrastRatio As Double
    cyan As Double
    magenta As Double
    yellow As Double
    black As Double
    hue As Double
    saturation As Double
    lightness As Double
    value As Double
    x As Double
    y As Double
    z As Double
    LStar As Double
    aStar As Double
    bStar As Double
    cStar As Double
    hStar As Double
End Type
Enum eCompareColor
    eccieDe2000
End Enum

'Reference white for XYZ space/Observer = 2 deg Illuminant = D65
Const refWhiteX As Double = 95.047
Const refWhiteY As Double = 100
Const refWhiteZ As Double = 108.883
Const ref1 = 11181951
Const ref2 = 5934250
Public Function getlStar(rgbColor As Long) As Double
    getlStar = makeColorProps(rgbColor).LStar
End Function
Public Function getCstar(rgbColor As Long) As Double
    getCstar = makeColorProps(rgbColor).cStar
End Function

Public Function getBstar(rgbColor As Long) As Double
    getBstar = makeColorProps(rgbColor).bStar
End Function
Public Function getHstar(rgbColor As Long) As Double
    getHstar = makeColorProps(rgbColor).hStar
End Function
Public Function getAstar(rgbColor As Long) As Double
    getAstar = makeColorProps(rgbColor).aStar
End Function
Public Function fromRef(rgbColor As Long, ref As Long) As Double
    fromRef = compareColors(rgbColor, ref)
End Function
Public Function fromRefX(rgbColor As Long) As Double
    fromRefX = fromRef(rgbColor, ref1)
End Function
Public Function fromRefY(rgbColor As Long) As Double
    fromRefY = fromRef(rgbColor, ref2)
End Function
Public Function cellProperty(r As Range, p As String) As String
   ' find the excel property given the requested style
    Select Case p
        Case "background-color"
            cellProperty = rgbToHTMLHex(r.Interior.Color)
          
        Case "color"
            cellProperty = rgbToHTMLHex(r.Font.Color)
            
        Case "font-size"
            cellProperty = r.Font.size
        
        Case Else
            Debug.Assert False
        
    End Select
End Function
Public Function cellCss(r As Range, p As String) As String
    cellCss = p & ":" & cellProperty(r, p) & ";"
End Function
Public Function heatmapColor(min As Variant, _
                max As Variant, value As Variant) As Long
        heatmapColor = rampLibraryRGB("heatmap", min, max, value)
                    
End Function
Public Function rgbExpose(r As Long, g As Long, b As Long) As Long
    ' so i can use it in worksheets
    rgbExpose = rgb(r, g, b)
    
End Function
Public Function rgbRed(rgbColor As Long) As Long
    rgbRed = rgbColor Mod &H100
End Function
Public Function rgbGreen(rgbColor As Long) As Long
    rgbGreen = (rgbColor \ &H100) Mod &H100
End Function
Public Function rgbBlue(rgbColor As Long) As Long
    rgbBlue = (rgbColor \ &H10000) Mod &H100
End Function
Public Function rgbToHex(rgbColor As Long) As String
    ' just a synonym
    rgbToHex = rgbToHTMLHex(rgbColor)
End Function
Public Function rgbToHTMLHex(rgbColor As Long) As String

    ' just swap the colors round for rgb to bgr
    rgbToHTMLHex = "#" & maskFormat(Hex(rgb(rgbBlue(rgbColor), _
            rgbGreen(rgbColor), rgbRed(rgbColor))), "000000")

End Function
Public Function htmlHexToRgb(htmlHex As String) As Long
    Dim x As Long, s As String
    
    s = LTrim(RTrim(htmlHex))
    Debug.Assert Len(htmlHex) > 1 And left(htmlHex, 1) = "#"
    x = val("&H" & Right(s, Len(s) - 1) & "&")
    ' these are purposefully reversed since byte order is different in unix
    htmlHexToRgb = rgb(rgbBlue(x), rgbGreen(x), rgbRed(x))

End Function
 
Private Function maskFormat(sIn As String, f As String) As String
    Dim s As String
    s = sIn
    If Len(s) < Len(f) Then
        s = left(f, Len(f) - Len(s)) & s
    End If
    maskFormat = s
End Function

Private Function lumRGB(rgbCom As Double, brighten As Double) As Double
    Dim x As Double
    x = rgbCom * brighten
    If x > 255 Then x = 255
    If x < 0 Then x = 0
    lumRGB = x
    
End Function
Public Function rgbToHsl(rgbColor As Long) As colorProps
    ' adapted from // http://www.easyrgb.com/
    Dim r As Double, g As Double, b As Double, d As Double, _
        dr As Double, dg As Double, db As Double, mn As Double, mx As Double, _
        p As colorProps
    
    r = rgbRed(rgbColor) / 255
    g = rgbGreen(rgbColor) / 255
    b = rgbBlue(rgbColor) / 255
    mn = min(r, g, b)
    mx = max(r, g, b)
    d = mx - mn
    
    ' HSL sets here
    p.hue = 0
    p.saturation = 0
    ' lightness
    p.lightness = (mx + mn) / 2
    
    If (d <> 0) Then
        ' saturation
        If (p.lightness < 0.5) Then
            p.saturation = d / (mx + mn)
        Else
            p.saturation = d / (2 - mx - mn)
        End If
        ' hue
        dr = (((mx - r) / 6) + (d / 2)) / d
        dg = (((mx - g) / 6) + (d / 2)) / d
        db = (((mx - b) / 6) + (d / 2)) / d
        
        If r = mx Then
            p.hue = db - dg
        ElseIf g = mx Then
            p.hue = (1 / 3) + dr - db
        Else
            p.hue = (2 / 3) + dg - dr
        End If
        
        'force between 0 and 1
        If p.hue < 0 Then p.hue = p.hue + 1
        If p.hue > 1 Then p.hue = p.hue - 1
        Debug.Assert p.hue >= 0 And p.hue <= 1
    End If
    p.hue = p.hue * 360
    p.saturation = p.saturation * 100
    p.lightness = p.lightness * 100
    rgbToHsl = p
    
End Function
Private Function rgbToHsv(rgbColor As Long) As colorProps
    ' adapted from // http://www.easyrgb.com/
    Dim r As Double, g As Double, b As Double, _
        mn As Double, mx As Double, _
        p As colorProps
    
    r = rgbRed(rgbColor) / 255
    g = rgbGreen(rgbColor) / 255
    b = rgbBlue(rgbColor) / 255
    mn = min(r, g, b)
    mx = max(r, g, b)
    
    ' this is the same as hsl and hsv are the same.
    p = rgbToHsl(rgbColor)
    
    ' HSV sets here
    p.value = mx
    
    rgbToHsv = p
End Function
Private Function xyzCorrection(v As Double) As Double
    If (v > 0.04045) Then
        xyzCorrection = ((v + 0.055) / 1.055) ^ 2.4
    Else
        xyzCorrection = v / 12.92
    End If
End Function


Private Function xyzCIECorrection(v As Double) As Double
    If (v > 0.008856) Then
        xyzCIECorrection = (v ^ (1 / 3))
    Else
        xyzCIECorrection = (7.787 * v) + (16 / 116)
    End If
End Function
Private Function rgbToXyz(rgbColor As Long) As colorProps
    ' adapted from // http://www.easyrgb.com/
    Dim r As Double, g As Double, b As Double, _
        p As colorProps
    
    r = xyzCorrection(rgbRed(rgbColor) / 255) * 100
    g = xyzCorrection(rgbGreen(rgbColor) / 255) * 100
    b = xyzCorrection(rgbBlue(rgbColor) / 255) * 100
    
    p.x = r * 0.4124 + g * 0.3576 + b * 0.1805
    p.y = r * 0.2126 + g * 0.7152 + b * 0.0722
    p.z = r * 0.0193 + g * 0.1192 + b * 0.9505

    rgbToXyz = p
End Function
Private Function rgbToLab(rgbColor As Long) As colorProps
    ' adapted from // http://www.easyrgb.com/
    Dim x As Double, y As Double, z As Double, _
        p As colorProps

    p = rgbToXyz(rgbColor)
    
    x = xyzCIECorrection(p.x / refWhiteX)
    y = xyzCIECorrection(p.y / refWhiteY)
    z = xyzCIECorrection(p.z / refWhiteZ)

    p.LStar = (116 * y) - 16
    p.aStar = 500 * (x - y)
    p.bStar = 200 * (y - z)

    rgbToLab = p
End Function
Public Function findNearestColorInRange(rSearchFor As Range, rSearchIn As Range) As Range
    Dim r As Range, d As Double, dmin As Double, dr As Range, t As Long
    Set dr = Nothing
    t = rgbColorOf(firstCell(rSearchFor))
    For Each r In rSearchIn.Cells
        d = compareColors(rgbColorOf(r), t)
        If d < dmin Or dr Is Nothing Then
            Set dr = r
            dmin = d
        End If
    Next r
    Set findNearestColorInRange = dr
End Function


Public Function compareColors(rgb1 As Long, rgb2 As Long, _
            Optional compareType As eCompareColor = eCompareColor.eccieDe2000) As Double
    Dim p1 As colorProps, p2 As colorProps
    p1 = makeColorProps(rgb1)
    p2 = makeColorProps(rgb2)
    Select Case compareType
        Case eCompareColor.eccieDe2000
            compareColors = cieDe2000(p1, p2)
            
        Case Else
            Debug.Assert False
    
    End Select
    
End Function


Public Function cieDe2000(p1 As colorProps, p2 As colorProps) As Double
    ' calculates the distance between 2 colors using CIEDE200
    ' see http://www.ece.rochester.edu/~gsharma/cieDe2000/cieDe2000noteCRNA.pdf
    Dim c1 As Double, c2 As Double, _
        c As Double, g As Double, a1 As Double, b1 As Double, _
        a2 As Double, b2 As Double, c1Tick As Double, c2Tick As Double, _
        h1 As Double, h2 As Double, dh As Double, dl As Double, dc As Double, _
        lTickAvg As Double, cTickAvg As Double, hTickAvg As Double, l50 As Double, sl As Double, _
        sc As Double, t As Double, sh As Double, dTheta As Double, kp As Double, _
        rc As Double, kl As Double, kc As Double, kh As Double, dlk As Double, _
        dck As Double, dhk As Double, rt As Double, dBigH As Double
    
    kp = 25 ^ 7
    kl = 1
    kc = 1
    kh = 1
    
    ' calculate c & g values
    c1 = Sqr(p1.aStar ^ 2 + p1.bStar ^ 2)
    c2 = Sqr(p2.aStar ^ 2 + p2.bStar ^ 2)
    c = (c1 + c2) / 2
    g = 0.5 * (1 - Sqr(c ^ 7 / (c ^ 7 + kp)))

    ' adjusted ab*
    a1 = (1 + g) * p1.aStar
    a2 = (1 + g) * p2.aStar

    ' adjusted cs
    c1Tick = Sqr(a1 ^ 2 + p1.bStar ^ 2)
    c2Tick = Sqr(a2 ^ 2 + p2.bStar ^ 2)

    ' adjusted h
    h1 = computeH(a1, p1.bStar)
    h2 = computeH(a2, p2.bStar)

    
    ' deltas
    If (h2 - h1 > 180) Then '1
        dh = h2 - h1 - 360
    ElseIf (h2 - h1 < -180) Then ' 2
        dh = h2 - h1 + 360
    Else '0
        dh = h2 - h1
    End If

    dl = p2.LStar - p1.LStar
    dc = c2Tick - c1Tick
    dBigH = (2 * Sqr(c1Tick * c2Tick) * sIn(toRadians(dh / 2)))

    ' averages
    lTickAvg = (p1.LStar + p2.LStar) / 2
    cTickAvg = (c1Tick + c2Tick) / 2

    
    If (c1Tick * c2Tick = 0) Then '3
        hTickAvg = h1 + h2
    
    ElseIf (Abs(h2 - h1) <= 180) Then '0
        hTickAvg = (h1 + h2) / 2
    
    ElseIf (h2 + h1 < 360) Then '1
        hTickAvg = (h1 + h2) / 2 + 180
    
    Else '2
        hTickAvg = (h1 + h2) / 2 - 180
    End If
    
    l50 = (lTickAvg - 50) ^ 2
    sl = 1 + (0.015 * l50 / Sqr(20 + l50))
    sc = 1 + 0.045 * cTickAvg
    t = 1 - 0.17 * Cos(toRadians(hTickAvg - 30)) + 0.24 * _
            Cos(toRadians(2 * hTickAvg)) + 0.32 * _
            Cos(toRadians(3 * hTickAvg + 6)) - 0.2 * _
            Cos(toRadians(4 * hTickAvg - 63))

    sh = 1 + 0.015 * cTickAvg * t

    dTheta = 30 * Exp(-1 * ((hTickAvg - 275) / 25) ^ 2)
    rc = 2 * Sqr(cTickAvg ^ 7 / (cTickAvg ^ 7 + kp))
    rt = -sIn(toRadians(2 * dTheta)) * rc
    dlk = dl / sl / kl
    dck = dc / sc / kc
    dhk = dBigH / sh / kh
    cieDe2000 = Sqr(dlk ^ 2 + dck ^ 2 + dhk ^ 2 + rt * dck * dhk)
    
End Function
Private Function computeH(a As Double, b As Double) As Double
    If (a = 0 And b = 0) Then
        computeH = 0
    ElseIf (b >= 0) Then
        computeH = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Atan2(a, b))
    Else
        computeH = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Atan2(a, b)) + 360
    End If
End Function

Public Function hslToRgb(p As colorProps) As Long
    ' adapted from // http://www.easyrgb.com/
    Dim x1 As Double, x2 As Double, h As Double, s As Double, l As Double, _
        red As Double, green As Double, blue As Double
    
    
    h = p.hue / 360
    s = p.saturation / 100
    l = p.lightness / 100
    
    If s = 0 Then
        red = l * 255
        green = l * 255
        blue = l * 255
    Else
        If l < 0.5 Then
            x2 = l * (1 + s)
        Else
            x2 = (l + s) - (l * s)
        End If
        x1 = 2 * l - x2
        
        red = 255 * hueToRgb(x1, x2, h + (1 / 3))
        green = 255 * hueToRgb(x1, x2, h)
        blue = 255 * hueToRgb(x1, x2, h - (1 / 3))
        
     End If
     hslToRgb = rgb(red, green, blue)
     
End Function
Private Function hueToRgb(a As Double, b As Double, h As Double) As Double
   ' adapted from // http://www.easyrgb.com/
    If h < 0 Then h = h + 1
    If h > 1 Then h = h - 1
    Debug.Assert h >= 0 And h <= 1
    
    If (6 * h < 1) Then
        hueToRgb = a + (b - a) * 6 * h
    ElseIf (2 * h < 1) Then
        hueToRgb = b
    ElseIf (3 * h < 2) Then
        hueToRgb = a + (b - a) * ((2 / 3) - h) * 6
    Else
        hueToRgb = a
    End If
    
End Function

Public Function makeColorProps(rgbColor As Long) As colorProps
    Dim p As colorProps, p2 As colorProps
    
    'store the source color
    p.rgb = rgbColor
    
    'split the components
    p.red = rgbRed(rgbColor)
    p.green = rgbGreen(rgbColor)
    p.blue = rgbBlue(rgbColor)
    
    'the html hex rgb equivalent
    p.htmlHex = rgbToHTMLHex(rgbColor)
    
    'the w3 algo for luminance
    p.luminance = w3Luminance(rgbColor)
    
    'determine whether black or white background
    If (p.luminance < 0.5) Then
        p.textColor = vbWhite
    Else
        p.textColor = vbBlack
    End If

    'contrast ratio - to comply with w3 recs 1.4 should be at least 10:1 for text
    p.contrastRatio = contrastRatio(p.textColor, p.rgb)
    
    ' myck - just an estimate
    p.black = min(1 - p.red / 255, 1 - p.green / 255, 1 - p.blue / 255)
    If p.black < 1 Then
        p.cyan = (1 - p.red / 255 - p.black) / (1 - p.black)
        p.magenta = (1 - p.green / 255 - p.black) / (1 - p.black)
        p.yellow = (1 - p.blue / 255 - p.black) / (1 - p.black)
    End If
    
    ' calculate hsl + hsv and other wierd things
    p2 = rgbToHsl(p.rgb)
    p.hue = p2.hue
    p.saturation = p2.saturation
    p.lightness = p2.lightness
    
    p.value = rgbToHsv(p.rgb).value
    
    p2 = rgbToXyz(p.rgb)
    p.x = p2.x
    p.y = p2.y
    p.z = p2.z
    
    p2 = rgbToLab(p.rgb)
    p.LStar = p2.LStar
    p.aStar = p2.aStar
    p.bStar = p2.bStar
    
    p2 = rgbToLch(p.rgb)
    p.cStar = p2.cStar
    p.hStar = p2.hStar
    
    makeColorProps = p

End Function
Public Function pokeLchH(p As colorProps, newH As Double) As colorProps
    p.hStar = newH
    pokeLchH = p
End Function

Public Function lchToLab(p As colorProps) As colorProps
    Dim h As Double
    h = toRadians(p.hStar)
    p.aStar = Cos(h) * p.cStar
    p.bStar = sIn(h) * p.cStar
    lchToLab = p
End Function
Private Function labxyzCorrection(x As Double) As Double
    If (x ^ 3 > 0.008856) Then
        labxyzCorrection = x ^ 3
    Else
        labxyzCorrection = (x - 16 / 116) / 7.787
    End If
    
End Function
Public Function lchToRgb(p As colorProps) As Long
    lchToRgb = xyzToRgb(labToXyz(lchToLab(p)))
End Function

Private Function labToXyz(p As colorProps) As colorProps
    
    p.y = (p.LStar + 16) / 116
    p.x = p.aStar / 500 + p.y
    p.z = p.y - p.bStar / 200
    
    p.x = labxyzCorrection(p.x) * refWhiteX
    p.y = labxyzCorrection(p.y) * refWhiteY
    p.z = labxyzCorrection(p.z) * refWhiteZ

    labToXyz = p

End Function

Private Function xyzrgbCorrection(x As Double) As Double
    If (x > 0.0031308) Then
        xyzrgbCorrection = 1.055 * (x ^ (1 / 2.4)) - 0.055
    Else
        xyzrgbCorrection = 12.92 * x
    End If
    
End Function
Public Function xyzToRgb(p As colorProps) As Long
    Dim r As Double, g As Double, b As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim x As Double, y As Double, z As Double, c As Double
    x = p.x / 100
    y = p.y / 100
    z = p.z / 100
    
    
    x1 = x * 0.8951 + y * 0.2664 + z * -0.1614
    y1 = x * -0.7502 + y * 1.7135 + z * 0.0367
    z1 = x * 0.0389 + y * -0.0685 + z * 1.0296
    
    x2 = x1 * 0.98699 + y1 * -0.14705 + z1 * 0.15997
    y2 = x1 * 0.43231 + y1 * 0.51836 + z1 * 0.04929
    z2 = x1 * -0.00853 + y1 * 0.04004 + z1 * 0.96849
    
    r = xyzrgbCorrection(x2 * 3.240479 + y2 * -1.53715 + z2 * -0.498535)
    g = xyzrgbCorrection(x2 * -0.969256 + y2 * 1.875992 + z2 * 0.041556)
    b = xyzrgbCorrection(x2 * 0.055648 + y2 * -0.204043 + z2 * 1.057311)

    c = rgb(min(255, max(0, CLng(r * 255))), _
                   min(255, max(0, CLng(g * 255))), _
                   min(255, max(0, CLng(b * 255))))
   
    
    xyzToRgb = c
End Function
Public Function rgbWashout(rgbColor As Long) As Long
    ' take a color and wash it out
    Dim p As colorProps
    p = makeColorProps(rgbColor)
    p.saturation = p.saturation * 0.2
    p.lightness = p.lightness * 0.9

    rgbWashout = hslToRgb(p)
End Function
Public Function rgbToLch(rgbColor As Long) As colorProps
    ' convert from cieL*a*b* to cieL*CH
    ' adapted from http://www.brucelindbloom.com/index.html?Equations.html


    Dim p As colorProps
    p = rgbToLab(rgbColor)
    If rgbColor = 0 Then
        p.hStar = 0
    Else
        p.hStar = Application.WorksheetFunction.Atan2(p.aStar, p.bStar)
        If p.hStar > 0 Then
            p.hStar = fromRadians(p.hStar)
        Else
            p.hStar = 360 - fromRadians(Abs(p.hStar))
        End If
    End If
    p.cStar = Sqr(p.aStar * p.aStar + p.bStar * p.bStar)
    rgbToLch = p

End Function
Public Function contrastRatio(rgbColorA As Long, rgbColorB As Long) As Double
    Dim lumA As Double, lumB As Double
    lumA = w3Luminance(rgbColorA)
    lumB = w3Luminance(rgbColorB)

    contrastRatio = (max(lumA, lumB) + 0.05) / (min(lumA, lumB) + 0.05)

End Function

Public Function w3Luminance(rgbColor As Long) As Double
' this is based on
' http://en.wikipedia.org/wiki/Luma_(video)

  w3Luminance = (0.2126 * ((rgbRed(rgbColor) / 255) ^ 2.2)) + _
         (0.7152 * ((rgbGreen(rgbColor) / 255) ^ 2.2)) + _
         (0.0722 * ((rgbBlue(rgbColor) / 255) ^ 2.2))

End Function
Public Function rampLibraryRGB(ramp As Variant, min As Variant, _
                max As Variant, value As Variant, _
                Optional brighten As Double = 1) As Long
                Dim x As Long
                
    If IsArray(ramp) Then
    ' ramp colors have been passed here
                rampLibraryRGB = colorRamp(min, max, value, _
                                ramp, , _
                                brighten)
    Else
    
        Select Case Trim(LCase(CStr(ramp)))
            Case "heatmaptowhite"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbBlue, vbGreen, vbYellow, vbRed, vbWhite), , _
                                brighten)
            
            Case "heatmap"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbBlue, vbGreen, vbYellow, vbRed), , _
                                brighten)
            
            Case "blacktowhite"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbBlack, vbWhite), , brighten)
            
            Case "whitetoblack"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbWhite, vbBlack), , brighten)
                         
            Case "hotinthemiddle"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbBlue, vbGreen, vbYellow, vbRed, _
                                        vbYellow, vbGreen, vbBlue), , brighten)
                                        
            Case "candylime"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(rgb(255, 77, 121), rgb(255, 121, 77), _
                                        rgb(255, 210, 77), rgb(210, 255, 77)), , _
                                        brighten)
                                        
            Case "heatcolorblind"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbBlack, vbBlue, vbRed, vbWhite), , brighten)
                                
            Case "gethotquick"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbBlue, vbGreen, vbYellow, vbRed), _
                                Array(0, 0.1, 0.25, 1), brighten)
                  
            Case "slowramp"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbBlack, rgb(0, 46, 184), rgb(0, 138, 184), _
                                rgb(0, 184, 138), _
                                rgb(138, 184, 0), rgb(184, 138, 0), _
                                rgb(138, 0, 184)), _
                                Array(0, 0.04, 0.1, 0.15, 0.22, 0.3, 1), brighten)
                  
            Case "greensweep"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(rgb(153, 204, 51), rgb(51, 204, 179)), , _
                                brighten)
                   
            Case "terrain"
                rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbBlack, rgb(0, 46, 184), rgb(0, 138, 184), _
                                rgb(0, 184, 138), _
                                rgb(138, 184, 0), rgb(184, 138, 0), _
                                rgb(138, 0, 184), vbWhite), , _
                                brighten)
                                
            Case "terrainnosea"
                 rampLibraryRGB = colorRamp(min, max, value, _
                                Array(vbGreen, rgb(0, 184, 138), _
                                rgb(138, 184, 0), rgb(184, 138, 0), _
                                rgb(138, 0, 184), vbWhite), , _
                                brighten)
            Case "greendollar"
                 rampLibraryRGB = colorRamp(min, max, value, _
                                Array(rgb(225, 255, 235), _
                                rgb(2, 202, 69)), , _
                                brighten)
    
            Case "lightblue"
                 rampLibraryRGB = colorRamp(min, max, value, _
                                Array(rgb(230, 237, 246), _
                                rgb(163, 189, 271)), , _
                                brighten)
                                
            Case "lightorange"
                 rampLibraryRGB = colorRamp(min, max, value, _
                                Array(rgb(253, 233, 217), _
                                rgb(244, 132, 40)), , _
                                brighten)
            Case Else
                MsgBox ramp & " is an unknown library entry"
                
        End Select
    End If
End Function
Public Function colorRamp(min As Variant, _
                max As Variant, value As Variant, _
                Optional mileStones As Variant, _
                Optional fractionStones As Variant, _
                Optional brighten As Double = 1) As Long
    
    ' create a value from a colorramp going through the array of milestones
    Dim spread As Double, ratio As Double, red As Double, _
                    green As Double, blue As Double, j As Long, _
                    lb As Long, ub As Long, cb As Long, r As Double, i As Long
    '----defaults and set up milestones on ramp
    Dim ms() As Long
    Dim fs() As Double


    If IsMissing(mileStones) Then
        ReDim ms(0 To 4)
        ms(0) = vbBlue
        ms(1) = vbGreen
        ms(2) = vbYellow
        ms(3) = vbRed
        ms(4) = vbWhite
    Else
        ReDim ms(0 To UBound(mileStones) - LBound(mileStones))
        j = 0
        For i = LBound(mileStones) To UBound(mileStones)
            ms(j) = mileStones(i)
            j = j + 1
        Next i
    End If
    ' tedious this is
    lb = LBound(ms)
    ub = UBound(ms)
    cb = ub - lb + 1
    ' only 1 milestone - thats the color
    If cb = 1 Then
        colorRamp = ms(lb)
        Exit Function
    End If
        
    If Not IsMissing(fractionStones) Then
        If UBound(fractionStones) - LBound(fractionStones) <> _
            cb - 1 Then
            MsgBox ("no of fractions must equal number of steps")
            Exit Function
        Else
            ReDim fs(lb To ub)
            j = lb
            For i = LBound(fractionStones) To UBound(fractionStones)
                fs(j) = fractionStones(i)
                j = j + 1
            Next i

        End If
    Else
        ReDim fs(lb To ub)
        For i = lb + 1 To ub
            fs(i) = i / (cb - 1)
        Next i
    End If
    'spread of range
    spread = max - min
    Debug.Assert spread >= 0
    If spread = 0 Then spread = 0.5
    ratio = (value - min) / spread
    Debug.Assert ratio >= 0 And ratio <= 1
    ' find which slot
    For i = lb + 1 To ub
        If ratio <= fs(i) Then
            r = (ratio - fs(i - 1)) / (fs(i) - fs(i - 1))
            red = rgbRed(ms(i - 1)) + (rgbRed(ms(i)) - rgbRed(ms(i - 1))) * r
            blue = rgbBlue(ms(i - 1)) + (rgbBlue(ms(i)) - rgbBlue(ms(i - 1))) * r
            green = rgbGreen(ms(i - 1)) + (rgbGreen(ms(i)) - rgbGreen(ms(i - 1))) * r
            colorRamp = rgb(lumRGB(red, brighten), _
                            lumRGB(green, brighten), _
                            lumRGB(blue, brighten))
            Exit Function
        End If
    Next i
    Debug.Assert False
   
End Function


Public Sub applyHeatMapToRange(rIn As Range, Optional libraryEntry As String = "heatmap")
    Dim mx As Variant, mn As Variant, r As Range, c As colorProps
    mx = Application.WorksheetFunction.max(rIn)
    mn = Application.WorksheetFunction.min(rIn)
    For Each r In rIn.Cells
        c = makeColorProps(rampLibraryRGB(libraryEntry, mn, mx, r.value))
        r.Interior.Color = c.rgb
        r.Font.Color = c.textColor
    Next r
End Sub
Public Function hexColorOf(r As Range) As String
    ' this should be volatile but turning off
    ''Application.Volatile
    hexColorOf = rgbToHTMLHex(rgbColorOf(r))
End Function
Public Function rgbColorOf(r As Range) As Long
    rgbColorOf = r.Interior.Color
End Function
Public Sub colorizeCell(target As Range, c As String)
    Dim p As colorProps
    If Len(c) > 1 And left(c, 1) = "#" Then
        p = makeColorProps(htmlHexToRgb(c))
        target.Interior.Color = p.rgb
        target.Font.Color = p.textColor
    End If
End Sub

Private Function colorPropBigger(a As colorProps, b As colorProps, byProp As String, _
            Optional descending As Boolean = False) As Boolean
    Dim result As Boolean
    
    Select Case LCase(byProp)
        Case "hue"
            result = a.hue > b.hue
        
        Case "saturation"
            result = a.saturation > b.saturation
            
        Case "lightness"
            result = a.lightness > b.lightness
            
        Case "lstar"
            result = a.LStar > b.LStar
            
        Case "cstar"
            result = a.cStar > b.cStar
            
        Case "hstar"
            result = a.hStar > b.hStar
        
        Case Else
            Debug.Assert False
    
    End Select
    If descending Then
        colorPropBigger = Not result
    Else
        colorPropBigger = result
    End If
    
End Function

Public Sub sortColorProp(pArray() As colorProps, inlow As Long, inhi As Long, byProp As String, _
        Optional descending As Boolean = False)

  Dim p As colorProps, swap As colorProps, low  As Long, hi   As Long, half As Long


    If inlow < inhi Then
        half = (inlow + inhi) \ 2
        p = pArray(half)
        low = inlow
        hi = inhi
        Do
            Do While colorPropBigger(p, pArray(low), byProp, descending)
                low = low + 1
            Loop
            Do While colorPropBigger(pArray(hi), p, byProp, descending)
                hi = hi - 1
            Loop
            If (low <= hi) Then
               swap = pArray(low)
               pArray(low) = pArray(hi)
               pArray(hi) = swap
               low = low + 1
               hi = hi - 1
            End If
            
        Loop Until low > hi

        If hi <= half Then
            sortColorProp pArray, inlow, hi, byProp, descending
            sortColorProp pArray, low, inhi, byProp, descending
        Else
            sortColorProp pArray, low, inhi, byProp, descending
            sortColorProp pArray, inlow, hi, byProp, descending
        End If
    End If

End Sub



