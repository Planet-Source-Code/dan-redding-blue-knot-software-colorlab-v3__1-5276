Attribute VB_Name = "modHSL"
Option Explicit

'For the FULL version of this module, please visit
' http://www.planet-source-code.com/vb
'(The darken & brighten routines in this module are
'slightly modified from that version)

'Portions of this code marked with *** are converted from
'C/C++ routines for RGB/HSL conversion found in the
'Microsoft Knowledge Base (PD sample code):
'http://support.microsoft.com/support/kb/articles/Q29/2/40.asp
'In addition to the language conversion, some internal
'calculations have been modified and converted to FP math to
'reduce rounding errors.
'Conversion to VB and original code by
'Dan Redding (bwsoft@revealed.net)
'http://home.revealed.net/bwsoft
'Free to use, please give proper credit

Public Const HSLMAX As Integer = 240 '***
    'H, S and L values can be 0 - HSLMAX
    '240 matches what is used by MS Win;
    'any number less than 1 byte is OK;
    'works best if it is evenly divisible by 6
Const RGBMAX As Integer = 255 '***
    'R, G, and B value can be 0 - RGBMAX
Const UNDEFINED As Integer = (HSLMAX * 2 / 3) '***
    'Hue is undefined if Saturation = 0 (greyscale)

Public Type HSLCol 'Datatype used to pass HSL Color values
    Hue As Integer
    Sat As Integer
    Lum As Integer
End Type

Public Function RGBRed(RGBCol As Long) As Integer
'Return the Red component from an RGB Color
    RGBRed = RGBCol And &HFF
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
'Return the Green component from an RGB Color
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
'Return the Blue component from an RGB Color
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function

Private Function iMax(a As Integer, B As Integer) _
    As Integer
'Return the Larger of two values
    iMax = IIf(a > B, a, B)
End Function

Private Function iMin(a As Integer, B As Integer) _
    As Integer
'Return the smaller of two values
    iMin = IIf(a < B, a, B)
End Function

Public Function RGBtoHSL(RGBCol As Long) As HSLCol '***
'Returns an HSLCol datatype containing Hue, Luminescence
'and Saturation; given an RGB Color value

Dim R As Integer, G As Integer, B As Integer
Dim cMax As Integer, cMin As Integer
Dim RDelta As Double, GDelta As Double, _
    BDelta As Double
Dim H As Double, s As Double, L As Double
Dim cMinus As Long, cPlus As Long
    
    R = RGBRed(RGBCol)
    G = RGBGreen(RGBCol)
    B = RGBBlue(RGBCol)
    
    cMax = iMax(iMax(R, G), B) 'Highest and lowest
    cMin = iMin(iMin(R, G), B) 'color values
    
    cMinus = cMax - cMin 'Used to simplify the
    cPlus = cMax + cMin  'calculations somewhat.
    
    'Calculate luminescence (lightness)
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    
    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        s = 0 'Saturation 0 for greyscale
        H = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation
        If L <= (HSLMAX / 2) Then
            s = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            s = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If
    
        'Calculate hue
        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus
    
        Select Case cMax
            Case CLng(R)
                H = BDelta - GDelta
            Case CLng(G)
                H = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(B)
                H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
        
        If H < 0 Then H = H + HSLMAX
    End If
    
    RGBtoHSL.Hue = CInt(H)
    RGBtoHSL.Lum = CInt(L)
    RGBtoHSL.Sat = CInt(s)
End Function

Public Function HSLtoRGB(HueLumSat As HSLCol) As Long '***
    Dim R As Double, G As Double, B As Double
    Dim H As Double, L As Double, s As Double
    Dim Magic1 As Double, Magic2 As Double
    
    H = HueLumSat.Hue
    L = HueLumSat.Lum
    s = HueLumSat.Sat
    
    If CInt(s) = 0 Then 'Greyscale
        R = (L * RGBMAX) / HSLMAX 'luminescence,
                'converted to the proper range
        G = R 'All RGB values same in greyscale
        B = R
        If CInt(H) <> UNDEFINED Then
            'This is technically an error.
            'The RGBtoHSL routine will always return
            'Hue = UNDEFINED (160 when HSLMAX is 240)
            'when Sat = 0.
            'if you are writing a color mixer and
            'letting the user input color values,
            'you may want to set Hue = UNDEFINED
            'in this case.
        End If
    Else
        'Get the "Magic Numbers"
        If L <= HSLMAX / 2 Then
            Magic2 = (L * (HSLMAX + s) + 0.5) / HSLMAX
        Else
            Magic2 = L + s - ((L * s) + 0.5) / HSLMAX
        End If
        
        Magic1 = 2 * L - Magic2
        
        'get R, G, B; change units from HSLMAX range
        'to RGBMAX range
        R = (HuetoRGB(Magic1, Magic2, H + (HSLMAX / 3)) _
            * RGBMAX + 0.5) / HSLMAX
        G = (HuetoRGB(Magic1, Magic2, H) * RGBMAX + 0.5) / HSLMAX
        B = (HuetoRGB(Magic1, Magic2, H - (HSLMAX / 3)) _
            * RGBMAX + 0.5) / HSLMAX
        
    End If
    
    HSLtoRGB = RGB(CInt(R), CInt(G), CInt(B))
    
End Function

Private Function HuetoRGB(mag1 As Double, mag2 As Double, _
    ByVal Hue As Double) As Double '***
'Utility function for HSLtoRGB

'Range check
    If Hue < 0 Then
        Hue = Hue + HSLMAX
    ElseIf Hue > HSLMAX Then
        Hue = Hue - HSLMAX
    End If
    
    'Return r, g, or b value from parameters
    Select Case Hue 'Values get progressively larger.
                'Only the first true condition will execute
        Case Is < (HSLMAX / 6)
            HuetoRGB = (mag1 + (((mag2 - mag1) * Hue + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Is < (HSLMAX / 2)
            HuetoRGB = mag2
        Case Is < (HSLMAX * 2 / 3)
            HuetoRGB = (mag1 + (((mag2 - mag1) * _
                ((HSLMAX * 2 / 3) - Hue) + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Else
            HuetoRGB = mag1
    End Select
End Function

Public Function ContrastingColor(RGBCol As Long) As Long
'Returns Black or White, whichever will show up better
'on the specified color.
'Useful for setting label forecolors with transparent
'backgrounds (send it the form backcolor - RGB value, not
'system value!)
'(also produces a monochrome negative when applied to
'all pixels in an image)

Dim HSL As HSLCol
    HSL = RGBtoHSL(RGBCol)
    If HSL.Lum > HSLMAX / 2 Then ContrastingColor = 0 _
        Else: ContrastingColor = &HFFFFFF
End Function

Public Function Brighten(RGBColor As Long, Percent As Single)
'Lightens the color by a specifie percent, given as a Single
'(10% = .10)

Dim HSL As HSLCol, L As Long
    If Percent <= 0 Then
        Brighten = RGBColor
        Exit Function
    End If
    
    HSL = RGBtoHSL(RGBColor)
    L = HSL.Lum + (HSLMAX * Percent)
    If L > HSLMAX Then L = HSLMAX
    HSL.Lum = L
    Brighten = HSLtoRGB(HSL)
End Function

Public Function Darken(RGBColor As Long, Percent As Single)
'Darkens the color by a specifie percent, given as a Single

Dim HSL As HSLCol, L As Long
    If Percent <= 0 Then
        Darken = RGBColor
        Exit Function
    End If
    
    HSL = RGBtoHSL(RGBColor)
    L = HSL.Lum - (HSLMAX * Percent)
    If L < 0 Then L = 0
    HSL.Lum = L
    Darken = HSLtoRGB(HSL)
End Function

Public Function Blend(RGB1 As Long, RGB2 As Long, _
    Percent As Single) As Long
'This one doesn't really use the HSL routines, just the
'RGB Component routines.  I threw it in as a bonus ;)
'Takes two colors and blends them according to a
'percentage given as a Single
'For example, .3 will return a color 30% of the way
'between the first color and the second.
'.5, or 50%, will be an even blend (halfway)
'Can create some nice effects inside a For loop

Dim R As Integer, R1 As Integer, R2 As Integer, _
    G As Integer, G1 As Integer, G2 As Integer, _
    B As Integer, B1 As Integer, B2 As Integer
    
    If Percent >= 1 Then
        Blend = RGB2
        Exit Function
    ElseIf Percent <= 0 Then
        Blend = RGB1
        Exit Function
    End If
    
    R1 = RGBRed(RGB1)
    R2 = RGBRed(RGB2)
    G1 = RGBGreen(RGB1)
    G2 = RGBGreen(RGB2)
    B1 = RGBBlue(RGB1)
    B2 = RGBBlue(RGB2)
    
    R = ((R2 * Percent) + (R1 * (1 - Percent)))
    G = ((G2 * Percent) + (G1 * (1 - Percent)))
    B = ((B2 * Percent) + (B1 * (1 - Percent)))
    
    Blend = RGB(R, G, B)
End Function


