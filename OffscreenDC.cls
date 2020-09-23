VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OffscreenDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Created by Jon Boyer Mooty 2001
' Comments, Suggestions, Harrasment send to youngbuck_66@yahoo.com

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBLT Lib "gdi32" Alias "StretchBlt" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal o As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Const FLOODFILLSURFACE = 1


Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10

Private Const HALFTONE = 4
Private Const COLORONCOLOR = 3

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

Private Const LOGPIXELSY = 90

'Uncomment the next 3 enums if using this class in another project

'Public Enum HorizAlign

    'HLeft
    'HCenter
    'HRight
    
 'End Enum
 
 'Public Enum VertAlign
 
    'VTop
    'VCenter
    'VBottom
    
 'End Enum

'Public Enum Rect3DStyle
    
    'None
    'Bump3D
    'Etched3D
    'Raised3D
    'Sunken3D
    
 'End Enum
 
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type
 
Private sfFont As New StdFont
Private lngDC As Long
Private lngOrigBMP As Long
Private lngHgt As Long, lngWdth As Long
Private lngOrigPen As Long
Private lngLineColor As Long
Private intLineWdth As Long
Private lngOrigBrsh As Long
Private lngFillColor As Long
Private lngBackColor As Long
Private lngOrigFont As Long
Private intFntAng As Integer

Private Sub Class_Initialize()
 Dim lBMP As Long
 Dim lPen As Long
 Dim lBrsh As Long
 Dim lFont As Long
    
    ' get a handle to a new Device that is compatible with the Desktop
    lngDC = CreateCompatibleDC(GetDC(GetDesktopWindow))
    
    If Not lngDC = 0 Then
        
        ' Get a handle to bitmap that is compatible with the desktop, initial size _
          of bitamp is 1 pixel by 1 pixel
        lBMP = CreateCompatibleBitmap(GetDC(GetDesktopWindow), 1, 1)
        
        If Not lBMP = 0 Then
            
            ' put the new bitmap into the DC and retrieve the handle to the original bitmap _
              created with DC and update the dimensions of are DC
            lngOrigBMP = SelectObject(lngDC, lBMP)
            lngHgt = 1: lngWdth = 1
            
        End If
        
        ' initialize the pen for the DC as a solid white pen with a drawing width of 1 pixel
        lPen = CreatePen(0, 1, vbWhite)
            
        If Not lPen = 0 Then
            
            ' put the pen into the DC and retrieve handle to the original pen and update _
              the classes local variable copies of the LineColor and LineWidth properties
            lngOrigPen = SelectObject(lngDC, lPen)
            lngLineColor = vbWhite
            intLineWdth = 1
            
        End If
        
        ' create a solid black brush
        lBrsh = CreateSolidBrush(vbBlack)
        
        If Not lBrsh = 0 Then
            
            ' put the brush into the DC and update the local variable
            lngOrigBrsh = SelectObject(lngDC, lBrsh)
            lngFillColor = vbWhite
            
        End If
        
        ' create a MS Sans Serif font with a 12 point font
        lFont = CreateFont(MulDiv(12, GetDeviceCaps(lngDC, 90), 72), 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, "MS Sans Serif")
        
        If Not lFont = 0 Then
            
            ' put the font into the DC and update the local copy of the stdFont object
            lngOrigFont = SelectObject(lngDC, lFont)
            sfFont.Name = "MS Sans Serif"
            sfFont.Size = 12
            
        End If
        
        ' init the  text color of the font to white
        SetTextColor lngDC, vbWhite
        
        ' init the backcolor to black and clear the jibberish off the bitmap
        lngBackColor = vbBlack
        Clear
        
    End If
    
End Sub

Private Sub Class_Terminate()

 
    'put all the original objects back into the DC and free all memory
    If Not lngDC = 0 Then
 
        If Not lngOrigBMP = 0 Then
        
            DeleteObject SelectObject(lngDC, lngOrigBMP)
        
        End If
    
        If Not lngOrigPen = 0 Then
        
            DeleteObject SelectObject(lngDC, lngOrigPen)
        
        End If
    
        If Not lngOrigBrsh = 0 Then
        
            DeleteObject SelectObject(lngDC, lngOrigBrsh)
        
        End If
    
        If Not lngOrigFont = 0 Then
            
            DeleteObject SelectObject(lngDC, lngOrigFont)
             
        End If
        
        ' delete the main DC from memory
        DeleteDC lngDC
    
    End If
    
    Set sfFont = Nothing
    
End Sub

Public Property Get hdc() As Long
    
    hdc = lngDC

End Property

Public Property Get ForeColor() As Long
    
    ForeColor = GetTextColor(lngDC)
    
End Property

Public Property Let ForeColor(ByVal pForeColor As Long)
    
    ' if the value passed into pForeColor is a system color than convert it _
      into a proper long value
    If IsSystemColor(pForeColor) Then
    
        SetTextColor lngDC, SystemToRGB(pForeColor)
    
    Else
    
        SetTextColor lngDC, pForeColor
        
    End If
    
End Property

Public Property Get BackColor() As OLE_COLOR
    
    BackColor = lngBackColor
    
End Property

Public Property Let BackColor(ByVal pBackColor As OLE_COLOR)
    
    ' if the value passed into pForeColor is a system color than convert it _
      into a proper long value
    If IsSystemColor(pBackColor) Then
    
        lngBackColor = SystemToRGB(pBackColor)
        
    Else
        
        lngBackColor = pBackColor
    
    End If
    
    Clear
    
End Property

Public Property Get LineColor() As Long
 
    LineColor = lngLineColor
    
End Property

Public Property Let LineColor(ByVal pLineColor As Long)
 Dim lPen As Long
 Dim lPrevPen As Long
    
    ' if the value passed into pForeColor is a system color than convert it _
      into a proper long value
    If IsSystemColor(pLineColor) Then
    
        lPen = CreatePen(0, intLineWdth, SystemToRGB(pLineColor))
    
    Else
    
        lPen = CreatePen(0, intLineWdth, pLineColor)
        
    End If
   
    If Not lPen = 0 Then
 
        lPrevPen = SelectObject(lngDC, lPen)
        
        ' remove the previous pen from memory
        DeleteObject lPrevPen
        
        ' if the passed color is a system color than convert to a RGB color _
          so that it can displayed properly
        If IsSystemColor(pLineColor) Then
            
            lngLineColor = SystemToRGB(pLineColor)
            
        Else
            
            lngLineColor = pLineColor
            
        End If
        
    End If
    
End Property

Public Property Get LineWidth() As Integer
    
    LineWidth = intLineWdth
    
End Property

Public Property Let LineWidth(ByVal pLineWidth As Integer)
    
    intLineWdth = pLineWidth
    
    LineColor = lngLineColor
    
End Property

Public Sub BltTo(ByVal pDestDC As Long, Optional ByVal pDestX As Long = 0, Optional ByVal pDestY As Long = 0, Optional ByVal pSrcX As Long = 0, Optional ByVal pSrcY As Long = 0, Optional ByVal pSrcH As Long = 0, Optional ByVal pSrcW As Long = 0, Optional ByVal pRaster As Long = vbSrcCopy)
    
    ' default the height of the Blt to the height of the main DC if nothing is passed
    If pSrcH = 0 Then pSrcH = lngHgt
    If pSrcW = 0 Then pSrcW = lngWdth
    
    ' copy the image of the DC to the Destination DC
    BitBlt pDestDC, pDestX, pDestY, pSrcW, pSrcH, lngDC, pSrcX, pSrcY, pRaster
    
End Sub

Public Property Get Height() As Long

    Height = lngHgt
    
End Property

Public Property Get Width() As Long
    
    Width = lngWdth
    
End Property

Public Sub Resize(ByVal pHeight As Long, ByVal pWidth As Long, Optional ByVal PreserveImage As Boolean = False)
 Dim lNewBMP As Long, lPrevBMP As Long
 Dim osdc As OffscreenDC
    
    ' create a new bitmap compatible with the desktop using size requested
    lNewBMP = CreateCompatibleBitmap(GetDC(GetDesktopWindow), pWidth, pHeight)
    
    ' if requested to preserve the current image of the DC then _
      copy the current image into a temporary OffscreenDC object
    If PreserveImage = True Then
        
        Set osdc = New OffscreenDC
        
        osdc.Resize Height, Width
        
        osdc.Clear
        
        BltTo osdc.hdc
    
    End If
    
    If Not lNewBMP = 0 Then
        
        ' change the local width and height
        lngWdth = pWidth
        lngHgt = pHeight
        
        lPrevBMP = SelectObject(lngDC, lNewBMP)
        
        'remove the previous bitmap from memory
        DeleteObject lPrevBMP
        
        If PreserveImage = True Then
            
            ' copy the previous image onto the new bitmap
            osdc.BltTo lngDC, 0, 0, 0, 0
            
            Set osdc = Nothing
            
        Else
            
            ' clear the DC to the current background color
            Clear
            
        End If
        
    End If
    
End Sub

Public Sub Clear()
 Dim lBrush As Long
 Dim rctBG As RECT
 
 ' clear the DC by drawing a rectangle over the entire size of the bitmap _
   and filling it with the current backcolor
 
 With rctBG
    .Left = 0
    .Top = 0
    .Right = lngWdth
    .Bottom = lngHgt
 End With
 
 lBrush = CreateSolidBrush(lngBackColor)
    
 FillRect lngDC, rctBG, lBrush
 
 DeleteObject lBrush
 
End Sub

Public Sub LoadSavedImage(ByVal pPath As String, Optional ByVal pSizeToFit As Boolean = False)
 ' if the path exists then put the image requested into the DC
 If Not Len(Dir(pPath)) = 0 Then
  Dim iPic As IPictureDisp
  Dim lPrevBMP As Long
  Dim lPrevH As Long
  Dim lPrevW As Long
  
  ' record the previous height and width of the DC in case resizing is _
    necessary
  lPrevH = lngHgt
  lPrevW = lngWdth
  
  ' load the picture file
  Set iPic = LoadPicture(pPath)
  
  ' select the bitmap into the DC
  lPrevBMP = SelectObject(lngDC, iPic)
  
  ' remove the previous bitmap
   DeleteObject lPrevBMP
  
   If pSizeToFit = False Then
   
    ' resize the DC and preserve the image contained
    Resize lPrevH, lPrevW, True
    
  Else
    
    ' set the local height and width to that of the picture loaded _
      convert from HiMetrics to Pixels
    lngHgt = HiMtoPix(iPic.Height)
    lngWdth = HiMtoPix(iPic.Width)
    
  End If
   
End If

Set iPic = Nothing

End Sub

Public Sub StretchTo(ByVal pDestDC As Long, ByVal pDestH As Long, ByVal pDestW As Long, Optional ByVal pDestX As Long = 0, Optional ByVal pDestY As Long = 0)
    
    ' set the stretchblt mode depending on dimensions of the destination
    If lngWdth > pDestW Or lngHgt > pDestH Then
        
            SetStretchBltMode pDestDC, HALFTONE
            
    Else
         
            SetStretchBltMode pDestDC, COLORONCOLOR
            
    End If
    
    ' stretch the image to the requested DC
    StretchBLT pDestDC, pDestX, pDestY, pDestW, pDestH, lngDC, 0, 0, lngWdth, lngHgt, vbSrcCopy
    
End Sub

Public Sub DrawRect(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, Optional ByVal pTransparent As Boolean = False)
        
    If Not pTransparent Then
        
        ' rectangle will be drawn and filled based on the current pen & brush
        Rectangle lngDC, pX1, pY1, pX2, pY2
        
    Else
     Dim ptRct(1 To 5) As POINTAPI
        
        ' just draw the outline of the rectangle using the current pen
        ptRct(1).x = pX1: ptRct(1).y = pY1
        ptRct(2).x = pX2: ptRct(2).y = pY1
        ptRct(3).x = pX2: ptRct(3).y = pY2
        ptRct(4).x = pX1: ptRct(4).y = pY2
        ptRct(5).x = pX1: ptRct(5).y = pY1
        
        Polyline lngDC, ptRct(1), 5
        
     End If
     
End Sub

Public Property Get FillColor() As Long
    
    FillColor = lngFillColor
    
End Property

Public Property Let FillColor(ByVal pFillColor As Long)
 Dim lBrsh As Long
    
    ' if the value passed into pForeColor is a system color than convert it _
      into a proper long value
    If IsSystemColor(pFillColor) Then
    
        lBrsh = CreateSolidBrush(SystemToRGB(pFillColor))
        
    Else
     
        lBrsh = CreateSolidBrush(pFillColor)
        
    End If
    
    If Not lBrsh = 0 Then
    
        ' select the new brush and remove the old
        DeleteObject SelectObject(lngDC, lBrsh)
        
        If IsSystemColor(pFillColor) Then
            
            lngFillColor = SystemToRGB(pFillColor)
        
        Else
            
            lngFillColor = pFillColor
        
        End If
        
    End If
    
End Property

Public Sub DrawLine(ByVal pXFrom As Long, ByVal pYFrom As Long, ByVal pXTo As Long, ByVal pYTo As Long)
 Dim ptLn(1 To 2) As POINTAPI
 
 ' define the points of the line to be drawn
 ptLn(1).x = pXFrom: ptLn(1).y = pYFrom
 ptLn(2).x = pXTo: ptLn(2).y = pYTo
 
 ' draw the line
 Polyline lngDC, ptLn(1), 2
 
End Sub

Public Sub DrawPoly(pXs() As Long, pYs() As Long)
 Dim iCt As Long, iPtCt As Long
 Dim ptPoly() As POINTAPI
 
 ' allocate the proper number of points for the polygon
 ReDim ptPoly(LBound(pXs) To UBound(pXs))
 
 ' assign the coordinates  of the polygon the POINTAPI structures
 For iCt = LBound(ptPoly) To UBound(ptPoly)
    
    iPtCt = iPtCt + 1
    
    ptPoly(iCt).x = pXs(iCt): ptPoly(iCt).y = pYs(iCt)
 
 Next iCt
 
 ' draw the polygon
 Polygon lngDC, ptPoly(LBound(ptPoly)), iPtCt
 
End Sub

Public Property Get Font() As StdFont
    
    Set Font = sfFont
    
    ChangeFont
    
End Property

Public Property Set Font(pFont As StdFont)

    Set sfFont = pFont
    
    ChangeFont
    
End Property

Public Property Get FontAngle() As Integer
    
    FontAngle = intFntAng
    
End Property

Public Property Let FontAngle(ByVal pAngle As Integer)
    
    ' only set the angle of the font if it is between 0 and 360
    If Not (pAngle > 360 Or pAngle < 0) Then
        
        intFntAng = pAngle * 10
        
        ' update the font
        ChangeFont
        
    End If
    
End Property

Private Sub ChangeFont()
 Dim lFont As Long
    
    ' create the font using the StdFont object, the MulDiv call is required _
      to convert the font size from points to pixels
    lFont = CreateFont(-MulDiv(sfFont.Size, GetDeviceCaps(lngDC, LOGPIXELSY), 72), 0, intFntAng, 0, sfFont.Weight, CLng(sfFont.Italic), CLng(sfFont.Underline), CLng(sfFont.Strikethrough), sfFont.Charset, 0, 0, 0, 0, sfFont.Name)
    
    ' select the new font to the main DC and remove the old one
    DeleteObject SelectObject(lngDC, lFont)
    
End Sub

Public Sub DrawTxt(ByVal pS As String, ByVal pX As Long, ByVal pY As Long, Optional ByVal pTransparent As Boolean = False)
 Dim lPrevBkMode As Long
    
    ' record the previous level of transparency
     If pTransparent = True Then
     
        lPrevBkMode = SetBkMode(lngDC, TRANSPARENT)
    
    Else
        
        lPrevBkMode = SetBkMode(lngDC, OPAQUE)
        
    End If
    
    ' draw the string
    TextOut lngDC, pX, pY, pS, Len(pS)
    
    ' set the DC back to its previous level of transparency
    SetBkMode lngDC, lPrevBkMode
    
End Sub

Public Sub Draw3DEdge(ByVal pX1 As Long, ByVal pY1 As Long, ByVal pX2 As Long, ByVal pY2 As Long, ByVal p3DEdge As Rect3DStyle)
Dim rctC As RECT

' define a rectangle the size of the current DC
With rctC
    
    .Left = pX1
    .Right = pX2
    .Top = pY1
    .Bottom = pY2
    
 End With

' Draw the 3D edge with the appropriate effect
Select Case p3DEdge
 
    Case Bump3D
    
        DrawEdge lngDC, rctC, EDGE_BUMP, BF_RECT
        
    Case Etched3D
        
        DrawEdge lngDC, rctC, EDGE_ETCHED, BF_RECT
        
    Case Raised3D
    
        DrawEdge lngDC, rctC, EDGE_RAISED, BF_RECT
        
    Case Sunken3D
    
        DrawEdge lngDC, rctC, EDGE_SUNKEN, BF_RECT
        
 End Select
    
End Sub

Public Sub DrawCircle(ByVal pRadius As Long, ByVal pCenterX As Long, ByVal pCenterY As Long, Optional ByVal pTransparent As Boolean = False)
If Not pTransparent Then
    
    'Use Ellipse function to draw filled circle
    Ellipse lngDC, _
                          pCenterX + pRadius, _
                           pCenterY + pRadius, _
                           pCenterX - pRadius, _
                           pCenterY - pRadius
    
Else
    
    ' use arc function so circle is not filled with current brush
    Arc lngDC, pCenterX - pRadius, pCenterY - pRadius, _
                        pCenterX + pRadius, pCenterY + pRadius, _
                        pCenterX + pRadius, pCenterY - pRadius, _
                        pCenterX + pRadius, pCenterY - pRadius
                        
End If
End Sub

Public Sub DrawAlignedTxt(ByVal pS As String, Optional ByVal haAlign As HorizAlign = HLeft, Optional ByVal vaAlign As VertAlign = VTop, Optional ByVal pTransparent As Boolean = True)
 Dim rctTxt As RECT
 Dim lFormat As Long
 Dim lPrevBkMode As Long
    
    ' record the main DC previous level of transparency
    If pTransparent = True Then
        
        lPrevBkMode = SetBkMode(lngDC, TRANSPARENT)
        
    Else
        
        lPrevBkMode = SetBkMode(lngDC, OPAQUE)
        
    End If
    
    ' set the format for drawing the text
    lFormat = DT_SINGLELINE
    
    ' set the horizontal align bits
    Select Case haAlign
        
        Case HLeft
            lFormat = lFormat Or DT_LEFT
        Case HCenter
            lFormat = lFormat Or DT_CENTER
        Case HRight
            lFormat = lFormat Or DT_RIGHT
    
    End Select
    
    ' set the vertical alignment bits
    Select Case vaAlign
        
        Case VTop
            lFormat = lFormat Or DT_TOP
        Case VCenter
            lFormat = lFormat Or DT_VCENTER
        Case VBottom
            lFormat = lFormat Or DT_BOTTOM
    
     End Select
    
    ' define a rectangle the size of the current DC with a 2 pixel buffer _
      in case there is a border drawn to the DC
    With rctTxt
    
        .Left = 2
        .Top = 2
        .Right = lngWdth - 2
        .Bottom = lngHgt - 2
    
    End With
    
    ' Draw the text aligned
    DrawTextEx lngDC, pS, Len(pS), rctTxt, lFormat, ByVal CLng(0)
    
    ' set the transparency level to its previous value
    SetBkMode lngDC, lPrevBkMode
    
End Sub

Public Function StringWidth(ByVal pS As String) As Long
Dim ptSize As POINTAPI

 ' return the width in pixels of the string passsed based on the current _
   font settings
 GetTextExtentPoint32 lngDC, pS, Len(pS), ptSize
 
 StringWidth = ptSize.x
 
End Function

Public Function StringHeight(ByVal pS As String) As Long
Dim ptSize As POINTAPI
 
  ' return the height in pixels of the string passsed based on the current _
   font settings
 GetTextExtentPoint32 lngDC, pS, Len(pS), ptSize
 
 StringHeight = ptSize.y
 
End Function

Public Sub LoadStdPicture(ByVal pPicture As StdPicture, Optional ByVal pSizeToFit As Boolean = False)

    If pSizeToFit = True Then
    
        ' resize the DC so the picture will fit
        Resize HiMtoPix(pPicture.Height), HiMtoPix(pPicture.Width)
        
    End If
    
    ' select the new bitmap into the DC and remove the old one
    DeleteObject SelectObject(lngDC, pPicture.Handle)
    
End Sub

Public Sub MaskTo(ByVal phDC As Long, pTranspColor As Long)
 Dim lX As Long, lY As Long
 
 ' Not the most efficient algorithm, but it works
 ' Colors every pixel not equal to the Transparent color white and _
   every color that is equal to the Transparent color black for use _
   in transparent Bltting
 For lX = 0 To lngWdth
    For lY = 0 To lngHgt
        
        If GetPixel(lngDC, lX, lY) = pTranspColor Then
            
            SetPixelV phDC, lX, lY, vbBlack
            
        Else
            
            SetPixelV phDC, lX, lY, vbWhite
            
        End If
        
    Next lY
 Next lX
 
End Sub

Public Sub SetPoint(ByVal pX As Long, ByVal pY As Long, pColor As Long)
    
    SetPixelV lngDC, pX, pY, pColor
    
End Sub

Public Function GetPoint(ByVal pX As Long, pY As Long) As Long
    
    GetPoint = GetPixel(lngDC, pX, pY)
    
End Function

Private Function HiMtoPix(ByVal pHiM As Long) As Long
    
    ' convert HiMetrics to Pixels
    HiMtoPix = pHiM / 26.45703125
    
End Function

Private Function IsSystemColor(ByVal pColor As Long) As Boolean
    
    ' no system colors are higher than the value of vbInfoBackground and _
      neither are any RGB values therefore any thing less than or equal to _
      vbInfoBackground is a System color
    If pColor <= vbInfoBackground Then
        
        IsSystemColor = True
        
    Else
    
        IsSystemColor = False
        
    End If
    
End Function

Private Function SystemToRGB(ByVal pColConst As Long) As Long
    
    ' retrive the system color
    SystemToRGB = GetSysColor(pColConst + 2147483648#)
    
End Function

Public Sub FloodPt(pX As Long, pY As Long)
    
    ' Fill with value of FillColor while Color at pX, pY is found
    ExtFloodFill lngDC, pX, pY, GetPoint(pX, pY), FLOODFILLSURFACE

End Sub
