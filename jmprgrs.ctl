VERSION 5.00
Begin VB.UserControl JMProgress 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "jmprgrs.ctx":0000
   Begin VB.Timer tmrRefresh 
      Left            =   240
      Top             =   2880
   End
End
Attribute VB_Name = "JMProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Mouse_Enter Lib "user32" Alias "SetCapture" (ByVal hwnd As Long) As Long
Private Declare Function Mouse_Leave Lib "user32" Alias "ReleaseCapture" () As Long
 
Private Const DARK_PURPLE As Long = &H800040

Public Enum HorizAlign

    HLeft
    HCenter
    HRight
    
 End Enum
 
 Public Enum VertAlign
 
    VTop
    VCenter
    VBottom
    
 End Enum
 
Public Enum Rect3DStyle
    
    None
    Bump3D
    Etched3D
    Raised3D
    Sunken3D
    
 End Enum
 
  
 Public Enum ProgressStyle

    apbClassic
    apbGradient
    apbImage
    
End Enum

Public Enum ProgressDirection

    apbLeftToRight
    apbRightToLeft
    apbBottomToTop
    apbTopToBottom
    apbFromCenterHorizontal
    apbFromCenterVertical
    
End Enum


Public Event ValChange(ByVal lngValue As Long, ByVal dblPercentage As Double)
Public Event ProgressComplete()
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseEnter()
Public Event MouseExit()

Private offUC As New OffscreenDC
Private offBB As New OffscreenDC
Private offImg As New OffscreenDC
Private ocBackColor As OLE_COLOR
Private ocForeColor As OLE_COLOR
Private ocFillColor As OLE_COLOR
Private ocGradS As OLE_COLOR
Private ocGradE As OLE_COLOR
Private r3dStyle As Rect3DStyle
Private apbStyle As ProgressStyle
Private apbDirection As ProgressDirection
Private stdProgPic As StdPicture
Private haAlignment As HorizAlign
Private vaAlignment As VertAlign
Private strCaption As String
Private boolOBErr As Boolean
Private lngMin As Long
Private lngMax As Long
Private lngVal As Long

Private Sub tmrRefresh_Timer()
    ' used for redrawing the control after the Font property is accessed _
      through its Get property
    
    DrawControl
    
    ' only run this sub once
    tmrRefresh.Interval = 0
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
        ' once mouse is captured you will continue to receive mouse events even _
          when the pointer exits the control therefore you can raise the MouseExit event _
          as soon as the coordinates are less than 0 or greater than the dimensions of the control
        If (x < 0) Or (y < 0) Or (x > UserControl.Width) Or (y > UserControl.Height) Then
            
            RaiseEvent MouseExit
            Mouse_Leave
        
        Else
            
            RaiseEvent MouseEnter
            Mouse_Enter UserControl.hwnd
        
        End If
        
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Initialize()
    
    DrawControl
    
End Sub

Private Sub UserControl_InitProperties()
    
    AlignmentHoriz = HCenter
    AlignmentVert = VCenter
    BackColor = UserControl.Parent.BackColor ' initialize to owners BackColor
    BorderStyle = Sunken3D
    Caption = UserControl.Name ' initialize Caption to name of control
    FillColor = vbHighlight
    Set Font = UserControl.Parent.Font ' initialize font ot owners Font
    ForeColor = UserControl.Parent.ForeColor ' initialize ForeColor to owners ForeColor
    GradientEnd = DARK_PURPLE
    GradientStart = vbYellow
    Max = 100
    Min = 1
    OutOfBoundsError = False
    ProgressDirection = apbLeftToRight
    ProgressStyle = apbClassic
    Value = 100
    
End Sub

Private Sub UserControl_Paint()

    DrawControl
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    ' Reload state
    AlignmentHoriz = PropBag.ReadProperty("AlignmentHoriz", HCenter)
    AlignmentVert = PropBag.ReadProperty("AlignmentVert", VCenter)
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    BorderStyle = PropBag.ReadProperty("BorderStyle", Sunken3D)
    Caption = PropBag.ReadProperty("Caption", "")
    FillColor = PropBag.ReadProperty("FillColor", vbBlack)
    Font.Name = PropBag.ReadProperty("FontName", "MS Sans Serif")
    Font.Size = PropBag.ReadProperty("FontSize", 8)
    Font.Bold = PropBag.ReadProperty("FontBold", False)
    Font.Italic = PropBag.ReadProperty("FontItalic", False)
    Font.Strikethrough = PropBag.ReadProperty("FontStrikethrough", False)
    Font.Underline = PropBag.ReadProperty("FontUnderline", False)
    ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    GradientEnd = PropBag.ReadProperty("GradientEnd", vbBlue)
    GradientStart = PropBag.ReadProperty("GradientStart", vbRed)
    Max = PropBag.ReadProperty("Max", 100)
    Min = PropBag.ReadProperty("Min", 1)
    OutOfBoundsError = PropBag.ReadProperty("OutOfBoundsError", False)
    ProgressDirection = PropBag.ReadProperty("ProgressDirection", apbLeftToRight)
    ProgressStyle = PropBag.ReadProperty("ProgressStyle", apbClassic)
    Value = Max
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' save state
    PropBag.WriteProperty "AlignmentHoriz", haAlignment, HCenter
    PropBag.WriteProperty "AlignmentVert", vaAlignment, VCenter
    PropBag.WriteProperty "BackColor", ocBackColor, vbButtonFace
    PropBag.WriteProperty "BorderStyle", r3dStyle, Sunken3D
    PropBag.WriteProperty "Caption", strCaption, ""
    PropBag.WriteProperty "FillColor", ocFillColor, vbBlack
    PropBag.WriteProperty "FontName", UserControl.Font.Name, "MS Sans Serif"
    PropBag.WriteProperty "FontSize", UserControl.Font.Size, 8
    PropBag.WriteProperty "FontBold", UserControl.Font.Bold, False
    PropBag.WriteProperty "FontItalic", UserControl.Font.Italic, False
    PropBag.WriteProperty "FontStrikethrough", UserControl.Font.Strikethrough, False
    PropBag.WriteProperty "FontUnderline", UserControl.Font.Underline, False
    PropBag.WriteProperty "ForeColor", ocForeColor, vbButtonText
    PropBag.WriteProperty "GradientEnd", ocGradE, vbBlue
    PropBag.WriteProperty "GradientStart", ocGradS, vbRed
    PropBag.WriteProperty "Max", lngMax, 100
    PropBag.WriteProperty "Min", lngMin, 1
    PropBag.WriteProperty "OutOfBoundsError", boolOBErr, False
    PropBag.WriteProperty "ProgressDirection", apbDirection, apbLeftToRight
    PropBag.WriteProperty "ProgressStyle", apbStyle, apbClassic
    
End Sub

Private Sub UserControl_Resize()
    
    ' resize the backbuffer and the control DC
    offBB.Resize UserControl.Height / Screen.TwipsPerPixelY, UserControl.Width / Screen.TwipsPerPixelX, False
    offUC.Resize offBB.Height, offBB.Width
    
    ' update the contents of the backbuffer and display the changes
    UpdateBB
    DrawControl
    
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the Background color of the control."
    
    BackColor = ocBackColor
    
End Property

Public Property Let BackColor(ByVal pBackColor As OLE_COLOR)
    
        offBB.BackColor = pBackColor
        offUC.BackColor = pBackColor
        offImg.BackColor = pBackColor
        
        ' store value locally in case color is a system color
        ocBackColor = pBackColor
    
         UpdateBB
         DrawControl
    
End Property

Private Sub UserControl_Terminate()

    'Free memory
    Set offBB = Nothing
    Set offUC = Nothing
    Set offImg = Nothing
    
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns or sets the color used to draw the Caption text."
    
    ForeColor = ocForeColor
    
End Property

Public Property Let ForeColor(ByVal pForeColor As OLE_COLOR)
    
    offBB.ForeColor = pForeColor
    offUC.ForeColor = pForeColor
    
    'store value locally in case value is system color
    ocForeColor = pForeColor
    
    ' no need to update display unless there is a caption to display
    If Not strCaption = "" Then DrawControl
    
End Property

Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns or sets the color used to fill shapes."
        
    FillColor = ocFillColor
    
End Property

Public Property Let FillColor(ByVal pFillColor As OLE_COLOR)
        
        offBB.FillColor = pFillColor
        offBB.LineColor = pFillColor
        
        ' store value locally in case value is system color
        ocFillColor = pFillColor
        
        ' only update display if fillcolor is used for current style
        If apbStyle = apbClassic Then UpdateBB: DrawControl
        
End Property

Public Property Get GradientStart() As OLE_COLOR
Attribute GradientStart.VB_Description = "Returns or sets the start color of the gradient fill."

    GradientStart = ocGradS
    
End Property

Public Property Let GradientStart(ByVal pGradS As OLE_COLOR)
    
    If Not pGradS = ocGradS Then
    
        ocGradS = pGradS
    
        ' only update the display if gradient style is selected
        If apbStyle = apbGradient Then UpdateBB: DrawControl
    
    End If
    
End Property

Public Property Get GradientEnd() As OLE_COLOR
Attribute GradientEnd.VB_Description = "Returns or sets the end color of the gradient fill."

    GradientEnd = ocGradE

End Property

Public Property Let GradientEnd(ByVal pGradE As OLE_COLOR)
    
    If Not pGradE = ocGradE Then
    
        ocGradE = pGradE
    
        'only update the display if gradient is selected
        If apbStyle = apbGradient Then UpdateBB: DrawControl
        
    End If
    
End Property

Public Property Get BorderStyle() As Rect3DStyle
    
    BorderStyle = r3dStyle
    
End Property

Public Property Let BorderStyle(ByVal pBorderStyle As Rect3DStyle)
    
    If Not r3dStyle = pBorderStyle Then
        
        ' proceed only if value is within bounds of Enum
        If r3dStyle >= None And r3dStyle <= Sunken3D Then
        
            r3dStyle = pBorderStyle
            
            ' update the control with new Border
            DrawControl
            
        End If
        
    End If
    
End Property

Public Property Get ProgressStyle() As ProgressStyle
Attribute ProgressStyle.VB_Description = "Returns or sets the style of the progress bar."
    
    ProgressStyle = apbStyle
    
End Property

Public Property Let ProgressStyle(ByVal pStyle As ProgressStyle)
    
    If Not pStyle = apbStyle Then
    
        apbStyle = pStyle
        
        ' update the Control and Display it
        UpdateBB
        DrawControl
    
    End If
    
End Property

Public Property Get AlignmentHoriz() As HorizAlign
Attribute AlignmentHoriz.VB_Description = "Determines the postitioning of the Caption text."
    
    AlignmentHoriz = haAlignment
    
End Property

Public Property Let AlignmentHoriz(ByVal pHAlign As HorizAlign)
    
    If Not pHAlign = haAlignment Then
    
        haAlignment = pHAlign
        
        ' update the display only if there is a caption to display
        If Not strCaption = "" Then DrawControl
        
    End If
    
End Property

Public Property Get AlignmentVert() As VertAlign
    
    AlignmentVert = vaAlignment
    
End Property

Public Property Let AlignmentVert(ByVal pVAlign As VertAlign)
    
    If Not pVAlign = vaAlignment Then
        
        vaAlignment = pVAlign
        
        ' redraw control only if there is a caption to display
        If Not strCaption = "" Then DrawControl
        
    End If
        
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns or sets the text displayed in the control."
    
    Caption = strCaption
    
End Property

Public Property Let Caption(ByVal pCaption As String)
    
    If Not pCaption = strCaption Then
    
        strCaption = pCaption
        
        ' redraw the control with the new caption
        DrawControl
    
    End If
    
End Property

Private Sub UpdateBB()

' Fill the backbuffer based on user's selection

    Select Case apbStyle
    
        Case apbGradient
            FillBBGrad
            
        Case apbImage
            FillBBImage
        
        Case apbClassic
            FillBBRects
        
     End Select
     
End Sub

Private Sub FillBBGrad()
 Dim iCt As Long
 Dim lIter As Long, lLoopCt As Long
 Dim lStart As Long, lEnd As Long, iStep As Integer
 Dim nRS As Integer, nGS As Integer, nBS As Integer
 Dim nRE As Integer, nGE As Integer, nBE As Integer
 Dim nRChg As Single, nGChg As Single, nBChg As Single
 Dim ocOrigLineCol As OLE_COLOR
 
 offBB.Clear
 
 ocOrigLineCol = offBB.LineColor
 
 ' system color constants are converted to long color values by the offscreen DC class _
   so pass the value of the Gradient color to the class to be converted and then retrieve _
   their red, green, and blue values based on the value that the class returns
 offBB.LineColor = ocGradS
 nRS = GetR(offBB.LineColor): nGS = GetG(offBB.LineColor): nBS = GetB(offBB.LineColor)
 offBB.LineColor = ocGradE
 nRE = GetR(offBB.LineColor): nGE = GetG(offBB.LineColor): nBE = GetB(offBB.LineColor)

 ' set the start, end, step, and number of iterations to the correct values to get the desired gradient effect
Select Case apbDirection
    
     Case apbFromCenterHorizontal, apbLeftToRight
    
        lEnd = offBB.Width
        lStart = 0
        iStep = 1
        lIter = offBB.Width
    
    Case apbBottomToTop
    
        lEnd = 0
        lStart = offBB.Height
        iStep = -1
        lIter = offBB.Height
    
    Case apbTopToBottom, apbFromCenterVertical
    
        lEnd = offBB.Height
        lStart = 0
        iStep = 1
        lIter = offBB.Height
    
    Case apbRightToLeft
    
        lEnd = 0
        lStart = offBB.Width
        iStep = -1
        lIter = offBB.Width
    
End Select
    
 ' retrieve the amount of change
 nRChg = (nRE - nRS) / lIter
 nGChg = (nGE - nGS) / lIter
 nBChg = (nBE - nBS) / lIter
 
 For iCt = lStart To lEnd Step iStep
    
    ' update the color of the line to be drawn
    offBB.LineColor = RGB(nRS + (nRChg * lLoopCt), nGS + (nGChg * lLoopCt), nBS + (nBChg * lLoopCt))
    
    If apbDirection = apbBottomToTop Or apbDirection = apbTopToBottom Or apbDirection = apbFromCenterVertical Then

        offBB.DrawLine 0, iCt, offBB.Width, iCt
        
    Else
        
        offBB.DrawLine iCt, 0, iCt, offBB.Height
    
    End If
    
    lLoopCt = lLoopCt + 1
 Next iCt
 
 'restore the original line color
 offBB.LineColor = ocOrigLineCol
    
End Sub

Private Sub FillBBImage()
Static lPrevBMP As Long

    If Not IsStdPicNothing(stdProgPic) Then
        
        If Not lPrevBMP = stdProgPic.Handle Then
    
            ' load the image into the temporary offscreen DC and resize the DC _
            to the size of the picture
            offImg.Resize stdProgPic.Height / 26.45703125, stdProgPic.Width / 26.45703125, True
    
            offImg.LoadStdPicture stdProgPic
        
            lPrevBMP = stdProgPic.Handle
        
        End If
        
    Else
    
        offImg.Clear
        
    End If
 
  ' stretch the picture to fit inside the backbuffer
 offImg.StretchTo offBB.hdc, offBB.Height, offBB.Width
 
End Sub

Private Sub FillBBRects() ' classic style
 Dim iRectSz As Integer
 Dim lX1 As Long
 Dim lY1 As Long
 
 offBB.Clear
 
 ' get a size to draw the rects based on the dimensions of the control
 iRectSz = offBB.Width / 20 + offBB.Height / 5
 
 ' initialize the x and y values to be offset 3 pixels from the side and top
 lX1 = 3
 lY1 = 3
 
 Do Until lX1 > offBB.Width Or lY1 > offBB.Height
        
    ' if rects need to be drawn vertically draw them horizontally
    If apbDirection = apbBottomToTop Or apbDirection = apbTopToBottom Then
        
        offBB.DrawRect lX1, lY1, offBB.Width - 3, lY1 + iRectSz
        
        lY1 = lY1 + iRectSz + 3
        
    Else
        
        offBB.DrawRect lX1, 3, lX1 + iRectSz, offBB.Height - 4
        
        lX1 = lX1 + iRectSz + 3
        
    End If
    
 Loop

End Sub

Private Function GetR(ByVal pCol As Long) As Byte
   
   'return the 'red' amount in the Long value
   GetR = pCol Mod 256

End Function

Private Function GetG(ByVal pCol As Long) As Byte
    
    'return the amount of 'Green' in the Long value
   pCol = pCol \ 256
   GetG = pCol Mod 256
    
End Function

Private Function GetB(ByVal pCol As Long) As Byte

    'return the amount of 'Blue' in the Long value
    pCol = (pCol \ 256) \ 256
    GetB = pCol Mod 256

End Function

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns or sets the minimum progress value."
    
   Min = lngMin
    
End Property

Public Property Let Min(ByVal pMin As Long)
    
    If pMin >= lngMax Or pMin <= 0 Then
        
        ' don't allow min to be more than max or less than 0
        lngMin = lngMax - 1
        
    Else
        
        lngMin = pMin
        
    End If
    
    If lngVal < lngMin Then lngVal = lngMin

    DrawControl
    
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns or sets the maximum progress value."
    
   Max = lngMax
    
End Property

Public Property Let Max(ByVal pMax As Long)
        
    If pMax <= lngMin Or pMax <= 0 Then
        
        ' don't allow max to be less than min or less than 0
        lngMax = lngMin + 1
        
    Else
        
        lngMax = pMax
        
    End If
    
    If lngVal > lngMax Then lngVal = lngMax
    
    DrawControl
    
End Property

Public Property Get Value() As Long
    
    Value = lngVal
    
End Property

Public Property Let Value(ByVal pValue As Long)
Attribute Value.VB_Description = "Returns or sets the current state of progress."
 Dim iPrevVal As Long
 
    If pValue < lngMin Then
        
        If boolOBErr = True Then
            
            Err.Raise vbObjectError + 66, UserControl.Name & ".Value", "Value out of bounds of Min and Max properties."
        
        Else
        
            lngVal = lngMin
            
        End If
        
     ElseIf pValue > lngMax Then
        
        If boolOBErr = True Then
            
            Err.Raise vbObjectError + 66, UserControl.Name & ".Value", "Value out of bounds of Min and Max properties."
        
        Else
        
            lngVal = lngMax
            
        End If
        
     Else
        
        lngVal = pValue
    
    End If
    
    
        If Not lngVal = iPrevVal Then RaiseEvent ValChange(lngVal, ProgressPercentage)
        If lngVal = lngMax Then RaiseEvent ProgressComplete
    
    DrawControl
    
End Property

Private Sub DrawControl()
 Dim iProgressW As Long
 Dim lStrW As Long, lStrH As Long
    
    'clear the control DC
    offUC.Clear
    
    'copy the backbuffer based on the Value, Max, and Min properties
    DisplayProgress
    
    ' update the font
    Set offUC.Font = UserControl.Font
    
    'display the caption if necessary
    If Not strCaption = "" Then offUC.DrawAlignedTxt strCaption, haAlignment, vaAlignment, True
    
    'display the border if necessary
    If Not r3dStyle = None Then offUC.Draw3DEdge 0, 0, offUC.Width, offUC.Height, r3dStyle
    
    ' display the control
    offUC.BltTo UserControl.hdc
    UserControl.Refresh
    
End Sub

Private Sub DisplayProgress()
 Dim lH As Long, lW As Long
 Dim lDestX As Long, lDestY As Long
 Dim lSrcX As Long, lSrcY As Long
 
 ' display the percentage of the control to display based on the Min, Max, and Value properties _
   and display it in the direction requested
 Select Case apbDirection
 
    Case apbRightToLeft
        lH = offUC.Height
        lW = ProgressPercentage * offUC.Width
        lDestX = offUC.Width - lW
        lSrcX = lDestX
        
    Case apbBottomToTop
        lH = ProgressPercentage * offUC.Height
        lW = offUC.Width
        lDestY = offUC.Height - lH
        lSrcY = lDestY
        
    Case apbTopToBottom
        lH = ProgressPercentage * offUC.Height
        lW = offUC.Width
    
    Case apbFromCenterHorizontal
        lH = offUC.Height
        lW = ProgressPercentage * offUC.Width
        lDestX = (offUC.Width / 2) - (lW / 2)
        lSrcX = lDestX
    
    Case apbFromCenterVertical
        lH = ProgressPercentage * offUC.Height
        lW = offUC.Width
        lDestY = (offUC.Height / 2) - (lH / 2)
        lSrcY = lDestY
        
    Case Else
        lH = offUC.Height
        lW = ProgressPercentage * offUC.Width
        
 End Select
 
 If lW = 0 Then lW = 1
 If lH = 0 Then lH = 1
  
' Blt the Percentage of the control based on the Val, Min, Max properties
 offBB.BltTo offUC.hdc, lDestX, lDestY, lSrcX, lSrcY, lH, lW
        
End Sub

Public Property Get ProgressDirection() As ProgressDirection
Attribute ProgressDirection.VB_Description = "Returns or sets the direction in which the progress bar is filled."

    ProgressDirection = apbDirection

End Property

Public Property Let ProgressDirection(ByVal pDirection As ProgressDirection)
        
        If Not pDirection = apbDirection Then
        
            apbDirection = pDirection
            
            ' only update display if there is something to display
            If Not lngVal = lngMin Then UpdateBB: DrawControl
            
        End If
    
End Property

Public Function ProgressPercentage() As Double

    If lngMin = 0 Or lngMax = 0 Or lngVal = 0 Then
        
        ProgressPercentage = 0
        
    Else
       
       ' calculate the percentage of the control to display
        ProgressPercentage = (lngVal - lngMin) / (lngMax - lngMin)
        
    End If
    
End Function

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
    
    Set Font = UserControl.Font
    
    ' if there is a caption to display use the timer to update with the new font
    If Not strCaption = "" Then tmrRefresh.Interval = 1
    
End Property

Public Property Set Font(pFont As StdFont)

    Set UserControl.Font = pFont
    
    If Not strCaption = "" Then DrawControl
    
End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns or sets the graphic to be displayed in the progress bar if the Style proprety is set to apbImage."

    Set Picture = stdProgPic

End Property

Public Property Set Picture(ByVal pPicture As StdPicture)
    
    Set stdProgPic = pPicture
    
    ' if control is using the Image option then redraw the control to reflect the new image
    If apbStyle = apbImage Then UpdateBB: DrawControl
    
End Property

Private Function IsStdPicNothing(stdPic As StdPicture) As Boolean
On Error GoTo x:
 Dim lHnd As Long

' use error trapping to determine whether the StdPic object is initialized
 lHnd = stdPic.Handle
x:
 If Err.Number = 91 Then
    
    IsStdPicNothing = True

 Else
    
    IsStdPicNothing = False

 End If

End Function

Public Property Get OutOfBoundsError() As Boolean
Attribute OutOfBoundsError.VB_Description = "If set to True an error will occur (error number is equal to vbObjectError + 66) if a value is set to greater than the Max property or less than the Min property."
    
    OutOfBoundsError = boolOBErr
    
End Property

Public Property Let OutOfBoundsError(ByVal pOBErr As Boolean)
    ' when set to true the control will raise an error when you try to set the _
      control to a value outside the Min and Max properties, if set to False then _
      the control will automatically change the value property to the closest _
      of the Min, Max properties when you set it oustide their bounds
    boolOBErr = pOBErr
    
End Property
