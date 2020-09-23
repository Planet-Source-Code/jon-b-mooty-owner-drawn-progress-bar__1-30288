VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "JM Progress Bar Test"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForeColor 
      Caption         =   "ForeColor"
      Height          =   375
      Left            =   1680
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
   End
   Begin PRJPBT1.JMProgress JMProgress1 
      Height          =   2295
      Left            =   120
      Top             =   4560
      Width           =   4695
      _extentx        =   8281
      _extenty        =   4048
      caption         =   "JMProgress"
      fillcolor       =   -2147483635
      fontname        =   "Comic Sans MS"
      fontsize        =   9.75
      fontbold        =   -1  'True
      gradientend     =   8388672
      gradientstart   =   65535
   End
   Begin VB.ComboBox cmbMessage 
      Height          =   405
      ItemData        =   "frmTest.frx":0000
      Left            =   2880
      List            =   "frmTest.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox cmbBorder 
      Height          =   405
      ItemData        =   "frmTest.frx":002D
      Left            =   2880
      List            =   "frmTest.frx":0040
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1080
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   0
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmStyle 
      Caption         =   "Classic Options"
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   4695
      Begin VB.CommandButton cmdFillColor 
         Caption         =   "FillColor"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdBackColor 
         Caption         =   "BackColor"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblFillColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.ComboBox cmbDirection 
      Height          =   405
      ItemData        =   "frmTest.frx":0070
      Left            =   2880
      List            =   "frmTest.frx":0086
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox cmbStyle 
      Height          =   405
      ItemData        =   "frmTest.frx":00F4
      Left            =   2880
      List            =   "frmTest.frx":0101
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtMax 
      Height          =   405
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtMin 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame frmStyle 
      Caption         =   "Gradient Options"
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   4695
      Begin VB.CommandButton cmdGradStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdGradEnd 
         Caption         =   "End"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblGradEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   19
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblGradStart 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame frmStyle 
      Caption         =   "Image File Options"
      Height          =   1455
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   4695
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse....."
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label lblForeColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      TabIndex        =   28
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblMousePos 
      Height          =   375
      Left            =   2235
      TabIndex        =   26
      Top             =   4080
      Width           =   2655
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327680
   End
   Begin VB.Label Label6 
      Caption         =   "Message"
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Border"
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Direction"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Style"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Max"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Min"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmbBorder_Click()
    
    ' change the 3D border of the ProgressBar
    JMProgress1.BorderStyle = cmbBorder.ItemData(cmbBorder.ListIndex)
    
End Sub

Private Sub cmbDirection_Click()
    
    ' change the direction in which the progress bar is displayed when the value property is changed
    JMProgress1.ProgressDirection = cmbDirection.ItemData(cmbDirection.ListIndex)
    
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo x:
    
    cdl.ShowOpen
    
    If Not Len(Dir(cdl.filename)) = 0 Then
        
        'This code is form sample only as you can see an imagelist is not essentialy needed _
         this is just to demonstrate the use of an imagelist with the ProgressBar control _
         you could also load a file directly into the picture property of the ProgressBar by _
         using the following REM'd statement
         Rem JMProgress1.Picture = LoadPicture(cdl.FileName)
        ImageList1.ListImages.Remove 1
        ImageList1.ListImages.Add , , LoadPicture(cdl.filename)
        Set JMProgress1.Picture = ImageList1.ListImages(1).Picture
    
    End If
    
x:
End Sub

Private Sub cmdExit_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdFillColor_Click()
 On Error GoTo x:
    
    cdl.ShowColor
    
    ' update the fill color which is only used when the ProgressStyle is set to classic
    JMProgress1.FillColor = cdl.Color
    lblFillColor.BackColor = cdl.Color
    
x:
End Sub

Private Sub cmdForeColor_Click()
On Error GoTo x:
    
    cdl.ShowColor
    
    ' change the forecolor of the progress bar
    JMProgress1.ForeColor = cdl.Color
    lblBackColor.BackColor = cdl.Color
    
x:
End Sub

Private Sub cmdGradEnd_Click()
On Error GoTo x:

    cdl.ShowColor
    
    'change the end color of the gradient fade
    JMProgress1.GradientEnd = cdl.Color
    lblGradEnd.BackColor = cdl.Color
    
x:
End Sub

Private Sub cmdGradStart_Click()
On Error GoTo x:

    cdl.ShowColor
    
    ' change the starting color of the gradient fade
    JMProgress1.GradientStart = cdl.Color
    lblGradStart.BackColor = cdl.Color
    
x:
End Sub

Private Sub cmdRun_Click()
On Error Resume Next
 Dim iCt As Long
    
    For iCt = JMProgress1.Min To JMProgress1.Max
        
        ' change the value property to update the progress display
        JMProgress1.Value = iCt
        
        ' if the form is closed then exit the update loop
        If DoEvents = 0 Then Exit For
        
    Next iCt
    
End Sub

Private Sub cmbStyle_Click()
    
    ' update the style, classic, gradient, or imagefile
    JMProgress1.ProgressStyle = cmbStyle.ItemData(cmbStyle.ListIndex)
    
    frmStyle(0).Visible = False
    frmStyle(1).Visible = False
    frmStyle(2).Visible = False
    
    ' display the approriate frame for each style
    frmStyle(cmbStyle.ListIndex).Visible = True
    
End Sub

Private Sub cmdBackColor_Click()
 On Error GoTo x:
    
    cdl.ShowColor
    
    ' change the backcolor of the progress bar
    JMProgress1.BackColor = cdl.Color
    lblBackColor.BackColor = cdl.Color
    
x:
End Sub

Private Sub Form_Load()
    
    lblMousePos.Caption = "Mouse Outside ProgressBar!"
    cdl.Filter = "Image Files (*.bmp;*.jpg;*.jpeg;*.gif;.ico)|*.BMP;*.JPG;*.JPEG;*.GIF;*.ICO"
    cdl.CancelError = True
    txtMax.Text = JMProgress1.Max
    txtMin.Text = JMProgress1.Min
    cmbStyle.ListIndex = JMProgress1.ProgressStyle
    cmbDirection.ListIndex = JMProgress1.ProgressDirection
    lblForeColor.BackColor = JMProgress1.ForeColor
    lblBackColor.BackColor = JMProgress1.BackColor
    lblFillColor.BackColor = JMProgress1.FillColor
    lblGradStart.BackColor = JMProgress1.GradientStart
    lblGradEnd.BackColor = JMProgress1.GradientEnd
    cmbBorder.ListIndex = JMProgress1.BorderStyle
    cmbMessage.ListIndex = 0
    ImageList1.ListImages.Add , , LoadPicture(IIf(Right(App.Path, 1) = "\", App.Path & "tstimg.jpg", App.Path & "\" & "tstimg.jpg"))
    Set JMProgress1.Picture = ImageList1.ListImages(1).Picture
    
End Sub

Private Sub JMProgress1_MouseEnter()
 lblMousePos.Caption = "Mouse Inside ProgressBar!"
End Sub

Private Sub JMProgress1_MouseExit()
lblMousePos.Caption = "Mouse Outside ProgressBar!"
End Sub

Private Sub JMProgress1_ProgressComplete()
' this event happens when the value property and the max property are the same

    If cmbMessage.ListIndex = 2 Then JMProgress1.Caption = "Finished!"

End Sub

Private Sub JMProgress1_ValChange(ByVal lngValue As Long, ByVal dblPercentage As Double)
' this event occurs when the value property of the progress bar is changed to any value _
  different then the previous change
  
    Select Case cmbMessage.ListIndex
    
        Case 0
            'display the percentage of the progress bar filled
            JMProgress1.Caption = CStr(CInt(dblPercentage * 100)) & " %"
        
        Case 1
            'display the value property of the progressbar
            JMProgress1.Caption = CStr(lngValue)
            
        Case 2
            'display a message based on the percentage of the progress bar filled
            Select Case dblPercentage
                
                Case Is >= 0.75
                    JMProgress1.Caption = "Almost There!"
                
                Case Is >= 0.5
                    JMProgress1.Caption = "Getting Closer!"
                
                Case Is >= 0.25
                    JMProgress1.Caption = "Warming Up!"
                
                Case Is >= 0
                    JMProgress1.Caption = "Getting Started!"
                    
             End Select
            
        End Select
End Sub

Private Sub txtMax_LostFocus()

    'set the max property of the progress bar based on user entry
    If IsNumeric(txtMax.Text) = False Then
        
        txtMax.Text = JMProgress1.Max
    
    Else
        
        JMProgress1.Max = CLng(txtMax.Text)
        txtMax.Text = JMProgress1.Max
        
    End If
    
End Sub

Private Sub txtMin_LostFocus()
    ' set the min property of the progress bar based on user entry
    
    If IsNumeric(txtMin.Text) = False Then
        
        txtMin.Text = JMProgress1.Min
    
    Else
        
        JMProgress1.Min = CLng(txtMin.Text)
        txtMin.Text = JMProgress1.Min
        
    End If
    
End Sub
