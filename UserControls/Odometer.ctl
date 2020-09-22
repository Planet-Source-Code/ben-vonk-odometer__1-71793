VERSION 5.00
Begin VB.UserControl Odometer 
   ClientHeight    =   864
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   384
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "Odometer.ctx":0000
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.PictureBox picDigits 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   252
   End
End
Attribute VB_Name = "Odometer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Odometer Control
'
'Author Ben Vonk
'20-02-2009 First version

Option Explicit

' Public Events
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ReachLimit(Limit As Double)

' Public Enumeration
Public Enum Scrollings
   Automatic
   Up
   Down
End Enum

Public Enum Speeds
   None
   Fast
   Normal
   Slow
End Enum

' Private Types
Private Type TriVertex
   X                           As Long
   Y                           As Long
   Red                         As Integer
   Green                       As Integer
   Blue                        As Integer
   Alpha                       As Integer
End Type

Private Type GradientRect
   UpperLeft                   As Long
   LowerRight                  As Long
End Type

' Private Variables
Private m_SpinOver             As Boolean
Private NotAutomatic           As Boolean
Private m_Digits               As Integer
Private m_Value                As Double
Private DigitWidth             As Integer
Private Waiting                As Integer
Private m_BackColorFirstDigit  As Long
Private m_BackColorOtherDigits As Long
Private m_BorderGradientColor  As Long
Private m_ForeColorFirstDigit  As Long
Private m_ForeColorOtherDigits As Long
Private m_Scroll               As Scrollings
Private Scrolling              As Scrollings
Private m_Speed                As Speeds
Private Number                 As String

' Private API's
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "Kernel32" () As Long
Private Declare Function GradientFill Lib "MSImg32" (ByVal hDC As Long, ByRef PTRIVERTEX As TriVertex, ByVal ulong As Long, pvoid As Any, ByVal ulong As Long, ByVal ulong As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Property Get BackColorFirstDigit() As OLE_COLOR
Attribute BackColorFirstDigit.VB_Description = "Returns/sets the background color used for the first digit."

   BackColorFirstDigit = m_BackColorFirstDigit

End Property

Public Property Let BackColorFirstDigit(ByVal NewBackColorFirstDigit As OLE_COLOR)

   m_BackColorFirstDigit = NewBackColorFirstDigit
   PropertyChanged "BackColorFirstDigit"
   
   Call CreateDisplay

End Property

Public Property Get BackColorOtherDigits() As OLE_COLOR
Attribute BackColorOtherDigits.VB_Description = "Returns/sets the background color used for the other digits."

   BackColorOtherDigits = m_BackColorOtherDigits

End Property

Public Property Let BackColorOtherDigits(ByVal NewBackColorOtherDigits As OLE_COLOR)

   m_BackColorOtherDigits = NewBackColorOtherDigits
   PropertyChanged "BackColorOtherDigits"
   
   Call CreateDisplay

End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object border."

   BorderColor = BackColor

End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)

   BackColor = NewBorderColor
   PropertyChanged "BorderColor"
   
   Call CreateDisplay

End Property

Public Property Get BorderGradientColor() As OLE_COLOR
Attribute BorderGradientColor.VB_Description = "Returns/sets the gradient color of an object border."

   BorderGradientColor = m_BorderGradientColor

End Property

Public Property Let BorderGradientColor(ByVal NewBorderGradientColor As OLE_COLOR)

   m_BorderGradientColor = NewBorderGradientColor
   PropertyChanged "BorderGradientColor"
   
   Call CreateDisplay

End Property

Public Property Get Digits() As Integer
Attribute Digits.VB_Description = "Returns/sets the number of digits."

   Digits = m_Digits

End Property

Public Property Let Digits(ByVal NewDigits As Integer)

   If NewDigits < 2 Then NewDigits = 2
   If NewDigits > 15 Then NewDigits = 15
   
   m_Digits = NewDigits
   PropertyChanged "Digits"
   
   Call CreateDisplay

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

    Set Font = UserControl.Font

End Property

Public Property Let Font(ByRef NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByRef NewFont As StdFont)

   Set UserControl.Font = NewFont
   PropertyChanged "Font"
   
   Call CreateDisplay

End Property

Public Property Get ForeColorFirstDigit() As OLE_COLOR
Attribute ForeColorFirstDigit.VB_Description = "Returns/sets the text color of the first digit."

   ForeColorFirstDigit = m_ForeColorFirstDigit

End Property

Public Property Let ForeColorFirstDigit(ByVal NewForeColorFirstDigit As OLE_COLOR)

   m_ForeColorFirstDigit = NewForeColorFirstDigit
   PropertyChanged "ForeColorFirstDigit"
   
   Call CreateDisplay

End Property

Public Property Get ForeColorOtherDigits() As OLE_COLOR
Attribute ForeColorOtherDigits.VB_Description = "Returns/sets the text color of the other digits."

   ForeColorOtherDigits = m_ForeColorOtherDigits

End Property

Public Property Let ForeColorOtherDigits(ByVal NewForeColorOtherDigits As OLE_COLOR)

   m_ForeColorOtherDigits = NewForeColorOtherDigits
   PropertyChanged "ForeColorOtherDigits"
   
   Call CreateDisplay

End Property

Public Property Get Scroll() As Scrollings
Attribute Scroll.VB_Description = "Returns/sets the direction to scroll the digits. If Automatic is selected the direction will be selected by the added value."

   Scroll = m_Scroll

End Property

Public Property Let Scroll(ByVal NewScroll As Scrollings)

   m_Scroll = NewScroll
   PropertyChanged "Scroll"

End Property

Public Property Get Speed() As Speeds
Attribute Speed.VB_Description = "Returns/sets the speed to scroll the digits."

   Speed = m_Speed

End Property

Public Property Let Speed(ByVal NewSpeed As Speeds)

   m_Speed = NewSpeed
   PropertyChanged "Speed"
   Waiting = SetWait

End Property

Public Property Get SpinOver() As Boolean
Attribute SpinOver.VB_Description = "Determines whether the min/max values are spinover or not."

   SpinOver = m_SpinOver

End Property

Public Property Let SpinOver(ByVal NewSpinOver As Boolean)

   m_SpinOver = NewSpinOver
   PropertyChanged "SpinOver"

End Property

Public Property Get Value() As Double
Attribute Value.VB_Description = "Returns/sets the value of an object."
Attribute Value.VB_UserMemId = 0

   Value = m_Value

End Property

Public Property Let Value(ByVal NewValue As Double)

Static blnBusy As Boolean

Dim blnLimit   As Boolean
Dim dblValue   As Double

   If blnBusy Then Exit Property
   
   blnBusy = True
   dblValue = Val(String(m_Digits, 57))
   NotAutomatic = False
   
   If NewValue < 0 Then
      If SpinOver Then
         NewValue = dblValue
         NotAutomatic = True
         Scrolling = Down
         
      Else
         NewValue = 0
         blnLimit = True
      End If
      
   ElseIf NewValue > dblValue Then
      If SpinOver Then
         NewValue = 0
         NotAutomatic = True
         Scrolling = Up
         
      Else
         NewValue = dblValue
         blnLimit = True
      End If
   End If
   
   m_Value = NewValue
   PropertyChanged "Value"
   
   Call DrawDisplay
   
   If blnLimit Then RaiseEvent ReachLimit(m_Value)
   
   blnBusy = False

End Property

Private Function SetWait() As Integer

   SetWait = 5 * (m_Speed + (1 And (m_Speed = 2)) + (3 And (m_Speed = 3)))

End Function

Private Sub CreateDisplay()

   With picBuffer
      .Picture = Nothing
      Set .Font = Font
      DigitWidth = .TextWidth("0") + 2
      .Width = DigitWidth * m_Digits + 2
      .Height = .TextHeight("0") * 2 - 4
      .BackColor = m_BackColorOtherDigits
      picBuffer.Line (.Width - DigitWidth - 1, 0)-(.Width, .Height), m_BackColorFirstDigit, BF
      .Picture = .Image
      picDigits.Width = .Width
      picDigits.Height = .Height / 2
   End With
   
   Call SetSize
   
   With picDigits
      .Top = 2
      .BackColor = BackColor
      FillGradient .hDC, 0, 0, .Width, .Height / 1.025, BackColor, m_BorderGradientColor
      .Picture = .Image
   End With
   
   Call DrawDisplay(True)

End Sub

Private Sub DrawDisplay(Optional ByVal ForceDraw As Boolean)

Static dblPrevValue  As Double
Static strPrevNumber As String

Dim intBufferY       As Integer
Dim intCount         As Integer
Dim intHeight        As Integer
Dim intLoop          As Integer
Dim intSourceY       As Integer
Dim intX             As Integer
Dim intY             As Integer
Dim intMiddle        As Integer
Dim lngWait          As Long

   If ForceDraw Then strPrevNumber = ""
   
   If (m_Scroll = Automatic) And Not NotAutomatic Then
      If dblPrevValue > m_Value Then
         Scrolling = Down
         
      Else
         Scrolling = Up
      End If
   End If
   
   dblPrevValue = m_Value
   Number = Right(String(m_Digits, 48) & m_Value, m_Digits)
   
   With picBuffer
      .Cls
      intMiddle = .Height / 2
      
      If Scrolling = Down Then
         intY = -2
         intBufferY = intMiddle
         
      ' Up
      Else
         intY = intMiddle
      End If
      
      BitBlt .hDC, 0, intBufferY, .Width, intMiddle, picDigits.hDC, intX, 0, vbSrcCopy
      
      For intCount = Len(Number) To 1 Step -1
         If Mid(Number, intCount, 1) <> Mid(strPrevNumber, intCount, 1) Then
            If intCount = Len(Number) Then
               .ForeColor = m_ForeColorFirstDigit
               
            Else
               .ForeColor = m_ForeColorOtherDigits
            End If
            
            intX = DigitWidth * (intCount - 1) + (DigitWidth - (DigitWidth - 5)) \ 2
            .CurrentX = intX
            .CurrentY = intY
            picBuffer.Print Mid(Number, intCount, 1)
         End If
      Next 'intCount
      
      If Scrolling = Down Then
         intY = .Height
         intHeight = -.Height
         intSourceY = intY - 1
         
      ' Up
      Else
         intY = 0
         intHeight = .Height
         intSourceY = 1
      End If
      
      For intLoop = .Height To intMiddle Step -1
         For intCount = Len(Number) To 1 Step -1
            If Mid(Number, intCount, 1) <> Mid(strPrevNumber, intCount, 1) Then
               intX = DigitWidth * (intCount - 1) + 2
               BitBlt .hDC, intX, intY, DigitWidth - 2, intHeight, .hDC, intX, intSourceY, vbSrcCopy
            End If
         Next 'intCount
         
         BitBlt picDigits.hDC, 0, 0, .Width, intMiddle, .hDC, 0, intBufferY, vbSrcCopy
         lngWait = GetTickCount + (Waiting And (Len(strPrevNumber) <> 0))
         picDigits.Refresh
         
         Do
            DoEvents
         Loop While GetTickCount < lngWait
      Next 'intLoop
   End With
   
   strPrevNumber = Number
   RaiseEvent Change

End Sub

Private Sub FillColors(ByRef Vertex() As Byte, ByRef Colors() As Byte)

   Vertex(1) = Colors(0)
   Vertex(3) = Colors(1)
   Vertex(5) = Colors(2)

End Sub

Private Sub FillGradient(ByVal hDC As Long, ByVal X As Double, ByVal Y As Double, ByVal Width As Double, ByVal Height As Double, ByVal ColorStart As Long, ByVal ColorEnd As Long)

Dim bytColors(3) As Byte
Dim bytVertex(7) As Byte
Dim gctObject    As GradientRect
Dim intCount     As Integer
Dim lngSwap      As Long
Dim tvxObject(1) As TriVertex

   Height = Height / 2
   
   For intCount = 0 To 1
      tvxObject(0).X = X
      tvxObject(0).Y = Y
      tvxObject(1).X = X + Width
      tvxObject(1).Y = Y + Height
      
      Call CopyMemory(bytColors(0), ColorStart, &H4)
      Call FillColors(bytVertex, bytColors)
      Call CopyMemory(tvxObject(0).Red, bytVertex(0), &H8)
      Call CopyMemory(bytColors(0), ColorEnd, &H4)
      Call FillColors(bytVertex, bytColors)
      Call CopyMemory(tvxObject(1).Red, bytVertex(0), &H8)
      
      gctObject.UpperLeft = 0
      gctObject.LowerRight = 1
      GradientFill hDC, tvxObject(0), 2, gctObject, 1, 1
      lngSwap = ColorStart
      ColorStart = ColorEnd
      ColorEnd = lngSwap
      Y = Y + Height
   Next 'intCount

End Sub

Private Sub SetSize()

Static blnBusy As Boolean

   If blnBusy Then Exit Sub
   
   blnBusy = True
   Width = picBuffer.Width * Screen.TwipsPerPixelX
   Height = (picBuffer.Height / 2 + 4) * Screen.TwipsPerPixelY
   blnBusy = False
   Refresh

End Sub

Private Sub picDigits_Click()

   RaiseEvent Click

End Sub

Private Sub picDigits_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub picDigits_KeyDown(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub picDigits_KeyPress(KeyAscii As Integer)

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub picDigits_KeyUp(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub picDigits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub picDigits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub picDigits_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Click()

   RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub UserControl_InitProperties()

   BackColor = &H404040
   m_BorderGradientColor = &H808080
   m_BackColorFirstDigit = vbWhite
   m_ForeColorFirstDigit = vbRed
   m_ForeColorOtherDigits = vbWhite
   m_Digits = 8
   m_Speed = Normal
   m_SpinOver = True
   Scrolling = Up
   Set Font = Parent.Font
   Waiting = SetWait
   
   Call CreateDisplay

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      m_BackColorFirstDigit = .ReadProperty("BackColorFirstDigit", vbWhite)
      m_BackColorOtherDigits = .ReadProperty("BackColorOtherDigits", vbBlack)
      BackColor = .ReadProperty("BorderColor", &H404040)
      m_BorderGradientColor = .ReadProperty("BorderGradientColor", &H808080)
      m_Digits = .ReadProperty("Digits", 8)
      m_ForeColorFirstDigit = .ReadProperty("ForeColorFirstDigit", vbRed)
      m_ForeColorOtherDigits = .ReadProperty("ForeColorOtherDigits", vbWhite)
      Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
      m_Scroll = .ReadProperty("Scroll", Automatic)
      m_Speed = .ReadProperty("Speed", Normal)
      m_SpinOver = .ReadProperty("SpinOver", True)
      m_Value = .ReadProperty("Value", 0)
      Waiting = SetWait
   End With
   
   Call CreateDisplay

End Sub

Private Sub UserControl_Resize()

   Call SetSize

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BackColorFirstDigit", m_BackColorFirstDigit, vbWhite
      .WriteProperty "BackColorOtherDigits", m_BackColorOtherDigits, vbBlack
      .WriteProperty "BorderColor", BackColor, &H404040
      .WriteProperty "BorderGradientColor", m_BorderGradientColor, &H808080
      .WriteProperty "Digits", m_Digits, 8
      .WriteProperty "ForeColorFirstDigit", m_ForeColorFirstDigit, vbRed
      .WriteProperty "ForeColorOtherDigits", m_ForeColorOtherDigits, vbWhite
      .WriteProperty "Font", UserControl.Font
      .WriteProperty "Scroll", m_Scroll, Automatic
      .WriteProperty "Speed", m_Speed, Normal
      .WriteProperty "SpinOver", m_SpinOver, True
      .WriteProperty "Value", m_Value, 0
   End With

End Sub
