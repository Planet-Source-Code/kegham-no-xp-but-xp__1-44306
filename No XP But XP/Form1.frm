VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E39F68&
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   11  'Not Xor Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You think it deserve a vote feel free :)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your program title here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E39F68&
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   0
      X2              =   6600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   6210
      MouseIcon       =   "Form1.frx":00C5
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":0217
      Top             =   90
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   6210
      MouseIcon       =   "Form1.frx":0799
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":08EB
      Top             =   90
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "Form1.frx":0E6D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00D8E8EF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dont have windows XP here is the solution"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6495
   End
   Begin VB.Shape shpxp 
      BackColor       =   &H00CBE0E9&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E39F68&
      BorderWidth     =   6
      FillColor       =   &H00CBE0E9&
      FillStyle       =   0  'Solid
      Height          =   3900
      Index           =   1
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private ResultRegion As Long

Private Const RectXRound As Integer = 28
Private Const RectYRound As Integer = 28
Dim OldX As Integer, OldY As Integer, MoveIt As Boolean

Private Sub Form_Load()
 Dim nRet As Long

    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)

End Sub

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim Corraction As Integer
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim i As Integer
    
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)
    
    For i = shpxp.LBound To shpxp.UBound
        Select Case shpxp(i).Shape
            Case 0: 'rectangle & square
                ObjectRegion = CreateRectRgn( _
                        shpxp(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpxp(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpxp(i).Left + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpxp(i).Top + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY)
            Case 1: 'circle
                If shpxp(i).Width > shpxp(i).Height Then
                    Corraction = (shpxp(i).Width - shpxp(i).Height) / 2
                        
                    ObjectRegion = CreateRectRgn( _
                            (shpxp(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpxp(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpxp(i).Left - Corraction + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                Else
                    Corraction = (shpxp(i).Height - shpxp(i).Width) / 2
                        
                    ObjectRegion = CreateRectRgn( _
                            (shpxp(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpxp(i).Left + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top - Corraction + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                End If
            Case 4:  'round rectangle
            
                ObjectRegion = CreateRoundRectRgn( _
                        shpxp(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpxp(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpxp(i).Left + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpxp(i).Top + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                        RectXRound, RectYRound)
            Case 5: 'round square
                If shpxp(i).Width > shpxp(i).Height Then
                    Corraction = (shpxp(i).Width - shpxp(i).Height) / 2
                        
                    ObjectRegion = CreateRoundRectRgn( _
                            (shpxp(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpxp(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpxp(i).Left - Corraction + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                            RectXRound, RectYRound)
                Else
                    Corraction = (shpxp(i).Height - shpxp(i).Width) / 2
                        
                    ObjectRegion = CreateRoundRectRgn( _
                            (shpxp(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpxp(i).Left + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top - Corraction + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                            RectXRound, RectYRound)
                End If
            Case 3: 'circle
                If shpxp(i).Width > shpxp(i).Height Then
                    Corraction = (shpxp(i).Width - shpxp(i).Height) / 2
                        
                    ObjectRegion = CreateEllipticRgn( _
                            (shpxp(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpxp(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpxp(i).Left - Corraction + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                Else
                    Corraction = (shpxp(i).Height - shpxp(i).Width) / 2
                        
                    ObjectRegion = CreateEllipticRgn( _
                            (shpxp(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpxp(i).Left + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpxp(i).Top - Corraction + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                End If
            Case Else:  'oval
                shpxp(i).Shape = 2
                ObjectRegion = CreateEllipticRgn( _
                        shpxp(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpxp(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpxp(i).Left + shpxp(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpxp(i).Top + shpxp(i).Height) / Screen.TwipsPerPixelY + OffsetY)
        End Select
        nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
        nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
        DeleteObject ObjectRegion
    Next i
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
        OldX = X
        OldY = Y
        MoveIt = True
    End If
   
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If MoveIt = True Then
    Form1.Top = Form1.Top + Y - OldY
    Form1.Left = Form1.Left + X - OldX
End If
Image2.Visible = False
Image4.Visible = True
 
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveIt = False
End Sub

Private Sub Image2_Click()
End

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True

Image2.Visible = False
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image2.Visible = True

End Sub
