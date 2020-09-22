VERSION 5.00
Begin VB.Form frmDraw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Drawing Demonstration"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   4425
      Picture         =   "frmDraw.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   17
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   3480
      Picture         =   "frmDraw.frx":014A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   16
      Top             =   240
      Width           =   240
   End
   Begin VB.OptionButton optIcon 
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   15
      Top             =   240
      Width           =   225
   End
   Begin VB.OptionButton optIcon 
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   14
      Top             =   240
      Width           =   225
   End
   Begin VB.OptionButton optIcon 
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   13
      Top             =   240
      Value           =   -1  'True
      Width           =   225
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Selected"
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Shadow"
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Disabled"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   375
      Index           =   3
      Left            =   4560
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   3360
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   2160
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   2475
      Picture         =   "frmDraw.frx":0294
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   960
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Normal"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select an icon to draw:"
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   300
      Width           =   1635
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   16
      X2              =   336
      Y1              =   41
      Y2              =   41
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Normal:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Disabled:"
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Shadow:"
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   960
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "XP Style:"
      Height          =   195
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   16
      X2              =   336
      Y1              =   40
      Y2              =   40
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Programmer:    Brent Culpepper (IDontKnow)
    'Project:       GDI Drawing Demo
    'Date:          February 07, 2004
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Option Explicit

' Used with TranslateColor
Private Const CLR_INVALID = -1
  
' Draw State - Image
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4
' Draw State - Flags
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000

' Flags for DrawIconEx
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Sub cmd_Click(Index As Integer)
    Dim hIcon As Long
    Dim lIconSize As Long
    Dim lr As Long
    Dim brush As Long
    Dim brushColor As Long
    Dim xpos As Long
    Dim ypos As Long
    Dim bar As RECT
    
    ' Get a pointer to the icon. If a new icon was create here,
    ' for instance with ExtractIcon, we would need to destroy it
    ' when finished to prevent a memory leak.
    hIcon = picSource(SelectedOption).Picture
    
    ' Because the source picboxes are set to AutoSize, we can
    ' just get the size from the picbox. API graphics use pixels.
    lIconSize = picSource(SelectedOption).Width
    If hIcon = 0 Then Exit Sub
    
    ' Calculate the x and y positions for the icon we are drawing:
    xpos = (pic(Index).ScaleWidth / 2) - (picSource(SelectedOption).Width / 2)
    ypos = (pic(Index).ScaleHeight / 2) - (picSource(SelectedOption).Height / 2)
    
    ' Clear any existing graphics from the destination:
    pic(Index).Cls
    
    Select Case Index
        Case 0
            ' Draw the icon in normal state:
            DrawIconEx pic(Index).hdc, xpos, ypos, hIcon, lIconSize, lIconSize, 0, 0, DI_NORMAL
        
        Case 1
            ' Draw the icon disabled
            lr = DrawState(pic(Index).hdc, 0, 0, hIcon, 0, xpos, ypos, lIconSize, lIconSize, DST_ICON Or DSS_DISABLED)
            
        Case 2
            ' Combine normal & mono states to create icon w/shadow
            ' First we create a brush with the shadow color:
            brush = CreateSolidBrush(RGB(136, 141, 157))
            
            ' Now we draw a single color icon to the source using the brush.
            ' The x and y positions are shifted right and down:
            lr = DrawState(pic(Index).hdc, brush, 0, hIcon, 0, xpos + 1, ypos + 1, lIconSize, lIconSize, DST_ICON Or DSS_MONO)
            
            ' We created a GDI object so we must destroy it to
            ' prevent memory leak!
            DeleteObject brush
            
            ' Now draw the icon in the normal state over the existing icon.
            ' The x and y positions are shifted left and up:
            lr = DrawState(pic(Index).hdc, 0, 0, hIcon, 0, xpos - 1, ypos - 1, lIconSize, lIconSize, DST_ICON Or DSS_NORMAL)
        
        Case 3
            ' Draw XP-style for a selected icon.
            ' First we define borders for the rectangle:
            Dim Left_x As Long, Top_y As Long
            Dim Right_x As Long, Bottom_y As Long
            
            Left_x = pic(Index).ScaleLeft
            Top_y = pic(Index).ScaleTop
            Right_x = pic(Index).ScaleWidth
            Bottom_y = pic(Index).ScaleHeight
            
            ' Draw a transparent backround in  a lighter version
            ' of the highlight color. First we set the rectangle:
            SetRect bar, 0, 0, Right_x, Bottom_y
            
            ' Now create a solid brush that combines the highlight color
            ' with the buttonface color (use any color you want though)
            brush = CreateSolidBrush(BlendColor(vbHighlight, vbButtonFace, 80))
            
            ' Fill in the rectangle using the brush:
            FillRect pic(Index).hdc, bar, brush
            
            ' Clean up:
            DeleteObject brush
             
            ' Set a rectangle for the border color:
            SetRect bar, 0, 0, Right_x, Bottom_y
            
            ' For safety sake convert OLE color to system color:
            brushColor = TranslateColor(vbHighlight)
            brush = CreateSolidBrush(brushColor)
            
            ' Draw the outlined border:
            FrameRect pic(Index).hdc, bar, brush
            
            ' Clean up:
            DeleteObject brush
            
            ' Now we draw the icon with the shadow using the same
            ' method as already described:
            brush = CreateSolidBrush(RGB(136, 141, 157))
            lr = DrawState(pic(Index).hdc, brush, 0, hIcon, 0, xpos + 1, ypos + 1, lIconSize, lIconSize, DST_ICON Or DSS_MONO)
            DeleteObject brush
            lr = DrawState(pic(Index).hdc, 0, 0, hIcon, 0, xpos - 1, ypos - 1, lIconSize, lIconSize, DST_ICON Or DSS_NORMAL)
    End Select
    
    ' Refresh the destination so the new drawing will be visible:
    pic(Index).Refresh
    
End Sub

Private Function SelectedOption() As Long
   SelectedOption = optIcon(0).Value * 0 Or _
                    optIcon(1).Value * -1 Or _
                    optIcon(2).Value * -2
End Function

'##################################################
' BlendColor is from http://www.vbaccelerator.com
'##################################################
Private Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                                ByVal oColorTo As OLE_COLOR, _
                                Optional ByVal alpha As Long = 128) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = TranslateColor(oColorFrom)
   lCTo = TranslateColor(oColorTo)
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB( _
      ((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), _
      ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), _
      ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255) _
      )
      
End Property

' Convert Automation color to Windows color
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub Form_Load()
    pic(3).BorderStyle = 0
End Sub
