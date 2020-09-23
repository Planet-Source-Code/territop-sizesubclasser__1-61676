VERSION 5.00
Begin VB.Form frmSDIForm 
   Caption         =   "SizeSubClasser - SDI Forms"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucSizeSubclass ucSizeSubclass2 
      Left            =   6480
      Top             =   4200
      _ExtentX        =   423
      _ExtentY        =   423
      Enabled         =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load MDI Form"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin Project1.ucSizeSubclass ucSizeSubclass1 
      Left            =   120
      Top             =   120
      _ExtentX        =   423
      _ExtentY        =   423
      MinHeight       =   3030
      MaxHeight       =   5085
      MinWidth        =   5400
      MaxWidth        =   7080
      MaximizedWidth  =   6000
      MaximizedXOffset=   1300
      MaximizedYOffset=   800
   End
   Begin VB.Label Label1 
      Caption         =   "Multiple Instances Automatically Disabled"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblResizing 
      Caption         =   "Resizing Parent Form..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblMaxHeight 
      Caption         =   "Maximum Height"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblMinHeight 
      Caption         =   "Minimum Height"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Line Line16 
      X1              =   6960
      X2              =   6720
      Y1              =   1200
      Y2              =   1080
   End
   Begin VB.Line Line15 
      X1              =   6960
      X2              =   6720
      Y1              =   1200
      Y2              =   1320
   End
   Begin VB.Line Line14 
      BorderStyle     =   2  'Dash
      X1              =   5280
      X2              =   6960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line13 
      X1              =   1560
      X2              =   1680
      Y1              =   4560
      Y2              =   4320
   End
   Begin VB.Line Line12 
      X1              =   1560
      X2              =   1440
      Y1              =   4560
      Y2              =   4320
   End
   Begin VB.Line Line11 
      BorderStyle     =   2  'Dash
      X1              =   1560
      X2              =   1560
      Y1              =   2520
      Y2              =   4560
   End
   Begin VB.Line Line10 
      X1              =   5280
      X2              =   5040
      Y1              =   1200
      Y2              =   1320
   End
   Begin VB.Line Line9 
      X1              =   5280
      X2              =   5040
      Y1              =   1200
      Y2              =   1080
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   240
      Y1              =   1200
      Y2              =   1320
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   240
      Y1              =   1200
      Y2              =   1080
   End
   Begin VB.Line Line6 
      X1              =   1560
      X2              =   1440
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line Line5 
      X1              =   1560
      X2              =   1680
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line Line4 
      X1              =   1560
      X2              =   1440
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   1680
      X2              =   1560
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5280
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   1560
      Y1              =   0
      Y2              =   2520
   End
   Begin VB.Label lblMinWidth 
      Caption         =   "Minimum Width"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblMaxWidth 
      Caption         =   "Maximum Width"
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmSDIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+  File Description:
'       ucSizeSubclass - Size Subclasser to provide Flicker-Free Size Restrictions
'
'   Product Name:
'       ucSizeSubclass.ctl
'
'   Compatability:
'       Windows: 98, ME, NT4, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'       Adapted from the following online article(s):
'       Based in large part from Paul Caton's Self-Subclassing Example (see URL below)...
'       http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       http://www.vb-helper.com/howtoint.htm (Article: "Find the system's color depth (bits per pixel)")
'
'   Legal Copyright & Trademarks (Current Implementation):
'       Copyright © 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this software.
'
'-  Modification(s) History:
'       23Jun05 - Initial build of the SizeSubclass Control
'       23Jun05 - Fixed bug with Screen size routine which reported the
'                 incorrect values of pixels
'       05Jul05 - Added additional public events and properties
'       07Jul05 - Added Error handling for multiple control instances loaded on
'                 the same form at on time.
'       12Jul05 - Added additional error checking for previous existance based on
'                 suggestions by LaVolpe and Fred.cpp. Current version now has a
'                 Public Enabled property which checks for other instances when
'                 it is set to true.
'               - Added MDI Form support and parent form subclassing for QueryClose
'                 events to make sure the Subclasser is shutdown correctly.
'               - Added User feedback to the user by "X"ing out the controls
'                 GUI when disabled...
'
'   Force Declarations
Option Explicit

Private Sub Command1_Click()
    Load frmMDIForm
    frmMDIForm.Show
End Sub

Private Sub Form_Load()
    '   Initalize the size
    With Me
        .Width = 5400
        .Height = 3030
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '   Unload everyone....
    Unload frmMDIForm
    Unload Me
End Sub

Private Sub ucSizeSubclass1_FinishedSizeMove()
    With Me
        .lblResizing.Visible = False
    End With
End Sub

Private Sub ucSizeSubclass1_ParentSizing()
    With Me
        .lblResizing.Visible = True
        .lblMaxWidth.Visible = False
        .lblMinWidth.Visible = False
        .lblMaxHeight.Visible = False
        .lblMinHeight.Visible = False
        If .WindowState <> vbMaximized Then
            '   Handing Window Width Changes...
            If (.Width = .ucSizeSubclass1.MaxWidth) Then
                .lblResizing.Visible = False
                .lblMaxWidth.Visible = True
                .lblMinWidth.Visible = False
            ElseIf (.Width = .ucSizeSubclass1.MinWidth) Then
                .lblResizing.Visible = False
                .lblMaxWidth.Visible = False
                .lblMinWidth.Visible = True
            Else
                .lblResizing.Visible = True
                Exit Sub
            End If
            '   Handle Window Height Changes...
            If (.Height = .ucSizeSubclass1.MaxHeight) Then
                .lblResizing.Visible = False
                .lblMaxHeight.Visible = True
                .lblMinHeight.Visible = False
            ElseIf (.Height = .ucSizeSubclass1.MinHeight) Then
                .lblResizing.Visible = False
                .lblMaxHeight.Visible = False
                .lblMinHeight.Visible = True
            Else
                .lblResizing.Visible = True
                Exit Sub
            End If
        Else
            '   Handle Window Height Changes...
            If (.Height = .ucSizeSubclass1.MaxHeight) Then
                .lblResizing.Visible = False
                .lblMaxHeight.Visible = True
                .lblMinHeight.Visible = False
            ElseIf (.Height = .ucSizeSubclass1.MinHeight) Then
                .lblResizing.Visible = False
                .lblMaxHeight.Visible = False
                .lblMinHeight.Visible = True
            Else
                .lblResizing.Visible = True
            End If
        End If
    End With
End Sub
