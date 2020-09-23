VERSION 5.00
Begin VB.MDIForm frmMDIForm 
   BackColor       =   &H80000004&
   Caption         =   "SizeSubClasser - MDI Forms"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucSizeSubclass ucSizeSubclass2 
      Left            =   4920
      Top             =   2640
      _ExtentX        =   423
      _ExtentY        =   423
      Enabled         =   0   'False
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
End
Attribute VB_Name = "frmMDIForm"
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
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '   Make sure the the calling form is at the top layer
    '   in the visual feild....
    frmSDIForm.ZOrder 0
End Sub
