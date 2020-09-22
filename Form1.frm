VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame OwnerFrame 
      Caption         =   "FrameCaption - 0"
      Height          =   2955
      Index           =   0
      Left            =   315
      TabIndex        =   1
      Top             =   630
      Width           =   4755
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "HELLO WORLD!!! #0"
         Height          =   195
         Left            =   1500
         TabIndex        =   10
         Top             =   1110
         Width           =   1560
      End
   End
   Begin VB.Frame OwnerFrame 
      Caption         =   "FrameCaption - 2"
      Height          =   2955
      Index           =   2
      Left            =   315
      TabIndex        =   4
      Top             =   630
      Width           =   4755
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "HELLO WORLD!!! #2"
         Height          =   195
         Left            =   1410
         TabIndex        =   8
         Top             =   465
         Width           =   1560
      End
   End
   Begin VB.Frame OwnerFrame 
      Caption         =   "FrameCaption - 1"
      Height          =   2955
      Index           =   1
      Left            =   315
      TabIndex        =   3
      Top             =   630
      Width           =   4755
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "HELLO WORLD!!! #1"
         Height          =   195
         Left            =   1470
         TabIndex        =   9
         Top             =   1080
         Width           =   1560
      End
   End
   Begin VB.PictureBox Cover1 
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   285
      ScaleHeight     =   165
      ScaleWidth      =   1065
      TabIndex        =   7
      Top             =   435
      Width           =   1065
   End
   Begin VB.CommandButton CmdTab 
      Caption         =   "Tab1"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   285
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdPlat 
      Enabled         =   0   'False
      Height          =   3285
      Left            =   195
      TabIndex        =   0
      Top             =   435
      Width           =   5010
   End
   Begin VB.CommandButton CmdTab 
      Caption         =   "Tab2"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   1410
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdTab 
      Caption         =   "Tab3"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   2535
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########################################################
'#  ATTENTION:                                          #
'#                                                      #
'#  yar-interactive software is happy to                #
'#  share our knowledge of Visual Basic with you.       #
'#  please visit our website at                         #
'#  http://www.yarinteractive.com                       #
'#                                                      #
'#  There, you can find more great VB resources and     #
'#  great software.                                     #
'#                      - Thanks.                       #
'########################################################

'HOW TO USE THIS...
'First of all, you should run this code to get a sence of
'what it does, it is meant to be an alternative to the
'tabed dialog control, for developers that prefer to keep
'there applications as self inclosed as possible
'
'PUTTING CONTROLS ON THE TAB-DIALOG...
'Obviously, you can place controls in the frame that
'you can see, but how do you place controls in
'another tab?...
'To place a control on the tab dialog in this example
'select OwnerFrame(2) from the PROPERTIES WINDOW.
'Move into design mode and you should see 8 highlighted boxes
'along the edges of a frame, Right-Click on one of these
'boxes, and select 'Bring to Front'. The frame that will appear
'when the button Captioned 'Tab3'. NOTE that you can change
'the caption of this frame, ect. by selecting it from the
'properties window without having to bring it to
'the front.

Private Sub CmdTab_Click(Index As Integer)

'Brings the platform to the front to cover edges of
'unselected buttons. (CmdPlat is a button behind all the
'frames, and tabs to make a 'raised platform effect'
CmdPlat.ZOrder 0

'Bring the frame associated with the pressed button
'to the front.
OwnerFrame(Index).ZOrder 0

'Bring the Cover over the bottom edge of the
'selected button, the make it look like
'it is 'attached' to the rest of the
' "raised platform"
Cover1.Left = CmdTab(Index).Left

'The cover that hides the edge of the selected
'button got hidden by the big "Raised platform"
'we need it to be on top, so bring it to the fron.
Cover1.ZOrder 0

'GO TO http://www.yarinteractive.com for more great VB codes
'and resources
End Sub
