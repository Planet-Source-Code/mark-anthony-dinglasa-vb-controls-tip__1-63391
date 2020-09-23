VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTips 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Vb Controls Tips !"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmTips.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      Picture         =   "frmTips.frx":8331
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin MSForms.Label Label6 
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   4500
      ForeColor       =   4194304
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   $"frmTips.frx":8BFB
      PicturePosition =   393224
      Size            =   "7937;3625"
      BorderColor     =   16711680
      BorderStyle     =   1
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Some Examples of Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   600
      TabIndex        =   10
      Top             =   240
      Width           =   3345
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
      VariousPropertyBits=   746588179
      BackColor       =   16777215
      ForeColor       =   16711680
      DisplayStyle    =   4
      Size            =   "3836;873"
      Value           =   "0"
      Caption         =   "Check1"
      PicturePosition =   327683
      Picture         =   "frmTips.frx":8D6E
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
      VariousPropertyBits=   746604563
      ForeColor       =   16711680
      Size            =   "3836;661"
      Value           =   "Text1"
      BorderColor     =   16711680
      SpecialEffect   =   6
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   2205
      ForeColor       =   16711680
      VariousPropertyBits=   8388627
      Caption         =   "Command1"
      PicturePosition =   327683
      Size            =   "3889;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   2220
      ForeColor       =   16711680
      BackColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "Label1"
      PicturePosition =   393224
      Size            =   "3916;873"
      BorderColor     =   16711680
      BorderStyle     =   1
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MS Form 2.0 Library"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   2445
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1485
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'xxxxxxxxxxxxxxxxxxxx Read The Tips ! xxxxxxxxxxxxxxxxxxxxx

'About the hWnd Property !

'The hWnd property doesn't appear in the Properties window
'because its value is available only at run time. Moreover,
'it's a read-only property, and therefore you can't assign
'a value to it. The hWnd property returns the 32-bit integer
'value that Windows uses internally to identify a control.
'This value is absolutely meaningless in standard Visual
'Basic programming and only becomes useful if you invoke
'Windows API routines.
'Even if you're not going to use this property in your code,
'it's good for you to know that not all controls support it
'and it's important to understand why.

'Visual Basic controls—both intrinsic controls and external
'Microsoft ActiveX controls—can be grouped in two categories:
'standard controls and windowless (or lightweight) controls.
'To grasp the difference between the two groups, let's
'compare the PictureBox control (a standard control) and the
'Image control (a windowless control). Even though they
'appear similar at a first glance, behind the scenes they
'are completely different.

'When you place a standard control on the form, Visual
'Basic asks the operating system to create an instance of
'that control's class, and in return Windows passes back
'to Visual Basic the internal handle to that control,
'which the language then exposes to the programmer through
'the hWnd property. All subsequent operations that Visual
'Basic performs on that control—resizing, font setting,
'and so on—are actually delegated to Windows. When the
'application raises an event (such as resizing), Visual
'Basic runtime calls an internal Windows API function and
'passes it the handle so that Windows knows which control
'is to be affected.

'Lightweight controls such as Image controls, on the other
'hand, don't correspond to any Windows object and are
'entirely managed by Visual Basic itself. In a sense,
'Visual Basic just simulates the existence of that control:
'It keeps track of all the lightweight controls and redraws
'them each time the form is refreshed. For this reason,
'lightweight controls don't expose an hWnd property because
'there aren't any Windows handles associated with them.
'Windows doesn't even know a control is there.

'From a practical point of view, the distinction between
'standard and lightweight controls is that the former
'consume system resources and memory while the latter
'don't. For this reason, you should always try to replace
'standard controls with lightweight controls.
'For example, use an Image control instead of a PictureBox
'control unless you really need some of PictureBox's
'specific features. To give you an idea of what this means
'in practice, a form with 100 PictureBox controls loads 10
'times slower than a form with 100 Image controls.
