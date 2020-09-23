VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AutoResize Example"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSample2 
      Height          =   285
      Left            =   3750
      TabIndex        =   4
      Top             =   990
      Width           =   3405
   End
   Begin VB.TextBox txtSample1 
      Height          =   285
      Left            =   3750
      TabIndex        =   3
      Top             =   180
      Width           =   3405
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5160
      TabIndex        =   2
      Top             =   4860
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   6210
      TabIndex        =   1
      Top             =   4860
      Width           =   855
   End
   Begin VB.ListBox lstSampleList 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3375
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   $"SampleResize.frx":0000
      Height          =   975
      Left            =   3750
      TabIndex        =   6
      Top             =   3390
      Width           =   3195
   End
   Begin VB.Label lblFontResize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Code By M@"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   3810
      TabIndex        =   5
      Top             =   1590
      Width           =   3195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project demonstrates a method I came up with to automatically rescale a form
'when a user resizes it. There are other methods, but they require customization for
'each form. This code can be pasted into any form and will work with no changes.
'There are no rules to where you can place the objects, the number of objects, etc.

'NOTE: This code SCALES every object. This includes text boxes, etc so that they take
'      up the same amount of real estate in the whole form. This may not be appropriate
'      for some types of controls. This is just an example, so you need to experiement with
'      your own forms.



'----------------------Add this to the declarations section--------------------

Dim ObjectList() As ScreenObject
Dim CurrentObject As Object
Dim intReferenceHeight As Integer
Dim intReferenceWidth As Integer
'-------------------------------------------------------------------------------

Private Sub Form_Load()
    ' This is the call to load the initial values of the objects on your form.
    ' It should be placed in the Form_Load event
    GetCurrentPositions
End Sub

Private Sub Form_Resize()
    'This call must be placed in the Form_Resize event. It will rescale the objects.
    Call AutoScale
    End Sub

Sub AutoScale()

Dim dblXMultiplier As Double
Dim dblYMultiplier As Double
Dim intObjectNumber As Integer
Dim intFontSize As Integer
        
'   Get ratio of initial form size to current form size
dblXMultiplier = Me.Height / intReferenceHeight
dblYMultiplier = Me.Width / intReferenceWidth

'   resize each object
For intObjectNumber = 0 To UBound(ObjectList)
    For Each CurrentObject In Me
        If CurrentObject.TabIndex = ObjectList(intObjectNumber).Index Then
             With CurrentObject
                .Left = ObjectList(intObjectNumber).Left * dblYMultiplier
                .Width = ObjectList(intObjectNumber).Width * dblYMultiplier
                .Height = ObjectList(intObjectNumber).Height * dblXMultiplier
                .Top = ObjectList(intObjectNumber).Top * dblXMultiplier
             End With
        End If
    Next CurrentObject
Next intObjectNumber

'   This is a sample of how to rescale the font of an object as well.You can see the
'   effect best when you maximize the form.
If Int(dblXMultiplier) > 0 Then
    intFontSize = Int(dblXMultiplier * 8)
    lblFontResize.FontSize = intFontSize
    txtSample1.FontSize = intFontSize
    lblInstructions.FontSize = intFontSize
End If


End Sub


Sub GetCurrentPositions()
Dim intObjectNumber As Integer
'   Load the current positions of each object into a user defined type array.
'   This information will be used to rescale them in the AutoScale function.
For Each CurrentObject In Me
    ReDim Preserve ObjectList(intObjectNumber)
    With ObjectList(intObjectNumber)
        .Name = CurrentObject
        .Index = CurrentObject.TabIndex
        .Left = CurrentObject.Left
        .Top = CurrentObject.Top
        .Width = CurrentObject.Width
        .Height = CurrentObject.Height
    End With
    intObjectNumber = intObjectNumber + 1
Next CurrentObject
    
'   This is what the object sizes will be compared to on rescaling.
    intReferenceHeight = Me.Height
    intReferenceWidth = Me.Width

End Sub

