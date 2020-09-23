VERSION 5.00
Begin VB.UserControl TranspContainer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   ControlContainer=   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   7875
End
Attribute VB_Name = "TranspContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_Transparent = True
Const m_def_TransparentColor = 0
'Property Variables:
Dim m_Transparent As Boolean
Dim m_TransparentColor As OLE_COLOR
'Event Declarations:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()



Private Sub UserControl_Initialize()
    DrawControl
End Sub

Private Sub UserControl_Resize()
    
    RaiseEvent Resize
    DrawControl
'Must On the UserControl_Resize the DrawControl becouse when reise the event resize from the form can be
'change the position of each control
'in Case THE movmend from the control's is whithout Resize For this case must be called form the form the DrawControl
'XXControlName.DrawControl
End Sub

Public Sub DrawControl()
    UserControl.BackStyle = 1   'Set control's backstyle  to NON Transparent''
    UserControl.Cls
If m_Transparent = True Then

    Dim FCont As Form
    Set FCont = UserControl.Extender.Parent 'Set the FCont to curient usercontrol extenter to his parent(Form)
    For Each Control In FCont 'Search The Form For Control's
        If Control.Container.Name = UserControl.Extender.Name Then 'If The Container of the control is same as the extender then
            UserControl.Line (Control.Left, Control.Top)-(Control.Left + Control.Width - 15, Control.Top + Control.Height - 15), , BF 'make Back ONE LINE SAME AS THE SIZE Of the control
        End If
    Next
    
    UserControl.MaskColor = TransparentColor    'Set the control's mask color to the background color of the drawing area
    UserControl.MaskPicture = UserControl.Image     'Set the mask image from what we created on the drawing area
    UserControl.BackStyle = 0   'Set control's backstyle to Transparent
Else
    UserControl.BackStyle = 1   'Set control's backstyle to NON Transparent'
    UserControl.MaskPicture = Nothing     'Set the mask image TO NOTHING to free memory
End If

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawControl
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
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

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
    DrawControl
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal New_Transparent As Boolean)
    m_Transparent = New_Transparent
    PropertyChanged "Transparent"
    DrawControl
End Property

Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = m_TransparentColor
End Property

Public Property Let TransparentColor(ByVal New_TransparentColor As OLE_COLOR)
    m_TransparentColor = New_TransparentColor
    PropertyChanged "TransparentColor"
    DrawControl
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Transparent = m_def_Transparent
    m_TransparentColor = m_def_TransparentColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Transparent = PropBag.ReadProperty("Transparent", m_def_Transparent)
    m_TransparentColor = PropBag.ReadProperty("TransparentColor", m_def_TransparentColor)
End Sub

Private Sub UserControl_Show()
    DrawControl
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Transparent", m_Transparent, m_def_Transparent)
    Call PropBag.WriteProperty("TransparentColor", m_TransparentColor, m_def_TransparentColor)
End Sub

