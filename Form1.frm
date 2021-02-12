VERSION 5.00
Begin VB.Form FormUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Array3D(), Row As Integer, Col As Integer, Page As Integer

Private Sub Combo1_Click()
    'Clear the form
    Me.Cls
    'Set page selected
    Page = Combo1.Text
    'Print the array info on the form
    For Row = 1 To UBound(Array3D, 2)
        For Col = 1 To UBound(Array3D, 1)
            Me.Print Array3D(Col, Row, Page)
        Next
    Next
End Sub

Private Sub Form_Load()
    'This dimensions a three element array
    ReDim Array3D(1, 10, 3)
    'This loads data into each element of the array starting with the
    'third and nesting to the first
    For Page = 1 To UBound(Array3D, 3)
        For Row = 1 To UBound(Array3D, 2)
            For Col = 1 To UBound(Array3D, 1)
                Array3D(Col, Row, Page) = Page & "-" & Row
            Next
        Next
        'This will add each page number to the combo box
        Combo1.AddItem Page
    Next
    'This will initially display the first page
    Combo1.ListIndex = 0
End Sub
