VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "APP 1.0"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Call main(textBoxFecha, opcionPuc, opcionCuif, TextBox_RutaOrigen, TextBox_RutaDestino)
End Sub

Private Sub textBoxFecha_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim fecha As Date
    
    fecha = CalendarForm.GetDate
    textBoxFecha.value = fecha
End Sub


