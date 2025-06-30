VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClaseCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public NumeroCuenta As String
Public Saldo As Double
Public UVR As Double
Public FechaUltimoMovimiento As Date
Public DiasInactivo As Long
Public SaldoUVR As Double

Public Function Esreportable() As String

If DiasInactivo > 365 And SaldoUVR < 322 Then

Esreportable = "Si"

Else

Esreportable = "No"

End If

End Function




