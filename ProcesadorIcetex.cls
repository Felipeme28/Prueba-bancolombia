VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcesadorIcetex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Procesar()
    On Error GoTo ManejoError
    
 Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ICETEX")
    
    Dim i As Long
    Dim cta As ClaseCuenta
    Dim UVR As Double
    Dim fechaHoy As Date
    fechaHoy = Date
        
      For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Set cta = New ClaseCuenta
        
        With cta
            .NumeroCuenta = ws.Cells(i, 4).Value
            .Saldo = CDbl(ws.Cells(i, 6).Value)
            .UVR = CDbl(ws.Cells(i, 13).Value)
            .FechaUltimoMovimiento = ws.Cells(i, 7).Value
            .DiasInactivo = fechaHoy - .FechaUltimoMovimiento
            .SaldoUVR = .Saldo / .UVR
        End With
        
        ' Cargar en la hoja
        ws.Cells(i, 11).Value = cta.DiasInactivo
        ws.Cells(i, 14).Value = Round(cta.SaldoUVR, 0)
        ws.Cells(i, 15).Value = cta.Esreportable
    Next i

    MsgBox "CALCULO EJECUTADO.", vbInformation
    Exit Sub

ManejoError:
    MsgBox "Error: " & Err.Description
End Sub
