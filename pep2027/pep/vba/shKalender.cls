Option Explicit

' Doppelklick auf einen Tag: Eingabe fuer dieses Feld
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If modPEP.ImRaster(Target) Then
        Cancel = True
        modPEP.EintragDialog Target
    End If
End Sub

' Rechtsklick in markierte Tage: Eingabe fuer alle markierten Felder
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    If modPEP.ImRaster(Target) Then
        Cancel = True
        modPEP.EintragDialog Selection
    End If
End Sub
