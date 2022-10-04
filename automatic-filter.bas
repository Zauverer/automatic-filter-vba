Option Explicit

Private Sub CommandButton1_Click()
    Range("A9").CurrentRegion.AutoFilter
    Hoja1.TextBox1.Value = ""
        
End Sub

Private Sub TextBox1_Change()

Dim Criterio As String

    If Hoja1.TextBox1.Value <> "" Then
        Criterio = "*" & Hoja1.TextBox1.Value & "*"
        
     Range("A7").CurrentRegion.AutoFilter Field:=2, Criteria1:=Criterio
    
    Else
        Criterio = ""
        Range("A7").CurrentRegion.AutoFilter
    End If

End Sub
