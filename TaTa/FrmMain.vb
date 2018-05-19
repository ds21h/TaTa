Public Class FrmMain

  Private Sub BtnCalc_Click(sender As Object, e As EventArgs) Handles BtnCalc.Click
    Dim lTot As Integer
    Dim lFraction As Double
    Dim lStandard As Integer

    Try
      lStandard = CInt(TxtStand.Text)
    Catch ex As Exception
      lStandard = 0
    End Try

    lTot = 0
    lTot = lTot + sCalcLine(0, TxtNumb0, TxtWidth0, LbTot0)
    lTot = lTot + sCalcLine(1, TxtNumb1, TxtWidth1, LbTot1)
    lTot = lTot + sCalcLine(2, TxtNumb2, TxtWidth2, LbTot2)
    lTot = lTot + sCalcLine(3, TxtNumb3, TxtWidth3, LbTot3)
    lTot = lTot + sCalcLine(4, TxtNumb4, TxtWidth4, LbTot4)
    lTot = lTot + sCalcLine(5, TxtNumb5, TxtWidth5, LbTot5)
    lTot = lTot + sCalcLine(6, TxtNumb6, TxtWidth6, LbTot6)
    lTot = lTot + sCalcLine(7, TxtNumb7, TxtWidth7, LbTot7)
    lTot = lTot + sCalcLine(8, TxtNumb8, TxtWidth8, LbTot8)
    lTot = lTot + sCalcLine(9, TxtNumb9, TxtWidth9, LbTot9)

    LbTot.Text = CStr(lTot)
    If lStandard = 0 Then
      lFraction = -1
    Else
      lFraction = lTot / lStandard
    End If
    LbFraction.Text = Format(lFraction, "#0.0000")
  End Sub

  Private Function sCalcLine(pLine As Integer, pNumber As TextBox, pWidth As TextBox, pTotal As Label) As Integer
    Dim lNumber As Integer
    Dim lWidth As Integer
    Dim lTot As Integer

    Try
      lNumber = CInt(pNumber.Text)
      lWidth = CInt(pWidth.Text)
    Catch ex As Exception
      lNumber = 0
      lWidth = 0
    End Try

    lTot = lNumber * lWidth
    pTotal.Text = CStr(lTot)

    Return lTot
  End Function

End Class
