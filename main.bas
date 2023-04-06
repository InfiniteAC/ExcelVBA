Rem 计算误差，标准差
Function PhyEvp(rng As Range)
    Dim DataNum As Integer
    Dim Mean As Double
    DataNum = rng.Rows.Count
    For i = 0 To DataNum - 1
    Mean = Cells(rng.Row + i, rng.Column) + Mean
    Next
    Mean = Mean / DataNum
    Dim SquareAdd As Double
    For i = 0 To DataNum - 1
    SquareAdd = (Cells(rng.Row + i, rng.Column) - Mean) ^ 2 + SquareAdd
    Next
    PhyEvp = (SquareAdd / (DataNum - 1)) ^ 0.5
End Function
