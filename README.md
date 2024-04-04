Private Sub Worksheet_Change(ByVal Target As Range)
    'エラー処理
    If Target.Count > 1 Then Exit Sub
    
    If Not Intersect(Target, Range("B2:B65536")) Is Nothing Then
        '前回更新日と更新日を比較
        If Target > Cells(Target.Row, Target.Column - 1) Then
            '更新されていれば
            '対応行(row)の列(column)C(前回回数) + 1を現回数へ書き込み
            Cells(Target.Row, Target.Column + 2) = Cells(Target.Row, Target.Column + 1) + 1
            
            '現回数を前回回数へ書き込み
            Cells(Target.Row, Target.Column + 1) = Cells(Target.Row, Target.Column + 2)
            
            '更新日を前回更新日へ書き込み
            Cells(Target.Row, Target.Column - 1) = Target
        End If
    End If
End Sub
