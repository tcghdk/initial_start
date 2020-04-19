Attribute VB_Name = "Module5"
Public Sub 重ね合わせるグラフ任意のデータ点数で200315()

Dim GraphSize_X As Integer 'グラフX方向サイズ
Dim GraphSize_Y As Integer 'グラフY方向サイズ
Dim GraphLength As Integer 'グラフデータの最大行
Dim Max_SheetNum As Integer '数値まとめシート数

Dim start As Variant '重ね合わせのはじめ行
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    start = Application.InputBox( _
                    "重ね合わせのはじめの行を入力してください。", Type:=1)
    If TypeName(start) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If
Dim data_num As Variant '重ね合わせデータ数
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    data_num = Application.InputBox( _
                    "重ね合わせるデータ数を入力してください。", Type:=1)
    If TypeName(data_num) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If
Dim keiretu_num As Variant '重ね合わせる系列数
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    keiretu_num = Application.InputBox( _
                    "重ね合わせる系列数を入力してください。", Type:=1)
    If TypeName(keiretu_num) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If
Dim x As Variant 'x軸の列数
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    x = Application.InputBox( _
                    "x軸の列数を入力してください。", Type:=1)
    If TypeName(x) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If
Dim y As Variant 'y軸の列数
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    y = Application.InputBox( _
                    "y軸の列数を入力してください。", Type:=1)
    If TypeName(y) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If

'グラフX方向サイズ
GraphSize_X = 270.321502685547
'グラフY方向サイズ
GraphSize_Y = 184.821411132813
'グラフデータの最大行
'GraphLength = Sheets("1").Range("A65535").End(x1Up).Row
GraphLength = 15
'数値まとめシート数
Max_SheetNum = 11

Dim i As Integer
Dim j As Integer


With ActiveSheet.Shapes.AddChart.Chart
'グラフの種類
.ChartType = xlXYScatter


'データ追加
j = start + data_num - 1
For i = 1 To keiretu_num
.SeriesCollection.NewSeries
With .SeriesCollection(i)
.Name = Sheets(CStr(1)).Cells(j - data_num + 1, 1)
.XValues = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(j - data_num + 1, x), Sheets(CStr(1)).Cells(j, x))
.Values = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(j - data_num + 1, y), Sheets(CStr(1)).Cells(j, y))
End With
    j = j + data_num
Next i



'With ActiveSheet.Shapes.AddChart.Chart
'
''グラフの種類選択
''.ChartType = xlXYScatter
'
'''データプロット
''.SetSourceData Source:=Range("B1:B15", "C1:C15")
'''例：.SetSourceData Source:=Sheets("Sheet1").Range("A2:H7")
''.FullSeriesCollection(1).Name = "=Sheet1!A1"
'
''データプロット
''With ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Chart
'''ActiveChart.SetSourceData Source:=Range("Sheet1!A1:B15")
''.SetSourceData Source:=Range("A1:A15", "B1:B15")
'
''データプロット
''Set Series = .SeriesCollection.Add(Range("B1:B15", "D1:D15"))
''With Series
'''.ChartType = xlXYScatter
''.Name = Range("A2")
''End With
'
''データプロット
''With ActiveSheet.Shapes.AddChart.Chart
'''散布図追加
''.ChartType = xlXYScatter 'グラフの種類
''.SetSourceData Range("A1:A15", "C1:C15") 'グラフの範囲
'
''データ追加
''.SeriesCollection.NewSeries
''.FullSeriesCollection(1).Name = "=Sheet1!A1"
''.FullSeriesCollection(1).XValues = "=Sheet1!B1:b15"
''.FullSeriesCollection(1).Values = "=Sheet1!C1:C15"
'
''.FullSeriesCollection(1).Select
''With Selection
''マーカーの種類の設定
''.MarkerStyle = 8
''マーカーサイズの設定
''.MarkerSize = 7
''マーカーの塗りつぶし設定 msoTrueで塗りつぶし
''.Format.Fill.Visible = msofulse
''マーカーの色設定
''.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
''マーカーを線でつなぐ
''.Format.Line.Visible = msoTrue
''マーカーを消す？
''.MarkerForegroundColorIndex = xlColorIndexNone
''-------------------------------------------------
''.Border.Weight = 2.5 'xlThin　　　　　　'線の太さ
''.Border.LineStyle = xlContinuous      '線種
''.Border.LineStyle = 1
''.Border.ColorIndex = 1 '線の色
''.MarkerBackgroundColorIndex = 7      'マーカーの色　塗りつぶし？
''.MarkerForegroundColorIndex = 5      'マーカーの色　枠線？
''マーカーの枠線変更
''.Format.Line.Visible = msoFalse
''.Format.Line.Weight = 0
''.Border.Weight = 4
''.MarkerStyle = xlDiamond              'マーカーの形
''.Smooth = True
'' .MarkerSize = 7 'マーカーのサイズ
''.Shadow = False                      '影（True or False）
''End With
'
'
'
'''データ追加
''.SeriesCollection.NewSeries
''.FullSeriesCollection(2).Name = "=Sheet1!A2"
''.FullSeriesCollection(2).XValues = "=Sheet1!B16:B30"
''.FullSeriesCollection(2).Values = "=Sheet1!C16:C30"
''
'''データ追加
''.SeriesCollection.NewSeries
''.FullSeriesCollection(3).Name = "=Sheet1!A3"
''.FullSeriesCollection(3).XValues = "=Sheet1!B31:B45"
''.FullSeriesCollection(3).Values = "=Sheet1!C31:C45"
'
'
'グラフタイトル追加
.HasTitle = True 'グラフタイトル追加
.ChartTitle.Text = "記入してください" 'グラフタイトル変更value
.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11

'ラベル追加
'縦軸
.Axes(xlValue).HasTitle = True
.Axes(xlValue).AxisTitle.Text = "記入してください"
.Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
.Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 11
'横軸
.Axes(xlCategory).HasTitle = True
.Axes(xlCategory).AxisTitle.Text = "記入してください"
.Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
.Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 11

'軸の最小値・最大値変更
'縦軸
.Axes(xlValue).MinimumScale = 0
.Axes(xlValue).MaximumScale = 50
'横軸
.Axes(xlCategory).MinimumScale = 0
.Axes(xlCategory).MaximumScale = 20

''行と列を入れ替える
'Select Case .PlotBy
'    Case xlRows
'        .PlotBy = xlColumns
'    Case xlColumns
'        .PlotBy = xlRows
'End Select

'グラフの枠線を黒色で太くする
.PlotArea.Format.Line.Visible = msoTrue
.PlotArea.Format.Line.ForeColor.ObjectThemeColor = msoThemeColorText1
.PlotArea.Format.Line.ForeColor.TintAndShade = 0
.PlotArea.Format.Line.ForeColor.Brightness = 0
.PlotArea.Format.Line.Weight = 2

'軸の枠線を消す
.Axes(xlValue).Format.Line.Visible = msoFalse
.Axes(xlCategory).Format.Line.Visible = msoFalse

End With


'If .HasLegend = False Then .HasLegend = True    ''凡例を表示する
'    .Legend.Position = xlLegendPositionTop          ''凡例を上に表示する
'    .Legend.IncludeInLayout = False                 ''凡例をグラフに重ねる
'With .Legend.Format.Fill
'    .Visible = msoTrue                          ''凡例を塗りつぶします
'    .ForeColor.RGB = RGB(255, 0, 0)             ''赤色
'    .ForeColor.TintAndShade = 0.5               ''明暗の設定
'End With

With ActiveSheet.ChartObjects
    .Top = Range("E5").Top
    .Left = Range("E5").Left '位置を設定
    .Height = 200
    .Width = 300 '大きさを設定
End With


End Sub

