Attribute VB_Name = "Module6"
Public Sub 任意のデータ点数で重ね合わせ各列複数グラフ作成200329()

Dim i As Integer 'カウンタ
Dim j As Integer 'カウンタ

Dim start As Integer '重ね合わせのはじめ行
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    start = Application.InputBox( _
                    "重ね合わせのはじめの行を入力してください。", Type:=1)
    If TypeName(start) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If
    
Dim GraphLength As Integer '重ね合わせデータ数
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    GraphLength = Application.InputBox( _
                    "重ね合わせるデータ数を入力してください。", Type:=1)
    If TypeName(GraphLength) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If
    
Dim keiretu_num As Integer '重ね合わせる系列数
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    keiretu_num = Application.InputBox( _
                    "重ね合わせる系列数を入力してください。", Type:=1)
    If TypeName(keiretu_num) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If
    
Dim x As Integer 'x軸の列数
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    x = Application.InputBox( _
                    "x軸の列数を入力してください。", Type:=1)
    If TypeName(x) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If
    
Dim y As Integer 'y軸のはじめの列数
'入力キャンセルされるとブール値が返ってくるので、
    'バリアント型変数で受けます。
    y = Application.InputBox( _
                    "y軸のはじめの列数を入力してください。", Type:=1)
    If TypeName(y) = "Boolean" Then
        MsgBox "入力がキャンセルされました。", vbExclamation
    End If

'列の数だけグラフ作成
j = 1
For i = Sheets("1").Range("B1").Offset(0, 0).Column To Sheets("1").Range("A1").End(xlToRight).Offset(0, -1).Column
    If j Mod 16 <> 0 Then
    Call makeGraph(i, Cells(2 + 13 * ((j - 1) \ 4), 2 + 5 * ((j - 1) Mod 4)), start, GraphLength, x, y, keiretu_num)
    Else
        i = i - 1
    End If
    j = j + 1
    y = y + 1
Next i

End Sub

Public Sub makeGraph(data_col As Integer, graph_pos As Range, start As Integer, GraphLength As Integer, x As Integer, y As Integer, keiretu_num As Integer)

Dim i As Integer
Dim j As Integer
Dim GraphSize_X As Integer 'グラフX方向サイズ
Dim GraphSize_Y As Integer 'グラフY方向サイズ

'グラフX方向サイズ
GraphSize_X = 270.321502685547
'グラフY方向サイズ
GraphSize_Y = 184.821411132813

'グラフ出現位置とサイズをセット
With ActiveSheet.ChartObjects.Add(graph_pos.Left, graph_pos.Top, GraphSize_X, GraphSize_Y).Chart '△
'グラフの種類
.ChartType = xlXYScatter

'系列データのプロット
'データ追加
j = start + GraphLength - 1
For i = 1 To keiretu_num
    .SeriesCollection.NewSeries
With .SeriesCollection(i)
    .Name = Sheets(CStr(1)).Cells(j - GraphLength + 1, 1)
    .XValues = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(j - GraphLength + 1, x), Sheets(CStr(1)).Cells(j, x))
    .Values = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(j - GraphLength + 1, y), Sheets(CStr(1)).Cells(j, y))
End With
    j = j + GraphLength
    
'----グラフ書式変更----
'グラフタイトル追加
.HasTitle = True 'グラフタイトル追加
.ChartTitle.Text = "記入してください"
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
''軸の最小値・最大値変更
''縦軸
'.Axes(xlValue).MinimumScale = 0
'.Axes(xlValue).MaximumScale = 200
''横軸
'.Axes(xlCategory).MinimumScale = 0
'.Axes(xlCategory).MaximumScale = 20
'グラフの枠線を黒色で太くする（作法）
.PlotArea.Format.Line.Visible = msoTrue
.PlotArea.Format.Line.ForeColor.ObjectThemeColor = msoThemeColorText1
.PlotArea.Format.Line.ForeColor.TintAndShade = 0
.PlotArea.Format.Line.ForeColor.Brightness = 0
.PlotArea.Format.Line.Weight = 2
'軸の枠線を消す(作法)
.Axes(xlValue).Format.Line.Visible = msoFalse
.Axes(xlCategory).Format.Line.Visible = msoFalse
'凡例設定
'凡例を表示する
.HasLegend = True
'凡例を上に表示する
.Legend.Position = xlLegendPositionTop
'凡例をグラフに重ねる
.Legend.IncludeInLayout = True
With .Legend.Format.Fill
'凡例を塗りつぶします
    .Visible = msoTrue
'凡例の塗りつぶしの色
    .ForeColor.RGB = vbWhite
'明暗の設定
    .ForeColor.TintAndShade = 0.5
End With
'--------------------
.FullSeriesCollection(i).Select '〇
With Selection
'マーカーの種類の設定
.MarkerStyle = 8
'マーカーのサイズ
.MarkerSize = 7
'マーカーの有無
.Format.Fill.Visible = msoTrue
'マーカーの色設定
'.Format.Fill.ForeColor.RGB = vbRed
'マーカーの枠線変更
.Format.Line.Visible = msoTrue
.Format.Line.Weight = 0
'マーカー塗りつぶしの色
.MarkerBackgroundColorIndex = -4105
'マーカー枠線の色
.MarkerForegroundColorIndex = -4105
'マーカーの枠線を消す_上の色の定義の後に置かないと枠線消えない
.MarkerForegroundColorIndex = xlColorIndexNone
'マーカーを線でつなぐ
.Format.Line.Visible = msoFalse
'線の太さ
'.Border.Weight = 4 'xlThin
'線種
'.Border.LineStyle = xlContinuous
'線の色
'.Border.ColorIndex = -4105
'平滑線でなめらかにするか直線で補間するか
.Smooth = True
'影 (True Or False)falseの方が良い。trueだと線がずれる？
.Shadow = False
End With '〇
'----------------------
''行と列を入れ替える　←このプログラムだとうまくいかない・・・
'Select Case .PlotBy
'    Case xlRows
'        .PlotBy = xlColumns
'    Case xlColumns
'        .PlotBy = xlRows
'End Select

Next i

End With '△

End Sub
