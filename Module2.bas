Attribute VB_Name = "Module2"
Public Sub グラフ書式1番目のみ()

With ActiveSheet.ChartObjects(1).Chart

Set Object = Selection

.FullSeriesCollection(1).Select
With Selection
'マーカーの種類の設定
.MarkerStyle = 8
'マーカーサイズの設定
.MarkerSize = 7
'マーカーの有無
.Format.Fill.Visible = msoFalse
'マーカーの色設定
.Format.Fill.ForeColor.RGB = vbRed
'マーカーを線でつなぐ_trueで線消える
.Format.Line.Visible = msoFalse
'-------------------------------------------------
'線種
.Border.LineStyle = xlContinuous
'線の色
.Border.ColorIndex = 1
''マーカー塗りつぶしの色
'.MarkerBackgroundColorIndex = 6
''マーカー枠線の色
'.MarkerForegroundColorIndex = 8
''マーカーの枠線を消す　上の色の定義の後に置かないと枠線消えない
'.MarkerForegroundColorIndex = xlColorIndexNone
'マーカーの枠線変更
.Format.Line.Visible = msoTrue
.Format.Line.Weight = 1
'線の太さ
.Border.Weight = 4
'平滑線でなめらかにするか直線で補間するか
.Smooth = True
'影（True or False）falseの方が良い
.Shadow = False
End With


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
.Axes(xlValue).MaximumScale = 80
'横軸
.Axes(xlCategory).MinimumScale = 0
.Axes(xlCategory).MaximumScale = 60

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


If .HasLegend = False Then .HasLegend = True    ''凡例を表示する
    .Legend.Position = xlLegendPositionTop          ''凡例を上に表示する
    .Legend.IncludeInLayout = False                 ''凡例をグラフに重ねる
With .Legend.Format.Fill
    .Visible = msoTrue                          ''凡例を塗りつぶします
    .ForeColor.RGB = vbWhite             ''赤色
    .ForeColor.TintAndShade = 1               ''明暗の設定
End With

End With

With ActiveSheet.ChartObjects
    .Top = Range("E5").Top
    .Left = Range("E5").Left '位置を設定
    .Height = 200
    .Width = 300 '大きさを設定
End With


End Sub



Public Sub 書式()

Dim x As Object
Dim xi As Object

    Set x = Selection
    If TypeName(x) <> "DrawingObjects" Then
        If Not ActiveChart Is Nothing Then
                 
                With ActiveChart
                
                '.FullSeriesCollection(1).Select
                'With Selection
                ''マーカーの種類の設定
                '.MarkerStyle = 8
                ''マーカーサイズの設定
                '.MarkerSize = 7
                ''マーカーの塗りつぶし設定 msoTrueで塗りつぶし
                '.Format.Fill.Visible = msofulse
                ''マーカーの色設定
                '.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                ''マーカーを線でつなぐ
                '.Format.Line.Visible = msoTrue
                ''マーカーを消す？
                '.MarkerForegroundColorIndex = xlColorIndexNone
                ''-------------------------------------------------
                '.Border.Weight = 2.5 'xlThin　　　　　　'線の太さ
                '.Border.LineStyle = xlContinuous      '線種
                '.Border.LineStyle = 1
                '.Border.ColorIndex = 1 '線の色
                '.MarkerBackgroundColorIndex = 7      'マーカーの色　塗りつぶし？
                '.MarkerForegroundColorIndex = 5      'マーカーの色　枠線？
                ''マーカーの枠線変更
                '.Format.Line.Visible = msoFalse
                '.Format.Line.Weight = 0
                '.Border.Weight = 4
                '.MarkerStyle = xlDiamond              'マーカーの形
                '.Smooth = True
                ' .MarkerSize = 7 'マーカーのサイズ
                '.Shadow = False                      '影（True or False）
                'End With
                
                
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
                .Axes(xlValue).MaximumScale = 80
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
                
                
                'If .HasLegend = False Then .HasLegend = True    ''凡例を表示する
                '    .Legend.Position = xlLegendPositionTop          ''凡例を上に表示する
                '    .Legend.IncludeInLayout = False                 ''凡例をグラフに重ねる
                'With .Legend.Format.Fill
                '    .Visible = msoTrue                          ''凡例を塗りつぶします
                '    .ForeColor.RGB = RGB(255, 0, 0)             ''赤色
                '    .ForeColor.TintAndShade = 0.5               ''明暗の設定
                'End With
                
                End With
                
                With ActiveSheet.ChartObjects
                    .Top = Range("E5").Top
                    .Left = Range("E5").Left '位置を設定
                    .Height = 200
                    .Width = 300 '大きさを設定
                End With
            
        End If
    Else
        For Each xi In x
            If TypeName(xi) = "ChartObject" Then
                
                With xi.Chart
                
                '.FullSeriesCollection(1).Select
                'With Selection
                ''マーカーの種類の設定
                '.MarkerStyle = 8
                ''マーカーサイズの設定
                '.MarkerSize = 7
                ''マーカーの塗りつぶし設定 msoTrueで塗りつぶし
                '.Format.Fill.Visible = msofulse
                ''マーカーの色設定
                '.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                ''マーカーを線でつなぐ
                '.Format.Line.Visible = msoTrue
                ''マーカーを消す？
                '.MarkerForegroundColorIndex = xlColorIndexNone
                ''-------------------------------------------------
                '.Border.Weight = 2.5 'xlThin　　　　　　'線の太さ
                '.Border.LineStyle = xlContinuous      '線種
                '.Border.LineStyle = 1
                '.Border.ColorIndex = 1 '線の色
                '.MarkerBackgroundColorIndex = 7      'マーカーの色　塗りつぶし？
                '.MarkerForegroundColorIndex = 5      'マーカーの色　枠線？
                ''マーカーの枠線変更
                '.Format.Line.Visible = msoFalse
                '.Format.Line.Weight = 0
                '.Border.Weight = 4
                '.MarkerStyle = xlDiamond              'マーカーの形
                '.Smooth = True
                ' .MarkerSize = 7 'マーカーのサイズ
                '.Shadow = False                      '影（True or False）
                'End With
                
                
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
                .Axes(xlValue).MaximumScale = 40
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
                
                
                'If .HasLegend = False Then .HasLegend = True    ''凡例を表示する
                '    .Legend.Position = xlLegendPositionTop          ''凡例を上に表示する
                '    .Legend.IncludeInLayout = False                 ''凡例をグラフに重ねる
                'With .Legend.Format.Fill
                '    .Visible = msoTrue                          ''凡例を塗りつぶします
                '    .ForeColor.RGB = RGB(255, 0, 0)             ''赤色
                '    .ForeColor.TintAndShade = 0.5               ''明暗の設定
                'End With
                
                End With
                
                With ActiveSheet.ChartObjects
                    .Top = Range("E5").Top
                    .Left = Range("E5").Left '位置を設定
                    .Height = 200
                    .Width = 300 '大きさを設定
                End With
                
            End If
        Next
    End If
  
    Set xi = Nothing

End Sub
