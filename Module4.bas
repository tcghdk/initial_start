Attribute VB_Name = "Module4"


Public Sub �����̃O���t���쐬����P�i�̂�����s�ɂ�������200315()

Dim GraphSize_X As Integer '�O���tX�����T�C�Y
Dim GraphSize_Y As Integer '�O���tY�����T�C�Y
'Dim GraphLength As Integer '�O���t�f�[�^�̍ő�s
Dim Max_SheetNum As Integer '���l�܂Ƃ߃V�[�g��
'Dim start As Integer '�d�ˍ��킹�̂͂��ߍs
'Dim data_num As Integer '�d�ˍ��킹�f�[�^��
'Dim x As Integer 'x���̗�
'
'start = 1
'data_num = 15
'x = 2

'�O���tX�����T�C�Y
GraphSize_X = 270.321502685547
'�O���tY�����T�C�Y
GraphSize_Y = 184.821411132813
'�O���t�f�[�^�̍ő�s
'GraphLength = Sheets("1").Range("A65535").End(x1Up).Row
'GraphLength = 15
'���l�܂Ƃ߃V�[�g��
Max_SheetNum = 11

Dim i As Integer
Dim j As Integer

Dim start As Integer '�d�ˍ��킹�̂͂��ߍs
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    start = Application.InputBox( _
                    "�d�ˍ��킹�̂͂��߂̍s����͂��Ă��������B", Type:=1)
    If TypeName(start) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If
Dim GraphLength As Integer '�d�ˍ��킹�f�[�^��
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    GraphLength = Application.InputBox( _
                    "�d�ˍ��킹��f�[�^������͂��Ă��������B", Type:=1)
    If TypeName(GraphLength) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If
Dim x As Integer 'x���̗�
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    x = Application.InputBox( _
                    "x���̗񐔂���͂��Ă��������B", Type:=1)
    If TypeName(x) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If
Dim y As Integer 'y���̂͂��߂̗�
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    y = Application.InputBox( _
                    "y���̂͂��߂̗񐔂���͂��Ă��������B", Type:=1)
    If TypeName(y) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If




'��̐������O���t�쐬
j = 1
For i = Sheets("1").Range("B1").Offset(0, 0).Column To Sheets("1").Range("A1").End(xlToRight).Offset(0, -1).Column
    If j Mod 16 <> 0 Then
    
''�O���t�o���ʒu�ƃT�C�Y���Z�b�g
'With ActiveSheet.ChartObjects.Add(, graph_pos.Top, GraphSize_X, GraphSize_Y).Chart
''Cells(2 + 13*((j-1)\4),2+5*((j-1)Mod4))
''�O���t�̎��
'.ChartType = xlXYScatter
''�n��f�[�^�̃v���b�g
''�f�[�^�ǉ�
'.SeriesCollection.NewSeries
'.FullSeriesCollection(1).Name = Cells(1, 1)
'.FullSeriesCollection(1).XValues = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(start, x), Sheets(CStr(1)).Cells(GraphLength, x))
'.FullSeriesCollection(1).Values = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(start, 3), Sheets(CStr(1)).Cells(GraphLength, 3))
'End With
    
    Call makeGraph(i, Cells(2 + 13 * ((j - 1) \ 4), 2 + 5 * ((j - 1) Mod 4)), start, GraphLength, x, y)
    
    Else
    
        i = i - 1
    End If
    j = j + 1
    y = y + 1
Next i







'With ActiveSheet.Shapes.AddChart.Chart
''
'''�O���t�̎�ޑI��
''.ChartType = xlXYScatter
''
''''�f�[�^�v���b�g
'''.SetSourceData Source:=Range("B1:B15", "C1:C15")
''''��F.SetSourceData Source:=Sheets("Sheet1").Range("A2:H7")
'''.FullSeriesCollection(1).Name = "=Sheet1!A1"
''
'''�f�[�^�v���b�g
'''With ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Chart
''''ActiveChart.SetSourceData Source:=Range("Sheet1!A1:B15")
'''.SetSourceData Source:=Range("A1:A15", "B1:B15")
''
'''�f�[�^�v���b�g
'''Set Series = .SeriesCollection.Add(Range("B1:B15", "D1:D15"))
'''With Series
''''.ChartType = xlXYScatter
'''.Name = Range("A2")
'''End With
''
'''�f�[�^�v���b�g
'''With ActiveSheet.Shapes.AddChart.Chart
''''�U�z�}�ǉ�
'''.ChartType = xlXYScatter '�O���t�̎��
'''.SetSourceData Range("A1:A15", "C1:C15") '�O���t�͈̔�
''
''Dim k As String
''k = CStr(1)
''
''
'''�f�[�^�ǉ�
''.SeriesCollection.NewSeries
''.FullSeriesCollection(1).Name = Cells(1, 1)
''.FullSeriesCollection(1).XValues = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(start, x), Sheets(CStr(1)).Cells(GraphLength, x)) '"=Sheet1!B1:b15"
''.FullSeriesCollection(1).Values = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(start, 3), Sheets(CStr(1)).Cells(GraphLength, 3)) '"=Sheet1!C1:C15"
''
''
''
'''.FullSeriesCollection(1).Select
'''With Selection
'''�}�[�J�[�̎�ނ̐ݒ�
'''.MarkerStyle = 8
'''�}�[�J�[�T�C�Y�̐ݒ�
'''.MarkerSize = 7
'''�}�[�J�[�̓h��Ԃ��ݒ� msoTrue�œh��Ԃ�
'''.Format.Fill.Visible = msofulse
'''�}�[�J�[�̐F�ݒ�
'''.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
'''�}�[�J�[����łȂ�
'''.Format.Line.Visible = msoTrue
'''�}�[�J�[�������H
'''.MarkerForegroundColorIndex = xlColorIndexNone
'''-------------------------------------------------
'''.Border.Weight = 2.5 'xlThin�@�@�@�@�@�@'���̑���
'''.Border.LineStyle = xlContinuous      '����
'''.Border.LineStyle = 1
'''.Border.ColorIndex = 1 '���̐F
'''.MarkerBackgroundColorIndex = 7      '�}�[�J�[�̐F�@�h��Ԃ��H
'''.MarkerForegroundColorIndex = 5      '�}�[�J�[�̐F�@�g���H
'''�}�[�J�[�̘g���ύX
'''.Format.Line.Visible = msoFalse
'''.Format.Line.Weight = 0
'''.Border.Weight = 4
'''.MarkerStyle = xlDiamond              '�}�[�J�[�̌`
'''.Smooth = True
''' .MarkerSize = 7 '�}�[�J�[�̃T�C�Y
'''.Shadow = False                      '�e�iTrue or False�j
'''End With
''
''
''
'''�f�[�^�ǉ�
''.SeriesCollection.NewSeries
''.FullSeriesCollection(2).Name = "=Sheet1!A2"
''.FullSeriesCollection(2).XValues = "=Sheet1!B16:B30"
''.FullSeriesCollection(2).Values = "=Sheet1!C16:C30"
''
'''�f�[�^�ǉ�
''.SeriesCollection.NewSeries
''.FullSeriesCollection(3).Name = "=Sheet1!A3"
''.FullSeriesCollection(3).XValues = "=Sheet1!B31:B45"
''.FullSeriesCollection(3).Values = "=Sheet1!C31:C45"
'
'
''�O���t�^�C�g���ǉ�
'.HasTitle = True '�O���t�^�C�g���ǉ�
'.ChartTitle.Text = "�L�����Ă�������" '�O���t�^�C�g���ύXvalue
'.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
'
''���x���ǉ�
''�c��
'.Axes(xlValue).HasTitle = True
'.Axes(xlValue).AxisTitle.Text = "�L�����Ă�������"
'.Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'.Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 11
''����
'.Axes(xlCategory).HasTitle = True
'.Axes(xlCategory).AxisTitle.Text = "�L�����Ă�������"
'.Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'.Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 11
'
''���̍ŏ��l�E�ő�l�ύX
''�c��
'.Axes(xlValue).MinimumScale = 0
'.Axes(xlValue).MaximumScale = 50
''����
'.Axes(xlCategory).MinimumScale = 0
'.Axes(xlCategory).MaximumScale = 20
'
'''�s�Ɨ�����ւ���
''Select Case .PlotBy
''    Case xlRows
''        .PlotBy = xlColumns
''    Case xlColumns
''        .PlotBy = xlRows
''End Select
'
''�O���t�̘g�������F�ő�������
'.PlotArea.Format.Line.Visible = msoTrue
'.PlotArea.Format.Line.ForeColor.ObjectThemeColor = msoThemeColorText1
'.PlotArea.Format.Line.ForeColor.TintAndShade = 0
'.PlotArea.Format.Line.ForeColor.Brightness = 0
'.PlotArea.Format.Line.Weight = 2
'
''���̘g��������
'.Axes(xlValue).Format.Line.Visible = msoFalse
'.Axes(xlCategory).Format.Line.Visible = msoFalse
'
'End With
'
'
''If .HasLegend = False Then .HasLegend = True    ''�}���\������
''    .Legend.Position = xlLegendPositionTop          ''�}�����ɕ\������
''    .Legend.IncludeInLayout = False                 ''�}����O���t�ɏd�˂�
''With .Legend.Format.Fill
''    .Visible = msoTrue                          ''�}���h��Ԃ��܂�
''    .ForeColor.RGB = RGB(255, 0, 0)             ''�ԐF
''    .ForeColor.TintAndShade = 0.5               ''���Â̐ݒ�
''End With

'With ActiveSheet.ChartObjects
'    .Top = Range("E5").Top
'    .Left = Range("E5").Left '�ʒu��ݒ�
'    .Height = 200
'    .Width = 300 '�傫����ݒ�
'End With


End Sub




Public Sub makeGraph(data_col As Integer, graph_pos As Range, start As Integer, GraphLength As Integer, x As Integer, y As Integer)

Dim j As Integer
Dim linecolor As Variant
Dim graph As ChartObject

'Dim start As Integer '�d�ˍ��킹�̂͂��ߍs
'''���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
''    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
''    start = Application.InputBox( _
''                    "�d�ˍ��킹�̂͂��߂̍s����͂��Ă��������B", Type:=1)
''    If TypeName(start) = "Boolean" Then
''        MsgBox "���͂��L�����Z������܂����B", vbExclamation
''    End If
'Dim GraphLength As Integer '�d�ˍ��킹�f�[�^��
'''���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
''    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
''    GraphLength = Application.InputBox( _
''                    "�d�ˍ��킹��f�[�^������͂��Ă��������B", Type:=1)
''    If TypeName(GraphLength) = "Boolean" Then
''        MsgBox "���͂��L�����Z������܂����B", vbExclamation
''    End If
'Dim x As Integer 'x���̗�
'''���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
''    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
''    x = Application.InputBox( _
''                    "x���̗񐔂���͂��Ă��������B", Type:=1)
''    If TypeName(x) = "Boolean" Then
''        MsgBox "���͂��L�����Z������܂����B", vbExclamation
''    End If
'Dim y As Integer 'y���̂͂��߂̗�
'''���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
''    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
''    y = Application.InputBox( _
''                    "y���̂͂��߂̗񐔂���͂��Ă��������B", Type:=1)
''    If TypeName(y) = "Boolean" Then
''        MsgBox "���͂��L�����Z������܂����B", vbExclamation
''    End If


Dim GraphSize_X As Integer '�O���tX�����T�C�Y
Dim GraphSize_Y As Integer '�O���tY�����T�C�Y
'�O���tX�����T�C�Y
GraphSize_X = 270.321502685547
'�O���tY�����T�C�Y
GraphSize_Y = 184.821411132813


'MsgBox (GraphSize_Y)


'�O���t�o���ʒu�ƃT�C�Y���Z�b�g
With ActiveSheet.ChartObjects.Add(graph_pos.Left, graph_pos.Top, GraphSize_X, GraphSize_Y).Chart



'�O���t�̎��
.ChartType = xlXYScatter

'�n��f�[�^�̃v���b�g
'�f�[�^�ǉ�
.SeriesCollection.NewSeries
.FullSeriesCollection(1).Name = Cells(1, 1)
.FullSeriesCollection(1).XValues = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(start, x), Sheets(CStr(1)).Cells(GraphLength, x))
.FullSeriesCollection(1).Values = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(start, y), Sheets(CStr(1)).Cells(GraphLength, y))

End With


End Sub
