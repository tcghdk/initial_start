Attribute VB_Name = "Module5"
Public Sub �d�ˍ��킹��O���t�C�ӂ̃f�[�^�_����200315()

Dim GraphSize_X As Integer '�O���tX�����T�C�Y
Dim GraphSize_Y As Integer '�O���tY�����T�C�Y
Dim GraphLength As Integer '�O���t�f�[�^�̍ő�s
Dim Max_SheetNum As Integer '���l�܂Ƃ߃V�[�g��

Dim start As Variant '�d�ˍ��킹�̂͂��ߍs
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    start = Application.InputBox( _
                    "�d�ˍ��킹�̂͂��߂̍s����͂��Ă��������B", Type:=1)
    If TypeName(start) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If
Dim data_num As Variant '�d�ˍ��킹�f�[�^��
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    data_num = Application.InputBox( _
                    "�d�ˍ��킹��f�[�^������͂��Ă��������B", Type:=1)
    If TypeName(data_num) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If
Dim keiretu_num As Variant '�d�ˍ��킹��n��
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    keiretu_num = Application.InputBox( _
                    "�d�ˍ��킹��n�񐔂���͂��Ă��������B", Type:=1)
    If TypeName(keiretu_num) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If
Dim x As Variant 'x���̗�
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    x = Application.InputBox( _
                    "x���̗񐔂���͂��Ă��������B", Type:=1)
    If TypeName(x) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If
Dim y As Variant 'y���̗�
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    y = Application.InputBox( _
                    "y���̗񐔂���͂��Ă��������B", Type:=1)
    If TypeName(y) = "Boolean" Then
        MsgBox "���͂��L�����Z������܂����B", vbExclamation
    End If

'�O���tX�����T�C�Y
GraphSize_X = 270.321502685547
'�O���tY�����T�C�Y
GraphSize_Y = 184.821411132813
'�O���t�f�[�^�̍ő�s
'GraphLength = Sheets("1").Range("A65535").End(x1Up).Row
GraphLength = 15
'���l�܂Ƃ߃V�[�g��
Max_SheetNum = 11

Dim i As Integer
Dim j As Integer


With ActiveSheet.Shapes.AddChart.Chart
'�O���t�̎��
.ChartType = xlXYScatter


'�f�[�^�ǉ�
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
''�O���t�̎�ޑI��
''.ChartType = xlXYScatter
'
'''�f�[�^�v���b�g
''.SetSourceData Source:=Range("B1:B15", "C1:C15")
'''��F.SetSourceData Source:=Sheets("Sheet1").Range("A2:H7")
''.FullSeriesCollection(1).Name = "=Sheet1!A1"
'
''�f�[�^�v���b�g
''With ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Chart
'''ActiveChart.SetSourceData Source:=Range("Sheet1!A1:B15")
''.SetSourceData Source:=Range("A1:A15", "B1:B15")
'
''�f�[�^�v���b�g
''Set Series = .SeriesCollection.Add(Range("B1:B15", "D1:D15"))
''With Series
'''.ChartType = xlXYScatter
''.Name = Range("A2")
''End With
'
''�f�[�^�v���b�g
''With ActiveSheet.Shapes.AddChart.Chart
'''�U�z�}�ǉ�
''.ChartType = xlXYScatter '�O���t�̎��
''.SetSourceData Range("A1:A15", "C1:C15") '�O���t�͈̔�
'
''�f�[�^�ǉ�
''.SeriesCollection.NewSeries
''.FullSeriesCollection(1).Name = "=Sheet1!A1"
''.FullSeriesCollection(1).XValues = "=Sheet1!B1:b15"
''.FullSeriesCollection(1).Values = "=Sheet1!C1:C15"
'
''.FullSeriesCollection(1).Select
''With Selection
''�}�[�J�[�̎�ނ̐ݒ�
''.MarkerStyle = 8
''�}�[�J�[�T�C�Y�̐ݒ�
''.MarkerSize = 7
''�}�[�J�[�̓h��Ԃ��ݒ� msoTrue�œh��Ԃ�
''.Format.Fill.Visible = msofulse
''�}�[�J�[�̐F�ݒ�
''.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
''�}�[�J�[����łȂ�
''.Format.Line.Visible = msoTrue
''�}�[�J�[�������H
''.MarkerForegroundColorIndex = xlColorIndexNone
''-------------------------------------------------
''.Border.Weight = 2.5 'xlThin�@�@�@�@�@�@'���̑���
''.Border.LineStyle = xlContinuous      '����
''.Border.LineStyle = 1
''.Border.ColorIndex = 1 '���̐F
''.MarkerBackgroundColorIndex = 7      '�}�[�J�[�̐F�@�h��Ԃ��H
''.MarkerForegroundColorIndex = 5      '�}�[�J�[�̐F�@�g���H
''�}�[�J�[�̘g���ύX
''.Format.Line.Visible = msoFalse
''.Format.Line.Weight = 0
''.Border.Weight = 4
''.MarkerStyle = xlDiamond              '�}�[�J�[�̌`
''.Smooth = True
'' .MarkerSize = 7 '�}�[�J�[�̃T�C�Y
''.Shadow = False                      '�e�iTrue or False�j
''End With
'
'
'
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
'�O���t�^�C�g���ǉ�
.HasTitle = True '�O���t�^�C�g���ǉ�
.ChartTitle.Text = "�L�����Ă�������" '�O���t�^�C�g���ύXvalue
.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11

'���x���ǉ�
'�c��
.Axes(xlValue).HasTitle = True
.Axes(xlValue).AxisTitle.Text = "�L�����Ă�������"
.Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
.Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 11
'����
.Axes(xlCategory).HasTitle = True
.Axes(xlCategory).AxisTitle.Text = "�L�����Ă�������"
.Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
.Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 11

'���̍ŏ��l�E�ő�l�ύX
'�c��
.Axes(xlValue).MinimumScale = 0
.Axes(xlValue).MaximumScale = 50
'����
.Axes(xlCategory).MinimumScale = 0
.Axes(xlCategory).MaximumScale = 20

''�s�Ɨ�����ւ���
'Select Case .PlotBy
'    Case xlRows
'        .PlotBy = xlColumns
'    Case xlColumns
'        .PlotBy = xlRows
'End Select

'�O���t�̘g�������F�ő�������
.PlotArea.Format.Line.Visible = msoTrue
.PlotArea.Format.Line.ForeColor.ObjectThemeColor = msoThemeColorText1
.PlotArea.Format.Line.ForeColor.TintAndShade = 0
.PlotArea.Format.Line.ForeColor.Brightness = 0
.PlotArea.Format.Line.Weight = 2

'���̘g��������
.Axes(xlValue).Format.Line.Visible = msoFalse
.Axes(xlCategory).Format.Line.Visible = msoFalse

End With


'If .HasLegend = False Then .HasLegend = True    ''�}���\������
'    .Legend.Position = xlLegendPositionTop          ''�}�����ɕ\������
'    .Legend.IncludeInLayout = False                 ''�}����O���t�ɏd�˂�
'With .Legend.Format.Fill
'    .Visible = msoTrue                          ''�}���h��Ԃ��܂�
'    .ForeColor.RGB = RGB(255, 0, 0)             ''�ԐF
'    .ForeColor.TintAndShade = 0.5               ''���Â̐ݒ�
'End With

With ActiveSheet.ChartObjects
    .Top = Range("E5").Top
    .Left = Range("E5").Left '�ʒu��ݒ�
    .Height = 200
    .Width = 300 '�傫����ݒ�
End With


End Sub

