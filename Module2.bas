Attribute VB_Name = "Module2"
Public Sub �O���t����1�Ԗڂ̂�()

With ActiveSheet.ChartObjects(1).Chart

Set Object = Selection

.FullSeriesCollection(1).Select
With Selection
'�}�[�J�[�̎�ނ̐ݒ�
.MarkerStyle = 8
'�}�[�J�[�T�C�Y�̐ݒ�
.MarkerSize = 7
'�}�[�J�[�̗L��
.Format.Fill.Visible = msoFalse
'�}�[�J�[�̐F�ݒ�
.Format.Fill.ForeColor.RGB = vbRed
'�}�[�J�[����łȂ�_true�Ő�������
.Format.Line.Visible = msoFalse
'-------------------------------------------------
'����
.Border.LineStyle = xlContinuous
'���̐F
.Border.ColorIndex = 1
''�}�[�J�[�h��Ԃ��̐F
'.MarkerBackgroundColorIndex = 6
''�}�[�J�[�g���̐F
'.MarkerForegroundColorIndex = 8
''�}�[�J�[�̘g���������@��̐F�̒�`�̌�ɒu���Ȃ��Ƙg�������Ȃ�
'.MarkerForegroundColorIndex = xlColorIndexNone
'�}�[�J�[�̘g���ύX
.Format.Line.Visible = msoTrue
.Format.Line.Weight = 1
'���̑���
.Border.Weight = 4
'�������łȂ߂炩�ɂ��邩�����ŕ�Ԃ��邩
.Smooth = True
'�e�iTrue or False�jfalse�̕����ǂ�
.Shadow = False
End With


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
.Axes(xlValue).MaximumScale = 80
'����
.Axes(xlCategory).MinimumScale = 0
.Axes(xlCategory).MaximumScale = 60

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


If .HasLegend = False Then .HasLegend = True    ''�}���\������
    .Legend.Position = xlLegendPositionTop          ''�}�����ɕ\������
    .Legend.IncludeInLayout = False                 ''�}����O���t�ɏd�˂�
With .Legend.Format.Fill
    .Visible = msoTrue                          ''�}���h��Ԃ��܂�
    .ForeColor.RGB = vbWhite             ''�ԐF
    .ForeColor.TintAndShade = 1               ''���Â̐ݒ�
End With

End With

With ActiveSheet.ChartObjects
    .Top = Range("E5").Top
    .Left = Range("E5").Left '�ʒu��ݒ�
    .Height = 200
    .Width = 300 '�傫����ݒ�
End With


End Sub



Public Sub ����()

Dim x As Object
Dim xi As Object

    Set x = Selection
    If TypeName(x) <> "DrawingObjects" Then
        If Not ActiveChart Is Nothing Then
                 
                With ActiveChart
                
                '.FullSeriesCollection(1).Select
                'With Selection
                ''�}�[�J�[�̎�ނ̐ݒ�
                '.MarkerStyle = 8
                ''�}�[�J�[�T�C�Y�̐ݒ�
                '.MarkerSize = 7
                ''�}�[�J�[�̓h��Ԃ��ݒ� msoTrue�œh��Ԃ�
                '.Format.Fill.Visible = msofulse
                ''�}�[�J�[�̐F�ݒ�
                '.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                ''�}�[�J�[����łȂ�
                '.Format.Line.Visible = msoTrue
                ''�}�[�J�[�������H
                '.MarkerForegroundColorIndex = xlColorIndexNone
                ''-------------------------------------------------
                '.Border.Weight = 2.5 'xlThin�@�@�@�@�@�@'���̑���
                '.Border.LineStyle = xlContinuous      '����
                '.Border.LineStyle = 1
                '.Border.ColorIndex = 1 '���̐F
                '.MarkerBackgroundColorIndex = 7      '�}�[�J�[�̐F�@�h��Ԃ��H
                '.MarkerForegroundColorIndex = 5      '�}�[�J�[�̐F�@�g���H
                ''�}�[�J�[�̘g���ύX
                '.Format.Line.Visible = msoFalse
                '.Format.Line.Weight = 0
                '.Border.Weight = 4
                '.MarkerStyle = xlDiamond              '�}�[�J�[�̌`
                '.Smooth = True
                ' .MarkerSize = 7 '�}�[�J�[�̃T�C�Y
                '.Shadow = False                      '�e�iTrue or False�j
                'End With
                
                
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
                .Axes(xlValue).MaximumScale = 80
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
                
                
                'If .HasLegend = False Then .HasLegend = True    ''�}���\������
                '    .Legend.Position = xlLegendPositionTop          ''�}�����ɕ\������
                '    .Legend.IncludeInLayout = False                 ''�}����O���t�ɏd�˂�
                'With .Legend.Format.Fill
                '    .Visible = msoTrue                          ''�}���h��Ԃ��܂�
                '    .ForeColor.RGB = RGB(255, 0, 0)             ''�ԐF
                '    .ForeColor.TintAndShade = 0.5               ''���Â̐ݒ�
                'End With
                
                End With
                
                With ActiveSheet.ChartObjects
                    .Top = Range("E5").Top
                    .Left = Range("E5").Left '�ʒu��ݒ�
                    .Height = 200
                    .Width = 300 '�傫����ݒ�
                End With
            
        End If
    Else
        For Each xi In x
            If TypeName(xi) = "ChartObject" Then
                
                With xi.Chart
                
                '.FullSeriesCollection(1).Select
                'With Selection
                ''�}�[�J�[�̎�ނ̐ݒ�
                '.MarkerStyle = 8
                ''�}�[�J�[�T�C�Y�̐ݒ�
                '.MarkerSize = 7
                ''�}�[�J�[�̓h��Ԃ��ݒ� msoTrue�œh��Ԃ�
                '.Format.Fill.Visible = msofulse
                ''�}�[�J�[�̐F�ݒ�
                '.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                ''�}�[�J�[����łȂ�
                '.Format.Line.Visible = msoTrue
                ''�}�[�J�[�������H
                '.MarkerForegroundColorIndex = xlColorIndexNone
                ''-------------------------------------------------
                '.Border.Weight = 2.5 'xlThin�@�@�@�@�@�@'���̑���
                '.Border.LineStyle = xlContinuous      '����
                '.Border.LineStyle = 1
                '.Border.ColorIndex = 1 '���̐F
                '.MarkerBackgroundColorIndex = 7      '�}�[�J�[�̐F�@�h��Ԃ��H
                '.MarkerForegroundColorIndex = 5      '�}�[�J�[�̐F�@�g���H
                ''�}�[�J�[�̘g���ύX
                '.Format.Line.Visible = msoFalse
                '.Format.Line.Weight = 0
                '.Border.Weight = 4
                '.MarkerStyle = xlDiamond              '�}�[�J�[�̌`
                '.Smooth = True
                ' .MarkerSize = 7 '�}�[�J�[�̃T�C�Y
                '.Shadow = False                      '�e�iTrue or False�j
                'End With
                
                
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
                .Axes(xlValue).MaximumScale = 40
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
                
                
                'If .HasLegend = False Then .HasLegend = True    ''�}���\������
                '    .Legend.Position = xlLegendPositionTop          ''�}�����ɕ\������
                '    .Legend.IncludeInLayout = False                 ''�}����O���t�ɏd�˂�
                'With .Legend.Format.Fill
                '    .Visible = msoTrue                          ''�}���h��Ԃ��܂�
                '    .ForeColor.RGB = RGB(255, 0, 0)             ''�ԐF
                '    .ForeColor.TintAndShade = 0.5               ''���Â̐ݒ�
                'End With
                
                End With
                
                With ActiveSheet.ChartObjects
                    .Top = Range("E5").Top
                    .Left = Range("E5").Left '�ʒu��ݒ�
                    .Height = 200
                    .Width = 300 '�傫����ݒ�
                End With
                
            End If
        Next
    End If
  
    Set xi = Nothing

End Sub
