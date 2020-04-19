Attribute VB_Name = "Module6"
Public Sub �C�ӂ̃f�[�^�_���ŏd�ˍ��킹�e�񕡐��O���t�쐬200329()

Dim i As Integer '�J�E���^
Dim j As Integer '�J�E���^

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
    
Dim keiretu_num As Integer '�d�ˍ��킹��n��
'���̓L�����Z�������ƃu�[���l���Ԃ��Ă���̂ŁA
    '�o���A���g�^�ϐ��Ŏ󂯂܂��B
    keiretu_num = Application.InputBox( _
                    "�d�ˍ��킹��n�񐔂���͂��Ă��������B", Type:=1)
    If TypeName(keiretu_num) = "Boolean" Then
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
Dim GraphSize_X As Integer '�O���tX�����T�C�Y
Dim GraphSize_Y As Integer '�O���tY�����T�C�Y

'�O���tX�����T�C�Y
GraphSize_X = 270.321502685547
'�O���tY�����T�C�Y
GraphSize_Y = 184.821411132813

'�O���t�o���ʒu�ƃT�C�Y���Z�b�g
With ActiveSheet.ChartObjects.Add(graph_pos.Left, graph_pos.Top, GraphSize_X, GraphSize_Y).Chart '��
'�O���t�̎��
.ChartType = xlXYScatter

'�n��f�[�^�̃v���b�g
'�f�[�^�ǉ�
j = start + GraphLength - 1
For i = 1 To keiretu_num
    .SeriesCollection.NewSeries
With .SeriesCollection(i)
    .Name = Sheets(CStr(1)).Cells(j - GraphLength + 1, 1)
    .XValues = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(j - GraphLength + 1, x), Sheets(CStr(1)).Cells(j, x))
    .Values = Sheets(CStr(1)).Range(Sheets(CStr(1)).Cells(j - GraphLength + 1, y), Sheets(CStr(1)).Cells(j, y))
End With
    j = j + GraphLength
    
'----�O���t�����ύX----
'�O���t�^�C�g���ǉ�
.HasTitle = True '�O���t�^�C�g���ǉ�
.ChartTitle.Text = "�L�����Ă�������"
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
''���̍ŏ��l�E�ő�l�ύX
''�c��
'.Axes(xlValue).MinimumScale = 0
'.Axes(xlValue).MaximumScale = 200
''����
'.Axes(xlCategory).MinimumScale = 0
'.Axes(xlCategory).MaximumScale = 20
'�O���t�̘g�������F�ő�������i��@�j
.PlotArea.Format.Line.Visible = msoTrue
.PlotArea.Format.Line.ForeColor.ObjectThemeColor = msoThemeColorText1
.PlotArea.Format.Line.ForeColor.TintAndShade = 0
.PlotArea.Format.Line.ForeColor.Brightness = 0
.PlotArea.Format.Line.Weight = 2
'���̘g��������(��@)
.Axes(xlValue).Format.Line.Visible = msoFalse
.Axes(xlCategory).Format.Line.Visible = msoFalse
'�}��ݒ�
'�}���\������
.HasLegend = True
'�}�����ɕ\������
.Legend.Position = xlLegendPositionTop
'�}����O���t�ɏd�˂�
.Legend.IncludeInLayout = True
With .Legend.Format.Fill
'�}���h��Ԃ��܂�
    .Visible = msoTrue
'�}��̓h��Ԃ��̐F
    .ForeColor.RGB = vbWhite
'���Â̐ݒ�
    .ForeColor.TintAndShade = 0.5
End With
'--------------------
.FullSeriesCollection(i).Select '�Z
With Selection
'�}�[�J�[�̎�ނ̐ݒ�
.MarkerStyle = 8
'�}�[�J�[�̃T�C�Y
.MarkerSize = 7
'�}�[�J�[�̗L��
.Format.Fill.Visible = msoTrue
'�}�[�J�[�̐F�ݒ�
'.Format.Fill.ForeColor.RGB = vbRed
'�}�[�J�[�̘g���ύX
.Format.Line.Visible = msoTrue
.Format.Line.Weight = 0
'�}�[�J�[�h��Ԃ��̐F
.MarkerBackgroundColorIndex = -4105
'�}�[�J�[�g���̐F
.MarkerForegroundColorIndex = -4105
'�}�[�J�[�̘g��������_��̐F�̒�`�̌�ɒu���Ȃ��Ƙg�������Ȃ�
.MarkerForegroundColorIndex = xlColorIndexNone
'�}�[�J�[����łȂ�
.Format.Line.Visible = msoFalse
'���̑���
'.Border.Weight = 4 'xlThin
'����
'.Border.LineStyle = xlContinuous
'���̐F
'.Border.ColorIndex = -4105
'�������łȂ߂炩�ɂ��邩�����ŕ�Ԃ��邩
.Smooth = True
'�e (True Or False)false�̕����ǂ��Btrue���Ɛ��������H
.Shadow = False
End With '�Z
'----------------------
''�s�Ɨ�����ւ���@�����̃v���O�������Ƃ��܂������Ȃ��E�E�E
'Select Case .PlotBy
'    Case xlRows
'        .PlotBy = xlColumns
'    Case xlColumns
'        .PlotBy = xlRows
'End Select

Next i

End With '��

End Sub
