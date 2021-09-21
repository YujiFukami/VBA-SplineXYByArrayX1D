Attribute VB_Name = "ModSplineXYByArrayX1D"
Option Explicit

'SplineXYByArrayX1D�E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineByArrayX1D  �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineKeisu       �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'F_MMult           �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_Minverse        �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'�����s�񂩃`�F�b�N�E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_MDeterm         �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_Mgyoirekae      �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_Mgyohakidasi    �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     
'F_Mjyokyo         �E�E�E���ꏊ�FFukamiAddins3.ModMatrix     

'------------------------------
'�V�[�g�֐��p�ߎ��A��Ԋ֐�
'------------------------------

'�s����g�����v�Z
'��֊֐�
'------------------------------


Function SplineXYByArrayX1D(ByVal ArrayXY2D, ByVal InputArrayX1D)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͔z��InputArrayX1D�ɑ΂����Ԓl�̔z��YList
    
    '�����͒l�̐�����
    'HariretuXY�F��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
    'ArrayXY2D��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
    'InputArrayX1D:��ԈʒuX���i�[���ꂽ�z��
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    Dim RangeNaraTrue As Boolean
    RangeNaraTrue = False
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
        RangeNaraTrue = True
    End If
    If IsObject(InputArrayX1D) Then
        InputArrayX1D = Application.Transpose(InputArrayX1D.Value)
    End If

    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D, ArrayY1D
    Dim I%, N%
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(1 To N)
    ReDim ArrayY1D(1 To N)
    
    For I = 1 To N
        ArrayX1D(I) = ArrayXY2D(I, 1)
        ArrayY1D(I) = ArrayXY2D(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim OutputArrayY1D
    OutputArrayY1D = SplineByArrayX1D(ArrayX1D, ArrayY1D, InputArrayX1D)
    
    '�o�́�����������������������������������������������������
    If RangeNaraTrue = True Then
        '���[�N�V�[�g�֐��̏ꍇ
        SplineXYByArrayX1D = Application.Transpose(OutputArrayY1D)
    Else
        'VBA��ł̏����̏ꍇ
        SplineXYByArrayX1D = OutputArrayY1D
    End If
    
End Function

Private Function SplineByArrayX1D(ByVal ArrayX1D, ByVal ArrayY1D, ByVal InputArrayX1D)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͔z��InputArrayX1D�ɑ΂����Ԓl�̔z��YList
    
    '�����͒l�̐�����
    'HairetuX�F��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
    'HairetuY�F��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
    'InputArrayX1D:��ԈʒuX���i�[���ꂽ�z��

    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    Dim RangeNaraTrue As Boolean
    RangeNaraTrue = False
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
        RangeNaraTrue = True
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    If IsObject(InputArrayX1D) Then
        InputArrayX1D = Application.Transpose(InputArrayX1D.Value)
    End If
    
    Dim StartNum%
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1D�̊J�n�v�f�ԍ�������Ă����i�o�͒l�����킹�邽�߁j
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%, JigenCheck3%
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck3 = UBound(InputArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '�v�Z����������������������������������������������������������
    Dim N%, K%, A, B, C, D
    
    '�X�v���C���v�Z�p�̊e�W�����v�Z����B�Q�Ɠn����A,B,C,D�Ɋi�[
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
    
    Dim SotoNaraTrue As Boolean
    N = UBound(ArrayX1D, 1) '��ԑΏۂ̗v�f��
    
    Dim OutputArrayY1D#() '�o�͂���Y�̊i�[
    Dim NX%
    NX = UBound(InputArrayX1D, 1) '��Ԉʒu�̌�
    ReDim OutputArrayY1D(1 To NX)
    Dim TmpX#, TmpY#
    
    For J = 1 To NX
        TmpX = InputArrayX1D(J)
        SotoNaraTrue = False
        For I = 1 To N - 1
            If ArrayX1D(I) < ArrayX1D(I + 1) Then 'X���P�������̏ꍇ
                If I = 1 And ArrayX1D(1) > TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) <= TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I) <= TmpX And ArrayX1D(I + 1) > TmpX Then '�͈͓�
                    K = I: Exit For
                
                End If
            Else 'X���P�������̏ꍇ
            
                If I = 1 And ArrayX1D(1) < TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) >= TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I + 1) < TmpX And ArrayX1D(I) >= TmpX Then '�͈͓�
                    K = I: Exit For
                
                End If
            
            End If
        Next I
        
        If SotoNaraTrue = False Then
            TmpY = A(K) + B(K) * (TmpX - ArrayX1D(K)) + C(K) * (TmpX - ArrayX1D(K)) ^ 2 + D(K) * (TmpX - ArrayX1D(K)) ^ 3
        End If
        
        OutputArrayY1D(J) = TmpY
        
    Next J
    
    '�o�́�����������������������������������������������������
    Dim Output
    
    '�o�͂���z�����͂����z��InputArrayX1D�̌`��ɍ��킹��
    If JigenCheck3 = 1 Then '���͂�InputArrayX1D���񎟌��z��
        ReDim Output(StartNum To StartNum + NX - 1, 1 To 1)
        For I = 1 To NX
            Output(StartNum + I - 1, 1) = OutputArrayY1D(I)
        Next I
    Else
        If StartNum = 1 Then
            Output = OutputArrayY1D
        Else
            ReDim Output(StartNum To StartNum + NX - 1)
            For I = 1 To NX
                Output(StartNum + I - 1) = OutputArrayY1D(I)
            Next I
        End If
    End If
    
    If RangeNaraTrue Then
        '���[�N�V�[�g�֐��̏ꍇ
        SplineByArrayX1D = Application.Transpose(Output)
    Else
        'VBA��ł̏����̏ꍇ
        SplineByArrayX1D = Output
    End If
    
End Function

Private Function SplineKeisu(ByVal ArrayX1D, ByVal ArrayY1D)

    '�Q�l�Fhttp://www5d.biglobe.ne.jp/stssk/maze/spline.html
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim A, B, C, D
    N = UBound(ArrayX1D, 1)
    ReDim A(1 To N)
    ReDim B(1 To N)
    ReDim D(1 To N)
    
    Dim h#()
    Dim ArrayL2D#() '���ӂ̔z�� �v�f��(1 to N,1 to N)
    Dim ArrayR1D#() '�E�ӂ̔z�� �v�f��(1 to N,1 to 1)
    Dim ArrayLm2D#() '���ӂ̔z��̋t�s�� �v�f��(1 to N,1 to N)
    
    ReDim h(1 To N - 1)
    ReDim ArrayL2D(1 To N, 1 To N)
    ReDim ArrayR1D(1 To N, 1 To 1)
    
    'hi = xi+1 - x
    For I = 1 To N - 1
        h(I) = ArrayX1D(I + 1) - ArrayX1D(I)
    Next I
    
    'di = yi
    For I = 1 To N
        A(I) = ArrayY1D(I)
    Next I
    
    '�E�ӂ̔z��̌v�Z
    For I = 1 To N
        If I = 1 Or I = N Then
            ArrayR1D(I, 1) = 0
        Else
            ArrayR1D(I, 1) = 3 * (ArrayY1D(I + 1) - ArrayY1D(I)) / h(I) - 3 * (ArrayY1D(I) - ArrayY1D(I - 1)) / h(I - 1)
        End If
    Next I
    
    '���ӂ̔z��̌v�Z
    For I = 1 To N
        If I = 1 Then
            ArrayL2D(I, 1) = 1
        ElseIf I = N Then
            ArrayL2D(N, N) = 1
        Else
            ArrayL2D(I - 1, I) = h(I - 1)
            ArrayL2D(I, I) = 2 * (h(I) + h(I - 1))
            ArrayL2D(I + 1, I) = h(I)
        End If
    Next I
    
    '���ӂ̔z��̋t�s��
    ArrayLm2D = F_Minverse(ArrayL2D)
    
    'C�̔z������߂�
    C = F_MMult(ArrayLm2D, ArrayR1D)
    C = Application.Transpose(C)
    
    'B�̔z������߂�
    For I = 1 To N - 1
        B(I) = (A(I + 1) - A(I)) / h(I) - h(I) * (C(I + 1) + 2 * C(I)) / 3
    Next I
    
    'D�̔z������߂�
    For I = 1 To N - 1
        D(I) = (C(I + 1) - C(I)) / (3 * h(I))
    Next I
    
    '�o��
    Dim Output(1 To 4)
    Output(1) = A
    Output(2) = B
    Output(3) = C
    Output(4) = D
    
    SplineKeisu = Output

End Function

Private Function F_MMult(ByVal Matrix1, ByVal Matrix2)
    'F_MMult(Matrix1, Matrix2)
    'F_MMult(�z��@,�z��A)
    '�s��̐ς��v�Z
    '20180213����
    '20210603����
    
    '���͒l�̃`�F�b�N�ƏC��������������������������������������������������������
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(Matrix1, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(Matrix2, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����1�Ȃ玟��2�ɂ���B��)�z��(1 to N)���z��(1 to N,1 to 1)
    If IsEmpty(JigenCheck1) Then
        Matrix1 = Application.Transpose(Matrix1)
    End If
    If IsEmpty(JigenCheck2) Then
        Matrix2 = Application.Transpose(Matrix2)
    End If
    
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If UBound(Matrix1, 1) = 0 Or UBound(Matrix1, 2) = 0 Then
        Matrix1 = Application.Transpose(Application.Transpose(Matrix1))
    End If
    If UBound(Matrix2, 1) = 0 Or UBound(Matrix2, 2) = 0 Then
        Matrix2 = Application.Transpose(Application.Transpose(Matrix2))
    End If
    
    '���͒l�̃`�F�b�N
    If UBound(Matrix1, 2) <> UBound(Matrix2, 1) Then
        MsgBox ("�z��1�̗񐔂Ɣz��2�̍s������v���܂���B" & vbLf & _
               "(�o��) = (�z��1)(�z��2)")
        Stop
        End
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim M2%
    Dim Output#() '�o�͂���z��
    N = UBound(Matrix1, 1) '�z��1�̍s��
    M = UBound(Matrix1, 2) '�z��1�̗�
    M2 = UBound(Matrix2, 2) '�z��2�̗�
    
    ReDim Output(1 To N, 1 To M2)
    
    For I = 1 To N '�e�s
        For J = 1 To M2 '�e��
            For K = 1 To M '(�z��1��I�s)��(�z��2��J��)���|�����킹��
                Output(I, J) = Output(I, J) + Matrix1(I, K) * Matrix2(K, J)
            Next K
        Next J
    Next I
    
    '�o�́�����������������������������������������������������
    F_MMult = Output
    
End Function

Private Function F_Minverse(ByVal Matrix)
    '20210603����
    'F_Minverse(input_M)
    'F_Minverse(�z��)
    '�]���q�s���p���ċt�s����v�Z
    
    '���͒l�`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '���͒l�̃`�F�b�N
    Call �����s�񂩃`�F�b�N(Matrix)
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, M2%, N% '�����グ�p(Integer�^)
    N = UBound(Matrix, 1)
    Dim Output#()
    ReDim Output(1 To N, 1 To N)
    
    Dim detM# '�s�񎮂̒l���i�[
    detM = F_MDeterm(Matrix) '�s�񎮂����߂�
    
    Dim Mjyokyo '�w��̗�E�s�����������z����i�[
    
    For I = 1 To N '�e��
        For J = 1 To N '�e�s
            
            'I��,J�s����������
            Mjyokyo = F_Mjyokyo(Matrix, J, I)
            
            'I��,J�s�̗]���q�����߂ďo�͂���t�s��Ɋi�[
            Output(I, J) = F_MDeterm(Mjyokyo) * (-1) ^ (I + J) / detM
    
        Next J
    Next I
    
    '�o�́�����������������������������������������������������
    F_Minverse = Output
    
End Function

Private Sub �����s�񂩃`�F�b�N(Matrix)
    '20210603�ǉ�
    
    If UBound(Matrix, 1) <> UBound(Matrix, 2) Then
        MsgBox ("�����s�����͂��Ă�������" & vbLf & _
                "���͂��ꂽ�z��̗v�f����" & "�u" & _
                UBound(Matrix, 1) & "�~" & UBound(Matrix, 2) & "�v" & "�ł�")
        Stop
        End
    End If

End Sub

Private Function F_MDeterm(Matrix)
    '20210603����
    'F_MDeterm(Matrix)
    'F_MDeterm(�z��)
    '�s�񎮂��v�Z
    
    '���͒l�`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '���͒l�̃`�F�b�N
    Call �����s�񂩃`�F�b�N(Matrix)
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(Matrix, 1)
    
    Dim Matrix2 '�|���o�����s���s��
    Matrix2 = Matrix
    
    For I = 1 To N '�e��
        For J = I To N '�|���o�����̍s�̒T��
            If Matrix2(J, I) <> 0 Then
                K = J '�|���o�����̍s
                Exit For
            End If
            
            If J = N And Matrix2(J, I) = 0 Then '�|���o�����̒l���S��0�Ȃ�s�񎮂̒l��0
                F_MDeterm = 0
                Exit Function
            End If
            
        Next J
        
        If K <> I Then '(I��,I�s)�ȊO�ő|���o���ƂȂ�ꍇ�͍s�����ւ�
            Matrix2 = F_Mgyoirekae(Matrix2, I, K)
        End If
        
        '�|���o��
        Matrix2 = F_Mgyohakidasi(Matrix2, I, I)
              
    Next I
    
    
    '�s�񎮂̌v�Z
    Dim Output#
    Output = 1
    
    For I = 1 To N '�e(I��,I�s)���|�����킹�Ă���
        Output = Output * Matrix2(I, I)
    Next I
    
    '�o�́�����������������������������������������������������
    F_MDeterm = Output
    
End Function

Private Function F_Mgyoirekae(Matrix, Row1%, Row2%)
    '20210603����
    'F_Mgyoirekae(Matrix, Row1, Row2)
    'F_Mgyoirekae(�z��,�w��s�ԍ��@,�w��s�ԍ��A)
    '�s��Matrix�̇@�s�ƇA�s�����ւ���
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    
    Output = Matrix
    M = UBound(Matrix, 2) '�񐔎擾
    
    For I = 1 To M
        Output(Row2, I) = Matrix(Row1, I)
        Output(Row1, I) = Matrix(Row2, I)
    Next I
    
    F_Mgyoirekae = Output
End Function

Private Function F_Mgyohakidasi(Matrix, Row%, Col%)
    '20210603����
    'F_Mgyohakidasi(Matrix, Row, Col)
    'F_Mgyohakidasi(�z��,�w��s,�w���)
    '�s��Matrix��Row�s�Col��̒l�Ŋe�s��|���o��
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    
    Output = Matrix
    N = UBound(Output, 1) '�s���擾
    
    Dim Hakidasi '�|���o�����̍s
    Dim X# '�|���o�����̒l
    Dim Y#
    ReDim Hakidasi(1 To N)
    X = Matrix(Row, Col)
    
    For I = 1 To N '�|���o������1�s���쐬
        Hakidasi(I) = Matrix(Row, I)
    Next I
    
    
    For I = 1 To N '�e�s
        If I = Row Then
            '�|���o�����̍s�̏ꍇ�͂��̂܂�
            For J = 1 To N
                Output(I, J) = Matrix(I, J)
            Next J
        
        Else
            '�|���o�����̍s�ȊO�̏ꍇ�͑|���o��
            Y = Matrix(I, Col) '�|���o����̗�̒l
            For J = 1 To N
                Output(I, J) = Matrix(I, J) - Hakidasi(J) * Y / X
            Next J
        End If
    
    Next I
    
    F_Mgyohakidasi = Output
    
End Function

Private Function F_Mjyokyo(Matrix, Row%, Col%)
    '20210603����
    'F_Mjyokyo(Matrix, Row, Col)
    'F_Mjyokyo(�z��,�w��s,�w���)
    '�s��Matrix��Row�s�ACol������������s���Ԃ�
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output '�w�肵���s�E���������̔z��
    
    N = UBound(Matrix, 1) '�s���擾
    M = UBound(Matrix, 2) '�񐔎擾
    ReDim Output(1 To N - 1, 1 To M - 1)
    
    Dim I2%, J2%
    
    I2 = 0 '�s���������グ������
    For I = 1 To N
        If I = Row Then
            '�Ȃɂ����Ȃ�
        Else
            I2 = I2 + 1 '�s���������グ
            
            J2 = 0 '����������グ������
            For J = 1 To M
                If J = Col Then
                    '�Ȃɂ����Ȃ�
                Else
                    J2 = J2 + 1 '����������グ
                    Output(I2, J2) = Matrix(I, J)
                End If
            Next J
            
        End If
    Next I
    
    F_Mjyokyo = Output

End Function

