Attribute VB_Name = "FileEncoding"
Option Explicit
'****************************************************************************
' �@�\��    : Module1.bas
' �@�\����  : �����R�[�h����
' ���l      :
' ���쌠    : Copyright(C) 2008 - 2009 �̂� All rights reserved
' ---------------------------------------------------------------------------
' �g�p����  : ���̃T�C�g�̓��e���g�p(���p/����/�]��/���S��)�������ʕ���s����
'           : �����Ɍ��J/�z�z����ꍇ�́A���̃T�C�g���Q�l�ɂ����|���L�q���Ă�
'           : �������B(��)WEB�y�[�W��ReadMe�Ƀ����N��\���Ă�������
' ---------------------------------------------------------------------------
'****************************************************************************
Private Const JUDGEFIX = 9999       '�����R�[�h���聓
Private Const JUDGESIZEMAX = 1000   '�����R�[�h����o�C�g��
Private Const SingleByteWeight = 1  '�P�o�C�g�@�����R�[�h�̈�v�d��
Private Const Multi_ByteWeight = 2  '�����o�C�g�����R�[�h�̈�v�d��
Private Enum JISMODE                'JIS�R�[�h�̃��[�h
    ctrl = 0                        '����R�[�h
    asci = 1                        'ASCII
    roma = 2                        'JIS���[�}��
    kana = 3                        'JIS�J�i�i���p�J�i�j
    kanO = 4                        '��JIS���� (1978)
    kanN = 5                        '�VJIS���� (1983/1990)
    kanH = 6                        'JIS�⏕����
End Enum

'----�����R�[�h����
' �֐���    : JudgeCode
' �Ԃ�l    : ���茋�ʕ����R�[�h��
' ������    : bytCode : ���蕶���f�[�^
' �@�\����  : �����R�[�h�𔻒肷��
' ���l      :
Public Function JudgeCode(ByRef bytCode() As Byte) As String
    JudgeCode = "SJIS"
    Dim lngSJIS As Long
    Dim lngJIS As Long
    Dim lngEUC As Long
    Dim lngUNI As Long
    Dim lngUTF7 As Long
    Dim lngUTF8 As Long
    
    lngJIS = JudgeJIS(bytCode, True): Debug.Print "JIS :" & lngJIS
    If lngJIS >= JUDGEFIX Then JudgeCode = "JIS": Exit Function
    
    lngUNI = JudgeUNI(bytCode, True): Debug.Print "UNI :" & lngUNI
    If lngUNI >= JUDGEFIX Then JudgeCode = "UNICODE": Exit Function
    
    lngUTF8 = JudgeUTF8(bytCode, True): Debug.Print "UTF8:" & lngUTF8
    If lngUTF8 >= JUDGEFIX Then JudgeCode = "UTF8": Exit Function

    lngUTF7 = JudgeUTF7(bytCode, True): Debug.Print "UTF7:" & lngUTF7
    If lngUTF7 >= JUDGEFIX Then JudgeCode = "UTF7": Exit Function
    
    lngSJIS = JudgeSJIS(bytCode, True): Debug.Print "SJIS:" & lngSJIS
    If lngSJIS >= JUDGEFIX Then JudgeCode = "SJIS": Exit Function
    
    lngEUC = JudgeEUC(bytCode, True): Debug.Print "EUC :" & lngEUC
    If lngEUC >= JUDGEFIX Then JudgeCode = "EUC": Exit Function
    Debug.Print "--------"

    If lngSJIS >= lngSJIS And lngSJIS >= lngUNI And lngSJIS >= lngJIS And _
       lngSJIS >= lngUTF7 And lngSJIS >= lngUTF8 And lngSJIS >= lngEUC Then
        JudgeCode = "SJIS"
        Exit Function
    End If
    
    If lngUNI >= lngSJIS And lngUNI >= lngUNI And lngUNI >= lngJIS And _
       lngUNI >= lngUTF7 And lngUNI >= lngUTF8 And lngUNI >= lngEUC Then
        JudgeCode = "UNICODE"
        Exit Function
    End If
    
    If lngJIS >= lngSJIS And lngJIS >= lngUNI And lngJIS >= lngJIS And _
       lngJIS >= lngUTF7 And lngJIS >= lngUTF8 And lngJIS >= lngEUC Then
        JudgeCode = "JIS"
        Exit Function
    End If
    
    If lngUTF7 >= lngSJIS And lngUTF7 >= lngUNI And lngUTF7 >= lngJIS And _
       lngUTF7 >= lngUTF7 And lngUTF7 >= lngUTF8 And lngUTF7 >= lngEUC Then
        JudgeCode = "UTF7"
        Exit Function
    End If
    
    If lngUTF8 >= lngSJIS And lngUTF8 >= lngUNI And lngUTF8 >= lngJIS And _
       lngUTF8 >= lngUTF7 And lngUTF8 >= lngUTF8 And lngUTF8 >= lngEUC Then
        JudgeCode = "UTF8"
        Exit Function
    End If
    
    If lngEUC >= lngSJIS And lngEUC >= lngUNI And lngEUC >= lngJIS And _
       lngEUC >= lngUTF7 And lngEUC >= lngUTF8 And lngEUC >= lngEUC Then
        JudgeCode = "EUC"
        Exit Function
    End If
    
End Function

'----SJIS�֌W
' �֐���    : JudgeSJIS
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : SJIS�̕����R�[�h����(�\��)�m�����v�Z����
' ���l      :
Private Function JudgeSJIS(ByRef bytCode() As Byte, _
                           Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        '81-9F,E0-EF(1�o�C�g��)
        If (bytCode(i) >= &H81 And bytCode(i) <= &H9F) Or _
           (bytCode(i) >= &HE0 And bytCode(i) <= &HEF) Then
           If i <= UBound(bytCode) - 1 Then
                '40-7E,80-FC(2�o�C�g��)
                If (bytCode(i + 1) >= &H40 And bytCode(i + 1) <= &H7E) Or _
                   (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HFC) Then
                    lngFit = lngFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If
        
        'A1-DF(1�o�C�g��)
        ElseIf (bytCode(i) >= &HA1 And bytCode(i) <= &HDF) Then
            lngFit = lngFit + (1 * SingleByteWeight)
        
        '20-7E(1�o�C�g��)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            lngFit = lngFit + (1 * SingleByteWeight)
        
        '00-1F, 7F(1�o�C�g��)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            lngFit = lngFit + (1 * SingleByteWeight)
        End If
    Next i
    JudgeSJIS = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----JIS�֌W
' �֐���    : JudgeJIS
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : JIS�̕����R�[�h����(�\��)�m�����v�Z����
' ���l      :
Private Function JudgeJIS(ByRef bytCode() As Byte, _
                          Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngMode As JISMODE
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        '1B(1�o�C�g��)
        If bytCode(i) = &H1B Then
           If i <= UBound(bytCode) - 2 Then
                '28 42(2�E3�o�C�g��)
                If bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H42 Then
                    lngMode = asci
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '28 4A(2�E3�o�C�g��)
                If bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H4A Then
                    lngMode = roma
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '28 49(2�E3�o�C�g��)
                If bytCode(i + 1) = &H28 And bytCode(i + 1) <= &H49 Then
                    lngMode = kana
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '24 40(2�E3�o�C�g��)
                If bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H40 Then
                    lngMode = kanO
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '24 42(2�E3�o�C�g��)
                If bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H42 Then
                    lngMode = kanN
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
                '24 44(2�E3�o�C�g��)
                If bytCode(i + 1) = &H24 And bytCode(i + 1) <= &H44 Then
                    lngMode = kanH
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                    If fixFlag Then
                        JudgeJIS = JUDGEFIX
                        Exit Function
                    End If
                End If
            End If
        Else
            Select Case lngMode
            Case ctrl, asci, roma
                '00-1F,7F
                If (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                    bytCode(i) = &H7F Then
                    lngFit = lngFit + (1 * SingleByteWeight)
                End If
                '20-7E
                If (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
                    lngFit = lngFit + (1 * SingleByteWeight)
                End If
            Case kana
                '21-5F
                If (bytCode(i) >= &H21 And bytCode(i) <= &H5F) Then
                    lngFit = lngFit + (1 * SingleByteWeight)
                End If
            Case kanO, kanN, kanH
               If i <= UBound(bytCode) - 1 Then
                    '21-7E
                    If (bytCode(i) >= &H21 And bytCode(i) <= &H7E) And _
                       (bytCode(i - 1) >= &H21 And bytCode(i - 1) <= &H7E) Then
                        lngFit = lngFit + (2 * Multi_ByteWeight)
                        i = i + 1
                    End If
                End If
            End Select
        End If
    Next i
    JudgeJIS = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----EUC�֌W
' �֐���    : JudgeEUC
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : EUC�̕����R�[�h����(�\��)�m�����v�Z����
' ���l      :
Private Function JudgeEUC(ByRef bytCode() As Byte, _
                          Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        '8E(1�o�C�g��) + A1-DF(2�o�C�g��)
        If bytCode(i) = &H8E Then
            If i <= UBound(bytCode) - 1 Then
                If bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HDF Then
                    lngFit = lngFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If
        
        '8F(1�o�C�g��) + A1-0xFE(2�E3�o�C�g��)
        ElseIf bytCode(i) = &H8F Then
            If i <= UBound(bytCode) - 2 Then
                If (bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HFE) And _
                   (bytCode(i + 2) >= &HA1 And bytCode(i + 2) <= &HFE) Then
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                End If
            End If
        
        'A1-FE(1�o�C�g��) + A1-FE(2�o�C�g��)
        ElseIf bytCode(i) >= &HA1 And bytCode(i) <= &HFE Then
            If i <= UBound(bytCode) - 1 Then
                If bytCode(i + 1) >= &HA1 And bytCode(i + 1) <= &HFE Then
                    lngFit = lngFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If
            
        '20-7E(1�o�C�g��)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            lngFit = lngFit + (1 * SingleByteWeight)

        '00-1F, 7F(1�o�C�g��)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            lngFit = lngFit + (1 * SingleByteWeight)
        End If
    Next i
    JudgeEUC = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----UNICODE�֌W
' �֐���    : JudgeUNI
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : UTF16�̕����R�[�h����(�\��)�m�����v�Z����
' ���l      :
Private Function JudgeUNI(ByRef bytCode() As Byte, _
                          Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        If fixFlag Then
            'BOM
            If bytCode(i) = &HFF Then
                If i <= UBound(bytCode) - 1 Then
                    If bytCode(i + 1) = &HFE Then
                        JudgeUNI = JUDGEFIX
                        Exit Function
                    End If
                End If
            End If
            '���p�̏�
            'If bytCode(i) = &H0 Then
            '    JudgeUNI = JUDGEFIX
            '    Exit Function
            'End If
        End If
        
        If i <= UBound(bytCode) - 1 Then
            '00(2�o�C�g��)
            If (bytCode(i + 1) = &H0) Then
                '00-FF(1�o�C�g��)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            '01-33(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &H1 And bytCode(i + 1) <= &H33) Then
                '00-FF(1�o�C�g��)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            '34-4D(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &H34 And bytCode(i + 1) <= &H4D) Then
                '00-FF(1�o�C�g��)----��----
                lngFit = 0
                Exit For
            
            '4E-9F(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &H4E And bytCode(i + 1) <= &H9F) Then
                '00-FF(1�o�C�g��)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            'A0-AB(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HA0 And bytCode(i + 1) <= &HAB) Then
                '00-FF(1�o�C�g��)----��----
                lngFit = 0
                Exit For
            
            'AC-D7(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HAC And bytCode(i + 1) <= &HD7) Then
                '00-FF(1�o�C�g��)----�n���O��----
                lngFit = 0
                Exit For
            
            'D8-DF(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HD8 And bytCode(i + 1) <= &HDF) Then
                '00-FF(1�o�C�g��)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            'E0-F7(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HE0 And bytCode(i + 1) <= &HF7) Then
                '00-FF(1�o�C�g��)----�O��----
                lngFit = 0
                Exit For
            
            'F8-FF(2�o�C�g��)
            ElseIf (bytCode(i + 1) >= &HF8 And bytCode(i + 1) <= &HFF) Then
                '00-FF(1�o�C�g��)
                lngFit = lngFit + (2 * Multi_ByteWeight)
            
            End If
            i = i + 1
        End If
    Next i
    JudgeUNI = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----UTF7�֌W
' �֐���    : JudgeUTF7
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : UTF7�̕����R�[�h����(�\��)�m�����v�Z����
' ���l      :
Private Function JudgeUTF7(ByRef bytCode() As Byte, _
                           Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngWrk As Long
    Dim str64 As String
    Dim bln64 As Boolean
    str64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim lngUB As Long
    Dim lngBY As Long
    Dim lngXB As Long
    Dim lngXX As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    lngWrk = 0
    
    For i = 0 To lngUB
        '+�`-�܂ł�BASE64ENCODE
        If bytCode(i) = Asc("+") And bln64 = False Then
            lngWrk = 1
            bln64 = True
        ElseIf bytCode(i) = Asc("-") Then
            If lngWrk <= 0 Then
                lngWrk = lngWrk + 1
                lngFit = lngFit + (lngWrk * SingleByteWeight)
            ElseIf lngWrk = 1 Then
                lngWrk = lngWrk + 1
                lngFit = lngFit + (lngWrk * Multi_ByteWeight)
            ElseIf lngWrk >= 4 And lngXB < 6 And _
                   ((InStr(str64, Chr(bytCode(i - 1))) - 1) And lngXX) = 0 Then
                lngWrk = lngWrk + 1
                lngFit = lngFit + (lngWrk * Multi_ByteWeight)
            End If
            lngWrk = 0
            bln64 = False
        Else
            If bln64 = True Then
                'BASE64ENCODE��
                If InStr(str64, Chr(bytCode(i))) > 0 Then
                    lngBY = Int((lngWrk * 6) / 8)
                    lngXB = (lngWrk * 6) - (lngBY * 8)
                    lngXX = (2 ^ lngXB) - 1
                    lngWrk = lngWrk + 1
                Else
                    lngWrk = 0
                    bln64 = False
                End If
            Else
                '20-7E(1�o�C�g��)
                If (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
                    lngFit = lngFit + (1 * SingleByteWeight)
        
                '00-1F, 7F(1�o�C�g��)
                ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                        bytCode(i) = &H7F Then
                     lngFit = lngFit + (1 * SingleByteWeight)
                End If
            End If
        End If
    Next i
    JudgeUTF7 = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

'----UTF8�֌W
' �֐���    : JudgeUTF8
' �Ԃ�l    : ���茋�ʊm���i���j
' ������    : bytCode : ���蕶���f�[�^
'           : fixFlag : �m�蔻�f�L��
' �@�\����  : UTF8�̕����R�[�h����(�\��)�m�����v�Z����
' ���l      :
Private Function JudgeUTF8(ByRef bytCode() As Byte, _
                           Optional fixFlag As Boolean = False) As Integer
    Dim i As Long
    Dim lngFit As Long
    Dim lngUB As Long
    
    lngUB = JUDGESIZEMAX - 1
    If lngUB > UBound(bytCode()) Then
        lngUB = UBound(bytCode())
    End If
    For i = 0 To lngUB
        If fixFlag Then
            'BOM
            If bytCode(i) = &HEF Then
                If i <= UBound(bytCode) - 2 Then
                    If bytCode(i + 1) = &HBB And _
                       bytCode(i + 2) = &HBF Then
                        JudgeUTF8 = JUDGEFIX
                        Exit Function
                    End If
                End If
            End If
        End If
        
        'AND FC(1�o�C�g��) + 80-BF(2-6�o�C�g��)
        If (bytCode(i) And &HFC) = &HFC Then
            If i <= UBound(bytCode) - 5 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) And _
                   (bytCode(i + 4) >= &H80 And bytCode(i + 4) <= &HBF) And _
                   (bytCode(i + 5) >= &H80 And bytCode(i + 5) <= &HBF) Then
                    lngFit = lngFit + (6 * Multi_ByteWeight)
                    i = i + 5
                End If
            End If
        
        'AND F8(1�o�C�g��) + 80-BF(2-5�o�C�g��)
        ElseIf (bytCode(i) And &HF8) = &HF8 Then
            If i <= UBound(bytCode) - 4 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) And _
                   (bytCode(i + 4) >= &H80 And bytCode(i + 4) <= &HBF) Then
                    lngFit = lngFit + (5 * Multi_ByteWeight)
                    i = i + 4
                End If
            End If
            
        'AND F0(1�o�C�g��) + 80-BF(2-4�o�C�g��)
        ElseIf (bytCode(i) And &HF0) = &HF0 Then
            If i <= UBound(bytCode) - 3 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) And _
                   (bytCode(i + 3) >= &H80 And bytCode(i + 3) <= &HBF) Then
                    lngFit = lngFit + (4 * Multi_ByteWeight)
                    i = i + 3
                End If
            End If
        
        'AND E0(1�o�C�g��) + 80-BF(2-3�o�C�g��)
        ElseIf (bytCode(i) And &HE0) = &HE0 Then
            If i <= UBound(bytCode) - 2 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) And _
                   (bytCode(i + 2) >= &H80 And bytCode(i + 2) <= &HBF) Then
                    lngFit = lngFit + (3 * Multi_ByteWeight)
                    i = i + 2
                End If
            End If
        
        'AND C0(1�o�C�g��) + 80-BF(2�o�C�g��)
        ElseIf (bytCode(i) And &HC0) = &HC0 Then
            If i <= UBound(bytCode) - 1 Then
                If (bytCode(i + 1) >= &H80 And bytCode(i + 1) <= &HBF) Then
                    lngFit = lngFit + (2 * Multi_ByteWeight)
                    i = i + 1
                End If
            End If

        '20-7E(1�o�C�g��)
        ElseIf (bytCode(i) >= &H20 And bytCode(i) <= &H7E) Then
            lngFit = lngFit + (1 * SingleByteWeight)

        '00-1F, 7F(1�o�C�g��)
        ElseIf (bytCode(i) >= &H0 And bytCode(i) <= &H1F) Or _
                bytCode(i) = &H7F Then
            lngFit = lngFit + (1 * SingleByteWeight)
        End If
    Next i
    JudgeUTF8 = (lngFit * 100) / ((lngUB + 1) * Multi_ByteWeight)
End Function

