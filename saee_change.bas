Attribute VB_Name = "ģ��1"
'���幫������
Public ssetofRegion As AcadSelectionSet    '���ڱ����޸ķ�Χ
Public objName As String  '���ڱ�����Ŀ����
Public sumBR As Integer   '���ڱ�����ҳ��

Public Sub demo1()
    Dim objBR As AcadBlockReference  '����һ�������
    Dim sset As AcadSelectionSet '����һ��ѡ��
    Dim fType(2) As Integer  '����һ�����飬����ɸѡѡ��
    Dim fData(2) As Variant '����һ�����飬����ɸѡѡ��
    Dim offsety As Integer
    Dim offsetX As Integer
    Dim ptmin As Variant
    Dim ptmax As Variant

    Set sset = creatSSet("poltSSet")  '����һ����ΪpoleSSet��ѡ��

    fType(0) = 8: fData(0) = "puresino-ͼ��ͼǩ"   '����ͼ��
    fType(1) = 100: fData(1) = "AcDbBlockReference" 'ɸѡͼԪ����Ϊ�����
    fType(2) = 2: fData(2) = "puresino-ͼ��ͼǩ" 'ɸѡ���������Ϊpuresino-ͼ��ͼǩ
    sset.SelectOnScreen fType, fData 'ѡ������ͼԪ

    '��ȡƫ����
    Set objBR = sset.Item(0)
    objBR.GetBoundingBox ptmin, ptmax
    offsety = Abs(ptmax(1) - ptmin(1))
    offsetX = Abs(ptmax(0) - ptmin(0))

    '�õ��������п���պͲ�������������
    Dim objBRarr() As Variant
    ReDim objBRarr(0 To (sset.Count - 1), 0 To 1) '����һ����ά����
    Dim i As Integer
    i = 0
    For Each objBR In sset
        Set objBRarr(i, 0) = objBR  '������շ��ڵ�һ��
        objBRarr(i, 1) = objBR.InsertionPoint '�������������ڵڶ���
        i = i + 1
    Next

    sset.Delete

    '����������
    sortBRATTyx objBRarr, offsety

    '����ͼԪ����
    Dim blockATT As Variant

    For i = 0 To UBound(objBRarr)  '����ѡ��
        blockATT = objBRarr(i, 0).GetAttributes '��ȡ���տ������
        blockATT(8).TextString = UBound(objBRarr) + 1 '�޸���ҳ��Ϊ
        blockATT(9).TextString = i + 1 '�޸�ҳ��Ϊi
        If i < 10 Then
            blockATT(13).TextString = "QP-DY-0" & i + 1 '�޸�ͼ��ΪQP-DY-0i
        Else
            blockATT(13).TextString = "QP-DY-" & i + 1
        End If

    Next i

End Sub



'����ѡ�񼯺���
Public Function creatSSet(ByVal SSetName As String) As AcadSelectionSet
    Dim sset As AcadSelectionSet
    Dim i As Integer
    For i = 0 To ThisDrawing.SelectionSets.Count - 1
        Set sset = ThisDrawing.SelectionSets.Item(i)
        If StrComp(sset.Name, SSetName, vbTextCompare) = 0 Then '���������ѡ�񼯺����е�ѡ������
            sset.Delete
            Exit For
        End If
    Next i
    Set creatSSet = ThisDrawing.SelectionSets.Add(SSetName)
End Function
'�����������к���
Public Function sortBRATTyx(ByRef objarr() As Variant, ByVal offsety)
    Dim i As Integer
    Dim j As Integer
    Dim objArrTemp As AcadBlockReference
    Dim ptInsert As Variant
    'y������
    For i = 0 To UBound(objarr) - 1
        For j = i + 1 To UBound(objarr)
            If objarr(i, 1)(1) < objarr(j, 1)(1) Then
                Set objArrTemp = objarr(i, 0)
                Set objarr(i, 0) = objarr(j, 0)
                Set objarr(j, 0) = objArrTemp
                ptInsert = objarr(i, 1)
                objarr(i, 1) = objarr(j, 1)
                objarr(j, 1) = ptInsert
            End If
        Next j
    Next i
    '��������
    For i = 0 To UBound(objarr) - 1
        For j = i + 1 To UBound(objarr)
            If Abs(objarr(i, 1)(1) - objarr(j, 1)(1)) < offsety And objarr(i, 1)(0) > objarr(j, 1)(0) Then
                Set objArrTemp = objarr(i, 0)
                Set objarr(i, 0) = objarr(j, 0)
                Set objarr(j, 0) = objArrTemp
                ptInsert = objarr(i, 1)
                objarr(i, 1) = objarr(j, 1)
                objarr(j, 1) = ptInsert
            End If
        Next j
    Next i

End Function

'�����������к���
Public Function sortBRATTxy(ByRef objarr() As Variant, ByVal offsety)
    Dim i As Integer
    Dim j As Integer
    Dim objArrTemp As AcadBlockReference
    Dim ptInsert As Variant
    'x������
    For i = 0 To UBound(objarr) - 1
        For j = i + 1 To UBound(objarr)
            If objarr(i, 1)(0) > objarr(j, 1)(0) Then
                Set objArrTemp = objarr(i, 0)
                Set objarr(i, 0) = objarr(j, 0)
                Set objarr(j, 0) = objArrTemp
                ptInsert = objarr(i, 1)
                objarr(i, 1) = objarr(j, 1)
                objarr(j, 1) = ptInsert
            End If
        Next j
    Next i
    '��������
    For i = 0 To UBound(objarr) - 1
        For j = i + 1 To UBound(objarr)
            If Abs(objarr(i, 1)(0) - objarr(j, 1)(0)) < offsety And objarr(i, 1)(1) < objarr(j, 1)(1) Then
                Set objArrTemp = objarr(i, 0)
                Set objarr(i, 0) = objarr(j, 0)
                Set objarr(j, 0) = objArrTemp
                ptInsert = objarr(i, 1)
                objarr(i, 1) = objarr(j, 1)
                objarr(j, 1) = ptInsert
            End If
        Next j
    Next i

End Function

Public Sub gtq()
    UserForm1.show
End Sub

Sub picN()
    Dim pic As Variant
    Dim picNum As String
    Dim picNew As String
    Dim i As Integer
    
    picNum = "qp-dd-uu-09"
    i = 1
    If InStr(1, picNum, "-") = 0 Then  '��������ַ����в�������-��
        picNew = picNum
    Else
        pic = Split(picNum, "-")
        picNew = pic(0)
        
        If UBound(pic) >= 1 Then
            Do
            picNew = picNew & "-" & pic(i)
            i = i + 1
            Loop Until i > UBound(pic) - 1
        End If
        
    End If

    Stop
End Sub
