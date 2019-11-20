VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} saee_change 
   Caption         =   "�����޸�ͼǩ��~saeeͼǩר��~��ʦ����Ʒ��"
   ClientHeight    =   6600
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6996
   OleObjectBlob   =   "saee_change.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "saee_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

End Sub
'��ʾ����ͼ��
Public Function addLayers()
    Dim objLayer As AcadLayer '����ͼ�����
    For Each objLayer In ThisDrawing.Layers
        cmbBoxLayer.AddItem objLayer.Name
    Next
    'cmbBoxLayer.Text = "puresino-ͼ��ͼǩ"
End Function
'����ѡ��Ŀɼ���
Private Sub cmbRegion_Change()
    'Call addRegionToSset(cmbBoxLayer.Text, cmbRegion.Text)
    If cmbRegion.Text = "��������" Then
        cmdPickRegion.Visible = False
    ElseIf cmbRegion.Text = "ѡ��Χ" Then
        cmdPickRegion.Visible = True
    End If
    
End Sub

Private Sub cmdFind_Click()
    If ssetofRegion.Count = 0 And cmbRegion.Text = "��������" Then
        Call addRegionToSset(cmbBoxLayer.Text, cmbRegion.Text)
    ElseIf ssetofRegion.Count = 0 And cmbRegion.Text = "ѡ��Χ" Then
        MsgBox "��ѡ��Χ"
    End If
    
    If ssetofRegion.Count = 0 Then
        MsgBox "����ѡ��Χ"
    Else
        sumBR = ssetofRegion.Count  '��ȡѡ������ҳ��
        
        Dim objBR As AcadBlockReference
        Dim blockATT As Variant
        
        Set objBR = ssetofRegion.Item(0)  'ȡһ�������
        blockATT = objBR.GetAttributes   '��ȡ����յ�����
        objName = blockATT(4).TextString  '��ȡ��Ŀ����1
        '�����޸����ݵ�Ĭ��ֵ
        sumTB.Text = sumBR '��������
        objNameTB.Text = objName '������Ŀ����
        numTB.Text = 1  '���ó�ʼֵ
        picNumTB.Text = picNumSplit(blockATT(13).TextString) '����Ĭ��ͼ��ĸ��
        specTB.Text = blockATT(12).TextString '����רҵĬ��ֵ
        stageTB.Text = blockATT(11).TextString '���ý׶�Ĭ��ֵ
        dateTB.Text = blockATT(6).TextString '��������Ĭ��ֵ
        editionTB.Text = blockATT(7).TextString '���ð��Ĭ��ֵ
        
    End If
    
End Sub
'����ͼ�Ž�ȡ��ĸ��
Public Function picNumSplit(ByVal picNum As String) As String
    Dim pic As Variant
    Dim i As Integer
    
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
    picNumSplit = picNew
End Function

'ʰȡͼ��
Private Sub cmdPickLayer_Click()
    On Error Resume Next
    
    Dim objEntity As AcadEntity
    Dim pt
    Me.Hide  '���ضԻ��򣬰ѽ������ͼ��
    ThisDrawing.Utility.GetEntity objEntity, pt 'ʰȡͼԪ
    cmbBoxLayer.Text = objEntity.Layer  '��ֵͼ��
    If Err <> 0 Then
        Err.Clear
        MsgBox "��ѡ��ͼ���ϵ�ͼԪ"
    End If
    Me.show
End Sub
'���÷�Χ
Public Function setRegion()
    cmbRegion.AddItem "��������"
    cmbRegion.AddItem "ѡ��Χ"  '�������
    cmbRegion.Text = "��������"   'Ĭ����ʾ
End Function
'������Χ�ڵ�ѡ��
Public Function addRegionToSset(ByVal layerName As String, ByVal region As String)
    Dim fType(2) As Integer  '����һ�����飬����ɸѡѡ��
    Dim fData(2) As Variant '����һ�����飬����ɸѡѡ��
    ssetofRegion.Clear '���ú��������ѡ��
    fType(0) = 8: fData(0) = layerName   '����ͼ��
    fType(1) = 100: fData(1) = "AcDbBlockReference" 'ɸѡͼԪ����Ϊ�����
    
    fType(2) = 2: fData(2) = "puresino-ͼ��ͼǩ" 'ɸѡ���������Ϊpuresino-ͼ��ͼǩ
    
    'ssetofRegion.Select acSelectionSetAll, , , fType, fData 'ѡ������ͼԪ
    
    If StrComp(region, "��������", vbTextCompare) = 0 Then
        ssetofRegion.Select acSelectionSetAll, , , fType, fData
    ElseIf StrComp(region, "ѡ��Χ", vbTextCompare) = 0 Then
        'UserForm1.Hide  '���ضԻ��򣬰ѽ������ͼ��
        ssetofRegion.SelectOnScreen fType, fData
       ' UserForm1.show
    End If
End Function
'ѡ������
Private Sub cmdPickRegion_Click()
    Me.Hide  '���ضԻ��򣬰ѽ������ͼ��
    Call addRegionToSset(cmbBoxLayer.Text, cmbRegion.Text)
    Me.show
End Sub
'ȷ�ϰ�ť,�޸�ͼǩ��Ϣ
Public Function okChange()
    If ssetofRegion.Count = 0 And cmbRegion.Text = "��������" Then
        Call addRegionToSset(cmbBoxLayer.Text, cmbRegion.Text)
    ElseIf ssetofRegion.Count = 0 And cmbRegion.Text = "ѡ��Χ" Then
        MsgBox "��ѡ��Χ"
    End If
    
    Dim objBR As AcadBlockReference  '����һ�������
    Dim ptmin As Variant
    Dim ptmax As Variant
    
    '��ȡƫ����
    If ssetofRegion.Count <> 0 Then
        Set objBR = ssetofRegion.Item(1)
        objBR.GetBoundingBox ptmin, ptmax
        offsety = Abs(ptmax(1) - ptmin(1))
        offsetX = Abs(ptmax(0) - ptmin(0))
   

        '�õ��������п���պͲ�������������
        Dim objBRarr() As Variant
        ReDim objBRarr(0 To (ssetofRegion.Count - 1), 0 To 1) '����һ����ά����
        Dim i As Integer
        i = 0
        For Each objBR In ssetofRegion
            Set objBRarr(i, 0) = objBR  '������շ��ڵ�һ��
            objBRarr(i, 1) = objBR.InsertionPoint '�������������ڵڶ���
            i = i + 1
        Next
        
        ssetofRegion.Clear
        
        '����������
        If opbyx.Value = True Then
            sortBRATTyx objBRarr, offsety
        ElseIf opbxy.Value = True Then
            sortBRATTxy objBRarr, offsetX
        End If
        
        '����ͼԪ����
        Dim blockATT As Variant
    
        For i = 0 To UBound(objBRarr)  '����ѡ��
            blockATT = objBRarr(i, 0).GetAttributes '��ȡ���տ������
            blockATT(4).TextString = objNameTB.Text '�޸���Ŀ����
            blockATT(8).TextString = sumTB.Text  '�޸���ҳ��Ϊ
            blockATT(9).TextString = i + CInt(numTB.Text) '�޸�ҳ��Ϊi
            If i < 10 Then
                blockATT(13).TextString = picNumTB.Text & "-0" & i + CInt(numTB.Text)  '�޸�ͼ��Ϊxx-xx-0i
            Else
                blockATT(13).TextString = picNumTB.Text & "-" & i + CInt(numTB.Text)  '�޸�ͼ��Ϊxx-xx-xx
            End If
            
            blockATT(12).TextString = specTB.Text '����רҵ
            blockATT(11).TextString = stageTB.Text '���ý׶�
            blockATT(6).TextString = dateTB.Text '��������
            blockATT(7).TextString = editionTB.Text '���ð��
            
        Next i
        
        MsgBox "�޸ĳɹ�"
    Else
         MsgBox "���Ȳ�ѯ����"
    End If
End Function

Private Sub CommandButton2_Click()
    Call okChange
End Sub
'ȡ����ť
Private Sub CommandButton3_Click()
    Unload Me
End Sub



Private Sub UserForm_Initialize()
    Set ssetofRegion = creatSSet("myRegionSset")
    Call addLayers
    Call setRegion
    opbyx.Value = True  'Ĭ�����к���
End Sub

