VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BLOCK_CHANGE 
   Caption         =   "�������޸���"
   ClientHeight    =   5652
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5256
   OleObjectBlob   =   "BLOCK_CHANGE.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "BLOCK_CHANGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���幫������
Public ssetofRegion As AcadSelectionSet    '���ڱ����޸ķ�Χ
Dim objBR As AcadBlockReference   '�������յ�һ������
Dim blockATT As Variant  '�������������
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
    Set creatSSet = ThisDrawing.SelectionSets.Add(SSetName) '���û�������򴴽���ѡ��
End Function


Private Sub cmbATTlist_Change()
    If cmbATTlist.Text <> "" Then
        For i = 0 To UBound(blockATT)
            If blockATT(i).TagString = cmbATTlist.Text Then
                TBatt.Text = blockATT(i).TextString
            End If
        Next
    End If
End Sub

Private Sub cmdbExit_Click()
    Unload Me
End Sub

'ʰȡ��
Private Sub cmdBpickBlock_Click()

    On Error Resume Next '��������ʱִ����һ�Σ������ѡû��ͼԪ���󣬽��ᱨ��
    
    Dim objBR As AcadBlockReference  '����һ������ն���
    Dim objEntity As AcadEntity '����һ��ͼԪ
    Dim pt
    
    BLOCK_CHANGE.Hide  '���ضԻ��򣬰ѽ������ͼ��
    
    ThisDrawing.Utility.GetEntity objEntity, pt 'ʰȡͼԪ
    
    If objEntity.EntityName <> "AcDbBlockReference" Then
    
        MsgBox "��ѡ���ͼԪ"
    Else
    
        TBpick.Text = objEntity.Name  '��ֵͼ��
    End If
    
    If Err <> 0 Then  '�����������
        Err.Clear  '��մ���
        'MsgBox "��ѡ�񲼾��еĿ�"
    End If
    
    BLOCK_CHANGE.show

End Sub

'���÷�Χ
Public Function setRegion()

    cmbRegion.AddItem "��������"
    
    cmbRegion.AddItem "ѡ��Χ"  '�������
    
    cmbRegion.Text = "��������"   'Ĭ����ʾ
    
End Function

'����ѡ��Ŀɼ���
Private Sub cmbRegion_Change()
    If cmbRegion.Text = "��������" Then
        cmdbPickRegion.Visible = False
    ElseIf cmbRegion.Text = "ѡ��Χ" Then
        cmdbPickRegion.Visible = True
    End If
    
End Sub

'������Χ�ڵ�ѡ��
Public Function addRegionToSset(ByVal blockName As String, ByVal region As String)
    Dim fType(1) As Integer  '����һ�����飬����ɸѡѡ��
    Dim fData(1) As Variant '����һ�����飬����ɸѡѡ��
    ssetofRegion.Clear '���ú��������ѡ��
    
    fType(0) = 100: fData(1) = "AcDbBlockReference" 'ɸѡͼԪ����Ϊ�����
    
    fType(1) = 2: fData(1) = blockName 'ɸѡ���������
    
    '����ѡ��Χѡ��ѡ��ʽ
    If StrComp(region, "��������", vbTextCompare) = 0 Then
        ssetofRegion.Select acSelectionSetAll, , , fType, fData
    ElseIf StrComp(region, "ѡ��Χ", vbTextCompare) = 0 Then
        ssetofRegion.SelectOnScreen fType, fData
      
    End If
End Function


'ѡ������
Private Sub cmdbPickRegion_Click()
    BLOCK_CHANGE.Hide  '���ضԻ��򣬰ѽ������ͼ��
    Call addRegionToSset(TBpick.Text, cmbRegion.Text)
    BLOCK_CHANGE.show
End Sub

'��ѯ����ʵ��
Private Sub cmdbFind_Click()
    '��ֹ�յ�ѡ��
    If ssetofRegion.Count = 0 And cmbRegion.Text = "��������" Then
        Call addRegionToSset(TBpick.Text, cmbRegion.Text)
    ElseIf ssetofRegion.Count = 0 And cmbRegion.Text = "ѡ��Χ" Then
        MsgBox "��ѡ��Χ"
    End If
    
    If ssetofRegion.Count = 0 Then
        MsgBox "����ѡ��Χ"
    Else
        Dim sumBR As Integer
        sumBR = ssetofRegion.Count  '��ȡѡ������ҳ��
        MsgBox "���ҵ�" & sumBR & "����"
        
        'ʹ�����Կ�
        frmATT.Enabled = True
        cmdbYES.Enabled = True
        
        Dim i As Integer
        
        Set objBR = ssetofRegion.Item(0)  'ȡһ�������
        blockATT = objBR.GetAttributes   '��ȡ����յ�����
        
        For i = 0 To UBound(blockATT)
            cmbATTlist.AddItem blockATT(i).TagString
        Next
        
    End If
End Sub
'ȷ�ϰ�ť,�޸Ŀ���Ϣ
Public Function okChange()
    'ȷ��ǰ����Ҫ�ȱ�֤ѡ�񼯷ǿ�
    On Error Resume Next
    If ssetofRegion.Count = 0 And cmbRegion.Text = "��������" Then
        Call addRegionToSset(TBpick.Text, cmbRegion.Text)
    ElseIf ssetofRegion.Count = 0 And cmbRegion.Text = "ѡ��Χ" Then
        MsgBox "��ѡ��Χ"
    End If
    
    Dim objBRt As AcadBlockReference  '����һ�������
    Dim ptmin As Variant
    Dim ptmax As Variant
    
    '��ȡƫ����
    If ssetofRegion.Count <> 0 Then
        'Set objBR = ssetofRegion.Item(0)
        objBR.GetBoundingBox ptmin, ptmax
        offsety = Abs(ptmax(1) - ptmin(1)) '����x���ƫ�Χ
        offsetX = Abs(ptmax(0) - ptmin(0)) '����y���ƫ�Χ
   

        '�õ��������п���պͲ�������������
        Dim objBRarr() As Variant
        ReDim objBRarr(0 To (ssetofRegion.Count - 1), 0 To 1) '����һ����ά����
        Dim i As Integer
        i = 0
        For Each objBRt In ssetofRegion
            Set objBRarr(i, 0) = objBRt  '������շ��ڵ�һ��
            objBRarr(i, 1) = objBRt.InsertionPoint '�������������ڵڶ���
            i = i + 1
        Next
        
        'ssetofRegion.Clear  '��Ԫ�ش浽��������ѡ��
        
        '����������
        If opYX.Value = True Then
            sortBRATTyx objBRarr, offsety
        ElseIf opXY.Value = True Then
            sortBRATTxy objBRarr, offsetX
        End If
        
        '����ͼԪ����
        Dim blockATTt As Variant
        Dim j As Integer
        Dim no As Integer
        
        For i = 0 To UBound(objBRarr)  '����ѡ��
            blockATTt = objBRarr(i, 0).GetAttributes '��ȡ���տ������
            
             If cmbATTlist.Text <> "" Then
                For j = 0 To UBound(blockATTt)
                    If blockATTt(j).TagString = cmbATTlist.Text Then
                        
                        If opYES.Value = True Then
                            no = i + CInt(sortStart.Text)
                            If no < 10 Then
                                blockATTt(j).TextString = TBatt.Text & "0" & no  '�޸�
                            Else
                                blockATTt(j).TextString = TBatt.Text & no  '�޸�ͼ��Ϊxx-xx-xx
                            End If
                        Else
                             blockATTt(j).TextString = TBatt.Text
                        End If
                        
                    End If
                Next j
            End If
            
        Next i
        
        MsgBox "�޸ĳɹ�"
    Else
         MsgBox "���Ȳ�ѯ����"
    End If
    
    If Err <> 0 Then  '�����������
        Err.Clear  '��մ���
        'MsgBox "��ѡ�񲼾��еĿ�"
    End If
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

Private Sub cmdbYES_Click()
    Call okChange
End Sub

'���ѡ�����ǣ���ʾ��������
Private Sub opYES_Change()

    If opYES.Value = True Then
        frmSort.Visible = True
        Label4.Visible = True
        sortStart.Visible = True
    Else
        frmSort.Visible = False
        Label4.Visible = False
        sortStart.Visible = False
    End If

End Sub



Private Sub UserForm_Initialize()
    Set ssetofRegion = creatSSet("myRegionSset")
    Call setRegion
    '�����ѯǰ�����Կ�Ĭ��ʧЧ
    frmATT.Enabled = False
    cmdbYES.Enabled = False
    opNO.Value = True 'Ĭ�ϲ����
    opYX.Value = True
    frmSort.Visible = False
    Label4.Visible = False
    sortStart.Visible = False
    
End Sub

