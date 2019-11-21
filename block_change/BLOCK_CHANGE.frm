VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BLOCK_CHANGE 
   Caption         =   "块属性修改器"
   ClientHeight    =   5652
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5256
   OleObjectBlob   =   "BLOCK_CHANGE.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "BLOCK_CHANGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'定义公共变量
Public ssetofRegion As AcadSelectionSet    '用于保存修改范围
Dim objBR As AcadBlockReference   '保存块参照的一个样本
Dim blockATT As Variant  '用来保存块属性
'创建选择集函数
Public Function creatSSet(ByVal SSetName As String) As AcadSelectionSet
    Dim sset As AcadSelectionSet
    Dim i As Integer
    For i = 0 To ThisDrawing.SelectionSets.Count - 1
        Set sset = ThisDrawing.SelectionSets.Item(i)
        If StrComp(sset.Name, SSetName, vbTextCompare) = 0 Then '如果创建的选择集和已有的选择集重名
            sset.Delete
            Exit For
        End If
    Next i
    Set creatSSet = ThisDrawing.SelectionSets.Add(SSetName) '如果没有重名则创建该选择集
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

'拾取块
Private Sub cmdBpickBlock_Click()

    On Error Resume Next '发生错误时执行下一段，如果点选没有图元对象，将会报错
    
    Dim objBR As AcadBlockReference  '定义一个块参照对象
    Dim objEntity As AcadEntity '定义一个图元
    Dim pt
    
    BLOCK_CHANGE.Hide  '隐藏对话框，把焦点给绘图区
    
    ThisDrawing.Utility.GetEntity objEntity, pt '拾取图元
    
    If objEntity.EntityName <> "AcDbBlockReference" Then
    
        MsgBox "请选择块图元"
    Else
    
        TBpick.Text = objEntity.Name  '赋值图层
    End If
    
    If Err <> 0 Then  '如果发生错误
        Err.Clear  '清空错误
        'MsgBox "请选择布局中的块"
    End If
    
    BLOCK_CHANGE.show

End Sub

'设置范围
Public Function setRegion()

    cmbRegion.AddItem "整个布局"
    
    cmbRegion.AddItem "选择范围"  '添加内容
    
    cmbRegion.Text = "整个布局"   '默认显示
    
End Function

'设置选择的可见性
Private Sub cmbRegion_Change()
    If cmbRegion.Text = "整个布局" Then
        cmdbPickRegion.Visible = False
    ElseIf cmbRegion.Text = "选择范围" Then
        cmdbPickRegion.Visible = True
    End If
    
End Sub

'创建范围内的选择集
Public Function addRegionToSset(ByVal blockName As String, ByVal region As String)
    Dim fType(1) As Integer  '定义一个数组，用于筛选选择集
    Dim fData(1) As Variant '定义一个数组，用于筛选选择集
    ssetofRegion.Clear '调用函数先清空选择集
    
    fType(0) = 100: fData(1) = "AcDbBlockReference" '筛选图元类型为块参照
    
    fType(1) = 2: fData(1) = blockName '筛选块参照名称
    
    '根据选择范围选择选择方式
    If StrComp(region, "整个布局", vbTextCompare) = 0 Then
        ssetofRegion.Select acSelectionSetAll, , , fType, fData
    ElseIf StrComp(region, "选择范围", vbTextCompare) = 0 Then
        ssetofRegion.SelectOnScreen fType, fData
      
    End If
End Function


'选择区域
Private Sub cmdbPickRegion_Click()
    BLOCK_CHANGE.Hide  '隐藏对话框，把焦点给绘图区
    Call addRegionToSset(TBpick.Text, cmbRegion.Text)
    BLOCK_CHANGE.show
End Sub

'查询功能实现
Private Sub cmdbFind_Click()
    '防止空的选择集
    If ssetofRegion.Count = 0 And cmbRegion.Text = "整个布局" Then
        Call addRegionToSset(TBpick.Text, cmbRegion.Text)
    ElseIf ssetofRegion.Count = 0 And cmbRegion.Text = "选择范围" Then
        MsgBox "请选择范围"
    End If
    
    If ssetofRegion.Count = 0 Then
        MsgBox "请先选择范围"
    Else
        Dim sumBR As Integer
        sumBR = ssetofRegion.Count  '获取选区的总页数
        MsgBox "共找到" & sumBR & "个块"
        
        '使能属性框
        frmATT.Enabled = True
        cmdbYES.Enabled = True
        
        Dim i As Integer
        
        Set objBR = ssetofRegion.Item(0)  '取一个块参照
        blockATT = objBR.GetAttributes   '获取块参照的属性
        
        For i = 0 To UBound(blockATT)
            cmbATTlist.AddItem blockATT(i).TagString
        Next
        
    End If
End Sub
'确认按钮,修改块信息
Public Function okChange()
    '确认前还是要先保证选择集非空
    On Error Resume Next
    If ssetofRegion.Count = 0 And cmbRegion.Text = "整个布局" Then
        Call addRegionToSset(TBpick.Text, cmbRegion.Text)
    ElseIf ssetofRegion.Count = 0 And cmbRegion.Text = "选择范围" Then
        MsgBox "请选择范围"
    End If
    
    Dim objBRt As AcadBlockReference  '定义一个块参照
    Dim ptmin As Variant
    Dim ptmax As Variant
    
    '获取偏移量
    If ssetofRegion.Count <> 0 Then
        'Set objBR = ssetofRegion.Item(0)
        objBR.GetBoundingBox ptmin, ptmax
        offsety = Abs(ptmax(1) - ptmin(1)) '设置x轴的偏差范围
        offsetX = Abs(ptmax(0) - ptmin(0)) '设置y轴的偏差范围
   

        '得到包含所有块参照和插入点坐标的数组
        Dim objBRarr() As Variant
        ReDim objBRarr(0 To (ssetofRegion.Count - 1), 0 To 1) '定义一个两维数组
        Dim i As Integer
        i = 0
        For Each objBRt In ssetofRegion
            Set objBRarr(i, 0) = objBRt  '将块参照放在第一列
            objBRarr(i, 1) = objBRt.InsertionPoint '将插入点坐标放在第二列
            i = i + 1
        Next
        
        'ssetofRegion.Clear  '将元素存到数组后清空选择集
        
        '按坐标排序
        If opYX.Value = True Then
            sortBRATTyx objBRarr, offsety
        ElseIf opXY.Value = True Then
            sortBRATTxy objBRarr, offsetX
        End If
        
        '更改图元属性
        Dim blockATTt As Variant
        Dim j As Integer
        Dim no As Integer
        
        For i = 0 To UBound(objBRarr)  '遍历选择集
            blockATTt = objBRarr(i, 0).GetAttributes '获取参照快的属性
            
             If cmbATTlist.Text <> "" Then
                For j = 0 To UBound(blockATTt)
                    If blockATTt(j).TagString = cmbATTlist.Text Then
                        
                        If opYES.Value = True Then
                            no = i + CInt(sortStart.Text)
                            If no < 10 Then
                                blockATTt(j).TextString = TBatt.Text & "0" & no  '修改
                            Else
                                blockATTt(j).TextString = TBatt.Text & no  '修改图号为xx-xx-xx
                            End If
                        Else
                             blockATTt(j).TextString = TBatt.Text
                        End If
                        
                    End If
                Next j
            End If
            
        Next i
        
        MsgBox "修改成功"
    Else
         MsgBox "请先查询内容"
    End If
    
    If Err <> 0 Then  '如果发生错误
        Err.Clear  '清空错误
        'MsgBox "请选择布局中的块"
    End If
End Function
'数组排序，先行后列
Public Function sortBRATTyx(ByRef objarr() As Variant, ByVal offsety)
    Dim i As Integer
    Dim j As Integer
    Dim objArrTemp As AcadBlockReference
    Dim ptInsert As Variant
    'y轴排序
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
    '二次排序
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


'数组排序，先列后行
Public Function sortBRATTxy(ByRef objarr() As Variant, ByVal offsety)
    Dim i As Integer
    Dim j As Integer
    Dim objArrTemp As AcadBlockReference
    Dim ptInsert As Variant
    'x轴排序
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
    '二次排序
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

'如果选中了是，显示其他内容
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
    '点击查询前，属性框默认失效
    frmATT.Enabled = False
    cmdbYES.Enabled = False
    opNO.Value = True '默认不编号
    opYX.Value = True
    frmSort.Visible = False
    Label4.Visible = False
    sortStart.Visible = False
    
End Sub

