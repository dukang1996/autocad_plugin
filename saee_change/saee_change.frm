VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} saee_change 
   Caption         =   "批量修改图签（~saee图签专用~康师傅出品）"
   ClientHeight    =   6600
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6996
   OleObjectBlob   =   "saee_change.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "saee_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

End Sub
'显示所有图层
Public Function addLayers()
    Dim objLayer As AcadLayer '定义图层对象
    For Each objLayer In ThisDrawing.Layers
        cmbBoxLayer.AddItem objLayer.Name
    Next
    'cmbBoxLayer.Text = "puresino-图框图签"
End Function
'设置选择的可见性
Private Sub cmbRegion_Change()
    'Call addRegionToSset(cmbBoxLayer.Text, cmbRegion.Text)
    If cmbRegion.Text = "整个布局" Then
        cmdPickRegion.Visible = False
    ElseIf cmbRegion.Text = "选择范围" Then
        cmdPickRegion.Visible = True
    End If
    
End Sub

Private Sub cmdFind_Click()
    If ssetofRegion.Count = 0 And cmbRegion.Text = "整个布局" Then
        Call addRegionToSset(cmbBoxLayer.Text, cmbRegion.Text)
    ElseIf ssetofRegion.Count = 0 And cmbRegion.Text = "选择范围" Then
        MsgBox "请选择范围"
    End If
    
    If ssetofRegion.Count = 0 Then
        MsgBox "请先选择范围"
    Else
        sumBR = ssetofRegion.Count  '获取选区的总页数
        
        Dim objBR As AcadBlockReference
        Dim blockATT As Variant
        
        Set objBR = ssetofRegion.Item(0)  '取一个块参照
        blockATT = objBR.GetAttributes   '获取块参照的属性
        objName = blockATT(4).TextString  '获取项目名称1
        '设置修改内容的默认值
        sumTB.Text = sumBR '设置总数
        objNameTB.Text = objName '设置项目名称
        numTB.Text = 1  '设置初始值
        picNumTB.Text = picNumSplit(blockATT(13).TextString) '设置默认图号母段
        specTB.Text = blockATT(12).TextString '设置专业默认值
        stageTB.Text = blockATT(11).TextString '设置阶段默认值
        dateTB.Text = blockATT(6).TextString '设置日期默认值
        editionTB.Text = blockATT(7).TextString '设置版次默认值
        
    End If
    
End Sub
'处理图号截取字母段
Public Function picNumSplit(ByVal picNum As String) As String
    Dim pic As Variant
    Dim i As Integer
    
    i = 1
    
    If InStr(1, picNum, "-") = 0 Then  '如果传入字符串中不包含“-”
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

'拾取图层
Private Sub cmdPickLayer_Click()
    On Error Resume Next
    
    Dim objEntity As AcadEntity
    Dim pt
    Me.Hide  '隐藏对话框，把焦点给绘图区
    ThisDrawing.Utility.GetEntity objEntity, pt '拾取图元
    cmbBoxLayer.Text = objEntity.Layer  '赋值图层
    If Err <> 0 Then
        Err.Clear
        MsgBox "请选择图层上的图元"
    End If
    Me.show
End Sub
'设置范围
Public Function setRegion()
    cmbRegion.AddItem "整个布局"
    cmbRegion.AddItem "选择范围"  '添加内容
    cmbRegion.Text = "整个布局"   '默认显示
End Function
'创建范围内的选择集
Public Function addRegionToSset(ByVal layerName As String, ByVal region As String)
    Dim fType(2) As Integer  '定义一个数组，用于筛选选择集
    Dim fData(2) As Variant '定义一个数组，用于筛选选择集
    ssetofRegion.Clear '调用函数先清空选择集
    fType(0) = 8: fData(0) = layerName   '设置图层
    fType(1) = 100: fData(1) = "AcDbBlockReference" '筛选图元类型为块参照
    
    fType(2) = 2: fData(2) = "puresino-图框图签" '筛选块参照名称为puresino-图框图签
    
    'ssetofRegion.Select acSelectionSetAll, , , fType, fData '选择所有图元
    
    If StrComp(region, "整个布局", vbTextCompare) = 0 Then
        ssetofRegion.Select acSelectionSetAll, , , fType, fData
    ElseIf StrComp(region, "选择范围", vbTextCompare) = 0 Then
        'UserForm1.Hide  '隐藏对话框，把焦点给绘图区
        ssetofRegion.SelectOnScreen fType, fData
       ' UserForm1.show
    End If
End Function
'选择区域
Private Sub cmdPickRegion_Click()
    Me.Hide  '隐藏对话框，把焦点给绘图区
    Call addRegionToSset(cmbBoxLayer.Text, cmbRegion.Text)
    Me.show
End Sub
'确认按钮,修改图签信息
Public Function okChange()
    If ssetofRegion.Count = 0 And cmbRegion.Text = "整个布局" Then
        Call addRegionToSset(cmbBoxLayer.Text, cmbRegion.Text)
    ElseIf ssetofRegion.Count = 0 And cmbRegion.Text = "选择范围" Then
        MsgBox "请选择范围"
    End If
    
    Dim objBR As AcadBlockReference  '定义一个块参照
    Dim ptmin As Variant
    Dim ptmax As Variant
    
    '获取偏移量
    If ssetofRegion.Count <> 0 Then
        Set objBR = ssetofRegion.Item(1)
        objBR.GetBoundingBox ptmin, ptmax
        offsety = Abs(ptmax(1) - ptmin(1))
        offsetX = Abs(ptmax(0) - ptmin(0))
   

        '得到包含所有块参照和插入点坐标的数组
        Dim objBRarr() As Variant
        ReDim objBRarr(0 To (ssetofRegion.Count - 1), 0 To 1) '定义一个两维数组
        Dim i As Integer
        i = 0
        For Each objBR In ssetofRegion
            Set objBRarr(i, 0) = objBR  '将块参照放在第一列
            objBRarr(i, 1) = objBR.InsertionPoint '将插入点坐标放在第二列
            i = i + 1
        Next
        
        ssetofRegion.Clear
        
        '按坐标排序
        If opbyx.Value = True Then
            sortBRATTyx objBRarr, offsety
        ElseIf opbxy.Value = True Then
            sortBRATTxy objBRarr, offsetX
        End If
        
        '更改图元属性
        Dim blockATT As Variant
    
        For i = 0 To UBound(objBRarr)  '遍历选择集
            blockATT = objBRarr(i, 0).GetAttributes '获取参照快的属性
            blockATT(4).TextString = objNameTB.Text '修改项目名称
            blockATT(8).TextString = sumTB.Text  '修改总页数为
            blockATT(9).TextString = i + CInt(numTB.Text) '修改页数为i
            If i < 10 Then
                blockATT(13).TextString = picNumTB.Text & "-0" & i + CInt(numTB.Text)  '修改图号为xx-xx-0i
            Else
                blockATT(13).TextString = picNumTB.Text & "-" & i + CInt(numTB.Text)  '修改图号为xx-xx-xx
            End If
            
            blockATT(12).TextString = specTB.Text '设置专业
            blockATT(11).TextString = stageTB.Text '设置阶段
            blockATT(6).TextString = dateTB.Text '设置日期
            blockATT(7).TextString = editionTB.Text '设置版次
            
        Next i
        
        MsgBox "修改成功"
    Else
         MsgBox "请先查询内容"
    End If
End Function

Private Sub CommandButton2_Click()
    Call okChange
End Sub
'取消按钮
Private Sub CommandButton3_Click()
    Unload Me
End Sub



Private Sub UserForm_Initialize()
    Set ssetofRegion = creatSSet("myRegionSset")
    Call addLayers
    Call setRegion
    opbyx.Value = True  '默认先行后列
End Sub

