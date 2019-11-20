Attribute VB_Name = "模块1"
'定义公共变量
Public ssetofRegion As AcadSelectionSet    '用于保存修改范围
Public objName As String  '用于保存项目名称
Public sumBR As Integer   '用于保存总页数

Public Sub demo1()
    Dim objBR As AcadBlockReference  '定义一个块参照
    Dim sset As AcadSelectionSet '定义一个选择集
    Dim fType(2) As Integer  '定义一个数组，用于筛选选择集
    Dim fData(2) As Variant '定义一个数组，用于筛选选择集
    Dim offsety As Integer
    Dim offsetX As Integer
    Dim ptmin As Variant
    Dim ptmax As Variant

    Set sset = creatSSet("poltSSet")  '设置一个名为poleSSet的选择集

    fType(0) = 8: fData(0) = "puresino-图框图签"   '设置图层
    fType(1) = 100: fData(1) = "AcDbBlockReference" '筛选图元类型为块参照
    fType(2) = 2: fData(2) = "puresino-图框图签" '筛选块参照名称为puresino-图框图签
    sset.SelectOnScreen fType, fData '选择所有图元

    '获取偏移量
    Set objBR = sset.Item(0)
    objBR.GetBoundingBox ptmin, ptmax
    offsety = Abs(ptmax(1) - ptmin(1))
    offsetX = Abs(ptmax(0) - ptmin(0))

    '得到包含所有块参照和插入点坐标的数组
    Dim objBRarr() As Variant
    ReDim objBRarr(0 To (sset.Count - 1), 0 To 1) '定义一个两维数组
    Dim i As Integer
    i = 0
    For Each objBR In sset
        Set objBRarr(i, 0) = objBR  '将块参照放在第一列
        objBRarr(i, 1) = objBR.InsertionPoint '将插入点坐标放在第二列
        i = i + 1
    Next

    sset.Delete

    '按坐标排序
    sortBRATTyx objBRarr, offsety

    '更改图元属性
    Dim blockATT As Variant

    For i = 0 To UBound(objBRarr)  '遍历选择集
        blockATT = objBRarr(i, 0).GetAttributes '获取参照快的属性
        blockATT(8).TextString = UBound(objBRarr) + 1 '修改总页数为
        blockATT(9).TextString = i + 1 '修改页数为i
        If i < 10 Then
            blockATT(13).TextString = "QP-DY-0" & i + 1 '修改图号为QP-DY-0i
        Else
            blockATT(13).TextString = "QP-DY-" & i + 1
        End If

    Next i

End Sub



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
    Set creatSSet = ThisDrawing.SelectionSets.Add(SSetName)
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

    Stop
End Sub
