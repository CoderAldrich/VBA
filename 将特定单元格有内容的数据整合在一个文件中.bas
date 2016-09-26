Sub mergerselection() '把多个excel文件每个工作表制定单元格不为空的条目汇总到一个excel文件中
	On Error Resume Next
	Dim x As Variant, x1 As Variant, input_workbook As Workbook, input_worksheet As Worksheet, ws As Worksheet
	Dim current_workbook As Workbook, current_worksheet As Worksheet, i As Integer, column As Integer, row As Long, j As Integer, mark As Integer
	Dim next_worksheet As Worksheet
	Dim temp_worksheet As Worksheet
	Dim col_total As Integer
	Dim col_count As Integer
	Dim first_idx As Integer
	Dim next_idx As Integer	
	Dim row_to_check As Integer
	Dim row_to_check_modify As Integer
	Dim row_total As Integer
	Dim row_count As Integer
	Dim first_pos As Integer
	Dim next_pos As Integer
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
		
	x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", Title:="选择需要合并的Excel", MultiSelect:=True)
	Set current_workbook = ThisWorkbook
	
	For Each x1 In x														'所选的多个excel文件的循环访问
		If x1 <> False Then
			
			Set input_workbook = Workbooks.Open(x1)						'打开工作簿
			
			For i = 1 To input_workbook.Sheets.Count					'单个excel中工作表的循环访问
				mark = 0
				Set input_worksheet = input_workbook.Sheets(i)
				'源文件表的名字在本文件中尚不存在就新建一个
				For j = 1 To current_workbook.Sheets.Count
					Set current_worksheet = current_workbook.Sheets(j)
					If Trim(current_worksheet.Name) = Trim(input_worksheet.Name) Then		'该名称的工作表已存在并且第一个单元不为空则复制源数据，否则创建工作表并复制前两行标题
						mark = 1
						Exit For
		 			End If
				Next
				
				If mark = 0 Then			'新建一个工作表并重命名
					current_workbook.Sheets.Add After:=current_workbook.Sheets(current_workbook.Sheets.Count)
		 			current_worksheet.Name = Trim(input_worksheet.Name)
				End If					

				Set input_worksheet = input_workbook.Sheets(i)
				row_total = input_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row				'获取本工作表总行数
				col_total = input_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
			
				'找到“说明原因”所在的列
 		    For col_count = 1 To col_total
	        If input_worksheet.Cells(2, col_count) = "说明原因" Then
		        row_to_check = col_count
 	         Exit For
 		      End If
	      Next		
      
				'找到“是否修改”所在的列
	      For col_count = 1 To col_total
	        If input_worksheet.Cells(2, col_count) = "是否修改" Then
		        row_to_check_modify = col_count
	          Exit For
	        End If
	      Next      	
      
 		    '如果结果工作表还是空的，先把前两行标题复制进去
	      column = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
				row = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
				If column = 1 And row = 1 And current_worksheet.Cells(1, 1) = "" Then
					input_worksheet.Rows("1:2").Copy current_worksheet.Cells(1, 1)
				End If
			
				'从第三行开始判断第row_to_check列或row_to_check_modify列是否为空，不为空则复制到结果文件对应的工作表中
				For row_count = 3 To row_total
					If Len(Trim(input_worksheet.Cells(row_count, row_to_check))) <> 0 Or Len(Trim(input_worksheet.Cells(row_count, row_to_check_modify))) <> 0 Then							'如果该行第row_to_check列不为空则复制本行
				    column = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
						row = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
						input_worksheet.Rows(row_count).Copy current_worksheet.Cells(row+1, 1)
					End If
				Next
			Next
			input_workbook.Close
		End If
	Next
	
	'删除空工作表
	For Each ws In ActiveWorkbook.Worksheets    
		ws.Activate    
		If ActiveWorkbook.Worksheets.Count > 1 Then      
			If IsEmpty(ActiveSheet.UsedRange) Then '如工作表为空
				ws.Delete    '则删除该表      
			End If     
		End If   
	Next	
	
	'用冒泡排序法对工作表排序
	'不可以直接比较字符串，因为前面的序号有两位数的
	For i = 1 To current_workbook.Sheets.Count
		For j = 1 To current_workbook.Sheets.Count-i
			'找到两个工作表名字中、的位置
			first_pos = Instr(Trim(current_workbook.Sheets(j).Name), "、")
			next_pos = Instr(Trim(current_workbook.Sheets(j+1).Name), "、")
			'取出、前的字符串并转换成数字
			first_idx = CInt(Left(Trim(current_workbook.Sheets(j).Name), first_pos-1))
			next_idx = CInt(Left(Trim(current_workbook.Sheets(j+1).Name), next_pos-1))
			If first_idx > next_idx Then
				current_workbook.Sheets(j).Move after:= current_workbook.Sheets(j+1)
			End If
		Next
	Next
	
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub