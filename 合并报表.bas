Sub merger() '将多个文件的多个工作簿下的工作表依次对应合并到本工作簿下的工作表，即第一张工作表对应合并到第一张，第二张对应合并到第二张……
	On Error Resume Next
	Dim x As Variant, x1 As Variant, input_workbook As Workbook, input_worksheet As Worksheet, ws As Worksheet
	Dim current_workbook As Workbook, current_worksheet As Worksheet, i As Integer, column As Integer, row As Long, j As Integer, mark As Integer
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
		
	x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", Title:="选择需要合并的Excel", MultiSelect:=True)
	Set current_workbook = ThisWorkbook
	
	For Each x1 In x														'所选的多个excel文件的循环访问
		If x1 <> False Then
			
		Set input_workbook = Workbooks.Open(x1)						'打开工作簿
		
		For i = 1 To input_workbook.Sheets.Count					'单个excel中工作表的循环访问
			mark = 0
			Set current_worksheet = current_workbook.Sheets(i)
			Set input_worksheet = input_workbook.Sheets(i)
			
			For j = 1 To current_workbook.Sheets.Count
				Set current_worksheet = current_workbook.Sheets(j)
				If current_worksheet.Name = input_worksheet.Name Then
					column = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
					row = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
					If column = 1 And row = 1 And current_worksheet.Cells(1, 1) = "" Then			
						input_worksheet.UsedRange.Copy current_worksheet.Cells(1, 1)				'如果当前是总表的第一行第一列，则直接填充
					Else
				 		input_worksheet.UsedRange.Offset(2, 0).Copy current_worksheet.Cells(row + 1, 1)	'并不是总表第一行第一列，则从下一行的第一列开始填充
				 	End If
				 	mark = 1
	 			End If
			Next
			
			If mark = 0 Then			'输入文件的当前工作表还不存在，需要在结果文件中新建一个工作表
				current_workbook.Sheets.Add After:=current_workbook.Sheets(current_workbook.Sheets.Count)
				column = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
				row = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
				If column = 1 And row = 1 And current_worksheet.Cells(1, 1) = "" Then			
					input_worksheet.UsedRange.Copy current_worksheet.Cells(1, 1)				'如果当前是总表的第一行第一列，则直接填充
				Else
			 		input_worksheet.UsedRange.Offset(2, 0).Copy current_worksheet.Cells(row + 1, 1)	'不是总表第一行第一列，则从下一行的第一列开始填充
			 	End If
	 			current_worksheet.Name = input_worksheet.Name
			End If
		Next
		input_workbook.Close
		End If
	Next
	
	'删除空工作表
	For Each ws In ActiveWorkbook.Worksheets    
		ws.Activate    
		If ActiveWorkbook.Worksheets.Count > 1 Then      
			If IsEmpty(ActiveSheet.UsedRange) Then '如表格为空          
				ws.Delete    '则删除该表      
			End If     
		End If   
	Next	
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub