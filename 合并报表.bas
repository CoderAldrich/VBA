Sub merger() '������ļ��Ķ���������µĹ��������ζ�Ӧ�ϲ������������µĹ���������һ�Ź������Ӧ�ϲ�����һ�ţ��ڶ��Ŷ�Ӧ�ϲ����ڶ��š���
	On Error Resume Next
	Dim x As Variant, x1 As Variant, input_workbook As Workbook, input_worksheet As Worksheet, ws As Worksheet
	Dim current_workbook As Workbook, current_worksheet As Worksheet, i As Integer, column As Integer, row As Long, j As Integer, mark As Integer
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
		
	x = Application.GetOpenFilename(FileFilter:="Excel�ļ� (*.xls; *.xlsx),*.xls; *.xlsx,�����ļ�(*.*),*.*", Title:="ѡ����Ҫ�ϲ���Excel", MultiSelect:=True)
	Set current_workbook = ThisWorkbook
	
	For Each x1 In x														'��ѡ�Ķ��excel�ļ���ѭ������
		If x1 <> False Then
			
		Set input_workbook = Workbooks.Open(x1)						'�򿪹�����
		
		For i = 1 To input_workbook.Sheets.Count					'����excel�й������ѭ������
			mark = 0
			Set current_worksheet = current_workbook.Sheets(i)
			Set input_worksheet = input_workbook.Sheets(i)
			
			For j = 1 To current_workbook.Sheets.Count
				Set current_worksheet = current_workbook.Sheets(j)
				If current_worksheet.Name = input_worksheet.Name Then
					column = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
					row = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
					If column = 1 And row = 1 And current_worksheet.Cells(1, 1) = "" Then			
						input_worksheet.UsedRange.Copy current_worksheet.Cells(1, 1)				'�����ǰ���ܱ�ĵ�һ�е�һ�У���ֱ�����
					Else
				 		input_worksheet.UsedRange.Offset(2, 0).Copy current_worksheet.Cells(row + 1, 1)	'�������ܱ��һ�е�һ�У������һ�еĵ�һ�п�ʼ���
				 	End If
				 	mark = 1
	 			End If
			Next
			
			If mark = 0 Then			'�����ļ��ĵ�ǰ�����������ڣ���Ҫ�ڽ���ļ����½�һ��������
				current_workbook.Sheets.Add After:=current_workbook.Sheets(current_workbook.Sheets.Count)
				column = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
				row = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
				If column = 1 And row = 1 And current_worksheet.Cells(1, 1) = "" Then			
					input_worksheet.UsedRange.Copy current_worksheet.Cells(1, 1)				'�����ǰ���ܱ�ĵ�һ�е�һ�У���ֱ�����
				Else
			 		input_worksheet.UsedRange.Offset(2, 0).Copy current_worksheet.Cells(row + 1, 1)	'�����ܱ��һ�е�һ�У������һ�еĵ�һ�п�ʼ���
			 	End If
	 			current_worksheet.Name = input_worksheet.Name
			End If
		Next
		input_workbook.Close
		End If
	Next
	
	'ɾ���չ�����
	For Each ws In ActiveWorkbook.Worksheets    
		ws.Activate    
		If ActiveWorkbook.Worksheets.Count > 1 Then      
			If IsEmpty(ActiveSheet.UsedRange) Then '����Ϊ��          
				ws.Delete    '��ɾ���ñ�      
			End If     
		End If   
	Next	
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub