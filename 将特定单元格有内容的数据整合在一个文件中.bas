Sub mergerselection() '�Ѷ��excel�ļ�ÿ���������ƶ���Ԫ��Ϊ�յ���Ŀ���ܵ�һ��excel�ļ���
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
		
	x = Application.GetOpenFilename(FileFilter:="Excel�ļ� (*.xls; *.xlsx),*.xls; *.xlsx,�����ļ�(*.*),*.*", Title:="ѡ����Ҫ�ϲ���Excel", MultiSelect:=True)
	Set current_workbook = ThisWorkbook
	
	For Each x1 In x														'��ѡ�Ķ��excel�ļ���ѭ������
		If x1 <> False Then
			
			Set input_workbook = Workbooks.Open(x1)						'�򿪹�����
			
			For i = 1 To input_workbook.Sheets.Count					'����excel�й������ѭ������
				mark = 0
				Set input_worksheet = input_workbook.Sheets(i)
				'Դ�ļ���������ڱ��ļ����в����ھ��½�һ��
				For j = 1 To current_workbook.Sheets.Count
					Set current_worksheet = current_workbook.Sheets(j)
					If Trim(current_worksheet.Name) = Trim(input_worksheet.Name) Then		'�����ƵĹ������Ѵ��ڲ��ҵ�һ����Ԫ��Ϊ������Դ���ݣ����򴴽�����������ǰ���б���
						mark = 1
						Exit For
		 			End If
				Next
				
				If mark = 0 Then			'�½�һ��������������
					current_workbook.Sheets.Add After:=current_workbook.Sheets(current_workbook.Sheets.Count)
		 			current_worksheet.Name = Trim(input_worksheet.Name)
				End If					

				Set input_worksheet = input_workbook.Sheets(i)
				row_total = input_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row				'��ȡ��������������
				col_total = input_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
			
				'�ҵ���˵��ԭ�����ڵ���
 		    For col_count = 1 To col_total
	        If input_worksheet.Cells(2, col_count) = "˵��ԭ��" Then
		        row_to_check = col_count
 	         Exit For
 		      End If
	      Next		
      
				'�ҵ����Ƿ��޸ġ����ڵ���
	      For col_count = 1 To col_total
	        If input_worksheet.Cells(2, col_count) = "�Ƿ��޸�" Then
		        row_to_check_modify = col_count
	          Exit For
	        End If
	      Next      	
      
 		    '�������������ǿյģ��Ȱ�ǰ���б��⸴�ƽ�ȥ
	      column = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
				row = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
				If column = 1 And row = 1 And current_worksheet.Cells(1, 1) = "" Then
					input_worksheet.Rows("1:2").Copy current_worksheet.Cells(1, 1)
				End If
			
				'�ӵ����п�ʼ�жϵ�row_to_check�л�row_to_check_modify���Ƿ�Ϊ�գ���Ϊ�����Ƶ�����ļ���Ӧ�Ĺ�������
				For row_count = 3 To row_total
					If Len(Trim(input_worksheet.Cells(row_count, row_to_check))) <> 0 Or Len(Trim(input_worksheet.Cells(row_count, row_to_check_modify))) <> 0 Then							'������е�row_to_check�в�Ϊ�����Ʊ���
				    column = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
						row = current_worksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
						input_worksheet.Rows(row_count).Copy current_worksheet.Cells(row+1, 1)
					End If
				Next
			Next
			input_workbook.Close
		End If
	Next
	
	'ɾ���չ�����
	For Each ws In ActiveWorkbook.Worksheets    
		ws.Activate    
		If ActiveWorkbook.Worksheets.Count > 1 Then      
			If IsEmpty(ActiveSheet.UsedRange) Then '�繤����Ϊ��
				ws.Delete    '��ɾ���ñ�      
			End If     
		End If   
	Next	
	
	'��ð�����򷨶Թ���������
	'������ֱ�ӱȽ��ַ�������Ϊǰ����������λ����
	For i = 1 To current_workbook.Sheets.Count
		For j = 1 To current_workbook.Sheets.Count-i
			'�ҵ����������������С���λ��
			first_pos = Instr(Trim(current_workbook.Sheets(j).Name), "��")
			next_pos = Instr(Trim(current_workbook.Sheets(j+1).Name), "��")
			'ȡ����ǰ���ַ�����ת��������
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