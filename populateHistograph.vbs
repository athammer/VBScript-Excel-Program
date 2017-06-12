'============     STARTING EXCEL AND CREATING GRAPH    ============
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
'new sheet
Set objWriteSheet = objWorkbook.Worksheets(1)
'new graph
Set histogram = objWriteSheet.ChartObjects.Add(50, 50, 1000, 500).Chart  

histogram.HasTitle = True 
histogram.AutoScaling = True
histogram.Axes(2).HasTitle = True
histogram.Axes(1).HasTitle = True

'============     USER EDITABLE VARS (EDIT TO YOUR CHOOSING)     ============

'Title of graph 
histogram.ChartTitle.Text = ""

'X-axis title goes here 
histogram.Axes(1).AxisTitle.Text = "" 

'Y-axis title goes here
histogram.Axes(2).AxisTitle.Text = ""
 
'Change type of chart and style (defualt is 51)
'To change visit the website below and scroll down till you see CHART TYPE ENUMERATION.
'Then find the graph you want and get the value and replace histogram.ChartType current value with your own
'http://billietconsulting.com/2013/11/creating-charts-with-vba-in-excel/
histogram.ChartType = 51 'insert number here
histogram.ChartStyle = 8 'figure this out later (does not work)
 


'============     GETTING USER INPUT AND PRINTING TO EXCEL     ============
strInput = InputBox( "How many data sets do you have? Max is 26 data sets until it can be fixed.", "User Input")
ReDim onArr(strInput)
ReDim offArr(strInput + 2)
ReDim offBadArr(strInput + 2)

'arrCount = 0 'not used for now as we need to skip an index in the array to show that there is no data there
count = 0
For i = 1 To strInput
	objExcel.cells(1, i).value = i
	input1 = InputBox( "What is the first (first being on) data point in set " & i & "? Enter numerical numbers only.", "User Input")
	onArr(i) = input1 
	objExcel.cells(2, i).value = input1
	
	input2 = InputBox( "What is the second (first being off) data point in set " & i & "? Enter numerical numbers only.", "User Input")
	
	If(input2 < 5) Then 'good
		offArr(i) = input2
		objExcel.cells(3, i).value = input2
		objExcel.cells(4, i).value = "" 'to create space
	Else 'bad
		offBadArr(i) = input2
		objExcel.cells(3, i).value = "" 
		objExcel.cells(4, i).value = input2
	End If
	count = i
Next

'============     CREATING NEW SERIES     ============
With histogram.SeriesCollection.NewSeries
        .Name = "On"
        .Values = objWriteSheet.Range("A2:" & Chr(65 + count) & 2)
        .XValues = objWriteSheet.Range("A1:" & Chr(65 + count) & 1)
		.Interior.Color = RGB(34,139,34)
End With
With histogram.SeriesCollection.NewSeries
        .Name = "Off"
        .Values = objWriteSheet.Range("A3:" & Chr(65 + count) & 3)
        .XValues = objWriteSheet.Range("A1:" & Chr(65 + count) & 1)
		.Interior.Color = RGB(70,130,180)
End With
With histogram.SeriesCollection.NewSeries
        .Name = "Off < 5"
        .Values = objWriteSheet.Range("A4:" & Chr(65 + count) & 4)
        .XValues = objWriteSheet.Range("A1:" & Chr(65 + count) & 1)
		.Interior.Color = RGB(128,0,0)
End With

'============     CLEAN UP     ============
'For i = 1 To 4
'	For j = 1 To strInput
'		objExcel.cells(i, j).value = ""
'	Next
'Next

'============     ERROR CATCHING     ============
'On Error Resume Next
'MsgBox "Possible error has been detected this may be caused by canceling an input. If error persists contact the creator of this script. Error # " & CStr(Err.Number) & " " & Err.Description, 16, "Error found"
'Err.Clear   ' Clear the error.
