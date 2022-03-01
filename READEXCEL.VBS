Option Explicit
REM We use "Option Explicit" to help us check for coding mistakes

Dim MyConn
Dim MdbFilePath
Dim var
Dim newoutput
Dim newname
Dim del_query
Dim ins_query
REM the Excel Application
Dim objExcel
REM the path to the excel file
Dim excelPath
REM how many worksheets are in the current excel file
Dim worksheetCount
Dim counter
REM the worksheet we are currently getting data from
Dim currentWorkSheet
REM the number of columns in the current worksheet that have data in them
Dim usedColumnsCount
REM the number of rows in the current worksheet that have data in them
Dim usedRowsCount
Dim row
Dim column
REM the topmost row in the current worksheet that has data in it
Dim top
REM the leftmost row in the current worksheet that has data in it
Dim left
Dim Cells
REM the current row and column of the current worksheet we are reading
Dim curCol
Dim curRow
REM the value of the current row and column of the current worksheet we are reading
Dim word
Dim test
Dim base
Dim run1
Dim run2
Dim combined
Dim race
Dim racefile
 
dim mountain
dim racedate
dim racetype
dim connStr 

REM excel record row
dim ussa
dim bib
dim name
dim cat
dim gender
dim club
dim yob
dim runtime
dim runplace
dim timegap

Dim myFSO, outputline, outputfilename, WriteStuff 
mountain = WScript.Arguments.Item(0)
racedate = WScript.Arguments.Item(1)
racetype = WScript.Arguments.Item(2)

race = mountain & "-" & racedate 
racefile = mountain & "-" & racedate & "-" & racetype
outputfilename = race & "-out.txt"
Set myFSO = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = myFSO.OpenTextFile(outputfilename, 2, True)

Set MyConn = CreateObject("ADODB.Connection")
MdbFilePath = "raceresults.mdb"
connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & MdbFilePath 
wscript.echo connStr
MyConn.Open connStr 


del_query = "DELETE * FROM [Race-import] WHERE ((race = " & chr(39) & mountain & "-" & racedate & chr(39) & ")  );"
wscript.echo del_query
MyConn.Execute del_query


excelPath = "C:\Users\ekikel\OneDrive - Waterville Valley Academy\NJSRA\2015-2016\" & racefile & ".csv"
WScript.Echo "Reading Data from " & excelPath


REM Create an invisible version of Excel
Set objExcel = CreateObject("Excel.Application")

REM don't display any messages about documents needing to be converted
REM from  old Excel file formats
objExcel.DisplayAlerts = 0

REM open the excel document as read-only
REM open (path, confirmconversions, readonly)
objExcel.Workbooks.open excelPath, false, true


REM How many worksheets are in this Excel documents
workSheetCount = objExcel.Worksheets.Count

WScript.Echo "We have " & workSheetCount & " worksheets"

REM Loop through each worksheet
For counter = 1 to workSheetCount
rem	WScript.Echo "-----------------------------------------------"
rem	WScript.Echo "Reading data from worksheet " & counter & vbCRLF

	Set currentWorkSheet = objExcel.ActiveWorkbook.Worksheets(counter)
	REM how many columns are used in the current worksheet
	usedColumnsCount = currentWorkSheet.UsedRange.Columns.Count
	REM how many rows are used in the current worksheet
	usedRowsCount = currentWorkSheet.UsedRange.Rows.Count

	REM What is the topmost row in the spreadsheet that has data in it
	top = currentWorksheet.UsedRange.Row
	REM What is the leftmost column in the spreadsheet that has data in it
	left = currentWorksheet.UsedRange.Column


	Set Cells = currentWorksheet.Cells
	REM Loop through each row in the worksheet 
	For row = 0 to (usedRowsCount-1)
		REM Loop through each column in the worksheet 
		For column = 0 to usedColumnsCount-1
			REM only look at rows that are in the "used" range
			curRow = row+top
			REM only look at columns that are in the "used" range
			curCol = column+left

			if curCol = 1 Then 
				ussa = Cells(curRow,curCol).Value 
			elseif curCol = 2 Then 
				bib = Cells(curRow,curCol).Value 
			elseif curCol = 3 Then 
				name = Cells(curRow,curCol).Value 
			elseif curCol = 4 Then 
				cat = Cells(curRow,curCol).Value 
			elseif curCol = 5 Then 
				gender = Cells(curRow,curCol).Value 
			elseif curCol = 6 Then 
				club = Cells(curRow,curCol).Value 
			elseif curCol = 7 Then 
				yob = Cells(curRow,curCol).Value 
			elseif curCol = 12 Then 
				runtime = Cells(curRow,curCol).Value 
			elseif curCol = 13 Then 
				runplace = Cells(curRow,curCol).Value * -1
			elseif curCol = 14 Then 
				timegap = Cells(curRow,curCol).Value 
			End If
		Next
		base = chr(34) & race & chr(34) & chr(44)
		base = base & chr(34) & racetype & chr(34) & chr(44)
		base = base & chr(34) & gender & chr(34) & chr(44)
		base = base & chr(34) & cat & chr(34) & chr(44)
		base = base & chr(34) & ussa & chr(34) & chr(44)
		base = base & chr(34) & bib & chr(34) & chr(44)
		base = base & chr(34) & name & chr(34) & chr(44)
		base = base & chr(34) & cat & chr(34) & chr(44)
		base = base & chr(34) & club & chr(34) & chr(44)
		base = base & chr(34) & yob & chr(34) & chr(44)
		base = base & chr(34) & "COMBINED" & chr(34) & chr(44)
		base = base & chr(34) & runtime & chr(34) & chr(44)
		base = base & chr(34) & runplace & chr(34) & chr(44)
		base = base & chr(34) & timegap & chr(34) & chr(44)

		WScript.Echo (base)
		outputline = base 
         	call writescore
	Next

	REM We are done with the current worksheet, release the memory
	Set currentWorkSheet = Nothing
Next

objExcel.Workbooks(1).Close
objExcel.Quit

Set currentWorkSheet = Nothing
REM We are done with the Excel object, release it from memory
Set objExcel = Nothing

WriteStuff.Close
SET WriteStuff = NOTHING
SET myFSO = NOTHING

MyConn.close
set MyConn = nothing


sub writescore
	
	newoutput = replace(outputline, """" ,"")
	var = Split(newoutput , ",")
	newname = replace(var(6), "'" ,"''")


	ins_query = "INSERT INTO [Race-import](race,racetype,Gender,Cat,ussa,bib,name,class,club,yob,run,runtme,runplace,racepoints,runadjplace) "
	ins_query = ins_query & "values (" 
	ins_query = ins_query & chr(39) & var(0) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(1) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(2) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(3) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(4) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(5) & chr(39)& "," 
	ins_query = ins_query & chr(39) & newname & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(7) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(8) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(9) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(10) & chr(39)& "," 
	ins_query = ins_query & chr(39) & var(11) & chr(39)& "," 
	if (var(12) = "") Then
		ins_query = ins_query & chr(39) & "99999" & chr(39) & "," 
	Else
		ins_query = ins_query & chr(39) & var(12) & chr(39)& "," 
	End If
	if (var(13) = "") Then
		ins_query = ins_query & chr(39) & "999" & chr(39) & "," 
	Else
		ins_query = ins_query & chr(39) & var(13) & chr(39)& "," 
	End If
	ins_query = ins_query & chr(39) & "99999" & chr(39)& ")" 

  	wscript.echo ins_query

   	MyConn.Execute ins_query
end sub









