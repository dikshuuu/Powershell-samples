#Script to pass Rakuten CSV Report file to Excel
 
#$INfile = Read-Host "Put the path of the CSV Report file you want to transform here: "  
#$OUTfile = Read-Host "Put the path of the XLSX Report file to be saved: "
 
#Or you can hardcode them
$INfile="path here"           #Enter the path of the csv file here"
$OUTfile="path here" #Enter the path of the excel file here"
$delimiter= ","
 
 
$excel = New-Object -ComObject excel.application 
$excel.visible = $false
$workbook = $excel.Workbooks.Add()
$worksheet= $workbook.Worksheets.Item(1) 
 
# Build the QueryTables.Add command and reformat the data(the cells value might be different for you so please tweak them to your pleasure)
$TxtConnector = ("TEXT;" + $INfile)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A7"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1
$worksheet.Cells.Item(7,1).EntireRow.Interior.ColorIndex = 4 #makes row coloured
$worksheet.Cells.Item(7,1).EntireRow.Font.Bold = $true       #makes row bold
$query.Refresh()
$query.Delete()
 
$start_date = $worksheet.Cells.Item(8, 1).Value().toString("dd-MM-yyyy")   #take date value from one cell and transform it into string to be added in the header)
$end_date = $worksheet.Cells.Item(8, 2).Value().toString("dd-MM-yyyy")
$worksheet.Name = "rakuten_usage_$end_date"
$row = 1 
$Column = 1 
$worksheet.Cells.Item($row,$column)= "
 
Reporting Period                           $start_date - $end_date
MO Project Manager                         manager_name
Infrastructure                             infrastructure_name
                                           Project_name"
                                                                                                              
 
#Styling the header
$excel.Rows.Item("8:20").Select()
$excel.ActiveWindow.FreezePanes = $true   #to freeze the header when you scroll down
$MergeCells = $worksheet.Range("A1:L6") 
$worksheet.Cells.Item(1,1).Font.Size = 12
$worksheet.Cells.Item(1,1).Font.Bold=$True
$worksheet.Cells.Item(1,1).Font.Name = "Helvetica Neue" 
$MergeCells.Select() 
$MergeCells.MergeCells = $true 
$worksheet.Cells(1, 1).HorizontalAlignment = -4131 #align the header, choose the font size, font name, and bold)
$item=$worksheet.Range("A8:L41")
$item.Borders.LineStyle = -4119
 
$worksheet.SaveAs($OUTfile) 
$excel.Quit()
