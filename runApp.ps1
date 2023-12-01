#this script will let you create multiple xml data for tag inside xml file that has multiple same tags containing different data 
$a = New-Object -comobject Excel.Application #kreiranje objekta za poziv aplikacije excel
$a.Visible = $false #sakrij otvoreni objekt excel aplikaciju
$a.DisplayAlerts = $False #blokiraj alert poruke


# Variables that contain data bookmarks that will be replaced with data from excel file colums-rows / add as many as you need
$fstTg = 'insertFirstTag';
$scnTg = 'insertSecondTag';
$thrdTg = 'insertThirdTag';
#.....

# 
$Workbook = $a.workbooks.open('PATH TO YOUR *XLSX FILE') 
$Sheet = $Workbook.Worksheets.Item(1)

$row = [int]2 #num of starting row in excel
# lists for console output
$ListA = @() #list of elements for column A
$ListB = @() #list of elements for column B
$ListC = @() #list of elements for column C
# ..... add as many columns you need for your xml file / each column contains data for each individual tag in xml

# files used in script
$header_tag = "PATH TO headerTAG.txt" # header is file containing top level info of xml file
$header_file = "PATH TO header.txt" 
$footer_file = "PATH TO footer.txt"# footer is a file containing bottom xml data
$original_file = 'PATH TO originalTAG.txt'
$destination_file =  'PATH TO bridge.txt' 
$adapterDatot = 'PATH TO adapterDatot.txt' #
$mergedXML = "PATH TO mergedXML.txt" #final file txt with data
$xmlOut = "PATH TO XMLFinal.xml" # final xml file with data

#copy header_tag into header_file / this is used so bookmarks inside headerTAG file don't get overwritten
[System.IO.File]::Copy($header_tag, $header_file, $true);

#copy header_file data to adapterDatot file - adapterDatot file / any data inside adapterDatot file if exists before copypaste will be overwritten
[System.IO.File]::Copy($header_file, $adapterDatot, $true);

Do {
	# variable that contains data from each individual row in excel
	$rowA = $Sheet.Cells.Item($row,1).Text # number 1 is equivalent for column A in excel
	$rowB = $Sheet.Cells.Item($row,2).Text  # number 2 is equivalent for column B in excel
	$rowC = $Sheet.Cells.Item($row,3).Text  # number 3 is equivalent for column C in excel
	
	
	#$variableX = $variableY.Replace(',', '.') #replace variableY number that is separeted with , to .
	
	[System.IO.File]::Copy($original_file, $destination_file, $true);
	
# appending data to list
	# append columnX, rowY data to list
	$ListA += $rowA	
	$ListB += $rowB
	$ListC += $rowC
	

	$row = $row + [int]1 #counter
	
	# replace original xml with tags with actual data from excel columns
	(Get-Content -Path $destination_file) | ForEach-Object {$_ -Replace $fstTg, $rowA} | Set-Content -Path $destination_file 
	(Get-Content -Path $destination_file) | ForEach-Object {$_ -Replace $scnTg, $rowB} | Set-Content -Path $destination_file  
	(Get-Content -Path $destination_file) | ForEach-Object {$_ -Replace $thrdTg, $rowC} | Set-Content -Path $destination_file
	
	# regex to prevent empty lines when copy pasteing data from one file to another
	((Get-Content $destination_file -Raw) -replace "(?m)^\s*`r`n",'').trim() | Add-Content -Path $adapterDatot
	
	} until (!$Sheet.Cells.Item($row, 1).Text) #Punjenje liste dok je kolona sa najviše podataka napunjena / zadnji red sa podacima 
# closing excel file when all is finished
$Workbook.Close()
$a.Quit()
 [System.Runtime.Interopservices.Marshal]::ReleaseComObject($a) | Out-Null # do not write 0-success output in PS 
	
#append footer content to a file / data appended after insert of all data for specific tag with multiple data lines
Get-Content -Path $footer_file |
Add-Content -Path $adapterDatot

#append all data into a final txt file
Get-Content -Path $adapterDatot |
Set-Content -Path $mergedXML

# erase all unnecessarely data from adapter files 
Clear-Content -Path $adapterDatot
Clear-Content -Path $destination_file
# convert txt to xml file
Get-Content $mergedXML | Set-Content "PATH TO XMLFinal.xml"
# write output in PS console	
Write-Output $ListA, $ListB, $ListC, $ListD, $ListF, $ListG, $ListH #Ispis svih elemenata liste
