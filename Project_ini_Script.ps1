##############################   PARSER   ###################################

#Var definition
$ini = Get-Content -Path .\folder_ini.ini 

$section = "none" #section name in ini file
$projectName = "no project name"
[int]$projectNumber = 000
$projectCode = "no project code"
$projectEngineer = "no engineer"
$projectManager = "no manager"
$projectCompany = "no company"
$projectCountry = "no country"

$projectSites = @() # list of project site
$ProjectDocuments = @() # list of hash tables of Project Documents Properties
#$ProjectDocuments = @( @{ DocumentName = ; DocumentQt = ; DocumentPath = ; DocumentType = ; DocumentTemplate = } )

$ProjectItems = @() # list of hash tables of Project items
#ProjectItems =  @( @{ ItemNumber = ; ItemTitle = ; ItemSite = ; ItemFileName =  ; ItemPath = ; ItemType = ; ItemTemplate = })

#Ini File parser
foreach ($line in $ini)
    {
        if ($line -ne "" -and $line.startswith(";") -ne $true )
            {
# Checking section name
                if ($line.StartsWith("[General]") -eq $True) {$section = 'General'; continue}
                if ($line.StartsWith("[Site]") -eq $True) {$section = 'Site'; continue}
                if ($line.StartsWith("[Document]") -eq $True) {$section = 'Document'; continue}
# Section General                
                if ($section -eq 'General')
                {
                $SpliteArray = $Line.Split("=")
                if ($SpliteArray[0].trim('') -eq "ProjectName") {$projectName = $SpliteArray[1].trim('')}
                if ($SpliteArray[0].trim('') -eq "ProjectNumber") {$ProjectNumber = $SpliteArray[1].trim('')}
                if ($SpliteArray[0].trim('') -eq "ProjectCode") {$ProjectCode = $SpliteArray[1].trim('')}
                if ($SpliteArray[0].trim('') -eq "ProjectEngineer") {$projectEngineer = $SpliteArray[1].trim('')}
                if ($SpliteArray[0].trim('') -eq "projectManager") {$projectManager = $SpliteArray[1].trim('')}
                if ($SpliteArray[0].trim('') -eq "projectCompany") {$projectCompany = $SpliteArray[1].trim('')}
                if ($SpliteArray[0].trim('') -eq "projectCountry") {$projectCountry = $SpliteArray[1].trim('')}
                }
# Section ProjectSites
                if ($section -eq 'Site')
                {
                    $ProjectSites += $Line.Trim('')
                }
# Section ProjectDocuments
                if ($section -eq 'Document')
                {
                    $SpliteArray = $Line.Split(",")
                    $ProjectDocuments += @{ DocumentName = $SpliteArray[0].trim(''); DocumentQt = $SpliteArray[1].trim('');
                     DocumentPath = $SpliteArray[2].trim(''); DocumentType = $SpliteArray[3].trim(''); DocumentTemplate = $SpliteArray[4].trim('')}
                }                

            }

    }

#Debug info
    "Project Name is $projectName"
    "Project Code is $ProjectCode"
    "Project Number is $ProjectNumber"    
    "Project Engineer is $ProjectEngineer"
    "Project Manager is $projectManager"
    "Project Company is $projectCompany"
    "Project Country is $projectCountry"
   for ($i = 0; $i -lt $projectSites.Length; $i++) {"Project Site $($i + 1) is $($ProjectSites[$i])"} 
"Number of Sites is $($projectSites.Count)"
# Calculation Qty of Items
$DocumentQty = 0
foreach ($Doc in $ProjectDocuments)
    { 
        if ($Doc.DocumentQt -like "single") { $DocumentQty += 1 }
        if ($Doc.DocumentQt -like "multi") { $DocumentQty += $($ProjectSites.Count) }
    }
"Number of Documents is $DocumentQty"

#$ProjectDocuments

$ItemSmallNumber = 1
foreach ($Doc in $ProjectDocuments)
    {


        if ($Doc.DocumentQt -like "single")
        { 
        [string]$tempNumber = ($projectNumber * 1000) + $ItemSmallNumber
        $tempNumber += "-01"

        [string]$tempFileName = $projectName.Replace(' ', '_') + '_' + $Doc.DocumentName.Replace(' ', '_') + '_' + $tempNumber
        [string]$tempItemType = $Doc.DocumentType

        if ($tempItemType -eq "visio") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.vsdx'}
        if ($tempItemType -eq "word") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.docx'}
        if ($tempItemType -eq "excel") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.xlsx'}

        $ProjectItems += @{ItemNumber = $tempNumber; ItemTitle = $Doc.DocumentName; ItemSite = "General"; ItemFileName = $tempFileName;
            ItemPath = $tempItemPath; ItemType = $tempItemType; ItemTemplate = $Doc.DocumentTemplate}
        $ItemSmallNumber++    
        }
        if ($Doc.DocumentQt -like "multi")
        {
            foreach ($site in $projectSites)
            {
                [string]$tempNumber = ($projectNumber * 1000) + $ItemSmallNumber
                $tempNumber += "-01"
                
                [string]$tempFileName = $Site.Replace(' ', '_') + '_' + $Doc.DocumentName.Replace(' ', '_') + '_' + $tempNumber
                [string]$tempItemType = $Doc.DocumentType
                
                if ($Doc.DocumentType -eq "visio") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.vsdx'}
                if ($Doc.DocumentType -eq "word") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.docx'}
                if ($Doc.DocumentType -eq "excel") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.xlsx'}

                $ProjectItems += @{ItemNumber = $tempNumber; ItemTitle = $Doc.DocumentName; ItemSite = $site; ItemFileName = $tempFileName;
                     ItemPath = $tempItemPath; ItemType = $Doc.DocumentType; ItemTemplate = $Doc.DocumentTemplate}
            $ItemSmallNumber++
            }
        }
        
    }
#Debug info
"Project Items"    
foreach ($item in $ProjectItems) 
    {   
        $item
        write-host "`n"   
    }

###############################    Create Folser Structure      ##################################################

$folderStructure = Get-Content -path .\Folder_Structure.ini
[string]$ProjectPath = (get-location).path + "\" + $projectCountry + "_" + $($projectName.Replace(' ','_'))

'Folder Structer:'
foreach ($line in $folderStructure)
    {
#        if ($line -ne "" -and $line.startswith(";") -ne $true)
        if ($line -match '^\\')
            {
                $lineFormat = $line.Trim("")
#                $lineFormat = $lineFormat.Replace('','_')

                New-Item -path $($ProjectPath + $lineFormat) -ItemType Directory -Force | Out-Null
                $lineFormat
            } 
    }

#############################  Create DoC list ##################################

$Word = New-Object -ComObject Word.application
$Word.visible = $true
$WordDocument = $Word.Documents.add()
$Selection = $Word.Selection
$WordRange = $Selection.range
$WordSection = $WordDocument.Sections.Item(1)

#Landscape orientation
$WordDocument.PageSetup.Orientation = 1

#Header
$Header = $WordSection.Headers.Item(1)
$Header.range.text = $projectCompany

#Footer
$Footer = $WordSection.footers.item(1)
$Footer.range.text = Get-Date -UFormat "%d/%m/%Y"

#Title
$Selection.ParagraphFormat.Alignment = 1
$Selection.Font.Bold = 1
$Selection.Font.Size = 18
$Selection.TypeText("Documentation List")

#Description
$Selection.TypeParagraph()
$Selection.ParagraphFormat.Alignment = 0
$Selection.Font.Bold = 0
$Selection.Font.Size = 11
$Selection.TypeText("Project Name: $projectName")
$Selection.TypeParagraph()
$Selection.TypeText("Project Code: $ProjectCode")
$Selection.TypeParagraph()
$Selection.TypeText("Project Manager: $projectManager")
$Selection.TypeParagraph()
$Selection.TypeText("Project Engineer: $projectEngineer")

#Table name Design Documentation
$Selection.TypeParagraph()
$Selection.ParagraphFormat.Alignment = 0
$Selection.Font.Bold = 1
$Selection.Font.Size = 18
$Selection.TypeText("Design Documentation")
$Selection.TypeParagraph()

#Table
$Selection.Font.Bold = 0
$Selection.Font.Size = 11
$Selection.ParagraphFormat.Alignment = 1
# $selection.Shading.BackgroundPatternColorIndex = 16

$table = $Selection.tables.add($selection.range,$($DocumentQty + 1),5) #Change number of Rows
$table.Borders.InsideLineStyle = 1
$table.Borders.OutsideLineStyle = 1

#First Raw
$table.Rows.Item(1).Shading.ForegroundPatternColorIndex = 16

$table.cell(1,1).range.text = "Doc No"
$table.cell(1,2).range.text = "Document name"
$table.cell(1,3).range.text = "Date"
$table.cell(1,4).range.text = "Creator"
$table.cell(1,5).range.text = "Note"

#Other Rows
$row = 2  
foreach($item in $ProjectItems){
$table.cell($row,1).range.text = $item.ItemNumber
$table.cell($row,2).range.text = $item.ItemFileName
$table.cell($row,3).range.text = Get-Date -UFormat "%d/%m/%Y" 
$table.cell($row,4).range.text = $projectEngineer
$table.cell($row,5).range.text = "Document template was created automaticaly"
$row++
}
[string]$ListDocName = $ProjectPath + "\" + $projectNumber + "-List-Doc"

$WordDocument.saveas($ListDocName)
$WordDocument.Close()
$word.Application.quit()

###################   Copy Docs ########################

#Sorting items by item type
$SortedProjectItems = @()
$SortedProjectItems = $SortedProjectItems + $Projectitems | Sort-Object {$_.Itemtype}
#Index number of first and last office type item

$ExcelIndexFirst = -1
$ExcelIndexLast = -1
$NoneIndexFirst = -1
$NoneIndexLast = -1
$VisioIndexFirst = -1
$VisioIndexLast = -1
$WordIndexFirst = -1
$WordIndexLast = -1

#Checking Index number of first and last office type item
$i = 0
foreach ($item in $SortedProjectItems)
    {   
        If ($item.Itemtype -eq "excel") 
        { 
            if ($ExcelIndexFirst -eq -1) {$ExcelIndexFirst = $i}
            $ExcelIndexLast = $i
        }
        If ($item.Itemtype -eq "none") {
            if ($NoneIndexFirst -eq -1) {$NoneIndexFirst = $i}
            $NoneIndexLast = $i
        }
        If ($item.Itemtype -eq "visio") {
            if ($VisioIndexFirst -eq -1) 
            {$VisioIndexFirst = $i}
            $VisioIndexLast = $i
        }
        If ($item.Itemtype -eq "word") {
            if ($WordIndexFirst -eq -1) {$WordIndexFirst = $i}
            $WordIndexLast = $i
        }
        $i++
    }
# Coping Files   
for ($i = 0; $i -lt $SortedProjectItems.Count; $i++) {

    if ($ExcelIndexLast -ne -1 -and $i -ge $ExcelIndexFirst -and $i -le $ExcelIndexLast )
    {
        if ($i -eq $ExcelIndexFirst) {$OfficeApplication = New-Object -ComObject excel.application}
        $OfficeApplication.Visible = $true       
        [string]$templatename =  (get-location).path + '\Templates\' + $SortedProjectItems[$i].itemtemplate
        
        $OfficeDocument = $OfficeApplication.Workbooks.open($templatename)
        $OfficeDocument.BuiltInDocumentProperties("title") = $SortedProjectItems[$i].ItemTitle
        $OfficeDocument.BuiltInDocumentProperties("comments") = $SortedProjectItems[$i].ItemNumber
        $OfficeDocument.BuiltInDocumentProperties("subject") = $projectName
        $OfficeDocument.BuiltInDocumentProperties("company") = $projectCompany
        $OfficeDocument.BuiltInDocumentProperties("category") = $SortedProjectItems[$i].ItemSite
        $OfficeDocument.BuiltInDocumentProperties("author") = $projectEngineer
        $OfficeDocument.BuiltInDocumentProperties("manager") = $projectManager

        [string]$path =  $ProjectPath + "\" + $SortedProjectItems[$i].ItemPath
        $OfficeDocument.saveas($path) 
        $SortedProjectItems[$i].ItemPath
        $OfficeDocument.Close()
        if ($i -eq $ExcelIndexLast) {$OfficeApplication.Quit()}
    }

    if ($NoneIndexLast -ne -1 -and $i -ge $NoneIndexFirst -and $i -le $NoneIndexLast )
    {
        [string]$templatename =  (get-location).path + '\Templates\' + $SortedProjectItems[$i].itemtemplate
            [string]$path =  $ProjectPath + "\" + $SortedProjectItems[$i].ItemPath
            $SortedProjectItems[$i].ItemPath
            Copy-Item -Path $templatename -Destination $path
    }

    if ($VisioIndexLast -ne -1 -and $i -ge $VisioIndexFirst -and $i -le $VisioIndexLast )
    {
        if ($i -eq $VisioIndexFirst) {$OfficeApplication = New-Object -ComObject Visio.application}
        [string]$templatename =  (get-location).path + '\Templates\' + $SortedProjectItems[$i].itemtemplate
        $OfficeDocument = $OfficeApplication.Documents.open($templatename)

        $OfficeDocument.title = $SortedProjectItems[$i].ItemTitle
        $OfficeDocument.Description = $SortedProjectItems[$i].ItemNumber
        $OfficeDocument.Subject = $projectName
        $OfficeDocument.Company = $projectCompany
        $OfficeDocument.Category = $SortedProjectItems[$i].ItemSite
        $OfficeDocument.Creator = $projectEngineer
        $OfficeDocument.Manager = $projectManager

        [string]$path =  $ProjectPath + "\" + $SortedProjectItems[$i].ItemPath
        $OfficeDocument.saveas($path) | out-null
        $SortedProjectItems[$i].ItemPath
        $OfficeDocument.Close()
        if ($i -eq $VisioIndexLast) {$OfficeApplication.Quit()}
        
    }

    if ($WordIndexLast -ne -1 -and $i -ge $WordIndexFirst -and $i -le $WordIndexLast )
    {
        if ($i -eq $WordIndexFirst) {$OfficeApplication = New-Object -ComObject word.application}
        $OfficeApplication.Visible = $true
        [string]$templatename =  (get-location).path + '\Templates\' + $SortedProjectItems[$i].itemtemplate

        $OfficeDocument = $OfficeApplication.Documents.open($templatename)
        $OfficeDocument.BuiltInDocumentProperties("title") = $SortedProjectItems[$i].ItemTitle
        $OfficeDocument.BuiltInDocumentProperties("comments") = $SortedProjectItems[$i].ItemNumber
        $OfficeDocument.BuiltInDocumentProperties("subject") = $projectName
        $OfficeDocument.BuiltInDocumentProperties("company") = $projectCompany
        $OfficeDocument.BuiltInDocumentProperties("category") = $SortedProjectItems[$i].ItemSite
        $OfficeDocument.BuiltInDocumentProperties("author") = $projectEngineer
        $OfficeDocument.BuiltInDocumentProperties("manager") = $projectManager

        [string]$path =  $ProjectPath + "\" + $SortedProjectItems[$i].ItemPath
        $OfficeDocument.saveas($path)
        $SortedProjectItems[$i].ItemPath
        $OfficeDocument.Close()
        if ($i -eq $VisioIndexLast) {$OfficeApplication.Quit()}
    }
        
}    