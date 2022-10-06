Start-Transcript -Path log.log -IncludeInvocationHeader

######    FUNCTIONS    #########################################
function read-ProjectData
{
[CmdletBinding()]
param (
    [string]$inifile = ".\folder_ini.ini" 
)
$ini = Get-Content -Path $inifile
$ProjectData = @{projectName = "no project name"; projectNumber = 0; projectCode = "no project code"; projectEngineer = "no engineer";
projectManager =  "no manager"; ProjectCompany = "no company"; projectCountry = "no country"} # hashtable of project data

$section = "none" #section name in ini file
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
            if ($SpliteArray[0].trim('') -eq "ProjectName") {$ProjectData.ProjectName = $SpliteArray[1].trim('')}
            if ($SpliteArray[0].trim('') -eq "ProjectNumber") {[int]$ProjectData.ProjectNumber = $SpliteArray[1].trim('')}
            if ($SpliteArray[0].trim('') -eq "ProjectCode") {$ProjectData.ProjectCode = $SpliteArray[1].trim('')}
            if ($SpliteArray[0].trim('') -eq "ProjectEngineer") {$ProjectData.ProjectEngineer = $SpliteArray[1].trim('')}
            if ($SpliteArray[0].trim('') -eq "projectManager") {$ProjectData.ProjectManager = $SpliteArray[1].trim('')}
            if ($SpliteArray[0].trim('') -eq "projectCompany") {$ProjectData.ProjectCompany = $SpliteArray[1].trim('')}
            if ($SpliteArray[0].trim('') -eq "projectCountry") {$ProjectData.ProjectCountry = $SpliteArray[1].trim('')}
            }
        }
    }
#Debug info
Write-Host "Project data" -BackgroundColor Green
Write-Host "Project Name is $($ProjectData.projectName)"
Write-Host "Project Code is $($ProjectData.ProjectCode)"
Write-Host "Project Number is $($ProjectData.ProjectNumber)"    
Write-Host "Project Engineer is $($ProjectData.ProjectEngineer)"
Write-Host "Project Manager is $($ProjectData.projectManager)"
Write-Host "Project Company is $($ProjectData.projectCompany)"
Write-Host "Project Country is $($ProjectData.projectCountry)"
Write-Host "`n"
#RETURN        
$ProjectData
}
function read-ProjectSites
{
    [CmdletBinding()]
    param (
        [string]$inifile = ".\folder_ini.ini" 
    )
    $ini = Get-Content -Path $inifile
    $projectSites = @() # list of project site
    
    $section = "none" #section name in ini file
    #Ini File parser
    foreach ($line in $ini)
        {
            if ($line -ne "" -and $line.startswith(";") -ne $true )
                {
    # Checking section name
                    if ($line.StartsWith("[General]") -eq $True) {$section = 'General'; continue}
                    if ($line.StartsWith("[Site]") -eq $True) {$section = 'Site'; continue}
                    if ($line.StartsWith("[Document]") -eq $True) {$section = 'Document'; continue}
    
    # Section ProjectSites
                    if ($section -eq 'Site')
                    {
                        $ProjectSites += $Line.Trim('')
                    }
                

                }
        }
#Debug info
    Write-Host "Project sites" -BackgroundColor Green
    for ($i = 0; $i -lt $projectSites.Length; $i++) {Write-Host "Project Site $($i + 1) is $($ProjectSites[$i])"} 
    Write-Host "Number of Sites is $($projectSites.Count)"
    Write-Host "`n"
    # Calculation Qty of Items
#    $DocumentQty = 0
#    foreach ($Doc in $ProjectDocuments)
#        { 
#            if ($Doc.DocumentQt -like "single") { $DocumentQty += 1 }
#            if ($Doc.DocumentQt -like "multi") { $DocumentQty += $($ProjectSites.Count) }
#        }
#    Write-Host "Number of Documents is $DocumentQty"

#RETURN        
    $projectSites
}
function Read-ProjectDocuments
{
    [CmdletBinding()]
    param (
        [string]$inifile = ".\folder_ini.ini" 
    )
    $ini = Get-Content -Path $inifile 
    $ProjectDocuments = @() # list of hash tables of Project Documents Properties
    #$ProjectDocuments = @( @{ DocumentName = ; DocumentQt = ; DocumentPath = ; DocumentType = ; DocumentTemplate = } )

    $section = "none" #section name in ini file
    #Ini File parser
    foreach ($line in $ini)
        {
            if ($line -ne "" -and $line.startswith(";") -ne $true )
                {
    # Checking section name
                    if ($line.StartsWith("[General]") -eq $True) {$section = 'General'; continue}
                    if ($line.StartsWith("[Site]") -eq $True) {$section = 'Site'; continue}
                    if ($line.StartsWith("[Document]") -eq $True) {$section = 'Document'; continue}
    
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
    $DocumentQty = 0
    foreach ($Doc in $ProjectDocuments)
        { 
            if ($Doc.DocumentQt -like "single") { $DocumentQty += 1 }
            if ($Doc.DocumentQt -like "multi") { $DocumentQty += $($ProjectSites.Count) }
        }
    Write-Host "Number of Documents is $DocumentQty"
#RETURN        
    $ProjectDocuments
}    
function New-ItemList
{
    [CmdletBinding()]
    param (
        [hashtable]$ProjectData, [array]$ProjectSites, [array]$ProjectDocuments
    )
    $ProjectItems = @()
    $ItemSmallNumber = 1
    foreach ($Doc in $ProjectDocuments)
        {
            if ($Doc.DocumentQt -like "single")
            { 
            [string]$tempNumber = ($ProjectData.ProjectNumber * 1000) + $ItemSmallNumber
            $tempNumber += "-01"

            [string]$tempFileName = $ProjectData.ProjectName.Replace(' ', '_') + '_' + $Doc.DocumentName.Replace(' ', '_') + '_' + $tempNumber
#            [string]$tempItemType = $Doc.DocumentType

            if ($Doc.DocumentType -eq "visio") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.vsdx'}
            if ($Doc.DocumentType -eq "word") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.docx'}
            if ($Doc.DocumentType -eq "excel") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.xlsx'}
            if ($Doc.DocumentType -eq "none") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.vsdx'}

            $ProjectItems += @{ItemNumber = $tempNumber; ItemTitle = $Doc.DocumentName; ItemSite = "General"; ItemFileName = $tempFileName;
                ItemPath = $tempItemPath; ItemType = $Doc.DocumentType; ItemTemplate = $Doc.DocumentTemplate}
            $ItemSmallNumber++    
            }
            if ($Doc.DocumentQt -like "multi")
            {
                foreach ($site in $projectSites)
                {
                    [string]$tempNumber = ($ProjectData.projectNumber * 1000) + $ItemSmallNumber
                    $tempNumber += "-01"
                    
                    [string]$tempFileName = $Site.Replace(' ', '_') + '_' + $Doc.DocumentName.Replace(' ', '_') + '_' + $tempNumber
#                    [string]$tempItemType = $Doc.DocumentType
                    
                    if ($Doc.DocumentType -eq "visio") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.vsdx'}
                    if ($Doc.DocumentType -eq "word") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.docx'}
                    if ($Doc.DocumentType -eq "excel") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.xlsx'}
                    if ($Doc.DocumentType -eq "none") {[string]$tempItemPath = $Doc.DocumentPath + '\' + $tempFileName + '.vsdx'}

                    $ProjectItems += @{ItemNumber = $tempNumber; ItemTitle = $Doc.DocumentName; ItemSite = $site; ItemFileName = $tempFileName;
                        ItemPath = $tempItemPath; ItemType = $Doc.DocumentType; ItemTemplate = $Doc.DocumentTemplate}
                $ItemSmallNumber++
                }
            }
            
        }
    #Debug info
    Write-Host "Project Items" -BackgroundColor Green
    write-host "`n"   
    foreach ($item in $ProjectItems) 
        {   
            write-host $item.ItemNumber -BackgroundColor Green
            write-host $item.ItemTitle 
            write-host $item.ItemSite
            write-host $item.ItemFileName
            write-host $item.ItemPath
            write-host $item.ItemType $item.ItemTemplate
            write-host $item.ItemTemplate
            write-host "`n"   
        }
#RETURN        
    $ProjectItems
}

function new-foldersctrucure
{
    [CmdletBinding()]
    param (
        [string]$FolderStructurePath = ".\Folder_Structure.ini", [hashtable]$ProjectData
    )
    $folderStructure = Get-Content -path $FolderStructurePath
    [string]$ProjectPath = (get-location).path + "\" + $ProjectData.ProjectCountry + "_" + $($ProjectData.ProjectName.Replace(' ','_'))
    If (Test-Path -Path $ProjectPath) {Remove-Item -Path $ProjectPath -Force -Recurse}
    Write-host 'Folder Structer:' -BackgroundColor Green
    foreach ($line in $folderStructure)
        {
    #        if ($line -ne "" -and $line.startswith(";") -ne $true)
            if ($line -match '^\\')
                {
                    $lineFormat = $line.Trim("")
    #                $lineFormat = $lineFormat.Replace('','_')

                    New-Item -path $($ProjectPath + $lineFormat) -ItemType Directory -Force | Out-Null
                    Write-host $lineFormat
                } 
        }
    $ProjectPath
}
function New-DocList {
    param (
        [hashtable]$ProjectData, [array]$ProjectSites, [array]$ProjectItems, [string]$ProjectPath
    )
    $Word = New-Object -ComObject Word.application
    $Word.visible = $true
    $WordDocument = $Word.Documents.add()
    $Selection = $Word.Selection
    $WordSection = $WordDocument.Sections.Item(1)

    #Landscape orientation
    $WordDocument.PageSetup.Orientation = 1

    #Header
    $Header = $WordSection.Headers.Item(1)
    $Header.range.text = $ProjectData.ProjectCompany

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
    $Selection.TypeText("Project Name: $($ProjectData.ProjectName)")
    $Selection.TypeParagraph()
    $Selection.TypeText("Project Code: $($ProjectData.ProjectCode)")
    $Selection.TypeParagraph()
    $Selection.TypeText("Project Manager: $($ProjectData.ProjectManager)")
    $Selection.TypeParagraph()
    $Selection.TypeText("Project Engineer: $($ProjectData.ProjectEngineer)")

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

    $table = $Selection.tables.add($selection.range,$($ProjectItems.Count + 1),5) #Change number of Rows
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
    $table.cell($row,4).range.text = $ProjectData.ProjectEngineer
    $table.cell($row,5).range.text = "Document template was created automaticaly"
    $row++
    }
    [string]$ListDocName = $ProjectPath + "\" + $ProjectData.projectNumber + "-List-Doc"
    
    $WordDocument.saveas($ListDocName)
    Write-Host "List-Doc located in $ListDocName" -BackgroundColor Green
    Write-host "`n"
    $WordDocument.Close()
    $word.Application.quit()
    

}
######## COPY DOCS ################################################
function Copy-Docs {
    param (
        [hashtable]$ProjectData, [array]$ProjectSites, [array]$ProjectItems, [string]$ProjectPath
    )

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
            $OfficeDocument.BuiltInDocumentProperties("subject") = $ProjectData.projectName
            $OfficeDocument.BuiltInDocumentProperties("company") = $ProjectData.projectCompany
            $OfficeDocument.BuiltInDocumentProperties("category") = $SortedProjectItems[$i].ItemSite
            $OfficeDocument.BuiltInDocumentProperties("author") = $ProjectData.projectEngineer
            $OfficeDocument.BuiltInDocumentProperties("manager") = $ProjectData.projectManager

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
            $OfficeDocument.Subject = $ProjectData.ProjectName
            $OfficeDocument.Company = $ProjectDataProjectCompany
            $OfficeDocument.Category = $SortedProjectItems[$i].ItemSite
            $OfficeDocument.Creator = $ProjectData.ProjectEngineer
            $OfficeDocument.Manager = $ProjectData.ProjectManager

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
            $OfficeDocument.BuiltInDocumentProperties("subject") = $ProjectData.ProjectName
            $OfficeDocument.BuiltInDocumentProperties("company") = $ProjectData.ProjectCompany
            $OfficeDocument.BuiltInDocumentProperties("category") = $SortedProjectItems[$i].ItemSite
            $OfficeDocument.BuiltInDocumentProperties("author") = $ProjectData.ProjectEngineer
            $OfficeDocument.BuiltInDocumentProperties("manager") = $ProjectData.ProjectManager

            [string]$path =  $ProjectPath + "\" + $SortedProjectItems[$i].ItemPath
            $OfficeDocument.saveas($path)
            Write-host $SortedProjectItems[$i].ItemPath
            $OfficeDocument.Close()
            if ($i -eq $WordIndexLast) {$OfficeApplication.Quit()}
        }
            
    }    
        
}

#######   MAIN   ###########################################
$ProjectIniFilePath = ".\Project_ini.ini"
$ProjectData = read-projectdata $ProjectIniFilePath
$Confirm = Read-Host "Please check Project Data and Confirm Y/N"
if ($Confirm -ne "y")
    {
        Stop-Transcript
        Exit
    }
$ProjectSites = read-projectsites $ProjectIniFilePath
$Confirm = Read-Host "Please check Project Sites and Confirm Y/N"
if ($Confirm -ne "y")
    {
        Stop-Transcript
        Exit
    }
$ProjectDocuments = read-projectdocuments $ProjectIniFilePath
$ProjectItems = New-ItemList $ProjectData $ProjectSites $ProjectDocuments
$Confirm = Read-Host "Please check Project Documents and Confirm to create Folder Structure Y/N"
if ($Confirm -ne "y")
    {
        Stop-Transcript
        Exit
    }
$ProjectPath = new-foldersctrucure .\Folder_Structure.ini $ProjectData
New-DocList $ProjectData $ProjectSites $ProjectItems $ProjectPath
$Confirm = Read-Host "Please Confirm to copy documens from templates Y/N"
if ($Confirm -ne "y")
    {
        Stop-Transcript
        Exit
    }
Copy-Docs $ProjectData $ProjectSites $ProjectItems $ProjectPath
$Confirm = Read-Host "Finished, press any button"

Stop-Transcript 