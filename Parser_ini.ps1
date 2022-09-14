
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
"Project Items"    
foreach ($item in $ProjectItems) 
    {   
        $item
        write-host "`n"   
    }
#$ProjectItems[0..$($ProjectItems.Length - 1)]
#>