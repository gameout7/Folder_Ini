
#Var definition
$ini = Get-Content -Path .\folder_ini.ini 

$section = "none" #section name in ini file
$projectName = "no project name"
[int]$projectNumber = 000
$projectCode = "no project code"
$projectEngineer = "no engineer"
$projectManager = "no manager"
$projectCompany = "no company"

$projectSites = @() # list of project site
$ProjectDocuments = @() # list of hash tables of Project Documents Properties
#$ProjectDocuments = @( @{ DocumentName = ; DocumentQt = ; DocumentPath = } )

$ProjectItems = @() # list of hash tables of Project items
#ProjectItems =  @( @{ ItemNumber = ; ItemTitle = ; ItemSite = ; ItemFileName =  ;ItemPath})

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
                    $ProjectDocuments += @{ DocumentName = $SpliteArray[0].trim(''); DocumentQt = $SpliteArray[1].trim(''); DocumentPath = $SpliteArray[2].trim('')}
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

            [string]$tempItemPath = $Doc.DocumentPath + '/' + $tempFileName

            $ProjectItems += @{ItemNumber = $tempNumber; ItemTitle = $Doc.DocumentName; ItemSite = "General"; ItemFileName = $tempFileName; ItemPath = $tempItemPath}
            $ItemSmallNumber++
        }
        if ($Doc.DocumentQt -like "multi")
        {
            
            foreach ($site in $projectSites)
            {
                [string]$tempNumber = ($projectNumber * 1000) + $ItemSmallNumber
                $tempNumber += "-01"
                
                [string]$tempFileName = $Site.Replace(' ', '_') + '_' + $Doc.DocumentName.Replace(' ', '_') + '_' + $tempNumber
                
                [string]$tempItemPath = $Doc.DocumentPath + '/' +$tempFileName

                $ProjectItems += @{ItemNumber = $tempNumber; ItemTitle = $Doc.DocumentName; ItemSite = $site; ItemFileName = $tempFileName; ItemPath = $tempItemPath}
                $ItemSmallNumber++
            }

            

        }
    }
"Project Items"    
$ProjectItems[0..$($ProjectItems.Length - 1)]
#>
