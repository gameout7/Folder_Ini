
$ini = Get-Content -Path .\folder_ini.ini 

$section = "none" #section name in ini file
$projectName = "no project name"
$ProjectNumber = "no project number"
$ProjectCode = "no project code"
$Author = "no author"
$Manager = "no manager"
$Company = "no company"

$Site = @()
$Document = @()
$iniHash = @{}
$iniTemp = @()

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
                    if ($SpliteArray[0].trim('') -eq "Author") {$Author = $SpliteArray[1].trim('')}
                    if ($SpliteArray[0].trim('') -eq "Manager") {$Manager = $SpliteArray[1].trim('')}
                    if ($SpliteArray[0].trim('') -eq "Company") {$Company = $SpliteArray[1].trim('')}
                }
# Section Site
                if ($section -eq 'Site')
                {
                    $Site += $Line.Trim('')
                }
# Section Document
                if ($section -eq 'Document')
                {
                    $SpliteArray = $Line.Split(",")
                    $Document += @{DocumentName=$SpliteArray[0].trim(''); DocumentQt=$SpliteArray[1].trim(''); DocumentPath=$SpliteArray[2].trim('')}
                }                

            }

    }

#Debug info
    "Project Name is $projectName"
    "Project Code is $ProjectCode"
    "Project Number is $ProjectNumber"    
    "Project Author is $Author"
    "Project Manager is $Manager"
    "Project Company is $Company"
    $site
    $Document
#>
