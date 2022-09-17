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
<#
foreach ($item in $SortedProjectItems)
    {
        If ($item.itemType -eq "visio")

        {
            $OfficeApplication = New-Object -ComObject Visio.application
            [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate
            $OfficeDocument = $OfficeApplication.Documents.open($templatename)

            $OfficeDocument.title = $item.ItemTitle
            $OfficeDocument.Description = $item.ItemNumber
            $OfficeDocument.Subject = $projectName
            $OfficeDocument.Company = $projectCompany
            $OfficeDocument.Category = $Item.ItemSite
            $OfficeDocument.Creator = $projectEngineer
            $OfficeDocument.Manager = $projectManager

            [string]$path =  $ProjectPath + "\" + $item.ItemPath
            $OfficeDocument.saveas($path)
            $OfficeDocument.Close()
            $OfficeApplication.Quit()
        }
        If ($item.itemType -eq "excel")
        {
            $OfficeApplication = New-Object -ComObject excel.application
            [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate
            
            $OfficeDocument = $OfficeApplication.Workbooks.open($templatename)
            $OfficeDocument.BuiltInDocumentProperties("title") = $item.ItemTitle
            $OfficeDocument.BuiltInDocumentProperties("comments") = $item.ItemNumber
            $OfficeDocument.BuiltInDocumentProperties("subject") = $projectName
            $OfficeDocument.BuiltInDocumentProperties("company") = $projectCompany
            $OfficeDocument.BuiltInDocumentProperties("category") = $Item.ItemSite
            $OfficeDocument.BuiltInDocumentProperties("author") = $projectEngineer
            $OfficeDocument.BuiltInDocumentProperties("manager") = $projectManager

            [string]$path =  $ProjectPath + "\" + $item.ItemPath
            $OfficeDocument.saveas($path)
            $OfficeDocument.Close()
            $OfficeApplication.Quit()
        }
        If ($item.itemType -eq "word")
        {
            $OfficeApplication = New-Object -ComObject word.application
            [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate

            $OfficeDocument = $OfficeApplication.Documents.open($templatename)
            $OfficeDocument.BuiltInDocumentProperties("title") = $item.ItemTitle
            $OfficeDocument.BuiltInDocumentProperties("comments") = $item.ItemNumber
            $OfficeDocument.BuiltInDocumentProperties("subject") = $projectName
            $OfficeDocument.BuiltInDocumentProperties("company") = $projectCompany
            $OfficeDocument.BuiltInDocumentProperties("category") = $Item.ItemSite
            $OfficeDocument.BuiltInDocumentProperties("author") = $projectEngineer
            $OfficeDocument.BuiltInDocumentProperties("manager") = $projectManager

            [string]$path =  $ProjectPath + "\" + $item.ItemPath
            $OfficeDocument.saveas($path)
            $OfficeDocument.Close()
            $OfficeApplication.Quit()
        }
        If ($item.itemType -eq "none")
        {
            [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate
            [string]$path =  $ProjectPath + "\" + $item.ItemPath
            Copy-Item -Path $templatename -Destination $path
        }


    }

#>
    
<#
$OfficeApplication = New-Object -ComObject Visio.application

foreach ($item in $ProjectItems)
    {
        [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate
        $OfficeDocument = $OfficeApplication.Documents.open($templatename)

        $OfficeDocument.title = $item.ItemTitle
        $OfficeDocument.Description = $item.ItemNumber
        $OfficeDocument.Subject = $projectName
        $OfficeDocument.Company = $projectCompany
        $OfficeDocument.Category = $Item.ItemSite
        $OfficeDocument.Creator = $projectEngineer
        $OfficeDocument.Manager = $projectManager

        [string]$path =  $ProjectPath + "\" + $item.ItemPath
        $OfficeDocument.saveas($path)
        $OfficeDocument.Close()

    }

    $OfficeApplication.Quit()
#>