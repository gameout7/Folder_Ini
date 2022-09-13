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