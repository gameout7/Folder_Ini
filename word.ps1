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
$Header.range.text = "header text"

#Footer
$Footer = $WordSection.footers.item(1)
$Footer.range.text = "footer text"

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
$Selection.TypeText("Project Name: STC")
$Selection.TypeParagraph()
$Selection.TypeText("Project Code: 123")
$Selection.TypeParagraph()
$Selection.TypeText("Project Manager: Name")
$Selection.TypeParagraph()
$Selection.TypeText("Project Manager: Name")

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

$table = $Selection.tables.add($selection.range,2,5) #Change number of Rows
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

$table.cell(2,1).range.text = "1 Doc No"
$table.cell(2,2).range.text = "1 Document name"
$table.cell(3,3).range.text = "1 Date"
$table.cell(4,4).range.text = "1 Creator"
$table.cell(5,5).range.text = "1 Note"