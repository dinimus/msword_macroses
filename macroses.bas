
Sub Change_Table_Formatting(tbl As Word.Table)
    tbl.Select
    With tbl
        .Style = "Table Grid"
    End With
    tbl.Rows.HeightRule = wdRowHeightAuto
    With Selection.Font
        .Name = "Courier New"
        .Size = 9
        .Color = wdColorAutomatic
        .Bold = False
    End With
    With tbl
        .Borders(wdBorderLeft).Color = wdColorAutomatic
        .Borders(wdBorderRight).Color = wdColorAutomatic
        .Borders(wdBorderTop).Color = wdColorAutomatic
        .Borders(wdBorderBottom).Color = wdColorAutomatic
        .Borders(wdBorderHorizontal).Color = wdColorAutomatic
        .Borders(wdBorderVertical).Color = wdColorAutomatic
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
''' background color:
  ' Selection.Shading.BackgroundPatternColor = -603923969
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .Alignment = wdAlignParagraphLeft
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    With Options
        .DefaultBorderColor = wdColorAutomatic
    End With
End Sub

Public Sub Format_Tables()
  Dim Table As Word.Table
  For Each Table In ActiveDocument.Tables
    Call Change_Table_Formatting(Table)
  Next Table
End Sub

Sub Change_Table_Type(tbl As Word.Table)
    tbl.Select
    With tbl
        .Style = "Table Grid"
    End With
    tbl.Rows.HeightRule = wdRowHeightAuto
    With Selection.Font
        .Color = wdColorAutomatic
        
    End With
    With tbl
        .Borders(wdBorderLeft).Color = wdColorAutomatic
        .Borders(wdBorderRight).Color = wdColorAutomatic
        .Borders(wdBorderTop).Color = wdColorAutomatic
        .Borders(wdBorderBottom).Color = wdColorAutomatic
        .Borders(wdBorderHorizontal).Color = wdColorAutomatic
        .Borders(wdBorderVertical).Color = wdColorAutomatic
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
 '''  background color:
    'Selection.Shading.BackgroundPatternColor = -603923969
    With Options
        .DefaultBorderColor = wdColorAutomatic
    End With
    With Selection.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
    End With
End Sub

Public Sub Format_Tables_in_Tables()
  Dim Table As Word.Table
  Dim Table2 As Word.Table
  For Each Table2 In ActiveDocument.Tables
    Call Change_Table_Type(Table2)
    If Table2.Columns.Count = 1 And Table2.Rows.Count = 1 Then
        Call Change_Table_Formatting(Table2)
    End If
    For Each Table In Table2.Tables
        Call Change_Table_Formatting(Table)
    Next Table
  Next Table2
End Sub

Sub InsertTableCode()
'
' InsertTableCode Macro
'
'
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:= _
        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
    End With
    Selection.Font.Name = "Courier New"
    Selection.Font.Size = 10
        ' background color:
  '  Selection.Shading.BackgroundPatternColor = -603923969
End Sub

Sub Remove_Hyperlinks()
    While ActiveDocument.Hyperlinks.Count > 0
    ActiveDocument.Hyperlinks(1).Delete
    Wend
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
End Sub

Sub Picture_Center_Border()
'
' Picture_Center_Border Macro
'
'
    Dim inshape As InlineShape, shape As shape
   
    '1. ????. ????????.
    Application.ScreenUpdating = False
   
    '2. ????????? ????????, ??????????? ? ??????.
    For Each inshape In ActiveDocument.InlineShapes
        With inshape.Range.ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .FirstLineIndent = CentimetersToPoints(0)
            '.Borders.OutsideLineColor = wdColorBlack
            '.Borders.OutsideLineStyle = wdLineStyleSingle
            .SpaceBefore = 6
            .SpaceBeforeAuto = False
            .SpaceAfter = 6
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
        End With
        inshape.Fill.Solid  ' picture line style
        inshape.Line.Weight = 0.5
    Next inshape
   
    '3. ????????? ???????? ????????.
    For Each shape In ActiveDocument.Shapes
        shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
        shape.Left = wdShapeCenter
    Next shape
   
    '4. ???. ????????.
    Application.ScreenUpdating = True
End Sub

Public Sub TableSizeCorrect()
  Dim Table As Word.Table
  For Each Table In ActiveDocument.Tables
    Call FuncTableSizeCorrect(Table)
  Next Table
End Sub

Sub FuncTableSizeCorrect(tbl As Word.Table)
    With tbl
        .Rows.LeftIndent = CentimetersToPoints(0.1)
        .PreferredWidthType = wdPreferredWidthPoints
        .PreferredWidth = CentimetersToPoints(18)
        .TopPadding = CentimetersToPoints(0)
        .BottomPadding = CentimetersToPoints(0)
        .LeftPadding = CentimetersToPoints(0.1)
        .RightPadding = CentimetersToPoints(0.1)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = False
        '.Columns.Item(1).PreferredWidth = 118.8
    End With
    tbl.Select
    Selection.Font.Color = wdColorAutomatic
End Sub

Sub Change_List_Format(lis As Word.List)
'
' List formatting steps
'
    Dim Paragr As Word.Paragraph
    For Each Paragr In lis.ListParagraphs
        With Paragr
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .LeftIndent = CentimetersToPoints(0.75)
            .RightIndent = CentimetersToPoints(0)
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceAfter = 0
            .SpaceBefore = 0
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = CentimetersToPoints(-0.5)
            .TabStops.ClearAll
        End With
    Next Paragr
End Sub

Public Sub Format_Lists()
  Dim List As Word.List
  For Each List In ActiveDocument.Lists
    Call Change_List_Format(List)
  Next List
End Sub

Sub Change_list()
'
' List formatting steps
'
        With Selection.Paragraphs
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .LeftIndent = CentimetersToPoints(0.75)
            .RightIndent = CentimetersToPoints(0)
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceAfter = 0
            .SpaceBefore = 0
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = CentimetersToPoints(-0.5)
            .TabStops.ClearAll
        End With
End Sub

Sub Change_type_of_quotes()
        'change " to elochki
        Dim blnQuotes As Boolean
        'save user settings
        blnQuotes = Options.AutoFormatAsYouTypeReplaceQuotes
        Options.AutoFormatAsYouTypeReplaceQuotes = False
        With Selection.Find
           .ClearFormatting
           .Replacement.ClearFormatting
           .Text = """(*)"""
           .Replacement.Text = "«\1»"
           .Forward = True
           .Wrap = wdFindContinue
           .MatchWildcards = True
           .Execute Replace:=wdReplaceAll
        End With
        'restore user settings
        Options.AutoFormatAsYouTypeReplaceQuotes = blnQuotes

End Sub
Sub Change_some_symbols()
' Change_some_symbols
'
    'Selection.Find.ClearFormatting
    'Selection.Find.Replacement.ClearFormatting
    With Selection.Find
    ''Change short dash to medium dash if it's between spaces
        .Text = " - "
        .Replacement.Text = " ^0150 "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
    ''Change long dash to medium dash
        .Text = "^0151"
        .Replacement.Text = "^0150"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
      With Selection.Find
      ''Change inseparable space to normal space
        .Text = "^0160"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


