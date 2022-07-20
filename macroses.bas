
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
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .AutoAdjustRightIndent = True
        .WordWrap = True
        .BaseLineAlignment = wdBaselineAlignAuto
       ' Selection.Style = ActiveDocument.Styles("Body_table Text")
    End With
End Sub

Sub Change_Table_Formatting_CodeBox(tbl As Word.Table)
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

Public Sub Format_Tables_in_Tables()
  Dim Table As Word.Table
  Dim Table2 As Word.Table
  Dim FirstCell As String
  Dim FirstCellStr As String
  Dim SecondColumn As String
  Dim SecondColumnStr As String
  'Vektor ataki
  Type1 = ChrW(1042) + ChrW(1077) + ChrW(1082) + ChrW(1090) + ChrW(1086) + ChrW(1088) + ChrW(32) + ChrW(1072) + ChrW(1090) + ChrW(1072) + ChrW(1082) + ChrW(1080)
  'Ustarevshee PO
  Type2 = ChrW(1059) + ChrW(1089) + ChrW(1090) + ChrW(1072) + ChrW(1088) + ChrW(1077) + ChrW(1074) + ChrW(1096) + ChrW(1077) + ChrW(1077) + ChrW(32) + ChrW(1055) + ChrW(1054)
  'Razdel
  Type3 = ChrW(1056) + ChrW(1072) + ChrW(1079) + ChrW(1076) + ChrW(1077) + ChrW(1083)
  'Identificator
  Type4 = ChrW(1048) + ChrW(1076) + ChrW(1077) + ChrW(1085) + ChrW(1090) + ChrW(1080) + ChrW(1092) + ChrW(1080) + ChrW(1082) + ChrW(1072) + ChrW(1090) + ChrW(1086) + ChrW(1088)
  For Each Table2 In ActiveDocument.Tables
    Table2.PreferredWidthType = wdPreferredWidthPoints
    If Table2.Columns.Count = 1 And Table2.Rows.Count = 1 Then
        Call Change_Table_Formatting_CodeBox(Table2)
    Else
        Call Change_Table_Type(Table2)
        Table2.PreferredWidth = CentimetersToPoints(16.8)
        FirstCell = Table2.Cell(1, 1).Range.Text
        If Table2.Rows(1).Cells.Count = 1 And Table2.Rows.Count > 1 Then
            SecondColumnStr = ""
        Else
            SecondColumn = Table2.Cell(1, 2).Range.Text
            SecondColumnStr = Left(SecondColumn, Len(SecondColumn) - 2)
        End If
       ' MsgBox (Len(FirstCell))
        FirstCellStr = Left(FirstCell, Len(FirstCell) - 2)
       ' MsgBox (FirstCellStr)
        If FirstCellStr = ChrW(8470) Or FirstCellStr = "#" Or FirstCellStr = Type1 Or FirstCellStr = Type2 Then
            Table2.Rows(1).Select
            With Selection.Font
                .Bold = True
            End With
        End If
        'Razdel OWASP
        If InStr(SecondColumnStr, Type3) <> 0 Then
            Table2.Columns(1).Width = CentimetersToPoints(0.63)
            Table2.Columns(2).Width = CentimetersToPoints(8.61)
            Table2.Columns(3).Width = CentimetersToPoints(1.75)
            Table2.Columns(4).Width = CentimetersToPoints(5.81)
        End If
        'Identifikator vulns
        If SecondColumnStr = Type4 Then
            Table2.Columns(1).Width = CentimetersToPoints(0.63)
            Table2.Columns(2).Width = CentimetersToPoints(3.4)
            Table2.Columns(3).Width = CentimetersToPoints(2.27)
            Table2.Columns(4).Width = CentimetersToPoints(5.5)
            Table2.Columns(5).Width = CentimetersToPoints(5)
            Table2.Select
            Selection.Font.Size = 10
        End If
        
        With Table2
            .TopPadding = CentimetersToPoints(0)
            .BottomPadding = CentimetersToPoints(0)
            .LeftPadding = CentimetersToPoints(0.1)
            .RightPadding = CentimetersToPoints(0.1)
            .Spacing = 0
            .AllowPageBreaks = True
            .AllowAutoFit = False
        End With
    End If
    For Each Table In Table2.Tables
        Call Change_Table_Formatting_CodeBox(Table)
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

Sub Picture_Center_Border()
'
' Picture to Center, with Border Macro
'
'
    Dim inshape As InlineShape, shape As shape
   
    '1. disable screenupdating
    Application.ScreenUpdating = False
   
    '2. Paragraph of picture formatting
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
   
    '3. Picture to the center
    For Each shape In ActiveDocument.Shapes
        shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
        shape.Left = wdShapeCenter
    Next shape
   
    '4. Enable screenupdating
    Application.ScreenUpdating = True
End Sub

Sub Picture_Center_Border_Anonym()
    '
    ' The same but also added picture effect to hide content
    '
    Dim inshape As InlineShape, shape As shape
   
    Application.ScreenUpdating = False
   
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
        inshape.Fill.PictureEffects.Insert msoEffectGlass
        inshape.Line.Weight = 0.5
    Next inshape
   
    For Each shape In ActiveDocument.Shapes
        shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
        shape.Left = wdShapeCenter
    Next shape
   
    Application.ScreenUpdating = True
End Sub

Sub Paste_only_text()
'
' Paste_only_text Macro for hot keys, e.g. Cmd+Option+V
'
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:= _
        wdInLine, DisplayAsIcon:=False
End Sub

Sub OneTable()
'
' OneTable Macro to insert a table with hot keys, e.g. Cmd+Option+T
'
'
    Selection.Tables(1).Style = "Table Grid"
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = True
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
    Selection.Font.Size = 10
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

Public Sub TableSizeCorrect()
  Dim Table As Word.Table
  For Each Table In ActiveDocument.Tables
    Call FuncTableSizeCorrect(Table)
  Next Table
End Sub


Sub Change_List_Format_by_params(lis As Word.List)
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

Sub Change_List_Format_by_Style(lis As Word.List)
'
' Apply style "List1" to all bulleted lists
'
    Dim list_type As String
    Dim Paragr As Word.Paragraph
    For Each Paragr In lis.ListParagraphs
        If Paragr.Range.ListFormat.ListType = "2" Then
            With Paragr
                .Range.Style = "List1"
            End With
        End If
    Next Paragr
End Sub

Public Sub Format_Lists()
'
' Change style of all lists
'
  Dim List As Word.List
  For Each List In ActiveDocument.Lists
    Call Change_List_Format_by_Style(List)
  Next List
End Sub

Sub Format_one_list()
'
' List of formatting steps
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
        With Selection.Find
           .ClearFormatting
           .Replacement.ClearFormatting
           .Text = "“(*)”"
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
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      ''Change inseparable space to normal space
        .Text = ChrW(160)
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
    With Selection.Find
    ''Change short dash to medium dash if it's between spaces
        .Text = " - "
        .Replacement.Text = " " + ChrW(8211) + " "
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
        .Text = ChrW(8212)
        .Replacement.Text = ChrW(8211)
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

Sub Footnote_add()
  Call Remove_Hyperlinks
  Dim oRng As Range
  Dim strText As String
  Dim i As Long
      Set oRng = ActiveDocument.Range
      i = 1
      With oRng.Find
          Do While .Execute(FindText:="\[footnote:(*)\]", MatchWildcards:=True)
              'oRng.MoveEndUntil Chr(46)
              If oRng.Characters.Last = "]" Then
                  'oRng.End = oRng.End + 1
                  oRng = Replace(oRng, "[footnote:", "")
                  oRng = Replace(oRng, "]", "")
                  'MsgBox (oRng)
                  ActiveDocument.Footnotes.Add Range:=oRng, Text:=oRng.Text
                  oRng.Text = ""
                  i = i + 1
              End If
              oRng.Collapse 0
          Loop
      End With
lbl_Exit:
    Set oRng = Nothing
    Exit Sub

End Sub

Sub Remove_Hyperlinks()
    While ActiveDocument.Hyperlinks.Count > 0
    ActiveDocument.Hyperlinks(1).Delete
    Wend
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
End Sub

Sub All_format_run()
    Call Change_some_symbols
    Call Change_type_of_quotes
    Call Remove_Hyperlinks
    Call Footnote_add
End Sub


Sub Change_Table_Formatting_old(tbl As Word.Table)
    tbl.Select
    With tbl
        .Style = "Table Grid"
    End With
    tbl.Rows.HeightRule = wdRowHeightAuto
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 12
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

Public Sub Format_Tables_old()
  Dim Table As Word.Table
  For Each Table In ActiveDocument.Tables
    Call Change_Table_Formatting_old(Table)
  Next Table
End Sub
