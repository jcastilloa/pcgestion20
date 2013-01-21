Attribute VB_Name = "FlexGridHTML"
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Function HTMLColor(ByVal c As Long) As String
'
' converts a VB color into HTML color code
'
    Dim s$
    
    ' convert to hex
    s = Hex(c)
    
    ' handle system colors
    If Left(s, 1) = "8" Then
        c = Val("&H" & Mid(s, 2))
        c = GetSysColor(c)
        s = Hex(c)
    End If
    
    ' build format
    s = String(6 - Len(s), "0") & s
    HTMLColor = """#" & Right(s, 2) & Mid(s, 3, 2) & Left(s, 2) & """"
End Function


Function HTMLText(ByVal s$) As String
'
' converts a VB string into an HTML string
' this involves replacing special characters "&", "<", and ">".
'
    StringReplace s, "&", "&amp;"
    StringReplace s, "<", "&lt;"
    StringReplace s, ">", "&gt;"
    If s = "" Then s = "&nbsp;"
    HTMLText = s
    
End Function

Function IsNumber(txt$) As Boolean
'
' checks a string to see whether it's a number
'
    Dim i%, c$, s$, hasdec%
    s = Trim(txt)
    For i = 1 To Len(s)
        
        ' get character to test
        c = Mid(s, i, 1)
        
        ' plus and minus are OK only when they are first
        If c = "+" Or c = "-" And i > 1 Then c = "x"
        
        ' only one decimal point is allowed
        If c = "." Then
            If hasdec Then c = "x"
            hasdec = True
        End If
        
        ' that's it (no currency signs or parenthesis allowed)
        If InStr("0123456789,.", c) = 0 Then
            IsNumber = False
            Exit Function
        End If
    Next
    
    ' if you got here, you're a number
    IsNumber = True
End Function

Function SaveFlexGridToHTML(fa As VSFlexGrid, fn$, pgb As FWProgressBar) As Boolean
'
' saves the given FlexGrid control into an HTML file.
' fn is the file name for the HTML file, including path and extension.
' returns True if successful, False otherwise.
'
' we don't do pictures
' we don't do font sizing
'
    ' additional width for HTML columns
    Const EXTRAWIDTH = 1.3
    
    ' table header (same for all tables)
    Const TABLEHDR = "<HTML>" & vbCrLf & _
        "<HEAD>" & vbCrLf & _
        "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html;charset=windows-1252"">" & vbCrLf & _
        "<META NAME=""Generator"" CONTENT=""VSFlexGrid Pro"">" & vbCrLf & _
        "<TITLE>VSFlexGrid Pro</TITLE>" & vbCrLf & _
        "</HEAD>" & vbCrLf & _
        "<BODY>" & vbCrLf & vbCrLf
    
    '----------------------------------------------------------------------
    ' open HTML output file
    ' UNDONE: On Error Resume Next
    Dim f%
    f = FreeFile
    Open fn For Output As f
    If Err <> 0 Then Exit Function
    
    '----------------------------------------------------------------------
    ' save heading information
    Print #f, TABLEHDR
              
    '----------------------------------------------------------------------
    ' get total table width in pixels
    Dim i%, tblwid!
    tblwid = 0
    For i = 0 To fa.Cols - 1
        tblwid = tblwid + fa.ColWidth(i)
    Next
    If tblwid < fa.Width Then tblwid = fa.Width
    tblwid = EXTRAWIDTH * tblwid / Screen.TwipsPerPixelX
    
    '----------------------------------------------------------------------
    ' save table header
    Dim s$
    s = "<FONT FACE=""" & fa.FontName & """ SIZE=1>" & vbCrLf & _
        "<TABLE BORDER CELLSPACING=0 CELLPADDING=2 VALIGN=CENTER" & _
        " BGCOLOR=" & HTMLColor(fa.BackColor) & _
        " BORDERCOLOR=" & HTMLColor(fa.GridColor) & _
        " WIDTH=" & Format(Int(tblwid)) & _
        ">" & vbCrLf
    Print #f, s
    
    '----------------------------------------------------------------------
    ' loop through the rows
    Dim r&, c&
    
    pgb.Max = fa.Rows
    
    For r = 0 To fa.Rows - 1
                
        pgb.Value = pgb.Value + 1
                
        '------------------------------------------------------------------
        ' skip hidden rows
        If fa.RowHidden(r) Or (fa.RowHeight(r) = 0) Then GoTo nextRow
        
        '------------------------------------------------------------------
        ' start row
        Print #f, "<TR>"
        
        '------------------------------------------------------------------
        ' loop through the columns
        For c = 0 To fa.Cols - 1
            
            '--------------------------------------------------------------
            ' skip hidden cols
            If fa.ColHidden(c) Or (fa.ColWidth(c) = 0) Then GoTo nextCol
            
            '--------------------------------------------------------------
            ' handle merges
            ' var: span
            Dim span$
            span = ""
            Dim r1&, c1&, r2&, c2&
            fa.GetMergedRange r, c, r1, c1, r2, c2
            If c1 < c Then GoTo nextCol
            If r1 < r Then GoTo nextCol
            If c2 > c Then span = " COLSPAN=" & (c2 - c + 1)
            If r2 > r Then span = " ROWSPAN=" & (r2 - r + 1)
            
            '--------------------------------------------------------------
            ' get col width
            ' var: wid
            Dim wid!
            wid = 0
            For i = c1 To c2
                wid = wid + fa.ColWidth(i)
            Next
            wid = EXTRAWIDTH * wid / Screen.TwipsPerPixelX
            
            '--------------------------------------------------------------
            ' get cell text
            ' var: txt
            Dim txt$
            txt = fa.Cell(flexcpTextDisplay, r, c)
            If r >= fa.FixedRows And fa.ColDataType(c) = flexDTBoolean Then
                If Val(txt) Then txt = "Y" Else txt = ""
            End If
            txt = HTMLText(txt)
            
            '--------------------------------------------------------------
            ' get outline indent
            ' var: txt
            Dim olni$
            If fa.OutlineBar > 0 And c = fa.OutlineCol Then
                If fa.IsSubtotal(r) Then
                    olni = ""
                    For i = 1 To fa.RowOutlineLevel(r)
                        olni = "&nbsp&nbsp&nbsp&nbsp" & olni
                    Next
                End If
                txt = olni & txt
            End If
            
            '--------------------------------------------------------------
            ' get back color
            ' var: bkg
            Dim bkg$, clr&
            bkg = ""
            clr = fa.Cell(flexcpBackColor, r, c)
            If clr <> 0 Then
                bkg = " BGCOLOR=" & HTMLColor(clr)
            ElseIf r < fa.FixedRows Or c < fa.FixedCols Then
                bkg = " BGCOLOR=" & HTMLColor(fa.BackColorFixed)
            End If
            
            '--------------------------------------------------------------
            ' get border color
            ' var: bdr
            Dim bdr$
            bdr = ""
            If r < fa.FixedRows Or c < fa.FixedCols Then
                bdr = " BORDERCOLOR=" & HTMLColor(fa.GridColorFixed)
            End If
            
            '--------------------------------------------------------------
            ' get fore color and font name
            ' var: fnt
            Dim fnt$
            fnt = ""
            s = fa.Cell(flexcpFontName, r, c)
            If s <> fa.FontName Then
                fnt = " FACE=" & """" & s & """"
            End If
            clr = fa.Cell(flexcpForeColor, r, c)
            If clr <> 0 Then
                fnt = " COLOR=" & HTMLColor(clr)
            ElseIf r < fa.FixedRows Or c < fa.FixedCols Then
                fnt = " COLOR=" & HTMLColor(fa.ForeColorFixed)
            End If
            
            '--------------------------------------------------------------
            ' get font effects
            ' var: ffx
            Dim ffx$
            ffx = ""
            If fa.Cell(flexcpFontBold, r, c) Then ffx = ffx & "<B>"
            If fa.Cell(flexcpFontItalic, r, c) Then ffx = ffx & "<I>"
            If fa.Cell(flexcpFontUnderline, r, c) Then ffx = ffx & "<U>"
            
            '--------------------------------------------------------------
            ' get alignment
            ' var: aln
            Dim aln$
            aln = ""
            Select Case fa.ColAlignment(c)
                Case flexAlignCenterBottom
                    aln = " ALIGN=CENTER VALIGN=BOTTOM"
                Case flexAlignCenterCenter
                    aln = " ALIGN=CENTER"
                Case flexAlignCenterTop
                    aln = " ALIGN=CENTER VALIGN=TOP"
                Case flexAlignLeftBottom
                    aln = " VALIGN=BOTTOM"
                Case flexAlignLeftCenter
                    aln = ""
                Case flexAlignLeftTop
                    aln = " VALIGN=TOP"
                Case flexAlignRightBottom
                    aln = " ALIGN=RIGHT VALIGN=BOTTOM"
                Case flexAlignRightCenter
                    aln = " ALIGN=RIGHT"
                Case flexAlignRightTop
                    aln = " ALIGN=RIGHT VALIGN=TOP"
                Case Else
                    Select Case fa.ColDataType(c)
                        Case flexDTBoolean
                            aln = " ALIGN=CENTER"
                        Case flexDTDate
                            aln = ""
                        Case Else
                            If IsNumber(fa.Cell(flexcpTextDisplay, r, c)) Then
                                aln = " ALIGN=RIGHT"
                            End If
                    End Select
            End Select
            
            '--------------------------------------------------------------
            ' build HTML cell string
            s = """" & Format(wid / tblwid, "#%") & """"
            s = "<TD WIDTH=" & s & bkg & aln & bdr & span & ">"
            If fnt <> "" Then s = s & "<FONT" & fnt & ">"
            s = s & ffx & txt
            If InStr(ffx, "B") > 0 Then s = s & "</B>"
            If InStr(ffx, "I") > 0 Then s = s & "</I>"
            If InStr(ffx, "U") > 0 Then s = s & "</U>"
            If fnt <> "" Then s = s & "</FONT>"
            
            '--------------------------------------------------------------
            ' end cell
            s = s & "</TD>"
            Print #f, s
nextCol:
        Next
        
        '------------------------------------------------------------------
        ' end row
        Print #f, "</TR>"
nextRow:
    Next r
    
    ' save table end
    Dim tblFtr$
    tblFtr = "</TABLE>" & vbCrLf & vbCrLf & _
             "</BODY>" & vbCrLf & _
             "</HTML>" & vbCrLf
    Print #f, tblFtr & vbCrLf
    
    ' close file
    Close f
    
    ' return success
    SaveFlexGridToHTML = True
End Function



Private Sub StringReplace(s$, ByVal f$, ByVal r$)
    Dim i&, l&

    ' sanity check
    l = Len(f)
    If l = 0 Then Exit Sub

    ' find and replace
    i = 0
    Do
        i = InStr(i + 1, s, f)
        If i = 0 Then Exit Do
        s = Left(s, i - 1) & r & Mid(s, i + l)
    Loop
End Sub


