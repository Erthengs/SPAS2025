Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports System.IO
Module Output


    Sub Print_Excasso_form()
        '------------------- checks vooraf ------------------------------------------------
        If SPAS.Dgv_Excasso_numbers.Rows(1).Cells(10).Value = "0" And SPAS.Dgv_Excasso_numbers.Rows(2).Cells(10).Value = "0" Then Exit Sub


        '-------------------- afhandeling bestandslocatie ------------------------------
        '--- kijkt eerst naar het vorig gekozen locatie, indien deze ongeldig is wordt de standaard dropbox locatie bepaald. 
        '--- als dat ook niet bestaat wordt de locatie dezelfde locatie als het programma

        Dim dropboxPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "\Dropbox\HulpoosteuropaTexel\SPAS2\Uitkeringen")
        Dim SelectFolder As New FolderBrowserDialog
        SelectFolder.ShowNewFolderButton = True

        ' Check if the directory stored in My.Settings._excassopath exists
        If Directory.Exists(My.Settings._excassopath) Then
            ' If it exists, set SelectFolder.SelectedPath to the stored path
            SelectFolder.SelectedPath = My.Settings._excassopath
        ElseIf Directory.Exists(dropboxPath) Then
            SelectFolder.SelectedPath = dropboxPath
        Else
            ' If it doesn't exist, default to the application's executable directory
            SelectFolder.SelectedPath = Application.StartupPath
        End If

        Dim pad As String = SelectFolder.SelectedPath & "\"
        Dim Journal_name As String = SPAS.Cmx_Excasso_Select.Text
        Dim filename As String = Journal_name
        Dim filenum As Integer = 0

        If (SelectFolder.ShowDialog() = DialogResult.OK) Then
            Do
                Try
                    filename = $"{SelectFolder.SelectedPath}\{Journal_name & filenum.ToString}.pdf"
                    ' Check if file exists without causing an error
                    If File.Exists(filename) Then
                        filenum += 1
                    Else
                        Exit Do
                    End If
                Catch
                    ' Handle unauthorized access
                    MsgBox("Unauthorized access. Please enter a new path.", vbExclamation)
                    filenum = 0
                End Try
            Loop
        End If

        '-----------------------document properties--------------------------
        Dim document As PdfDocument = New PdfDocument
        document.Info.Title = "Sponsor program form"
        Dim filenameShort = Mid(filename, Strings.Len(SelectFolder.SelectedPath) + 2)


        '-----------excasso properties ---------------------
        Dim aantal_begunstigden As Integer = SPAS.CalculateColumnCount(SPAS.Dgv_Excasso2, 9)
        Dim totalpages As Integer = Math.Ceiling(aantal_begunstigden / 14) + 1

        Dim cp_id = SPAS.Cmx_Excasso_Select.SelectedValue
        Dim Con_tot_qty As Integer = Tbx2Int(SPAS.Dgv_Excasso_numbers.Rows(0).Cells(4).Value)
        Dim Ext_tot_qty As Integer = Tbx2Dec(SPAS.Dgv_Excasso_numbers.Rows(1).Cells(4).Value) + Tbx2Dec(SPAS.Dgv_Excasso_numbers.Rows(2).Cells(4).Value) * 1
        Dim CP_tot_eur As Integer = Tbx2Dec(SPAS.Dgv_Excasso_numbers.Rows(0).Cells(8).Value) * 1 + Tbx2Dec(SPAS.Dgv_Excasso_numbers.Rows(1).Cells(8).Value) * 1 + Tbx2Dec(SPAS.Dgv_Excasso_numbers.Rows(2).Cells(8).Value) * 1
        Dim Con_tot_eur As Integer = Tbx2Dec(SPAS.Dgv_Excasso_numbers.Rows(0).Cells(5).Value)
        Dim Ext_tot_eur As Integer = Tbx2Dec(SPAS.Dgv_Excasso_numbers.Rows(1).Cells(5).Value) + Tbx2Dec(SPAS.Dgv_Excasso_numbers.Rows(2).Cells(5).Value) * 1
        Dim Gen_tot_eur = Tbx2Int(SPAS.Dgv_Excasso_numbers.Rows(2).Cells(10).Value)
        Dim xr As Decimal = Tbx2Dec(SPAS.Tbx_Excasso_Exchange_rate.Text)



        Dim page As PdfPage = document.AddPage
        Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
        Dim pen1 As XPen = New XPen(XColors.Black)
        pen1.Width = 3
        Dim pen2 As XPen = New XPen(XColors.Black)
        pen2.Width = 1
        Dim pen3 As XPen = New XPen(XColors.Black)
        pen3.DashStyle = XDashStyle.Dot
        pen3.Width = 1
        Dim pen4 As XPen = New XPen(XColors.Black)
        pen4.Width = 3

        Dim font As XFont = New XFont("Verdana", 12, XFontStyle.Regular)
        Dim font2 As XFont = New XFont("Verdana", 14, XFontStyle.Bold)
        Dim fontbold As XFont = New XFont("Verdana", 12, XFontStyle.Bold)
        Dim fontnumber As XFont = New XFont("Verdana", 12, XFontStyle.Regular)

        '
        Dim pages As Integer = 1
        Dim line As Integer = 80
        Dim Sponsored As String
        Dim Contract, Extra, Total As Integer

        Dim CP_name = QuerySQL($"Select name from CP WHERE id ='{cp_id}'")
        Dim CP_bank = QuerySQL($"SELECT bankacc.accountno FROM bankacc 
                LEFT JOIN cp on bankacc.id = cp.fk_bankacc_id
                WHERE cp.id='{cp_id}'")
        Dim dat As Date = SPAS.Dtp_Excasso_Start.Value.ToString
        Dim ximg As XImage = XImage.FromFile(Application.StartupPath & "\HOEZH3.jpg")

        gfx.DrawImage(ximg, 10, 0)
        gfx.DrawString("East Europe Support South Holland", font2, XBrushes.Black,
        New XRect(150, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString("Page " & pages & "/" & totalpages, font, XBrushes.Black, New XRect(500, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Contact person:", font, XBrushes.Black, New XRect(150, 40, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Date:", font, XBrushes.Black, New XRect(150, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("File name:", font, XBrushes.Black, New XRect(150, 80, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(CP_name, font, XBrushes.Black, New XRect(270, 40, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(dat, font, XBrushes.Black, New XRect(270, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(filenameShort, font, XBrushes.Black, New XRect(270, 80, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawLine(pen4, New XPoint(20, 140), New XPoint(560, 140))

        'horizontal
        gfx.DrawLine(pen2, New XPoint(20, 200), New XPoint(400, 200))
        gfx.DrawLine(pen2, New XPoint(20, 230), New XPoint(400, 230))
        gfx.DrawLine(pen2, New XPoint(20, 260), New XPoint(400, 260))
        gfx.DrawLine(pen2, New XPoint(20, 290), New XPoint(400, 290))
        gfx.DrawLine(pen2, New XPoint(20, 320), New XPoint(400, 320))
        gfx.DrawLine(pen2, New XPoint(20, 350), New XPoint(400, 350))
        'vertical
        gfx.DrawLine(pen2, New XPoint(20, 200), New XPoint(20, 350))
        gfx.DrawLine(pen2, New XPoint(200, 200), New XPoint(200, 350))
        gfx.DrawLine(pen2, New XPoint(260, 200), New XPoint(260, 350))
        gfx.DrawLine(pen2, New XPoint(320, 200), New XPoint(320, 350))
        gfx.DrawLine(pen2, New XPoint(400, 200), New XPoint(400, 350))

        gfx.DrawString("Summary", font2, XBrushes.Black, New XRect(20, 180, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("QTY", font, XBrushes.Black, New XRect(215, 180, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("EUR", font, XBrushes.Black, New XRect(275, 180, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("MDL", font, XBrushes.Black, New XRect(350, 180, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

        gfx.DrawString("Distribution", font, XBrushes.Black, New XRect(30, 210, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Monthly gifts", font, XBrushes.Black, New XRect(30, 240, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Extra gifts", font, XBrushes.Black, New XRect(30, 270, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)



        gfx.DrawString("General totals", font2, XBrushes.Black, New XRect(30, 330, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("1", font, XBrushes.Black, New XRect(190, 210, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Con_tot_qty, font, XBrushes.Black, New XRect(190, 240, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Ext_tot_qty, font, XBrushes.Black, New XRect(190, 270, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(CP_tot_eur * 1, font, XBrushes.Black, New XRect(260, 210, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Con_tot_eur * 1, font, XBrushes.Black, New XRect(260, 240, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Ext_tot_eur, font, XBrushes.Black, New XRect(260, 270, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Gen_tot_eur, font2, XBrushes.Black, New XRect(260, 330, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Tbx2Int(CP_tot_eur * xr), font, XBrushes.Black, New XRect(340, 210, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Tbx2Int(Con_tot_eur * xr), font, XBrushes.Black, New XRect(340, 240, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Tbx2Int(Ext_tot_eur * xr), font, XBrushes.Black, New XRect(340, 270, 50, font.Height), XStringFormats.TopRight)
        gfx.DrawString(Tbx2Int(Gen_tot_eur * xr), font2, XBrushes.Black, New XRect(340, 330, 50, font.Height), XStringFormats.TopRight)



        page = document.AddPage()
        gfx = XGraphics.FromPdfPage(page)
        pages = pages + 1

        Dim linecounter As Integer
        For x As Integer = 0 To SPAS.Dgv_Excasso2.Rows.Count - 1
            If SPAS.Dgv_Excasso2.Rows(x).Cells(6).Value <> 0 Or SPAS.Dgv_Excasso2.Rows(x).Cells(7).Value <> 0 Or SPAS.Dgv_Excasso2.Rows(x).Cells(8).Value <> 0 Then
                If linecounter Mod 14 = 0 Then
                    If linecounter > 13 Then
                        page = document.AddPage()
                        gfx = XGraphics.FromPdfPage(page)
                        pages = pages + 1
                    End If

                    gfx.DrawString("HOEZH", font2, XBrushes.Black, New XRect(20, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(Journal_name, font, XBrushes.Black, New XRect(240, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    'gfx.DrawString("Page " & pages & "/" & totalpages, font, XBrushes.Black,New XRect(500, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

                    'column headers
                    gfx.DrawString("Name", font, XBrushes.Black,
                        New XRect(20, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Monthly", font, XBrushes.Black, New XRect(175, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Extra gift", font, XBrushes.Black, New XRect(240, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Total(MLD)", font, XBrushes.Black, New XRect(315, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Signature", font, XBrushes.Black, New XRect(390, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawLine(pen1, New XPoint(20, 85), New XPoint(560, 85))
                    gfx.DrawLine(pen3, New XPoint(175, 90), New XPoint(175, 145))
                    gfx.DrawLine(pen3, New XPoint(385, 90), New XPoint(385, 145))
                    line = 65

                End If
                line = line + 50
                Sponsored = SPAS.Dgv_Excasso2.Rows(x).Cells(1).Value
                Contract = Tbx2Int(SPAS.Dgv_Excasso2.Rows(x).Cells(6).Value) * xr
                Extra = (SPAS.Dgv_Excasso2.Rows(x).Cells(7).Value + SPAS.Dgv_Excasso2.Rows(x).Cells(8).Value) * xr
                Total = SPAS.Dgv_Excasso2.Rows(x).Cells(10).Value

                gfx.DrawString(Sponsored, font, XBrushes.Black, New XRect(20, line, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                gfx.DrawString(IIf(Contract > 0, Contract, ""), fontnumber, XBrushes.Black, New XRect(200, line, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                gfx.DrawString(IIf(Extra > 0, Extra, ""), fontnumber, XBrushes.Black, New XRect(265, line, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                gfx.DrawString(Total, font, XBrushes.Black, New XRect(340, line, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                gfx.DrawLine(pen2, New XPoint(20, line + 25), New XPoint(560, line + 25))
                linecounter = linecounter + 1
                If x < SPAS.Dgv_Excasso2.Rows.Count - 1 And linecounter Mod 14 <> 0 Then
                    gfx.DrawLine(pen3, New XPoint(175, line + 25), New XPoint(175, line + 75))
                    gfx.DrawLine(pen3, New XPoint(385, line + 25), New XPoint(385, line + 75))
                End If

            End If
        Next

        document.Save($"{SelectFolder.SelectedPath}\{Journal_name & filenum.ToString}.pdf")
        MsgBox("De uitkeringslijst " & filename & " is opgeslagen.")
        Process.Start(filename)
        My.Settings._excassopath = SelectFolder.SelectedPath


    End Sub


End Module
