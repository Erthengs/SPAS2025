Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Module Output


    Sub Print_Excasso_form()
        Dim document As PdfDocument = New PdfDocument
        document.Info.Title = "Sponsor program form"


        Dim page As PdfPage = document.AddPage
        Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
        Dim pen1 As XPen = New XPen(XColors.Black)
        pen1.Width = 4
        Dim pen2 As XPen = New XPen(XColors.Black)
        pen2.Width = 1
        Dim font As XFont = New XFont("Verdana", 12, XFontStyle.Regular)
        Dim font2 As XFont = New XFont("Verdana", 14, XFontStyle.Bold)
        Dim fontbold As XFont = New XFont("Verdana", 20, XFontStyle.Bold)

        Dim totalpages As Integer = Math.Ceiling((SPAS.Dgv_Excasso2.Rows.Count - 1) / 14) + 1
        Dim pages As Integer = 1
        Dim line As Integer = 80
        Dim Sponsored As String
        Dim Amt2 As Integer
        Dim Journal_name As String = SPAS.Cmx_Excasso_Select.Text
        Dim CP_name = QuerySQL("Select name from CP WHERE id ='" & SPAS.Lbl_Excasso_CPid.Text & "'")
        Dim CP_bank = QuerySQL(
                "SELECT bankacc.accountno FROM bankacc 
                LEFT JOIN cp on bankacc.id = cp.fk_bankacc_id
                WHERE cp.id='" & SPAS.Lbl_Excasso_CPid.Text & "'")
        Dim dat As Date = SPAS.Dtp_Excasso_Start.Value.ToString

        'Dim img
        Dim ximg As XImage = XImage.FromFile("C:\Users\werne\OneDrive\Pictures\Logo HOET.jpg")
        'img = Image.FromFile("C:\Users\werne\OneDrive\Pictures\Logo HOET.jpg")
        gfx.DrawImage(ximg, 10, 10)

        'gfx.DrawString(Journal_name & " / " & dat, font2, XBrushes.Black,
        'New XRect(20, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

        'header:
        gfx.DrawString("Support East Europe Texel", font2, XBrushes.Black,
        New XRect(120, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Page " & pages & "/" & totalpages, font, XBrushes.Black,
                    New XRect(500, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)


        gfx.DrawString("Contact person:", font, XBrushes.Black,
                    New XRect(120, 40, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Date:", font, XBrushes.Black,
                    New XRect(120, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Bank Account:", font, XBrushes.Black,
                    New XRect(120, 80, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Transaction name:", font, XBrushes.Black,
                    New XRect(120, 100, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

        gfx.DrawString(CP_name, font, XBrushes.Black,
                    New XRect(300, 40, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(dat, font, XBrushes.Black,
                    New XRect(300, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(CP_bank, font, XBrushes.Black,
                    New XRect(300, 80, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Journal_name, font, XBrushes.Black,
                    New XRect(300, 100, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

        gfx.DrawLine(pen1, New XPoint(20, 140), New XPoint(560, 140))

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

        gfx.DrawString("Summary", font2, XBrushes.Black,
                    New XRect(20, 180, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("QTY", font, XBrushes.Black,
                    New XRect(220, 180, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("EUR", font, XBrushes.Black,
                    New XRect(280, 180, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("MDL", font, XBrushes.Black,
                    New XRect(335, 180, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)




        gfx.DrawString("Distribution", font, XBrushes.Black,
                    New XRect(30, 210, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Contractual donations", font, XBrushes.Black,
                    New XRect(30, 240, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("Extra donations", font, XBrushes.Black,
                    New XRect(30, 270, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString("Internal donations", font, XBrushes.Black,
        'New XRect(30, 300, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString("Total persons/donations", font, XBrushes.Black,
        'New XRect(30, 330, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

        Dim Con_tot_qty As Integer = Tbx2Int(SPAS.Lbl_Excasso_Items_Contract.Text)
        Dim Ext_tot_qty As Integer = Tbx2Dec(SPAS.Lbl_Excasso_Items_Extra.Text) + Tbx2Dec(SPAS.Lbl_Excasso_Items_Intern.Text) * 1
        Dim CP_tot_eur As Integer = Tbx2Int(SPAS.Lbl_Excasso_CP_Totaal.Text)
        Dim Con_tot_eur As Integer = Tbx2Int(SPAS.Lbl_Excasso_Contractwaarde.Text)
        Dim Ext_tot_eur As Integer = Tbx2Dec(SPAS.Lbl_Excasso_Extra.Text) + Tbx2Dec(SPAS.Lbl_Excasso_Intern.Text) * 1
        Dim Gen_tot_eur = GetDouble(Tbx2Dec(SPAS.Lbl_Excasso_Totaal.Text)) + GetDouble(Tbx2Dec(SPAS.Lbl_Excasso_CP_Totaal.Text))
        Dim xr As Decimal = Tbx2Dec(SPAS.Tbx_Excasso_Exchange_rate.Text)
        'Dim mld_tot As Integer = Format(Gen_tot_eur * xr, "#,###")

        gfx.DrawString("General totals", font2, XBrushes.Black, New XRect(20, 360, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString("1", font, XBrushes.Black, New XRect(220, 210, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Con_tot_qty, font, XBrushes.Black, New XRect(220, 240, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Ext_tot_qty, font, XBrushes.Black, New XRect(220, 270, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString(SPAS.Lbl_Excasso_Items_Intern.Text, font, XBrushes.Black, New XRect(220, 300, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString(SPAS.Lbl_Excasso_Items_Totaal.Text, font, XBrushes.Black, New XRect(220, 330, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

        gfx.DrawString(CP_tot_eur, font, XBrushes.Black, New XRect(270, 210, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Con_tot_eur, font, XBrushes.Black, New XRect(270, 240, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Ext_tot_eur, font, XBrushes.Black, New XRect(270, 270, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString(SPAS.Lbl_Excasso_Intern.Text, font, XBrushes.Black, New XRect(270, 300, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString(Tbx2Dec(SPAS.Lbl_Excasso_Totaal.Text), font, XBrushes.Black,New XRect(270, 330, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Gen_tot_eur, font2, XBrushes.Black, New XRect(270, 360, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)


        gfx.DrawString(Tbx2Int(CP_tot_eur * xr), font, XBrushes.Black, New XRect(330, 210, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Tbx2Int(Con_tot_eur * xr), font, XBrushes.Black, New XRect(330, 240, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Tbx2Int(Ext_tot_eur * xr), font, XBrushes.Black, New XRect(330, 270, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString(Tbx2Int(Tbx2Int(SPAS.Lbl_Excasso_Intern.Text) * xr), font, XBrushes.Black, New XRect(330, 300, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        'gfx.DrawString(Tbx2Int(Tbx2Dec(SPAS.Lbl_Excasso_Totaal.Text) * xr), font, XBrushes.Black, New XRect(330, 330, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        gfx.DrawString(Tbx2Int(Gen_tot_eur * xr), font2, XBrushes.Black, New XRect(330, 360, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

        page = document.AddPage()
        gfx = XGraphics.FromPdfPage(page)
        pages = pages + 1
        'header

        'gfx.DrawLine(pen1, New XPoint(20, 10), New XPoint(500, 60))
        'end header

        Dim linecounter As Integer
        For x As Integer = 0 To SPAS.Dgv_Excasso2.Rows.Count - 1
            If SPAS.Dgv_Excasso2.Rows(x).Cells(6).Value <> 0 Then
                If linecounter Mod 14 = 0 Then
                    If linecounter > 13 Then
                        'If x > 13 Then
                        page = document.AddPage()
                        gfx = XGraphics.FromPdfPage(page)
                        pages = pages + 1
                    End If

                    gfx.DrawString("Texel East Europe Support - " & Journal_name, font2, XBrushes.Black,
                    New XRect(20, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Page " & pages & "/" & totalpages, font, XBrushes.Black,
                        New XRect(500, 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

                    'column headers
                    gfx.DrawString("Name", font, XBrushes.Black,
                        New XRect(20, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Amount", font, XBrushes.Black,
                        New XRect(250, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Signature", font, XBrushes.Black,
                        New XRect(400, 60, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawLine(pen1, New XPoint(20, 80), New XPoint(560, 80))
                    line = 80

                End If
                line = line + 50
                    'gfx.DrawString("This is my first pdf document", font, XBrushes.Black,
                    'New XRect(0, 0, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    Sponsored = SPAS.Dgv_Excasso2.Rows(x).Cells(1).Value
                    Amt2 = SPAS.Dgv_Excasso2.Rows(x).Cells(7).Value

                    gfx.DrawString(Sponsored, font, XBrushes.Black,
                New XRect(20, line, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("MDL " & Amt2, font, XBrushes.Black,
                New XRect(250, line, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawLine(pen2, New XPoint(20, line + 20), New XPoint(560, line + 20))
                linecounter = linecounter + 1
            End If
        Next


        Dim filename As String = Journal_name & ".pdf"

        Dim SelectFolder As New FolderBrowserDialog

        With SelectFolder
            .SelectedPath = My.Settings._excassopath
            .ShowNewFolderButton = True
        End With

        If (SelectFolder.ShowDialog() = DialogResult.OK) Then
            filename = SelectFolder.SelectedPath & "\" & filename
            document.Save(filename)
            MsgBox("De uitkeringslijst " & filename & " is opgeslagen.")
            Process.Start(filename)
            My.Settings._excassopath = SelectFolder.SelectedPath
        End If


    End Sub


End Module
