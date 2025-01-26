Imports System.ComponentModel.DataAnnotations
Imports System.Deployment.Application
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.EntityFrameworkCore.Metadata.Conventions
Imports Microsoft.EntityFrameworkCore.Update.Internal

Module acct


    Sub Fill_Cmx_Journal_List()

        Dim k As String = "", lf As String = ""
        Dim f As String = SPAS.Searchbox.Text 'SPAS.Tbx_Journal_Filter.Text
        Dim act As Boolean = (SPAS.Cbx_LifeCycle.Text = "Actief")
        Dim verwerkt As Boolean = SPAS.Cbx_Journal_Status_Verwerkt.Checked
        Dim open As Boolean = SPAS.Cbx_Journal_Status_Open.Checked
        Dim sqlstr As String

        Dim nulsaldo As String = ""

        Dim st As String = " AND (j.status "
        If open And verwerkt Then st &= "IN ('Open','Verwerkt') or j.status isnull)"
        If open And Not verwerkt Then st &= "IN ('Open') or j.status isnull)"
        If Not open And verwerkt Then st &= "IN ('Verwerkt') or j.status isnull)"
        If Not open And Not verwerkt Then st &= "Not IN ('Open','Verwerkt') or j.status isnull)"
        If SPAS.Cbx_Journal_Saldo_Open.Checked Then nulsaldo = "having sum(amt1) !=0::money "

        Load_Listview(SPAS.Lv_Journal_List, "SELECT DISTINCT name, name, date FROM journal j
                                                WHERE name ILIKE '%" & f & "%'" & st & " 
                                                ORDER BY date desc, name")


        With SPAS.Lv_Journal_List
            .Columns.Item(0).Width = 0
            .Columns(1).Text = "Naam"
            .Columns.Item(1).Width = 150
            .Columns.Item(2).Width = 100
            .Columns(2).Text = "Date"
        End With


    End Sub

    '===============================================================================================

    Sub Divide_among_targets()
        Dim cnt As Integer = 0
        Dim amt As Decimal = Tbx2Dec(SPAS.Tbx_Journal_Source_Amt.Text)

        'count number of selected target accounts
        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1
            cnt = cnt + 1
        Next

        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1
            SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value = Int(amt / cnt)
        Next
        Calculate_Journal_Booking_Data()

    End Sub

    Sub Save_Internal_Booking()
        'invoerchecks
        Dim err As String = ""
        Dim rst = SPAS.Lbl_Journal_Source_Restamt.Text
        Dim name As String = SPAS.Tbx_Journal_Name.Text
        If QuerySQL("SELECT COUNT(Distinct(name)) FROM journal WHERE name ILIKE '%" & name & "%'") > 0 Then
            name = name & "_" & DateTime.Now.Day & "." & Date.Now.Month & "." & Right(DateTime.Now.Year, 2) & ":" & DateTime.Now.Second '"Intern " & SPAS.Lbl_Journal_Source_Name.Text
        End If

        If SPAS.Dgv_Journal_Intern.RowCount = 0 Then err &= "Er is geen doelaccount geselecteerd" & vbCr
        If rst < 0 Then err = "Het te verdelen bedrag is hoger dan het saldo van de bronaccount." & vbCr
        If SPAS.Rbn_Journal_Intern.Checked = False And SPAS.Rbn_Journal_Contract.Checked = False And SPAS.Rbn_Journal_Extra.Checked = False Then
            err = "Het transactietype (intern, contract of extra) is nog niet aangegeven." & vbCr
        End If

        If err <> "" Then
            MsgBox(err)
            Exit Sub
        End If

        If rst > 0 Then
            Dim answ = MsgBox("Een bedrag ad €" & rst & " is nog niet verdeeld. Wilt u doorgaan met bewaren?", vbYesNo)
            If answ = vbNo Then Exit Sub
        End If

        Dim SQLstr As String = ""
        Dim SQLroot As String = ""

        Dim dat As Date = SPAS.Dtp_Journal_intern.Value
        Dim dat1 As String = dat.Year & "-" & dat.Month & "-" & dat.Day
        Dim src_amt As Integer '= Cur2(Tbx2Int(SPAS.Tbx_Journal_Source_Amt.Text) - rst)
        Dim desc As String = SPAS.Tbx_Journal_Description.Text
        Dim selectedItem As ComboBoxItem = TryCast(SPAS.Cmbx_Overboeking_Bron.SelectedItem, ComboBoxItem)
        Dim fka As String = selectedItem.Column1
        Dim type As String = IIf(SPAS.Rbn_Journal_Intern.Checked, "'Internal'", IIf(SPAS.Rbn_Journal_Contract.Checked, "'Contract'", "'Extra'"))

        'save source 
        SQLroot = "INSERT INTO journal(name,date,status,type,source,description,amt1,fk_account)
                   VALUES('" & name & "','" & dat1 & "'::date,'Verwerkt'," & type & ",'Intern','" & desc & "','"

        With SPAS.Dgv_Journal_Intern
            For i = 0 To .Rows.Count - 1
                If .Rows(i).Cells(2).Value > 0 Then
                    SQLstr &= SQLroot & Cur2(CLng(.Rows(i).Cells(2).Value)) & "','" & .Rows(i).Cells(0).Value & "');"

                End If
                'nulwaarden overslaan
                src_amt = src_amt + Cur2(CLng(SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value))
            Next i
        End With

        SQLstr &= SQLroot & -Cur2(Tbx2Int(src_amt)) & "','" & fka & "');"
        Clipboard.Clear()
        Clipboard.SetText(SQLstr)
        Try
            RunSQL(SQLstr, "NULL", "Save_Internal_Booking")
        Catch ex As Exception
            MsgBox($"Incorrecte SQL statement: {SQLstr}")
            Exit Sub
        End Try
        MsgBox("Deze interne boeking is opgeslagen met de naam " & name & ".")
        SPAS.Leeg_overboeking_scherm()


    End Sub
    Sub Add_Internal_Contract_Bookings(ByVal accnt1 As String, accnt2 As String, amt1 As Decimal,
                                       name As String, rel As String)
        'calculate the internal bookings date
        Dim d As Date, m_add As Integer, SQLstr As String
        If Strings.Len(accnt1) = 0 Then accnt1 = 0
        m_add = IIf(SPAS.Dtp_31_contract__startdate.Value.Day > 1, 1, 0)
        SQLstr = "INSERT INTO journal (amt1, fk_account,date, status,
                                    description,source,fk_relation,name,type) VALUES"

        For m = m_add + SPAS.Dtp_31_contract__startdate.Value.Month To 12

            d = CDate("01-" & m & "-" & SPAS.Dtp_Incasso_start.Value.Year)
            SQLstr &= "('" & -amt1 & "','" & accnt1 & "','" & d & "','Verwerkt','Gegenereerde interne contractboeking " & name & "','Internal','
            " & rel & "','Intern contract " & name & "','Contract'),"

            SQLstr &= "('" & amt1 & "','" & accnt2 & "','" & d & "','Verwerkt','Gegenereerde interne contractboeking" & name & "','Internal','
            " & rel & "','Intern contract " & name & "','Contract'),"

        Next m
        SQLstr = Strings.Left(SQLstr, Strings.Len(SQLstr) - 1)  'remove last comma

        'Clipboard.Clear()
        'Clipboard.SetText(SQLstr)
        RunSQL(SQLstr, "NULL", "Add_Internal_Contract_Bookings")

    End Sub


    Sub Calculate_Journal_Booking_Data()
        'calculate values of target accounts
        Dim tgt_tot As Decimal = 0
        Dim tname As String
        Dim selectedItem As ComboBoxItem = TryCast(SPAS.Cmbx_Overboeking_Bron.SelectedItem, ComboBoxItem)
        If SPAS.Cmbx_Overboeking_Bron.SelectedIndex = -1 Then
            MsgBox("Selecteer eerst een bronaccount.")
            Exit Sub
        End If

        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1
            tgt_tot = tgt_tot + SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value
            tname = SPAS.Dgv_Journal_Intern.Rows(i).Cells(1).Value
        Next

        SPAS.Lbl_Journal_Source_Restamt.Text = Tbx2Dec(Tbx2Dec(SPAS.Tbx_Journal_Source_Amt.Text) - tgt_tot)


        If SPAS.Dgv_Journal_Intern.Rows.Count > 1 Then
            SPAS.Tbx_Journal_Name.Text = selectedItem.Column2 & ">" & tname & "+" & SPAS.Dgv_Journal_Intern.Rows.Count - 1
        Else
            SPAS.Tbx_Journal_Name.Text = selectedItem.Column2 & ">" & tname
        End If


    End Sub


    Function Create_Journal_SQL()

        Dim i As Integer
        Dim id
        Dim dat
        Dim tbl, SQL_Where, SQLStr, dateselect As String
        Dim stat As String = ""
        SQL_Where = ""
        SQLStr = ""
        tbl = "j.name"

        If SPAS.Cbx_Journal_Status_Open.Checked And Not SPAS.Cbx_Journal_Status_Verwerkt.Checked Then
            stat = " AND j.status = 'Open' "
        ElseIf Not SPAS.Cbx_Journal_Status_Open.Checked And SPAS.Cbx_Journal_Status_Verwerkt.Checked Then
            stat = " AND j.status = 'Verwerkt' "
        End If


        SQLStr = "
             SELECT j.date::date As Datum, j.name As Name, 
             CASE 
 	            WHEN j.amt1::decimal > 0.00 THEN j.amt1
 	            WHEN j.amt1::decimal < 0.00 THEN '0'
             END As Bij,
              CASE 
 	            WHEN j.amt1::decimal < 0.00 THEN j.amt1::decimal * -1
 	            WHEN j.amt1::decimal > 0.00 THEN '0'
             END As Af,
             TRIM(j.description) As Omschrijving,
             ac.name As Account, 
             j.status As Status,
             j.source As Bron,
             j.iban As IBAN,
            j.type As Soort,
             --,cp.name As CP
             --,r.name||','||r.name_add As Relatie,
             --,b.name||'/'||b.description||b.batchid As Bankinfo
             j.id As Id,
             j.fk_account As Accountnr, j.cpinfo As CPinfo, j.fk_relation AS relatie, cp.name As cp
             ,CASE 
 	 	        WHEN j.amt2 is distinct from null and j.amt2 !=0::money THEN j.amt2::decimal/j.amt1::decimal
                WHEN j.amt2 is not distinct from null or j.amt2=0::money THEN 0 END As Rate
             ,j.fk_bank As Banktransaction
             ,r.name||','||r.name_add As Relatienaam
             FROM journal j 
             LEFT JOIN account ac ON j.fk_account = ac.id
             LEFT JOIN relation r ON j.fk_relation = r.id
             LEFT JOIN target ta ON ta.id = ac.id  
             LEFT JOIN bank b ON b.id = j.fk_bank
             LEFT JOIN cp ON cp.id = ta.fk_cp_id
             "


        'Dim saldo As Decimal = 0.00
        For i = 0 To SPAS.Lv_Journal_List.Items.Count - 1
            With SPAS.Lv_Journal_List.Items(i)

                If (.Selected) Then
                    id = SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(0).Text

                    Dim _dat As Date
                    _dat = SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(2).Text
                    dat = _dat.Year & "-" & _dat.Month & "-" & _dat.Day
                    dateselect = " and j.date ='" & dat & "'::date"

                    SQL_Where &= IIf(SQL_Where = "", " WHERE ", " OR ") & tbl & "='" & id & "'" & stat & dateselect
                End If

            End With
        Next
        SQLStr &= SQL_Where & " ORDER BY j.date, j.name"
        ToClipboard(SQLStr, True)

        Return SQLStr

    End Function


    Sub Fill_Journal_List_journaalposten()

        ' Dim cred, deb As Decimal

        Load_Datagridview(SPAS.Dgv_journaalposten, Create_Journal_SQL, "Fill_Journal_List_journaalposten")

        'End If
        With SPAS.Dgv_journaalposten

            .Columns(0).Width = 60
            .Columns(0).HeaderText = "Datum"
            .Columns(0).DefaultCellStyle.Format = "dd-MM"
            .Columns(0).ReadOnly = True
            .Columns(0).Visible = False

            .Columns(1).Width = 160
            .Columns(1).HeaderText = "Naam"
            .Columns(1).Visible = False

            .Columns(2).Width = 60
            .Columns(2).ReadOnly = False
            .Columns(2).HeaderText = "Bij"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(2).DefaultCellStyle.Format = "N2"
            .Columns(2).DefaultCellStyle.ForeColor = Color.Blue

            .Columns(3).Width = 60
            .Columns(3).ReadOnly = False
            .Columns(3).HeaderText = "Af"
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(3).DefaultCellStyle.Format = "N2"
            .Columns(3).DefaultCellStyle.ForeColor = Color.Blue

            .Columns(4).Width = 250
            .Columns(4).HeaderText = "Omschrijving"
            .Columns(4).DefaultCellStyle.ForeColor = Color.Blue

            .Columns(5).Visible = True
            .Columns(5).Width = 160
            .Columns(5).HeaderText = "Account"
            .Columns(5).ReadOnly = True


            For i = 6 To 16
                .Columns(i).Visible = True
            Next
            .Columns(10).Visible = True
            .Columns(17).Visible = True
            .Columns(17).Width = 160
            .Columns(17).HeaderText = "Relatie"
            .Columns(17).ReadOnly = True


        End With
        SPAS.Calculate_Journaalposten_totalen(SPAS.Dgv_journaalposten)

    End Sub
    Sub Calculate_Budget(ByVal id As String)

        Dim SQLStr As String = ""
        Dim sd As String
        Dim mon As String = ""
        Dim where As String = ""
        If id <> "" Then where = "WHERE ac1.id=" & id


        For m = 1 To 12
            sd = Date.Today.Year & "-" & m & "-01"            'ed = CDate("01-" & m + 1 & "-" & Date.Now.Year).AddDays(-1)
            Select Case m
                Case 1 : mon = "b_jan"
                Case 2 : mon = "b_feb"
                Case 3 : mon = "b_mar"
                Case 4 : mon = "b_apr"
                Case 5 : mon = "b_may"
                Case 6 : mon = "b_jun"
                Case 7 : mon = "b_jul"
                Case 8 : mon = "b_aug"
                Case 9 : mon = "b_sep"
                Case 10 : mon = "b_oct"
                Case 11 : mon = "b_nov"
                Case 12 : mon = "b_dec"
            End Select

            SQLStr &= "
                    UPDATE account ac1
                    SET " & mon & "=
                    (
                    SELECT  sum(co.donation/co.term)
                    FROM contract co
                    LEFT JOIN target ta ON co.fk_target_id = ta.id
                    LEFT JOIN account ac ON ac.f_key = ta.id
                    WHERE co.startdate <='" & sd & "'
                    AND co.enddate > '" & sd & "'
                    AND ac1.f_key = ta.id
                    )
                    " & where & ";
"

        Next

        'MsgBox(SQLStr)
        Clipboard.Clear()
        Clipboard.SetText(SQLStr)
        RunSQL(SQLStr, "NULL", "Calculate_Budget")
        RunSQL(Budget_Year_Totals, "NULL", "Calculate_Budget/Budget_Year_Totals")

    End Sub

    Function Budget_Year_Totals()
        Dim SQLStr As String = "
                    UPDATE account
                    SET b_year=
                    (
                    Select 
							(CASE
								WHEN b_jan IS NOT NULL Then b_jan
								WHEN b_jan IS NULL Then '0'
							END) +
							(CASE
								WHEN b_feb IS NOT NULL Then b_feb
								WHEN b_feb IS NULL Then '0'
							END) +
							(CASE
								WHEN b_mar IS NOT NULL Then b_mar
								WHEN b_mar IS NULL Then '0'
							END) +
							(CASE
								WHEN b_apr IS NOT NULL Then b_apr
								WHEN b_apr IS NULL Then '0'
							END)+
							(CASE
								WHEN b_may IS NOT NULL Then b_may
								WHEN b_may IS NULL Then '0'
							END) +
							(CASE
								WHEN b_jun IS NOT NULL Then b_jun
								WHEN b_jun IS NULL Then '0'
							END) +
							(CASE
								WHEN b_jul IS NOT NULL Then b_jul
								WHEN b_jul IS NULL Then '0'
							END) +
							(CASE
								WHEN b_aug IS NOT NULL Then b_aug
								WHEN b_aug IS NULL Then '0'
							END) +
							(CASE
								WHEN b_sep IS NOT NULL Then b_sep
								WHEN b_sep IS NULL Then '0'
							END) +
							(CASE
								WHEN b_oct IS NOT NULL Then b_oct
								WHEN b_oct IS NULL Then '0'
							END) +
							(CASE
								WHEN b_nov IS NOT NULL Then b_nov
								WHEN b_nov IS NULL Then '0'
							END) +
							(CASE
								WHEN b_dec IS NOT NULL Then b_dec
								WHEN b_dec IS NULL Then '0'
							END)
                            ); 

"
        Return SQLStr


    End Function

    Sub Calculate_Manual_Budgets()


        SPAS.Lbl_Account_Budget_Difference.Text = Tbx2Dec(
                    GetDouble(SPAS.Tbx_10_Account__b_year.Text) -
                    (
                    GetDouble(SPAS.Tbx_10_Account__b_jan.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_feb.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_mar.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_apr.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_may.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_jun.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_jul.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_aug.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_sep.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_oct.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_nov.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_dec.Text))
        )


    End Sub

    Sub Load_Account_Settings()

        Load_Datagridview(SPAS.Dgv_Settings, QuerySQL("select sql from query where name='Haal settings op'"), "load account settings")
        Dim formatting As String = QuerySQL("select formatting from query where name='Haal settings op'")
        Dim arr_format() As String
        If Not IsNothing(formatting) Then arr_format = formatting.Split(",")
        SPAS.Format_Datagridview(SPAS.Dgv_Settings, arr_format, False)

        With SPAS.Dgv_Settings
            Dim unit As String = ""
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            .DefaultCellStyle.WrapMode = DataGridViewTriState.True
            For r As Integer = 0 To .Rows.Count - 1
                unit = .Rows(r).Cells(1).Value
                Select Case unit
                    Case "text"
                        .Rows(r).Height = 50
                        .Rows(r).Cells(2).ReadOnly = False
                        .Rows(r).Cells(2).Style.ForeColor = Color.Blue
                    Case "currency"
                        .Rows(r).Cells(2).Style.Format = "N2"
                        .Rows(r).Cells(2).ReadOnly = False
                        .Rows(r).Cells(2).Style.ForeColor = Color.Blue
                    Case Else
                        .Rows(r).Cells(2).ReadOnly = True

                End Select
            Next
        End With



    End Sub
    Sub Save_Settings()

        Dim sql As String = ""
        With SPAS.Dgv_Settings
            If .SelectedCells.Count = 0 Then Exit Sub
            .Rows(.SelectedCells(0).RowIndex).Cells(2).Value = .CurrentCell.EditedFormattedValue
            For r As Integer = 0 To .Rows.Count - 1
                If .Rows(r).Cells(1).Value = "text" Or .Rows(r).Cells(1).Value = "currency" Then
                    sql &= "Update settings set value ='" & .Rows(r).Cells(2).Value & "' where label = '" & .Rows(r).Cells(0).Value & "';" & vbCr
                End If
            Next
        End With
        RunSQL(sql, "NULL", "Save settings")

        ToClipboard(sql, True)


    End Sub

    Sub Select_in_Lv_Journal_list_ByNameAndDate(name As String, dateValue As String, idx As String, fltr As String)
        ' Loop through each item in the ListView
        Dim dat_ As Date
        Dim dat As String

        SPAS.TC_Boeking.SelectedIndex = idx
        For Each item As ListViewItem In SPAS.Lv_Journal_List.Items
            dat_ = item.SubItems(2).Text
            dat = dat_.Year & "-" & dat_.Month & "-" & dat_.Day
            If item.SubItems(0).Text = name AndAlso dat = dateValue Then
                item.Selected = True
                item.EnsureVisible()

                Exit For
            End If
        Next
    End Sub

    Sub SelectRowById(dgv As DataGridView, idToFind As Integer)
        ' Loop through all rows in the DataGridView
        For Each row As DataGridViewRow In dgv.Rows
            ' Check if the value in the "IdColumn" matches the idToFind
            If row.Cells("Id").Value IsNot Nothing AndAlso row.Cells("Id").Value = idToFind Then
                ' Select the row
                row.Selected = True
                dgv.FirstDisplayedScrollingRowIndex = row.Index

                ' Exit the loop once the row is found
                Exit For
            End If
        Next
    End Sub

    Sub Close_Year()
        '----------- uitvoeren controles

        If Check_administratie() = False Then Exit Sub
        If MsgBox("Wilt u het jaar " & report_year & " definitief afsluiten? Dit kan niet meer worden teruggedraaid!", vbYesNo) = vbNo Then Exit Sub

        'Ophalen transitieposten ten behoeve van table t1
        Dim Sql1 As String = QuerySQL("Select sql from query where category = 'Overzicht' and name='Transitieposten'")
        Sql1 = Sql1.Replace("[Year]", report_year)
        Dim Sql2 As String = QuerySQL("Select sql from query where category = 'Transaction' and name='Jaarafsluiting'")
        Sql2 = Sql2.Replace("[Year]", report_year)
        Dim Sql As String = Sql1 & vbCr & Sql2

        'Uitvoeren jaarafsluiting
        Clipboard.Clear()
        Clipboard.SetText(Sql)

        RunSQL(Sql, "NULL", "Close_Year")

        If MsgBox("Wilt u de budgetten voor " & Now.Year & " berekenen (eventuele handmatige aanpassen gaan verloren)? ", vbYesNo) = vbYes Then
            Calculate_Budget("")
        End If
    End Sub

    Function Check_administratie()


        Collect_data(QuerySQL("Select sql from query where name= 'Haal checks op'"))
        Dim result As VariantType
        Dim response As String = "Uitkomsten controles:" & vbCr
        Dim Res As Boolean = True


        For c = 0 To dst.Tables(0).Rows.Count - 1
            'dst.Tables(0).Rows(c)(1) = dst.Tables(0).Rows(c)(1).Replace("p1", Bank_table(rep_year))
            'dst.Tables(0).Rows(c)(1) = dst.Tables(0).Rows(c)(1).Replace("p2", Report_table(rep_year))
            dst.Tables(0).Rows(c)(1) = dst.Tables(0).Rows(c)(1).Replace("[year]", report_year)
            Debug.Print(dst.Tables(0).Rows(c)(1))
            'result = QuerySQL(dst.Tables(0).Rows(c)(1))

            If Not IsDBNull(QuerySQL(dst.Tables(0).Rows(c)(1))) Then

                result = QuerySQL(dst.Tables(0).Rows(c)(1))
                If result <> 0 Then
                    dst.Tables(0).Rows(c)(2) = dst.Tables(0).Rows(c)(2).Replace("#", result)
                    dst.Tables(0).Rows(c)(2) = "- FOUT: " & dst.Tables(0).Rows(c)(2)
                    response &= dst.Tables(0).Rows(c)(2) & vbCr
                    Res = False

                End If
            End If
            If Strings.Left(dst.Tables(0).Rows(c)(2), 6) <> "- FOUT" Then
                dst.Tables(0).Rows(c)(3) = "- " & dst.Tables(0).Rows(c)(3) & " OK"
                response &= dst.Tables(0).Rows(c)(3) & vbCr
            End If

        Next c

        MsgBox(response, IIf(Res, vbInformation, vbCritical))
        If Res Then Return True Else Return False

    End Function
End Module
