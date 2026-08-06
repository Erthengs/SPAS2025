Imports System.Xml.Linq
Imports System.Data

Public Class SepaParser

    ''' <summary>
    ''' Leest een SEPA pain.008 XML bestand en zet dit om naar een DataTable.
    ''' </summary>
    ''' <param name="filePath">Het pad naar het XML bestand.</param>
    ''' <returns>Een DataTable met de transacties.</returns>
    Public Function ConvertSepaXmlToDataTable(filePath As String) As DataTable
        ' 1. Maak de tabelstructuur aan
        Dim dt As New DataTable("IncassoOverzicht")
        dt.Columns.Add("Incassodatum", GetType(DateTime))
        dt.Columns.Add("Naam Debiteur", GetType(String))
        dt.Columns.Add("IBAN", GetType(String))
        dt.Columns.Add("Bedrag", GetType(Decimal))
        dt.Columns.Add("Mandaatcode", GetType(String))
        dt.Columns.Add("Omschrijving", GetType(String))
        dt.Columns.Add("EndToEndId", GetType(String))

        Try
            ' 2. Laad de XML
            Dim doc As XDocument = XDocument.Load(filePath)

            ' 3. Haal de Namespace op (Cruciaal voor SEPA bestanden!)
            ' De root heeft vaak iets als xmlns="urn:iso:std:iso:20022:tech:xsd:pain.008..."
            Dim ns As XNamespace = doc.Root.GetDefaultNamespace()

            ' 4. LINQ Query om de data te "flatten" (PmtInf -> DrctDbtTxInf)
            ' We loopen door elke PaymentInformation block (PmtInf)
            For Each pmtInf In doc.Descendants(ns + "PmtInf")

                ' Haal de datum op die geldt voor deze hele groep transacties
                Dim collectDateStr As String = pmtInf.Element(ns + "ReqdColltnDt")?.Value
                Dim collectDate As DateTime
                DateTime.TryParse(collectDateStr, collectDate)

                ' Loop nu door de individuele transacties binnen deze groep
                For Each tx In pmtInf.Descendants(ns + "DrctDbtTxInf")

                    Dim row As DataRow = dt.NewRow()

                    ' Vul de data (gebruik ?.Value en CStr() om crashes bij lege velden te voorkomen)
                    row("Incassodatum") = collectDate

                    ' Bedrag (attribuut Ccy is currency, Value is het bedrag)
                    Dim amtElement = tx.Element(ns + "InstdAmt")
                    If amtElement IsNot Nothing Then
                        row("Bedrag") = Decimal.Parse(amtElement.Value, System.Globalization.CultureInfo.InvariantCulture)
                    Else
                        row("Bedrag") = 0
                    End If

                    ' Naam & IBAN
                    row("Naam Debiteur") = CStr(tx.Element(ns + "Dbtr")?.Element(ns + "Nm")?.Value)
                    row("IBAN") = CStr(tx.Element(ns + "DbtrAcct")?.Element(ns + "Id")?.Element(ns + "IBAN")?.Value)

                    ' Mandaat & ID
                    row("Mandaatcode") = CStr(tx.Element(ns + "DrctDbtTx")?.Element(ns + "MndtRltdInf")?.Element(ns + "MndtId")?.Value)
                    row("EndToEndId") = CStr(tx.Element(ns + "PmtId")?.Element(ns + "EndToEndId")?.Value)

                    ' Omschrijving (Unstructured)
                    row("Omschrijving") = CStr(tx.Element(ns + "RmtInf")?.Element(ns + "Ustrd")?.Value)

                    dt.Rows.Add(row)
                Next
            Next

        Catch ex As Exception
            ' Log de error of gooi hem omhoog naar je UI
            Throw New Exception("Fout bij het inlezen van het SEPA bestand: " & ex.Message, ex)
        End Try

        Return dt

    End Function

End Class
