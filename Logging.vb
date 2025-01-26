Imports Npgsql
Imports System.Text.RegularExpressions


Public Module LoggingModule

    ''' <param name="sql">The SQL statement to log.</param>
    ''' <param name="caller">The name of the calling procedure or method.</param>
    ''' <param name="username">The username performing the operation.</param>

    Public Sub WriteLog(ByVal sql As String, ByVal caller As String)
        Dim operationType As String = GetOperationType(sql)
        If String.IsNullOrEmpty(operationType) Then
            ' Only log INSERT, UPDATE, DELETE operations
            Exit Sub
        End If
        sql = CleanString(sql)
        Dim dbName As String = GetTableName(sql)
        Dim records As List(Of String) = GetOperationDetails(sql)

        ' Insert log for each record in the operation
        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()

            For Each details In records
                Dim commandText As String = "
                    INSERT INTO public.operation_logs 
                        (db_name, caller, operation_type, username, operation_details) 
                    VALUES 
                        (@db_name, @caller, @operation_type, @username, @operation_details)
                "

                Using cmd As New NpgsqlCommand(commandText, conn)
                    cmd.Parameters.AddWithValue("db_name", dbName)
                    cmd.Parameters.AddWithValue("caller", caller)
                    cmd.Parameters.AddWithValue("operation_type", operationType)
                    cmd.Parameters.AddWithValue("username", username)
                    cmd.Parameters.AddWithValue("operation_details", details)
                    cmd.ExecuteNonQuery()
                End Using
            Next
        End Using
    End Sub

    ''' <summary>
    ''' Extracts the operation type (INSERT, UPDATE, DELETE) from an SQL statement.
    ''' </summary>
    ''' <param name="sql">The SQL statement.</param>
    ''' <returns>The operation type or an empty string if not applicable.</returns>
    Private Function GetOperationType(ByVal sql As String) As String
        Dim match As Match = Regex.Match(sql, "^\s*(INSERT|UPDATE|DELETE)", RegexOptions.IgnoreCase)
        If match.Success Then
            Return match.Groups(1).Value.ToUpper()
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Extracts the table name from an SQL statement.
    ''' </summary>
    ''' <param name="sql">The SQL statement.</param>
    ''' <returns>The table name.</returns>
    Private Function GetTableName(ByVal sql As String) As String
        Dim match As Match = Regex.Match(sql, "^\s*(?:INSERT INTO|UPDATE|DELETE FROM)\s+([a-zA-Z0-9_]+)", RegexOptions.IgnoreCase)
        If match.Success Then
            Return match.Groups(1).Value
        End If
        Return "Unknown"
    End Function

    ''' <summary>
    ''' Extracts the operation details (e.g., inserted values) from an SQL statement.
    ''' For INSERT operations with multiple rows, each row is treated separately.
    ''' </summary>
    ''' <param name="sql">The SQL statement.</param>
    ''' <returns>A list of operation details.</returns>
    Private Function GetOperationDetails(ByVal sql As String) As List(Of String)
        Dim operationDetails As New List(Of String)

        ' Handle INSERT with multiple rows
        Dim insertMatch As Match = Regex.Match(sql, "INSERT INTO\s+([a-zA-Z0-9_]+)\s+\((.*?)\)\s+VALUES\s+(.*)", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
        If insertMatch.Success Then
            Dim valuesPart As String = insertMatch.Groups(3).Value
            Dim valuesMatches As MatchCollection = Regex.Matches(valuesPart, "\((.*?)\)")
            For Each match As Match In valuesMatches
                operationDetails.Add($"INSERT VALUES: {match.Groups(1).Value}")
            Next
        Else
            ' For UPDATE and DELETE, log the full SQL statement
            operationDetails.Add(sql)
        End If

        Return operationDetails
    End Function
    Public Function CleanString(ByVal input As String) As String
        ' Replace line breaks with a single space
        Dim noLineBreaks As String = input.Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")

        ' Remove unnecessary spaces (multiple spaces replaced with a single space)
        Dim cleanedString As String = Regex.Replace(noLineBreaks, "\s+", " ").Trim()

        Return cleanedString
    End Function
End Module
