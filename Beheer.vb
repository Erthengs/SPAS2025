Imports System.Data
Imports System.Windows.Forms
Imports Npgsql




Public Class AccountRepository

    ' ---------------------------------------------------------
    ' METHOD 1: The existing method for loading the TreeView
    ' ---------------------------------------------------------
    Public Shared Function GetAccountHierarchyData(searchTerm As String) As DataTable
        Dim filter As String = If(String.IsNullOrWhiteSpace(searchTerm), "", searchTerm.Trim())

        Dim sql As String = "
            SELECT g.type AS level1_id, g.type AS level1_name,
                   g.id AS level2_id, g.name AS level2_name,
                   a.id AS level3_id, a.name AS level3_name
            FROM accgroup g
            LEFT JOIN account a ON g.id = a.fk_accgroup_id AND a.active = true
            WHERE g.active = true "

        If Not String.IsNullOrEmpty(filter) Then
            sql &= $" AND (g.type ILIKE '%{filter}%' OR g.name ILIKE '%{filter}%' OR a.name ILIKE '%{filter}%') "
        End If

        sql &= " ORDER BY g.type, g.name, a.name;"

        Return Collect_data2(sql)
    End Function

    ' ---------------------------------------------------------
    ' METHOD 2: The NEW method for fetching Group Details
    ' Must be inside the class and marked 'Public Shared'
    ' ---------------------------------------------------------
    Public Shared Function GetAccountGroupDetails(groupId As String) As DataRow
        ' Fetch the specific row from the accgroup table using the provided ID
        Dim sql As String = $"SELECT description, type, subtype, active,
        (select count(j.id) from accgroup ag
        left join account ac on ac.fk_accgroup_id = ag.id 
        left join journal j on j.fk_account = ac.id) as posts
        FROM accgroup WHERE id = {groupId};"

        Dim dt As DataTable = Collect_data2(sql)

        ' Return the first row if data exists, otherwise return Nothing
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)
        End If

        Return Nothing
    End Function

    ' ---------------------------------------------------------
    ' METHOD 3: The NEW method for fetching Account Details
    ' ---------------------------------------------------------
    Public Shared Function GetAccountDetails(accountId As String) As DataRow
        ' Fetch the specific row from the account table using the provided ID
        ' Add any other columns you need to retrieve here (e.g., source, bankcode)
        Dim sql As String = $"SELECT name, description, type, active, source, bankcode, searchword, 
        (select count(id) from journal where fk_account={accountId}) As posts
        FROM account WHERE id = {accountId};"

        Dim dt As DataTable = Collect_data2(sql)

        ' Return the first row if data exists, otherwise return Nothing
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)
        End If

        Return Nothing
    End Function

    ' --- VALIDATION LOGIC ---

    ''' <summary>
    ''' Checks if a subtype is unique, excluding the current group being edited.
    ''' </summary>
    Public Shared Function IsSubtypeUnique(subtype As String, excludeId As Integer) As Boolean
        If String.IsNullOrWhiteSpace(subtype) Then Return True ' Optional, so empty is always valid

        Dim sql As String = "SELECT COUNT(id) FROM accgroup WHERE subtype ILIKE @subtype AND id <> @excludeId;"

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@subtype", subtype.Trim())
                cmd.Parameters.AddWithValue("@excludeId", excludeId)
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count = 0
            End Using
        End Using
    End Function


    ' --- SAVE OPERATIONS ---

    ''' <summary>
    ''' Validates and Upserts an AccountGroup into the database.
    ''' </summary>
    Public Shared Sub SaveAccountGroup(model As AccountGroupModel)
        ' 1. Business Validations
        If String.IsNullOrWhiteSpace(model.Name) Then Throw New ArgumentException("Naam is verplicht voor een accountgroep.")
        If String.IsNullOrWhiteSpace(model.Type) Then Throw New ArgumentException("Type is verplicht voor een accountgroep.")

        ' Ensure the subtype is unique using your existing IsSubtypeUnique method
        If Not IsSubtypeUnique(model.Subtype, model.Id) Then
            Throw New ArgumentException($"Subtype '{model.Subtype}' is al in gebruik.")
        End If

        ' 2. Upsert Logic (Insert if Id is 0, otherwise Update)
        Dim sql As String
        If model.Id = 0 Then
            sql = "INSERT INTO accgroup (name, type, subtype, description, active) VALUES (@name, @type, @subtype, @description, @active);"
        Else
            sql = "UPDATE accgroup SET name = @name, type = @type, subtype = @subtype, description = @description, active = @active WHERE id = @id;"
        End If

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@name", model.Name.Trim())
                cmd.Parameters.AddWithValue("@type", model.Type.Trim())
                cmd.Parameters.AddWithValue("@subtype", If(String.IsNullOrWhiteSpace(model.Subtype), DBNull.Value, model.Subtype.Trim()))
                cmd.Parameters.AddWithValue("@description", If(String.IsNullOrWhiteSpace(model.Description), DBNull.Value, model.Description.Trim()))
                cmd.Parameters.AddWithValue("@active", model.Active)

                If model.Id > 0 Then cmd.Parameters.AddWithValue("@id", model.Id)

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub


    ''' <summary>
    ''' Validates and Upserts an Account into the database.
    ''' </summary>
    Public Shared Sub SaveAccount(model As AccountModel)
        ' 1. Business Validations
        If String.IsNullOrWhiteSpace(model.Name) Then Throw New ArgumentException("Naam is verplicht voor een account.")
        If String.IsNullOrWhiteSpace(model.Type) Then Throw New ArgumentException("Type is verplicht voor een account.")
        If model.FkAccGroupId <= 0 Then Throw New ArgumentException("Een account moet altijd gekoppeld zijn aan een geldige accountgroep.")

        ' 2. Upsert Logic with ALL fields
        Dim sql As String
        If model.Id = 0 Then
            sql = "INSERT INTO account (name, fk_accgroup_id, type, source, f_key, active, description, bankcode, searchword) 
                   VALUES (@name, @fk_group, @type, @source, @fkey, @active, @description, @bankcode, @searchword);"
        Else
            sql = "UPDATE account 
                   SET name = @name, fk_accgroup_id = @fk_group, type = @type, source = @source, f_key = @fkey, active = @active, 
                       description = @description, bankcode = @bankcode, searchword = @searchword 
                   WHERE id = @id;"
        End If

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                ' Basic properties
                cmd.Parameters.AddWithValue("@name", model.Name.Trim())
                cmd.Parameters.AddWithValue("@fk_group", model.FkAccGroupId)
                cmd.Parameters.AddWithValue("@type", model.Type.Trim())
                cmd.Parameters.AddWithValue("@source", If(String.IsNullOrWhiteSpace(model.Source), DBNull.Value, model.Source.Trim()))
                cmd.Parameters.AddWithValue("@fkey", If(model.FKey <= 0, DBNull.Value, model.FKey))
                cmd.Parameters.AddWithValue("@active", model.Active)

                ' ---> THE MISSING PROPERTIES ARE MAPPED HERE <---
                cmd.Parameters.AddWithValue("@description", If(String.IsNullOrWhiteSpace(model.Description), DBNull.Value, model.Description.Trim()))
                cmd.Parameters.AddWithValue("@bankcode", If(String.IsNullOrWhiteSpace(model.Bankcode), DBNull.Value, model.Bankcode.Trim()))
                cmd.Parameters.AddWithValue("@searchword", If(String.IsNullOrWhiteSpace(model.Searchword), DBNull.Value, model.Searchword.Trim()))

                If model.Id > 0 Then cmd.Parameters.AddWithValue("@id", model.Id)

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Refactored automated account creation for Targets and CPs.
    ''' Safely looks up the fk_accgroup_id using a parameterized subquery.
    ''' </summary>
    Public Shared Sub CreateAutomaticAccount(source As String, name As String, targetSubtype As String, fKey As Integer, accType As String)
        Dim sql As String = "
            INSERT INTO account (name, source, type, f_key, active, fk_accgroup_id) 
            VALUES (@name, @source, @type, @fkey, true, (SELECT id FROM accgroup WHERE subtype = @subtype LIMIT 1));"

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@name", name.Trim())
                cmd.Parameters.AddWithValue("@source", source.Trim())
                cmd.Parameters.AddWithValue("@type", accType.Trim())
                cmd.Parameters.AddWithValue("@fkey", fKey)
                cmd.Parameters.AddWithValue("@subtype", targetSubtype.Trim())

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

End Class
Public Class AccountGroupModel
    Public Property Id As Integer
    Public Property Name As String
    Public Property Type As String
    Public Property Subtype As String
    Public Property Description As String
    Public Property Active As Boolean
End Class

Public Class AccountModel
    Public Property Id As Integer
    Public Property FkAccGroupId As Integer
    Public Property Name As String
    Public Property Type As String
    Public Property Source As String
    Public Property FKey As Integer
    Public Property Active As Boolean
    Public Property Description As String
    Public Property Bankcode As String
    Public Property Searchword As String
End Class







