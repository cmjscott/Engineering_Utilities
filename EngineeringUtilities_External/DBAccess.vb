

'Component Name: DBAccess
'Description: All the database access is done through DBAccess Component 
'Author:  Srinivasan Subbanchattiar 
'Created Date: 01/04/2006 
'Modified Date: 01/04/2006
'Version: 4 


Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class DBAccess

    Public Sub New()
    End Sub

    Private dbCs As String
    Private dbCmdTimeout As Integer

    Public Sub New(ByVal ConnectionString As String)
        'Allows us to use a dbCS String Other then the Default
        dbCs = ConnectionString
    End Sub

    Public Function GetConnectionString() As String

        dbCs = Configuration.Configure
        'Configuration.ConfigurationManager.ConnectionStrings("PALDBConnectionString").ConnectionString

        GetConnectionString = dbCs

    End Function

    Public Function GetCommandTimeOut() As Integer

        dbCmdTimeout = Configuration.WebConfigurationManager.ConnectionStrings("PALCMDTimeOut").ConnectionString

        GetCommandTimeOut = dbCmdTimeout

    End Function

    Protected Function GetConnection() As SqlConnection

        Dim dbCon As New SqlConnection

        If dbCs = String.Empty Then
            dbCs = GetConnectionString()
        End If

        dbCon = New SqlConnection(dbCs)
        dbCon.Open()
        GetConnection = dbCon

    End Function

    Protected Sub CloseConnection(ByVal dbCon As SqlConnection)

        dbCon.Close()
        dbCon = Nothing

    End Sub

    Public Function RunPassSQL(ByVal strSQL As String) As SqlDataReader

        Dim dbCon As SqlConnection = GetConnection()
        Dim dbRs As SqlDataReader
        Dim dbCmd As New SqlCommand(strSQL, dbCon)

        dbRs = dbCmd.ExecuteReader(CommandBehavior.CloseConnection)

        dbCmd.Dispose()

        Return dbRs

    End Function

    Public Sub RunActionQuery(ByVal strSQL As String)

        Dim dbCon As SqlConnection = GetConnection()
        Dim dbCmd As New SqlCommand(strSQL, dbCon)

        Try
            dbCmd.ExecuteNonQuery()
            dbCmd.Dispose()
        Finally
            CloseConnection(dbCon)
        End Try

    End Sub

    Public Overloads Function RunSQLReturnDataSet(ByVal strSQL As String) As DataSet

        Dim dbCon As SqlConnection = GetConnection()

        Dim ds As New DataSet
        Dim da As New SqlDataAdapter(strSQL, dbCon)

        da.SelectCommand.CommandType = CommandType.Text

        da.Fill(ds)

        CloseConnection(dbCon)

        da.Dispose()

        Return ds

    End Function

    Public Function RunSPReturnInteger(ByVal strSP As String, ByVal ParamArray commandParameters() As SqlParameter) As Integer

        Dim dbCon As SqlConnection = GetConnection()
        Dim retVal As Integer

        Try

            Dim dbCmd As New SqlCommand(strSP, dbCon)
            dbCmd.CommandTimeout = GetCommandTimeOut()
            dbCmd.CommandType = CommandType.StoredProcedure

            Dim para As SqlParameter
            For Each para In commandParameters
                para = dbCmd.Parameters.Add(para)
                para.Direction = ParameterDirection.Input
            Next

            para = dbCmd.Parameters.Add(New SqlParameter("@RetVal", SqlDbType.Int))
            para.Direction = ParameterDirection.Output

            dbCmd.ExecuteNonQuery()
            retVal = dbCmd.Parameters("@RetVal").Value()
            dbCmd.Dispose()

        Finally
            CloseConnection(dbCon)
        End Try

        Return retVal

    End Function

    Public Function RunSPReturnId(ByVal strSP As String, ByVal ParamArray commandParameters() As SqlParameter) As String

        Dim dbCon As SqlConnection = GetConnection()
        Dim retVal As String

        Try

            Dim dbCmd As New SqlCommand(strSP, dbCon)
            dbCmd.CommandTimeout = GetCommandTimeOut()
            dbCmd.CommandType = CommandType.StoredProcedure

            Dim para As SqlParameter
            For Each para In commandParameters
                para = dbCmd.Parameters.Add(para)
                para.Direction = ParameterDirection.Input
            Next

            para = dbCmd.Parameters.Add(New SqlParameter("@RetVal", SqlDbType.VarChar, 200))
            para.Direction = ParameterDirection.Output

            dbCmd.ExecuteNonQuery()
            retVal = dbCmd.Parameters("@RetVal").Value
            dbCmd.Dispose()

        Finally
            CloseConnection(dbCon)
        End Try

        Return retVal

    End Function

    Public Overloads Function RunSPReturnDataSet(ByVal strSP As String) As DataSet

        Dim dbCon As SqlConnection = GetConnection()

        Dim ds As New DataSet
        Dim da As New SqlDataAdapter(strSP, dbCon)

        da.SelectCommand.CommandType = CommandType.StoredProcedure

        da.Fill(ds)

        CloseConnection(dbCon)

        da.Dispose()

        Return ds

    End Function

    Public Overloads Function RunSPReturnDataSet(ByVal strSP As String, ByVal ParamArray commandParameters() As SqlParameter) As DataSet

        Dim dbCon As SqlConnection = GetConnection()

        Dim ds As New DataSet
        Dim da As New SqlDataAdapter(strSP, dbCon)

        da.SelectCommand.CommandType = CommandType.StoredProcedure

        Dim para As SqlParameter

        For Each para In commandParameters
            da.SelectCommand.Parameters.Add(para)
            para.Direction = ParameterDirection.Input
        Next

        da.Fill(ds)

        CloseConnection(dbCon)

        da.Dispose()

        Return ds

    End Function

    Public Overloads Function RunSPReturnDataSet(ByVal strSP As String, ByVal DataTableName As String, ByVal ParamArray commandParameters() As SqlParameter) As DataSet

        Dim dbCon As SqlConnection = GetConnection()

        Dim ds As New DataSet
        Dim da As New SqlDataAdapter(strSP, dbCon)

        da.SelectCommand.CommandType = CommandType.StoredProcedure

        Dim para As SqlParameter

        For Each para In commandParameters
            da.SelectCommand.Parameters.Add(para)
            para.Direction = ParameterDirection.Input
        Next

        da.Fill(ds, DataTableName)

        CloseConnection(dbCon)

        da.Dispose()

        Return ds

    End Function

    Public Overloads Function RunSPReturnRS(ByVal strSP As String, ByVal ParamArray commandParameters() As SqlParameter) As SqlDataReader

        Dim dbCon As SqlConnection = GetConnection()
        Dim dbRs As SqlDataReader

        Dim dbCmd As New SqlCommand(strSP, dbCon)
        dbCmd.CommandTimeout = GetCommandTimeOut()
        dbCmd.CommandType = CommandType.StoredProcedure

        Dim para As SqlParameter
        For Each para In commandParameters
            para = dbCmd.Parameters.Add(para)
            para.Direction = ParameterDirection.Input
        Next

        dbRs = dbCmd.ExecuteReader(CommandBehavior.CloseConnection)

        dbCmd.Dispose()

        Return dbRs

    End Function


    Public Overloads Function RunSPReturnRS(ByVal strSP As String) As SqlDataReader

        Dim dbCon As SqlConnection = GetConnection()
        Dim dbRs As SqlDataReader

        Dim dbCmd As New SqlCommand(strSP, dbCon)
        dbCmd.CommandTimeout = GetCommandTimeOut()
        dbCmd.CommandType = CommandType.StoredProcedure

        dbRs = dbCmd.ExecuteReader(CommandBehavior.CloseConnection)
        dbCmd.Dispose()

        Return dbRs

    End Function


    Public Overloads Function RunSQLReturnRS(ByVal strSQL As String, ByVal ParamArray commandParameters() As SqlParameter) As SqlDataReader

        Dim dbCon As SqlConnection = GetConnection()
        Dim dbRs As SqlDataReader

        Dim dbCmd As New SqlCommand(strSQL, dbCon)
        dbCmd.CommandTimeout = GetCommandTimeOut()
        dbCmd.CommandType = CommandType.Text

        Dim para As SqlParameter
        For Each para In commandParameters
            para = dbCmd.Parameters.Add(para)
            para.Direction = ParameterDirection.Input
        Next

        dbRs = dbCmd.ExecuteReader(CommandBehavior.CloseConnection)

        dbCmd.Dispose()

        Return dbRs

    End Function

End Class



