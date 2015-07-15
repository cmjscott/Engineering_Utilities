
'Component Name: PAL_Report
'Description: Used to Create, Update, View and Delete PAL Report Information
'Author:  Srinivasan Subbanchattiar 
'Created By: Srinivasan Subbanchattiar 
'Created Date: 05/17/2012 
'Modified By: Srinivasan Subbanchattiar 
'Modified Date: 05/17/2012  
'Version: 1.0

Imports System.Web
Imports System.Data
Imports System.Data.SqlClient

Public Class PAL_Report

    Protected intErr As Integer
    Protected strErr As String
    Protected strBy As String

    Protected strReportId As String
    Protected strReportName As String
    Protected strReportQuery As String
    Protected strReportDescription As String
    Protected strActiveReport As String
    Protected strReportComments As String

    Protected strFilter As String

    Protected strUserId As String

    Protected intRoleId As Integer

    Protected gvData As GridView
    Protected dsData As SqlDataSource

    'for view purpose only
    Protected strCreatedBy As String
    Protected dtCreatedDate As Date
    Protected strModifiedBy As String
    Protected dtModifiedDate As Date


    Public Sub New()

        Clear()

    End Sub


    Public Sub Clear()

        intErr = Nothing
        strErr = Nothing
        strBy = Nothing

        strReportId = Nothing
        strReportName = Nothing
        strReportQuery = Nothing
        strReportDescription = Nothing
        strActiveReport = Nothing
        strReportComments = Nothing

        strUserId = Nothing
        intRoleId = Nothing

        strFilter = Nothing

        gvData = Nothing
        dsData = Nothing

        strCreatedBy = Nothing
        dtCreatedDate = Nothing
        strModifiedBy = Nothing
        dtModifiedDate = Nothing

    End Sub

    Public Property Error_Id() As Integer
        Get
            Return intErr
        End Get
        Set(ByVal Value As Integer)
            intErr = Value
        End Set
    End Property

    Public Property Error_Message() As String
        Get
            Return strErr
        End Get
        Set(ByVal Value As String)
            strErr = Value
        End Set
    End Property

    Public Property Report_Id() As String
        Get
            Return strReportId
        End Get
        Set(ByVal Value As String)
            strReportId = Value
        End Set
    End Property

    Public Property Report_Name() As String
        Get
            Return strReportName
        End Get
        Set(ByVal Value As String)
            strReportName = Value
        End Set
    End Property

    Public Property Report_Query() As String
        Get
            Return strReportQuery
        End Get
        Set(ByVal Value As String)
            strReportQuery = Value
        End Set
    End Property


    Public Property Report_Description() As String
        Get
            Return strReportDescription
        End Get
        Set(ByVal Value As String)
            strReportDescription = Value
        End Set
    End Property


    Public Property Active_Report() As String
        Get
            Return strActiveReport
        End Get
        Set(ByVal Value As String)
            strActiveReport = Value
        End Set
    End Property

    Public Property Report_Comments() As String
        Get
            Return strReportComments
        End Get
        Set(ByVal Value As String)
            strReportComments = Value
        End Set
    End Property

    Public Property Filter() As String
        Get
            Return strFilter
        End Get
        Set(ByVal Value As String)
            strFilter = Value
        End Set
    End Property

    Public Property User_Id() As String
        Get
            Return strUserId
        End Get
        Set(ByVal Value As String)
            strUserId = Value
        End Set
    End Property

    Public Property Role_Id() As Integer
        Get
            Return intRoleId
        End Get
        Set(ByVal Value As Integer)
            intRoleId = Value
        End Set
    End Property

    Public Property Created_By() As String
        Get
            Return strCreatedBy
        End Get
        Set(ByVal Value As String)
            strCreatedBy = Value
        End Set
    End Property

    Public Property Modified_By() As String
        Get
            Return strModifiedBy
        End Get
        Set(ByVal Value As String)
            strModifiedBy = Value
        End Set
    End Property

    Public Property Created_Date() As Date
        Get
            Return dtCreatedDate
        End Get
        Set(ByVal Value As Date)
            dtCreatedDate = Value
        End Set
    End Property

    Public Property Modified_Date() As Date
        Get
            Return dtModifiedDate
        End Get
        Set(ByVal Value As Date)
            dtModifiedDate = Value
        End Set
    End Property

    Public Property DS_Data() As SqlDataSource
        Get
            Return dsData
        End Get
        Set(ByVal Value As SqlDataSource)
            dsData = Value
        End Set
    End Property

    Public Property GV_Data() As GridView
        Get
            Return gvData
        End Get
        Set(ByVal Value As GridView)
            gvData = Value
        End Set
    End Property

    Public Property By() As String
        Get
            Return strBy
        End Get
        Set(ByVal Value As String)
            strBy = Value
        End Set
    End Property

    Public Sub selectAllReports()

        Dim dbCon As New DBAccess

        'Get all the User information from the database

        DS_Data.SelectCommand = "dbo.sp_get_all_reports"
        DS_Data.SelectCommandType = 1
        DS_Data.ConnectionString = dbCon.GetConnectionString()
        DS_Data.DataBind()

        dbCon = Nothing

    End Sub

    Public Sub selectAllReportsbyUser()

        Dim dbCon As New DBAccess

        'Get all the Report Result information from the database

        DS_Data.SelectParameters.Clear()
        DS_Data.SelectCommand = "dbo.sp_get_all_reports_by_user"
        DS_Data.SelectParameters.Add("User_id", strBy)
        DS_Data.SelectCommandType = 1
        DS_Data.ConnectionString = dbCon.GetConnectionString()
        DS_Data.DataBind()

        dbCon = Nothing

    End Sub

    Public Sub selectAllReportsbyRoleUser()

        Dim dbCon As New DBAccess

        'Get all the Report Result information from the database

        DS_Data.SelectParameters.Clear()
        DS_Data.SelectCommand = "dbo.sp_get_all_reports_by_role_user"
        DS_Data.SelectParameters.Add("User_id", strBy)
        DS_Data.SelectCommandType = 1
        DS_Data.ConnectionString = dbCon.GetConnectionString()
        DS_Data.DataBind()

        dbCon = Nothing

    End Sub

    Public Sub selectAllUsersbyReport()

        Dim dbCon As New DBAccess

        'Get all the Report Result information from the database

        DS_Data.SelectParameters.Clear()
        DS_Data.SelectCommand = "dbo.sp_get_all_users_by_report"
        DS_Data.SelectParameters.Add("Report_id", strReportId)
        DS_Data.SelectCommandType = 1
        DS_Data.ConnectionString = dbCon.GetConnectionString()
        DS_Data.DataBind()

        dbCon = Nothing

    End Sub

    Public Sub selectAllRolesbyReport()

        Dim dbCon As New DBAccess

        'Get all the Report Result information from the database

        DS_Data.SelectParameters.Clear()
        DS_Data.SelectCommand = "dbo.sp_get_all_roles_by_report"
        DS_Data.SelectParameters.Add("Report_id", strReportId)
        DS_Data.SelectCommandType = 1
        DS_Data.ConnectionString = dbCon.GetConnectionString()
        DS_Data.DataBind()

        dbCon = Nothing

    End Sub

    Public Sub runReport()

        Dim dbCon As New DBAccess

        'Get all the Report Result information from the database

        DS_Data.SelectParameters.Clear()
        DS_Data.SelectCommand = "dbo.sp_run_report"
        DS_Data.SelectParameters.Add("Report_id", strReportId)
        DS_Data.SelectCommandType = 1
        DS_Data.ConnectionString = dbCon.GetConnectionString()
        DS_Data.DataBind()

        dbCon = Nothing

    End Sub

    Public Sub executeSelectReport()

        If Not IsDBNull(strReportId) Then

            Dim dbCon As New DBAccess
            Dim dbRs As SqlDataReader

            'Get all the user information from the database

            dbRs = dbCon.RunSPReturnRS("dbo.sp_get_report", _
                         New SqlParameter("@report_id", strReportId))

            If dbRs.Read Then


                If Not IsDBNull(dbRs("Report_name")) Then
                    strReportName = dbRs("Report_name")
                End If

                If Not IsDBNull(dbRs("Report_query")) Then
                    strReportQuery = dbRs("Report_query")
                End If

                If Not IsDBNull(dbRs("Report_description")) Then
                    strReportDescription = dbRs("Report_description")
                End If

                If Not IsDBNull(dbRs("Active_report")) Then
                    strActiveReport = dbRs("Active_report")
                End If

                If Not IsDBNull(dbRs("Report_comments")) Then
                    strReportComments = dbRs("Report_comments")
                End If

                If Not IsDBNull(dbRs("Created_by")) Then
                    strCreatedBy = dbRs("Created_by")
                End If

                If Not IsDBNull(dbRs("Modified_by")) Then
                    strModifiedBy = dbRs("Modified_by")
                End If

                If Not IsDBNull(dbRs("Created_date")) Then
                    dtCreatedDate = dbRs("Created_date")
                End If

                If Not IsDBNull(dbRs("Modified_date")) Then
                    dtModifiedDate = dbRs("Modified_date")
                End If

                intErr = 0 'Record Found
                strErr = ""

            Else

                intErr = -1 'Record NOT Found
                strErr = "Record Not Found"

            End If

            dbRs.Close()
            dbRs = Nothing
            dbCon = Nothing

        Else

            intErr = -2 'Id is Nothing
            strErr = "Report Id is Nothing"

        End If

    End Sub 'executeSelectReport()


    Public Sub executeCreateReport()

        Dim dbCon As New DBAccess
        Dim T_id As String

        'Create new report to the database 

        T_id = dbCon.RunSPReturnId("dbo.sp_create_report_wiz", _
                         New SqlParameter("@report_name", strReportName), _
                         New SqlParameter("@report_description", strReportDescription), _
                         New SqlParameter("@report_comments", strReportComments), _
                         New SqlParameter("@Created_by", strBy))

        If T_id = "-1" Then

            intErr = -1 'Create New Report Failed
            strErr = "Create New Report Failed"

        Else

            intErr = 0 'New Report Created Successfully
            strErr = "New Report Created Successfully"

        End If

        dbCon = Nothing

    End Sub 'executeCreateReport()


    Public Sub executeUpdateReport()

        Dim dbCon As New DBAccess
        Dim T_id As Integer

        'Save Report Information to the database 

        T_id = dbCon.RunSPReturnInteger("dbo.sp_update_report_wiz", _
                         New SqlParameter("@Report_id", strReportId), _
                         New SqlParameter("@Report_name", strReportName), _
                         New SqlParameter("@Report_description", strReportDescription), _
                         New SqlParameter("@Active_report", strActiveReport), _
                         New SqlParameter("@report_comments", strReportComments), _
                         New SqlParameter("@Modified_by", strBy))

        If T_id = -1 Then

            intErr = -1 'Update Report Failed
            strErr = "Update Report Failed"

        Else

            intErr = 0 'Report  Saved Successfully
            strErr = "Report  Saved Successfully"

        End If

        dbCon = Nothing

    End Sub 'executeUpdateReport()

    Public Sub executeDeleteReport()

        Dim dbCon As New DBAccess
        Dim T_id As Integer

        'Delete Report Information from the database 

        T_id = dbCon.RunSPReturnInteger("dbo.sp_delete_report_wiz", _
                         New SqlParameter("@Report_id", strReportId), _
                         New SqlParameter("@Modified_by", strBy))

        If T_id = -1 Then

            intErr = -1 'Delete Report Failed
            strErr = "Delete Report Failed"

        Else

            intErr = 0 'Report  Deleted Successfully
            strErr = "Report  Deleted Successfully"

        End If

        dbCon = Nothing

    End Sub 'executeDeletereport()


    Public Sub executeAddUser()

        Dim dbCon As New DBAccess
        Dim T_id As Integer

        'Add Report User Information to the database 

        T_id = dbCon.RunSPReturnInteger("dbo.sp_add_user_to_report", _
                         New SqlParameter("@Report_id", strReportId), _
                         New SqlParameter("@User_id", strUserId), _
                         New SqlParameter("@Modified_by", strBy))

        If T_id = "-1" Then

            intErr = -1 'Add User Failed
            strErr = "Add User Failed"

        Else

            intErr = 0 'User Added Successfully
            strErr = "User Added Successfully"

        End If

        dbCon = Nothing


    End Sub 'executeAddUser()

    Public Sub executeAddRole()

        Dim dbCon As New DBAccess
        Dim T_id As Integer

        'Add Report Role Information to the database 

        T_id = dbCon.RunSPReturnInteger("dbo.sp_add_role_to_report", _
                         New SqlParameter("@Report_id", strReportId), _
                         New SqlParameter("@Role_id", intRoleId), _
                         New SqlParameter("@Modified_by", strBy))

        If T_id = "-1" Then

            intErr = -1 'Add Role Failed
            strErr = "Add Role Failed"

        Else

            intErr = 0 'Role Added Successfully
            strErr = "Role Added Successfully"

        End If

        dbCon = Nothing


    End Sub 'executeAddRole()

    Public Sub executeRemoveUser()

        Dim dbCon As New DBAccess
        Dim T_id As Integer

        'Remove Report User from the database 

        T_id = dbCon.RunSPReturnInteger("dbo.sp_remove_user_from_report", _
                         New SqlParameter("@Report_id", strReportId), _
                         New SqlParameter("@User_id", strUserId), _
                         New SqlParameter("@Modified_by", strBy))

        If T_id = -1 Then

            intErr = -1 'Remove User Failed
            strErr = "Remove User Failed"

        Else

            intErr = 0 'User Removed Successfully
            strErr = "User Removed Successfully"

        End If

        dbCon = Nothing

    End Sub 'executeRemoveUser()

    Public Sub executeRemoveRole()

        Dim dbCon As New DBAccess
        Dim T_id As Integer

        'Remove Report Role from the database 

        T_id = dbCon.RunSPReturnInteger("dbo.sp_remove_role_from_report", _
                         New SqlParameter("@Report_id", strReportId), _
                         New SqlParameter("@Role_id", intRoleId), _
                         New SqlParameter("@Modified_by", strBy))

        If T_id = -1 Then

            intErr = -1 'Remove Role Failed
            strErr = "Remove Role Failed"

        Else

            intErr = 0 'Role Removed Successfully
            strErr = "Role Removed Successfully"

        End If

        dbCon = Nothing

    End Sub 'executeRemoveRole()

    Public Sub executeAddFilter(ByVal FieldName As String, ByVal FieldValue As String, ByVal TypeString As Boolean)

        FieldValue = Replace(FieldValue, "select", "")
        FieldValue = Replace(FieldValue, "insert", "")
        FieldValue = Replace(FieldValue, "update", "")
        FieldValue = Replace(FieldValue, "delete", "")

        If FieldValue <> "" Then

            If TypeString Then

                If strFilter <> "" Then
                    strFilter = strFilter + " AND " + FieldName + " Like '" + Replace(FieldValue, "*", "%") + "'"
                Else
                    strFilter = FieldName + " Like '" + Replace(FieldValue, "*", "%") + "'"
                End If

            Else

                If strFilter <> "" Then
                    strFilter = strFilter + " AND " + FieldName + "=" + FieldValue
                Else
                    strFilter = FieldName + "=" + FieldValue
                End If

            End If

        End If

    End Sub 'executeAddFilter()

    Public Sub executeAddFilterOR(ByVal FieldName As String, ByVal FieldValue1 As String, ByVal FieldValue2 As String, ByVal TypeString As Boolean)

        Dim T_check As Boolean = False
        Dim FieldValue As String = ""

        FieldValue1 = Replace(FieldValue1, "select", "")
        FieldValue1 = Replace(FieldValue1, "insert", "")
        FieldValue1 = Replace(FieldValue1, "update", "")
        FieldValue1 = Replace(FieldValue1, "delete", "")

        FieldValue2 = Replace(FieldValue2, "select", "")
        FieldValue2 = Replace(FieldValue2, "insert", "")
        FieldValue2 = Replace(FieldValue2, "update", "")
        FieldValue2 = Replace(FieldValue2, "delete", "")

        If FieldValue1 <> "" And FieldValue2 = "" Then
            FieldValue = FieldValue1
            T_check = True
        ElseIf FieldValue1 = "" And FieldValue2 <> "" Then
            FieldValue = FieldValue2
            T_check = True
        End If

        If T_check Then

            If TypeString Then

                If strFilter <> "" Then
                    strFilter = strFilter + " AND " + FieldName + " Like '" + Replace(FieldValue, "*", "%") + "'"
                Else
                    strFilter = FieldName + " Like '" + Replace(FieldValue, "*", "%") + "'"
                End If

            Else

                If strFilter <> "" Then
                    strFilter = strFilter + " AND " + FieldName + "=" + FieldValue
                Else
                    strFilter = FieldName + "=" + FieldValue
                End If

            End If

        ElseIf FieldValue1 <> "" And FieldValue2 <> "" Then

            If TypeString Then

                If strFilter <> "" Then
                    strFilter = strFilter + " AND (" + FieldName + " Like '" + Replace(FieldValue1, "*", "%") + "' OR " + FieldName + " Like '" + Replace(FieldValue2, "*", "%") + "') "
                Else
                    strFilter = " (" + FieldName + " Like '" + Replace(FieldValue1, "*", "%") + "' OR " + FieldName + " Like '" + Replace(FieldValue2, "*", "%") + "') "
                End If

            Else

                If strFilter <> "" Then
                    strFilter = strFilter + " AND (" + FieldName + "=" + FieldValue1 + " OR " + FieldName + "=" + FieldValue2 + ") "
                Else
                    strFilter = " (" + FieldName + "=" + FieldValue1 + " OR " + FieldName + "=" + FieldValue2 + ") "
                End If

            End If

        End If

    End Sub 'executeAddFilterOR()

    Public Sub executeAddFilterList(ByVal FieldName As String, ByVal FieldValue As String, ByVal TypeString As Boolean)

        FieldValue = Replace(FieldValue, "select", "")
        FieldValue = Replace(FieldValue, "insert", "")
        FieldValue = Replace(FieldValue, "update", "")
        FieldValue = Replace(FieldValue, "delete", "")

        If FieldValue <> "" Then

            If TypeString Then

                If strFilter <> "" Then
                    strFilter = strFilter + " AND " + FieldName + " = '" + FieldValue + "'"
                Else
                    strFilter = FieldName + " = '" + FieldValue + "'"
                End If

            Else

                If strFilter <> "" Then
                    strFilter = strFilter + " AND " + FieldName + "=" + FieldValue
                Else
                    strFilter = FieldName + "=" + FieldValue
                End If

            End If

        End If

    End Sub 'executeAddFilterList()

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
