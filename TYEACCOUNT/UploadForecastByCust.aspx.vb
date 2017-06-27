Imports DataAccess
Imports System.IO

Public Class WebForm9
    Inherits System.Web.UI.Page
    Private _dbConnect As DBConnection = Nothing
    Private strLib As String = "TYEACC"
    Public ReadOnly Property DbConnect As DBConnection
        Get
            If _dbConnect Is Nothing Then
                _dbConnect = New DBConnection
                Return _dbConnect
            End If
            Return _dbConnect
        End Get
    End Property
    Private CURate As Integer
    Private AlRate As Integer
    Public Enum eColumn
        CSGroup = 0
        CUSTOM
        MATHERIAL
        Qty
        Amount
        FC
        Profit
        Month
        MonthName
        policy
        Year
    End Enum
    Private custom As String
    Private isCustom As Boolean
    Private NameType As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            LoadAllLot()
            LoadCheckMapping()
        End If
    End Sub

    Protected Sub btnImportForecast_Click(sender As Object, e As EventArgs) Handles btnImportForecast.Click
        If FileUpload1.HasFile Then
            Dim dt As New DataTable
            Dim FileName As String = Path.GetFileName(FileUpload1.PostedFile.FileName)

            Dim Extension As String = Path.GetExtension(FileUpload1.PostedFile.FileName)

            Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")

            Dim FilePath As String = Server.MapPath(FolderPath & "\" & FileName)

            FileUpload1.SaveAs(FilePath)

            dt = DbConnect.ExcuteQueryExcel(FilePath, Extension, "test", "Plan Sale by Customer")
            dt = ConvertDataTable(dt)
            ImportDataToAS400(dt)
            LoadAllLot()
            LoadCheckMapping()
        End If
    End Sub
   

    Public Function GetInteger(str As Object) As String
        If str.ToString.ToUpper.Contains("M") Then
            str = str.ToString.ToUpper.Replace("M", "")
        End If
        Dim num As Decimal
        If Not Decimal.TryParse(str.ToString, num) Then
            str = "0"
        End If
        Return str
    End Function
    Public Sub ImportDataToAS400(dt As DataTable)
        Dim cmd As String = ""
        Dim cmdChk As String = ""
        Dim cmdUpdate As String = ""
        Dim strcmd As String
        Dim result As Integer
        Dim cnt As Integer
        Dim cname As String
        Dim start As Boolean = False
        'cmdChk = "select count(*) from " & strLib & ".TACLMA where CLCLCD = {0}"
        cname = "CSGRUP,CSCUST,CSMAT,CSQTY,CSAMNT,CSFCST,CSPROF,CSMONM,CSMNNM,CSPLYR,CSADDT,CSADTM,CSYEAR,CSLOT,CSCURT,CSALRT"

        cmd = "insert into  " & strLib & ".TASPCS(" & cname & ") values ({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}," & Utility.GenLot() & ",{13},{14})"

        For Each dr As DataRow In dt.Rows
            strcmd = String.Format(cmd, Utility.AddSingleQuoat(dr(eColumn.CSGroup)), Utility.AddSingleQuoat(dr(eColumn.CUSTOM).ToString.Replace("-", "")), Utility.AddSingleQuoat(dr(eColumn.MATHERIAL)), _
                                   GetInteger(dr(eColumn.Qty)), GetInteger(dr(eColumn.Amount)), GetInteger(dr(eColumn.FC)), GetInteger(dr(eColumn.Profit)), dr(eColumn.Month), _
                            Utility.AddSingleQuoat(dr(eColumn.MonthName)), dr(eColumn.policy), Utility.GetDate, Utility.GetTime, dr(eColumn.Year), CURate, AlRate)
            result = result + DbConnect.ExcuteNonQueryString(strcmd, DBConnection.DatabaseType.AS400)


        Next
        MessageBox("Upload customer " & result & "records")
    End Sub
    Public Sub LoadCheckMapping()
        Dim str As String

        str = "select distinct A.CSCUST from {0}.TASPCS A left join {0}.TACSMA B "
        str &= "on trim(upper(A.CSCUST)) = trim(upper(B.CSCUTM)) where B.CSCUTM is null"
        str = String.Format(str, strLib)
        gvCheck.DataSource = DbConnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        gvCheck.DataBind()
    End Sub

    Public Sub LoadAllLot()
        Dim str As String
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/GetLotPlanByCust.lbs")
            str = String.Format(rw.ReadToEnd, strLib)
        End Using
        gvLot.DataSource = DbConnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        gvLot.DataBind()
    End Sub

    Public Function ConvertDataTable(dt As DataTable) As DataTable
        Dim startRow As Integer
        Dim rowStart As Integer = 3
        Dim count As Integer = 0
        isCustom = True
        startRow = 15  'Dummy data
        Dim newDt As New DataTable
        newDt.Columns.Add("CSGROUP", GetType(String))
        newDt.Columns.Add("CUSTOM", GetType(String))
        newDt.Columns.Add("MATHERIAL", GetType(String))
        newDt.Columns.Add("Qty", GetType(Double))
        newDt.Columns.Add("Amount", GetType(Double))
        newDt.Columns.Add("FC", GetType(Double))
        newDt.Columns.Add("Profit", GetType(Double))
        newDt.Columns.Add("Month", GetType(Integer))
        newDt.Columns.Add("MonthName", GetType(String))
        newDt.Columns.Add("policy", GetType(Integer))
        newDt.Columns.Add("Year", GetType(Integer))
        Dim dr As DataRow
        NameType = ""
        For Each r As DataRow In dt.Rows
            If r(1).ToString.ToUpper.Contains("GRAND") Then
                Exit For
            End If

            rowStart += 1
            If rowStart < 3 Then
                Continue For
            End If
            'If r(1).ToString.ToUpper.Contains("TOTAL") Or (NameType.Contains("TOTAL") And r(1).ToString = "") Then
            '    NameType = "TOTAL"
            '    Continue For
            'End If

            If r(3).ToString.ToUpper = "CU" Then
                CURate = GetInteger(r(4))
            End If
            If r(3).ToString.ToUpper = "AL" Then
                AlRate = GetInteger(r(4))
            End If
            If r(3).ToString.ToUpper = "GROUP" Then
                startRow = count + 2
            End If

            If r(3).ToString.ToUpper.Contains("TOTAL") Then
                isCustom = True
                Continue For
            End If

            If isCustom Then
                custom = r(2).ToString
                If r(2).ToString.Contains(":") Then
                    custom = r(2).ToString.Split(":")(0)
                    NameType = r(2).ToString.Split(":")(1)
                End If

                If custom <> "" Then
                    isCustom = False
                Else
                    If r(1).ToString <> "" And r(1).ToString.Contains(":") Then

                        custom = r(1).ToString.Split(":")(0).Trim
                        isCustom = False
                    End If

                End If

            End If

            If custom.ToUpper.Contains("TOTAL") Then
                Continue For
            End If

            If count >= startRow Then

                For month As Integer = 7 To 18
                    dr = newDt.NewRow
                    MapRow(dr, r, month)
                    newDt.Rows.Add(dr)
                Next

            End If
            count = count + 1
        Next

        Return newDt
    End Function

    Protected Sub delete_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lotno As String
        lotno = CType(sender, LinkButton).Attributes("LOTNO")
        Dim result As Integer
        Dim str As String
        str = "delete " & strLib & ".TASPCS where CSLOT = '" & lotno & "'"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
       
        MessageBox("Delete completed!")
        LoadAllLot()

    End Sub

    Protected Sub gvLot_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvLot.RowDataBound
        Dim control As New LinkButton
        'Dim clbl As New Label
        'Dim strSpilt As String

        If e.Row.RowIndex > -1 And e.Row.RowType = DataControlRowType.DataRow Then
            control = CType(e.Row.FindControl("imgEdit"), LinkButton)
            If Not control Is Nothing Then

                control.Attributes("LOTNO") = DataBinder.Eval(e.Row.DataItem, "LOTNO").ToString()

            End If

        End If
    End Sub

    Public Function MapRow(dr As DataRow, r As DataRow, month As Integer) As DataRow
       
        If r(1).ToString <> "" Then
            If r(1).ToString.Contains(":") Then
                NameType = r(1).ToString.Split(":")(1).Trim
            Else
                NameType = r(1).ToString
            End If

            If r(2).ToString.Contains(":") Then
                NameType = r(2).ToString.Split(":")(1).Trim
            End If
        End If

       

        dr(0) = NameType
        dr(1) = custom
        dr(2) = r(3)
        If month > 12 Then
            dr(3) = r(15 + ((month - 7) * 4))
            dr(4) = r(16 + ((month - 7) * 4))
            dr(5) = r(17 + ((month - 7) * 4))
            dr(6) = r(18 + ((month - 7) * 4))
        Else
            dr(3) = r(7 + ((month - 7) * 4))
            dr(4) = r(8 + ((month - 7) * 4))
            dr(5) = r(9 + ((month - 7) * 4))
            dr(6) = r(10 + ((month - 7) * 4))
        End If

        month = IIf(month Mod 12 = 0, 12, month Mod 12)
        dr(7) = month
        dr(8) = Utility.GetMonthName(month)
        dr(9) = ddlPolicy.SelectedValue

        dr(10) = GetYear(month, CInt(ddlPolicy.SelectedValue))
        Return dr
    End Function

    Public Function GetYear(month As Integer, policy As Integer) As Integer
        Dim year As Integer
        year = 1940
        year = year + policy

        If month >= 1 And month <= 6 Then
            year = year + 1
        End If
        Return year
    End Function

    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub

    Protected Sub btnDownload_Click(sender As Object, e As ImageClickEventArgs) Handles btnDownload.Click
        Dim filename As String
        filename = "Template\Plan Sale By Customer.xlsx"
        Response.ContentType = ContentType

        Response.AppendHeader("Content-Disposition", ("attachment; filename=" + Path.GetFileName(filename)))

        Response.WriteFile(filename)

        Response.End()
    End Sub
End Class