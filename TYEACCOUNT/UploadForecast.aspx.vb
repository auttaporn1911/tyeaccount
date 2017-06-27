Imports DataAccess
Imports System.IO

Public Class WebForm6
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
    Private NameType As String

    Public Enum eColumn
        NameType = 0
        Name
        ClassName
        Qty
        Amount
        FC
        Profit
        Month
        MonthName
        policy
        ITEMGROUP
        Year
    End Enum
    Protected Sub ShowClassUnmatch()
        Dim str As String
        str = "select distinct SPITCL from " & strLib & ".TASPSA " & _
                "left join " & strLib & ".TACLMA on SPITCL = CLCLCD where CLCLCD is null "
        Dim dt As New DataTable
        dt = DbConnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        GridView1.DataSource = dt
        GridView1.DataBind()

    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            ShowClassUnmatch()
            LoadAllLot()
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

            dt = DbConnect.ExcuteQueryExcel(FilePath, Extension, "test", "Plan Sale by Class")
            dt = ConvertDataTable(dt)
            ImportDataToAS400(dt)
        End If

    End Sub

    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
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

    Protected Sub delete_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lotno As String
        lotno = CType(sender, LinkButton).Attributes("LOTNO")
        Dim result As Integer
        Dim str As String
        str = "delete " & strLib & ".TASPSA where SPLOT = '" & lotno & "'"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)

        MessageBox("Delete completed!")
        ShowClassUnmatch()
        LoadAllLot()
    End Sub

    Public Sub LoadAllLot()
        Dim str As String
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/GetLotPlanByClass.lbs")
            str = String.Format(rw.ReadToEnd, strLib)
        End Using
        gvLot.DataSource = DbConnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        gvLot.DataBind()
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
    Public Function ConvertDataTable(dt As DataTable) As DataTable
        Dim startRow As Integer
        Dim count As Integer = 0

        startRow = 5
        Dim newDt As New DataTable
        newDt.Columns.Add("NameType", GetType(String))
        newDt.Columns.Add("Name", GetType(String))
        newDt.Columns.Add("ClassName", GetType(String))
        newDt.Columns.Add("Qty", GetType(Double))
        newDt.Columns.Add("Amount", GetType(Double))
        newDt.Columns.Add("FC", GetType(Double))
        newDt.Columns.Add("Profit", GetType(Double))
        newDt.Columns.Add("Month", GetType(Integer))
        newDt.Columns.Add("MonthName", GetType(String))
        newDt.Columns.Add("policy", GetType(Integer))
        newDt.Columns.Add("ItemGroup", GetType(Integer))
        newDt.Columns.Add("Year", GetType(Integer))
        Dim dr As DataRow
        Dim ItemGroup As Integer = 1
        For Each r As DataRow In dt.Rows
            If r(2).ToString = "" Then
                count = count + 1
                If Trim(r(0).ToString).PadRight(5, " ").Substring(0, 5).ToUpper = "TOTAL" Then
                    ItemGroup = ItemGroup + 1
                End If
                Continue For
            End If
            If count >= startRow Then

                For month As Integer = 7 To 18
                    dr = newDt.NewRow
                    MapRow(dr, r, month, ItemGroup)
                    newDt.Rows.Add(dr)
                Next

            End If
            count = count + 1
        Next

        Return newDt
    End Function
    Public Function MapRow(dr As DataRow, r As DataRow, month As Integer, ItemGroup As Integer) As DataRow
        If r(0).ToString <> "" Then
            NameType = r(0).ToString.Trim
        End If

        dr(0) = NameType
        dr(1) = r(1)
        dr(2) = r(2)
        If month > 12 Then
            dr(3) = r(14 + ((month - 7) * 4))
            dr(4) = r(15 + ((month - 7) * 4))
            dr(5) = r(16 + ((month - 7) * 4))
            dr(6) = r(17 + ((month - 7) * 4))
        Else
            dr(3) = r(6 + ((month - 7) * 4))
            dr(4) = r(7 + ((month - 7) * 4))
            dr(5) = r(8 + ((month - 7) * 4))
            dr(6) = r(9 + ((month - 7) * 4))
        End If

        month = IIf(month Mod 12 = 0, 12, month Mod 12)
        dr(7) = month
        dr(8) = Utility.GetMonthName(month)
        dr(9) = ddlPolicy.SelectedValue
        dr(10) = ItemGroup
        dr(11) = GetYear(month, CInt(ddlPolicy.SelectedValue))
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
    'Public Function GenLot() As String
    '    Dim lotno As String
    '    lotno = String.Format("{0:yyMMddHHmm}", DateTime.Now)
    '    Return lotno
    'End Function

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
        cname = "SPITTY,SPITNM,SPITCL,SPQTY,SPAMNT,SPFCST,SPPROF,SPMONM,SPMNNM,SPPLYR,SPADDT,SPADTM,SPITGP,SPYEAR,SPLOT"

        cmd = "insert into  " & strLib & ".TASPSA(" & cname & ") values ({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}," & Utility.GenLot() & ")"

        For Each dr As DataRow In dt.Rows
            strcmd = String.Format(cmd, Utility.AddSingleQuoat(dr(eColumn.NameType)), Utility.AddSingleQuoat(dr(eColumn.Name).ToString.Replace("-", "")), Utility.AddSingleQuoat(dr(eColumn.ClassName)), _
                                   dr(eColumn.Qty), dr(eColumn.Amount), dr(eColumn.FC), dr(eColumn.Profit), dr(eColumn.Month), _
                            Utility.AddSingleQuoat(dr(eColumn.MonthName)), dr(eColumn.policy), Utility.GetDate, Utility.GetTime, dr(eColumn.ITEMGROUP), dr(eColumn.Year))
            result = result + DbConnect.ExcuteNonQueryString(strcmd, DBConnection.DatabaseType.AS400)


        Next
        ShowClassUnmatch()
        LoadAllLot()
        MessageBox("Upload customer " & result & "records")
    End Sub

    Protected Sub btnDownload_Click(sender As Object, e As ImageClickEventArgs) Handles btnDownload.Click
        Dim filename As String
        filename = "Template\Plan Sale By Class.xlsx"
        Response.ContentType = ContentType

        Response.AppendHeader("Content-Disposition", ("attachment; filename=" + Path.GetFileName(filename)))

        Response.WriteFile(filename)

        Response.End()
    End Sub
End Class