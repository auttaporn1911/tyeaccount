Imports DataAccess
Imports System.IO
Imports System.Data.OleDb

Public Class UploadAgent
    Inherits System.Web.UI.Page
    Private _dbConnect As DBConnection = Nothing
    Private strLib As String = "TYEACC"
    Public ReadOnly Property Dbconnect As DBConnection
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
        ItemGroup
        Year
        AgentNumber
        AgentName
        PriceCU
        PriceAL

    End Enum

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            GetDDlYears()
            GenGrid()
        End If

    End Sub

    Private Sub GenGrid()
        Dim str As String


        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qryUploadAgent.lbs")
            str = rw.ReadToEnd
        End Using
        str = String.Format(str, strLib)
        Dim dt As New DataTable
        dt = Dbconnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub
    Private SheetName As String
    Protected Sub btnImportAgent_Click(sender As Object, e As EventArgs) Handles btnImportAgent.Click
        If FileUpload1.HasFile Then
            Dim dt As New DataTable
            Dim dtTemp1 As DataTable
            Dim dtTemp2 As DataTable
            Dim dtTemp3 As DataTable

            Dim Filename As String = Path.GetFileName(FileUpload1.PostedFile.FileName)
            Dim Extension As String = Path.GetExtension(FileUpload1.PostedFile.FileName)
            Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")
            Dim FilePath As String = Server.MapPath(FolderPath & "\" & Filename)
            FileUpload1.SaveAs(FilePath)

            SheetName = "BKK"
            dtTemp1 = ExcuteQueryExcelAgent(FilePath, Extension, "test", SheetName.Trim)
            If Not dtTemp1 Is Nothing Then
                dtTemp1 = ConvertDataTable(dtTemp1)
                dt.Merge(dtTemp1)
            End If


            SheetName = "UPCOUNTRY"
            dtTemp2 = ExcuteQueryExcelAgent(FilePath, Extension, "test", SheetName.Trim)
            If Not dtTemp2 Is Nothing Then
                dtTemp2 = ConvertDataTable(dtTemp2)
                dt.Merge(dtTemp2)
            End If


            'SheetName = "ProjectAgent"
            'dtTemp3 = ExcuteQueryExcelAgent(FilePath, Extension, "test", SheetName.Trim)
            'If Not dtTemp3 Is Nothing Then
            '    dtTemp3 = ConvertDataTable(dtTemp3)
            '    dt.Merge(dtTemp3)
            'End If


            'GridView2.DataSource = dt
            'GridView2.DataBind()
            ImportDataToAS400(dt)


            MessageBox("Upload Agent " & result & "records")
            GenGrid()
        End If
    End Sub
    Public Function ExcuteQueryExcelAgent(FilePath As String, Extension As String, isHDR As String, sheet As String) As DataTable
        Dim conStr As String = ""
        Select Case Extension
            Case ".xls"
                'Excel 97-03
                conStr = ConfigurationManager.ConnectionStrings("Excel03ConString") _
                         .ConnectionString
                Exit Select
            Case ".xlsx"
                'Excel 07
                conStr = ConfigurationManager.ConnectionStrings("Excel07ConString") _
                          .ConnectionString
                Exit Select
        End Select
        conStr = String.Format(conStr, FilePath, isHDR)
        Dim connExcel As New OleDbConnection(conStr)
        Dim cmdExcel As New OleDbCommand()
        Dim oda As New OleDbDataAdapter()
        Dim dt As New DataTable()
        cmdExcel.Connection = connExcel
        connExcel.Open()
        Dim dtExcelSchema As DataTable
        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString

        SheetName = sheet + "$"

        For i As Integer = 0 To (dtExcelSchema.Rows.Count - 1) Step 1
            Dim SheetNameS As String = dtExcelSchema.Rows(i)("TABLE_NAME").ToString
            If dtExcelSchema.Rows(i)("TABLE_NAME").ToString = SheetName Then
                connExcel.Close()
                connExcel.Open()
                cmdExcel.CommandText = "SELECT * From [" + SheetName + "]"
                oda.SelectCommand = cmdExcel
                oda.Fill(dt)
                connExcel.Close()
                Exit For
            End If
        Next
        Return dt

    End Function
    Public Function ConvertDataTable(dt As DataTable) As DataTable
        Dim startRow As Integer
        Dim count As Integer = 0
        startRow = 5
        Dim newDt As New DataTable
        newDt.Columns.Add("Nametype", GetType(String))
        newDt.Columns.Add("Name", GetType(String))
        newDt.Columns.Add("ClassName", GetType(String))
        newDt.Columns.Add("Qty", GetType(Double))
        newDt.Columns.Add("Amount", GetType(Double))
        newDt.Columns.Add("FC", GetType(Double))
        newDt.Columns.Add("Profit", GetType(Double))
        newDt.Columns.Add("Month", GetType(Integer))
        newDt.Columns.Add("Monthname", GetType(String))
        newDt.Columns.Add("policy", GetType(Integer))
        newDt.Columns.Add("ItemGroup", GetType(Integer))
        newDt.Columns.Add("Year", GetType(Integer))
        newDt.Columns.Add("AgentNumber", GetType(Integer))
        newDt.Columns.Add("AgentName", GetType(String))
        newDt.Columns.Add("PriceCU", GetType(Integer))
        newDt.Columns.Add("PriceAL", GetType(Integer))
        Dim dr As DataRow
        Dim ItemGroup As Integer = 1
        Dim PricCu As Integer = 0
        Dim PricAL As Integer = 0

        For Each r As DataRow In dt.Rows
            If r(2).ToString = "" Then
                count = count + 1
                If Trim(r(0).ToString).PadRight(5, " ").Substring(0, 5).ToUpper = "TOTAL" Then
                    ItemGroup = ItemGroup + 1
                End If
                Continue For
            ElseIf r(2).ToString.Trim = "CU" Then
                count = count + 1
                PricCu = r(3)
                Continue For
            ElseIf r(2).ToString.Trim = "AL" Then
                count = count + 1
                PricAL = r(3)
                Continue For
            ElseIf r(2).ToString.Trim <> "" And Trim(r(1).ToString).ToUpper = "TOTAL" Then
                count = count + 1
                Continue For
            End If

            If count >= startRow Then
                For month As Integer = 7 To 18
                    dr = newDt.NewRow
                    MapRow(dr, r, month, ItemGroup, PricCu, PricAL)
                    newDt.Rows.Add(dr)
                Next
            End If
            count = count + 1
        Next
        Return newDt
    End Function

    Private name As String
    Public Function MapRow(dr As DataRow, r As DataRow, month As Integer, ItemGroup As Integer, PricCu As Integer, PricAL As Integer) As DataRow


        If r(0).ToString <> "" Then
            NameType = r(0).ToString.Trim
        End If

        dr(0) = NameType

        If r(1).ToString <> "" Then
            name = r(1)
        Else
            name = name
        End If
        dr(1) = name

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
        dr(9) = ddltest.SelectedItem.ToString.Trim
        dr(10) = ItemGroup
        If (month < 7) Then
            dr(11) = Integer.Parse(ddltest.SelectedValue) + 1
        Else
            dr(11) = ddltest.SelectedValue
        End If

        If SheetName = "BKK" Then
            dr(12) = 1
        ElseIf SheetName = "UPCOUNTRY" Then
            dr(12) = 2
        Else
            dr(12) = 3
        End If
        dr(13) = SheetName.Trim
        dr(14) = PricCu
        dr(15) = PricAL
        Return dr

    End Function
    'Public Function GetYear(month As Integer, policy As Integer) As Integer
    '    Dim year As Integer
    '    year = 1940
    '    year = year + policy

    '    If month >= 1 And month <= 6 Then '
    '        year = year + 1
    '    End If
    '    Return year

    'End Function
    Public Sub GetDDlYears()
        Dim nowYear As Integer = Year(Now)
        Dim loopYear As Integer
        For i As Integer = 5 To 1 Step -1
            Dim list As New ListItem()
            loopYear = nowYear - i
            list.Text = GetConvYear(loopYear)
            list.Value = loopYear
            ddltest.Items.Add(list)
        Next
        For j As Integer = 1 To 5 Step 1
            If j = 1 Then
                loopYear = nowYear
            Else
                loopYear = nowYear + j
            End If

            Dim list As New ListItem()
            list.Text = GetConvYear(loopYear)
            list.Value = loopYear
            ddltest.Items.Add(list)
        Next

        Dim SetYear As Integer = Year(Now)
        Dim nowMonth As Integer = Month(Now)
        If nowMonth >= 1 And nowMonth <= 6 Then
            ddltest.SelectedValue = Convert.ToString(SetYear - 1)
        ElseIf nowMonth >= 7 And nowMonth <= 12 Then
            ddltest.SelectedValue = Convert.ToString(SetYear)
        End If

    End Sub
    Public Function GetConvYear(RealYear As Integer) As Integer
        Dim NowYear, CountY As Integer
        NowYear = RealYear
        CountY = NowYear - 1940
        Return CountY
    End Function

    Protected Sub btnDownload_Click(sender As Object, e As ImageClickEventArgs) Handles btnDownload.Click
        Dim filename As String
        filename = "Template\" & "PlanSaleByAgent" & ".xlsx"
        Response.ContentType = ContentType

        Response.AppendHeader("Content-Disposition", ("attachment; filename=" + Path.GetFileName(filename)))

        Response.WriteFile(filename)

        Response.End()
    End Sub

    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub

    Public Function AddSingleQuoat(str As Object) As String
        str = "'" & str.ToString & "'"
        Return str
    End Function

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
    Protected Function ConvertInt(str As Object) As Integer
        If str.ToString = "" Then
            Return 0
        End If
        Return CDec(str)
    End Function
    Protected Function ConvertDec(str As Object) As Decimal
        If str.ToString = "" Then
            Return 0
        End If
        Return CDec(str)
    End Function

    Private result As Integer
    Public Sub ImportDataToAS400(dt As DataTable)
        Dim cmd As String = ""
        Dim cmdChk As String = ""
        Dim cmdUpdate As String = ""
        Dim strcmd As String
        Dim cnt As Integer
        Dim cname As String
        Dim start As Boolean = False
        'cmdChk = "select count(*) from " & strLib & ".TACLMA where CLCLCD = {0}"
        cname = "AGITTY,AGITNM,AGITCL,AGQTY,AGAMNT,AGFCST,AGPROF,AGMONM,AGMNNM,AGPLYR,AGADDT,AGADTM,AGITGP,AGYEAR,AGLOT,AGCODE,AGNAME,AGCURT,AGALRT"

        cmd = "insert into  " & strLib & ".TASPAG(" & cname & ") values ({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}," & Utility.GenLot() & ",{14},{15},{16},{17})"

        For Each dr As DataRow In dt.Rows
            strcmd = String.Format(cmd, _
                                  AddSingleQuoat(dr(eColumn.NameType)), _
                                  AddSingleQuoat(dr(eColumn.Name).ToString.Replace("-", "")), _
                                  AddSingleQuoat(dr(eColumn.ClassName)), _
                                  ConvertDec(dr(eColumn.Qty)), _
                                  ConvertDec(dr(eColumn.Amount)), _
                                  ConvertDec(dr(eColumn.FC)), _
                                  ConvertDec(dr(eColumn.Profit)), _
                                  ConvertInt(dr(eColumn.Month)), _
                                  AddSingleQuoat(dr(eColumn.MonthName)), _
                                  ConvertDec(dr(eColumn.policy)), _
                                  Utility.GetDate, _
                                  Utility.GetTime, _
                                  ConvertInt(dr(eColumn.ItemGroup)), _
                                  ConvertInt(dr(eColumn.Year)), _
                                  ConvertInt(dr(eColumn.AgentNumber)), _
                                  AddSingleQuoat(dr(eColumn.AgentName)), _
                                  ConvertInt(dr(eColumn.PriceCU)), _
                                  ConvertInt(dr(eColumn.PriceAL)))

            result = result + Dbconnect.ExcuteNonQueryString(strcmd, DBConnection.DatabaseType.AS400)

        Next

    End Sub

    Protected Sub GridView1_RowCommand(sender As Object, e As GridViewCommandEventArgs)


    End Sub

    Protected Sub GridView1_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)

        Dim MyIdentity As Integer = e.RowIndex
        Dim lbldeleteID As Label = DirectCast(GridView1.Rows(e.RowIndex).FindControl("lblLotno"), Label)
        Dim lblAgent As Label = DirectCast(GridView1.Rows(e.RowIndex).FindControl("lblAgent"), Label)
        Dim cmd As String = ""
        Dim strcmd As String = ""
        cmd = "Delete from  " & strLib & ".TASPAG where AGLOT = {0} and AGNAME = {1}"
        strcmd = String.Format(cmd, ConvertInt(lbldeleteID.Text.Trim), AddSingleQuoat(lblAgent.Text.Trim))
        result = result + Dbconnect.ExcuteNonQueryString(strcmd, DBConnection.DatabaseType.AS400)

        If result > 0 Then
            MessageBox("Success!")
            GenGrid()

        Else
            MessageBox("Unsuccess!")
        End If

    End Sub


    Protected Sub GridView1_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim item As Label = DirectCast(e.Row.FindControl("lblLotno"), Label)
            Dim Ag As Label = DirectCast(e.Row.FindControl("lblAgent"), Label)
            Select Case e.Row.RowType
                Case DataControlRowType.DataRow
                    Dim BtnDelete = DirectCast(e.Row.FindControl("btnDelete"), LinkButton)
                    BtnDelete.OnClientClick = "if(!confirm('Do you want to delete Agent: " + Ag.Text + " lot: " + item.Text + " ?')){ return false; };"
            End Select
        End If
    End Sub
End Class