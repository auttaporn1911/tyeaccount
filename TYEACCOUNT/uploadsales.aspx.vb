Imports DataAccess
Imports System.IO
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine

Public Class WebForm2
    Inherits System.Web.UI.Page

    Private _dbConnect As DBConnection = Nothing
    Public ReadOnly Property DbConnect As DBConnection
        Get
            If _dbConnect Is Nothing Then
                _dbConnect = New DBConnection
                Return _dbConnect
            End If
            Return _dbConnect
        End Get
    End Property
    Private strLib As String = "TYEACC"
    Public Enum eColumn
        ItemClass = 0
        ItemCode
        Desc
        Lenght
        CustCode
        CustName
        DocType
        DocNo
        InvoiceDate
        OrderDate
        RequestDate
        DelNo
        DelDate
        OrderType
        OrderNo
        Quantity
        UnitPrice
        Amount
        ProfitFP
        UnitCost
        TotalCost
        Weight
        TotalWeight
        PercentExp
        CustomerType
        Region
        ProductCode
        ProfitFC
        InvoiceDate400
        OrderDate400
        ReqDate400
        DeliveryDate400

    End Enum

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            LoadAllLot()
        End If
    End Sub

    Protected Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click

        If lbFileName.Text <> "" Then
            Dim dt As New DataTable
            Dim FileName As String = Path.GetFileName(lbFileName.Text)

            Dim Extension As String = Path.GetExtension(lbFileName.Text)

            Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")

            Dim FilePath As String = Server.MapPath(FolderPath & "\" & FileName)
            Dim lotno As String

            lbFileName.Text = FilePath
            dt = DbConnect.ExcuteQueryExcel(FilePath, Extension, "test", ddlSheet.SelectedItem.Text)
            dt = TableEnterprise(dt)
            lotno = Utility.GenLot()
            If Not ImportDataToAS400(dt, lotno) Then
                MessageBox("Error Step 1. Import Sale Data.")
                Exit Sub
            End If
            If Not SammarySaleByClass(lotno) Then
                MessageBox("Error Step 2. Summarize sales by class.")
                Exit Sub
            End If
            If Not SammarySaleByCustomer(lotno) Then
                MessageBox("Error Step 3. Summarize sales by customer.")
                Exit Sub
            End If

            If Not SammarySaleByAgent(lotno) Then
                MessageBox("Error Step 4. Summarize sales by Agent.")
                Exit Sub
            End If

            LoadAllLot()
        Else
            MessageBox("Please upload files")
        End If


    End Sub
    Public Function TableEnterprise(dt As DataTable) As DataTable
        dt.Columns.Add("OrderDate400", GetType(Integer))
        dt.Columns.Add("ReqDate400", GetType(Integer))
        dt.Columns.Add("InvoiceDate400", GetType(Integer))
        dt.Columns.Add("DeliveryDate400", GetType(Integer))
        'Dim dt2 As New DataTable
        'dt2 = dt.Clone
        Dim lst As New List(Of Integer)
        Dim isStart As Boolean = False
        Dim i As Integer = 0
        For Each dr As DataRow In dt.Rows

            If dr(0).ToString = "" Then
                If i < 10 Then
                    lst.Add(i)
                End If

                i += 1
                Continue For
            Else
                If isStart Then
                    dr(eColumn.OrderDate400) = ConvertAS400Date(dr(eColumn.OrderDate))
                    dr(eColumn.InvoiceDate400) = ConvertAS400Date(dr(eColumn.InvoiceDate))
                    dr(eColumn.ReqDate400) = ConvertAS400Date(dr(eColumn.RequestDate))
                    If dr(eColumn.DelDate) Is Nothing Then
                        dr(eColumn.DeliveryDate400) = ConvertAS400Date(dr(eColumn.DelDate))
                    Else
                        dr(eColumn.DeliveryDate400) = 0
                    End If
                Else
                    If i < 10 Then
                        lst.Add(i)
                    End If
                    isStart = True
                End If
            End If
            i += 1
        Next
        For j As Integer = lst.Count - 1 To 0 Step -1

            dt.Rows.RemoveAt(lst(j))
        Next j
        Return dt
    End Function

    Public Function ConvertAS400Date(strdate As DateTime) As Integer
        Return CInt(String.Format("{0:yyyyMMdd}", strdate))
    End Function
    Public Function ImportDataToAS400(dt As DataTable, lotno As String) As Boolean
        Dim cmd As String = ""
        Dim strcmd As String
        Dim resultAll As Integer
        Dim result As Integer
        Dim cname As String
        Dim dtErr As New DataTable
        dtErr.Columns.Add("Class", GetType(String))
        dtErr.Columns.Add("Itemcode", GetType(String))
        dtErr.Columns.Add("Invoice No", GetType(String))
        cname = "MAITCL,MAITCD,	MADESC,	MALEN,	MACSCD,	MACSNM,	MADOCT,	MADOCN	,MADODT,MAORDT,	MARQDT,	MADLNO,	MADLDT,	MAODTY,	MAORDN,	MAQTY," & _
                "MAUPRC, MAAMT, MAPRFP, MAUCST,MATCST,MAWGHT,MATWGH,MAPCEX,	MACSTP,	MARGN,	MAPDCD,MAPRFC,MALOT"
        cmd = "insert into  " & strLib & ".TASMAS(" & cname & ") values ({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27}," & lotno & ")"

        For Each dr As DataRow In dt.Rows
            If dr(eColumn.ItemClass).ToString.Trim = "" Then
                Continue For
            End If
            strcmd = String.Format(cmd, Utility.AddSingleQuoat(dr(eColumn.ItemClass)), Utility.AddSingleQuoat(dr(eColumn.ItemCode)), Utility.AddSingleQuoat(dr(eColumn.Desc)), Utility.AddSingleQuoat(dr(eColumn.Lenght)), _
                                dr(eColumn.CustCode), Utility.AddSingleQuoat(dr(eColumn.CustName)), _
                                Utility.AddSingleQuoat(dr(eColumn.DocType)), dr(eColumn.DocNo), dr(eColumn.InvoiceDate400), dr(eColumn.OrderDate400), dr(eColumn.ReqDate400), _
                                GetInteger(dr(eColumn.DelNo)), dr(eColumn.DeliveryDate400), Utility.AddSingleQuoat(dr(eColumn.OrderType)), dr(eColumn.OrderNo), dr(eColumn.Quantity), dr(eColumn.UnitPrice), _
                                GetInteger(dr(eColumn.Amount)), GetInteger(dr(eColumn.ProfitFP)), GetInteger(dr(eColumn.UnitCost)), GetInteger(dr(eColumn.TotalCost)), GetInteger(dr(eColumn.Weight)), _
                                GetInteger(dr(eColumn.TotalWeight)), GetInteger(dr(eColumn.PercentExp)), _
                                dr(eColumn.CustomerType), Utility.AddSingleQuoat(dr(eColumn.Region)), dr(eColumn.ProductCode), GetInteger(dr(eColumn.ProfitFC)))
            result = DbConnect.ExcuteNonQueryString(strcmd, DBConnection.DatabaseType.AS400)
            If result = 0 Then
                dtErr.Rows.Add(dr(eColumn.ItemClass), dr(eColumn.ItemCode), dr(eColumn.DocNo))
            Else
                resultAll = resultAll + result
            End If

        Next
        gvErr.DataSource = dtErr
        gvErr.DataBind()
        MessageBox("Upload Sales Data " & resultAll & "records Lot number " & lotno)
        If resultAll > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub

    Public Sub LoadAllLot()
        Dim str As String
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/GetLotUploadSale.lbs")
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

    Public Function ValidateData(dt As DataTable) As Boolean
        Dim isValid As Boolean = False
        Dim err As String
        For Each dr As DataRow In dt.Rows
            If dr(eColumn.ItemClass).ToString().Length > 3 Then
                err += dr(eColumn.ItemClass) & "Item class is incorrect format"
            End If

            If dr(eColumn.ItemCode) <> 19 Then
                err += dr(eColumn.ItemCode) & "Item code is incorrect format"
            End If
            If dr(eColumn.CustCode) Is Nothing Then

            End If
            If dr(eColumn.DocNo) Then

            End If


        Next

        Return isValid
    End Function



    'Public Sub ImportDataToAS400(dt As DataTable)
    '    Dim cmd As String = ""
    '    Dim strcmd As String

    '    Dim cname As String
    '    cname = "MAITCL,MAITCD,	MADESC,	MALEN,	MACSCD,	MACSNM,	MADOCT,	MADOCN	,MADODT,MAORDT,	MARQDT,	MADLNO,	MADLDT,	MAODTY,	MAORDN,	MAQTY," & _
    '            "MAUPRC, MAAMT, MAPROF, MAUCST,MATCST,MAWGHT,MATWGH,MAPCEX,	MACSTP,	MARGN,	MAPDCD"
    '    cmd = "insert into  #pik.TASMAS(" & cname & ") values ({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26})"

    '    For Each dr As DataRow In dt.Rows

    '        strcmd = String.Format(cmd, Utility.AddSingleQuoat(dr(eColumn.ItemClass)), Utility.AddSingleQuoat(dr(eColumn.ItemCode)), Utility.AddSingleQuoat(dr(eColumn.Desc)), GetInteger(dr(eColumn.Lenght)), _
    '                            dr(eColumn.CustCode), Utility.AddSingleQuoat(dr(eColumn.CustName)), _
    '                            Utility.AddSingleQuoat(dr(eColumn.DocType)), dr(eColumn.DocNo), dr(eColumn.InvoiceDate400), dr(eColumn.OrderDate400), dr(eColumn.ReqDate400), _
    '                            GetInteger(dr(eColumn.DelNo)), dr(eColumn.DeliveryDate400), Utility.AddSingleQuoat(dr(eColumn.OrderType)), dr(eColumn.OrderNo), dr(eColumn.Quantity), dr(eColumn.UnitPrice), _
    '                            dr(eColumn.Amount), dr(eColumn.Profit), dr(eColumn.UnitCost), dr(eColumn.TotalCost), dr(eColumn.Weight), dr(eColumn.TotalWeight), GetInteger(dr(eColumn.PercentExp)), _
    '                            dr(eColumn.CustomerType), Utility.AddSingleQuoat(dr(eColumn.Region)), dr(eColumn.ProductCode))
    '        DbConnect.ExcuteNonQueryString(strcmd, DBConnection.DatabaseType.AS400)
    '    Next
    'End Sub
    Public Function SammarySaleByCustomer(lotno As String) As Boolean
        Dim str As String
        Dim result As Integer
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qrySummarybyCust.lbs")
            str = String.Format(rw.ReadToEnd, strLib, lotno)
        End Using

        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)

        If result > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
   
    Public Function SammarySaleByClass(lotno As String) As Boolean
        Dim str As String
        Dim result As Integer
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qrySummarybyClass.lbs")
            str = String.Format(rw.ReadToEnd, strLib, lotno)
        End Using

        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
        If result > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function SammarySaleByAgent(lotno As String) As Boolean
        Dim str As String
        Dim result As Integer
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qrySummaryAgent.lbs")
            str = String.Format(rw.ReadToEnd, strLib, lotno)
        End Using

        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
        If result > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

   
    
    Protected Sub btnDownload_Click(sender As Object, e As ImageClickEventArgs) Handles btnDownload.Click
        Dim filename As String
        filename = "Template\Sales Data.xlsx"
        Response.ContentType = ContentType

        Response.AppendHeader("Content-Disposition", ("attachment; filename=" + Path.GetFileName(filename)))

        Response.WriteFile(filename)

        Response.End()
    End Sub

    Protected Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        Dim list As List(Of String)


        If FileUpload1.HasFile Then
            Dim dt As New DataTable
            Dim FileName As String = Path.GetFileName(FileUpload1.PostedFile.FileName)

            Dim Extension As String = Path.GetExtension(FileUpload1.PostedFile.FileName)

            Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")

            Dim FilePath As String = Server.MapPath(FolderPath & "\" & FileName)

            FileUpload1.SaveAs(FilePath)
            lbFileName.Text = FilePath
            list = DbConnect.ListSheetInExcel(FilePath)
            ddlSheet.DataSource = list
            ddlSheet.DataBind()
        Else
            lbFileName.Text = ""
        End If
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
        str = "delete " & strLib & ".TASMAS where MALOT = '" & lotno & "'"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
        str = "delete " & strLib & ".TASMCC where CCLOT = '" & lotno & "'"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
        str = "delete " & strLib & ".TASMSC where SCLOT = '" & lotno & "'"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
        str = "delete " & strLib & ".TASMAG where SALOT = '" & lotno & "'"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
        MessageBox("Delete completed!")
        LoadAllLot()
    End Sub
End Class