Imports DataAccess
Imports System.IO
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Public Class WebForm4
    Inherits System.Web.UI.Page

   

    Private _dbConnect As DBConnection = Nothing
    Private strLib As String = "TYEACC"
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/report/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("Mapping Class.xlt")
    Private ReportName As String = "Mapping Product"
    Private crtDate As String = Date.Now.ToString("ddMMyyyy")
    Private crtTime As String = Date.Now.ToString("HHmm")
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private intcurrow, intno, intstart As Integer
    Private filename As String
    Public ReadOnly Property DbConnect As DBConnection
        Get
            If _dbConnect Is Nothing Then
                _dbConnect = New DBConnection
                Return _dbConnect
            End If
            Return _dbConnect
        End Get
    End Property

    Public Enum eColumn
        ClassName = 0
        Material
        MaterialType
    End Enum
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            BindGridview()
        End If
    End Sub


    Public Sub BindGridview()
        Dim dt As New DataTable
        Dim ds As New DataSet
        Dim cmd As String
        cmd = "select * from " & strLib & ".TACLMA"
        dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.AS400)
        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub
    Public Function ConvertAS400Date(strdate As DateTime) As Integer
        Return CInt(String.Format("{0:yyyyMMdd}", strdate))
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
        cmdChk = "select count(*) from " & strLib & ".TACLMA where CLCLCD = {0}"
        cname = "CLCLCD,CLMTMN,CLADDT,CLADTM,CLMTTP"

        cmd = "insert into  " & strLib & ".TACLMA(" & cname & ") values ({0},{1},{2},{3},{4})"
        cmdUpdate = "update " & strLib & ".TACLMA set CLMTMN={0}," & _
                    "CLADDT={1},CLADTM={2},CLMTTP={4} where CLCLCD = {3}"
        For Each dr As DataRow In dt.Rows
            strcmd = String.Format(cmdChk, Utility.AddSingleQuoat(dr(eColumn.ClassName)))

            'If dr(eColumn.ClassName) = "Class" Then
            '    start = True
            'End If
            start = True
            If start Then
                cnt = CInt(DbConnect.ExcuteScalar(strcmd, DBConnection.DatabaseType.AS400))
                If cnt > 0 Then
                    strcmd = String.Format(cmdUpdate, Utility.AddSingleQuoat(dr(eColumn.Material)), Utility.GetDate, Utility.GetTime, Utility.AddSingleQuoat(dr(eColumn.ClassName)), Utility.AddSingleQuoat(dr(eColumn.MaterialType)))

                    result = result + DbConnect.ExcuteNonQueryString(strcmd, DBConnection.DatabaseType.AS400)
                Else
                    strcmd = String.Format(cmd, Utility.AddSingleQuoat(dr(eColumn.ClassName)), Utility.AddSingleQuoat(dr(eColumn.Material)), _
                            Utility.GetDate, Utility.GetTime, Utility.AddSingleQuoat(dr(eColumn.MaterialType)))
                    result = result + DbConnect.ExcuteNonQueryString(strcmd, DBConnection.DatabaseType.AS400)

                End If
            End If
            


        Next
        MessageBox("Upload customer " & result & "records")
        SummarizeAll()
    End Sub
   
    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub
    Public Function GetInteger(str As Object) As String
        If str.ToString = "" Then
            str = "0"
        End If
        Return str
    End Function

    Protected Sub btnImportClass_Click(sender As Object, e As EventArgs) Handles btnImportClass.Click
        If FileUpload1.HasFile Then
            Dim dt As New DataTable
            Dim FileName As String = Path.GetFileName(FileUpload1.PostedFile.FileName)

            Dim Extension As String = Path.GetExtension(FileUpload1.PostedFile.FileName)

            Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")

            Dim FilePath As String = Server.MapPath(FolderPath & "\" & FileName)

            FileUpload1.SaveAs(FilePath)

            dt = DbConnect.ExcuteQueryExcel(FilePath, Extension, "test", "Mapping Class")

            ImportDataToAS400(dt)

            BindGridview()
        End If
    End Sub

    Protected Sub btnDownload_Click(sender As Object, e As ImageClickEventArgs) Handles btnDownload.Click
        Dim cmd As String
        cmd = "select * from " & strLib & ".TACLMA"
        Dim dt As New DataTable


        dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.AS400)

        ExportExcel(dt)
    End Sub

    

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

    Public Sub SummarizeAll()
        ClearSummarize()
       
        If Not SammarySaleByCustomer("0") Then
            MessageBox("Error Step 2. Summarize sales by customer.")
            Exit Sub
        End If

        If Not SammarySaleByAgent("0") Then
            MessageBox("Error Step 3. Summarize sales by Agent.")
            Exit Sub
        End If
    End Sub
    Public Function ExportExcel(oTable As DataTable) As Boolean
        Dim oRow As DataRow


        If oTable.Rows.Count <= 0 Then
            Return False

        End If
        oRow = oTable.Rows(0)

        intcurrow = 2
        intstart = 0

        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create("Mapping Class")
        With sheet2
            Workbook.Worksheets(0).Range("A1:M1").CopyTo(.Range("A1"))
            '.Range(2, 6).Value = Trim(oTable.Rows(0)("IDESC")) & "-" & oTable.Rows(0)("IDSCE")
            '.Range(3, 1).Value = "AT " & String.Format("{0:dd-MMM-yyyy}", DateTime.Now)
            '.Range(3, 2).Value = "As " & Right(Trim(txtfromdate.Text), 2) & "/" & Mid(Trim(txtfromdate.Text), 5, 2) & "/" & Mid(Trim(txtfromdate.Text), 3, 2) & " - " & Right(Trim(txttodate.Text), 2) & "/" & Mid(Trim(txttodate.Text), 5, 2) & "/" & Mid(Trim(txttodate.Text), 3, 2) & ""

            For Each oRow In oTable.Rows
                Workbook.Worksheets(0).Range("A2:M2").CopyTo(.Range("A" & intcurrow))
                .Range(intcurrow, 1).Text = oRow("CLCLCD").ToString()
                .Range(intcurrow, 2).Text = oRow("CLMTMN").ToString()
                .Range(intcurrow, 3).Value = oRow("CLMTTP").ToString()
                
                intcurrow = intcurrow + 1
            Next

            ' Workbook.Worksheets(0).Range("A31:O38").CopyTo(.Range("A" & intcurrow))


        End With
        Call SetPageProperties()

        'Workbook.Worksheets("Data").Range("A2:AZ10000").CopyTo(Workbook.Worksheets(0).Range("A2:AZ10000"))
        Workbook.Worksheets(0).Remove()
        ' Workbook.Worksheets(0).Range("AF1").Value = "As of " & DateTime.Now.ToString("dd MMM yyyy")
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
        Return True
    End Function
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

   
    Public Sub ClearSummarize()

        Dim result As Integer
        Dim str As String

        str = "delete " & strLib & ".TASMCC"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
       
        str = "delete " & strLib & ".TASMAG"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
    End Sub

    Protected Sub SetPageProperties()

        With sheet2
            .ShowColumn(256, False)
            .SetColumnWidth(1, 15)
            .SetColumnWidth(2, 15)
            .SetColumnWidth(3, 15)
            .SetColumnWidth(4, 8)
            .SetColumnWidth(5, 8)
            .SetColumnWidth(6, 8)
            .SetColumnWidth(7, 8)
            .SetColumnWidth(8, 8)
            .SetColumnWidth(9, 8)
            .SetColumnWidth(10, 8)
            .SetColumnWidth(11, 11)
            .SetColumnWidth(12, 8)
            .SetColumnWidth(13, 8)
            .SetColumnWidth(14, 11)
            .SetColumnWidth(15, 9)
            For i As Integer = 6 To intcurrow
                .SetRowHeight(i, 15.75)
            Next
            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Landscape
            .PageSetup.LeftMargin = 0.5
            .PageSetup.RightMargin = 0.5
            .PageSetup.TopMargin = 0.5
            .PageSetup.BottomMargin = 0.5
            .PageSetup.Zoom = 95

        End With

    End Sub
End Class