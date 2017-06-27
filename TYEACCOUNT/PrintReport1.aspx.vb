Imports DataAccess
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Imports System.IO

Public Class WebForm5
    Inherits System.Web.UI.Page

    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/report/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("tpReport1.xlt")
    Private ReportName As String = "REPORT 1"
    Private crtDate As String = Date.Now.ToString("ddMMyyyy")
    Private crtTime As String = Date.Now.ToString("HHmm")
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = ExcelEngine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private intcurrow, intno, intstart As Integer
    Private filename As String
    Private strLib As String = "TYEACC"
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim str As String

        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qryReport1.lbs")
            str = rw.ReadToEnd
        End Using
        str = String.Format(str, strLib)
        Dim dt As New DataTable
        dt = DbConnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        ExportExcel(dt)
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
        sheet2 = Workbook.Worksheets.Create("Data")
        With sheet2
            Workbook.Worksheets(0).Range("A1:AZ1").CopyTo(.Range("A1"))
            '.Range(2, 6).Value = Trim(oTable.Rows(0)("IDESC")) & "-" & oTable.Rows(0)("IDSCE")
            '.Range(3, 1).Value = "AT " & String.Format("{0:dd-MMM-yyyy}", DateTime.Now)
            '.Range(3, 2).Value = "As " & Right(Trim(txtfromdate.Text), 2) & "/" & Mid(Trim(txtfromdate.Text), 5, 2) & "/" & Mid(Trim(txtfromdate.Text), 3, 2) & " - " & Right(Trim(txttodate.Text), 2) & "/" & Mid(Trim(txttodate.Text), 5, 2) & "/" & Mid(Trim(txttodate.Text), 3, 2) & ""
           
            For Each oRow In oTable.Rows
                Workbook.Worksheets(0).Range("A2:AZ2").CopyTo(.Range("A" & intcurrow))
                .Range(intcurrow, 1).Text = oRow("ITEMCLASS").ToString()
                .Range(intcurrow, 2).Text = "'" & oRow("ITEM").ToString()
                .Range(intcurrow, 3).Value = oRow("DESCRIPTION").ToString()
                .Range(intcurrow, 4).Value = oRow("LENGTH").ToString()
                .Range(intcurrow, 5).Value = oRow("CUSTCODE").ToString()
                .Range(intcurrow, 6).Value = oRow("CUSTNAME").ToString()
                .Range(intcurrow, 7).Value = oRow("DOCTYPE").ToString()
                .Range(intcurrow, 8).Value = oRow("INVOICENO").ToString()
                .Range(intcurrow, 9).Value = oRow("INVOICEDATE").ToString()
                .Range(intcurrow, 10).Value = oRow("ORDERDATE").ToString()
                .Range(intcurrow, 11).Value = oRow("REQDATE").ToString()
                .Range(intcurrow, 12).Value = oRow("ORDERTYPE").ToString()
                .Range(intcurrow, 13).Value = oRow("ORDERNO").ToString()
                .Range(intcurrow, 14).Value = oRow("MAQTY").ToString()
                .Range(intcurrow, 15).Value = ConvertDec(oRow("UNITPRICE"))
                .Range(intcurrow, 16).Value = ConvertDec(oRow("AMOUNT"))
                .Range(intcurrow, 17).Value = ConvertDec(oRow("UNITCOST"))
                .Range(intcurrow, 18).Value = ConvertDec(oRow("TOTALCOST"))
                .Range(intcurrow, 19).Value = ConvertDec(oRow("TOTALWEIGHT"))
                .Range(intcurrow, 20).Value = ConvertDec(oRow("PERCENTEXP"))
                .Range(intcurrow, 21).Value = ConvertDec(oRow("MACSTP"))
                .Range(intcurrow, 22).Value = oRow("MARGN").ToString()
                .Range(intcurrow, 23).Value = ConvertDec(oRow("MAPDCD"))
                .Range(intcurrow, 24).Value = ConvertDec(oRow("PROFIC"))
                .Range(intcurrow, 25).Value = oRow("ACCOUNTCODE").ToString()
                .Range(intcurrow, 26).Value = oRow("CUSTNAME").ToString()
                .Range(intcurrow, 27).Value = oRow("CUSTOM").ToString()
                .Range(intcurrow, 28).Value = oRow("CUSTGROUP").ToString()
                .Range(intcurrow, 29).Value = oRow("SECTION").ToString()
                .Range(intcurrow, 30).Value = oRow("MATNAME").ToString()
                intcurrow = intcurrow + 1
            Next
            
            ' Workbook.Worksheets(0).Range("A31:O38").CopyTo(.Range("A" & intcurrow))

           
        End With
        Call SetPageProperties()

        Workbook.Worksheets("Data").Range("A2:AZ10000").CopyTo(Workbook.Worksheets(0).Range("A2:AZ10000"))
        Workbook.Worksheets("Data").Remove()
        Workbook.Worksheets(0).Range("AF1").Value = "As of " & DateTime.Now.ToString("dd MMM yyyy")
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
        Return True
    End Function
    Protected Function ConvertDec(str As Object) As Decimal
        If str.ToString = "" Then
            Return 0
        End If
        Return CDec(str)
    End Function
    Protected Sub SetPageProperties()
       
                With sheet2
                    .SetColumnWidth(1, 8)
                    .SetColumnWidth(2, 9)
                    .SetColumnWidth(3, 9)
                    .SetColumnWidth(4, 8)
                    .SetColumnWidth(5, 8)
                    .SetColumnWidth(6, 11)
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
            '.PageSetup.RightFooter = "KPN-004-1"
                    '.PageSetup.PrintTitleRows = "$5:$15"
            '.Range("A7").FreezePanes()
                End With
            


    End Sub
End Class