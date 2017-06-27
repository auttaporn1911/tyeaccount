Imports DataAccess
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Imports System.IO
Imports System.Drawing
Public Class WebForm11
    Inherits System.Web.UI.Page

    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/report/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("tpReport6.xlt")
    Private ReportName As String = "REPORT 6"
    Private crtDate As String = Date.Now.ToString("ddMMyyyy")
    Private crtTime As String = Date.Now.ToString("HHmm")
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
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

        CallQuery()

    End Sub

    Public Function CallQuery() As Boolean
        If txtDateS.Text.ToString = "" Then
            txtDateS.Focus()
            AlertMessagebox("Please choose Start Date")
            Return False
        End If

        If txtDateE.Text.ToString = "" Then
            AlertMessagebox("Please choose End Date")
            txtDateE.Focus()
            Return False
        End If
        Dim str As String
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qryReport6.lbs")
            str = String.Format(rw.ReadToEnd, txtDateS.Text, txtDateE.Text, strLib)
        End Using

        Dim dt As New DataTable
        dt = DbConnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        ExportExcel(dt)

    End Function

    Public Function ExportExcel(oTable As DataTable) As Boolean
        Dim oRow As DataRow
        Dim DType As String
        Dim Quantity, Amount, Profit, TotalC, TotalW, SUnit, SAmount, SProfit, STotalC, STotalW As Decimal
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        'Quantity = 0
        'Amount = 0
        'Profit = 0
        'TotalC = 0
        'TotalW = 0
        Dim rowStartCal As Integer
        DType = "Start"
        If oTable.Rows.Count <= 0 Then
            AlertMessagebox("Data not found.")
            Return False

        End If
        oRow = oTable.Rows(0)

        intcurrow = 5
        intstart = 0
        rowStartCal = 5
        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create("Data")
        With sheet2
            Workbook.Worksheets(0).Range("A1:AZ1").CopyTo(.Range("A1"))
            Workbook.Worksheets(0).Range("A2:AZ2").CopyTo(.Range("A2"))
            Workbook.Worksheets(0).Range("A3:AZ3").CopyTo(.Range("A3"))
            Workbook.Worksheets(0).Range("A4:AZ4").CopyTo(.Range("A4"))

            For Each oRow In oTable.Rows
                If oRow("MADOCT").ToString() = DType Or DType = "Start" Then
                    Workbook.Worksheets(0).Range("A5:AZ5").CopyTo(.Range("A" & intcurrow))
                    .Range(intcurrow, 1).Text = oRow("MAITCL").ToString()
                    .Range(intcurrow, 2).Text = oRow("MADESC").ToString()
                    .Range(intcurrow, 3).Value = oRow("MALEN").ToString()
                    .Range(intcurrow, 4).Value = "'" + oRow("MACSCD").ToString()
                    .Range(intcurrow, 5).Value = oRow("MACSNM").ToString()
                    .Range(intcurrow, 6).Value = oRow("MADOCT").ToString()
                    .Range(intcurrow, 7).Value = "'" + oRow("MADOCN").ToString()
                    .Range(intcurrow, 8).Value = oRow("MAUPRC").ToString()
                    .Range(intcurrow, 9).Value = oRow("MAAMT").ToString()
                    .Range(intcurrow, 10).Value = oRow("MAPRFP").ToString()
                    .Range(intcurrow, 11).Value = oRow("MATCST").ToString()
                    .Range(intcurrow, 12).Value = oRow("MATWGH").ToString()
                    DType = oRow("MADOCT").ToString()
                    'Quantity = Decimal.Parse(Quantity) + Decimal.Parse(oRow("MAQTY").ToString())
                    'Amount = Decimal.Parse(Amount) + Decimal.Parse(oRow("MAAMT").ToString())
                    'Profit = Decimal.Parse(Profit) + Decimal.Parse(oRow("MAPRFP").ToString())
                    'TotalC = Decimal.Parse(TotalC) + Decimal.Parse(oRow("MATCST").ToString())
                    'TotalW = Decimal.Parse(TotalW) + Decimal.Parse(oRow("MATWGH").ToString())
                    intcurrow = intcurrow + 1
                Else
                    Workbook.Worksheets(0).Range("A7:AZ7").CopyTo(.Range("A" & intcurrow))
                    .Range("A" & intcurrow & ":L" & intcurrow).CellStyle.Color = Color.FromArgb(191, 191, 191)
                    .Range(intcurrow, 5).RowHeight = 25
                    .Range(intcurrow, 5).Value = "Total : " + DType
                    If Quantity = 0 Then
                        Dim aa As String = ""
                    End If
                    'Dim tot1 As Double
                    'Try
                    '    tot1 = Quantity / Amount
                    'Catch e As DivideByZeroException
                    '    tot1 = 0
                    'End Try
                    '.Range(intcurrow, 8).Value = tot1
                    .Range(intcurrow, 9).Value = "=SUM(I" & rowStartCal & ":I" & intcurrow - 1 & ")"
                    str1 = str1 & "I" & intcurrow & ","
                    .Range(intcurrow, 10).Value = "=SUM(J" & rowStartCal & ":J" & intcurrow - 1 & ")"
                    str2 = str2 & "J" & intcurrow & ","
                    .Range(intcurrow, 11).Value = "=SUM(K" & rowStartCal & ":K" & intcurrow - 1 & ")"
                    str3 = str3 & "K" & intcurrow & ","
                    .Range(intcurrow, 12).Value = "=SUM(L" & rowStartCal & ":L" & intcurrow - 1 & ")"
                    str4 = str4 & "L" & intcurrow & ","
                    DType = oRow("MADOCT").ToString()
                    'SUnit += .Range(intcurrow, 8).Value
                    'SAmount += .Range(intcurrow, 9).Value
                    'SProfit += .Range(intcurrow, 10).Value
                    'STotalC += .Range(intcurrow, 11).Value
                    'STotalW += .Range(intcurrow, 12).Value
                    Quantity = 0
                    Amount = 0
                    Profit = 0
                    TotalC = 0
                    TotalW = 0
                    intcurrow = intcurrow + 1
                    '------------------------------------------------------
                    rowStartCal = intcurrow
                    Workbook.Worksheets(0).Range("A5:AZ5").CopyTo(.Range("A" & intcurrow))
                    .Range(intcurrow, 1).Text = oRow("MAITCL").ToString()
                    .Range(intcurrow, 2).Text = oRow("MADESC").ToString()
                    .Range(intcurrow, 3).Value = oRow("MALEN").ToString()
                    .Range(intcurrow, 4).Value = "'" + oRow("MACSCD").ToString()
                    .Range(intcurrow, 5).Value = oRow("MACSNM").ToString()
                    .Range(intcurrow, 6).Value = oRow("MADOCT").ToString()
                    .Range(intcurrow, 7).Value = "'" + oRow("MADOCN").ToString()
                    .Range(intcurrow, 8).Value = oRow("MAUPRC").ToString()
                    .Range(intcurrow, 9).Value = oRow("MAAMT").ToString()
                    .Range(intcurrow, 10).Value = oRow("MAPRFP").ToString()
                    .Range(intcurrow, 11).Value = oRow("MATCST").ToString()
                    .Range(intcurrow, 12).Value = oRow("MATWGH").ToString()
                    DType = oRow("MADOCT").ToString()
                    'Quantity = Decimal.Parse(Quantity) + Decimal.Parse(oRow("MAQTY").ToString())
                    'Amount = Decimal.Parse(Amount) + Decimal.Parse(oRow("MAAMT").ToString())
                    'Profit = Decimal.Parse(Profit) + Decimal.Parse(oRow("MAPRFP").ToString())
                    'TotalC = Decimal.Parse(TotalC) + Decimal.Parse(oRow("MATCST").ToString())
                    'TotalW = Decimal.Parse(TotalW) + Decimal.Parse(oRow("MATWGH").ToString())
                    intcurrow = intcurrow + 1

                End If
            Next

            Workbook.Worksheets(0).Range("A7:AZ7").CopyTo(.Range("A" & intcurrow))
            .Range("A" & intcurrow & ":L" & intcurrow).CellStyle.Color = Color.FromArgb(191, 191, 191)
            .Range(intcurrow, 5).Value = "Total :" + DType
            .Range(intcurrow, 5).RowHeight = 25
            'Dim tot2 As Double
            'Try
            '    tot2 = Quantity / Amount
            'Catch e As DivideByZeroException
            '    tot2 = 0
            'End Try
            '.Range(intcurrow, 8).Value = tot2
            .Range(intcurrow, 9).Value = "=SUM(I" & rowStartCal & ":I" & intcurrow - 1 & ")"
            str1 = str1 & "I" & intcurrow & ","
            .Range(intcurrow, 10).Value = "=SUM(J" & rowStartCal & ":J" & intcurrow - 1 & ")"
            str2 = str2 & "J" & intcurrow & ","
            .Range(intcurrow, 11).Value = "=SUM(K" & rowStartCal & ":K" & intcurrow - 1 & ")"
            str3 = str3 & "K" & intcurrow & ","
            .Range(intcurrow, 12).Value = "=SUM(L" & rowStartCal & ":L" & intcurrow - 1 & ")"
            str4 = str4 & "L" & intcurrow & ","
            DType = oRow("MADOCT").ToString()
            'SUnit += .Range(intcurrow, 8).Value
            'SAmount += .Range(intcurrow, 9).Value
            'SProfit += .Range(intcurrow, 10).Value
            'STotalC += .Range(intcurrow, 11).Value
            'STotalW += .Range(intcurrow, 12).Value
            Quantity = 0
            Amount = 0
            Profit = 0
            TotalC = 0
            TotalW = 0

            str1 = str1.Remove(str1.Length - 1, 1)
            str2 = str2.Remove(str2.Length - 1, 1)
            str3 = str3.Remove(str3.Length - 1, 1)
            str4 = str4.Remove(str4.Length - 1, 1)

            Workbook.Worksheets(0).Range("A7:AZ7").CopyTo(.Range("A" & intcurrow + 1))
            .Range("A" & intcurrow + 1 & ":L" & intcurrow + 1).CellStyle.Color = Color.FromArgb(191, 191, 191)
            .Range(intcurrow + 1, 5).Value = "Grand Total :"
            .Range(intcurrow + 1, 5).RowHeight = 25
            '.Range(intcurrow + 1, 8).Value = SUnit
            .Range(intcurrow + 1, 9).Value = "=SUM(" & str1 & ")"
            .Range(intcurrow + 1, 10).Value = "=SUM(" & str2 & ")"
            .Range(intcurrow + 1, 11).Value = "=SUM(" & str3 & ")"
            .Range(intcurrow + 1, 12).Value = "=SUM(" & str4 & ")"


        End With

        Call SetPageProperties()
        Workbook.Worksheets(0).Remove()
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

    Protected Sub AlertMessagebox(Message As String)
        Dim sb As New System.Text.StringBuilder()
        sb.Append("<script type = 'text/javascript'>")
        sb.Append("window.onload=function(){")
        sb.Append("alert('")
        sb.Append(Message)
        sb.Append("')};")
        sb.Append("</script>")
        ClientScript.RegisterClientScriptBlock(Me.GetType(), "alert", sb.ToString())
    End Sub

    Protected Sub SetPageProperties()

        With sheet2
            .Rows(0).RowHeight = 25
            .Rows(1).RowHeight = 25
            .Rows(2).RowHeight = 25
            .Range("A4:L4").CellStyle.Color = Color.FromArgb(230, 220, 241)
            .Range("A4:L4").CellStyle.Font.RGBColor = Color.FromArgb(22, 54, 92)
            .SetColumnWidth(1, 5)
            .SetColumnWidth(2, 18)
            .SetColumnWidth(3, 8)
            .SetColumnWidth(4, 13)
            .SetColumnWidth(5, 20)
            .SetColumnWidth(6, 5)
            .SetColumnWidth(7, 10)
            .SetColumnWidth(8, 12)
            .SetColumnWidth(9, 15)
            .SetColumnWidth(10, 15)
            .SetColumnWidth(11, 15)
            .SetColumnWidth(12, 15)
            .SetRowHeight(4, 25)
            .PageSetup.PaperSize = ExcelPaperSize.PaperA4
            .PageSetup.Orientation = ExcelPageOrientation.Landscape
            .PageSetup.PrintTitleRows = "$4:$4"
            .PageSetup.RightHeader = "Page &P of &N"
            .PageSetup.LeftMargin = 0.2
            .PageSetup.RightMargin = 0.2
            .PageSetup.TopMargin = 0.4
            .PageSetup.BottomMargin = 0.2
            .PageSetup.HeaderMargin = 0.1
            .PageSetup.Zoom = 90
        End With

    End Sub


End Class