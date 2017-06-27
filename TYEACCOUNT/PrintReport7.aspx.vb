Imports DataAccess
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Imports System.IO

Public Class WebForm10
    Inherits System.Web.UI.Page

    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/report/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("tpReport7.xlt")
    Private ReportName As String = "REPORT 7"
    Private crtDate As String = Date.Now.ToString("ddMMyyyy")
    Private crtTime As String = Date.Now.ToString("HHmm")
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = ExcelEngine.Excel
    Private Workbook As IWorkbook
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private DateMonth As Date
    Private intcurrow, intno, intstart, inttotwr, inttotdc, intend, intcustr, intcuend, intGrnTTC As Integer
    Private filename, Type, strsql, monName As String
    Private dt, oTable As New DataTable
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        ExportExcel()
    End Sub
    Protected Function GetData1() As DataTable
        strsql = String.Format("select * from {0}.TASMAS left join {0}.TACSMA on MACSCD=CSCSCD where maitcl='B1' and madodt between {1} and {2} order by madoct desc,maitcd", strLib, txtDateS.Text, txtDateE.Text)
        dt = DbConnect.ExcuteQueryString(strsql, DBConnection.DatabaseType.AS400)
        Return dt
    End Function
    Protected Function GetData2() As DataTable
        strsql = String.Format("select distinct macsnm from {0}.TASMAS left join {0}.TACSMA on MACSCD=CSCSCD where maitcl='B1' and madodt between {1} and {2} order by macsnm", strLib, txtDateS.Text, txtDateE.Text)
        dt = DbConnect.ExcuteQueryString(strsql, DBConnection.DatabaseType.AS400)
        Return dt
    End Function
    Protected Function GetData3() As DataTable
        strsql = String.Format("select distinct madesc from {0}.TASMAS left join {0}.TACSMA on MACSCD=CSCSCD where maitcl='B1' and madodt between {1} and {2} order by madesc", strLib, txtDateS.Text, txtDateE.Text)
        dt = DbConnect.ExcuteQueryString(strsql, DBConnection.DatabaseType.AS400)
        Return dt
    End Function
    Public Function ExportExcel()

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Dim oRow As DataRow
        Dim ConvertMonth As Integer

        oTable = GetData1()
        If oTable.Rows.Count <= 0 Then
            Return False
        End If

        ConvertMonth = Left(txtDateS.Text, 6)
        DateMonth = DateTime.ParseExact(ConvertMonth, "yyyyMM", Nothing)
        monName = MonthName(DateMonth.Month, False)
        oRow = oTable.Rows(0)
        intcurrow = 4
        intstart = intcurrow

        appExcel.DefaultFilePath = Server.MapPath(".")
        Workbook = appExcel.Workbooks.Open(TemplateFile)
        sheet = Workbook.Worksheets(0)
        sheet2 = Workbook.Worksheets.Create("Data")
        Type = oTable.Rows(0).Item("madoct").ToString()

        With sheet2
            Workbook.Worksheets(0).Range("A1:J3").CopyTo(.Range("A1"))
            .Range("A2").Value = "CU W/R AS AT " & monName.ToUpper & " " & Left(txtDateS.Text, 4)

            For Each oRow In oTable.Rows
                If Type <> oRow("madoct").ToString() Then 'SUB TOTAL WIRE ROD SALE
                    Workbook.Worksheets(0).Range("A6:J6").CopyTo(.Range("A" & intcurrow))
                    .Range(intcurrow, 6).Value = "=SUM(F" & intstart & ":F" & intcurrow - 1 & ")"
                    .Range(intcurrow, 8).Value = "=SUM(H" & intstart & ":H" & intcurrow - 1 & ")"
                    .Range(intcurrow, 9).Value = "=SUM(I" & intstart & ":I" & intcurrow - 1 & ")"
                    .Range(intcurrow, 10).Value = "=SUM(J" & intstart & ":J" & intcurrow - 1 & ")"
                    Type = oRow("madoct").ToString()
                    inttotwr = intcurrow
                    intcurrow = intcurrow + 1
                    intstart = intcurrow
                End If

                Workbook.Worksheets(0).Range("A4:J4").CopyTo(.Range("A" & intcurrow))
                .Range(intcurrow, 1).Text = oRow("maitcl").ToString()                       'Class
                .Range(intcurrow, 2).Text = oRow("maitcd").ToString()                       'Item Code
                .Range(intcurrow, 3).Value = oRow("madesc").ToString()                      'Description
                .Range(intcurrow, 4).Value = oRow("macsnm").ToString()                      'Customer
                .Range(intcurrow, 5).Value = oRow("mauprc").ToString()                      'Unit Sale price
                .Range(intcurrow, 6).Value = oRow("maamt").ToString()                       'Amount 
                .Range(intcurrow, 7).Value = oRow("maucst").ToString()                      'Unit Cost
                .Range(intcurrow, 8).Value = oRow("matcst").ToString()                      'Total Cost
                .Range(intcurrow, 9).Value = oRow("matwgh").ToString()                      'Total Weight
                intcurrow = intcurrow + 1
            Next

            Workbook.Worksheets(0).Range("A7:J9").CopyTo(.Range("A" & intcurrow)) 'SUB TOTAL DEBIT/CREDIT NOTE/DISCOUNT
            .Range(intcurrow, 6).Value = "=SUM(F" & intstart & ":F" & intcurrow - 1 & ")"
            .Range(intcurrow, 8).Value = "=SUM(H" & intstart & ":H" & intcurrow - 1 & ")"
            .Range(intcurrow, 9).Value = "=SUM(I" & intstart & ":I" & intcurrow - 1 & ")"
            .Range(intcurrow, 10).Value = "=SUM(J" & intstart & ":J" & intcurrow - 1 & ")"
            inttotdc = intcurrow
            '  intend = intcurrow
            .Range(intcurrow + 2, 6).Value = "=+F" & inttotwr & "+F" & inttotdc 'GRAND TOTAL
            .Range(intcurrow + 2, 8).Value = "=+H" & inttotwr & "+H" & inttotdc
            .Range(intcurrow + 2, 9).Value = "=+I" & inttotwr & "+I" & inttotdc
            .Range(intcurrow + 2, 10).Value = "=+J" & inttotwr & "+J" & inttotdc

            intcurrow = intcurrow + 5

            'TOTAL BY CUSTOMER
            oTable = GetData2()

            If oTable.Rows.Count >= 0 Then
                Workbook.Worksheets(0).Range("C12:J12").CopyTo(.Range("C" & intcurrow))
                intcurrow += 1
                intstart = intcurrow

                For n As Integer = 0 To oTable.Rows.Count - 1
                    oRow = oTable.Rows(n)
                    Workbook.Worksheets(0).Range("C14:J14").CopyTo(.Range("C" & intcurrow))

                    If n = 0 Then .Range(intcurrow, 3).Text = "TOTAL BY CUSTOMER"
                    .Range(intcurrow, 4).Text = oRow("macsnm").ToString()
                    .Range(intcurrow, 6).Value = "=SUMIF($D$4:$D$" & inttotdc & ",D" & intcurrow & ",$F$4:$F$" & inttotdc & ")"
                    .Range(intcurrow, 8).Value = "=SUMIF($D$4:$D$" & inttotdc & ",D" & intcurrow & ",$H$4:$H$" & inttotdc & ")"
                    .Range(intcurrow, 9).Value = "=SUMIF($D$4:$D$" & inttotdc & ",D" & intcurrow & ",$I$4:$I$" & inttotdc & ")"
                    .Range(intcurrow, 10).Value = "=SUMIF($D$4:$D$" & inttotdc & ",D" & intcurrow & ",$J$4:$J$" & inttotdc & ")"
                    intcurrow = intcurrow + 1
                Next

                'GRAND TOTAL BY CUSTOMER
                Workbook.Worksheets(0).Range("C15:J15").CopyTo(.Range("C" & intcurrow))
                .Range(intcurrow, 6).Value = "=SUM(F" & intstart & ":F" & intcurrow - 1 & ")"
                .Range(intcurrow, 8).Value = "=SUM(H" & intstart & ":H" & intcurrow - 1 & ")"
                .Range(intcurrow, 9).Value = "=SUM(I" & intstart & ":I" & intcurrow - 1 & ")"
                .Range(intcurrow, 10).Value = "=SUM(J" & intstart & ":J" & intcurrow - 1 & ")"
                intGrnTTC = intcurrow
                intcurrow = intcurrow + 1
            End If

            'TOTAL BY PRODUCT
            oTable = GetData3()

            If oTable.Rows.Count >= 0 Then
                Workbook.Worksheets(0).Range("C16:J17").CopyTo(.Range("C" & intcurrow))
                intcurrow = intcurrow + 1
                intstart = intcurrow

                For n As Integer = 0 To oTable.Rows.Count - 1
                    oRow = oTable.Rows(n)
                    Workbook.Worksheets(0).Range("C14:J14").CopyTo(.Range("C" & intcurrow))

                    If n = 0 Then
                        .Range(intcurrow, 3).Text = "TOTAL BY PRODUCT"
                        .Range(intcurrow, 3).CellStyle.Font.Color = ExcelKnownColors.Blue
                    End If

                    .Range(intcurrow, 4).Text = oRow("madesc").ToString()
                    .Range(intcurrow, 6).Value = "=SUMIF($C$4:$C$" & inttotdc & ",D" & intcurrow & ",$F$4:$F$" & inttotdc & ")"
                    .Range(intcurrow, 8).Value = "=SUMIF($C$4:$C$" & inttotdc & ",D" & intcurrow & ",$H$4:$H$" & inttotdc & ")"
                    .Range(intcurrow, 9).Value = "=SUMIF($C$4:$C$" & inttotdc & ",D" & intcurrow & ",$I$4:$I$" & inttotdc & ")"
                    .Range(intcurrow, 10).Value = "=SUMIF($C$4:$C$" & inttotdc & ",D" & intcurrow & ",$J$4:$J$" & inttotdc & ")"

                    If .Range(intcurrow, 4).Text.Trim = "CU W/R  1/11   MM" Or .Range(intcurrow, 4).Text.Trim = "CU W/R  1/8.0  MM" Or .Range(intcurrow, 4).Text.Trim = "CU W/R  1/8.0  MM (EXPORT)" Then
                        If intcustr = 0 Then intcustr = intcurrow
                        For i As Integer = 4 To 10
                            .Range(intcurrow, i).CellStyle.Font.Color = ExcelKnownColors.Blue
                        Next
                        intcuend = intcurrow
                    End If

                    intcurrow = intcurrow + 1
                Next

                'GRAND TOTAL BY PRODUCTS
                Workbook.Worksheets(0).Range("C19:J23").CopyTo(.Range("C" & intcurrow))
                .Range(intcurrow, 6).Value = "=SUBTOTAL(9,F" & intstart & ":F" & intcurrow - 1 & ")"
                .Range(intcurrow, 8).Value = "=SUBTOTAL(9,H" & intstart & ":H" & intcurrow - 1 & ")"
                .Range(intcurrow, 9).Value = "=SUBTOTAL(9,I" & intstart & ":I" & intcurrow - 1 & ")"
                .Range(intcurrow, 10).Value = "=SUBTOTAL(9,J" & intstart & ":J" & intcurrow - 1 & ")"

                'TOTAL : CU W/R 1/8.0,11.0 MM.
                .Range(intcurrow + 1, 6).Value = "=SUM(F" & intcustr & ":F" & intcuend & ")"
                .Range(intcurrow + 1, 8).Value = "=SUM(H" & intcustr & ":H" & intcuend & ")"
                .Range(intcurrow + 1, 9).Value = "=SUM(I" & intcustr & ":I" & intcuend & ")"
                .Range(intcurrow + 1, 10).Value = "=SUM(J" & intcustr & ":J" & intcuend & ")"

                'Check diff by costomer
                .Range(intcurrow + 3, 5).Value = "=+E" & inttotdc + 2 & "-E" & intGrnTTC
                .Range(intcurrow + 3, 6).Value = "=+F" & inttotdc + 2 & "-F" & intGrnTTC
                .Range(intcurrow + 3, 7).Value = "=+G" & inttotdc + 2 & "-G" & intGrnTTC
                .Range(intcurrow + 3, 8).Value = "=+H" & inttotdc + 2 & "-H" & intGrnTTC
                .Range(intcurrow + 3, 9).Value = "=+I" & inttotdc + 2 & "-I" & intGrnTTC
                .Range(intcurrow + 3, 10).Value = "=+J" & inttotdc + 2 & "-J" & intGrnTTC

                'Check diff by product
                .Range(intcurrow + 4, 5).Value = "=+E" & inttotdc + 2 & "-E" & intcurrow
                .Range(intcurrow + 4, 6).Value = "=+F" & inttotdc + 2 & "-F" & intcurrow
                .Range(intcurrow + 4, 7).Value = "=+G" & inttotdc + 2 & "-G" & intcurrow
                .Range(intcurrow + 4, 8).Value = "=+H" & inttotdc + 2 & "-H" & intcurrow
                .Range(intcurrow + 4, 9).Value = "=+I" & inttotdc + 2 & "-I" & intcurrow
                .Range(intcurrow + 4, 10).Value = "=+J" & inttotdc + 2 & "-J" & intcurrow
            End If

        End With

        'Workbook.Worksheets("Data").Range("A2:AZ200").CopyTo(Workbook.Worksheets(0).Range("A2:AZ200"))
        'Workbook.Worksheets("Data").Remove()
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
    Protected Sub SetPageProperties()

        With sheet2
            .SetColumnWidth(1, 8)
            .SetColumnWidth(2, 9)
            .SetColumnWidth(3, 35)
            .SetColumnWidth(4, 36)
            .SetColumnWidth(5, 8)
            .SetColumnWidth(6, 14)
            .SetColumnWidth(7, 8)
            .SetColumnWidth(8, 14)
            .SetColumnWidth(9, 14)
            .SetColumnWidth(10, 14)

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