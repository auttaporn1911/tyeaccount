Imports DataAccess
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Imports System.IO
Imports System.Globalization

Public Class WebForm7
    Inherits System.Web.UI.Page
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/report/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("Reporttp2.xlt")
    Private ReportName As String = "REPORT 2"
    Private crtDate As String = Date.Now.ToString("ddMMyyyy")
    Private crtTime As String = Date.Now.ToString("HHmm")
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private Workbook As IWorkbook
    Private _gc As GregorianCalendar = New GregorianCalendar

    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private sheet3 As IWorksheet
    Private sheetPrint As IWorksheet
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
    Public Sub Confirm(ByVal message As String)
        'ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("return confirm('{0}');", message), True)
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            ShowClassUnmatch()
        End If

    End Sub

    Protected Sub ShowClassUnmatch()
        Dim str As String
        str = "select distinct SPITCL from " & strLib & ".TASPSA " & _
                "left join " & strLib & ".TACLMA on SPITCL = CLCLCD where CLCLCD is null "
        Dim dt As New DataTable
        dt = DbConnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        GridView1.DataSource = dt
        GridView1.DataBind()

    End Sub
    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub
    Protected Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim str As String

        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qryReport2.lbs")
            str = rw.ReadToEnd
        End Using
        Dim cmd As String
        Dim tMonth As Integer

        Dim dt As New DataTable
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim dateS As Date
        Dim dateE As Date


        dateS = Date.ParseExact(txtDateS.Text, "yyyyMMdd", provider)
        dateE = Date.ParseExact(txtDateE.Text, "yyyyMMdd", provider)
        If dateS > dateE Then
            MessageBox("Warning : Date Start is more than Date End.")
            Exit Sub
        End If
        tMonth = DateDiff(DateInterval.Month, dateS, dateE) + 1
        If tMonth > 6 Then
            cmd = String.Format(str, dateS.ToString("yyyyMM"), dateS.AddMonths(5).ToString("yyyyMM"), strLib)
            dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.SQL)

            ExportExcel(dt, 6, 1)
            tMonth = tMonth - 6
            cmd = String.Format(str, dateS.AddMonths(6).ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
            dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.SQL)
            ExportExcel(dt, tMonth, 2)
        Else
            cmd = String.Format(str, dateS.ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
            dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.SQL)
            ExportExcel(dt, tMonth, 1)
        End If
        Workbook.Worksheets(0).Remove()
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
    End Sub


    Public Function SetColumnMonthName(dateS As Date, totalMonth As Integer)

        With sheetPrint
            For i As Integer = 0 To totalMonth - 1
                .Range(3, (i * 25) + 10).Value = "'" & dateS.ToString("MMMM", CultureInfo.InvariantCulture) & " " & dateS.Year.ToString
                dateS = dateS.AddMonths(1)
            Next

        End With

    End Function

    Public Function ExportExcel(oTable As DataTable, tMonth As Integer, sheetNo As Integer) As Boolean
        Dim oRow As DataRow
        Dim datatypeid As Integer
        Dim cmonth As Integer
        Dim classname As String = ""
        Dim group As String = ""
        Dim custom As String = ""
        Dim MatID As Integer = 0
        Dim subDummy As Integer = 300
        Dim GrandDummy As Integer = 310
        Dim Summary As Integer = 320
        Dim memRow As Integer
        'Dim dummy As Integer = 265
        Dim dummy2 As Integer = 275
        Dim rowcd As Integer = 200
        Dim strCol As String
        Dim extMonth As Integer
        Dim itemname As String = ""
        Dim isSubTotal As Boolean
        Dim isInitial As Boolean = False
        Dim isGrandTotal As Boolean
        Dim isFirstSummary As Boolean = True
        Dim isFisrtRecord As Boolean = True
        Dim isFisrtSheet2 As Boolean = True
        Dim lastRecord As Integer
        Dim dt As New DataTable
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim dateS As Date


        isSubTotal = False
        isGrandTotal = False


        If oTable.Rows.Count <= 0 Then
            Return False
        End If


        oRow = oTable.Rows(0)
        intcurrow = 5

        appExcel.DefaultFilePath = Server.MapPath(".")


        strCol = Utility.GetExcelColumnName((tMonth * 25) + 3)

        If sheetNo = 1 Then
            Workbook = appExcel.Workbooks.Open(TemplateFile, ExcelOpenType.Automatic)

            sheet = Workbook.Worksheets(0)
            sheetPrint = Workbook.Worksheets.Create("Data")
        Else
            sheet = Workbook.Worksheets(0)
            sheetPrint = Workbook.Worksheets.Create("Data2")
        End If
        With sheetPrint
            Workbook.Worksheets(0).Range("A1:" & strCol & "5").CopyTo(.Range("A1"))
            dateS = Date.ParseExact(txtDateS.Text, "yyyyMMdd", provider)
            If sheetNo = 2 Then
                dateS = dateS.AddMonths(6)
            End If
            SetColumnMonthName(dateS, tMonth)
            Workbook.Worksheets(0).Range("A18:" & strCol & "18").CopyTo(sheet.Range("A" & subDummy))

            Workbook.Worksheets(0).Range("A14:" & strCol & "16").CopyTo(sheet.Range("A" & Summary))
            Workbook.Worksheets(0).Range("A19:" & strCol & "19").CopyTo(sheet.Range("A" & GrandDummy))

            For Each oRow In oTable.Rows
                isSubTotal = False
                isGrandTotal = False
                If classname <> oRow("CLASS").ToString().Trim Then
                    If classname.Trim = "CD" Then

                    Else
                        If intcurrow = rowcd Then
                            intcurrow = memRow
                        End If

                        intcurrow = intcurrow + 1
                    End If

                    Workbook.Worksheets(0).Range("A7:" & strCol & "7").CopyTo(.Range("A" & intcurrow))

                    If group <> CStr(oRow("GROUP")).ToString And classname.ToUpper <> "CD" Then
                        If isFisrtRecord Then
                            '.Range(intcurrow, 1).Text = oRow("GROUP").ToString()

                        Else
                            isGrandTotal = True
                        End If

                        group = CStr(oRow("GROUP")).ToString


                    End If

                    If custom <> oRow("ITEMTYPE").ToString And classname.ToUpper <> "CD" Then
                        .Range(intcurrow, 1).Text = "'" & oRow("ITEMTYPE").ToString()

                        custom = oRow("ITEMTYPE").ToString

                        If isFisrtRecord = False Then
                            isSubTotal = True
                        End If

                    End If

                    classname = oRow("CLASS").ToString().ToUpper.Trim
                    '.Range(intcurrow, 2).Value = oRow("ITEM_NAME").ToString()
                    '.Range(intcurrow, 3).Value = oRow("CLASS").ToString()
                    '.Range(intcurrow, 256).Value = oRow("CLMTMN").ToString().Trim
                ElseIf itemname.Trim <> oRow("ITEM_NAME").ToString().Trim Then
                    intcurrow = intcurrow + 1
                    Workbook.Worksheets(0).Range("A7:" & strCol & "7").CopyTo(.Range("A" & intcurrow))
                End If



                itemname = oRow("ITEM_NAME")
                datatypeid = CInt(oRow("DATATYPEID").ToString())
                cmonth = CInt(oRow("MONTH").ToString())

                If sheetNo = 1 Then
                    If cmonth < 7 Then
                        cmonth = cmonth + 5 - (IIf(dateS.Month < 7, dateS.Month - 7 + 12, dateS.Month - 7))
                    Else
                        cmonth = cmonth - 7 - (dateS.Month - 7)
                    End If
                Else
                    If isFisrtSheet2 Then
                        extMonth = CInt(oRow("MONTH").ToString())
                        isFisrtSheet2 = False
                        cmonth = 0
                    Else
                        cmonth = cmonth - extMonth
                    End If
                    ''''extMonth =
                    '''''cmonth = 
                End If


                If isSubTotal Then
                    sheet.Range("A" & subDummy.ToString & ":" & strCol & subDummy.ToString).CopyTo(.Range("A" & intcurrow.ToString))
                    Workbook.Worksheets(0).Range("A18:" & strCol & "18").CopyTo(sheet.Range("A" & subDummy))
                    If isGrandTotal = False Then
                        intcurrow += 1
                        Workbook.Worksheets(0).Range("A7:" & strCol & "7").CopyTo(.Range("A" & intcurrow))
                        .Range(intcurrow, 1).Text = "'" & oRow("ITEMTYPE").ToString()
                        .Range(intcurrow, 2).Text = "'" & oRow("ITEM_NAME").ToString()
                        .Range(intcurrow, 3).Value = oRow("CLASS").ToString()
                        .Range(intcurrow, 256).Value = oRow("CLMTMN").ToString().Trim
                    End If


                End If
                If isGrandTotal Then

                    intcurrow += 1
                    sheet.Range("A" & (GrandDummy).ToString & ":" & strCol & (GrandDummy).ToString).CopyTo(.Range("A" & intcurrow.ToString))
                    Workbook.Worksheets(0).Range("A19:" & strCol & "19").CopyTo(sheet.Range("A" & GrandDummy))
                    If classname.Trim <> "CD" Then
                        intcurrow += 1
                        Workbook.Worksheets(0).Range("A7:" & strCol & "7").CopyTo(.Range("A" & intcurrow))
                        .Range(intcurrow, 1).Text = oRow("ITEMTYPE").ToString()
                        .Range(intcurrow, 2).Text = "'" & oRow("ITEM_NAME").ToString()
                        .Range(intcurrow, 3).Value = oRow("CLASS").ToString()
                        .Range(intcurrow, 256).Value = oRow("CLMTMN").ToString().Trim

                    End If

                End If


                If classname.Trim = "CD" And intcurrow <> rowcd Then
                    memRow = intcurrow
                    rowcd = memRow + 50
                    intcurrow = rowcd
                    Workbook.Worksheets(0).Range("A39:" & strCol & "39").CopyTo(.Range("A" & intcurrow))

                End If


                .Range(intcurrow, 2).Value = oRow("ITEM_NAME").ToString()
                .Range(intcurrow, 3).Value = oRow("CLASS").ToString()
                .Range(intcurrow, 256).Value = oRow("CLMTMN").ToString().Trim


                .Range(intcurrow, (cmonth * 25) + ((5 * datatypeid) - 1)).Value = ConvertDec(oRow("QTY").ToString())
                .Range(intcurrow, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value = ConvertDec(oRow("AMOUNT").ToString())
                .Range(intcurrow, (cmonth * 25) + (5 * datatypeid - 1) + 2).Value = ConvertDec(oRow("COST").ToString())
                .Range(intcurrow, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value = ConvertDec(oRow("PROFIT").ToString())
                .Range(intcurrow, (cmonth * 25) + (5 * datatypeid - 1) + 4).Value = Percent(ConvertDec(oRow("PROFIT")), ConvertDec(oRow("AMOUNT")))


                sheet.Range(subDummy, 3).Value = "SUB TOTAL"
                sheet.Range(subDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value = ConvertDec(oRow("QTY")) + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value)
                sheet.Range(subDummy, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value = ConvertDec(oRow("AMOUNT").ToString()) + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value)
                sheet.Range(subDummy, (cmonth * 25) + (5 * datatypeid - 1) + 2).Value = ConvertDec(oRow("COST").ToString()) + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * datatypeid - 1) + 2).Value)
                sheet.Range(subDummy, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value = ConvertDec(oRow("PROFIT").ToString()) + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value)
                sheet.Range(subDummy, (cmonth * 25) + (5 * datatypeid - 1) + 4).Value = Percent(ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value), ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value))

                sheet.Range(GrandDummy, 3).Value = "TOTAL"
                sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value = ConvertDec(oRow("QTY")) + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value)
                sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value = ConvertDec(oRow("AMOUNT").ToString()) + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value)
                sheet.Range(GrandDummy, (cmonth * 25) + (5 * datatypeid - 1) + 2).Value = ConvertDec(oRow("COST").ToString()) + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * datatypeid - 1) + 2).Value)
                sheet.Range(GrandDummy, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value = ConvertDec(oRow("PROFIT").ToString()) + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value)
                sheet.Range(GrandDummy, (cmonth * 25) + (5 * datatypeid - 1) + 4).Value = Percent(ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value), ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value))

                If datatypeid = 2 Then
                    .Range(intcurrow, (cmonth * 25) + ((5 * 3) - 1)).Value = ConvertDec(.Range(intcurrow, (cmonth * 25) + ((5 * 2) - 1)).Value) - ConvertDec(.Range(intcurrow, (cmonth * 25) + ((5 * 1) - 1)).Value)
                    .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 1).Value = ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 2 - 1) + 1).Value) - ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 1 - 1) + 1).Value)
                    .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 2).Value = ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 2 - 1) + 2).Value) - ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 1 - 1) + 2).Value)
                    .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 3).Value = ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 2 - 1) + 3).Value) - ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 1 - 1) + 3).Value)
                    .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 4).Value = Percent(ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 3).Value), ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 1).Value))

                    sheet.Range(subDummy, 3).Value = "SUB TOTAL"
                    sheet.Range(subDummy, (cmonth * 25) + ((5 * 3) - 1)).Value = .Range(intcurrow, (cmonth * 25) + ((5 * 3) - 1)).Value + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * 3) - 1)).Value)
                    sheet.Range(subDummy, (cmonth * 25) + ((5 * 3) - 1) + 1).Value = .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 1).Value + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * 3) - 1) + 1).Value)
                    sheet.Range(subDummy, (cmonth * 25) + (5 * 3 - 1) + 2).Value = .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 2).Value + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * 3 - 1) + 2).Value)
                    sheet.Range(subDummy, (cmonth * 25) + (5 * 3 - 1) + 3).Value = .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 3).Value + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * 3 - 1) + 3).Value)
                    sheet.Range(subDummy, (cmonth * 25) + (5 * 3 - 1) + 4).Value = Percent(ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * 3 - 1) + 3).Value), ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * 3) - 1) + 1).Value))

                    sheet.Range(GrandDummy, 3).Value = "TOTAL"
                    sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 3) - 1)).Value = .Range(intcurrow, (cmonth * 25) + ((5 * 3) - 1)).Value + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 3) - 1)).Value)
                    sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 3) - 1) + 1).Value = .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 1).Value + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 3) - 1) + 1).Value)
                    sheet.Range(GrandDummy, (cmonth * 25) + (5 * 3 - 1) + 2).Value = .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 2).Value + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * 3 - 1) + 2).Value)
                    sheet.Range(GrandDummy, (cmonth * 25) + (5 * 3 - 1) + 3).Value = .Range(intcurrow, (cmonth * 25) + (5 * 3 - 1) + 3).Value + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * 3 - 1) + 3).Value)
                    sheet.Range(GrandDummy, (cmonth * 25) + (5 * 3 - 1) + 4).Value = Percent(ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * 3 - 1) + 3).Value), ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 3) - 1) + 1).Value))

                End If

                If datatypeid = 4 Then
                    .Range(intcurrow, (cmonth * 25) + ((5 * 5) - 1)).Value = ConvertDec(.Range(intcurrow, (cmonth * 25) + ((5 * 2) - 1)).Value) - ConvertDec(.Range(intcurrow, (cmonth * 25) + ((5 * 4) - 1)).Value)
                    .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 1).Value = ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 2 - 1) + 1).Value) - ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 4 - 1) + 1).Value)
                    .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 2).Value = ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 2 - 1) + 2).Value) - ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 4 - 1) + 2).Value)
                    .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 3).Value = ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 2 - 1) + 3).Value) - ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 4 - 1) + 3).Value)
                    .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 4).Value = Percent(ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 3).Value), ConvertDec(.Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 1).Value))


                    sheet.Range(subDummy, 3).Value = "SUB TOTAL"
                    sheet.Range(subDummy, (cmonth * 25) + ((5 * 5) - 1)).Value = .Range(intcurrow, (cmonth * 25) + ((5 * 5) - 1)).Value + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * 5) - 1)).Value)
                    sheet.Range(subDummy, (cmonth * 25) + ((5 * 5) - 1) + 1).Value = .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 1).Value + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * 5) - 1) + 1).Value)
                    sheet.Range(subDummy, (cmonth * 25) + (5 * 5 - 1) + 2).Value = .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 2).Value + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * 5 - 1) + 2).Value)
                    sheet.Range(subDummy, (cmonth * 25) + (5 * 5 - 1) + 3).Value = .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 3).Value + ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * 5 - 1) + 3).Value)
                    sheet.Range(subDummy, (cmonth * 25) + (5 * 5 - 1) + 4).Value = Percent(ConvertDec(sheet.Range(subDummy, (cmonth * 25) + (5 * 5 - 1) + 3).Value), ConvertDec(sheet.Range(subDummy, (cmonth * 25) + ((5 * 5) - 1) + 1).Value))

                    sheet.Range(GrandDummy, 3).Value = "TOTAL"
                    sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 5) - 1)).Value = .Range(intcurrow, (cmonth * 25) + ((5 * 5) - 1)).Value + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 5) - 1)).Value)
                    sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 5) - 1) + 1).Value = .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 1).Value + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 5) - 1) + 1).Value)
                    sheet.Range(GrandDummy, (cmonth * 25) + (5 * 5 - 1) + 2).Value = .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 2).Value + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * 5 - 1) + 2).Value)
                    sheet.Range(GrandDummy, (cmonth * 25) + (5 * 5 - 1) + 3).Value = .Range(intcurrow, (cmonth * 25) + (5 * 5 - 1) + 3).Value + ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * 5 - 1) + 3).Value)
                    sheet.Range(GrandDummy, (cmonth * 25) + (5 * 5 - 1) + 4).Value = Percent(ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + (5 * 5 - 1) + 3).Value), ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * 5) - 1) + 1).Value))

                End If

                If group.ToUpper = "TYNS" Then

                End If
                'sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value = ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value) + ConvertDec(sheet.Range(MatID + dummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value)

                isFisrtRecord = False
            Next
            If intcurrow = rowcd Then
                intcurrow = memRow
            End If
            lastRecord = intcurrow

            If classname <> "CD" Then
                intcurrow += 1
                sheet.Range("A" & subDummy.ToString & ":" & strCol & "" & subDummy.ToString).CopyTo(.Range("A" & intcurrow.ToString))
                'For i As Integer = 1 To totalMat - 1
                '    intcurrow += 1
                '    sheet.Range("A" & (dummy + i).ToString & ":" & strCol & "" & (dummy + i).ToString).CopyTo(.Range("A" & intcurrow.ToString))

                'Next
                intcurrow += 1
                sheet.Range("A" & (GrandDummy).ToString & ":" & strCol & "" & (GrandDummy).ToString).CopyTo(.Range("A" & intcurrow.ToString))
            End If

            

            intcurrow += 1
            sheet.Range("A30:" & strCol & "39").CopyTo(.Range("A" & intcurrow))
            Dim sumStart As Integer
            Dim endMonth As Integer
            sumStart = intcurrow

            If tMonth > 10 Then
                endMonth = 9
            Else
                endMonth = tMonth - 1
            End If
            

            For x As Integer = 0 To endMonth
                intcurrow = sumStart
                For i As Integer = 0 To 9
                    For d As Integer = 1 To 5

                        .Range(intcurrow, (x * 25) + ((5 * d) - 1)).Value = "=SUMIF($IV$6:$IV$" & lastRecord & ",B" _
                        & intcurrow & ",$" _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1)) & "$6:$" _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1)) & "$" & lastRecord & ")"

                        .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 1)).Value = "=SUMIF($IV$6:$IV$" & lastRecord & ",B" _
                        & intcurrow & ",$" _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 1)) & "$6:$" _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 1)) & "$" & lastRecord & ")"

                        .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 2)).Value = "=SUMIF($IV$6:$IV$" & lastRecord & ",B" _
                        & intcurrow & ",$" _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 2)) & "$6:$" _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 2)) & "$" & lastRecord & ")"

                        .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 3)).Value = "=SUMIF($IV$6:$IV$" & lastRecord & ",B" _
                        & intcurrow & ",$" _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 3)) & "$6:$" _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 3)) & "$" & lastRecord & ")"

                        '.Range(intcurrow, (x * 25) + ((5 * d) - 1 + 4)).Value = Percent(ConvertDec(.Range(intcurrow, (x * 25) + ((5 * d) - 1 + 3)).Value), ConvertDec(.Range(intcurrow, (x * 25) + ((5 * d) - 1 + 1)).Value))

                        '.Range(intcurrow, (x * 25) + ((5 * d) - 1 + 4)).Value = "=SUMIF($IV$6:$IV$" & lastRecord & ",B" _
                        '& intcurrow & ",$" _
                        '& Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 4)) & "$6:$" _
                        '& Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 4)) & "$" & lastRecord & ")"
                    Next
                    intcurrow += 1
                Next
                If isFirstSummary Then
                    sheet.Range("A" & (Summary).ToString & ":" & strCol & "" & (Summary + 2).ToString).CopyTo(.Range("A" & intcurrow.ToString))
                    isFirstSummary = False

                End If

                For d As Integer = 1 To 5
                    '============Electric Wire=====================================================
                    .Range(intcurrow, (x * 25) + ((5 * d) - 1)).Value = "=SUM(SUMIF($IV$6:$IV$" & lastRecord & ",$D$1," _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1)) & "$6:$" _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1)) & "$" & lastRecord _
                       & "),SUMIF($IV$6:$IV$" & lastRecord & ",$D$2," _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1)) & "$6:$" _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1)) & "$" & lastRecord & "))"

                    .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 1)).Value = "=SUM(SUMIF($IV$6:$IV$" & lastRecord & ",$D$1," _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 1)) & "$6:$" _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 1)) & "$" & lastRecord _
                       & "),SUMIF($IV$6:$IV$" & lastRecord & ",$D$2," _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 1)) & "$6:$" _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 1)) & "$" & lastRecord & "))"

                    .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 2)).Value = "=SUM(SUMIF($IV$6:$IV$" & lastRecord & ",$D$1," _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 2)) & "$6:$" _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 2)) & "$" & lastRecord _
                       & "),SUMIF($IV$6:$IV$" & lastRecord & ",$D$2," _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 2)) & "$6:$" _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 2)) & "$" & lastRecord & "))"

                    .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 3)).Value = "=SUM(SUMIF($IV$6:$IV$" & lastRecord & ",$D$1," _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 3)) & "$6:$" _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 3)) & "$" & lastRecord _
                       & "),SUMIF($IV$6:$IV$" & lastRecord & ",$D$2," _
                        & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 3)) & "$6:$" _
                       & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 3)) & "$" & lastRecord & "))"

                    ' .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 4)).Value = Percent(ConvertDec(.Range(intcurrow, (x * 25) + ((5 * d) - 1 + 3)).Value), ConvertDec(.Range(intcurrow, (x * 25) + ((5 * d) - 1 + 1)).Value))

                    '.Range(intcurrow, (x * 25) + ((5 * d) - 1 + 4)).Value = "=SUM(SUMIF($IV$6:$IV$" & lastRecord & ",$D$1," _
                    '   & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 4)) & "$6:$" _
                    '   & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 4)) & "$" & lastRecord _
                    '   & "),SUMIF($IV$6:$IV$" & lastRecord & ",$D$2," _
                    '    & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 4)) & "$6:$" _
                    '   & Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 4)) & "$" & lastRecord & "))"
                    ' =SUM(SUMIF($IV$6:$IV$87,"CO*",$D$6:$D$87),SUMIF($IV$6:$IV$87,"AL*",$D$6:$D$87))
                    '============End Electric Wire =========================================================
                    '------Grand Total : Electric Wire + Digital Tachograph-------------------------------
                    '.Range(intcurrow + 2, (x * 25) + ((5 * d) - 1)).Value = .Range(intcurrow, (x * 25) + ((5 * d) - 1)).Value _
                    '    + .Range(intcurrow + 1, (x * 25) + ((5 * d) - 1)).Value

                    '.Range(intcurrow + 2, (x * 25) + ((5 * d) - 1 + 1)).Value = .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 1)).Value _
                    '    + .Range(intcurrow + 1, (x * 25) + ((5 * d) - 1 + 1)).Value
                    '.Range(intcurrow + 2, (x * 25) + ((5 * d) - 1 + 2)).Value = .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 2)).Value _
                    '    + .Range(intcurrow + 1, (x * 25) + ((5 * d) - 1 + 2)).Value

                    '.Range(intcurrow + 2, (x * 25) + ((5 * d) - 1 + 3)).Value = .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 3)).Value _
                    '    + .Range(intcurrow + 1, (x * 25) + ((5 * d) - 1 + 3)).Value

                    '.Range(intcurrow + 2, (x * 25) + ((5 * d) - 1 + 4)).Value = Percent(ConvertDec(.Range(intcurrow + 2, (x * 25) + ((5 * d) - 1 + 3)).Value), ConvertDec(.Range(intcurrow + 2, (x * 25) + ((5 * d) - 1 + 1)).Value))

                    '.Range(intcurrow + 2, (x * 25) + ((5 * d) - 1 + 4)).Value = .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 4)).Value _
                    '    + .Range(intcurrow + 1, (x * 25) + ((5 * d) - 1 + 4)).Value
                Next

            Next

            '----Copy Cash Discount---------------------------------------------------------
            If rowcd = 200 Then
                Workbook.Worksheets(0).Range("A39:" & strCol & "39").CopyTo(.Range("A" & rowcd))
            End If
            .Range("D" & rowcd & ":" & strCol & rowcd).CopyTo(.Range("D" & intcurrow - 1))
            .Range("A" & rowcd + 1 & ":" & strCol & rowcd + 1).CopyTo(.Range("A" & rowcd & ":" & strCol & rowcd))
            '------------------------------------------------------------------------



            'intcurrow += 1



        End With
        Call SetPageProperties()
        If sheetNo = 1 Then
            sheet2 = sheetPrint
        Else
            sheet3 = sheetPrint
        End If
        'Workbook.Worksheets("Data").Range("A2:AZ65000").CopyTo(Workbook.Worksheets(0).Range("A2:AZ65000"))
        ' Workbook.Worksheets(0).Remove()

        Return True
    End Function

    Protected Function Percent(Profit As Decimal, Amount As Decimal) As Decimal
        If Amount = 0 Then
            Percent = 0
        Else
            Percent = Profit / Amount
        End If
    End Function

    Protected Function ConvertDec(str As Object) As Decimal
        If str.ToString.Trim = "" Then
            Return 0
        End If
        Return CDec(str)
    End Function
    Protected Sub SetPageProperties()

        With sheetPrint
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

    Public Function GetWeekOfMonth(ByVal xYear As Integer, ByVal xMonth As Integer, ByVal xDay As Integer) As Integer
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim dateS As Date
        dateS = Date.ParseExact(xYear.ToString & xMonth.ToString.PadLeft(2, "0") & xDay.ToString.PadLeft(2, "0"), "yyyyMMdd", provider)
        Dim first As Date = New Date(dateS.Year, dateS.Month, 1)



        GetWeekOfMonth = GetWeekofYear(dateS) - GetWeekofYear(first) + 1
        Return GetWeekOfMonth

    End Function

    Private Function GetWeekofYear(time As Date) As Integer
        Return _gc.GetWeekOfYear(time, CalendarWeekRule.FirstDay, DayOfWeek.Sunday)
    End Function

    'Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    GetWeekOfMonth(2017, 4, 1)
    '    GetWeekOfMonth(2017, 4, 2)

    '    GetWeekOfMonth(2017, 4, 3)
    '    GetWeekOfMonth(2017, 4, 30)
    'End Sub
End Class