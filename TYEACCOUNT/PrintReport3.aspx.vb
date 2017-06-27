Imports DataAccess
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Imports System.IO
Imports System.Globalization

Public Class WebForm12
    Inherits System.Web.UI.Page

    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/report/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("Reporttp3.xlt")
    Private ReportName As String = "REPORT 3"
    Private crtDate As String = Date.Now.ToString("ddMMyyyy")
    Private crtTime As String = Date.Now.ToString("HHmm")
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private Workbook As IWorkbook
    Private c_maxMonth As Integer
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
            LoadCheckMapping()
        End If

    End Sub
    Public Sub LoadCheckMapping()
        Dim str As String

        str = "select distinct A.CSCUST from {0}.TASPCS A left join {0}.TACSMA B "
        str &= "on trim(upper(A.CSCUST)) = trim(upper(B.CSCUTM)) where B.CSCUTM is null"
        str = String.Format(str, strLib)
        gvCheck.DataSource = DbConnect.ExcuteQueryString(str, DBConnection.DatabaseType.AS400)
        gvCheck.DataBind()
    End Sub
    Protected Function PrepareData() As Boolean
        Dim str As String
        Dim result As Integer
        str = String.Format("delete {0}.TACSCL", strLib)
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/tmpRpt3.lbs")
            str = String.Format(rw.ReadToEnd, strLib)
        End Using

        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
        Return True
    End Function
    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub
    Protected Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim str As String
        Dim strGetMonth As String
        Dim cmd2 As String
        strGetMonth = "select distinct ccmonm from {2}.TASMCC " & _
            "where (CCYEAR || RIGHT('0' || CCMONM,2)) between {0} and {1}"

        PrepareData()
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qryReport3.lbs")
            str = rw.ReadToEnd
        End Using
        Dim cmd As String
        Dim tMonth As Integer
        Dim dtMonth As New DataTable
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
            cmd2 = String.Format(strGetMonth, dateS.ToString("yyyyMM"), dateS.AddMonths(5).ToString("yyyyMM"), strLib)
            dtMonth = DbConnect.ExcuteQueryString(cmd2, DBConnection.DatabaseType.AS400)
            c_maxMonth = dtMonth.Rows.Count - 1
            ExportExcel(dt, 6, 1)
            tMonth = tMonth - 6
            cmd = String.Format(str, dateS.AddMonths(6).ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
            dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.SQL)
            cmd2 = String.Format(strGetMonth, dateS.AddMonths(6).ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
            dtMonth = DbConnect.ExcuteQueryString(cmd2, DBConnection.DatabaseType.AS400)
            c_maxMonth = dtMonth.Rows.Count - 1
            ExportExcel(dt, tMonth, 2)
        Else
            cmd = String.Format(str, dateS.ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
            dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.SQL)
            cmd2 = String.Format(strGetMonth, dateS.ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
            dtMonth = DbConnect.ExcuteQueryString(cmd2, DBConnection.DatabaseType.AS400)
            c_maxMonth = dtMonth.Rows(0)(0).ToString
            ExportExcel(dt, tMonth, 1)
        End If
        Workbook.Worksheets(0).Remove()
        filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
        Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
    End Sub
   


    Public Function SetColumnMonthName(dateS As Date, totalMonth As Integer)
        Dim total As String
        total = dateS.ToString("MMMM", CultureInfo.InvariantCulture) & " " & dateS.Year.ToString
        With sheetPrint
            For i As Integer = 0 To totalMonth - 1
                .Range(3, (i * 25) + 10).Value = "'" & dateS.ToString("MMMM", CultureInfo.InvariantCulture) & " " & dateS.Year.ToString
                dateS = dateS.AddMonths(1)
            Next
            dateS = dateS.AddMonths(-1)
            total = total & " TO " & dateS.ToString("MMMM", CultureInfo.InvariantCulture) & " " & dateS.Year.ToString
            .Range(3, (totalMonth * 25) + 10).Value = total
        End With

    End Function

    Public Function ExportExcel(oTable As DataTable, tMonth As Integer, sheetNo As Integer) As Boolean
        Dim list1 As New List(Of Integer)
        Dim rowCalStart As Integer
        Dim rowMatStart As Integer
        Dim rowGrandTotal As Integer
        Dim rowMatEnd As Integer
        Dim rowsub As Integer
        Dim oRow As DataRow
        Dim datatypeid As Integer
        Dim cmonth As Integer
        Dim orderno As Integer = 0
        Dim group As String = ""
        Dim custom As String = ""
        Dim MatID As Integer = 0
        Dim subDummy As Integer = 300
        Dim GrandDummy As Integer = 310
        Dim Summary As Integer = 320
        Dim firstmonth As Integer
        Dim dummy As Integer = 265
        Dim dummy2 As Integer = 275
        Dim strCol As String
        Dim extMonth As Integer
        Dim ColAmtCal As String
        Dim isSubTotal As Boolean
        Dim isInitial As Boolean = False
        Dim isGrandTotal As Boolean
        Dim totalMat As Integer
        Dim isFisrtRecord As Boolean = True
        Dim isFisrtSheet2 As Boolean = True
        Dim strsubCol As String
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
        rowMatStart = 6
        appExcel.DefaultFilePath = Server.MapPath(".")
        
        strCol = Utility.GetExcelColumnName((tMonth * 25) + 3 + 15)

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
           

            For Each oRow In oTable.Rows
                isSubTotal = False
                isGrandTotal = False
                If orderno <> CInt(oRow("OrderMat").ToString()) Then
                    orderno = CInt(oRow("OrderMat").ToString())
                    intcurrow = intcurrow + 1
                    Workbook.Worksheets(0).Range("A7:" & strCol & "7").CopyTo(.Range("A" & intcurrow))

                    MatID += 1
                    If group <> CStr(oRow("GROUP")).ToString Then
                        If isFisrtRecord Then
                            .Range(intcurrow, 1).Text = oRow("GROUP").ToString()

                        Else
                            isGrandTotal = True
                        End If

                        group = CStr(oRow("GROUP")).ToString


                    End If

                    If custom <> CStr(oRow("CUSTOM")).ToString Then
                        .Range(intcurrow, 2).Text = "'" & oRow("CUSTOM").ToString()

                        custom = CStr(oRow("CUSTOM")).ToString
                        totalMat = MatID
                        MatID = 1
                        If isFisrtRecord = False Then
                            isSubTotal = True
                        End If

                    End If

                    .Range(intcurrow, 3).Value = oRow("MATERIAL").ToString()

                End If

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
                    firstmonth = cmonth
                    'sheet.Range("A" & subDummy.ToString & ":" & strCol & subDummy.ToString).CopyTo(.Range("A" & intcurrow.ToString))
                    Workbook.Worksheets(0).Range("A18:" & strCol & "18").CopyTo(.Range("A" & intcurrow))
                    rowsub = intcurrow
                    .Range(rowsub, 3).Value = "Sub Total"
                    For i As Integer = cmonth To tMonth - 1
                        For d As Integer = 1 To 5
                            strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1))
                            .Range(rowsub, (i * 25) + ((5 * d) - 1)).Value = "=SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")"

                            strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 1)
                            .Range(rowsub, (i * 25) + ((5 * d) - 1) + 1).Value = "=SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")"
                            ColAmtCal = strsubCol
                            strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 2)
                            .Range(rowsub, (i * 25) + (5 * d - 1) + 2).Value = "=SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")"

                            strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 3)
                            .Range(rowsub, (i * 25) + (5 * d - 1) + 3).Value = "=SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")"

                            'strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 4)
                            '.Range(rowsub, (i * 25) + (5 * d - 1) + 4).Value = "=IF(ISERROR(SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")" & "/" & "SUM(" & ColAmtCal & rowCalStart & ":" & ColAmtCal & rowsub - 1 & ")),0,SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")" & "/" & "SUM(" & ColAmtCal & rowCalStart & ":" & ColAmtCal & rowsub - 1 & "))"
                            .Range(rowsub, (i * 25) + (5 * d - 1) + 4).Value = "=IF(ISERROR(" & strsubCol & rowsub & "/" & ColAmtCal & rowsub & "),0," & strsubCol & rowsub & "/" & ColAmtCal & rowsub & ")"
                            '-----------------------------------------------------------------------
                        Next
                    Next




                    If isGrandTotal = False Then
                        intcurrow += 1
                        Workbook.Worksheets(0).Range("A7:" & strCol & "7").CopyTo(.Range("A" & intcurrow))
                        .Range(intcurrow, 2).Text = "'" & oRow("CUSTOM").ToString()
                        .Range(intcurrow, 3).Value = oRow("MATERIAL").ToString()
                    End If


                End If
                If isGrandTotal Then
                    rowGrandTotal = intcurrow + 1
                    list1.Add(rowGrandTotal)
                    For i As Integer = 1 To totalMat - 1
                        intcurrow += 1
                        rowsub = intcurrow
                        rowMatEnd = intcurrow - 1
                        Workbook.Worksheets(0).Range("A" & 18 + i & ":" & strCol & 18 + i).CopyTo(.Range("A" & intcurrow))
                        '  .Range(rowsub, 3).Value = oRow("MATERIAL").ToString()
                        'sheet.Range("A" & (dummy + i).ToString & ":" & strCol & (dummy + i).ToString).CopyTo(.Range("A" & intcurrow.ToString))

                        For x As Integer = cmonth To tMonth - 1
                            For d As Integer = 1 To 5

                                strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1))
                                .Range(rowsub, (x * 25) + ((5 * d) - 1)).Value = "=SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                                    strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")"

                                strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 1)
                                .Range(rowsub, (x * 25) + ((5 * d) - 1) + 1).Value = "=SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                                   strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")"
                                ColAmtCal = strsubCol
                                strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 2)
                                .Range(rowsub, (x * 25) + (5 * d - 1) + 2).Value = "=SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                                   strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")"

                                strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 3)
                                .Range(rowsub, (x * 25) + (5 * d - 1) + 3).Value = "=SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                                   strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")"

                                'strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 4)
                                '.Range(rowsub, (x * 25) + (5 * d - 1) + 4).Value = "=IF(ISERROR(SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                                '   strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")" & "/" & "SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                                '   ColAmtCal & rowMatStart & ":" & ColAmtCal & rowMatEnd & ")" & "),0," & _
                                '"SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                                '   strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")" & "/" & "SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                                '   ColAmtCal & rowMatStart & ":" & ColAmtCal & rowMatEnd & "))"
                                '--------
                                .Range(rowsub, (x * 25) + (5 * d - 1) + 4).Value = "=IF(ISERROR(" & strsubCol & rowsub & "/" & ColAmtCal & rowsub & "),0," & strsubCol & rowsub & "/" & ColAmtCal & rowsub & ")"
                            Next
                        Next
                    Next




                    intcurrow += 1
                    Workbook.Worksheets(0).Range("A28:" & strCol & "28").CopyTo(.Range("A" & intcurrow))
                    rowsub = intcurrow
                    'sheet.Range("A" & (GrandDummy).ToString & ":" & strCol & (GrandDummy).ToString).CopyTo(.Range("A" & intcurrow.ToString))
                    .Range(rowsub, 3).Value = "Sub Total"
                    For i As Integer = cmonth To tMonth - 1
                        For d As Integer = 1 To 5
                            strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1))
                            .Range(rowsub, (i * 25) + ((5 * d) - 1)).Value = "=SUM(" & strsubCol & rowGrandTotal & ":" & strsubCol & rowsub - 1 & ")"

                            strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 1)
                            .Range(rowsub, (i * 25) + ((5 * d) - 1) + 1).Value = "=SUM(" & strsubCol & rowGrandTotal & ":" & strsubCol & rowsub - 1 & ")"
                            ColAmtCal = strsubCol
                            strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 2)
                            .Range(rowsub, (i * 25) + (5 * d - 1) + 2).Value = "=SUM(" & strsubCol & rowGrandTotal & ":" & strsubCol & rowsub - 1 & ")"

                            strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 3)
                            .Range(rowsub, (i * 25) + (5 * d - 1) + 3).Value = "=SUM(" & strsubCol & rowGrandTotal & ":" & strsubCol & rowsub - 1 & ")"

                            ' strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 4)
                            '.Range(rowsub, (i * 25) + ((5 * d) - 1 + 4)).Value = "=IF(ISERROR(SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")" & "/" & "SUM(" & ColAmtCal & rowCalStart & ":" & ColAmtCal & rowsub - 1 & ")),0,SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")" & "/" & "SUM(" & ColAmtCal & rowCalStart & ":" & ColAmtCal & rowsub - 1 & "))"
                            .Range(rowsub, (i * 25) + (5 * d - 1) + 4).Value = "=IF(ISERROR(" & strsubCol & rowsub & "/" & ColAmtCal & rowsub & "),0," & strsubCol & rowsub & "/" & ColAmtCal & rowsub & ")"


                            '-----------------------------------------------------------------------
                        Next
                    Next



                    intcurrow += 1
                    Workbook.Worksheets(0).Range("A7:" & strCol & "7").CopyTo(.Range("A" & intcurrow))
                    .Range(intcurrow, 1).Text = oRow("GROUP").ToString()
                    .Range(intcurrow, 2).Text = "'" & oRow("CUSTOM").ToString()
                    .Range(intcurrow, 3).Value = oRow("MATERIAL").ToString()
                    rowMatStart = intcurrow


                End If

                If MatID = 1 And isGrandTotal Then
                    isInitial = True
                ElseIf MatID > 1 Then
                    isInitial = False

                End If
                If MatID = 1 Then
                    'Initial Sub Total
                    rowCalStart = intcurrow
                    sheet.Range(subDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value = 0
                    sheet.Range(subDummy, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value = 0
                    sheet.Range(subDummy, (cmonth * 25) + (5 * datatypeid - 1) + 2).Value = 0
                    sheet.Range(subDummy, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value = 0
                    If isInitial Then
                        'Initial Grand Total
                        Workbook.Worksheets(0).Range("A28:" & strCol & "28").CopyTo(sheet.Range("A" & GrandDummy))
                        Workbook.Worksheets(0).Range("A20:" & strCol & "25").CopyTo(sheet.Range("A" & dummy))
                        Workbook.Worksheets(0).Range("A20:" & strCol & "25").CopyTo(sheet.Range("A" & dummy + 5))
                        sheet.Range(MatID + dummy, 2).Value = "TOTAL"
                        isInitial = False
                    End If



                End If

                .Range(intcurrow, (cmonth * 25) + ((5 * datatypeid) - 1)).Value = oRow("QTY").ToString()
                .Range(intcurrow, (cmonth * 25) + ((5 * datatypeid) - 1) + 1).Value = oRow("AMOUNT").ToString()
                .Range(intcurrow, (cmonth * 25) + (5 * datatypeid - 1) + 2).Value = oRow("COST").ToString()
                .Range(intcurrow, (cmonth * 25) + (5 * datatypeid - 1) + 3).Value = oRow("PROFIT").ToString()
                ColAmtCal = Utility.GetExcelColumnName((cmonth * 25) + ((5 * datatypeid) - 1) + 1)
                strsubCol = Utility.GetExcelColumnName((cmonth * 25) + ((5 * datatypeid) - 1) + 3)
                .Range(intcurrow, (cmonth * 25) + (5 * datatypeid - 1) + 4).Value = "=IF(ISERROR(" & strsubCol & intcurrow & "/" & ColAmtCal & intcurrow & "),0," & strsubCol & intcurrow & "/" & ColAmtCal & intcurrow & ")"


               
                If group.ToUpper = "TYNS" Then

                End If
                'sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value = ConvertDec(sheet.Range(GrandDummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value) + ConvertDec(sheet.Range(MatID + dummy, (cmonth * 25) + ((5 * datatypeid) - 1)).Value)

                isFisrtRecord = False
            Next
            '--------------------------------------------------------------------------
            '--------------------------------------------------------------------------
            '------------------------------------------------------------------------
            intcurrow += 1
            Workbook.Worksheets(0).Range("A18:" & strCol & "18").CopyTo(.Range("A" & intcurrow))
            rowsub = intcurrow
            .Range(rowsub, 3).Value = "Sub Total"
            For i As Integer = firstmonth To tMonth - 1
                For d As Integer = 1 To 5
                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1))
                    .Range(rowsub, (i * 25) + ((5 * d) - 1)).Value = "=SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")"

                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 1)
                    .Range(rowsub, (i * 25) + ((5 * d) - 1) + 1).Value = "=SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")"
                    ColAmtCal = strsubCol

                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 2)
                    .Range(rowsub, (i * 25) + (5 * d - 1) + 2).Value = "=SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")"

                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 3)
                    .Range(rowsub, (i * 25) + (5 * d - 1) + 3).Value = "=SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")"

                    ' strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 4)
                    .Range(rowsub, (i * 25) + ((5 * d) - 1 + 4)).Value = "=IF(ISERROR(SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")" & "/" & "SUM(" & ColAmtCal & rowCalStart & ":" & ColAmtCal & rowsub - 1 & ")),0,SUM(" & strsubCol & rowCalStart & ":" & strsubCol & rowsub - 1 & ")" & "/" & "SUM(" & ColAmtCal & rowCalStart & ":" & ColAmtCal & rowsub - 1 & "))"
                    '-----------------------------------------------------------------------
                Next
            Next

            rowGrandTotal = intcurrow + 1
            list1.Add(rowGrandTotal)
            For i As Integer = 1 To totalMat - 1
                intcurrow += 1
                rowsub = intcurrow
                rowMatEnd = intcurrow - 1
                Workbook.Worksheets(0).Range("A" & 18 + i & ":" & strCol & 18 + i).CopyTo(.Range("A" & intcurrow))
                '  .Range(rowsub, 3).Value = oRow("MATERIAL").ToString()
                'sheet.Range("A" & (dummy + i).ToString & ":" & strCol & (dummy + i).ToString).CopyTo(.Range("A" & intcurrow.ToString))
                For x As Integer = firstmonth To tMonth - 1
                    For d As Integer = 1 To 5

                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1))
                        .Range(rowsub, (x * 25) + ((5 * d) - 1)).Value = "=SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                            strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")"

                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 1)
                        .Range(rowsub, (x * 25) + ((5 * d) - 1) + 1).Value = "=SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                           strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")"
                        ColAmtCal = strsubCol
                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 2)
                        .Range(rowsub, (x * 25) + (5 * d - 1) + 2).Value = "=SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                           strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")"

                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 3)
                        .Range(rowsub, (x * 25) + (5 * d - 1) + 3).Value = "=SUMIF($C$" & rowMatStart & ":$C$" & rowMatEnd & ",$C$" & rowsub & "," & _
                           strsubCol & rowMatStart & ":" & strsubCol & rowMatEnd & ")"

                       
                        '-----------------------------------------------------------------------
                        .Range(rowsub, (x * 25) + (5 * d - 1) + 4).Value = "=IF(ISERROR(" & strsubCol & rowsub & "/" & ColAmtCal & rowsub & "),0," & strsubCol & rowsub & "/" & ColAmtCal & rowsub & ")"

                    Next
                Next
            Next

            intcurrow += 1
            Workbook.Worksheets(0).Range("A28:" & strCol & "28").CopyTo(.Range("A" & intcurrow))
            rowsub = intcurrow
            .Range(rowsub, 3).Value = "Sub Total"
            'sheet.Range("A" & (GrandDummy).ToString & ":" & strCol & (GrandDummy).ToString).CopyTo(.Range("A" & intcurrow.ToString))
            For i As Integer = firstmonth To tMonth - 1
                For d As Integer = 1 To 5
                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1))
                    .Range(rowsub, (i * 25) + ((5 * d) - 1)).Value = "=SUM(" & strsubCol & rowGrandTotal & ":" & strsubCol & rowsub - 1 & ")"

                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 1)
                    .Range(rowsub, (i * 25) + ((5 * d) - 1) + 1).Value = "=SUM(" & strsubCol & rowGrandTotal & ":" & strsubCol & rowsub - 1 & ")"
                    ColAmtCal = strsubCol
                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 2)
                    .Range(rowsub, (i * 25) + (5 * d - 1) + 2).Value = "=SUM(" & strsubCol & rowGrandTotal & ":" & strsubCol & rowsub - 1 & ")"

                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 3)
                    .Range(rowsub, (i * 25) + (5 * d - 1) + 3).Value = "=SUM(" & strsubCol & rowGrandTotal & ":" & strsubCol & rowsub - 1 & ")"

                    ' strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 4)
                    .Range(rowsub, (i * 25) + ((5 * d) - 1 + 4)).Value = "=IF(ISERROR(" & strsubCol & rowsub & "/" & ColAmtCal & rowsub & "),0," & strsubCol & rowsub & "/" & ColAmtCal & rowsub & ")"
                    '-----------------------------------------------------------------------
                Next
            Next
            Dim strRow As String
            Dim strRowCal As String
            For i As Integer = 1 To totalMat - 1
                strRow = ""
                intcurrow += 1
                rowsub = intcurrow
                rowMatEnd = intcurrow - 1
                Workbook.Worksheets(0).Range("A" & 39 + i & ":" & strCol & 39 + i).CopyTo(.Range("A" & intcurrow))
                For x As Integer = firstmonth To tMonth - 1
                    For d As Integer = 1 To 5

                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1))
                        strRow = ""
                        For Each l As Integer In list1
                            strRow = strRow & strsubCol & (l + i - 1).ToString & ","
                        Next
                        strRow = strRow.Remove(strRow.Length - 1, 1)

                        .Range(rowsub, (x * 25) + ((5 * d) - 1)).Value = "=SUM(" & strRow & ")"

                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 1)
                        strRow = ""
                        For Each l As Integer In list1
                            strRow = strRow & strsubCol & (l + i - 1).ToString & ","
                        Next
                        strRow = strRow.Remove(strRow.Length - 1, 1)
                        strRowCal = strRow
                        .Range(rowsub, (x * 25) + ((5 * d) - 1) + 1).Value = "=SUM(" & strRow & ")"

                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 2)
                        strRow = ""
                        For Each l As Integer In list1
                            strRow = strRow & strsubCol & (l + i - 1).ToString & ","
                        Next
                        strRow = strRow.Remove(strRow.Length - 1, 1)
                        .Range(rowsub, (x * 25) + (5 * d - 1) + 2).Value = "=SUM(" & strRow & ")"

                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1) + 3)
                        strRow = ""
                        For Each l As Integer In list1
                            strRow = strRow & strsubCol & (l + i - 1).ToString & ","
                        Next
                        strRow = strRow.Remove(strRow.Length - 1, 1)
                        .Range(rowsub, (x * 25) + (5 * d - 1) + 3).Value = "=SUM(" & strRow & ")"

                        .Range(rowsub, (x * 25) + ((5 * d) - 1 + 4)).Value = "=IF(ISERROR(SUM(" & strRow & ")/SUM(" & strRowCal & ")),0,SUM(" & strRow & ")/SUM(" & strRowCal & "))"

                    Next
                Next
            Next

            Dim lastrecord As Integer
            Dim startrecord As Integer
            lastrecord = intcurrow

            intcurrow += 1
            Workbook.Worksheets(0).Range("A14:" & strCol & "16").CopyTo(.Range("A" & intcurrow))
            startrecord = intcurrow - totalMat + 1
            For x As Integer = firstmonth To tMonth - 1
                For d As Integer = 1 To 5
                    '============Electric Wire=====================================================

                    strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1))
                    .Range(intcurrow, (x * 25) + ((5 * d) - 1)).Value = "=SUM(SUMIF($C$" & startrecord & ":$C$" & lastrecord & ",$D$1," _
                       & strsubCol & "$" & startrecord & ":$" _
                       & strsubCol & "$" & lastrecord _
                       & "),SUMIF($C$" & startrecord & ":$C$" & lastrecord & ",$D$2," _
                        & strsubCol & "$" & startrecord & ":$" _
                       & strsubCol & "$" & lastrecord & "))"

                    strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 1))
                    .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 1)).Value = "=SUM(SUMIF($C$" & startrecord & ":$C$" & lastrecord & ",$D$1," _
                       & strsubCol & "$" & startrecord & ":$" _
                       & strsubCol & "$" & lastrecord _
                       & "),SUMIF($C$" & startrecord & ":$C$" & lastrecord & ",$D$2," _
                        & strsubCol & "$" & startrecord & ":$" _
                       & strsubCol & "$" & lastrecord & "))"
                    ColAmtCal = strsubCol


                    strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 2))
                    .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 2)).Value = "=SUM(SUMIF($C$" & startrecord & ":$C$" & lastrecord & ",$D$1," _
                       & strsubCol & "$" & startrecord & ":$" _
                       & strsubCol & "$" & lastrecord _
                       & "),SUMIF($C$" & startrecord & ":$C$" & lastrecord & ",$D$2," _
                       & strsubCol & "$" & startrecord & ":$" _
                       & strsubCol & "$" & lastrecord & "))"

                    strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 3))
                    .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 3)).Value = "=SUM(SUMIF($C$" & startrecord & ":$C$" & lastrecord & ",$D$1," _
                       & strsubCol & "$" & startrecord & ":$" _
                       & strsubCol & "$" & lastrecord _
                       & "),SUMIF($C$" & startrecord & ":$C$" & lastrecord & ",$D$2," _
                       & strsubCol & "$" & startrecord & ":$" _
                       & strsubCol & "$" & lastrecord & "))"

                    
                    .Range(intcurrow, (x * 25) + ((5 * d) - 1 + 4)).Value = "=" & strsubCol & intcurrow & "/" & ColAmtCal & intcurrow
                Next
            Next
            intcurrow += 1
            intcurrow += 1
            For i As Integer = firstmonth To tMonth - 1
                For d As Integer = 1 To 5
                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1))
                    .Range(intcurrow, (i * 25) + ((5 * d) - 1)).Value = "=SUM(" & strsubCol & intcurrow - 2 & ":" & strsubCol & intcurrow - 1 & ")"

                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 1)
                    .Range(intcurrow, (i * 25) + ((5 * d) - 1) + 1).Value = "=SUM(" & strsubCol & intcurrow - 2 & ":" & strsubCol & intcurrow - 1 & ")"
                    ColAmtCal = strsubCol
                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 2)
                    .Range(intcurrow, (i * 25) + (5 * d - 1) + 2).Value = "=SUM(" & strsubCol & intcurrow - 2 & ":" & strsubCol & intcurrow - 1 & ")"

                    strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 3)
                    .Range(intcurrow, (i * 25) + (5 * d - 1) + 3).Value = "=SUM(" & strsubCol & intcurrow - 2 & ":" & strsubCol & intcurrow - 1 & ")"

                    ' strsubCol = Utility.GetExcelColumnName((i * 25) + ((5 * d) - 1) + 4)
                    .Range(intcurrow, (i * 25) + ((5 * d) - 1 + 4)).Value = "=IF(ISERROR(SUM(" & strsubCol & intcurrow - 2 & ":" & strsubCol & intcurrow - 1 & ")" & "/" & "SUM(" & ColAmtCal & intcurrow - 2 & ":" & ColAmtCal & intcurrow - 1 & ")),0,SUM(" & strsubCol & intcurrow - 2 & ":" & strsubCol & intcurrow - 1 & ")" & "/" & "SUM(" & ColAmtCal & intcurrow - 2 & ":" & ColAmtCal & intcurrow - 1 & "))"

                    '-----------------------------------------------------------------------
                Next
            Next

            Dim strRow2 As String
            Dim strRow3 As String
            Dim strRow4 As String

            Dim maxmonth As Integer
            If c_maxMonth < 0 Then
                maxmonth = 0
            Else
                maxmonth = c_maxMonth
            End If
            For r As Integer = 6 To intcurrow
               
                For d As Integer = 1 To 3
                    strRow = ""
                    strRow2 = ""
                    strRow3 = ""
                    strRow4 = ""

                    For x As Integer = firstmonth To maxmonth


                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1))
                        strRow = strRow & strsubCol & (r).ToString & ","

                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 1))
                        strRow2 = strRow2 & strsubCol & (r).ToString & ","
                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 2))
                        strRow3 = strRow3 & strsubCol & (r).ToString & ","
                        strsubCol = Utility.GetExcelColumnName((x * 25) + ((5 * d) - 1 + 3))
                        strRow4 = strRow4 & strsubCol & (r).ToString & ","
                    Next
                    strRow = strRow.Remove(strRow.Length - 1, 1)
                    strRow2 = strRow2.Remove(strRow2.Length - 1, 1)
                    strRow3 = strRow3.Remove(strRow3.Length - 1, 1)
                    strRow4 = strRow4.Remove(strRow4.Length - 1, 1)

                    If c_maxMonth < 0 Then
                        .Range(r, (tMonth * 25) + ((5 * d) - 1)).Value = 0
                        .Range(r, (tMonth * 25) + ((5 * d) - 1 + 1)).Value = 0
                        .Range(r, (tMonth * 25) + ((5 * d) - 1 + 2)).Value = 0
                        .Range(r, (tMonth * 25) + ((5 * d) - 1 + 3)).Value = 0
                        .Range(r, (tMonth * 25) + ((5 * d) - 1 + 4)).Value = 0
                    Else
                        .Range(r, (tMonth * 25) + ((5 * d) - 1)).Value = "=SUM(" & strRow & ")"
                        .Range(r, (tMonth * 25) + ((5 * d) - 1 + 1)).Value = "=SUM(" & strRow2 & ")"
                        .Range(r, (tMonth * 25) + ((5 * d) - 1 + 2)).Value = "=SUM(" & strRow3 & ")"
                        .Range(r, (tMonth * 25) + ((5 * d) - 1 + 3)).Value = "=SUM(" & strRow4 & ")"
                        .Range(r, (tMonth * 25) + ((5 * d) - 1 + 4)).Value = "=IF(ISERROR(SUM(" & strRow4 & ")/SUM(" & strRow2 & ")),0,SUM(" & strRow4 & ")/SUM(" & strRow2 & "))"

                    End If
                    'If sheetNo = 2 Then
                    '    ColAmtCal = Utility.GetExcelColumnName((5 * 25) + ((5 * d) - 1))
                    '    strsubCol = Utility.GetExcelColumnName((tMonth * 25) + ((5 * d) - 1))
                    '    .Range(r, (tMonth * 25) + ((5 * d) - 1 + 15)).Value = "=SUM(Data$!" & ColAmtCal & r & "," & strsubCol & r & ")"
                    '    '.Range(r, (tMonth * 25) + ((5 * d) - 1 + 1 + 15)).Value = 0
                    '    '.Range(r, (tMonth * 25) + ((5 * d) - 1 + 2 + 15)).Value = 0
                    '    '.Range(r, (tMonth * 25) + ((5 * d) - 1 + 3 + 15)).Value = 0
                    '    '.Range(r, (tMonth * 25) + ((5 * d) - 1 + 4 + 15)).Value = 0
                    'End If
                   
                Next
                
            Next

            
            'sheet.Range("A" & (GrandDummy).ToString & ":" & strCol & "" & (GrandDummy).ToString).CopyTo(.Range("A" & intcurrow.ToString))

            'For i As Integer = 1 To totalMat - 1
            '    intcurrow += 1
            '    sheet.Range("A" & (dummy2 + i).ToString & ":" & strCol & "" & (dummy2 + i).ToString).CopyTo(.Range("A" & intcurrow.ToString))
            'Next
            'intcurrow += 1
            'sheet.Range("A" & (Summary).ToString & ":" & strCol & "" & (Summary + 2).ToString).CopyTo(.Range("A" & intcurrow.ToString))


        End With
        Call SetPageProperties()
        If sheetNo = 1 Then
            sheet2 = sheetPrint
        Else
            sheet3 = sheetPrint
        End If
        'Workbook.Worksheets("Data").Range("A2:AZ65000").CopyTo(Workbook.Worksheets(0).Range("A2:AZ65000"))


        Return True
    End Function




    Protected Function Percent(Amount As Decimal, DivideBy As Decimal) As Decimal
        If DivideBy = 0 Then
            Percent = 0
        Else
            Percent = Amount / DivideBy
        End If
    End Function

    Protected Function ConvertDec(str As Object) As Decimal
        If str.ToString = "" Then
            Return 0
        End If
        Return CDec(str)
    End Function
    Protected Sub SetPageProperties()

        With sheetPrint
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

