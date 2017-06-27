Imports DataAccess
Imports Syncfusion.XlsIO
Imports Syncfusion.XlsIO.ExcelEngine
Imports System.IO
Imports System.Globalization
Imports System.Drawing
Public Class PrintReport4
    Inherits System.Web.UI.Page
    Private CurTempPath As String = Server.MapPath("~/tmp/")
    Private DefaultPath As String = ("~/tmp/")
    Private CurReportPath As String = Server.MapPath("~/report/ExcelReport/")
    Private TemplateFile As String = CurReportPath & ("tpReport4.xlt")
    Private ReportName As String = "REPORT 4"
    Private crtDate As String = Date.Now.ToString("ddMMyyyy")
    Private crtTime As String = Date.Now.ToString("HHmm")
    Private excelengine As New ExcelEngine
    Private appExcel As IApplication = excelengine.Excel
    Private Workbook As IWorkbook
    Private FirstRecordSheet2 = True
    Private sheet As IWorksheet
    Private sheet2 As IWorksheet
    Private sheet3 As IWorksheet
    Private sheetPrint As IWorksheet
    Private sheetPrint2 As IWorksheet
    Private intcurrow, intno, intstart As Integer
    Dim Cloneintcurrow As Integer = 0
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
        If Not IsPostBack Then
            checkClass()
            CallQuery()
        End If

    End Sub
    Sub checkClass()
        Dim cmd1 As String
        Dim dt1 As New DataTable
        cmd1 = "Select AGITCL as Class from {0}.TACLMA full join {0}.TASPAG on CLCLCD = AGITCL where CLCLCD Is null group by AGITCL"
        cmd1 = String.Format(cmd1, strLib)
        dt1 = DbConnect.ExcuteQueryString(cmd1, DBConnection.DatabaseType.AS400)
        GridView1.DataSource = dt1
        GridView1.DataBind()
        If dt1.Rows.Count > 0 Then
            Label1.Visible = True
            Label2.Visible = True
        End If
    End Sub
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
    Dim EMonth As Integer
    Public Function CallQuery() As Boolean
        If txtDateS.Text.ToString = "" Then
            txtDateS.Focus()
            AlertMessagebox("Please choose Start Date!")
            Return False
        End If

        If txtDateE.Text.ToString = "" Then
            AlertMessagebox("Please choose End Date!")
            txtDateE.Focus()
            Return False
        End If

        Dim Sdat As Integer = Convert.ToInt32(txtDateS.Text.Trim.Substring(0, 6))
        Dim Edat As Integer = Convert.ToInt32(txtDateE.Text.Trim.Substring(0, 6))
        Dim Over As Integer = Edat - Sdat
        If Edat < Sdat Then
            AlertMessagebox("Your End Date is less then Start Date !!")
            txtDateE.Text = ""
            txtDateE.Focus()
            Return False
        End If

        If Over > 99 Then
            AlertMessagebox("Please select up to 12 months!!")
            txtDateE.Text = ""
            txtDateE.Focus()
            Return False
        End If


        Dim cmd, cmd2 As String
        Dim dt, dt2 As New DataTable
        Dim str, str2 As String
        Using rw As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qryReport4.lbs")
            str = rw.ReadToEnd
        End Using

        Using rw2 As StreamReader = New StreamReader(Server.MapPath(".") & "/query/qryReport4_CheckMonthTotal.lbs")
            str2 = rw2.ReadToEnd
        End Using

        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim tMonth As Integer
        Dim dateS As Date
        Dim dateE As Date
        dateS = Date.ParseExact(txtDateS.Text.ToString.Trim, "yyyyMMdd", provider)
        dateE = Date.ParseExact(txtDateE.Text.ToString.Trim, "yyyyMMdd", provider)
        tMonth = DateDiff(DateInterval.Month, dateS, dateE) + 1

        If tMonth > 6 Then
            cmd = String.Format(str, dateS.ToString("yyyyMM"), dateS.AddMonths(5).ToString("yyyyMM"), strLib)
            dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.SQL)
            If dt.Rows.Count > 0 Then
                EMonth = dateS.AddMonths(5).ToString("MM")
                cmd2 = String.Format(str2, dateS.ToString("yyyyMM"), dateS.AddMonths(5).ToString("yyyyMM"), strLib)
                dt2 = DbConnect.ExcuteQueryString(cmd2, DBConnection.DatabaseType.SQL)
                ExportExcel(dt, 6, 1, dt2)
                tMonth = tMonth - 6
            End If


            cmd = String.Format(str, dateS.AddMonths(6).ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
            If dt.Rows.Count > 0 Then
                dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.SQL)
                cmd2 = String.Format(str2, dateS.AddMonths(6).ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
                dt2 = DbConnect.ExcuteQueryString(cmd2, DBConnection.DatabaseType.SQL)

                EMonth = dateE.ToString("MM")
                ExportExcel(dt, tMonth, 2, dt2)
            End If
        Else

            cmd = String.Format(str, dateS.ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
            dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.SQL)
            If dt.Rows.Count > 0 Then
                cmd2 = String.Format(str2, dateS.ToString("yyyyMM"), dateE.ToString("yyyyMM"), strLib)
                dt2 = DbConnect.ExcuteQueryString(cmd2, DBConnection.DatabaseType.SQL)
                EMonth = dateE.ToString("MM")
                ExportExcel(dt, tMonth, 1, dt2)
            End If
        End If


        If dt.Rows.Count > 0 Then
            filename = ReportName & "-" & crtDate & "-" & crtTime & ".xls"
            Workbook.Worksheets(0).Remove()
            Workbook.SaveAs(filename, Response, ExcelHttpContentType.Excel2000, ExcelDownloadType.Open)
        End If
        Return True
    End Function
    Protected Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click

        CallQuery()


    End Sub
    Public Function getTotalMonth(startDate As String, enddate As String) As Integer
        Dim yearE As Integer
        Dim yearS As Integer
        Dim monthE As Integer
        Dim monthS As Integer
        Dim tMonth As Integer
        yearE = CInt(enddate.Substring(0, 4))
        yearS = CInt(startDate.Substring(0, 4))
        monthE = CInt(enddate.Substring(4, 2))
        monthS = CInt(startDate.Substring(4, 2))
        If yearE - yearS >= 1 Then
            tMonth = (13 - monthS) + monthE + 1
        Else
            tMonth = monthE - monthS + 1
        End If
        Return tMonth
    End Function


    Public Function SetColumnMonthName(dateS As Date, totalMonth As Integer)

        With sheetPrint
            For i As Integer = 0 To totalMonth - 1
                .Range(3, (i * 35) + 20).Value = "'" & dateS.ToString("MMMM", CultureInfo.InvariantCulture) & " " & dateS.Year.ToString
                dateS = dateS.AddMonths(1)
            Next

        End With

    End Function

    Public Function CalculateTotalRow(Rintcurrow As Integer, Startintcurrow As Integer, Endintcurrow As Integer, rangColumn As Integer)
        Dim Counts As Integer = 0
        With sheetPrint
            .Range("D" & Rintcurrow).Formula = "SUM(D" & Startintcurrow & ":D" & Endintcurrow & ")"
            Counts = 2
            rangColumn = rangColumn + 35
            For j As Integer = 5 To rangColumn Step 1
                Dim Colname As String = GetExcelColumnName(j)
                If Counts < 5 Then
                    Counts = Counts + 1
                    .Range("D" & Rintcurrow).CopyTo(.Range(Colname & Rintcurrow))
                Else
                    Dim ProColumn As String = GetExcelColumnName(j - 1)
                    Dim AmoColumn As String = GetExcelColumnName(j - 3)
                    .Range(Colname & Rintcurrow).Formula = "=IF(ISERROR(" & ProColumn & Rintcurrow & "/" & AmoColumn & Rintcurrow & "),0," & ProColumn & Rintcurrow & "/" & AmoColumn & Rintcurrow & ")"
                    Counts = 1
                End If
            Next
        End With
        Return True
    End Function


    Public Function ExportExcel(oTable As DataTable, tMonth As Integer, sheetNo As Integer, YMTable As DataTable) As Boolean
        Dim oRow As DataRow
        Dim oRow2 As DataRow
        Dim group As String = ""
        Dim ordersub As String = ""
        Dim classT As String = ""
        Dim typedoc As String = ""
        Dim cmonth As Integer
        Dim cmonths As Integer
        Dim datatypeid As Integer
        Dim DiffPer As Decimal
        Dim isFisrtRecord As Boolean = True
        Dim isFisrtRecordC As Boolean = True
        Dim TypeTTP As String = ""
        Dim strCol, strCols As String
        Dim extMonth As Integer
        Dim dateS As Date
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim isFisrtSheet2 As Boolean = True
        Dim isFisrtcmoth As Boolean = True
        Dim AQuality, AAmount, ACost, AProfit, BQuality, BAmount, BCost, BProfit As Decimal
        oRow = oTable.Rows(0)
        oRow2 = YMTable.Rows(0)
        intcurrow = 6
        Dim Startintcurrow As Integer
        Dim TotalALFabicRowS, TotalALFabicRowE As Integer
        Dim TotalALFinishRowS, TotalALFinishRowE As Integer
        Dim TotalCUFabicRowS, TotalCUFabicRowE As Integer
        Dim TotalCUFinishRowS, TotalCUFinishRowE As Integer
        Dim TotalCUWireRowS, TotalCUWireRowE As Integer
        Dim TotalVTARowS, TotalVTARowE As Integer
        Dim TotalOtherRowS, TotalOtherRowE As Integer
        Dim Counts As Integer = 0
        Dim rangColumnO As Integer


        If oTable.Rows.Count <= 0 Then
            Return False
        End If
        Startintcurrow = intcurrow + 1
        TotalALFabicRowS = intcurrow + 1
        '---------------- Set Sheet Name ---------------------------
        appExcel.DefaultFilePath = Server.MapPath(".")
        strCol = Utility.GetExcelColumnName((tMonth * 35) + 3 + 35)
        If sheetNo = 1 Then
            Workbook = appExcel.Workbooks.Open(TemplateFile, ExcelOpenType.Automatic)
            sheet = Workbook.Worksheets(0)
            sheetPrint = Workbook.Worksheets.Create("1 to 6")
        Else
            sheet = Workbook.Worksheets(0)
            sheetPrint = Workbook.Worksheets.Create("6 to 12")
        End If
        '---------------------===============------------------------

        With sheetPrint

            strCols = Utility.GetExcelColumnName((tMonth * 35) + 3)
            '---------------- Set Header Date ---------------------------
            Workbook.Worksheets(0).Range("A1:" & strCols & "6").CopyTo(.Range("A1"))

            dateS = Date.ParseExact(txtDateS.Text.ToString.Trim, "yyyyMMdd", provider)
            If sheetNo = 2 Then
                dateS = dateS.AddMonths(6)
            End If
            SetColumnMonthName(dateS, tMonth)
            '---------------------===============------------------------


            '==================== Loop Put Data =============================
            Dim start As Integer = 60
            For Each oRow In oTable.Rows

                Dim RowStand As Integer = 0
                Dim LoopRow As Integer = 0
                Dim PointTotal As Integer
                Dim rangColumnN As Integer

                '-----------Set Tap Total------------
                If isFisrtRecordC = True Then
                    TypeTTP = oRow("CheckType").ToString()
                    isFisrtRecordC = False
                End If


                If oRow("CheckType").ToString() <> TypeTTP Then
                    If oRow("CheckType").ToString <> "" Then
                        If TypeTTP.Trim.Substring(0, 2) = "AL" Then
                            If TypeTTP.Trim.Substring(3) = "Fabrication" Then

                                Workbook.Worksheets(0).Range("A25:" & strCol & "25").CopyTo(.Range("A" & intcurrow + 1))
                                CalculateTotalRow(intcurrow + 1, Startintcurrow, intcurrow, rangColumnO)

                                If intcurrow <= TotalALFabicRowS Then
                                    TotalALFabicRowS = intcurrow
                                End If
                                If intcurrow > TotalALFabicRowS Then
                                    TotalALFabicRowE = intcurrow
                                End If

                                intcurrow = intcurrow + 1
                                Startintcurrow = intcurrow + 1
                            Else

                                Workbook.Worksheets(0).Range("A31:" & strCol & "31").CopyTo(.Range("A" & intcurrow + 1))
                                Workbook.Worksheets(0).Range("A27:" & strCol & "27").CopyTo(.Range("A" & intcurrow + 2))
                                Workbook.Worksheets(0).Range("A33:" & strCol & "33").CopyTo(.Range("A" & intcurrow + 3))
                                CalculateTotalRow(intcurrow + 1, Startintcurrow, intcurrow, rangColumnO)
                                CalculateTotalRow(intcurrow + 2, TotalALFabicRowS, TotalALFabicRowE, rangColumnO)
                                CalculateTotalRow(intcurrow + 3, Startintcurrow, intcurrow, rangColumnO)

                                If TotalALFinishRowS = 0 Then
                                    TotalALFinishRowS = Startintcurrow
                                End If

                                If intcurrow <= TotalALFinishRowS Then
                                    TotalALFinishRowS = intcurrow
                                End If
                                If intcurrow > TotalALFinishRowS Then
                                    TotalALFinishRowE = intcurrow
                                End If
                                intcurrow = intcurrow + 3
                                Startintcurrow = intcurrow + 1
                            End If


                        ElseIf TypeTTP.Trim.Substring(0, 2) <> oRow("CheckType").Trim.Substring(0, 2) And TypeTTP.Trim.Substring(0, 2) = "CU" Then
                            Workbook.Worksheets(0).Range("A37:" & strCol & "37").CopyTo(.Range("A" & intcurrow + 1))
                            Workbook.Worksheets(0).Range("A41:" & strCol & "41").CopyTo(.Range("A" & intcurrow + 2))

                            TotalCUFabicRowS = Startintcurrow
                            TotalCUFabicRowE = TotalCUFabicRowE + TotalCUFabicRowS

                            TotalCUFinishRowS = TotalCUFabicRowE + 1
                            TotalCUFinishRowE = TotalCUFinishRowE + TotalCUFinishRowS

                            TotalCUWireRowS = TotalCUFinishRowE + 1
                            TotalCUWireRowE = TotalCUWireRowE + TotalCUWireRowS

                            CalculateTotalRow(intcurrow + 1, Startintcurrow, TotalCUFabicRowE, rangColumnO)
                            CalculateTotalRow(intcurrow + 2, TotalCUFinishRowS, TotalCUFinishRowE, rangColumnO)

                            intcurrow = intcurrow + 2
                            Startintcurrow = intcurrow + 1

                        ElseIf TypeTTP.Trim.ToUpper <> oRow("CheckType").Trim.ToString.ToUpper And TypeTTP = "VTA" Then
                            Workbook.Worksheets(0).Range("A45:" & strCol & "45").CopyTo(.Range("A" & intcurrow + 1))

                            TotalVTARowS = Startintcurrow
                            TotalVTARowE = TotalVTARowE + TotalVTARowS

                            CalculateTotalRow(intcurrow + 1, TotalVTARowS, TotalVTARowE, rangColumnO)
                            intcurrow = intcurrow + 1
                            Startintcurrow = intcurrow + 1
                        End If
                    Else
                        Workbook.Worksheets(0).Range("A45:" & strCol & "45").CopyTo(.Range("A" & intcurrow + 1))

                        TotalVTARowS = Startintcurrow
                        TotalVTARowE = TotalVTARowE + TotalVTARowS

                        CalculateTotalRow(intcurrow + 1, TotalVTARowS, TotalVTARowE, rangColumnO)
                        intcurrow = intcurrow + 1
                        Startintcurrow = intcurrow + 1
                    End If
                    TotalOtherRowS = Startintcurrow
                End If
                '-----------===============------------

                If oRow("CheckType").ToString.Trim.ToUpper <> "CASH DISCOUNT" Then

                    '-----------Set Header Total------------
                    If classT <> oRow("Class").ToString().Trim Then
                        If TypeTTP.ToUpper.Trim = oRow("CheckType").ToString().Trim.ToUpper Then
                            If TypeTTP.ToUpper.Trim = "CU FABRICATION" Then
                                TotalCUFabicRowE += 1
                            End If
                            If TypeTTP.ToUpper.Trim = "CU FINISHED" Then
                                TotalCUFinishRowE += 1
                            End If
                            If TypeTTP.ToUpper.Trim = "CU WIRE ROD" Then
                                TotalCUWireRowE += 1
                            End If
                            If TypeTTP.ToUpper.Trim = "VTA" Then
                                TotalVTARowE += 1
                            End If
                            If TypeTTP.ToUpper.Trim = "" Or TypeTTP.ToUpper.Trim = "Other" Then
                                TotalOtherRowE += 1
                            End If
                        End If
                        intcurrow = intcurrow + 1
                        Workbook.Worksheets(0).Range("A8:" & strCol & "8").CopyTo(.Range("A" & intcurrow))
                        Workbook.Worksheets(0).Range("D8:AL8").CopyTo(Workbook.Worksheets(0).Range("D" & 60 + intcurrow))
                        classT = oRow("Class").ToString().Trim
                    End If
                    '-----------===============------------
                    '-----------Set Header Total------------
                    If ordersub <> oRow("SubType").ToString().Trim Then
                        Workbook.Worksheets(0).Range("A8:" & strCol & "8").CopyTo(.Range("A" & intcurrow))
                        ordersub = oRow("SubType").ToString().Trim
                    End If
                    '-----------===============------------
                    '-----------Set Column A Merch Type Total------------
                    If group <> oRow("MainType").ToString.Trim Then
                        If isFisrtRecord Then
                            Workbook.Worksheets(0).Range("A23:" & strCol & "23").CopyTo(.Range("A" & intcurrow))
                            Workbook.Worksheets(0).Range("D23:AL23").CopyTo(Workbook.Worksheets(0).Range("D" & 60 + intcurrow))
                            .Range(intcurrow, 1).Text = oRow("MainType").ToString().Trim
                        End If
                        group = oRow("MainType").ToString.Trim
                    End If
                    '-----------============================------------

                    .Range(intcurrow, 2).Text = "'" & oRow("SubType").ToString()
                    .Range(intcurrow, 3).Text = "'" & oRow("Class").ToString().Trim()
                End If

                Dim cmonthN As Integer = CInt(oRow("MONTH").ToString())
                datatypeid = CInt(oRow("TypeLoop").ToString())
                If datatypeid > 2 Then
                    datatypeid = datatypeid + 1
                End If

                TypeTTP = oRow("CheckType").ToString()

                If cmonths <> cmonthN Then
                    If isFisrtcmoth = True Then
                        cmonth = 0
                        isFisrtcmoth = False
                    Else
                        ' Dim a As String = cmonth
                        If cmonth > 4 Then
                            cmonth = 0
                        ElseIf cmonth = (tMonth - 1) Then
                            cmonth = 0
                        Else
                            cmonth = cmonth + 1
                        End If
                    End If
                    cmonths = cmonthN
                End If

                If cmonth = 0 Then
                    If isFisrtSheet2 Then
                        sheet.Range("D25:HE29").Value = 0
                        sheet.Range("D31:HE35").Value = 0
                        sheet.Range("D37:HE39").Value = 0
                        sheet.Range("D41:HE43").Value = 0
                        sheet.Range("D45:HE47").Value = 0
                        sheet.Range("D49:HE51").Value = 0
                        sheet.Range("D53:HE53").Value = 0
                        sheet.Range("D55:HE55").Value = 0
                        isFisrtSheet2 = False
                    End If
                End If

                Dim StatusSum As Boolean = False
                Dim Mn As String
                Dim CheckM, CheckM2 As Integer
                Mn = oRow("Month").ToString.Length

                If Mn < 2 Then
                    CheckM = Convert.ToString(oRow("Year").ToString.Trim + "0" + oRow("Month").ToString)
                Else
                    CheckM = Convert.ToString(oRow("Year").ToString.Trim + oRow("Month").ToString)
                End If

                For Each oRow2 In YMTable.Rows
                    If oRow2("YM").ToString.Trim = CheckM Then
                        StatusSum = True
                        Exit For
                    End If
                Next

                CheckM2 = oRow2("YM").ToString()


                '_____________________________ PUT NORMAL DATA_____________________________________
                Dim QTB, ATB, CTB, PTB, QTU, ATU, CTU, PTU, PTT, ATT, Q, A, C, P, Q1, A1, C1, P1 As String
                If oRow("CheckType").ToString.Trim.ToUpper <> "CASH DISCOUNT" Then
                    .Range(intcurrow, (cmonth * 35) + ((5 * datatypeid) - 1)).Value = oRow("Quality").ToString()
                    .Range(intcurrow, (cmonth * 35) + ((5 * datatypeid) - 1) + 1).Value = oRow("Amount").ToString()
                    .Range(intcurrow, (cmonth * 35) + (5 * datatypeid - 1) + 2).Value = oRow("Cost").ToString()
                    .Range(intcurrow, (cmonth * 35) + (5 * datatypeid - 1) + 3).Value = oRow("Profit").ToString()
                    PTT = GetExcelColumnName((cmonth * 35) + (5 * datatypeid - 1) + 3)
                    ATT = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 1)
                    .Range(intcurrow, (cmonth * 35) + (5 * datatypeid - 1) + 4).Formula = "=IF(ISERROR(" & PTT & intcurrow & "/" & ATT & intcurrow & "),0," & PTT & intcurrow & "/" & ATT & intcurrow & ")"

                    Q = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1)) & intcurrow
                    A = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 1) & intcurrow
                    C = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 2) & intcurrow
                    P = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 3) & intcurrow

                    If StatusSum = True Then
                        sheet.Range(60 + intcurrow, ((5 * datatypeid) - 1)).Value = sheet.Range(60 + intcurrow, ((5 * datatypeid) - 1)).Value + "+" + Q
                        sheet.Range(60 + intcurrow, ((5 * datatypeid) - 1) + 1).Value = sheet.Range(60 + intcurrow, ((5 * datatypeid) - 1) + 1).Value + "+" + A
                        sheet.Range(60 + intcurrow, (5 * datatypeid - 1) + 2).Value = sheet.Range(60 + intcurrow, (5 * datatypeid - 1) + 2).Value + "+" + C
                        sheet.Range(60 + intcurrow, (5 * datatypeid - 1) + 3).Value = sheet.Range(60 + intcurrow, (5 * datatypeid - 1) + 3).Value + "+" + P


                        'If CheckM = CheckM2 Then
                        '    Q1 = sheet.Range(60 + intcurrow, ((5 * datatypeid) - 1)).Value
                        '    A1 = sheet.Range(60 + intcurrow, ((5 * datatypeid) - 1) + 1).Value
                        '    C1 = sheet.Range(60 + intcurrow, (5 * datatypeid - 1) + 2).Value
                        '    P1 = sheet.Range(60 + intcurrow, (5 * datatypeid - 1) + 3).Value

                        '    sheet.Range(60 + intcurrow, ((5 * datatypeid) - 1)).Formula = Q1
                        '    sheet.Range(60 + intcurrow, ((5 * datatypeid) - 1) + 1).Formula = A1
                        '    sheet.Range(60 + intcurrow, (5 * datatypeid - 1) + 2).Formula = C1
                        '    sheet.Range(60 + intcurrow, (5 * datatypeid - 1) + 3).Formula = P1

                        'End If
                    End If

                Else
                    sheet.Range(19, (cmonth * 35) + ((5 * datatypeid) - 1)).Value = oRow("Quality").ToString()
                    sheet.Range(19, (cmonth * 35) + ((5 * datatypeid) - 1) + 1).Value = oRow("Amount").ToString()
                    sheet.Range(19, (cmonth * 35) + (5 * datatypeid - 1) + 2).Value = oRow("Cost").ToString()
                    sheet.Range(19, (cmonth * 35) + (5 * datatypeid - 1) + 3).Value = oRow("Profit").ToString()
                    sheet.Range(19, (cmonth * 35) + (5 * datatypeid - 1) + 4).Value = Percent(ConvertDec(oRow("Profit")), ConvertDec(oRow("Amount")))
                    If StatusSum = True Then
                        sheet.Range(63, ((5 * datatypeid) - 1)).Value = ConvertDec(sheet.Range(63, ((5 * datatypeid) - 1)).Value) + ConvertDec(oRow("Quality").ToString())
                        sheet.Range(63, ((5 * datatypeid) - 1) + 1).Value = ConvertDec(sheet.Range(63, ((5 * datatypeid) - 1) + 1).Value) + ConvertDec(oRow("Amount").ToString())
                        sheet.Range(63, (5 * datatypeid - 1) + 2).Value = ConvertDec(sheet.Range(63, (5 * datatypeid - 1) + 2).Value) + ConvertDec(oRow("Cost").ToString())
                        sheet.Range(63, (5 * datatypeid - 1) + 3).Value = ConvertDec(sheet.Range(63, (5 * datatypeid - 1) + 3).Value) + ConvertDec(oRow("Profit").ToString())
                    End If
                End If


                '_____________________________ CALCULATE AND PUT TOTAL (BKK + UPCOUNTRY) _____________________________________

                If datatypeid < 3 Then
                    PointTotal = 3
                Else
                    PointTotal = 6
                End If

                If oRow("TYPELOOP").ToString() = "1" Or oRow("TYPELOOP").ToString() = "4" Then
                    QTB = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1))
                    ATB = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 1)
                    CTB = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 2)
                    PTB = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 3)
                Else
                    QTU = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1))
                    ATU = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 1)
                    CTU = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 2)
                    PTU = GetExcelColumnName((cmonth * 35) + ((5 * datatypeid) - 1) + 3)
                End If

                Dim QDB, ADB, CDB, PDB, QDU, ADU, CDU, PDU As String
                If oRow("CheckType").ToString.Trim.ToUpper <> "CASH DISCOUNT" Then
                    .Range(intcurrow, (cmonth * 35) + ((5 * PointTotal) - 1)).Formula = "=SUM(" & QTB & intcurrow & "+" & QTU & intcurrow & ")"
                    .Range(intcurrow, (cmonth * 35) + ((5 * PointTotal) - 1) + 1).Formula = "=SUM(" & ATB & intcurrow & "+" & ATU & intcurrow & ")"
                    .Range(intcurrow, (cmonth * 35) + (5 * PointTotal - 1) + 2).Formula = "=SUM(" & CTB & intcurrow & "+" & CTU & intcurrow & ")"
                    .Range(intcurrow, (cmonth * 35) + (5 * PointTotal - 1) + 3).Formula = "=SUM(" & PTB & intcurrow & "+" & PTU & intcurrow & ")"
                    PTT = GetExcelColumnName((cmonth * 35) + (5 * PointTotal - 1) + 3)
                    ATT = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1) + 1)
                    .Range(intcurrow, (cmonth * 35) + (5 * PointTotal - 1) + 4).Formula = "=IF(ISERROR(" & PTT & intcurrow & "/" & ATT & intcurrow & "),0," & PTT & intcurrow & "/" & ATT & intcurrow & ")"
                End If

                '_____________________________ CALCULATE AND PUT DiFF (B - A) _____________________________________

                If PointTotal = 3 Then
                    QDB = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1))
                    ADB = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1) + 1)
                    CDB = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1) + 2)
                    PDB = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1) + 3)
                Else
                    QDU = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1))
                    ADU = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1) + 1)
                    CDU = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1) + 2)
                    PDU = GetExcelColumnName((cmonth * 35) + ((5 * PointTotal) - 1) + 3)
                End If


                If datatypeid = 5 Then
                    If oRow("CheckType").ToString.Trim.ToUpper <> "CASH DISCOUNT" Then
                        .Range(intcurrow, (cmonth * 35) + ((5 * 7) - 1)).Formula = "=" & QDU & intcurrow & "-" & QDB & intcurrow & ")"
                        .Range(intcurrow, (cmonth * 35) + ((5 * 7) - 1) + 1).Value = "=" & ADU & intcurrow & "-" & ADB & intcurrow & ")"
                        .Range(intcurrow, (cmonth * 35) + (5 * 7 - 1) + 2).Value = "=" & CDU & intcurrow & "-" & CDB & intcurrow & ")"
                        .Range(intcurrow, (cmonth * 35) + (5 * 7 - 1) + 3).Value = "=" & PDU & intcurrow & "-" & PDB & intcurrow & ")"
                        PTT = GetExcelColumnName((cmonth * 35) + (5 * 7 - 1) + 3)
                        ATT = GetExcelColumnName((cmonth * 35) + ((5 * 7) - 1) + 1)
                        .Range(intcurrow, (cmonth * 35) + (5 * 7 - 1) + 4).Formula = "=IF(ISERROR(" & PTT & intcurrow & "/" & ATT & intcurrow & "),0," & PTT & intcurrow & "/" & ATT & intcurrow & ")"
                        If oRow("Month").ToString.Trim = EMonth.ToString Then
                            rangColumnN = (cmonth * 35) + (5 * 7 - 1) + 4
                            If rangColumnN > rangColumnO Then
                                rangColumnO = rangColumnN
                            End If
                            Workbook.Worksheets(0).Range("D" & 60 + intcurrow & ":AL" & 60 + intcurrow).CopyTo(.Range(GetExcelColumnName(rangColumnO + 1) & intcurrow))
                            For i As Integer = 1 To 35 Step 1
                                If i < 25 Then
                                    If i > 15 Then
                                        If i <> 20 Then
                                            Dim Formu As String = .Range(GetExcelColumnName(rangColumnO + i) & intcurrow).Value
                                            .Range(GetExcelColumnName(rangColumnO + i) & intcurrow).Formula = "=" & Formu
                                        End If
                                    End If
                                    If i < 10 Then
                                        If i <> 5 Then
                                            Dim Formu As String = .Range(GetExcelColumnName(rangColumnO + i) & intcurrow).Value
                                            .Range(GetExcelColumnName(rangColumnO + i) & intcurrow).Formula = "=" & Formu
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            Next

            strCols = GetExcelColumnName(rangColumnO + 1)
            Workbook.Worksheets(0).Range("HF1:IN6").CopyTo(.Range(strCols & "1"))
            Dim SMon, DMon As String
            Dim MON As Date
            Dim FirstR As Integer
            FirstR = 1
            For i As Integer = 0 To YMTable.Rows.Count - 1 Step 1
                SMon = YMTable.Rows(i)("YM").ToString() + "01"
                MON = Date.ParseExact(SMon, "yyyyMMdd", provider)
                If FirstR = 1 Then
                    If YMTable.Rows.Count = 1 Then
                        DMon = "(" & MON.ToString("MMMM", CultureInfo.InvariantCulture) & " " & MON.ToString("yyy", CultureInfo.InvariantCulture) & ")"
                    Else
                        DMon = "(" & MON.ToString("MMMM", CultureInfo.InvariantCulture) & " " & MON.ToString("yyy", CultureInfo.InvariantCulture)
                    End If
                    FirstR = 0
                ElseIf i = YMTable.Rows.Count - 1 Then
                    DMon = DMon & " - " & MON.ToString("MMMM", CultureInfo.InvariantCulture) & " " & MON.ToString("yyy", CultureInfo.InvariantCulture) & ")"
                End If
            Next



            .Range(strCols & "3").Value = "TOTAL " & DMon


            If TypeTTP.Trim = "Other" Or TypeTTP.Trim = "" Then
                Workbook.Worksheets(0).Range("A49:" & strCol & "49").CopyTo(.Range("A" & intcurrow + 1))
                TotalOtherRowS = Startintcurrow
                TotalOtherRowE = TotalOtherRowE + TotalOtherRowS
                CalculateTotalRow(intcurrow + 1, TotalOtherRowS, TotalOtherRowE, rangColumnO)
                intcurrow = intcurrow + 1
            End If
            '----------- Set / Place Grand total anf all total when last Row  ------------
            Workbook.Worksheets(0).Range("A10:" & strCol & "20").CopyTo(.Range("A" & intcurrow + 1))
            CalculateTotalRow(intcurrow + 1, TotalCUFinishRowS, TotalCUFinishRowE, rangColumnO)
            CalculateTotalRow(intcurrow + 2, TotalCUWireRowS, TotalCUWireRowE, rangColumnO)
            CalculateTotalRow(intcurrow + 5, TotalCUFabicRowS, TotalCUFabicRowE, rangColumnO)
            CalculateTotalRow(intcurrow + 6, TotalALFinishRowS, TotalALFinishRowE, rangColumnO)
            CalculateTotalRow(intcurrow + 7, TotalALFabicRowS, TotalALFabicRowE, rangColumnO)
            CalculateTotalRow(intcurrow + 8, TotalVTARowS, TotalVTARowE, rangColumnO)
            CalculateTotalRow(intcurrow + 9, TotalOtherRowS, TotalOtherRowE, rangColumnO)
            Workbook.Worksheets(0).Range("D63:AL63").CopyTo(.Range(strCols & intcurrow + 10))


            '----------- ==================================================== ---------
            .Range.AutofitColumns()
            Dim lastcolumn As Integer
            lastcolumn = rangColumnO + 34
            For i As Integer = 3 To lastcolumn
                .Columns(i).ColumnWidth = 16.82
            Next


            SetPageProperties()
        End With

        Return True
    End Function

    Protected Function Percent(Profit As Decimal, Amount As Decimal) As Decimal
        If Profit = 0 Then
            Percent = 0
        Else
            Percent = Profit / Amount
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
    Private Function GetExcelColumnName(columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = [String].Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function
End Class
