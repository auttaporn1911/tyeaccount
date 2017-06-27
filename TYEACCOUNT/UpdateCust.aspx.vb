Imports DataAccess
Imports System.IO
Public Class WebForm14
    Inherits System.Web.UI.Page

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
    Public Function GetInteger(str As Object) As String
        If str.ToString = "" Then
            str = "0"
        End If
        Return str
    End Function
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            If Request("state") = "add" Then
                txtCustCode.Enabled = True
            Else
                txtCustCode.Enabled = False
                txtCustCode.Text = Request("cust").ToString
                Binddata()
            End If


        End If
    End Sub
    Public Sub Binddata()
        Dim cmd As String
        cmd = "select * from " & strLib & ".TACSMA where  CSCSCD = '" & txtCustCode.Text & "' " & _
            "and CSEFFD in (select max(CSEFFD) from " & strLib & ".TACSMA where CSCSCD = '" & txtCustCode.Text & "')"

        Dim dt As New DataTable
        dt = DbConnect.ExcuteQueryString(cmd, DBConnection.DatabaseType.AS400)
        If dt.Rows.Count > 0 Then
            Dim r As DataRow
            r = dt.Rows(0)
            txtAccCode.Text = r("CSACCD")
            txtCustom.Text = r("CSCUTM")
            txtEffDate.Text = r("CSEFFD")
            txtGroup.Text = r("CSCSGP")
            txtName.Text = r("CSCSNM")
            txtSection.Text = r("CSSECT")
            txtType.Text = r("CSTYPE")
        End If
    End Sub

    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim str As String
        Dim result As Integer
        Dim cname As String
        Dim cmdUpdate As String
        Dim cmd As String
        Dim strcmd As String
        Dim cmdChk As String
        Dim cnt As Integer
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim dateS As Date

        Dim firstDate As Date = New Date(DateTime.Now.Year, DateTime.Now.Month, 1)
        Dim endDate As Date = firstDate.AddDays(-1)
        cname = "CSCSCD,CSACCD,CSCSNM,CSTYPE,CSCUTM,CSCSGP,CSSECT,CSADDT,CSADTM,CSEFFD,CSENDD"


        cmd = "insert into  " & strLib & ".TACSMA(" & cname & ") values ({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},0)"
        cmd = String.Format(cmd, Utility.AddSingleQuoat(txtCustCode.Text), Utility.AddSingleQuoat(txtAccCode.Text), Utility.AddSingleQuoat(txtName.Text), GetInteger(txtType.Text), _
                                Utility.AddSingleQuoat(txtCustom.Text), Utility.AddSingleQuoat(txtGroup.Text), _
                                Utility.AddSingleQuoat(txtSection.Text), Utility.GetDate, Utility.GetTime, firstDate.ToString("yyyyMMdd"))

        cmdUpdate = "update " & strLib & ".TACSMA set CSENDD={0} where CSCSCD = {1} and CSENDD = 0"
        cmdUpdate = String.Format(cmdUpdate, endDate.ToString("yyyyMMdd"), Utility.AddSingleQuoat(txtCustCode.Text))
        cmdChk = "select count(*) from " & strLib & ".TACSMA where CSCSCD = {0}"
       
        strcmd = String.Format(cmdChk, Utility.AddSingleQuoat(txtCustCode.Text))
        cnt = CInt(DbConnect.ExcuteScalar(strcmd, DBConnection.DatabaseType.AS400))

        If cnt > 0 Then
            If Request("state") = "add" Then
                MessageBox("Warning : Customer Code is Duplicated.")
                Exit Sub
            End If
            result = DbConnect.ExcuteNonQueryString(cmdUpdate, DBConnection.DatabaseType.AS400)
            If result > 0 Then
                result = DbConnect.ExcuteNonQueryString(cmd, DBConnection.DatabaseType.AS400)
                If result > 0 Then
                    MessageBox("Successful : Update customer completed.")
                Else
                    MessageBox("Found Error : Cannot add customer.")
                End If
            Else
                MessageBox("Found Error : Cannot update customer.")
            End If
        Else
            result = DbConnect.ExcuteNonQueryString(cmd, DBConnection.DatabaseType.AS400)
            If result > 0 Then
                MessageBox("Successful : Added customer completed.")
            Else
                MessageBox("Found Error : Cannot add customer.")
            End If
        End If
        SummarizeAll()
        Binddata()
    End Sub
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
    Public Sub ClearSummarize()

        Dim result As Integer
        Dim str As String

        str = "delete " & strLib & ".TASMCC"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)

        str = "delete " & strLib & ".TASMAG"
        result = DbConnect.ExcuteNonQueryString(str, DBConnection.DatabaseType.AS400)
    End Sub

    Public Sub MessageBox(ByVal message As String)
        ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), True)
    End Sub
End Class