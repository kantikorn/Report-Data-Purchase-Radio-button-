Imports System.Data.SqlClient
Public Class WebForm1
    Inherits System.Web.UI.Page

    'เชื่อมต่อกับฐานข้อมูล'
    Dim connectionString As String = "Data Source=DESKTOP-LJ58AB0\SQLEXPRESS;Initial Catalog=AdventureWorks2008R2;Integrated Security=True;"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BindDropDownList()
            GridView2.DataSourceID = ""
            GridView2.DataBind()
        End If
    End Sub
    'เมธอดสำหรับเตรียมข้อมูลวันที่'
    Protected Sub BindDropDownList()
        'เชื่อมต่อกับฐานข้อมูล'
        Using con As New SqlConnection(connectionString)
            con.Open()
            'สร้างคำสั่ง sql เพื่อดึงค่าวันที่มาเเสดง'
            Dim query As String = "SELECT DISTINCT CONVERT(VARCHAR, OrderDate, 23) AS FormattedOrderDate FROM Purchasing.PurchaseOrderHeader"

            'ดึงข้อมูลจาก Sql Data Reader'
            Using cmd As New SqlCommand(query, con)
                'เตรียมข้อมูลใน Dropdownlist'
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dd_dateinput.DataSource = reader
                    dd_dateinput.DataTextField = "FormattedOrderDate"
                    dd_dateinput.DataValueField = "FormattedOrderDate"
                    dd_dateinput.DataBind()
                End Using
            End Using
        End Using
    End Sub
    Protected Sub btn_search_Click(sender As Object, e As EventArgs) Handles btn_search.Click
        'สร้างตัวเเปรเพื่อเก็บข้อมูลจาก select date'
        Dim selectDate As String = dd_dateinput.SelectedValue

        'สร้างตัวเก็บข้อมูลเเบบ dataTable  ประกาศตัวเเปร dt'

        Dim dt As New DataTable()

        'เชื่อมต่อกับฐานข้อมูล'
        Using con As New SqlConnection(connectionString)
            con.Open()

            'สร้างคำสั่ง sql เพื่อดึงรายกาารสั่งซื้อสำหรับวันที่ที่เลือก'
            Dim query As String = "SELECT Poh.PurchaseOrderID, Poh.EmployeeID, Poh.VendorID, Pod.OrderQty, Poh.OrderDate
                      FROM Purchasing.PurchaseOrderHeader Poh
                      JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID
                      WHERE CONVERT(VARCHAR, OrderDate, 23) = @SelectDate"


            'ดำเนินการ Execute Sql Command'
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@SelectDate", selectDate)

                'สร้าง DataAdapter เพื่อเตรียมข้อมูลใน DataTable'
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using

            'สร้างคำสั่ง SQL เพื่อดึงผลรวมของ OrderQty สำหรับวันที่ที่เลือก'
            Dim querysum_qty As String = "SELECT ISNULL(SUM(Pod.OrderQty), 0) AS TotalOrderQty " &
                                      "FROM Purchasing.PurchaseOrderHeader Poh " &
                                      "JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID " &
                                      "WHERE CONVERT(DATE, Poh.OrderDate) = @SelectDate"
            'ดำเนินการ Execute Sql Command เพื่อคำนวณผลรวม OrderQty'
            Using cmdsumqty As New SqlCommand(querysum_qty, con)
                cmdsumqty.Parameters.AddWithValue("@SelectDate", selectDate)

                'ดึงผลรวมของ OrderQty'
                Dim totalOrderQty As Integer = Convert.ToInt32(cmdsumqty.ExecuteScalar())

                'เเสดงจำนวนรายการ row count ผลรวม ของ OrderQty'
                lbl_count_order.Text = "รวมป้อน " & dt.Rows.Count.ToString() & " รายการ"
                lbl_sum_qty.Text = "รวม Qty = " & totalOrderQty.ToString()
            End Using
        End Using

        'เอา DataTable มาเพื่อ Bind() เข้า Gridview'
        GridView2.DataSourceID = ""
        GridView2.DataSource = dt
        GridView2.DataBind()
    End Sub

    Protected Sub btn_searchdata_Click(sender As Object, e As EventArgs) Handles btn_searchdata.Click
        ' Get the selected option from RadioButtonList'
        Dim selectedOption As String = RadioButtonList1.SelectedItem.Text

        ' Build the appropriate SQL query based on the selected option'
        Dim query As String = ""

        Select Case selectedOption
            Case "Year ASC"
                query = "SELECT Poh.PurchaseOrderID, Poh.EmployeeID, Poh.VendorID, Pod.OrderQty, Poh.OrderDate
                     FROM Purchasing.PurchaseOrderHeader Poh
                     JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID
                     WHERE YEAR(Poh.OrderDate) = (SELECT MAX(YEAR(OrderDate)) FROM Purchasing.PurchaseOrderHeader)"

            Case "Month ASC"
                query = "SELECT Poh.PurchaseOrderID, Poh.EmployeeID, Poh.VendorID, Pod.OrderQty, Poh.OrderDate
                     FROM Purchasing.PurchaseOrderHeader Poh
                     JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID
                     WHERE YEAR(Poh.OrderDate) = (SELECT MAX(YEAR(OrderDate)) FROM Purchasing.PurchaseOrderHeader)
                     AND MONTH(Poh.OrderDate) = (SELECT MAX(MONTH(OrderDate)) FROM Purchasing.PurchaseOrderHeader WHERE YEAR(OrderDate) = (SELECT MAX(YEAR(OrderDate)) FROM Purchasing.PurchaseOrderHeader))"

            Case "Day ASC"
                query = "SELECT Poh.PurchaseOrderID, Poh.EmployeeID, Poh.VendorID, Pod.OrderQty, Poh.OrderDate
                     FROM Purchasing.PurchaseOrderHeader Poh
                     JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID
                     WHERE CONVERT(DATE, Poh.OrderDate) = (SELECT MAX(CONVERT(DATE, OrderDate)) FROM Purchasing.PurchaseOrderHeader)"
        End Select

        ' Retrieve data from the database based on the constructed query'
        Dim dt As New DataTable()

        Using con As New SqlConnection(connectionString)
            con.Open()

            Using cmd As New SqlCommand(query, con)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using

            ' Calculate and display the total order quantity'
            Dim totalOrderQty As Integer = 0

            If dt.Rows.Count > 0 Then
                totalOrderQty = dt.AsEnumerable().Sum(Function(row) Convert.ToInt32(row("OrderQty")))
            End If

            ' Display the data and total order quantity in the GridView and labels'
            lbl_count_order.Text = "รวมป้อน " & dt.Rows.Count.ToString() & " รายการ"
            lbl_sum_qty.Text = "รวม Qty = " & totalOrderQty.ToString()


            GridView2.DataSourceID = ""
            GridView2.DataSource = dt
            GridView2.DataBind()
        End Using
    End Sub

End Class
