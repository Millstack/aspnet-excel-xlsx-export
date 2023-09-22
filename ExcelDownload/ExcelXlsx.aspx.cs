using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class ExcelXlsx : System.Web.UI.Page
{
    DataTable dt;
    string connectionString = ConfigurationManager.ConnectionStrings["Ginie"].ConnectionString;

    protected void Page_Load(object sender, EventArgs e)
    {
        BindGrid();
    }

    protected void BindGrid()
    {
        using (SqlConnection con = new SqlConnection(connectionString))
        {
            con.Open();
            string sql = "select Name, Email, MobileNo, Password from UserMaster";
            SqlCommand cmd = new SqlCommand(sql, con);
            con.Close();

            SqlDataAdapter ad = new SqlDataAdapter(cmd);
            dt = new DataTable();
            ad.Fill(dt);

            GridUser.DataSource = dt;
            GridUser.DataBind();
        }
    }

    protected void btnExp_Click(object sender, EventArgs e)
    {
        ExcelExport(dt);
    }

    protected void ExcelExport(DataTable dt)
    {
        // Set the license context
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            //Excel sheet name
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("EventList");

            // Add header row
            worksheet.Cells["A1"].Value = "ID";
            worksheet.Cells["B1"].Value = "Name";
            worksheet.Cells["C1"].Value = "Email";
            worksheet.Cells["D1"].Value = "Mobile No";
            worksheet.Cells["E1"].Value = "Password";

            using (var range = worksheet.Cells["A1:AH1"])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue); //color to light blue

                // setting individual column width
                worksheet.Column(2).Width = 15;
                worksheet.Column(3).Width = 15;
                worksheet.Column(4).Width = 15;
               
                // setting individual columns format as text only (in order to prevent that data from getting chance based on different data type)
                worksheet.Cells["D:D"].Style.Numberformat.Format = "@";
                worksheet.Cells["E:E"].Style.Numberformat.Format = "@";

                //ExcelRange fromDateColumn = worksheet.Cells["H:H"];
                //ExcelRange toDateColumn = worksheet.Cells["I:I"];

                //fromDateColumn.Style.Numberformat.Format = "@";
                //toDateColumn.Style.Numberformat.Format = "@";
            }

            int rowIndex = 2;
            int srNo = 1; //For fix Sr no starting from 1

            foreach (DataRow row in dt.Rows)
            {
                // row[1] for position or single column name for better any column missmatch row["UserName"]
                worksheet.Cells["A" + rowIndex].Value = srNo.ToString();
                worksheet.Cells["B" + rowIndex].Value = row["Name"].ToString();
                worksheet.Cells["C" + rowIndex].Value = row["Email"].ToString();
                worksheet.Cells["D" + rowIndex].Value = row["MobileNo"].ToString();
                worksheet.Cells["E" + rowIndex].Value = row["Password"].ToString();

                rowIndex++;
                srNo++;
            }

            byte[] excelBytes = excelPackage.GetAsByteArray();

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename = Excel FIle Name.xlsx");   //Excel file name
            Response.BinaryWrite(excelBytes);
            Response.End();
        }
    }

    
}