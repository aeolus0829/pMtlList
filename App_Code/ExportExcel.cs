using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Data;
using ClosedXML.Excel;

/// <summary>
/// Summary description for ExportExcel
/// </summary>
public class ExportExcel
{

    public static void ExportToExcel(GridView gridView)
    {
        //
        // TODO: Add constructor logic here
        //
        string FileName = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");
        HttpContext.Current.Response.Clear();

        HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ".xls");
        //HttpContext.Current.Response.AddHeader("content-disposition","attachment;filename=PoolExport.xls");
        //HttpContext.Current.Response.ContentType = "application/vnd.xls";
        System.IO.StringWriter sw = new System.IO.StringWriter();
        System.Web.UI.HtmlTextWriter htw = new HtmlTextWriter(sw);
        htw.Write("<center><table><tr><td colspan='16' align='center'><font size=\"32\">進料狀況表</font></td></tr><tr><td colspan='16'>&nbsp;</td></tr></center>");
        gridView.AllowSorting = false;
        gridView.AllowPaging = false;
        Page page = new Page();
        HtmlForm form = new HtmlForm();
        gridView.EnableViewState = false;
        // Deshabilitar la validación de eventos, sólo asp.net 2
        page.EnableEventValidation = false;
        // Realiza las inicializaciones de la instancia de la clase Page que requieran los diseñadores RAD.
        page.DesignerInitialize();
        page.Controls.Add(form);
        htw.Write("<center>");
        form.Controls.Add(gridView);
        page.RenderControl(htw);
        htw.Write("</center>");
        htw.Write("<table><tr><td colspan='16'></td></tr><tr><td colspan='8' align='right'>審核：________________</td><td colspan='8' align='right'>製表：________________</td></tr>");
        //HttpContext.Current.Response.Write(sw.ToString());
        HttpContext.Current.Response.Write(sw.ToString().Substring(sw.ToString().IndexOf("<table")));
        HttpContext.Current.Response.End();        
    }

    public void exportXlsx(GridView gv)
    {

        DataTable mainDt = GetDataTable(gv);

        string FileName = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");

        using (XLWorkbook wb = new XLWorkbook())
        {
            wb.Worksheets.Add(mainDt, "Sheet1");

            //wb.SaveAs(folderPath + "DataGridViewExport.xlsx");
            string strFileName = HttpContext.Current.Server.UrlEncode(FileName + ".xlsx");
            System.IO.MemoryStream stream = GetStream(wb); // The method is defined below
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=" + strFileName);
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
            HttpContext.Current.Response.BinaryWrite(stream.ToArray());
            HttpContext.Current.Response.End();
        }
    }

    private System.IO.MemoryStream GetStream(XLWorkbook wb)
    {
        System.IO.MemoryStream fs = new System.IO.MemoryStream();
        wb.SaveAs(fs);
        fs.Position = 0;
        return fs;
    }

        DataTable GetDataTable(GridView dtg)
    {
        DataTable dt = new DataTable();

        // add the columns to the datatable           
        if (dtg.HeaderRow != null)
        {

            for (int i = 0; i < dtg.HeaderRow.Cells.Count; i++)
            {
                dt.Columns.Add(dtg.HeaderRow.Cells[i].Text);
            }
        }

        //  add each of the data rows to the table
        foreach (GridViewRow row in dtg.Rows)
        {
            DataRow dr;
            dr = dt.NewRow();

            for (int i = 0; i < row.Cells.Count; i++)
            {
                dr[i] = row.Cells[i].Text.Replace(" ", "");
            }
            dt.Rows.Add(dr);
        }
        return dt;
    }

}