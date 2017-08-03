using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.UI;
using System.Web.UI.HtmlControls;

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
        HttpContext.Current.Response.ContentType = "application/vnd.xls";
        System.IO.StringWriter sw = new System.IO.StringWriter();
        System.Web.UI.HtmlTextWriter htw = new HtmlTextWriter(sw);
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
        form.Controls.Add(gridView);
        page.RenderControl(htw);
        //HttpContext.Current.Response.Write(sw.ToString());
        HttpContext.Current.Response.Write(sw.ToString().Substring(sw.ToString().IndexOf("<table")));
        HttpContext.Current.Response.End();
        
    }
}