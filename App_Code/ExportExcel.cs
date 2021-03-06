﻿using System;
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
/// 

namespace ExportNs {
    public class ExportExcel
    {
        public string FormName { get; set; }
        public string FormPaperType { get; set; }
        public void Html2Xls(GridView gridView)
        {
            /// TODO: Add constructor logic here
            string FileName = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");
            HttpContext.Current.Response.Clear();

            HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=" + FileName + ".xls");
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

        /// <summary>       
        /// <para>FormName 報表名稱，列印時作為抬頭使用</para>
        /// <para>FormPaperType 報表的紙張格式, A4P / A4L / A3P / A3L</para>
        /// </summary>
        /// <param name="gv"></param>
        public void ExportToXlsx(GridView gv)
        {
            DataTable mainDt = GetDataTable(gv);

            string FileName = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");            

            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add(mainDt, "Sheet1");

                formatWorkSheet(ws, FormPaperType);

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

        private void formatWorkSheet(IXLWorksheet ws, string paperFormat)
        {
            checkIfEmpty("FormName", FormName);
            checkIfEmpty("FormPaperType", FormPaperType);

            ws.PageSetup.Header.Center.AddText("富荃工業股份有限公司\n").SetFontSize(36).SetBold();
            ws.PageSetup.Header.Center.AddText(FormName).SetFontSize(24).SetBold();

            ws.PageSetup.Footer.Center.AddText("會計：___________________        審核：___________________        製表：___________________").SetFontSize(16);

            //紙張右下角放  n / m , n=頁數, m=總頁數
            ws.PageSetup.Footer.Right.AddText(XLHFPredefinedText.PageNumber, XLHFOccurrence.AllPages).SetFontSize(14);
            ws.PageSetup.Footer.Right.AddText(" / ", XLHFOccurrence.AllPages).SetFontSize(14);
            ws.PageSetup.Footer.Right.AddText(XLHFPredefinedText.NumberOfPages, XLHFOccurrence.AllPages).SetFontSize(14);

            //設定列印標題列
            ws.PageSetup.SetRowsToRepeatAtTop("1:1");

            //列印內容縮放到頁寬
            ws.PageSetup.PagesWide = 1;

            //所有欄依內容自動調整欄寬
            ws.Columns().AdjustToContents();

            ws.PageSetup.SetHorizontalDpi(300);
            ws.PageSetup.SetVerticalDpi(300);
            ws.PageSetup.AlignHFWithMargins = true;
            ws.PageSetup.CenterHorizontally = true;

            ws.PageSetup.Margins.Header = cmToInch(0.8);
            ws.PageSetup.Margins.Footer = cmToInch(0.8);

            switch (paperFormat)
            {
                case "A4L":
                    ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
                    ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;

                    //設定紙張邊界，單位：英吋    
                    ws.PageSetup.Margins.Top = cmToInch(4);
                    ws.PageSetup.Margins.Bottom = cmToInch(1.9);
                    ws.PageSetup.Margins.Left = cmToInch(0.6);
                    ws.PageSetup.Margins.Right = cmToInch(0.6);
                    break;
                case "A4P":
                    break;
                case "A3L":
                    break;
                case "A3P":
                    break;  
            }

        }

        private void checkIfEmpty(string key, string value)
        {
            if (string.IsNullOrEmpty(value)) throw new NotImplementedException(key);
        }

        private double cmToInch(double cm)
        {
            const double inch = 0.393700787;
            var result = cm * inch;
            return result;
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
                    dr[i] = row.Cells[i].Text.Replace("&nbsp;", "");
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}