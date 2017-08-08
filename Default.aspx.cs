using System;
using System.Data;
using SAPFunctionsOCX;
using SAPTableFactoryCtrl;
using ADAuth;
using ExportNs;
using System.Linq;
using System.Collections.Generic;
using System.Web.UI.WebControls;

namespace mtlList
{
    public partial class _Default : System.Web.UI.Page
    {
        private string clientNum;
        private string rfcName;

        public string sDate1 { get; private set; }
        public string sDate2 { get; private set; }
        public string sPO { get; private set; }
        public string sMaterial { get; private set; }
        public string sZflag { get; private set; }
        public string sMtlDoc { get; private set; }
        public string sRefDoc { get; private set; }
        public string sVendorName { get; private set; }
        public string sVendorId { get; private set; }
        public string sMvt { get; private set; }
        public string FormName { get; set; }
        public string FormPaperType { get; private set; }
        public DataTable sapDataDt { get; set; }
        public DataTable originDT { get; set; }
        public DataTable formatedDT { get; set; }

        public static T CType<T>(object obj)
        {
            try { return (T)obj; }
            catch { return default(T); }

        }

        protected void Page_Load(object sender, EventArgs e)
        {
            clientNum = "620";
            rfcName = "ZMMRFC002";
            FormName = "進料狀況表";
            FormPaperType = "A4L";

            originDT = new DataTable();
            sapDataDt = new DataTable();
            formatedDT = new DataTable();

            Auth auth = new Auth();

            string domainUserName = auth.GetDomainUserName();
            string currentUserID = auth.GetUserID(domainUserName);
            var groupList = auth.GetGroupLists(currentUserID);
            var isInGroup = auth.SearchInGroups(groupList);

            //使用者不在允許名單
            if (isInGroup == false)
            {
                Response.Redirect("disabled.html");
            }
        }


        protected void btnQry_Click(object sender, EventArgs e)
        {
            gvData.Visible = true;
            lblMsg.Text = "";
    
            //Create a new thread and set the method test() run in this thread
            System.Threading.Thread s = new System.Threading.Thread(new System.Threading.ThreadStart(connSAP));
            //Set the run mode 'STA'
            s.SetApartmentState(System.Threading.ApartmentState.STA);
            s.Start();
            s.Join();
        }
        
        private void connSAP()
        {
            Sapcon oLogon = new Sapcon(clientNum);
            SAPLogonCtrl.Connection oSAPConn = CType<SAPLogonCtrl.Connection>(oLogon.NewConnection());
            
            
            try
            {
                if (oSAPConn.Logon(0, true))
                {
                    sDate1 = txtDate1.Text.Trim();
                    sDate2 = txtDate2.Text.Trim();
                    sPO = txtPO.Text.Trim();
                    if (ddlMvt.SelectedIndex > 0) sMvt = ddlMvt.SelectedValue.ToString();
                    sMtlDoc = txtMtlDoc.Text.Trim();
                    sRefDoc = txtRefDoc.Text.Trim();
                    sVendorName = txtVendorName.Text.Trim();
                    sMaterial = txtMaterial.Text.Trim();
                    sZflag = convertFlag(cbZflag.Checked);

                    SAPFunctionsClass func = new SAPFunctionsClass();
                    func.Connection = oSAPConn;

                    if (string.IsNullOrEmpty(sDate1)) sDate1 = DateTime.Today.ToString("yyyyMMdd");
                    if (string.IsNullOrEmpty(sDate2)) sDate2 = sDate1;

                    //功能模組名稱
                    IFunction ifunc = (IFunction)func.Add(rfcName);
                    //查詢參數：起始日期
                    IParameter pDate1 = (IParameter)ifunc.get_Exports("DATE1");
                    pDate1.Value = sDate1;
                    //查詢參數：結束日期
                    IParameter pDate2 = (IParameter)ifunc.get_Exports("DATE2");
                    if (txtDate2.Text == "") pDate2.Value = sDate1;
                    else pDate2.Value = sDate2;
                    //查詢參數：未產生105
                    IParameter pFlag = (IParameter)ifunc.get_Exports("ZFLAG");
                    pFlag.Value = sZflag;
                    //查詢參數：採購單
                    IParameter pPO = (IParameter)ifunc.get_Exports("PO");
                    pPO.Value = sPO;
                    //查詢參數：物料號碼
                    IParameter pMaterial = (IParameter)ifunc.get_Exports("MATERIAL");
                    pMaterial.Value =  sMaterial;
                    //查詢參數：廠商名稱
                    IParameter pVendorName  = (IParameter)ifunc.get_Exports("VENDORNM");
                    pVendorName.Value = sVendorName;
                    //查詢參數：異動類型
                    IParameter pMvt = (IParameter)ifunc.get_Exports("MVT");
                    pMvt.Value = sMvt;
                    //查詢參數：物料文件
                    IParameter pMtlDoc = (IParameter)ifunc.get_Exports("MTLDOC");
                    pMtlDoc.Value = sMtlDoc;
                    //查詢參數：參考文件
                    IParameter pRefDoc = (IParameter)ifunc.get_Exports("REFDOC");
                    pRefDoc.Value = sRefDoc;

                    ifunc.Call();

                    Tables tables = (Tables)ifunc.Tables;
                    SAPTableFactoryCtrl.Table ITAB = (SAPTableFactoryCtrl.Table)tables.get_Item("ITAB");

                    int itabRowCount = ITAB.RowCount;

                    if (itabRowCount > 0)
                    {
                        for (int i = 1; i <= itabRowCount; i++)
                        {
                            DataRow dr = sapDataDt.NewRow();
                            if (i == 1)
                            {
                                //  dt.Columns.Add("No");
                                sapDataDt.Columns.Add("輸入日期");
                                sapDataDt.Columns.Add("輸入時間");
                                sapDataDt.Columns.Add("採購文件");
                                sapDataDt.Columns.Add("文件項次");
                                sapDataDt.Columns.Add("物料文件");
                                sapDataDt.Columns.Add("異動類型");
                                sapDataDt.Columns.Add("參考文件");
                                sapDataDt.Columns.Add("數量");
                                sapDataDt.Columns.Add("單價");
                                sapDataDt.Columns.Add("幣別");
                                sapDataDt.Columns.Add("料號");
                                sapDataDt.Columns.Add("工單");
                                sapDataDt.Columns.Add("群組說明");
                                sapDataDt.Columns.Add("品名");
                                sapDataDt.Columns.Add("供應商");
                                sapDataDt.Columns.Add("備註");

                            }

                            //dr["No"] = i.ToString();
                            dr["輸入日期"] = Convert.ToDateTime(ITAB.get_Cell(i, "CPUDT")).ToString("yyyy-MM-dd");
                            dr["輸入時間"] = Convert.ToDateTime(ITAB.get_Cell(i, "CPUTM")).ToString("HH:mm:ss");
                            dr["採購文件"] = ITAB.get_Cell(i, "EBELN").ToString();
                            dr["文件項次"] = ITAB.get_Cell(i, "EBELP").ToString().TrimStart('0');
                            dr["物料文件"] = ITAB.get_Cell(i, "BELNR").ToString();
                            dr["異動類型"] = ITAB.get_Cell(i, "BWART").ToString();
                            dr["參考文件"] = ITAB.get_Cell(i, "LFBNR").ToString();
                            dr["數量"] = ITAB.get_Cell(i, "MENGE").ToString().TrimEnd('0').TrimEnd('.');
                            dr["單價"] = ITAB.get_Cell(i, "NETPR").ToString();
                            dr["幣別"] = ITAB.get_Cell(i, "WAERS").ToString();
                            dr["料號"] = ITAB.get_Cell(i, "MATNR").ToString().TrimStart('0');
                            dr["工單"] = ITAB.get_Cell(i, "AUFNR").ToString().TrimStart('0');
                            dr["群組說明"] = ITAB.get_Cell(i, "TEXT20").ToString();
                            dr["品名"] = ITAB.get_Cell(i, "TXZ01").ToString();
                            var vendorName = checkStringLen(ITAB.get_Cell(i, "NAME1").ToString(),4);

                            dr["供應商"] = vendorName;
                            dr["備註"] = ITAB.get_Cell(i, "MD_MEMO").ToString().TrimEnd(':');

                            sapDataDt.Rows.Add(dr);
                        }
                        formatedDT = arrangeExcelLayout(sapDataDt);

                        gvData.DataSource = formatedDT.DefaultView;
                        gvData.DataBind();
                        btnConvert.Visible = true;
                    }
                    else
                    {
                        lblMsg.Text = "無符合條件的資料";
                        gvData.Visible = false;
                    }
                }
                else
                {
                    lblMsg.Text = "無法連接 sap";
                }
            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
        }

        /// <summary>
        /// 檢查字串長度 v1 是否超過 v2，若超過就縮短
        /// </summary>
        /// <param name="v1"></param>
        /// <param name="v2"></param>
        /// <returns></returns>
        private string checkStringLen(string v1, int v2)
        {
            var result = "";
            if (v1.Length > v2)
            {
                result = v1.Substring(0, v2);
            }
            else result = v1;

            return result;
        }

        private string convertFlag(bool @checked)
        {
            string flag = "";

            switch(@checked)
            {
                case false:
                    flag = "";
                    break;
                case true:
                    flag = "X";
                    break;
            }
            return flag;
        }

        protected void btnClr_Click(object sender, EventArgs e)
        {
            Response.Redirect("Default.aspx");
        }

        protected void btnDate1_Click(object sender, EventArgs e)
        {
        }
        protected void btnDate2_Click(object sender, EventArgs e)
        {
        }
        protected void cldrDate1_SelectionChanged(object sender, EventArgs e)
        {
        }
        protected void cldrDate2_SelectionChanged(object sender, EventArgs e)
        {
        }
        protected void btnConvert_Click(object sender, EventArgs e)
        {
            ExportExcel ee = new ExportExcel();
            ee.FormName = changeFormName(rdblExcelFormat.SelectedValue);
            ee.FormPaperType = this.FormPaperType;

            ee.ExportToXlsx(gvData);
        }

        private string changeFormName(string selectedValue)
        {
            var result = "";
            switch (selectedValue)
            {
                case "0":
                    result = this.FormName;
                    break;
                case "1":
                    result = rdblExcelFormat.SelectedItem.ToString();
                    break;
                case "2":
                    result = rdblExcelFormat.SelectedItem.ToString();
                    break;
            }
            return result;
        }

        private DataTable arrangeExcelLayout(DataTable arrangeLayoutDt)
        {
            switch (rdblExcelFormat.SelectedValue)
            {
                case "1": //格式：每日列印                    
                    arrangeLayoutDt.Columns.Remove("輸入時間");
                    arrangeLayoutDt.Columns.Remove("參考文件");
                    arrangeLayoutDt.Columns.Remove("群組說明");
                    arrangeLayoutDt.Columns.Remove("備註");
                    break;
                case "2": //格式：對帳用
                    arrangeLayoutDt.Columns.Remove("輸入時間");
                    arrangeLayoutDt.Columns.Remove("文件項次");
                    arrangeLayoutDt.Columns.Remove("物料文件");
                    arrangeLayoutDt.Columns.Remove("異動類型");
                    arrangeLayoutDt.Columns.Remove("參考文件");
                    arrangeLayoutDt.Columns.Remove("工單");
                    arrangeLayoutDt.Columns.Remove("群組說明");
                    break;
            }

            return arrangeLayoutDt;

        }
    }
}
