using OOSCommon;
using OOSLibrary.PrimarkOriginDB.Business;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace OOSPrimarkWeb.Report
{
    public partial class DailyDatPOReport : System.Web.UI.Page
    {
        private OriginDatBase _OriDatBLL;
        private OriginDatBase OriDataBLL
        {
            get
            {
                if (_OriDatBLL == null)
                {
                    _OriDatBLL = new OriginDatBase();
                }
                return _OriDatBLL;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                BindGrid();
            }
        }

        private void BindGrid()
        {
            DataSet dtDatH = OriDataBLL.GetDBDataSetBySQL(@"
select * from  dbo.OriginOrderHead order by Createtime
select FileID, SUM(Convert(int, Quantity)) as Qty from dbo.OriginOrderDetail group by FileID");

            dtDatH.Tables[0].Columns.Add("Qty", typeof(int));
            foreach (DataRow dr in dtDatH.Tables[0].Rows)
            {
                DataRow[] drs = dtDatH.Tables[1].Select(string.Format(" FileID = '{0}'", dr["FileID"].ToString()));
                dr["Qty"] = drs[0]["Qty"].ToString();
            }

            this.gvDailyDatlist.DataSource = dtDatH;
            this.gvDailyDatlist.DataBind();
        }
        protected void gvDailyDatlist_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            string _FileId = e.CommandArgument.ToString();
            this.hfFileId.Value = _FileId;
            this.btnExportExcel_Detal.Visible = true;
            BindDetailGrid(_FileId);
        }

        private DataTable BindDetailGrid(string _fileId)
        {
            DataTable dtDetailGroup = OriDataBLL.GetDBDataBySQL(string.Format(@"
SELECT     [Date], Tag_Type, Purchase_Order_Number, SUM(CONVERT(int, Quantity)) AS Quantity, Supplier_Ref_code, Supplier_Name_Address1, Supplier_Name_Address2, 
                      Supplier_Name_Address3, Supplier_Name_Address4, Supplier_Name_Address5, Supplier_Name_Address6, Contact1Name_or_PhoneNumber, 
                      Contact2Name_or_PhoneNumber, Tag_Country_Code, User_emailaddress
FROM         OriginOrderDetail
WHERE     (FileID = '{0}')
GROUP BY [Date], Tag_Type, Purchase_Order_Number, Supplier_Ref_code, Supplier_Name_Address1, Supplier_Name_Address2, Supplier_Name_Address3, 
                      Supplier_Name_Address4, Supplier_Name_Address5, Supplier_Name_Address6, Contact1Name_or_PhoneNumber, Contact2Name_or_PhoneNumber, 
                      Tag_Country_Code, User_emailaddress
order by Tag_Country_Code, Purchase_Order_Number", _fileId));

            this.gvDetailRpt.DataSource = dtDetailGroup;
            this.gvDetailRpt.DataBind();

            return dtDetailGroup;
        }

        protected void btnExportExcel_head_Click(object sender, EventArgs e)
        {
            string style = @"<style> .text { mso-number-format:\@; } </style> ";

            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "GB2312";

            string FileName = "Report-" + DateTime.Now.ToFileTime();

            string fileName = HttpUtility.UrlEncode(FileName);// +System.DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss");
            Response.AppendHeader("Content-Disposition", "attachment;filename=" + fileName + ".xls");
            //Response.ContentEncoding = System.Text.Encoding.UTF8;
            Response.ContentType = "application/ms-excel";//设置输出文件类型为excel文件。

            StringWriter sw = new StringWriter();

            HtmlTextWriter htw = new HtmlTextWriter(sw);

            DisableControls(gvDailyDatlist);

            this.gvDailyDatlist.RowStyle.BackColor = System.Drawing.Color.White;
            this.gvDailyDatlist.FooterStyle.BackColor = System.Drawing.Color.White;
            this.gvDailyDatlist.FooterStyle.ForeColor = System.Drawing.Color.Black;
            this.gvDailyDatlist.HeaderStyle.BackColor = System.Drawing.Color.White;
            this.gvDailyDatlist.HeaderStyle.ForeColor = System.Drawing.Color.Black;
            this.gvDailyDatlist.HeaderStyle.Height = 20;
            this.gvDailyDatlist.AlternatingRowStyle.BackColor = System.Drawing.Color.White;
            this.gvDailyDatlist.BorderWidth = 1;
            this.gvDailyDatlist.Columns[this.gvDailyDatlist.Columns.Count - 1].Visible = false;

            gvDailyDatlist.RenderControl(htw);

            Response.Write(style);

            Response.Output.Write(sw.ToString());
            //Session.Remove("gvReportResult");
            Response.Flush();
            Response.End();
        }

        public override void VerifyRenderingInServerForm(Control control)
        {

        }

        private void DisableControls(Control gv)
        {

            LinkButton lb = new LinkButton();

            Literal l = new Literal();

            string name = String.Empty;

            for (int i = 0; i < gv.Controls.Count; i++)
            {

                if (gv.Controls[i].GetType() == typeof(LinkButton))
                {

                    l.Text = (gv.Controls[i] as LinkButton).Text;

                    gv.Controls.Remove(gv.Controls[i]);

                    gv.Controls.AddAt(i, l);

                }
                else if (gv.Controls[i].GetType() == typeof(DropDownList))
                {
                    l.Text = (gv.Controls[i] as DropDownList).SelectedItem.Text;

                    gv.Controls.Remove(gv.Controls[i]);

                    gv.Controls.AddAt(i, l);

                }

                if (gv.Controls[i].HasControls())
                {
                    DisableControls(gv.Controls[i]);
                }

            }

        }

        protected void btnExportExcel_Detal_Click(object sender, EventArgs e)
        {

            DataSet dsAll = new DataSet();
            DataTable _dtRaw = BindDetailGrid(this.hfFileId.Value);            
            List<string> _lstCountry = new List<string>();

            foreach (DataRow dr in _dtRaw.Rows)
            {
                if (!_lstCountry.Contains(dr["Tag_Country_Code"].ToString().Trim()))
                {
                    _lstCountry.Add(dr["Tag_Country_Code"].ToString().Trim());
                }
            }

            if (_lstCountry.Count <= 1)
            {
                DataTable _dtOneContry = new DataTable();
                _dtOneContry = _dtRaw.Clone();
                _dtOneContry = _dtRaw.Copy();
                _dtOneContry.TableName = "PrimarkT-" + _lstCountry[0];
                dsAll.Tables.Add(_dtOneContry);
            }
            else
            {
                foreach (string item in _lstCountry)
                {
                    DataTable _dtOneContry = new DataTable("PrimarkT-" + item);
                    _dtOneContry = _dtRaw.Clone();
                    _dtOneContry.TableName = "PrimarkT-" + item;

                    foreach (DataRow dr in _dtRaw.Rows)
                    {
                        if (item == dr["Tag_Country_Code"].ToString().Trim())
                        {
                            //DataRow ndr = _dtOneContry.NewRow();
                            _dtOneContry.ImportRow(dr);
                        }
                    }

                    dsAll.Tables.Add(_dtOneContry);
                }
            }

            #region 用NPOI导出

            string FilePath = System.Configuration.ConfigurationManager.AppSettings["OOSExportTempary"];
                //@"E:\CRMWeb\WebDocuments\NLBarcodePageDataExport\";
            
            if (!Directory.Exists(FilePath))
            {
                Directory.CreateDirectory(FilePath);
            }
            string FileName = "Report-DailyPORpt" + DateTime.Now.ToFileTime() + ".xls";
            bool flag = NPOIRenderExcelDataAgent.DataSetToExcelMergeByProductColumn(dsAll, FilePath + FileName, "Purchase_Order_Number");
            if (!flag)
            {
                //sw.WriteLine("Output Faild");
                new Exception("导出Excel失败");
            }

            string _ExcelFile = FilePath + FileName;
            //sw.WriteLine();
            //sw.WriteLine("将 Excel 提供给用户下载." + DateTime.Now);
            FileInfo fi = new FileInfo(_ExcelFile); //excelFile为文件在服务器上的地址
            HttpResponse contextResponse = HttpContext.Current.Response;
            contextResponse.Clear();
            contextResponse.Buffer = true;
            contextResponse.Charset = "GB2312"; //设置了类型为中文防止乱码的出现 
            contextResponse.AppendHeader("Content-Disposition", String.Format("attachment;filename={0}", _ExcelFile)); //定义输出文件和文件名 
            contextResponse.AppendHeader("Content-Length", fi.Length.ToString());
            contextResponse.ContentEncoding = Encoding.Default;
            contextResponse.ContentType = "application/xls";//设置输出文件类型为excel文件。 


            contextResponse.WriteFile(fi.FullName);
            contextResponse.Flush();
            contextResponse.End();

            #endregion
        }
    }
}