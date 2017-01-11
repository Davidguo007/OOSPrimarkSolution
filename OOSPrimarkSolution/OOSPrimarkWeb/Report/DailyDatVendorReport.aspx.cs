using OOSCommon;
using OOSLibrary.PrimarkOriginDB.Business;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace OOSPrimarkWeb.Report
{
    public partial class DailyDatVendorReport : System.Web.UI.Page
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
SELECT     [Date], Tag_Type, Quantity, Supplier_Ref_code, Supplier_Name_Address1, Supplier_Name_Address2, 
                      Supplier_Name_Address3, Supplier_Name_Address4, Supplier_Name_Address5, Supplier_Name_Address6, Contact1Name_or_PhoneNumber, 
                      Contact2Name_or_PhoneNumber, Tag_Country_Code, Tag_Print_CountryID, User_emailaddress
FROM         OriginOrderDetail
WHERE     (FileID = '{0}')
order by Tag_Country_Code, Tag_Type, Supplier_Ref_code", _fileId));

            DataTable dtNew = new DataTable();
            dtNew = dtDetailGroup.Clone();
            List<string> _lstCCTTSR = new List<string>();
            foreach (DataRow dr in dtDetailGroup.Rows)
            {
                string _keys = dr["Tag_Country_Code"].ToString().Trim() + "^" + dr["Tag_Type"].ToString().Trim() + "^" + dr["Supplier_Ref_code"].ToString().Trim();
                if (!_lstCCTTSR.Contains(_keys))
                {
                    _lstCCTTSR.Add(_keys);
                }
            }

            foreach (string item in _lstCCTTSR)
            {
                string[] _Arra = item.Split('^');
                int _TQTY = 0;
                foreach (DataRow dr in dtDetailGroup.Rows)
                {
                    if (_Arra[0] == dr["Tag_Country_Code"].ToString().Trim() && _Arra[1] == dr["Tag_Type"].ToString().Trim() && _Arra[2] == dr["Supplier_Ref_code"].ToString().Trim())
                    {
                        _TQTY += int.Parse(dr["Quantity"].ToString());
                        dtNew.ImportRow(dr);
                    }
                }
                DataRow ndr = dtNew.NewRow();
                ndr["Quantity"] = _TQTY;
                ndr["Tag_Country_Code"] = _Arra[0];
                dtNew.Rows.Add(ndr);                
            }

            this.gvDetailRpt.DataSource = dtNew;
            this.gvDetailRpt.DataBind();

            return dtNew;
        }

        protected void btnExportExcel_head_Click(object sender, EventArgs e)
        {

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