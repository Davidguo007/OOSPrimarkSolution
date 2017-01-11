using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOSLibrary.PrimarkOriginDB.Business
{
    public class PrimarkOrig
    {
        public static DataSet DatToDataSet(List<string> list, string filename, ref StringWriter sw)
        {
            DataSet ds = DatFileTable();
            try
            {
                foreach (string item in list)
                {
                    #region Header
                    string split = "\",\"";
                    string[] datalist = item.Replace(split, "^").Split('^');
                    if (datalist.Length == ds.Tables["OrderHeader"].Columns.Count - 1)
                    {
                        DataRow dr = ds.Tables["OrderHeader"].NewRow();

                        for (int i = 0; i < datalist.Length; i++)
                        {
                            dr[i] = datalist[i].Replace("\"", "");
                        }
                        //dr["FileName"] = filename; //另一张表头 OriginOrderHead 中有存，所以此表中无需再存
                        dr["OrigFile"] = item;
                        ds.Tables["OrderHeader"].Rows.Add(dr);
                    }
                    #endregion

                }
                return ds;
            }
            catch (Exception ex)
            {
                sw.WriteLine(ex.Message);
                sw.WriteLine(ex.StackTrace);
                return ds;
            }
        }
        public static DataSet DatFileTable()
        {
            DataTable header = new DataTable("OrderHeader");

            header.Columns.Add("Line_No");
            header.Columns.Add("Date");
            header.Columns.Add("Tag_Type");
            header.Columns.Add("Tag_Sequence_Number_for_PO");
            header.Columns.Add("Kimball");
            header.Columns.Add("Purchase_Order_Number");
            header.Columns.Add("Size");
            header.Columns.Add("Colour");
            header.Columns.Add("Base_Currency_Sell_Price");
            header.Columns.Add("Punt_Sell_Price");
            header.Columns.Add("Quantity");
            header.Columns.Add("Supplier_Ref_code");
            header.Columns.Add("Supplier_Name_Address1");
            header.Columns.Add("Supplier_Name_Address2");
            header.Columns.Add("Supplier_Name_Address3");
            header.Columns.Add("Supplier_Name_Address4");
            header.Columns.Add("Supplier_Name_Address5");
            header.Columns.Add("Supplier_Name_Address6");
            header.Columns.Add("Contact1Name_or_PhoneNumber");
            header.Columns.Add("Contact2Name_or_PhoneNumber");
            header.Columns.Add("Date_Required");
            header.Columns.Add("Urgency_indicator");
            header.Columns.Add("Short_Style_Description");
            header.Columns.Add("Tag_Country_Code");
            header.Columns.Add("Tag_Print_CountryID");
            header.Columns.Add("VAT_Code");
            header.Columns.Add("User_emailaddress");
            header.Columns.Add("Department_Section_SubSection");
            header.Columns.Add("Size_UK_Ireland");
            header.Columns.Add("Size_UK_IRL");
            header.Columns.Add("Size_ES_FR");
            header.Columns.Add("Size_EUR");
            header.Columns.Add("Size_USA");
            header.Columns.Add("IP_StyleCode");
            header.Columns.Add("IP_ColourCode");
            header.Columns.Add("IP_SizeCode");
            header.Columns.Add("IP_SKUCode");
            header.Columns.Add("IP_EAN13_PrimaryEAN");
            header.Columns.Add("IP_ClassNumber");
            header.Columns.Add("IP_VendorNumber");
            header.Columns.Add("Size_Italy");
            //header.Columns.Add("FileName"); //另一张表头 OriginOrderHead 中有存，所以此表中无需再存
            header.Columns.Add("OrigFile");


            DataSet ds = new DataSet();
            ds.Tables.Add(header);
            //ds.Tables.Add(detail);

            return ds;
        }
    }
}
