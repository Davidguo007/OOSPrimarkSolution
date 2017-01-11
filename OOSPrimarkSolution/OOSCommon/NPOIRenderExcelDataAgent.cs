using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;

namespace OOSCommon
{
    public class NPOIRenderExcelDataAgent
    {
        public NPOIRenderExcelDataAgent()
        {

        }
        
        /// <summary>
        /// 将 DataSet 导出到 一个 Excel 中， 每一个 Table 对应一个 Sheet
        /// </summary>
        /// <param name="_ds"></param>
        /// <param name="_Path"></param>
        /// <returns></returns>
        public static bool DataSetToExcel(DataSet _ds, string _Path)
        {
            try
            {
                #region NPOI 导出方式
                HSSFWorkbook hw = new HSSFWorkbook();

                #region styleH for 报表 Body 列名，正常字体，加粗，不加边框 左对齐
                IFont fontH = hw.CreateFont();
                fontH.Boldweight = (short)FontBoldWeight.BOLD;

                ICellStyle styleH = hw.CreateCellStyle();
                styleH.SetFont(fontH);
                styleH.Alignment = HorizontalAlignment.CENTER;
                styleH.VerticalAlignment = VerticalAlignment.CENTER;

                #endregion

                #region output every sheet

                for (int t = 0; t < _ds.Tables.Count; t++)
                {
                    HSSFSheet sheet2 = (HSSFSheet)hw.CreateSheet(_ds.Tables[t].TableName);

                    HSSFRow rowCol2 = (HSSFRow)sheet2.CreateRow(0);
                    for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                    {
                        HSSFCell cell = (HSSFCell)rowCol2.CreateCell(j);
                        //cell.SetCellValue("1000000000000000000000000000000000000000000");
                        cell.SetCellValue(_ds.Tables[t].Columns[j].ColumnName);
                        cell.CellStyle = styleH;
                    }

                    for (int i = 0; i < _ds.Tables[t].Rows.Count; i++)
                    {
                        HSSFRow row = (HSSFRow)sheet2.CreateRow(i + 1);
                        for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                        {
                            HSSFCell cell = (HSSFCell)row.CreateCell(j);
                            string _celldata = _ds.Tables[t].Rows[i][j].ToString();
                            cell.SetCellValue(_celldata);
                            if (_celldata.Trim().Contains("<br>"))
                            {
                                cell.SetCellValue(_celldata.Trim().Replace("<br>", "\r\n"));
                                cell.CellStyle.WrapText = true;
                            }
                            cell.CellStyle.Alignment = HorizontalAlignment.CENTER;
                            cell.CellStyle.VerticalAlignment = VerticalAlignment.CENTER;
                        }
                    }
                }


                #endregion

                //hw.Write(System.IO.Stream(@"c:\poi0.xls"));
                FileStream file = new FileStream(_Path, FileMode.Create);
                hw.Write(file);
                file.Close();

                #endregion

                //sw.WriteLine("生产的excel 建立完成，再次检查文件是否成功建立");
                #region check if the excel file is created
                if (!File.Exists(_Path))
                {
                    //sw.WriteLine("生产的excel 未建立成功");
                    return false;
                }
                #endregion


                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        /// <summary>
        /// 将 DataSet 导出到 一个 Excel 中， 每一个 Table 对应一个 Sheet
        /// 注:此方法适用于导出的数据的一行中有用 br 字段分隔多个产品的的情况，并且产品字段的下一个字段为数量列，并且产品的列名为Product
        /// 此方法是为了导出的Excel的数量列可使用Excel的自动汇总功能，一个订单有多少个产品就会有多少行，其它列会合并成一行。
        /// </summary>
        /// <param name="_ds"></param>
        /// <param name="_Path"></param>
        /// <returns></returns>
        public static bool DataSetToExcelMergeByProductColumn(DataSet _ds, string _Path, string _ProdColumnName)
        {
            //string _SheetNameStr = "";
            //string _theCurrentSheet = "";
            //int _OutSheetCount = 0;

            try
            {
                #region NPOI 导出方式
                HSSFWorkbook hw = new HSSFWorkbook();

                #region styleH for 报表 Body 列名，正常字体，加粗，不加边框
                IFont fontH = hw.CreateFont();
                fontH.Boldweight = (short)FontBoldWeight.BOLD;

                ICellStyle styleH = hw.CreateCellStyle();
                styleH.SetFont(fontH);
                styleH.Alignment = HorizontalAlignment.CENTER;
                styleH.VerticalAlignment = VerticalAlignment.CENTER;


                IFont fontB = hw.CreateFont();
                ICellStyle styleB = hw.CreateCellStyle();
                styleB.SetFont(fontB);
                styleB.Alignment = HorizontalAlignment.CENTER;
                styleB.VerticalAlignment = VerticalAlignment.CENTER;
                #endregion

                #region output every sheet

                for (int t = 0; t < _ds.Tables.Count; t++)
                {
                    //_OutSheetCount = t;
                    //_SheetNameStr += _ds.Tables[t].TableName + "^^";
                    //_theCurrentSheet = _ds.Tables[t].TableName;

                    HSSFSheet sheet2 = (HSSFSheet)hw.CreateSheet(_ds.Tables[t].TableName);

                    HSSFRow rowCol2 = (HSSFRow)sheet2.CreateRow(0);
                    int _prodPosition = 0;

                    for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                    {
                        HSSFCell cell = (HSSFCell)rowCol2.CreateCell(j);
                        //cell.SetCellValue("1000000000000000000000000000000000000000000");
                        cell.SetCellValue(_ds.Tables[t].Columns[j].ColumnName);
                        cell.CellStyle = styleH;

                        if (_ds.Tables[t].Columns[j].ColumnName.Trim().ToLower() == _ProdColumnName.Trim().ToLower())
                        {
                            _prodPosition = j;
                        }
                    }

                    int _theVeriablRowcount = 0;

                    for (int i = 0; i < _ds.Tables[t].Rows.Count; i++)
                    {
                        //add every row data
                        //判断该行数据是否有 包含 <br> , 如果有，则要分成多行，但不含 <br> 的行要合并。
                        bool _isMultLine = false;
                        #region 判断该行数据是否有 包含 <br> , 如果有，则要分成多行，但不含 <br> 的行要合并。

                        Dictionary<string, int> _dicPdQty = new Dictionary<string, int>();
                        for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                        {
                            string _CellProdData = _ds.Tables[t].Rows[i][_prodPosition].ToString();
                            if (_CellProdData.Trim().Contains("<br>"))
                            {
                                _isMultLine = true;
                                string _cellQty = _ds.Tables[t].Rows[i][_prodPosition + 1].ToString();

                                string[] _arr = _CellProdData.Replace("<br>", "^").Split('^');
                                string[] _arr1 = _cellQty.Replace("<br>", "^").Split('^');

                                for (int z = 0; z < _arr.Length; z++)
                                {
                                    if (_arr1.Length - 1 >= z)
                                    {
                                        _dicPdQty.Add(_arr[z], int.Parse(_arr1[z]));
                                    }
                                    else
                                    {
                                        _dicPdQty.Add(_arr[z], 0);
                                    }
                                }
                                break;
                            }
                        }

                        #endregion

                        if (_isMultLine)
                        {
                            #region

                            int _theRowNO = _theVeriablRowcount + 1;

                            HSSFRow rowDataBody = (HSSFRow)sheet2.CreateRow(_theRowNO);
                            for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                            {
                                if (j != _prodPosition && j != _prodPosition + 1)
                                {
                                    HSSFCell cell = (HSSFCell)rowDataBody.CreateCell(j);
                                    string _celldata = _ds.Tables[t].Rows[i][j].ToString();
                                    cell.SetCellValue(_celldata);
                                    cell.CellStyle = styleB;
                                }
                            }

                            foreach (KeyValuePair<string, int> kvp in _dicPdQty)
                            {
                                if (_theRowNO != _theVeriablRowcount + 1)
                                {
                                    rowDataBody = (HSSFRow)sheet2.CreateRow(_theRowNO);
                                }
                                HSSFCell cellDataProd = (HSSFCell)rowDataBody.CreateCell(_prodPosition);
                                cellDataProd.SetCellValue(kvp.Key);

                                HSSFCell cellDataQty = (HSSFCell)rowDataBody.CreateCell(_prodPosition + 1);
                                cellDataQty.SetCellValue(kvp.Value);

                                _theRowNO++;

                            }

                            for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                            {
                                if (j != _prodPosition && j != _prodPosition + 1)
                                {
                                    CellRangeAddress celrang = new CellRangeAddress(_theRowNO - 2, _theRowNO - 1, j, j);
                                    sheet2.AddMergedRegion(celrang);

                                    //sheet2.AutoSizeColumn(j);
                                }
                            }

                            _theVeriablRowcount += _dicPdQty.Count;

                            #endregion
                        }
                        else
                        {

                            HSSFRow row = (HSSFRow)sheet2.CreateRow(_theVeriablRowcount + 1);
                            for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                            {
                                HSSFCell cell = (HSSFCell)row.CreateCell(j);
                                string _celldata = _ds.Tables[t].Rows[i][j].ToString();
                                cell.SetCellValue(_celldata);
                                if (_celldata.Trim().Contains("<br>"))
                                {
                                    cell.SetCellValue(_celldata.Trim().Replace("<br>", "\r\n"));
                                    cell.CellStyle.WrapText = true;
                                }
                                else
                                {
                                    if (j == _prodPosition + 1)
                                    {
                                        int _testInt = 0;
                                        if (int.TryParse(_celldata, out _testInt))
                                        {
                                            cell.SetCellValue(_testInt);
                                        }
                                    }
                                }
                                cell.CellStyle.Alignment = HorizontalAlignment.CENTER;
                                cell.CellStyle.VerticalAlignment = VerticalAlignment.CENTER;
                            }

                            _theVeriablRowcount++;
                        }
                    }
                }

                #endregion

                //hw.Write(System.IO.Stream(@"c:\poi0.xls"));
                FileStream file = new FileStream(_Path, FileMode.Create);
                hw.Write(file);
                file.Close();

                #endregion

                //sw.WriteLine("生产的excel 建立完成，再次检查文件是否成功建立");
                #region check if the excel file is created
                if (!File.Exists(_Path))
                {
                    //sw.WriteLine("生产的excel 未建立成功");
                    return false;
                }
                #endregion


                return true;
            }
            catch (Exception ex)
            {
                //string _Result = "一个处理了[" + _OutSheetCount + "]个, All = " + _SheetNameStr;

                throw new Exception(ex.Message, ex);
            }
        }

        /// <summary>
        /// 将 DataSet 导出到 一个 Excel 中， 每一个 Table 对应一个 Sheet
        /// 注:此方法适用于导出的数据的一行中有用 br 字段分隔多个产品的的情况，并且产品字段的下一个字段为数量列，并且产品的列名为Product
        /// 此方法是为了导出的Excel的数量列可使用Excel的自动汇总功能，一个订单有多少个产品就会有多少行，其它列会合并成一行。
        /// </summary>
        /// <param name="_ds"></param>
        /// <param name="_Path"></param>
        /// <returns></returns>
        public static bool DataSetToExcelMergeByProductColumnLadmark(DataSet _ds, string _Path, string _ProdColumnName, Dictionary<int, int[]> dis)
        {
            //string _SheetNameStr = "";
            //string _theCurrentSheet = "";
            //int _OutSheetCount = 0;

            try
            {
                #region NPOI 导出方式
                IWorkbook hw = new HSSFWorkbook();

                #region styleH for 报表 Body 列名，正常字体，加粗，不加边框
                IFont fontH = hw.CreateFont();
                fontH.Boldweight = (short)FontBoldWeight.BOLD;

                ICellStyle styleH = hw.CreateCellStyle();
                styleH.SetFont(fontH);
                styleH.Alignment = HorizontalAlignment.CENTER;
                styleH.VerticalAlignment = VerticalAlignment.CENTER;


                IFont fontB = hw.CreateFont();
                ICellStyle styleB = hw.CreateCellStyle();
                styleB.SetFont(fontB);
                styleB.Alignment = HorizontalAlignment.CENTER;
                styleB.VerticalAlignment = VerticalAlignment.CENTER;

                //日期
                ICellStyle cellStyle2 = hw.CreateCellStyle();
                IDataFormat format = hw.CreateDataFormat();
                cellStyle2.DataFormat = format.GetFormat("yyyy/mm/dd");
                cellStyle2.SetFont(fontB);

                //数量
                ICellStyle cellStyle3 = hw.CreateCellStyle();
                cellStyle3.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");
                cellStyle3.SetFont(fontB);

                //价格
                ICellStyle cellStyle4 = hw.CreateCellStyle();
                cellStyle4.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0000_");
                cellStyle4.SetFont(fontB);

                //金额
                ICellStyle cellStyle5 = hw.CreateCellStyle();
                cellStyle5.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.00");
                cellStyle5.SetFont(fontB);

                #endregion

                #region output every sheet

                for (int t = 0; t < _ds.Tables.Count; t++)
                {
                    //_OutSheetCount = t;
                    //_SheetNameStr += _ds.Tables[t].TableName + "^^";
                    //_theCurrentSheet = _ds.Tables[t].TableName;

                    HSSFSheet sheet2 = (HSSFSheet)hw.CreateSheet(_ds.Tables[t].TableName);

                    HSSFRow rowCol2 = (HSSFRow)sheet2.CreateRow(0);
                    //int _prodPosition = 0;

                    for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                    {
                        HSSFCell cell = (HSSFCell)rowCol2.CreateCell(j);
                        //cell.SetCellValue("1000000000000000000000000000000000000000000");
                        cell.SetCellValue(_ds.Tables[t].Columns[j].ColumnName);


                        cell.CellStyle = styleH;


                        //if (_ds.Tables[t].Columns[j].ColumnName.Trim().ToLower() == _ProdColumnName.Trim().ToLower())
                        //{
                        //    _prodPosition = j;
                        //}
                    }

                    int _theVeriablRowcount = 0;

                    for (int i = 0; i < _ds.Tables[t].Rows.Count; i++)
                    {
                        //add every row data
                        //判断该行数据是否有 包含 <br> , 如果有，则要分成多行，但不含 <br> 的行要合并。
                        //bool _isMultLine = false;
                        //#region 判断该行数据是否有 包含 <br> , 如果有，则要分成多行，但不含 <br> 的行要合并。

                        //Dictionary<string, int> _dicPdQty = new Dictionary<string, int>();
                        //for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                        //{
                        //    string _CellProdData = _ds.Tables[t].Rows[i][_prodPosition].ToString();
                        //    if (_CellProdData.Trim().Contains("<br>"))
                        //    {
                        //        _isMultLine = true;
                        //        string _cellQty = _ds.Tables[t].Rows[i][_prodPosition + 1].ToString();

                        //        string[] _arr = _CellProdData.Replace("<br>", "^").Split('^');
                        //        string[] _arr1 = _cellQty.Replace("<br>", "^").Split('^');

                        //        for (int z = 0; z < _arr.Length; z++)
                        //        {
                        //            if (_arr1.Length - 1 >= z)
                        //            {
                        //                _dicPdQty.Add(_arr[z], int.Parse(_arr1[z]));
                        //            }
                        //            else
                        //            {
                        //                _dicPdQty.Add(_arr[z], 0);
                        //            }
                        //        }
                        //        break;
                        //    }
                        //}

                        //#endregion

                        //if (_isMultLine)
                        //{
                        //    #region

                        //    int _theRowNO = _theVeriablRowcount + 1;

                        //    HSSFRow rowDataBody = (HSSFRow)sheet2.CreateRow(_theRowNO);
                        //    for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                        //    {
                        //        if (j != _prodPosition && j != _prodPosition + 1)
                        //        {
                        //            HSSFCell cell = (HSSFCell)rowDataBody.CreateCell(j);
                        //            string _celldata = _ds.Tables[t].Rows[i][j].ToString();
                        //            cell.SetCellValue(_celldata);
                        //            cell.CellStyle = styleB;
                        //        }
                        //    }

                        //    foreach (KeyValuePair<string, int> kvp in _dicPdQty)
                        //    {
                        //        if (_theRowNO != _theVeriablRowcount + 1)
                        //        {
                        //            rowDataBody = (HSSFRow)sheet2.CreateRow(_theRowNO);
                        //        }
                        //        HSSFCell cellDataProd = (HSSFCell)rowDataBody.CreateCell(_prodPosition);
                        //        cellDataProd.SetCellValue(kvp.Key);

                        //        HSSFCell cellDataQty = (HSSFCell)rowDataBody.CreateCell(_prodPosition + 1);
                        //        cellDataQty.SetCellValue(kvp.Value);

                        //        _theRowNO++;

                        //    }

                        //    for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                        //    {
                        //        if (j != _prodPosition && j != _prodPosition + 1)
                        //        {
                        //            CellRangeAddress celrang = new CellRangeAddress(_theRowNO - 2, _theRowNO - 1, j, j);
                        //            sheet2.AddMergedRegion(celrang);

                        //            //sheet2.AutoSizeColumn(j);
                        //        }
                        //    }

                        //    _theVeriablRowcount += _dicPdQty.Count;

                        //    #endregion
                        //}
                        //else
                        //{

                        IRow row = (IRow)sheet2.CreateRow(_theVeriablRowcount + 1);
                        for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                        {
                            bool celladd = true;
                            ICell cell = (ICell)row.CreateCell(j);

                            foreach (var item in dis)
                            {
                                int[] k = item.Value;
                                if (k.Contains(j))
                                {
                                    if (item.Key == 0)
                                    {

                                        if (_ds.Tables[t].Rows[i][j].ToString() != "")
                                        {
                                            DateTime _celldata = Convert.ToDateTime(_ds.Tables[t].Rows[i][j].ToString());

                                            cell.CellStyle = cellStyle2;

                                            cell.SetCellValue(_celldata);

                                        }

                                        else
                                        {
                                            string _celldata = _ds.Tables[t].Rows[i][j].ToString();
                                            cell.CellStyle = styleH;
                                            cell.SetCellValue(_celldata);

                                        }
                                        celladd = false;
                                        break;


                                    }
                                    else if (item.Key == 1)
                                    {
                                        int _celldata = 0;
                                        if (_ds.Tables[t].Rows[i][j].ToString() != "")
                                        {
                                            _celldata = int.Parse(_ds.Tables[t].Rows[i][j].ToString());
                                        }


                                        cell.CellStyle = cellStyle3;
                                        cell.SetCellValue(_celldata);
                                        celladd = false;
                                        break;

                                    }
                                    else if (item.Key == 2)
                                    {
                                        double _celldata = 0;
                                        if (_ds.Tables[t].Rows[i][j].ToString() != "")
                                        {
                                            _celldata = double.Parse(_ds.Tables[t].Rows[i][j].ToString());
                                        }

                                        cell.CellStyle = cellStyle4;


                                        cell.SetCellValue(_celldata);
                                        celladd = false;
                                        break;
                                    }
                                    else if (item.Key == 3)
                                    {

                                        double _celldata = 0;
                                        if (_ds.Tables[t].Rows[i][j].ToString() != "")
                                        {
                                            _celldata = double.Parse(_ds.Tables[t].Rows[i][j].ToString());
                                        }

                                        cell.CellStyle = cellStyle5;

                                        cell.SetCellValue(_celldata);
                                        celladd = false;
                                        break;
                                    }

                                }

                            }

                            if (celladd)
                            {
                                string _celldata = _ds.Tables[t].Rows[i][j].ToString();
                                cell.CellStyle = styleH;
                                cell.SetCellValue(_celldata);
                            }

                            //if (_celldata.Trim().Contains("<br>"))
                            //{
                            //    cell.SetCellValue(_celldata.Trim().Replace("<br>", "\r\n"));
                            //    cell.CellStyle.WrapText = true;
                            //}
                            //else
                            //{
                            //    if (j == _prodPosition + 1)
                            //    {
                            //        int _testInt = 0;
                            //        if (int.TryParse(_celldata, out _testInt))
                            //        {
                            //            cell.SetCellValue(_testInt);
                            //        }
                            //    }
                            //}
                            cell.CellStyle.Alignment = HorizontalAlignment.CENTER;
                            cell.CellStyle.VerticalAlignment = VerticalAlignment.CENTER;

                        }

                        _theVeriablRowcount++;
                        //}
                    }
                }

                #endregion

                //hw.Write(System.IO.Stream(@"c:\poi0.xls"));
                FileStream file = new FileStream(_Path, FileMode.Create);
                hw.Write(file);
                file.Close();

                #endregion

                //sw.WriteLine("生产的excel 建立完成，再次检查文件是否成功建立");
                #region check if the excel file is created
                if (!File.Exists(_Path))
                {
                    //sw.WriteLine("生产的excel 未建立成功");
                    return false;
                }
                #endregion


                return true;
            }
            catch (Exception ex)
            {
                //string _Result = "一个处理了[" + _OutSheetCount + "]个, All = " + _SheetNameStr;

                throw new Exception(ex.Message, ex);
            }
        }


        public static bool DataSetToExcelXSSF(DataSet _ds, string _Path)
        {
            try
            {
                #region NPOI 导出方式
                //HSSFWorkbook hw = new HSSFWorkbook();
                XSSFWorkbook hw = new XSSFWorkbook();

                #region output every sheet

                for (int t = 0; t < _ds.Tables.Count; t++)
                {
                    ISheet sheet2 = (ISheet)hw.CreateSheet(_ds.Tables[t].TableName);

                    IRow rowCol2 = (IRow)sheet2.CreateRow(0);
                    for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                    {
                        ICell cell = (ICell)rowCol2.CreateCell(j);
                        //cell.SetCellValue("1000000000000000000000000000000000000000000");
                        cell.SetCellValue(_ds.Tables[t].Columns[j].ColumnName);
                    }

                    for (int i = 0; i < _ds.Tables[t].Rows.Count; i++)
                    {
                        IRow row = (IRow)sheet2.CreateRow(i + 1);
                        for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                        {
                            ICell cell = (ICell)row.CreateCell(j);
                            if (_ds.Tables[t].Rows[i][j].ToString().Contains("<br>"))
                            {
                                ICellStyle cs = hw.CreateCellStyle();
                                cs.WrapText = true;
                                cell.CellStyle = cs;
                                //sheet2.GetRow(i+1).HeightInPoints = 2 * sheet2.DefaultRowHeight / 20;

                            }
                            //cell.SetCellValue("1000000000000000000000000000000000000000000");
                            cell.SetCellValue(_ds.Tables[t].Rows[i][j].ToString().Replace("<br>", "\r\n"));
                        }
                    }
                }


                #endregion

                FileStream file = new FileStream(_Path, FileMode.Create);
                hw.Write(file);
                file.Close();

                //hw.Write(System.IO.Stream(@"c:\poi0.xls"));
                //FileStream fs = new FileStream(_Path, FileMode.Create, FileAccess.Write);
                //long fileSize = fs.Length;


                //HttpContext.Current.Response.AddHeader("Content-Length", fileSize.ToString());

                //byte[] fileBuffer = new byte[fileSize];
                //fs.Read(fileBuffer, 0, (int)fileSize);
                //HttpContext.Current.Response.BinaryWrite(fileBuffer);
                //hw.Write(fs);
                //fs.Close();

                #endregion

                //sw.WriteLine("生产的excel 建立完成，再次检查文件是否成功建立");
                #region check if the excel file is created
                if (!File.Exists(_Path))
                {
                    //sw.WriteLine("生产的excel 未建立成功");
                    return false;
                }
                #endregion


                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        public static bool DataSetToExcelXSSF(DataSet _ds, string _Path, string _HoriAlignment, Dictionary<int, int> _dicColWidth)
        {
            try
            {
                #region NPOI 导出方式
                //HSSFWorkbook hw = new HSSFWorkbook();
                XSSFWorkbook hw = new XSSFWorkbook();

                #region style Head and Body  报表 Footer 数据，正常字体，加粗，加边框 右对齐
                IFont fontHead = hw.CreateFont();
                fontHead.Color = HSSFColor.BLACK.index;
                fontHead.Boldweight = (short)FontBoldWeight.BOLD;

                ICellStyle styleHead = hw.CreateCellStyle();
                styleHead.SetFont(fontHead);

                if (_HoriAlignment.Trim().ToLower() == "center")
                {
                    styleHead.Alignment = HorizontalAlignment.CENTER;
                }
                else if (_HoriAlignment.Trim().ToLower() == "left")
                {
                    styleHead.Alignment = HorizontalAlignment.LEFT;
                }
                else if (_HoriAlignment.Trim().ToLower() == "right")
                {
                    styleHead.Alignment = HorizontalAlignment.RIGHT;
                }

                styleHead.VerticalAlignment = VerticalAlignment.CENTER;

                styleHead.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                styleHead.BottomBorderColor = HSSFColor.BLACK.index;
                styleHead.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
                styleHead.LeftBorderColor = HSSFColor.BLACK.index;
                styleHead.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
                styleHead.RightBorderColor = HSSFColor.BLACK.index;
                styleHead.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
                styleHead.TopBorderColor = HSSFColor.BLACK.index;



                IFont fontBody = hw.CreateFont();
                fontBody.Color = HSSFColor.BLACK.index;
                //font6.Boldweight = (short)FontBoldWeight.BOLD;

                ICellStyle styleBody = hw.CreateCellStyle();
                styleBody.SetFont(fontBody);

                if (_HoriAlignment.Trim().ToLower() == "center")
                {
                    styleBody.Alignment = HorizontalAlignment.CENTER;
                }
                else if (_HoriAlignment.Trim().ToLower() == "left")
                {
                    styleBody.Alignment = HorizontalAlignment.LEFT;
                }
                else if (_HoriAlignment.Trim().ToLower() == "right")
                {
                    styleBody.Alignment = HorizontalAlignment.RIGHT;
                }

                styleBody.VerticalAlignment = VerticalAlignment.CENTER;

                styleBody.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                styleBody.BottomBorderColor = HSSFColor.BLACK.index;
                styleBody.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
                styleBody.LeftBorderColor = HSSFColor.BLACK.index;
                styleBody.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
                styleBody.RightBorderColor = HSSFColor.BLACK.index;
                styleBody.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
                styleBody.TopBorderColor = HSSFColor.BLACK.index;

                #endregion

                #region output every sheet

                for (int t = 0; t < _ds.Tables.Count; t++)
                {
                    ISheet sheet2 = (ISheet)hw.CreateSheet(_ds.Tables[t].TableName);

                    IRow rowCol2 = (IRow)sheet2.CreateRow(0);
                    for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                    {
                        ICell cell = (ICell)rowCol2.CreateCell(j);
                        //cell.SetCellValue("1000000000000000000000000000000000000000000");
                        cell.SetCellValue(_ds.Tables[t].Columns[j].ColumnName);
                        cell.CellStyle = styleHead;
                    }

                    for (int i = 0; i < _ds.Tables[t].Rows.Count; i++)
                    {
                        IRow row = (IRow)sheet2.CreateRow(i + 1);
                        for (int j = 0; j < _ds.Tables[t].Columns.Count; j++)
                        {
                            ICell cell = (ICell)row.CreateCell(j);
                            cell.CellStyle = styleBody;
                            if (_ds.Tables[t].Rows[i][j].ToString().Contains("<br>"))
                            {
                                //ICellStyle cs = hw.CreateCellStyle();
                                cell.CellStyle.WrapText = true;

                                //sheet2.GetRow(i+1).HeightInPoints = 2 * sheet2.DefaultRowHeight / 20;
                            }
                            //cell.SetCellValue("1000000000000000000000000000000000000000000");
                            cell.SetCellValue(_ds.Tables[t].Rows[i][j].ToString().Replace("<br>", "\r\n"));

                            //int _DataLen = _ds.Tables[t].Rows[i][j].ToString().Length;
                            //int _colWidth = 3000;

                            //if (_DataLen > 20)
                            //{
                            //    _colWidth = _colWidth * _DataLen / 10;
                            //}
                            int _defaultWidth = 2000;
                            if (_dicColWidth.ContainsKey(j))
                            {
                                _defaultWidth = _dicColWidth[j];
                            }

                            sheet2.SetColumnWidth(j, _defaultWidth);
                        }
                    }
                }


                #endregion

                FileStream file = new FileStream(_Path, FileMode.Create);
                hw.Write(file);
                file.Close();

                //hw.Write(System.IO.Stream(@"c:\poi0.xls"));
                //FileStream fs = new FileStream(_Path, FileMode.Create, FileAccess.Write);
                //long fileSize = fs.Length;


                //HttpContext.Current.Response.AddHeader("Content-Length", fileSize.ToString());

                //byte[] fileBuffer = new byte[fileSize];
                //fs.Read(fileBuffer, 0, (int)fileSize);
                //HttpContext.Current.Response.BinaryWrite(fileBuffer);
                //hw.Write(fs);
                //fs.Close();

                #endregion

                //sw.WriteLine("生产的excel 建立完成，再次检查文件是否成功建立");
                #region check if the excel file is created
                if (!File.Exists(_Path))
                {
                    //sw.WriteLine("生产的excel 未建立成功");
                    return false;
                }
                #endregion


                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        public static DataTable MSTranslationExcelRead(string path)
        {
            try
            {
                int cellCount = 40;
                //DataSet _ds = new DataSet();
                HSSFWorkbook hssfworkbook;
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new HSSFWorkbook(file);
                }

                #region 判断 Excel 中的 Sheet
                if (hssfworkbook.NumberOfSheets <= 0)
                {
                    throw new Exception("There is no sheet in the Excel.(Excel中无Sheet资料。)");
                }
                //else if (hssfworkbook.NumberOfSheets > 1)
                //{
                //    throw new Exception("Uploaded Excel file contains multiple sheets. Cannot not upload.(上传的Excel存在多个Sheet，不允许上传！)");
                //}
                #endregion

                int _WarningSheet = 7;

                HSSFSheet sheet = (HSSFSheet)hssfworkbook.GetSheetAt(_WarningSheet);

                DataTable table = new DataTable();

                #region

                for (int i = 1; i < cellCount + 1; i++)
                {
                    DataColumn column = new DataColumn("F" + i);
                    table.Columns.Add(column);
                }

                for (int i = (sheet.FirstRowNum); i < sheet.LastRowNum + 1; i++)
                {
                    HSSFRow row = (HSSFRow)sheet.GetRow(i);
                    DataRow dataRow = table.NewRow();

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            dataRow[j] = row.GetCell(j).ToString();
                    }

                    table.Rows.Add(dataRow);
                }
                //_ds.Tables.Add(table);
                #endregion


                //ExcelFileStream.Close();
                hssfworkbook = null;
                sheet = null;



                DataTable t = new DataTable();
                for (int i = 0; i < cellCount; i++)
                {
                    t.Columns.Add(i.ToString());
                }
                foreach (DataRow item in table.Rows)
                {
                    DataRow row = t.NewRow();
                    for (int i = 0; i < cellCount; i++)
                    {
                        row[i] = item[i];
                    }
                    t.Rows.Add(row);
                }
                return t;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <param name="filePath">路径</param>
        /// <param name="startRow">每个表格从多少行开始 格式：["",""]</param>
        /// <returns></returns>
        public static DataSet ReadExcel(string filePath)
        {
            return ReadExcel(filePath, null);
        }

        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <param name="filePath">路径</param>
        /// <param name="startRow">每个表格从多少行开始 格式：["",""]</param>
        /// <returns></returns>
        public static DataSet ReadExcel(string filePath, string[] startRow)
        {
            IWorkbook book = null;
            DataSet ds = new DataSet();
            string excelType = "";
            string str = filePath.Substring(filePath.LastIndexOf('.') + 1);
            if (str == "xlsx")
                excelType = "xlsx";
            else if (str == "xls")
                excelType = "xls";
            else
                throw new Exception("Excel格式不正确");
            try
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (excelType == "xlsx")
                        book = new XSSFWorkbook(file);
                    else
                        book = new HSSFWorkbook(file);

                    for (int a = 0; a < book.NumberOfSheets; a++)
                    {
                        ISheet sheet = (ISheet)book.GetSheetAt(a);
                        int sheetRowsCount = sheet.LastRowNum;

                        int row = 0;
                        if (startRow != null)
                        {
                            if (startRow.Length >= a)
                                row = int.Parse(startRow[a]);
                        }

                        IRow iRow = (IRow)sheet.GetRow(row);
                        DataTable dt = new DataTable();

                        if (sheet.GetRow(row) == null)
                            return ds;
                        for (int i = 0; i < (sheet.GetRow(row).LastCellNum); i++)
                        {
                            object obj = GetValueType(iRow.GetCell(i) as ICell);
                            if (obj == null || obj.ToString() == string.Empty)
                                dt.Columns.Add(new DataColumn("Columns" + i.ToString().Trim()));
                            else
                                dt.Columns.Add(new DataColumn(obj.ToString().Trim()));
                        }
                        for (int i = row + 1; i <= sheetRowsCount; i++)
                        {
                            DataRow dr = dt.NewRow();
                            IRow ro = (IRow)sheet.GetRow(i);
                            for (int j = 0; j < ro.LastCellNum; j++)
                            {
                                dr[j] = GetValueType(sheet.GetRow(i).GetCell(j) as ICell);
                            }
                            dt.Rows.Add(dr);
                        }
                        dt.TableName = sheet.SheetName;
                        ds.Tables.Add(dt);
                    }
                    return ds;
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        /// <summary>  
        /// 获取单元格类型(xls)  
        /// </summary>  
        /// <param name="cell"></param>  
        /// <returns></returns>  
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.BLANK: //BLANK:  
                    return null;
                case CellType.BOOLEAN: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.NUMERIC: //NUMERIC:
                    short format = cell.CellStyle.DataFormat;
                    if (DateUtil.IsCellDateFormatted(cell) || format == 14 || format == 31 || format == 57 || format == 58)
                    {
                        if (format == 180 || format == 178)
                        {
                            return cell.DateCellValue.ToString("yyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            return cell.DateCellValue.ToString("yyy-MM-dd");
                        }
                    }
                    else
                    {
                        return cell.ToString(); //cell.NumericCellValue;//此处不使用返回 NumericCellValue 的原因为：当内容为100% 或60.00时会返回1 和 60.0
                    }
                case CellType.STRING: //STRING:  
                    return cell.StringCellValue;
                case CellType.ERROR: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.FORMULA: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }

        /// <summary>
        /// 用 NPOI 方式取得 Excel 中的数据。读出的列栏位以F1，F2，F3.....为列名。
        /// IsSheetIndex 为true 表示要按照参数SheetIndex来读指定的sheet；为false则直接按照sheet名（SheetName）来读取。
        /// </summary>
        public static bool ReadExcelAsTableNOPI(string fileName, int SheetIndex, string SheetName, ref DataSet DS, bool IsSheetIndex, ref string ErrDsc)
        {
            ErrDsc = string.Empty;
            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    //HSSFWorkbook wb = new HSSFWorkbook(fs);
                    IWorkbook wb;// = new HSSFWorkbook(fs);

                    if (fileName.Trim().ToLower().EndsWith(".xls"))
                    {
                        wb = new HSSFWorkbook(fs);
                    }
                    else if (fileName.Trim().ToLower().EndsWith(".xlsx"))
                    {
                        wb = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                    }
                    else
                    {
                        throw new Exception("file is not a excel file.");
                        //return null;
                    }

                    ISheet sheet;
                    if (IsSheetIndex)
                    {
                        sheet = wb.GetSheetAt(SheetIndex);
                    }
                    else
                    {
                        bool _isFindSheetName = false;
                        for (int i = 0; i < wb.NumberOfSheets; i++)
                        {
                            if (wb.GetSheetName(i).Trim().ToLower() == SheetName.Trim().ToLower())
                            {
                                _isFindSheetName = true;
                                break;
                            }
                        }

                        if (!_isFindSheetName)
                        {
                            throw new Exception("The Excel does not has the sheet name [" + SheetName + "] (Excel中没有名为 [" + SheetName + "] 的Sheet).");
                        }

                        sheet = wb.GetSheet(SheetName);
                    }
                    DataTable table = new DataTable();
                    DS.Tables.Add(table);
                    //由第一列取標題做為欄位名稱          
                    // IRow headerRow = sheet.GetRow(0);
                    // int cellCount = 36;// headerRow.LastCellNum;

                    int maxcount = -1;
                    for (int i = (sheet.FirstRowNum); i < sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        int count = row.LastCellNum;
                        if (count > maxcount)
                        {
                            maxcount = count;
                        }
                    }

                    if (maxcount <= 0) throw new Exception("Excel no Data for maxcount");
                    for (int d = 0; d < maxcount; d++)                      //以F1字為名新增欄位，此處全視為字串型別以求簡化      
                    {
                        table.Columns.Add(new DataColumn("F" + (d + 1)));
                    }

                    //略過第零列(標題列)，一直處理至最後一列         
                    for (int i = (sheet.FirstRowNum); i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        DataRow dataRow = table.NewRow();
                        //依先前取得的欄位數逐一設定欄位內容    
                        // if (row.LastCellNum > cellCount) throw new Exception("LastCellNum < cellCount  LastCellNum=" + row.LastCellNum);

                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                //如要針對不同型別做個別處理，可善用.CellType判斷型別        
                                //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值      
                                //此處只簡單轉成字串                    

                                string s = row.GetCell(j).ToString();
                                dataRow[j] = row.GetCell(j).ToString();
                            }
                        }
                        table.Rows.Add(dataRow);
                    }

                    if (table.Rows.Count <= 0)
                        throw new Exception("Excel has No Data");

                    return true;
                }
            }
            catch (Exception ex)
            {
                ErrDsc = ex.Message;

                return false;
            }
        }

        /// <summary>
        /// 用 NPOI 方式取得 Excel 中的数据。读出的列栏位以F1，F2，F3.....为列名。
        /// IsSheetIndex 为true 表示要按照参数SheetIndex来读指定的sheet；为false则直接按照sheet名（SheetName）来读取。
        /// 加了一个参数据：_ColumnCount 代表读出几列，为0，则不生效。
        /// 此方法已改成根据单元格类型来读出资料。
        /// </summary>
        public static bool ReadExcelAsTableNOPI(string fileName, int SheetIndex, string SheetName, ref DataSet DS, bool IsSheetIndex, int _ColumnCount, ref string ErrDsc)
        {
            ErrDsc = string.Empty;
            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    //HSSFWorkbook wb = new HSSFWorkbook(fs);
                    IWorkbook wb;// = new HSSFWorkbook(fs);

                    if (fileName.Trim().ToLower().EndsWith(".xls"))
                    {
                        wb = new HSSFWorkbook(fs);
                    }
                    else if (fileName.Trim().ToLower().EndsWith(".xlsx"))
                    {
                        wb = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                    }
                    else
                    {
                        throw new Exception("file is not a excel file.");
                        //return null;
                    }

                    ISheet sheet;
                    #region IsSheetIndex 值来决定是由 Sheet 序号来读

                    if (IsSheetIndex)
                    {
                        sheet = wb.GetSheetAt(SheetIndex);
                    }
                    else
                    {
                        bool _isFindSheetName = false;
                        for (int i = 0; i < wb.NumberOfSheets; i++)
                        {
                            if (wb.GetSheetName(i).Trim().ToLower() == SheetName.Trim().ToLower())
                            {
                                _isFindSheetName = true;
                                break;
                            }
                        }

                        if (!_isFindSheetName)
                        {
                            throw new Exception("The Excel does not has the sheet name [" + SheetName + "] (Excel中没有名为 [" + SheetName + "] 的Sheet).");
                        }

                        sheet = wb.GetSheet(SheetName);
                    }
                    #endregion

                    DataTable table = new DataTable();
                    DS.Tables.Add(table);
                    //由第一列取標題做為欄位名稱          
                    // IRow headerRow = sheet.GetRow(0);
                    // int cellCount = 36;// headerRow.LastCellNum;

                    #region 由 _ColumnCount 决定读出指定列数，还是根据Excel内容来读

                    int maxcount = -1;
                    for (int i = (sheet.FirstRowNum); i < sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        int count = row.LastCellNum;
                        if (count > maxcount)
                        {
                            maxcount = count;
                        }
                    }

                    if (_ColumnCount > 0)
                    {
                        maxcount = _ColumnCount;
                    }

                    if (maxcount <= 0) throw new Exception("Excel no Data for maxcount");
                    for (int d = 0; d < maxcount; d++)                      //以F1字為名新增欄位，此處全視為字串型別以求簡化      
                    {
                        table.Columns.Add(new DataColumn("F" + (d + 1)));
                    }

                    #endregion

                    //略過第零列(標題列)，一直處理至最後一列         
                    for (int i = (sheet.FirstRowNum); i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        DataRow dataRow = table.NewRow();
                        //依先前取得的欄位數逐一設定欄位內容    
                        // if (row.LastCellNum > cellCount) throw new Exception("LastCellNum < cellCount  LastCellNum=" + row.LastCellNum);

                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                //如要針對不同型別做個別處理，可善用.CellType判斷型別        
                                //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值      
                                //此處只簡單轉成字串                    

                                //string s = row.GetCell(j).ToString();
                                dataRow[j] = GetValueType(row.GetCell(j) as ICell);
                            }
                        }
                        table.Rows.Add(dataRow);
                    }

                    if (table.Rows.Count <= 0)
                        throw new Exception("Excel has No Data");

                    return true;
                }
            }
            catch (Exception ex)
            {
                ErrDsc = ex.Message;

                return false;
            }
        }

        public static bool ReadExcelAsTableNOPIforNLBarc(string fileName, int SheetIndex, string SheetName, ref DataSet DS, bool IsSheetIndex, ref string ErrDsc)
        {
            ErrDsc = string.Empty;
            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    //HSSFWorkbook wb = new HSSFWorkbook(fs);
                    IWorkbook wb;// = new HSSFWorkbook(fs);

                    if (fileName.Trim().ToLower().EndsWith(".xls"))
                    {
                        wb = new HSSFWorkbook(fs);
                    }
                    else if (fileName.Trim().ToLower().EndsWith(".xlsx"))
                    {
                        wb = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                    }
                    else
                    {
                        throw new Exception("file is not a excel file.");
                        //return null;
                    }

                    ISheet sheet;
                    if (IsSheetIndex)
                    {
                        sheet = wb.GetSheetAt(SheetIndex);
                    }
                    else
                    {
                        bool _isFindSheetName = false;
                        for (int i = 0; i < wb.NumberOfSheets; i++)
                        {
                            if (wb.GetSheetName(i).Trim().ToLower() == SheetName.Trim().ToLower())
                            {
                                _isFindSheetName = true;
                                break;
                            }
                        }

                        if (!_isFindSheetName)
                        {
                            throw new Exception("The Excel does not has the sheet name [" + SheetName + "] (Excel中没有名为 [" + SheetName + "] 的Sheet).");
                        }

                        sheet = wb.GetSheet(SheetName);
                    }
                    DataTable table = new DataTable();
                    DS.Tables.Add(table);
                    //由第一列取標題做為欄位名稱          
                    // IRow headerRow = sheet.GetRow(0);
                    // int cellCount = 36;// headerRow.LastCellNum;

                    int maxcount = -1;

                    if (sheet.FirstRowNum == sheet.LastRowNum)
                    {
                        maxcount = sheet.GetRow(0).LastCellNum;
                    }
                    else
                    {
                        for (int i = (sheet.FirstRowNum); i < sheet.LastRowNum; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null) continue;

                            int count = row.LastCellNum;
                            if (count > maxcount)
                            {
                                maxcount = count;
                            }
                        }
                    }

                    if (maxcount <= 0) throw new Exception("Excel no Data for maxcount");
                    for (int d = 0; d < maxcount; d++)                      //以F1字為名新增欄位，此處全視為字串型別以求簡化      
                    {
                        table.Columns.Add(new DataColumn("F" + (d + 1)));
                    }

                    //略過第零列(標題列)，一直處理至最後一列         
                    for (int i = (sheet.FirstRowNum); i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        DataRow dataRow = table.NewRow();
                        //依先前取得的欄位數逐一設定欄位內容    
                        // if (row.LastCellNum > cellCount) throw new Exception("LastCellNum < cellCount  LastCellNum=" + row.LastCellNum);

                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                //如要針對不同型別做個別處理，可善用.CellType判斷型別        
                                //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值      
                                //此處只簡單轉成字串                    

                                string s = row.GetCell(j).ToString();
                                dataRow[j] = row.GetCell(j).ToString();
                            }
                        }
                        table.Rows.Add(dataRow);
                    }

                    if (table.Rows.Count <= 0)
                        throw new Exception("Excel has No Data");

                    return true;
                }
            }
            catch (Exception ex)
            {
                ErrDsc = ex.Message;

                return false;
            }
        }

        /// <summary>
        /// 用 NPOI 方式取得 Excel 中的数据。以Excel中第一行为列名。
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static DataTable ReadExcelAsTableNOPIUsingFirstRow(string fileName, int SheetIndex, string SheetName, bool IsSheetIndex)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open))
            {
                //HSSFWorkbook wb = new HSSFWorkbook(fs);
                IWorkbook wb;// = new HSSFWorkbook(fs);

                if (fileName.Trim().ToLower().EndsWith(".xls"))
                {
                    wb = new HSSFWorkbook(fs);
                }
                else if (fileName.Trim().ToLower().EndsWith(".xlsx"))
                {
                    wb = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                }
                else
                {
                    throw new Exception("file is not a excel file.");
                    //return null;
                }

                ISheet sheet;
                if (IsSheetIndex)
                {
                    sheet = wb.GetSheetAt(SheetIndex);
                }
                else
                {
                    sheet = wb.GetSheet(SheetName);
                }

                DataTable table = new DataTable();
                //由第一列取標題做為欄位名稱          
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)                      //以欄位文字為名新增欄位，此處全視為字串型別以求簡化      
                    table.Columns.Add(
                        new DataColumn(headerRow.GetCell(i).StringCellValue));
                //略過第零列(標題列)，一直處理至最後一列         
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    DataRow dataRow = table.NewRow();
                    //依先前取得的欄位數逐一設定欄位內容           
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                        if (row.GetCell(j) != null)
                            //如要針對不同型別做個別處理，可善用.CellType判斷型別        
                            //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值      
                            //此處只簡單轉成字串                    
                            dataRow[j] = row.GetCell(j).ToString();
                    table.Rows.Add(dataRow);
                }
                return table;
            }
        }
    }
}
