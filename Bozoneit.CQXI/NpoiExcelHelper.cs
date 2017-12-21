using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace Bozoneit.CQXI
{
    public class NpoiExcelHelper
    {
        /// <summary>
        /// 读取excel转为DataTable
        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="sheetName">指定sheet</param>
        /// <param name="isColumnName">第一行是否为列名</param>
        /// <param name="startRow">从第几行开始</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string fileName, ref string sheetName, bool isColumnName, int startRow = 0)
        {
            IWorkbook workBook = null;
            ISheet sheet = null;
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                //低于2007版本
                if (Path.GetExtension(fileName) == ".xls")
                {
                    workBook = new HSSFWorkbook(fs);
                }
                //2007及以上版本
                else if (Path.GetExtension(fileName) == ".xlsx")
                {
                    workBook = new XSSFWorkbook(fs);
                }
            }
            //判断是否指定sheet上传
            if (sheetName != null)
            {
                //获取指定sheet
                sheet = workBook.GetSheet(sheetName);
                if (sheet == null)
                {
                    //获取不到时取第一个sheet
                    sheet = workBook.GetSheetAt(0);

                }
            }
            else
            {
                sheet = workBook.GetSheetAt(0);
            }
            if (sheet != null)
            {
                sheetName = workBook.GetSheetName(0);
                //sheet中第一行
                IRow firstRow = sheet.GetRow(0);
                if (firstRow == null)
                {
                    throw new Exception("首行无数据");
                }

                //遍历第一行的单元格
                for (int i = firstRow.FirstCellNum; i < firstRow.LastCellNum; i++)
                {
                    //得到列名
                    ICell cell = firstRow.GetCell(i);
                    if (cell != null)
                    {
                        //得到列名的值,若列名不是字符则不能使用StringCellValue，最好使用ToString()
                        string cellValue = cell.ToString();
                        if (cellValue != null)
                        {
                            try
                            {
                                //判断第一行是否是列名
                                if (isColumnName)
                                {
                                    //将列放入datatable中
                                    DataColumn column = new DataColumn(cellValue);
                                    dt.Columns.Add(column);
                                }
                                else
                                {
                                    //将空列放入datatable中
                                    DataColumn column = new DataColumn();
                                    dt.Columns.Add(column);
                                }
                            }
                            catch
                            {
                                throw new Exception("列名有误！");
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                }
                //遍历所有行
                for (int i = startRow; i <= sheet.LastRowNum; i++)
                {
                    //得到i行
                    IRow row = sheet.GetRow(i);
                    if (row == null)
                    {
                        continue;
                    }
                    //datatable新增行
                    DataRow dr = dt.NewRow();
                    //遍历i行的单元格
                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            dr[j] = row.GetCell(j).ToString();
                        }
                    }
                    try
                    {
                        //将行放入datatable中
                        dt.Rows.Add(dr);
                    }
                    catch
                    {
                        throw new Exception("第" + i + "行有误！");
                    }
                }
            }
            return dt;
        }
        /// <summary>
        /// DataTable导出到Excel
        /// </summary>
        /// <param name="fileName">导出文件的路径</param>
        /// <param name="templetName">导出模板路径</param>
        /// <param name="dt">DataTable</param>
        /// <param name="titleName">文件标题</param>
        /// <param name="sheetName">文件sheet名称</param>
        public static void DataTableToExcel(string fileName, string templetName, DataTable dt, string titleName, string sheetName)
        {
            FileStream fs1 = new FileStream(templetName, FileMode.Open, FileAccess.Read);
            IWorkbook workBook = new XSSFWorkbook(fs1);
            ISheet sheet = workBook.GetSheet(sheetName);

            

            ////第一行
            //IRow row0 = sheet.GetRow(0);
            //ICell cellTitle = row0.GetCell(0);
            //cellTitle.SetCellValue(titleName);
            //第二行
            IRow row1 = sheet.GetRow(0);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = row1.GetCell(j);
                cell.SetCellValue(dt.Columns[j].ColumnName);
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow rowi = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //创建单元格
                    ICell cell = rowi.CreateCell(j);
                    //给单元格赋值
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                    cell.CellStyle.BorderBottom = BorderStyle.Thin;
                    cell.CellStyle.BorderLeft = BorderStyle.Thin;
                    cell.CellStyle.BorderDiagonalColor = IndexedColors.Red.Index;

                    cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                    cell.CellStyle.FillForegroundColor = HSSFColor.Yellow.Index;

                }
            }
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workBook.Write(fs);
            }
        }

        public static void ListToExcel(string fileName, string templetName, List<RecordMode> dt, string titleName, string sheetName)
        {
            FileStream fs1 = new FileStream(templetName, FileMode.Open, FileAccess.Read);
            IWorkbook workBook = new XSSFWorkbook(fs1);
            ISheet sheet = workBook.GetSheet(sheetName);

            List<Color> colorList = new List<Color>();
            
            colorList.Add(Color.FromArgb(255, 0, 0));//红色
            colorList.Add(Color.FromArgb(255, 255, 255));//白色

            HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
            HSSFPalette palette = hssfWorkbook.GetCustomPalette();
            palette.SetColorAtIndex(999, colorList[0].R, colorList[0].G, colorList[0].B);
            palette.SetColorAtIndex(998, colorList[1].R, colorList[1].G, colorList[1].B);

            ////第一行
            //IRow row0 = sheet.GetRow(0);
            //ICell cellTitle = row0.GetCell(0);
            //cellTitle.SetCellValue(titleName);
            //第二行
            IRow row1 = sheet.GetRow(0);
            for (int j = 0; j < 8; j++)
            {
                ICell cell = row1.GetCell(j);
                cell.SetCellValue(cell.ToString());
            }
            for (int i = 0; i < dt.Count; i++)
            {
                IRow rowi = sheet.CreateRow(i + 1);
                for (int j = 0; j < 8; j++)
                {
                    //创建单元格
                    ICell cell = rowi.CreateCell(j);
                    //给单元格赋值
                    cell.SetCellValue(dt[i][j].ToString());

                    //给单元格设置样式
                    ICellStyle colorStyle = workBook.CreateCellStyle();
                    colorStyle.FillPattern = FillPattern.SolidForeground;


                    //cell.CellStyle.BorderBottom = BorderStyle.Thin;
                    //cell.CellStyle.BorderLeft = BorderStyle.Thin;
                    if (dt[i].isYC == true&&(j==4||j==5))
                    {
                        var v1 = palette.FindColor(colorList[0].R, colorList[0].G, colorList[0].B);
                        if (v1 == null)
                        {
                            throw new Exception("Color is not in Palette");
                        }
                        colorStyle.FillForegroundColor = v1.Indexed;
                        cell.CellStyle = colorStyle;
                        //cell.CellStyle.BorderDiagonalColor = IndexedColors.Yellow.Index;
                        //cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        //cell.CellStyle.FillForegroundColor = HSSFColor.Yellow.Index;
                    }
                    else
                    {
                        var v1 = palette.FindColor(colorList[1].R, colorList[1].G, colorList[1].B);
                        if (v1 == null)
                        {
                            throw new Exception("Color is not in Palette");
                        }
                        colorStyle.FillForegroundColor = v1.Indexed;
                        cell.CellStyle = colorStyle;
                        //cell.CellStyle.BorderDiagonalColor = IndexedColors.Yellow.Index;
                        //cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                        //cell.CellStyle.FillForegroundColor = HSSFColor.White.Index;
                    }

                }
            }
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workBook.Write(fs);
            }
        }
    }
}
