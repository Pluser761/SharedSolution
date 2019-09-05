using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExelWork
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, string sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
            
        }

        public string ReadCell(int line, string column)
        {
            int col = columnToInt(column);

            if (Convert.ToString(ws.Cells[line, col].Value2) != "")
                return Convert.ToString(ws.Cells[line, col].Value2);
            else
                return "";
        }

        public string ReadCell(int line, int col)
        {
            if (Convert.ToString(ws.Cells[line, col].Value2) != "")
                return Convert.ToString(ws.Cells[line, col].Value2);
            else
                return "";
        }

        public void WriteCell(string str, int line, string column)
        {
            ws.Cells[line, columnToInt(column)].Value2 = str;
        }
        public void WriteCell(string str, int line, int column)
        {
            ws.Cells[line, column].Value2 = str;
        }

        public void Save()
        {
            wb.Save();
        }

        public void Close()
        {
            wb.Close();
            excel.Quit();
        }

        public int columnToInt(string s)
        {
            int col = 0;
            for (int i = 0; i < s.Length; i++)
            {
                col *= 10;
                col += (int)(s[i]) - 64;
            }
            return col;
        }
    }

    
}