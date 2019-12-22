using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.SqlClient;

namespace Directory
{
    class ExcelControl
    {
        Excel.Application excelApp = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        #region 엑셀

        public void ConnectExcel()
        {
            try
            {
                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Open(@"C:\Users\DUMB\Desktop\새 Microsoft Excel 워크시트.xlsx");
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
            }
            catch(Exception ex)
            {
                MessageBox.Show("[에러] : " + ex.Message);
            }
        }

        public void ExcelSave()
        {
            try
            {
                wb.Save();
                ws.Delete();
                wb.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                wb.Close();
                excelApp.Quit();
            }

        }

        public void Exit()
        {
            try
            {
                DialogResult r= MessageBox.Show("종료 전 : 엑셀을 저장하시겠습니까?", "알림", MessageBoxButtons.OKCancel);
                if(r == DialogResult.OK)
                {
                    wb.Save();
                }
               wb.Close(Type.Missing, Type.Missing, Type.Missing); excelApp.Quit(); 
                releaseObject(excelApp); 
                releaseObject(ws); 
                releaseObject(wb);

                 }
            catch (Exception ex)
            {
             
            }
        }

          #region 메모리해제
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion


        public bool SaveExcelFile()
        {
            wb.Save();
            return true;
        }
        public void PrintAll(ListView listView2)
        {
            Microsoft.Office.Interop.Excel.Range userrange = ws.UsedRange;
            int row = userrange.Rows.Count;
            
            listView2.Items.Clear();

            object[] text = { "" };
            for (int r = 1; r <= row; r++)
            {
                try
                {

                    String[] aa = { ws.Cells[r, 1].Value2.ToString(), ws.Cells[r, 2].Value2.ToString(), ws.Cells[r, 3].Value2.ToString(), ws.Cells[r, 5].Value2.ToString(), ws.Cells[r, 6].Value2.ToString(), ws.Cells[r, 7].Value2.ToString(), ws.Cells[r, 8].Value2.ToString(), ws.Cells[r, 9].Value2.ToString(), ws.Cells[r, 10].Value2.ToString(), ws.Cells[r, 11].Value2.ToString(), ws.Cells[r, 12].Value2.ToString(), ws.Cells[r, 4].Value2.ToString() };
                    ListViewItem newitem = new ListViewItem(aa);
                    listView2.Items.Add(newitem);
                }
                catch (Exception ex)
                {
                  //  MessageBox.Show("[DB 리스트뷰 에러]" + ex.Message);
                }

            }
        }



        public void InsertMember(int a, string[] msg)
        {

            string str = Excel_NullToSom(msg[1], msg[2], msg[3], msg[4], msg[5], msg[6], msg[7], msg[8], msg[9], msg[10], msg[11]);
            string[] temp = str.Split('#');     
            ((Excel.Range)ws.Cells[a, 1]).Value =  msg[0];
            ((Excel.Range)ws.Cells[a, 2]).Value =  temp[0];
            ((Excel.Range)ws.Cells[a, 3]).Value =  temp[1];
            ((Excel.Range)ws.Cells[a, 4]).Value =  temp[2];
            ((Excel.Range)ws.Cells[a, 5]).Value =  temp[3];
            ((Excel.Range)ws.Cells[a, 6]).Value =  temp[4];
            ((Excel.Range)ws.Cells[a, 7]).Value =  temp[5];
            ((Excel.Range)ws.Cells[a, 8]).Value =  temp[6];
            ((Excel.Range)ws.Cells[a, 9]).Value =  temp[7];
            ((Excel.Range)ws.Cells[a, 10]).Value = temp[8];
            ((Excel.Range)ws.Cells[a, 11]).Value = temp[9];
            ((Excel.Range)ws.Cells[a, 12]).Value = temp[10];


        }

        public void UpdateExcelFile(int idx)
        {
            //idx에 있는 모든 정보를 불러온후
            //Excel상의 idx배열에 업데이트

            ((Excel.Range)ws.Cells[idx, 1]).Value = "저장";
            ((Excel.Range)ws.Cells[idx, 2]).Value = "저장";
            ((Excel.Range)ws.Cells[idx, 3]).Value = "저장";
            ((Excel.Range)ws.Cells[idx, 4]).Value = "저장";
            ((Excel.Range)ws.Cells[idx, 5]).Value = "저장";

        }

        public bool DeleteExcelFile(int idx)
        {
            Excel.Range startRange = ws.Cells[idx, 1] as Excel.Range;

            Excel.Range endRange = ws.Cells[idx, ws.Columns.Count] as Excel.Range;

            ws.Rows.Range[startRange, endRange].Delete();
            wb.Save();

            return true;
        }

        public void DeleteAll(int rows)
        {
            Excel.Range startRange = ws.Cells[1, 1] as Excel.Range;

            Excel.Range endRange = ws.Cells[rows, ws.Columns.Count] as Excel.Range;

            ws.Rows.Range[startRange, endRange].Delete();
            wb.Save();

        }

        public int FindLastIdx()
        {
            Microsoft.Office.Interop.Excel.Range userrange = ws.UsedRange;
            int row = userrange.Rows.Count;
            if (row == 0)
                return -1;
            return row;

        }


        public string Excel_NullToSom(string group_id, string name,
                         string subject, string number, string phone, string email, string state,
                         string company, string company_part, string company_location, string Picture)
        {
           
            if (group_id.Equals(string.Empty))
                group_id = "x";
            if (name.Equals(string.Empty))
                name = "x";
            if (subject.Equals(string.Empty))
                subject = "x";
            if (number.Equals(string.Empty))
                number = "x";
            if (phone.Equals(string.Empty))
                phone = "x";
            if (email.Equals(string.Empty))
                email = "x";
            if (state.Equals(string.Empty))
                state = "x";
            if (company.Equals(string.Empty))
                company = "x";
            if (company_part.Equals(string.Empty))
                company_part = "x";
            if (company_location.Equals(string.Empty))
                company_location = "x";
            if (Picture.Equals(string.Empty))
                Picture = "x";

            string msg = group_id + "#" + name + "#" + subject + "#" + number + "#" + phone + "#" + email + "#" + state + "#" + company + "#"
                + company_part + "#" + company_location + "#" + Picture;
            return msg;
        }

        public void Excel_Update_Member(int idx, string group_id, string name,
                         string subject, string number, string phone, string email, string state,
                         string company, string company_part, string company_location, string Picture)
        {

            string str = Excel_NullToSom(group_id, name, subject, number, phone, email, state, company, company_part, company_location, Picture);
            string[] temp = str.Split('#');
            
            ((Excel.Range)ws.Cells[idx, 2]).Value =temp[1];
            ((Excel.Range)ws.Cells[idx, 3]).Value =temp[2];
            ((Excel.Range)ws.Cells[idx, 4]).Value =temp[3];
            ((Excel.Range)ws.Cells[idx, 5]).Value =temp[4];
            ((Excel.Range)ws.Cells[idx, 6]).Value =temp[5];
            ((Excel.Range)ws.Cells[idx, 7]).Value =temp[6];
            ((Excel.Range)ws.Cells[idx, 8]).Value =temp[7];
            ((Excel.Range)ws.Cells[idx, 9]).Value =temp[8];
            ((Excel.Range)ws.Cells[idx, 10]).Value  = temp[9];
            ((Excel.Range)ws.Cells[idx, 11]).Value =  temp[10];
            ((Excel.Range)ws.Cells[idx, 12]).Value  = temp[11];



        }



        public void Excel_InsertMember(int idx, string group_id, string name,
                        string subject, string number, string phone, string email, string state,
                        string company, string company_part, string company_location, string Picture)
        {

            string str = Excel_NullToSom(group_id, name, subject, number, phone, email, state, company, company_part, company_location, Picture);
            string[] temp = str.Split('#');

            ((Excel.Range)ws.Cells[idx, 1]).Value = 0;
            ((Excel.Range)ws.Cells[idx, 2]).Value = temp[1];
            ((Excel.Range)ws.Cells[idx, 3]).Value = temp[2];
            ((Excel.Range)ws.Cells[idx, 4]).Value = temp[3];
            ((Excel.Range)ws.Cells[idx, 5]).Value = temp[4];
            ((Excel.Range)ws.Cells[idx, 6]).Value = temp[5];
            ((Excel.Range)ws.Cells[idx, 7]).Value = temp[6];
            ((Excel.Range)ws.Cells[idx, 8]).Value = temp[7];
            ((Excel.Range)ws.Cells[idx, 9]).Value = temp[8];
            ((Excel.Range)ws.Cells[idx, 10]).Value = temp[9];
            ((Excel.Range)ws.Cells[idx, 11]).Value = temp[10];
            ((Excel.Range)ws.Cells[idx, 12]).Value = temp[11];



        }

        #endregion
    }
}
