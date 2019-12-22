using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Directory
{
    class MainControl
    {
        public static MainControl main;

        public DBControl dbcon = new DBControl();
        public ExcelControl exlctr = new ExcelControl();

        string searchtext = string.Empty;
        string searchsort = string.Empty;
        static MainControl()
        {
            main = new MainControl();
        }
        private MainControl()
        {

        }
        ~MainControl()
        {
            DB_DisConnect();
        }

        #region DB ui

        public void DB_Connect(ListView listView1)
        {
            dbcon.DB_Connect();
            List<string> temp = dbcon.DB_PrintAllMembers();
            dbcon.DBPrintAll(listView1, temp);
            exlctr.ConnectExcel();
        }

     
        public void DB_DisConnect()
        {
            dbcon.DB_DisConnect();
        }


        public void DB_PrintAll(ListView listView1)
        {
            List<string> temp = dbcon.DB_PrintAllMembers();
            dbcon.DBPrintAll(listView1, temp);
        }
        public void DB_AddBtn_Ui(string group_id, string name,
                            string subject, string number, string phone, string email, string state,
                            string company, string company_part, string company_location, string Picture)
        {

            dbcon.Insert_Member(group_id, name, subject,
             number, phone, email, state, company,
             company_part, company_location, Picture);

        }


        public void DB_UpdateBtn_Ui(int id, string group_id, string name,
                            string subject, string number, string phone, string email, string state,
                            string company, string company_part, string company_location, string Picture)
        {


            dbcon.Update_Member(id, group_id, name, subject,
             number, phone, email, state, company,
             company_part, company_location, Picture);

        }

        public void DB_DeleteBtn_Ui(int id)
        {
            dbcon.Delete_Member(id);

        }



        #endregion



        #region Excel UI
        public void ExcelConnect()
        {

        }
        public void ExcelClose()
        {

        }




        #region 엑셀 검색 창 속 UI
        public void Excel_SelectBtn_Ui(string Text, string TextBoxSeach,ListView listView2)
        {
            searchtext = TextBoxSeach;
            searchsort = Text;

            List<string> temp = dbcon.DB_Select_Members(Text, TextBoxSeach);
            dbcon.DBPrintAll(listView2, temp);
        }


        public void Excel_LoadBtn_Ui(ListView listView3, TabControl tabControl2)
        {
         
            exlctr.PrintAll(listView3);
            tabControl2.SelectTab(tabControl2.TabPages[1].Name);
            //  tabControl2.FindForm().

        }

        public void Excel_SaveBtn_Ui(TextBox TextBoxSeach)
        {
            try
            {
                if (searchtext.Equals(string.Empty) || searchsort.Equals(string.Empty))
                    return;

                List<string> str = dbcon.DB_Select_Members(searchsort, searchtext);
                int idx = exlctr.FindLastIdx();
    
                foreach (string temp in str)
                {

                    string[] msg = temp.Split('#');
              
                    exlctr.InsertMember(idx, msg);
                    idx++;
                    
                }
                exlctr.SaveExcelFile();
            }
            catch(Exception ex)
            {
           
                return;
            }
    
        }
        #endregion

        #region 엑셀 리스튜 뷰 UI




        public void Excel_UpdaeMember_Ui(ListView listView3, string TextBoxNum, string TextBoxName,
                         string TextBoxSubject, string TextBoxSubjectNum, string TextBoxPhone, string TextBoxEMail, string TextBoxState,
                         string TextBoxCompany, string TextBoxCompanyPart, string TextBoxLocation, string Picture)
        {
            try
            {
    
                int num = listView3.SelectedIndices[0];

                exlctr.Excel_Update_Member(num + 1, TextBoxNum, TextBoxName, TextBoxSubject,
                 TextBoxSubjectNum, TextBoxPhone, TextBoxEMail, TextBoxState, TextBoxCompany,
                 TextBoxCompanyPart, TextBoxLocation, Picture);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Excel_InsertMember_Ui(string TextBoxNum, string TextBoxName,
                         string TextBoxSubject, string TextBoxSubjectNum, string TextBoxPhone, string TextBoxEMail, string TextBoxState,
                         string TextBoxCompany, string TextBoxCompanyPart, string TextBoxLocation, string Picture)
        {
            
            int idx = exlctr.FindLastIdx();
            //int num = listView1.SelectedIndices[0];
            exlctr.Excel_InsertMember(idx + 1, TextBoxNum, TextBoxName, TextBoxSubject,
              TextBoxSubjectNum, TextBoxPhone, TextBoxEMail, TextBoxState, TextBoxCompany,
              TextBoxCompanyPart, TextBoxLocation, Picture);
         
        }

        // [DB 저장 버튼]
        public void Excel_InsertMemberDB_Ui(ListView listView3, string TextBoxNum, string TextBoxName,
                         string TextBoxSubject, string TextBoxSubjectNum, string TextBoxPhone, string TextBoxEMail, string TextBoxState,
                         string TextBoxCompany, string TextBoxCompanyPart, string TextBoxLocation, string Picture)
        {


            for (int i = 0; i < listView3.Items.Count; i++)
            {

                TextBoxName = listView3.Items[i].SubItems[2].Text;
                TextBoxSubject = listView3.Items[i].SubItems[3].Text;
                TextBoxSubjectNum = listView3.Items[i].SubItems[4].Text;
                TextBoxEMail = listView3.Items[i].SubItems[6].Text;
                TextBoxPhone = listView3.Items[i].SubItems[5].Text;
                TextBoxState = listView3.Items[i].SubItems[7].Text;
                TextBoxCompany = listView3.Items[i].SubItems[8].Text;
                TextBoxCompanyPart = listView3.Items[i].SubItems[9].Text;
                TextBoxNum = listView3.Items[i].SubItems[1].Text;

                dbcon.Insert_Member(TextBoxNum, TextBoxName, TextBoxSubject,
               TextBoxSubjectNum, TextBoxPhone, TextBoxEMail, TextBoxState, TextBoxCompany,
               TextBoxCompanyPart, TextBoxLocation, Picture);
            }

        }

        // [Excel List초기화 버튼]
        public void Excel_ResetExcel_Ui(ListView listView3)
        {
          

            exlctr.DeleteAll(exlctr.FindLastIdx());

            exlctr.PrintAll(listView3);

   
        }

        #endregion


        #endregion



    }
}
