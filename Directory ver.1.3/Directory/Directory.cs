using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;


namespace Directory
{
    public partial class Directory : Form
    {

        string Picture;

        public Directory()
        {
            InitializeComponent();
            //MainControl.main.DB_Connect(listView1);
            ComboBoxSeach.Text = "이름";
        }




        #region 메시지 처리 정리

        // TextBox ReadOnly처리
        void TextReadTrue(int a)
        {
            if (a == 0)
            {
                TextBoxName.ReadOnly = true;
                TextBoxSubject.ReadOnly = true;
                TextBoxSubjectNum.ReadOnly = true;
                TextBoxEMail.ReadOnly = true;
                TextBoxPhone.ReadOnly = true;

                TextBoxState.ReadOnly = true;
                TextBoxCompany.ReadOnly = true;
                TextBoxCompanyPart.ReadOnly = true;
                TextBoxNum.ReadOnly = true;
                TextBoxLocation.ReadOnly = true;
            }
            else
            {
                TextBoxName.ReadOnly = false;
                TextBoxSubject.ReadOnly = false;
                TextBoxSubjectNum.ReadOnly = false;
                TextBoxEMail.ReadOnly = false;
                TextBoxPhone.ReadOnly = false;

                TextBoxState.ReadOnly = false;
                TextBoxCompany.ReadOnly = false;
                TextBoxCompanyPart.ReadOnly = false;
                TextBoxNum.ReadOnly = false;
                TextBoxLocation.ReadOnly = false;
            }
        }

        // TextBox들 초기화
        void TextEmty()
        {

            TextBoxName.Text = string.Empty;
            TextBoxSubject.Text = string.Empty;
            TextBoxSubjectNum.Text = string.Empty;
            TextBoxCompany.Text = string.Empty;
            TextBoxEMail.Text = string.Empty;
            TextBoxPhone.Text = string.Empty;
            TextBoxState.Text = string.Empty;
            TextBoxCompany.Text = string.Empty;
            TextBoxNum.Text = string.Empty;
            TextBoxCompanyPart.Text = string.Empty;
            TextBoxLocation.Text = string.Empty;
        }

        void MessageChoice(int sort)
        {
            try
            {
               // int idx = listView1.SelectedIndices[0];
                ListView listTemp;
                switch(sort)
                {
                    case 1:  listTemp = listView1; break;
                    case 2:  listTemp = listView2; break;
                    case 3:  listTemp = listView3; break;
                    default: return;
                }
               
                TextBoxName.Text = listTemp.SelectedItems[0].SubItems[2].Text;
                TextBoxSubject.Text = listTemp.SelectedItems[0].SubItems[3].Text;
                TextBoxSubjectNum.Text = listTemp.SelectedItems[0].SubItems[5].Text;
                TextBoxCompanyPart.Text = listTemp.SelectedItems[0].SubItems[9].Text;
                TextBoxEMail.Text = listTemp.SelectedItems[0].SubItems[6].Text;
                TextBoxPhone.Text = listTemp.SelectedItems[0].SubItems[5].Text;
                TextBoxState.Text = listTemp.SelectedItems[0].SubItems[7].Text;
                TextBoxCompany.Text = listTemp.SelectedItems[0].SubItems[8].Text;
                TextBoxLocation.Text = listTemp.SelectedItems[0].SubItems[10].Text;
                TextBoxNum.Text = listTemp.SelectedItems[0].SubItems[1].Text;
                pictureBox1.Image = Bitmap.FromFile( listTemp.SelectedItems[0].SubItems[11].Text);
            }
            catch (Exception ex)
            {
                pictureBox1.Image = null;
            }
        }


        #endregion


        // [ 이미지 등록 ]
        private void BtnPic_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // 사진 불러오기
            string file_path = null;
            openFileDialog1.InitialDirectory = @"C:\\";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file_path = openFileDialog1.FileName;
            }
            else if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            pictureBox1.Image = Bitmap.FromFile(file_path);

            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;

            Picture = file_path;
        }


        // [ 회원 등록 ] 
        bool Addmem = true;
        private void AddBtn_Click(object sender, EventArgs e)
        {

            if (TextBoxName.Text.Equals(string.Empty)|| Addmem == false)
            {
                TextReadTrue(1);
                TextEmty();
                Addmem = true;
                return;
            }
            else
            {
                MainControl.main.DB_AddBtn_Ui(TextBoxNum.Text, TextBoxName.Text, TextBoxSubject.Text,
              TextBoxSubjectNum.Text, TextBoxPhone.Text, TextBoxEMail.Text, TextBoxState.Text, TextBoxCompany.Text,
              TextBoxCompanyPart.Text, TextBoxLocation.Text,Picture);
                MessageBox.Show("저장 되었습니다.");
                MainControl.main.DB_PrintAll(listView1);

                Addmem = false;
            }

            
        }

        // [ 회원 수정 ] 
        private void UpdateBtn_Click(object sender, EventArgs e)
        {
            try
            {
                TextReadTrue(1);
                int num = listView1.SelectedIndices[0];
                MainControl.main.DB_UpdateBtn_Ui(int.Parse(listView3.Items[num].SubItems[0].Text), TextBoxNum.Text, TextBoxName.Text, TextBoxSubject.Text,
                 TextBoxSubjectNum.Text, TextBoxPhone.Text, TextBoxEMail.Text, TextBoxState.Text, TextBoxCompany.Text,
                 TextBoxCompanyPart.Text, TextBoxLocation.Text, Picture);
                MessageBox.Show("수정 되었습니다.");
                MainControl.main.DB_PrintAll(listView1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("수정 실패. 이유: " + ex.Message);
            }


        }



        // [ 회원 삭제 ] 
        private void DeleteBtn_Click(object sender, EventArgs e)
        {
            try
            {
                // TextReadTrue(1);
                TextEmty();
                int idx = listView1.SelectedIndices[0];
                MainControl.main.DB_DeleteBtn_Ui(int.Parse(listView1.Items[idx].SubItems[0].Text));
                MessageBox.Show("삭제 되었습니다.");
                MainControl.main.DB_PrintAll(listView1);

            }
            catch (Exception ex)
            {
                MessageBox.Show("삭제 실패");
            }
        }



        // [DB 리스트뷰] 선택시 해당 Row 위치를 가져옴.
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Addmem = false;
            MessageChoice(1);
        }




        // [ Excel VIew - 선택정보삭제 버튼 ]
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // TextReadTrue(1);
                int idx = listView3.SelectedIndices[0];
                
                MainControl.main.exlctr.DeleteExcelFile(idx+1);
                MainControl.main.Excel_LoadBtn_Ui(listView3, tabControl2);
            }
            catch (Exception ex)
            {

            }
        }



        #region 검색 Control

        // [ 검색 - 검색버튼 ]
        private void SeachBtn_Click(object sender, EventArgs e)
        {
            MainControl.main.Excel_SelectBtn_Ui(ComboBoxSeach.SelectedItem.ToString(), TextBoxSeach.Text, listView2);
        }


        // [ 검색 - Excel 불러오기 버튼 ]
        private void ExcelLoadBtn_Click(object sender, EventArgs e)
        {
            MainControl.main.Excel_LoadBtn_Ui(listView3, tabControl2);
        }


        // [ 검색 - Excel 저장하기 버튼 ]
        private void ExcelSaveBtn_Click(object sender, EventArgs e)
        {
            MainControl.main.Excel_SaveBtn_Ui(TextBoxSeach);
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                //    TextReadTrue(1);        
                MainControl.main.Excel_UpdaeMember_Ui(listView3, TextBoxNum.Text, TextBoxName.Text, TextBoxSubject.Text,
                 TextBoxSubjectNum.Text, TextBoxPhone.Text, TextBoxEMail.Text, TextBoxState.Text, TextBoxCompany.Text,
                 TextBoxCompanyPart.Text, TextBoxLocation.Text, Picture);
                MainControl.main.Excel_LoadBtn_Ui(listView3, tabControl2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MainControl.main.Excel_InsertMember_Ui(TextBoxNum.Text, TextBoxName.Text, TextBoxSubject.Text,
               TextBoxSubjectNum.Text, TextBoxPhone.Text, TextBoxEMail.Text, TextBoxState.Text, TextBoxCompany.Text,
               TextBoxCompanyPart.Text, TextBoxLocation.Text, Picture);
            MainControl.main.Excel_LoadBtn_Ui(listView3, tabControl2);
        }

        // [DB 저장 버튼]
        private void button4_Click(object sender, EventArgs e)
        {
            MainControl.main.Excel_InsertMemberDB_Ui(listView3, TextBoxNum.Text, TextBoxName.Text, TextBoxSubject.Text,
               TextBoxSubjectNum.Text, TextBoxPhone.Text, TextBoxEMail.Text, TextBoxState.Text, TextBoxCompany.Text,
               TextBoxCompanyPart.Text, TextBoxLocation.Text, Picture);
        }

        // [Excel List초기화 버튼]
        private void button5_Click(object sender, EventArgs e)
        {
            TextEmty();
            MainControl.main.Excel_ResetExcel_Ui(listView3);
            MainControl.main.Excel_LoadBtn_Ui(listView3, tabControl2);
        }

        private void listView3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Addmem = false;
            MessageChoice(3);
        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Addmem = false;
            MessageChoice(2);
        }


    }
}
