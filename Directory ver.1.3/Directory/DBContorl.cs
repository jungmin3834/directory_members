using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Directory
{
    class DBControl
    {
        string connstring1 = @"Server=192.168.0.52;database=
                                   WB28;uid=ccm;pwd=ccm;";


        SqlConnection conn = new SqlConnection();
        #region DB접속
        public void DB_Connect()
        {
            conn.ConnectionString = connstring1;
            conn.Open();    //  데이터베이스 연결     
        }

        public void DB_DisConnect()
        {
            conn.Close();
        }


        #endregion

        #region DB 검색

        public List<string> DB_PrintAllMembers()
        {
            try
            {

                string comtext = "select * from WBMEMBER";

                SqlCommand command = new SqlCommand(comtext, conn);
                //SqlParameter param_title = new SqlParameter("@ACCESS", access);
                //command.Parameters.Add(param_title);


                //SqlParameter title = new SqlParameter("@ASK", ask);
                //command.Parameters.Add(title);

                SqlDataReader myDataReader;
                myDataReader = command.ExecuteReader();


                List<string> stringList = new List<string>();
                while (myDataReader.Read())
                {

                    string msg = string.Empty;
                    msg += myDataReader["memberid"].ToString() + "#";
                    msg += myDataReader["group_id"].ToString() + "#";
                    msg += myDataReader["name"].ToString() + "#";
                    msg += myDataReader["picture"].ToString() + "#";
                    msg += myDataReader["subject"].ToString() + "#";
                    msg += myDataReader["number"].ToString() + "#";
                    msg += myDataReader["phone"].ToString() + "#";
                    msg += myDataReader["email"].ToString() + "#";
                    msg += myDataReader["state"].ToString() + "#";
                    msg += myDataReader["company"].ToString() + "#";
                    msg += myDataReader["company_part"].ToString() + "#";
                    msg += myDataReader["company_location"].ToString();
                    stringList.Add(msg);
                }

                myDataReader.Close();
                return stringList;
            }
            catch (Exception ex)
            {
                MessageBox.Show("[DB 검색 에러]" + ex.Message);
                return null;
            }
        }

        public void DBPrintAll(ListView listView1, List<string> str)
        {
            listView1.Items.Clear();

            foreach (string temp in str)
            {
                string[] msg = temp.Split('#');
                String[] aa = {msg[0], msg[1], msg[2], msg[4], msg[5], msg[6], msg[7], msg[8], msg[9], msg[10], msg[11] ,msg[3]};
                ListViewItem newitem = new ListViewItem(aa);
                listView1.Items.Add(newitem);

            }
        }
        public List<string> DB_Select_Members(string access, string ask)
        {
            try
            {

                string comtext = string.Empty;
                switch (access)
                {
                    case "이름": comtext = "select *     from WBMEMBER where name = @ASK"; break;
                    case "학번": comtext = "select *  from WBMEMBER where number = @ASK"; break;
                    case "학과": comtext = "select *    from WBMEMBER where Phone = @ASK"; break;
                    case "기수": comtext = "select * from WBMEMBER where group_id = @ASK"; break;
                }

                SqlCommand command = new SqlCommand(comtext, conn);

                SqlParameter title = new SqlParameter("@ASK", ask);
                command.Parameters.Add(title);

                SqlDataReader myDataReader;
                myDataReader = command.ExecuteReader();


                Console.WriteLine("select * from WBMEMBER where CNAME = @NAME");
                List<string> stringList = new List<string>();
                while (myDataReader.Read())
                {
                    string msg = string.Empty;
                    msg += myDataReader["memberid"].ToString() + "#";     
                    msg += myDataReader["group_id"].ToString() + "#";
                    msg += myDataReader["name"].ToString() + "#";
                    msg += myDataReader["picture"].ToString() + "#";
                    msg += myDataReader["subject"].ToString() + "#";
                    msg += myDataReader["number"].ToString() + "#";
                    msg += myDataReader["phone"].ToString() + "#";
                    msg += myDataReader["email"].ToString() + "#";
                    msg += myDataReader["state"].ToString() + "#";
                    msg += myDataReader["company"].ToString() + "#";
                    msg += myDataReader["company_part"].ToString() + "#";
                    msg += myDataReader["company_location"].ToString();
                    stringList.Add(msg);

                }

                myDataReader.Close();

                return stringList;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        #endregion

        #region DB 기능 /삭제/추가/입력

        // [맴버 삭제 기능 함수]
        public bool Delete_Member(int memberid)
        {

            //1. 쿼리문 작성
            string sql = "Delete from WBMEMBER " +
                "where memberid = @MEMBERID";

            //2. 명령객체 생성
            SqlCommand cmd = new SqlCommand(sql, conn);

            //3. 파라미터 설정
            SqlParameter param_memberid = new SqlParameter("@MEMBERID", memberid);
            param_memberid.SqlDbType = System.Data.SqlDbType.Int;
            cmd.Parameters.Add(param_memberid);


            //4. 쿼리문 실행
            cmd.ExecuteNonQuery();

            return true;
        }

        // [맴버 추가 기능 함수]
        public bool Insert_Member(string group_id, string name,
                            string subject, string number, string phone, string email, string state,
                            string company, string company_part, string company_location, string Picture)
        {


            //1. 쿼리문 작성
            string sql = "INSERT INTO WBMEMBER " +
                "(group_id,name,subject,number,phone,email,state,company,company_part,company_location,picture) VALUES" +
                " (@GROUP_ID,@NAME,@SUBJECT,@NUMBER,@PHONE,@EMAIL,@STATE,@COMPANY,@COMPANY_PART,@COMPANY_LOCATION,@PICTURE)";

            //2. 명령객체 생성
            SqlCommand cmd = new SqlCommand(sql, conn);

            ////3. 파라미터 설정
            //SqlParameter param_memberid = new SqlParameter("@MEMBERID", memberid);
            //param_memberid.SqlDbType = System.Data.SqlDbType.Int;
            //cmd.Parameters.Add(param_memberid);

            //SqlParameter param_group_id = new SqlParameter("@GROUP_ID", group_id);
            //cmd.Parameters.Add(param_group_id);

            SqlParameter param_memberid = new SqlParameter("@GROUP_ID", int.Parse(group_id));
            param_memberid.SqlDbType = System.Data.SqlDbType.Int;
            cmd.Parameters.Add(param_memberid);

            SqlParameter param_name = new SqlParameter("@NAME", name);
            cmd.Parameters.Add(param_name);

            SqlParameter param_subject = new SqlParameter("@SUBJECT", subject);
            cmd.Parameters.Add(param_subject);

            SqlParameter param_number = new SqlParameter("@NUMBER", number);
            cmd.Parameters.Add(param_number);

            SqlParameter param_phone = new SqlParameter("@PHONE", phone);
            cmd.Parameters.Add(param_phone);

            SqlParameter Param_picture = new SqlParameter("@PICTURE", Picture);
            cmd.Parameters.Add(Param_picture);


            SqlParameter param_email = new SqlParameter("@EMAIL", email);
            cmd.Parameters.Add(param_email);

            SqlParameter param_state = new SqlParameter("@STATE", state);
            cmd.Parameters.Add(param_state);

            SqlParameter param_company = new SqlParameter("COMPANY", company);
            cmd.Parameters.Add(param_company);

            SqlParameter param_company_part = new SqlParameter("@COMPANY_PART", company_part);
            cmd.Parameters.Add(param_company_part);

            SqlParameter param_company_location = new SqlParameter("@COMPANY_LOCATION", company_location);
            cmd.Parameters.Add(param_company_location);

            cmd.ExecuteNonQuery();


            return true;
        }

        // [맴버 추가 기능 함수]
        public bool Update_Member(int idx, string group_id, string name,
                     string subject, string number, string phone, string email, string state,
                     string company, string company_part, string company_location, string Picture)
        {

            //1. 쿼리문 작성
            string sql = "UPDATE WBMEMBER " +
                         "SET GROUP_ID = @GROUP_ID," +
                         "NAME = @NAME," +
                         "SUBJECT = @SUBJECT," +
                         "NUMBER = @NUMBER," +
                         "PHONE = @PHONE," +
                         "EMAIL = @EMAIL," +
                         "STATE = @STATE," +
                         "COMPANY = @COMPANY," +
                         "COMPANY_PART = @COMPANY_PART," +
                         "COMPANY_LOCATION = @COMPANY_LOCATION," +
                         "picture = @PICTURE "+
                         "WHERE MEMBERID = @MEMBERID";
            // "WHERE MEMBERID = @MEMBERID";

            //2. 명령객체 생성
            SqlCommand cmd = new SqlCommand(sql, conn);

            //3. 파라미터 설정


            SqlParameter param_memberid = new SqlParameter("@MEMBERID", idx);
            param_memberid.SqlDbType = System.Data.SqlDbType.Int;
            cmd.Parameters.Add(param_memberid);

            SqlParameter param_group_id = new SqlParameter("@GROUP_ID", group_id);
            cmd.Parameters.Add(param_group_id);


            SqlParameter Param_picture = new SqlParameter("@PICTURE", Picture);
            cmd.Parameters.Add(Param_picture);

            SqlParameter param_name = new SqlParameter("@NAME", name);
            cmd.Parameters.Add(param_name);

            SqlParameter param_subject = new SqlParameter("@SUBJECT", subject);
            cmd.Parameters.Add(param_subject);

            SqlParameter param_number = new SqlParameter("@NUMBER", number);
            cmd.Parameters.Add(param_number);

            SqlParameter param_phone = new SqlParameter("@PHONE", phone);
            cmd.Parameters.Add(param_phone);

            SqlParameter param_email = new SqlParameter("@EMAIL", email);
            cmd.Parameters.Add(param_email);

            SqlParameter param_state = new SqlParameter("@STATE", state);
            cmd.Parameters.Add(param_state);

            SqlParameter param_company = new SqlParameter("COMPANY", company);
            cmd.Parameters.Add(param_company);

            SqlParameter param_company_part = new SqlParameter("@COMPANY_PART", company_part);
            cmd.Parameters.Add(param_company_part);

            SqlParameter param_company_location = new SqlParameter("@COMPANY_LOCATION", company_location);
            cmd.Parameters.Add(param_company_location);

            //4. 쿼리문 실행
            if (cmd.ExecuteNonQuery() == 1)
                return true;
            else
                return false;
        }


        #endregion


    }
}
