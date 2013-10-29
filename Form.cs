using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace TestMailMerge
{
    public partial class Form : BaseForm
    {
        public Form()
        {
            InitializeComponent();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            //首先把建立的範本檔案讀入MemoryStream
            System.IO.MemoryStream _memoryStream = new System.IO.MemoryStream(Properties.Resources.TestMailMerge);

            //建立一個Document物件
            //並傳入MemoryStream
            Aspose.Words.Document doc = new Aspose.Words.Document(_memoryStream);

            //新增一個DataTable
            DataTable table = new DataTable();
            //建立Column
            table.Columns.Add("學校代碼");
            table.Columns.Add("預設學年度");
            table.Columns.Add("預設學期");
            table.Columns.Add("學校中文名稱");
            table.Columns.Add("學校英文名稱");
            table.Columns.Add("學校中文地址");

            table.Columns.Add("學校英文地址");

            table.Columns.Add("學校電話");
            table.Columns.Add("學校傳真");

            table.Columns.Add("校長中文名稱");
            table.Columns.Add("校長英文名稱");
            table.Columns.Add("教務主任姓名");
            table.Columns.Add("學務主任姓名");
            table.Columns.Add("列印日期");

            //透過建立的DataTable物件來New一個儲存資料的Row
            DataRow row = table.NewRow();
            //這些Row具有上面所建立相同的Column欄位
            //因此可以直接指定欄位名稱將資料填入裡面
            row["學校代碼"] = K12.Data.School.Code;
            row["預設學年度"] = K12.Data.School.DefaultSchoolYear;
            row["預設學期"] = K12.Data.School.DefaultSemester;
            row["學校中文名稱"] = K12.Data.School.ChineseName;
            row["學校英文名稱"] = K12.Data.School.EnglishName;
            row["學校中文地址"] = K12.Data.School.Address;
            row["學校電話"] = K12.Data.School.Telephone;
            row["學校傳真"] = K12.Data.School.Fax;
            row["列印日期"] = DateTime.Today.ToShortDateString();

            //把所建立的資料行加入Table的Row清單內
            table.Rows.Add(row);


            //將DataTable傳入Document的MailMerge.Execute()方法
            doc.MailMerge.Execute(table);
            //清空所有未被合併的功能變數
            doc.MailMerge.DeleteFields();
            //將檔案儲存至c:\
            doc.Save("C:\\學校基本資料說明書.doc");

        }
    }
}
