using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using MixERP.Net.VCards;
using MixERP.Net.VCards.Models;
using MixERP.Net.VCards.Serializer;
using MixERP.Net.VCards.Types;
using MessageBox = System.Windows.MessageBox;
using Path = System.Windows.Shapes.Path;

namespace Excel2vCard
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public List<ContactModel> ContactList { get; set; }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (ContactList==null||ContactList.Count < 0)
            {
                MessageBox.Show("没有数据");
                return;
            }

            StringBuilder serializedStr = new StringBuilder();
            foreach (var contactModel in ContactList)
            {
                var vcard = new VCard
            {
                Version = VCardVersion.V2_1,
                FormattedName = contactModel.Name,
                FirstName = "",
                LastName = "",
                //Classification = ClassificationType.Confidential,
                //Categories = new[] { "Friend", "Fella", "Amsterdam" }
            };
                var Telephones = new List<Telephone>();
                if (!string.IsNullOrEmpty(contactModel.Phone1))
                {
                    var telephone = new Telephone() {Number = contactModel.Phone1, Type = TelephoneType.Cell};
                    Telephones.Add(telephone);
                }
                if (!string.IsNullOrEmpty(contactModel.Phone2))
                {
                    var telephone = new Telephone() {Number = contactModel.Phone2, Type = TelephoneType.Home};
                    Telephones.Add(telephone);
                }
                if (!string.IsNullOrEmpty(contactModel.Phone3))
                {
                    var telephone = new Telephone() {Number = contactModel.Phone3, Type = TelephoneType.Work};
                    Telephones.Add(telephone);
                }
                vcard.Telephones = Telephones;
                string serialized = vcard.Serialize();
                serializedStr.AppendLine(serialized);
            }
            
            string saveFilePath="";
            System.Windows.Forms.FolderBrowserDialog openFileDialog = new System.Windows.Forms.FolderBrowserDialog();  //选择文件夹
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                saveFilePath = openFileDialog.SelectedPath;
            }

            if (string.IsNullOrEmpty(saveFilePath))
            {
                MessageBox.Show("请选择要保存的文件夹");
            }
            var fileName= System.IO.Path.GetFileNameWithoutExtension(this.txt_FilePath.Text);// 没有扩展名的文件名
            string path = saveFilePath + "\\"+fileName + ".vcf";
            File.WriteAllText(path, serializedStr.ToString());
            MessageBox.Show("保存成功！");
        }

        private void btn_OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择文件";
            openFileDialog.Filter = "excel2003文件|*.xls|excel2007及以上文件|*.xlsx";
            openFileDialog.FileName = string.Empty;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            //openFileDialog.DefaultExt = "zip";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            string fileName = openFileDialog.FileName;
            this.txt_FilePath.Text = fileName;
            var dt = ExcelHelper.ExcelToDataTable(fileName, null, true);
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("Excel第一个表格中没有任何数据");
                return;
            }

            List<string> nameFieldAliasList = new List<string>() { "姓名","联系人","名字","人名"};
            List<string> phoneFieldAliasList = new List<string>() { "电话","手机","电话号码","手机号码","座机","固定电话","移动电话","家庭电话"};
            ContactList = new List<ContactModel>();
            foreach (DataRow row in dt.Rows)
            {
                ContactModel contact = new ContactModel();
                foreach (var nameFieldAlias in nameFieldAliasList)
                {
                    if (!dt.Columns.Contains(nameFieldAlias))
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(contact.Name)&&row[nameFieldAlias] != DBNull.Value)
                    {
                        contact.Name = row[nameFieldAlias].ToString();
                        continue;
                    }
                }
                //电话1
                foreach (var phoneFieldAlias in phoneFieldAliasList)
                {
                    if (!dt.Columns.Contains(phoneFieldAlias))
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(contact.Phone1) && row[phoneFieldAlias] != DBNull.Value)
                    {
                        contact.Phone1 = row[phoneFieldAlias].ToString();
                        continue;
                    }
                }
                //电话2
                foreach (var phoneFieldAlias in phoneFieldAliasList)
                {
                    if (!dt.Columns.Contains(phoneFieldAlias))
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(contact.Phone2) && row[phoneFieldAlias] != DBNull.Value)
                    {
                        var phone=row[phoneFieldAlias].ToString();
                        if (!string.IsNullOrEmpty(contact.Phone1) &&contact.Phone1!=phone)
                        {
                            contact.Phone2 = phone;
                            continue;
                        }
                    }
                }
                //电话3
                foreach (var phoneFieldAlias in phoneFieldAliasList)
                {
                    if (!dt.Columns.Contains(phoneFieldAlias))
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(contact.Phone3) && row[phoneFieldAlias] != DBNull.Value)
                    {
                        var phone=row[phoneFieldAlias].ToString();
                        if (!string.IsNullOrEmpty(contact.Phone1) &&
                            contact.Phone1!=phone&&
                            !string.IsNullOrEmpty(contact.Phone2) &&
                            contact.Phone2!=phone)
                        {
                            contact.Phone3 = phone;
                            continue;
                        }
                    }
                }
                
                if (string.IsNullOrEmpty(contact.Phone1)&&
                    dt.Columns.Contains("电话1") && 
                    row["电话1"] != DBNull.Value
                    )
                {
                    contact.Phone1 = row["电话1"].ToString();
                }

                if (string.IsNullOrEmpty(contact.Phone2)&&
                    dt.Columns.Contains("电话2") && 
                    row["电话2"] != DBNull.Value
                    )
                {
                    contact.Phone2 = row["电话2"].ToString();
                }

                if (string.IsNullOrEmpty(contact.Phone3)&&
                    dt.Columns.Contains("电话3") && 
                    row["电话3"] != DBNull.Value
                    )
                {
                    contact.Phone3 = row["电话3"].ToString();
                }
                ContactList.Add(contact);
            }

            if (ContactList.Count > 0)
            {
                this.grid_Contact.ItemsSource = ContactList;
            }
        }
    }
}
