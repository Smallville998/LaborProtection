using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace FireSecurityWF
{
    public partial class Form1 : Form
    {
        private string fullName = "";
        private string post = "";
        private string pathDocument;
        

        public Form1()
        {

            InitializeComponent();
            Console.WriteLine((comboBox1.Items.Count));

        }

        private void button1_Click(object sender, EventArgs e)
        {


             try 
             {

                 Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                 Document doc = app.Documents.Add(Visible: false);
                 Microsoft.Office.Interop.Word.Range r = doc.Range();
                 r.Paragraphs[1].Range.Font.Name = "Times New Roman";
                 r.Paragraphs[1].Range.Font.Size = 14;
                 r.Paragraphs[1].Range.Text = $"{fullName} работает на должности {post}";

                /* //app.Documents.Open(@"C:\Users\Smallville\Desktop\Code\C#\MS_VS\Fire_Safity\FireSecurityWF");
                 *//*if (fullName.Equals("") && post.Equals(""))
                 {
                     button1.Visible = false;
                 }
                 else
                     button1.Visible = true;*/

                 doc.SaveAs2($"{fullName} {post}.doc" );
                 pathDocument = doc.FullName;
                 doc.Close();
                 app.Quit();

                Post.FillPositions(post);

                MessageBox.Show($"Файл сохранен по пути {pathDocument}");

             }

             catch(Exception a) 
             {
                 MessageBox.Show(a.Message);
             }    
            
            
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            post = textBox1.Text;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            fullName = textBox2.Text;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = Convert.ToString(comboBox1.Items[comboBox1.SelectedIndex]);
            post = Convert.ToString(comboBox1.Items[comboBox1.SelectedIndex]);
            

            
        }
    }
}

