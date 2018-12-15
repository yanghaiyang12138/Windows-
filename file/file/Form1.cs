using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.IO;

namespace file
{
    public partial class Form1 : Form
    {
        //private static ArrayList fileList = new ArrayList();
        private static string line="";

        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            int lbxLength = this.listBox1.Items.Count;//listbox的长度   
            int iselect = this.listBox1.SelectedIndex;//listbox选择的索引   
            if (lbxLength > iselect && iselect < lbxLength - 1)
            {
                object oTempItem = this.listBox1.SelectedItem;
                this.listBox1.Items.RemoveAt(iselect);
                this.listBox1.Items.Insert(iselect + 1, oTempItem);
                this.listBox1.SelectedIndex = iselect + 1;
            }

        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            //textBox2.
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int lbxLength = this.listBox1.Items.Count;//listbox的长度   
            int iselect = this.listBox1.SelectedIndex;//listbox选择的索引   
            if (lbxLength > iselect && iselect > 0)
            {
                object oTempItem = this.listBox1.SelectedItem;
                this.listBox1.Items.RemoveAt(iselect);
                this.listBox1.Items.Insert(iselect - 1, oTempItem);
                this.listBox1.SelectedIndex = iselect - 1;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog opf = new FolderBrowserDialog();
            opf.ShowDialog();

            label2.Text = opf.SelectedPath;

            
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //读入流
            string[] sfile = (string[])listBox1.Items.Cast<string>().ToArray();
            for(int i = 0; i < sfile.Length; i++)
            {
                StreamReader sr = new StreamReader(sfile[i], Encoding.Default);
                //写入流
                FileStream fs = new FileStream("F:\\"+textBox1.Text, FileMode.Append);
                StreamWriter sw = new StreamWriter(fs);

                for (; (line = sr.ReadLine()) != null;)
                    sw.Write(line += "\r\n");
                sw.Flush();
                sw.Close();
                fs.Close();
                
            }
            MessageBox.Show("任务完成");
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.listBox1.Items.RemoveAt(this.listBox1.SelectedIndex);
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add(this.listBox2.SelectedItem);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string[] files = Directory.GetFiles(label2.Text, textBox2.Text);
            foreach (string file in files)
            {
                listBox2.Items.Add(file);
            }
        }
    }
}
