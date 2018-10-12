using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AntPPTUpLoad
{
    public partial class Form2 : Form
    {
        public Form1 f1 = new Form1();
        public int total = 0;
        Thread th = null;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.total = Form1.ResTxtPathList.Count;
            this.label1.Text = "0/" + this.total;
        }

        private void button1_Click(object sender, EventArgs e)
        {

                f1.F2StartF1();
                this.total = Form1.ResTxtPathList.Count;
               
                th = new Thread(UpdataPro);
                th.IsBackground = true;
                th.Start();           
        }

        public void UpdataPro()
        {
            while (true)
            {
                this.total = Form1.ResTxtPathList.Count;

                this.label1.Invoke(new Action(() =>
               {
                   this.label1.Text = Form1.suc + "/" + total;
               }));

                this.progressBar1.Invoke(new Action(() =>
                {
                    this.progressBar1.Maximum = this.total;
                    this.progressBar1.Value = Form1.suc;
                }));
                Thread.Sleep(1000);
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (f1 != null)
                f1.Show();
            if (th != null)
                th.Abort();
        }
    }
}
