using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Threading;

namespace Ultimate_executer
{
    public partial class Form1 : Form
    {
        string ep_path;
        string idf_folder;
        string[] idfs;
        string epw_folder;
        string[] epws;
        string output_folder;
        bool expobj;

        bool ep_path_ok;
        bool idfs_ok;
        bool epws_ok;
        bool output_folder_ok;

        List<string[]> paras;

        private string EP_Path
        {
            set
            {
                if (File.Exists(value + @"\energyplus.exe") && File.Exists(value + @"\Energy+.idd") && File.Exists(value + @"\ExpandObjects.exe") && File.Exists(value + @"\PostProcess\ReadVarsESO.exe"))
                {
                    ep_path = value;
                    ep_path_ok = true;
                }
            }
            get
            {
                if (ep_path_ok) { return ep_path; }
                else { return "EnergyPlus path Error"; }
            }
        }

        private string IDF_FOLDER
        {
            set
            {
                if (Directory.Exists(value))
                {
                    idf_folder = value;
                    idfs = Directory.GetFiles(idf_folder, "*.idf");
                    idfs_ok = idfs.Length > 0;
                }
            }
            get
            {
                if (idfs_ok) { return idf_folder; }
                else { return "IDFs folder Error"; }
            }
        }

        private string EPW_FOLDER
        {
            set
            {
                if (Directory.Exists(value))
                {
                    epw_folder = value; 
                    epws = Directory.GetFiles(epw_folder, "*.epw");
                    epws_ok = epws.Length > 0;
                }
            }
            get
            {
                if (epws_ok) { return epw_folder; }
                else { return "EPWs folder Error"; }
            }
        }

        private string OUTPUT_FOLDER
        {
            set
            {
                if (Directory.Exists(value))
                {
                    output_folder = value;
                    output_folder_ok = true;
                }
            }
            get
            {
                if (output_folder_ok) { return output_folder; }
                else { return "Output path Error"; }
            }
        }

        public Form1()
        {
            InitializeComponent();

            FResize();

            ep_path = "";
            idf_folder = "";
            epw_folder = "";
            output_folder = "";
            expobj = checkBox1.Checked;

            ep_path_ok = false;
            idfs_ok = false;
            epws_ok = false;
            output_folder_ok = false;

            comboBox1.Items.Add("ALL");

            for (int i = 1; i < 33; i++)
            {
                comboBox1.Items.Add(i.ToString());
            }

            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = false;
            checkBox1.Checked = true;

            richTextBox1.Text = "";

            paras = new List<string[]>();

            //textBox1.Text = @"C:\EnergyPlusV8-8-0";
            //textBox2.Text = @"C:\FTP\JPmodels\Ultimate_demo";
            //textBox3.Text = @"C:\FTP\JPmodels";
            //textBox4.Text = @"C:\temp";
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fbd.SelectedPath;
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            EP_Path = textBox1.Text;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = fbd.SelectedPath;
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            IDF_FOLDER = textBox2.Text;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = fbd.SelectedPath;
            }
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            EPW_FOLDER = textBox3.Text;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = fbd.SelectedPath;
            }
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            OUTPUT_FOLDER = textBox4.Text;
        }
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            expobj = checkBox1.Checked;
        }
        private void Button5_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;

            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;

            checkBox1.Enabled = false;
            comboBox1.Enabled = false;

            richTextBox1.Text = "";

            if (!ep_path_ok) { richTextBox1.Text += EP_Path + Environment.NewLine; }
            if (!idfs_ok) { richTextBox1.Text += IDF_FOLDER + Environment.NewLine; }
            if (!epws_ok) { richTextBox1.Text += EPW_FOLDER + Environment.NewLine; }
            if (!output_folder_ok) { richTextBox1.Text += OUTPUT_FOLDER + Environment.NewLine; }

            if (ep_path_ok && idfs_ok && epws_ok && output_folder_ok)
            {
                richTextBox1.Text += "Making task queue" + Environment.NewLine;
                List<EP_run> tasks = new List<EP_run>();
                foreach (string epw in epws)
                {
                    foreach (string idf in idfs)
                    {
                        tasks.Add(new EP_run(ep_path, idf, epw, output_folder, expobj));
                    }
                }

                richTextBox1.Text += "Creating temp folders" + Environment.NewLine;
                Directory.CreateDirectory(output_folder + @"\temp\");
                Directory.CreateDirectory(output_folder + @"\Warnings\");
                Directory.CreateDirectory(output_folder + @"\Errors\");

                foreach(string epw in epws)
                {
                    Directory.CreateDirectory(String.Format(@"{0}\{1}\", output_folder, Path.GetFileNameWithoutExtension(epw)));
                }

                foreach (EP_run task in tasks) { task.Creat_Temp_Folder(richTextBox1); }

                richTextBox1.Text += Environment.NewLine;

                richTextBox1.Text += "Simulation starts ... " + Environment.NewLine;
                Directory.SetCurrentDirectory(ep_path);
                Parallel.ForEach(tasks, (item, loopState) => { item.Run(); });

                richTextBox1.Text += Environment.NewLine;

                richTextBox1.Text += "Converting ESO to CSV" + Environment.NewLine;
                foreach (EP_run task in tasks) { task.ReadVarEso(richTextBox1); }

                richTextBox1.Text += Environment.NewLine + "ALL DONE";
            }

            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;

            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;

            checkBox1.Enabled = true;
            comboBox1.Enabled = false;
        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            FResize();
        }
        private void FResize()
        {
            textBox1.Width = this.Width - 175;
            textBox2.Width = this.Width - 175;
            textBox3.Width = this.Width - 175;
            textBox4.Width = this.Width - 175;

            richTextBox1.Height = this.Height - 192;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }
    }

    public class EP_run
    {
        string ep_path;
        string idf;
        string epw;
        string out_path;
        bool expobj;
        public string report;

        string idf_name;
        string epw_name;
        string local_temp_path;
        public EP_run(string _ep_path, string _idf_path, string _epw_path, string _out_path, bool _expobj)
        {
            ep_path = _ep_path;
            idf = _idf_path;
            epw = _epw_path;
            out_path = _out_path;
            expobj = _expobj;

            idf_name = Path.GetFileNameWithoutExtension(idf);
            epw_name = Path.GetFileNameWithoutExtension(epw);
            local_temp_path = String.Format(@"{0}\temp\{1}.{2}\", out_path, idf_name, epw_name);
        }

        public void Creat_Temp_Folder(RichTextBox rtb)
        {
            Directory.CreateDirectory(local_temp_path);
            File.Copy(idf, local_temp_path + "in.idf", true);

            rtb.Text += String.Format("{0}.{1} created", idf_name, epw_name) + Environment.NewLine;

            if (expobj)
            {
                ProcessStartInfo eob_info = new ProcessStartInfo(ep_path + @"\ExpandObjects.exe");
                eob_info.UseShellExecute = true;
                eob_info.WorkingDirectory = local_temp_path;
                eob_info.CreateNoWindow = true;
                eob_info.WindowStyle = ProcessWindowStyle.Hidden;

                Process expobject = Process.Start(eob_info);
                expobject.WaitForExit();
                expobject.Close();
                if (File.Exists(local_temp_path + "expanded.idf"))
                {
                    if(File.Exists(local_temp_path + "_in._idf")) { File.Delete(local_temp_path + "_in._idf"); }
                    File.Move(local_temp_path + "in.idf", local_temp_path + "_in._idf");
                    File.Move(local_temp_path + "expanded.idf", local_temp_path + "in.idf");
                    rtb.Text += String.Format("{0} expamded", idf_name) + Environment.NewLine;
                }
            }
        }
        public void Run()
        {
            ProcessStartInfo run_info = new ProcessStartInfo(ep_path + @"\energyplus.exe");
            run_info.UseShellExecute = true;
            run_info.WorkingDirectory = local_temp_path;

            run_info.Arguments = "-w " + epw;
            run_info.Arguments += " -p eplusout";
            run_info.Arguments += " -d " + local_temp_path;
            run_info.Arguments += " -s D ";
            run_info.Arguments += local_temp_path + @"\in.idf";

            Process runep = Process.Start(run_info);
            runep.WaitForExit();
            runep.Close();

            string[] temp_files = Directory.GetFiles(local_temp_path);
            foreach (string f in temp_files)
            {
                string extension = Path.GetExtension(f);
                if (extension != ".err" && extension != ".eso") 
                { 
                    File.Delete(f); 
                }
            }
        }

        public void ReadVarEso(RichTextBox rtb)
        {
            ProcessStartInfo rve_info = new ProcessStartInfo(ep_path + @"\PostProcess\ReadVarsESO.exe");
            rve_info.UseShellExecute = true;
            rve_info.WorkingDirectory = local_temp_path;
            rve_info.CreateNoWindow = true;
            rve_info.WindowStyle = ProcessWindowStyle.Hidden;

            Process rve = Process.Start(rve_info);
            rve.WaitForExit();
            rve.Close();

            bool done = false;

            if (File.Exists(local_temp_path + "eplusout.csv"))
            {
                if (new FileInfo(local_temp_path + "eplusout.csv").Length > 10) { done = true; }
            }

            if (done)
            {
                string t0 = String.Format(@"{0}\Warnings\{1}.{2}.err", out_path, idf_name, epw_name);
                File.Copy(local_temp_path + "eplusout.err", t0, true);
                string t1 = String.Format(@"{0}\{1}\{2}.csv", out_path, epw_name, idf_name);
                File.Copy(local_temp_path + "eplusout.csv", t1, true);
                report = String.Format("{0}.{1}: Successfully", idf_name, epw_name);
            }
            else
            {
                File.Copy(local_temp_path + "eplusout.err", String.Format(@"{0}\Errors\{1}.{2}.err", out_path, idf_name, epw_name), true);
                report = String.Format("{0}.{1}: Errors happen", idf_name, epw_name);
            }

            string[] temp_files = Directory.GetFiles(local_temp_path);
            foreach (string f in temp_files) { File.Delete(f); }

            rtb.Text += report + Environment.NewLine;
        }
    }
}
