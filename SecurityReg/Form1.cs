using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SecurityReg
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try//32位
            {
                RegistryKey key = null;
                //key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel");
                key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                    .CreateSubKey(@"SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel");
                key.SetValue("MyComputer", "Enabled");
                key.SetValue("LocalIntranet", "Enabled");
                key.SetValue("Internet", "Enabled");
                key.SetValue("TrustedSites", "Enabled");
                key.SetValue("UntrustedSites", "Enabled");
                key.Close();
                MessageBox.Show("Security Reg Done!");

                //判断操作系统版本是64还是32，以此确定打开那个注册表
                //key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);
                //string str = key.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel").GetValue("MyComputer", "").ToString();
                //MessageBox.Show(str);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
}

        private void button2_Click(object sender, EventArgs e)
        {
            try//64位
            {
                Microsoft.Win32.RegistryKey key;
                key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                    .CreateSubKey(@"SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel");
                key.SetValue("MyComputer", "Enabled");
                key.SetValue("LocalIntranet", "Enabled");
                key.SetValue("Internet", "Enabled");
                key.SetValue("TrustedSites", "Enabled");
                key.SetValue("UntrustedSites", "Enabled");
                key.Close();
                MessageBox.Show("Security Reg Done!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
