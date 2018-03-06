using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WizBolt
{
    public partial class LTF : Form
    {
        public decimal Minimum_LTF = 0M;
        public decimal Maximum_LTF = 0M;
        public decimal Entered_LTF = 0M;
        public LTF()
        {
            InitializeComponent();
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Ok_Button_Click(object sender, EventArgs e)
        {
            Entered_LTF = Convert.ToDecimal(LTF_TextBox.Text);
            if (Entered_LTF < Minimum_LTF)
            {
                DialogResult UserResponse = MessageBox.Show("Load Transfer Factor cannot be less than " + LTF_LowerValue_Label.Text + ".", "Lower LTF Error!");
                Entered_LTF = 0;
            }
            else if (Entered_LTF > Maximum_LTF)
            {
                DialogResult UserResponse = MessageBox.Show("Load Transfer Factor cannot be more than " + LFT_HigherValue_Label.Text + ".", "Higher LTF Error!");
                Entered_LTF = 0;
            }
            else if ((Entered_LTF > Minimum_LTF) && (Entered_LTF < Maximum_LTF))
            {
                this.Close();
            }
            
        }

        private void LTF_Load(object sender, EventArgs e)
        {
            ActiveControl = LTF_TextBox;
            LTF_LowerValue_Label.Text = Minimum_LTF.ToString();
            LFT_HigherValue_Label.Text = Maximum_LTF.ToString();
        }
    }
}
