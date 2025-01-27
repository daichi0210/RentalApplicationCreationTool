﻿using RentalApplicationCreationTool;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace RentalApplicationCreationTool
{
    public partial class FormUserList : Form
    {
        Form1 f1;

        public FormUserList(Form1 form1)
        {
            InitializeComponent();
            f1 = form1;
        }

        private void FormUserList_FormClosed(object sender, FormClosedEventArgs e)
        {
            f1.Show();
        }
    }
}
