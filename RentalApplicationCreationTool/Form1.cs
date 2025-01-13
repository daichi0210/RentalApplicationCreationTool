namespace RentalApplicationCreationTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // 使用室名の[その他]がチェックされた場合、テキストボックスを有効化する
        private void checkBoxOtherRooms_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxOtherRooms.Checked == true)
            {
                textBoxOtherRooms.Enabled = true;
            }
            else
            {
                textBoxOtherRooms.Enabled = false;
            }
        }

        // 使用附属設備の[その他]がチェックされた場合、テキストボックスを有効化する
        private void checkBoxOtherEquipment_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxOtherEquipment.Checked)
            {
                textBoxOtherEquipment.Enabled = true;
            }
            else
            {
                textBoxOtherEquipment.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 使用室名を確認
            if (checkBoxConferenceRoom1.Checked && checkBoxConferenceRoom2.Checked)
            {
                MessageBox.Show("会議室①②");
            }
            if (checkBoxJapaneseStyleRoom1.Checked && checkBoxJapaneseStyleRoom2.Checked)
            {
                MessageBox.Show("和室①②");
            }
        }
    }
}
