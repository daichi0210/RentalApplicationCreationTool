namespace RentalApplicationCreationTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // �g�p������[���̑�]���`�F�b�N���ꂽ�ꍇ�A�e�L�X�g�{�b�N�X��L��������
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

        // �g�p�����ݔ���[���̑�]���`�F�b�N���ꂽ�ꍇ�A�e�L�X�g�{�b�N�X��L��������
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
            // �g�p�������m�F
            if (checkBoxConferenceRoom1.Checked && checkBoxConferenceRoom2.Checked)
            {
                MessageBox.Show("��c���@�A");
            }
            if (checkBoxJapaneseStyleRoom1.Checked && checkBoxJapaneseStyleRoom2.Checked)
            {
                MessageBox.Show("�a���@�A");
            }
        }
    }
}
