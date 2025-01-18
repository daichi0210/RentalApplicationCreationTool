using Word = Microsoft.Office.Interop.Word;

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
            string applicationDate;     // �\����
            string name;                // ���O
            string telephoneNumber;     // �d�b�ԍ�
            string address;             // �Z��
            string organizationName;    // �c�̖�
            string purposeOfUse;        // �g�p�ړI
            string numberOfPeople;         // �g�p�l��
            string dateOfUse;           // �g�p��
            string startTime;           // �J�n�����i���j
            string startMinutes;        // �J�n�����i���j
            string endTime;             // �I�������i���j
            string endMinutes;          // �I�������i���j
            string roomName = "";       // �g�p����
            string auxiliaryEquipmentUsed = "";     // �g�p�����ݔ�
            string reasonForApplyingForExemption;   // �g�p���̖Ə��\��



            // �\��������
            applicationDate = dateTimePickerApplicationDate.Value.ToString();
            // ���O����
            name = textBoxLastName.Text + "�@" + textBoxFirstName.Text;
            // �d�b�ԍ�����
            telephoneNumber = textBoxTelephoneNumber.Text;
            // �Z������
            address = textBoxAddress.Text;
            // �c�̖�����
            organizationName = textBoxOrganizationName.Text;
            // �g�p�ړI����
            purposeOfUse = textBoxPurposeOfUse.Text;
            // �g�p�l������
            numberOfPeople = textBoxNumberOfPeople.Text;
            // �g�p������
            dateOfUse = dateTimePickerDateOfUse.Value.ToString();
            // �J�n�����i���j����
            startTime = comboBoxStartTimeHour.Text;
            // �J�n�����i���j����
            startMinutes = comboBoxStartTimeMinutes.Text;
            // �I�������i���j����
            endTime = comboBoxEndTimeHour.Text;
            // �I�������i���j����
            endMinutes = comboBoxEndTimeMinutes.Text;


            //�������v����
            // �g�p��������
            //var roomNameList = new List<CheckBox>();
            //var roomNameList = new List<Control>();
            //�����������t���Œǉ�����
            //����c���@�Ɖ�c���A���ǉ����ꂽ�ꍇ�A��c���@�A�ɂ���
            //�������̕������I������Ă���ꍇ�A��_�ŋ�؂�悤�ɂ���
            foreach (CheckBox cb in groupBoxRoomName.Controls.OfType<CheckBox>())
            {
                AddToList list = new AddToList();

                if (cb.Checked)
                {
                    if (cb.Text == "���̑�")
                    {
                        roomName += list.TextFormatting(roomName, textBoxOtherRooms.Text);
                    }
                    else
                    {
                        roomName += list.TextFormatting(roomName, cb.Text);
                    }
                }
            }
            // �g�p�������m�F
            //if (checkBoxConferenceRoom1.Checked && checkBoxConferenceRoom2.Checked)
            //{
            //    MessageBox.Show("��c���@�A");
            //}
            //if (checkBoxJapaneseStyleRoom1.Checked && checkBoxJapaneseStyleRoom2.Checked)
            //{
            //    MessageBox.Show("�a���@�A");
            //}



            // �g�p�����ݔ�����
            foreach (CheckBox cb in groupBoxAuxiliaryEquipmentUsed.Controls.OfType<CheckBox>())
            {
                AddToList list = new AddToList();

                if (cb.Checked)
                {
                    if (cb.Text == "���̑�")
                    {
                        auxiliaryEquipmentUsed += list.TextFormatting(auxiliaryEquipmentUsed, textBoxOtherEquipment.Text);
                    }
                    else
                    {
                        auxiliaryEquipmentUsed += list.TextFormatting(auxiliaryEquipmentUsed, cb.Text);
                    }
                }
            }

            // �g�p���̖Ə��\������
            reasonForApplyingForExemption = comboBoxReasonForApplyingForExemption.Text;


            //MessageBox.Show(applicationDate);
            //MessageBox.Show(name);
            //MessageBox.Show(telephoneNumber);
            //MessageBox.Show(address);
            //MessageBox.Show(organizationName);
            //MessageBox.Show(purposeOfUse);
            //MessageBox.Show(numberOfPeople);
            //MessageBox.Show(dateOfUse);
            //MessageBox.Show(startTime);
            //MessageBox.Show(startMinutes);
            //MessageBox.Show(endTime);
            //MessageBox.Show(endMinutes);
            MessageBox.Show(roomName);
            MessageBox.Show(auxiliaryEquipmentUsed);
            MessageBox.Show(reasonForApplyingForExemption);


            //if (int.Parse(comboBoxApplicationYear.Text) < 10)
            //{
            //    applicationYear = "�@" + int.Parse(comboBoxApplicationYear.Text);
            //}
            //else
            //{
            //    applicationYear = comboBoxApplicationYear.Text;
            //}

            //if (int.Parse(comboBoxApplicationMonth.Text) < 10)
            //{
            //    applicationMonth = "�@" + int.Parse(comboBoxApplicationMonth.Text);
            //}
            //else
            //{
            //    applicationMonth = comboBoxApplicationMonth.Text;
            //}

            //if (int.Parse(comboBoxApplicationDate.Text) < 10)
            //{
            //    applicationDate = "�@" + int.Parse(comboBoxApplicationDate.Text);
            //}
            //else
            //{
            //    applicationDate = comboBoxApplicationDate.Text;
            //}

            //if (int.Parse(comboBoxYearOfUse.Text) < 10)
            //{
            //    yearOfUse = "�@" + int.Parse(comboBoxYearOfUse.Text);
            //}
            //else
            //{
            //    yearOfUse = comboBoxYearOfUse.Text;
            //}

            //if (int.Parse(comboBoxMonthOfUse.Text) < 10)
            //{
            //    monthOfUse = "�@" + int.Parse(comboBoxMonthOfUse.Text);
            //}
            //else
            //{
            //    monthOfUse = comboBoxMonthOfUse.Text;
            //}

            //if (int.Parse(comboBoxDateOfUse.Text) < 10)
            //{
            //    dateOfUse = "�@" + int.Parse(comboBoxDateOfUse.Text);
            //}
            //else
            //{
            //    dateOfUse = comboBoxDateOfUse.Text;
            //}

            //if (int.Parse(comboBoxStartTime.Text) < 10)
            //{
            //    startTime = "�@" + int.Parse(comboBoxStartTime.Text);
            //}
            //else
            //{
            //    startTime = comboBoxStartTime.Text;
            //}


            //if (int.Parse(comboBoxEndTime.Text) < 10)
            //{
            //    endTime = "�@" + int.Parse(comboBoxEndTime.Text);
            //}
            //else
            //{
            //    endTime = comboBoxEndTime.Text;
            //}



            //if (checkBoxRoom1.Checked)
            //{
            //    rooms += checkBoxRoom1.Text;
            //}



            //if (checkBoxOtherEquipment.Checked)
            //{
            //    otherEquipment = "�`�F�b�N";
            //}
            //else
            //{
            //    otherEquipment = "�`�F�b�N�Ȃ�";
            //}
            //if (checkBoxFeeExemption.Checked)
            //{
            //    feeExemption = "�`�F�b�N";
            //}
            //else
            //{
            //    feeExemption = "�`�F�b�N�Ȃ�";
            //}

            //if (checkBoxFeeExemption.Checked && comboBoxReasonForExemption != null)
            //{
            //    reasonForExemption = comboBoxReasonForExemption.Text;
            //}
            //else
            //{
            //    reasonForExemption = " ";
            //}

            // �u������P����`
            //Dictionary<string, string> replaceWords = new Dictionary<string, string>()
            //{
            //    {"%AY%", applicationYear},
            //    {"%AM%", applicationMonth},
            //    {"%AD%", applicationDate},
            //    {"%LAST_NAME%", textBoxLastName.Text},
            //    {"%FIRST_NAME%", textBoxFirstName.Text},
            //    {"%TEL%", textBoxTEL.Text},
            //    {"%ADDRESS%", textBoxAddress.Text},
            //    {"%ORGANIZATION%", textBoxOrganization.Text},
            //    {"%PURPOSE%", textBoxPurpose.Text},
            //    {"%NoP%", textBoxNumberOfPeople.Text},
            //    {"%YoU%", yearOfUse},
            //    {"%MoU%", monthOfUse},
            //    {"%DoU%", dateOfUse},
            //    {"%ST%", startTime},
            //    {"%SM%", startMinutes},
            //    {"%ET%", endTime},
            //    {"%EM%", endMinutes},
            //    {"%ROOMS%", rooms},
            //    {"%AC%", airConditioner},
            //    {"%OE%", otherEquipment},
            //    {"%FE%", feeExemption},
            //    {"%RFE%", reasonForExemption},
            //};

            //���X�y�[�X�̐���v�C��
            //if (textBoxOtherEquipment.Text != null)
            //{
            //    replaceWords.Add("%OTHER_EQUIPMENT%", textBoxOtherEquipment.Text);
            //}
            //else
            //{
            //    replaceWords.Add("%OTHER_EQUIPMENT%", "");
            //}

            // Word �t�@�C��
            string wordFile = @"D:\dev\windows\src\repos\RentalApplicationCreationTool\RentalApplicationCreationTool\bin\Debug\net8.0-windows\template\mousikomisyo.docx";

            // Application ��錾����
            Word.Application app = null;

            // �h�L�������g�̃R���N�V������錾����
            Word.Documents documents = null;

            // �h�L�������g��錾����
            Word.Document document = null;

            try
            {
                // Application ���쐬����
                app = new Word.Application();

                // �h�L�������g�̃R���N�V�������擾����
                documents = app.Documents;

                // Word �̕����t�@�C�����J��
                document = documents.Open(wordFile);

                // �����ꂪ�Ȃ��� missing ���Ȃ��Ď��s�ł��Ȃ�
                //object missing = Type.Missing;

                Word.Find findObject = app.Selection.Find;
                /*
                foreach (var replaceWord in replaceWords)
                {
                    findObject.ClearFormatting();
                    findObject.Text = replaceWord.Key;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = replaceWord.Value;

                    // ���Y���ӏ��@�ꂩ���̂ݒu��������悤�ɕύX����
                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    findObject.Execute(Replace: replaceAll);
                }
                */

                // �\������
                app.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        // Word�̃C���X�^���X���쐬
        //Word.Application app;
        // Word�A�v���P�[�V�����I�u�W�F�N�g���쐬
        //app = new Word.Application();

        // �\������
        //app.Visible = true;

            //Documents �R���N�V������Add ���\�b�h���g�p���āA
            //Normal.dot �Ɋ�Â��V�����������쐬���܂��B
            /*
            object missingValue = Type.Missing;
            app.Documents.Add(ref missingValue, ref missingValue,
                ref missingValue, ref missingValue);
            */

            // �e���v���[�g���J��
            // �����΃p�X�Ŏw�肷��
            // ��Word�e���v���[�g���J���ق����ǂ�����
            //app.Documents.Open(@"C:\Users\nora\source\repos\RentalFacility\RentalFacility\bin\Debug\template\mousikomisyo.docx");

            // �f�[�^����������

        private void buttonUserList_Click(object sender, EventArgs e)
        {
            FormUserList formUserList = new FormUserList();
            formUserList.ShowDialog();
        }
    }
}
