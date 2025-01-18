using Word = Microsoft.Office.Interop.Word;

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
            string applicationDate;     // 申請日
            string name;                // 名前
            string telephoneNumber;     // 電話番号
            string address;             // 住所
            string organizationName;    // 団体名
            string purposeOfUse;        // 使用目的
            string numberOfPeople;         // 使用人数
            string dateOfUse;           // 使用日
            string startTime;           // 開始時刻（時）
            string startMinutes;        // 開始時刻（分）
            string endTime;             // 終了時刻（時）
            string endMinutes;          // 終了時刻（分）
            string roomName = "";       // 使用室名
            string auxiliaryEquipmentUsed = "";     // 使用附属設備
            string reasonForApplyingForExemption;   // 使用料の免除申請



            // 申請日を代入
            applicationDate = dateTimePickerApplicationDate.Value.ToString();
            // 名前を代入
            name = textBoxLastName.Text + "　" + textBoxFirstName.Text;
            // 電話番号を代入
            telephoneNumber = textBoxTelephoneNumber.Text;
            // 住所を代入
            address = textBoxAddress.Text;
            // 団体名を代入
            organizationName = textBoxOrganizationName.Text;
            // 使用目的を代入
            purposeOfUse = textBoxPurposeOfUse.Text;
            // 使用人数を代入
            numberOfPeople = textBoxNumberOfPeople.Text;
            // 使用日を代入
            dateOfUse = dateTimePickerDateOfUse.Value.ToString();
            // 開始時刻（時）を代入
            startTime = comboBoxStartTimeHour.Text;
            // 開始時刻（分）を代入
            startMinutes = comboBoxStartTimeMinutes.Text;
            // 終了時刻（時）を代入
            endTime = comboBoxEndTimeHour.Text;
            // 終了時刻（分）を代入
            endMinutes = comboBoxEndTimeMinutes.Text;


            //★★★要理解
            // 使用室名を代入
            //var roomNameList = new List<CheckBox>();
            //var roomNameList = new List<Control>();
            //★部屋名を逆順で追加する
            //★会議室①と会議室②が追加された場合、会議室①②にする
            //★複数の部屋が選択されている場合、句点で区切るようにする
            foreach (CheckBox cb in groupBoxRoomName.Controls.OfType<CheckBox>())
            {
                AddToList list = new AddToList();

                if (cb.Checked)
                {
                    if (cb.Text == "その他")
                    {
                        roomName += list.TextFormatting(roomName, textBoxOtherRooms.Text);
                    }
                    else
                    {
                        roomName += list.TextFormatting(roomName, cb.Text);
                    }
                }
            }
            // 使用室名を確認
            //if (checkBoxConferenceRoom1.Checked && checkBoxConferenceRoom2.Checked)
            //{
            //    MessageBox.Show("会議室①②");
            //}
            //if (checkBoxJapaneseStyleRoom1.Checked && checkBoxJapaneseStyleRoom2.Checked)
            //{
            //    MessageBox.Show("和室①②");
            //}



            // 使用附属設備を代入
            foreach (CheckBox cb in groupBoxAuxiliaryEquipmentUsed.Controls.OfType<CheckBox>())
            {
                AddToList list = new AddToList();

                if (cb.Checked)
                {
                    if (cb.Text == "その他")
                    {
                        auxiliaryEquipmentUsed += list.TextFormatting(auxiliaryEquipmentUsed, textBoxOtherEquipment.Text);
                    }
                    else
                    {
                        auxiliaryEquipmentUsed += list.TextFormatting(auxiliaryEquipmentUsed, cb.Text);
                    }
                }
            }

            // 使用料の免除申請を代入
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
            //    applicationYear = "　" + int.Parse(comboBoxApplicationYear.Text);
            //}
            //else
            //{
            //    applicationYear = comboBoxApplicationYear.Text;
            //}

            //if (int.Parse(comboBoxApplicationMonth.Text) < 10)
            //{
            //    applicationMonth = "　" + int.Parse(comboBoxApplicationMonth.Text);
            //}
            //else
            //{
            //    applicationMonth = comboBoxApplicationMonth.Text;
            //}

            //if (int.Parse(comboBoxApplicationDate.Text) < 10)
            //{
            //    applicationDate = "　" + int.Parse(comboBoxApplicationDate.Text);
            //}
            //else
            //{
            //    applicationDate = comboBoxApplicationDate.Text;
            //}

            //if (int.Parse(comboBoxYearOfUse.Text) < 10)
            //{
            //    yearOfUse = "　" + int.Parse(comboBoxYearOfUse.Text);
            //}
            //else
            //{
            //    yearOfUse = comboBoxYearOfUse.Text;
            //}

            //if (int.Parse(comboBoxMonthOfUse.Text) < 10)
            //{
            //    monthOfUse = "　" + int.Parse(comboBoxMonthOfUse.Text);
            //}
            //else
            //{
            //    monthOfUse = comboBoxMonthOfUse.Text;
            //}

            //if (int.Parse(comboBoxDateOfUse.Text) < 10)
            //{
            //    dateOfUse = "　" + int.Parse(comboBoxDateOfUse.Text);
            //}
            //else
            //{
            //    dateOfUse = comboBoxDateOfUse.Text;
            //}

            //if (int.Parse(comboBoxStartTime.Text) < 10)
            //{
            //    startTime = "　" + int.Parse(comboBoxStartTime.Text);
            //}
            //else
            //{
            //    startTime = comboBoxStartTime.Text;
            //}


            //if (int.Parse(comboBoxEndTime.Text) < 10)
            //{
            //    endTime = "　" + int.Parse(comboBoxEndTime.Text);
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
            //    otherEquipment = "チェック";
            //}
            //else
            //{
            //    otherEquipment = "チェックなし";
            //}
            //if (checkBoxFeeExemption.Checked)
            //{
            //    feeExemption = "チェック";
            //}
            //else
            //{
            //    feeExemption = "チェックなし";
            //}

            //if (checkBoxFeeExemption.Checked && comboBoxReasonForExemption != null)
            //{
            //    reasonForExemption = comboBoxReasonForExemption.Text;
            //}
            //else
            //{
            //    reasonForExemption = " ";
            //}

            // 置換する単語を定義
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

            //★スペースの数を要修正
            //if (textBoxOtherEquipment.Text != null)
            //{
            //    replaceWords.Add("%OTHER_EQUIPMENT%", textBoxOtherEquipment.Text);
            //}
            //else
            //{
            //    replaceWords.Add("%OTHER_EQUIPMENT%", "");
            //}

            // Word ファイル
            string wordFile = @"D:\dev\windows\src\repos\RentalApplicationCreationTool\RentalApplicationCreationTool\bin\Debug\net8.0-windows\template\mousikomisyo.docx";

            // Application を宣言する
            Word.Application app = null;

            // ドキュメントのコレクションを宣言する
            Word.Documents documents = null;

            // ドキュメントを宣言する
            Word.Document document = null;

            try
            {
                // Application を作成する
                app = new Word.Application();

                // ドキュメントのコレクションを取得する
                documents = app.Documents;

                // Word の文書ファイルを開く
                document = documents.Open(wordFile);

                // ★これがないと missing がなくて実行できない
                //object missing = Type.Missing;

                Word.Find findObject = app.Selection.Find;
                /*
                foreach (var replaceWord in replaceWords)
                {
                    findObject.ClearFormatting();
                    findObject.Text = replaceWord.Key;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = replaceWord.Value;

                    // ★該当箇所　一か所のみ置き換えるように変更する
                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    findObject.Execute(Replace: replaceAll);
                }
                */

                // 表示する
                app.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        // Wordのインスタンスを作成
        //Word.Application app;
        // Wordアプリケーションオブジェクトを作成
        //app = new Word.Application();

        // 表示する
        //app.Visible = true;

            //Documents コレクションのAdd メソッドを使用して、
            //Normal.dot に基づく新しい文書を作成します。
            /*
            object missingValue = Type.Missing;
            app.Documents.Add(ref missingValue, ref missingValue,
                ref missingValue, ref missingValue);
            */

            // テンプレートを開く
            // ★相対パスで指定する
            // ★Wordテンプレートを開くほうが良いかも
            //app.Documents.Open(@"C:\Users\nora\source\repos\RentalFacility\RentalFacility\bin\Debug\template\mousikomisyo.docx");

            // データを差し込む

        private void buttonUserList_Click(object sender, EventArgs e)
        {
            FormUserList formUserList = new FormUserList();
            formUserList.ShowDialog();
        }
    }
}
