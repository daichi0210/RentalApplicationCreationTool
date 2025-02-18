﻿using Microsoft.VisualBasic;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Globalization;
using System.Windows.Forms;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Word = Microsoft.Office.Interop.Word;

namespace RentalApplicationCreationTool
{
    public partial class FormMain : Form
    {
        public FormMain()
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
            string numberOfPeople;      // 使用人数
            string dateOfUse;           // 使用日
            string dayOfWeek;           // 使用日の曜日
            string startTime;           // 開始時刻（時）
            string startMinutes;        // 開始時刻（分）
            string endTime;             // 終了時刻（時）
            string endMinutes;          // 終了時刻（分）
            string roomName = "";       // 使用室名
            string airConditioner;      // 使用附属設備（冷暖房）
            string otherEquipment; // 使用附属設備（その他）
            string auxiliaryEquipmentUsed;  // 使用附属設備（その他の内容）
            string exemptionApplication;    // 使用料の免除申請
            string reasonForApplyingForExemption;   // 使用料の免除申請理由

            // 入力された値の確認
            if (inputConfirmation() == false)
            {
                return;
            }

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
            // 使用日の曜日を代入
            dayOfWeek = dateTimePickerDateOfUse.Value.ToString("ddd");
            // 開始時刻（時）を代入
            startTime = comboBoxStartTimeHour.Text;
            // 開始時刻（分）を代入
            startMinutes = comboBoxStartTimeMinutes.Text;
            // 終了時刻（時）を代入
            endTime = comboBoxEndTimeHour.Text;
            // 終了時刻（分）を代入
            endMinutes = comboBoxEndTimeMinutes.Text;
            // 使用室名を代入
            // 部屋名をTabIndex順にするため、foreachを逆順にしている
            //★★★要理解
            foreach (CheckBox cb in Enumerable.Reverse(groupBoxRoomName.Controls.OfType<CheckBox>()))
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
            // 使用附属設備（冷暖房）の状態
            if (checkBoxAirConditioningAndHeating.Checked)
            {
                airConditioner = "☑";
            }
            else
            {
                airConditioner = "□";
            }
            // 使用附属設備（その他）の状態
            if (checkBoxOtherEquipment.Checked)
            {
                otherEquipment = "☑";
            }
            else
            {
                otherEquipment = "□";
            }
            // 使用附属設備（その他の内容）
            auxiliaryEquipmentUsed = textBoxOtherEquipment.Text;
            // 使用料の免除申請の状態
            if (comboBoxReasonForApplyingForExemption.Text != "")
            {
                exemptionApplication = "☑";
            }
            else
            {
                exemptionApplication = "□";
            }
            // 使用料の免除申請を代入
            reasonForApplyingForExemption = comboBoxReasonForApplyingForExemption.Text;



            // 入力内容の確認
            //★もう少しスマートにできると思う
            contentConfirmation(dateOfUse, applicationDate, startTime, startMinutes, endTime, endMinutes, roomName, auxiliaryEquipmentUsed);




            // 年を和暦で表示するための準備
            CultureInfo Japanese = new CultureInfo("ja-JP");
            Japanese.DateTimeFormat.Calendar = new JapaneseCalendar();

            // 置換する単語を定義
            Dictionary<string, string> replaceWords = new Dictionary<string, string>()
            {
                {"%AY%", dateTimePickerApplicationDate.Value.ToString("%y", Japanese)},
                {"%AM%", dateTimePickerApplicationDate.Value.ToString("%M")},
                {"%AD%", dateTimePickerApplicationDate.Value.ToString("%d")},
                {"%NAME%", name},
                {"%TEL%", telephoneNumber},
                {"%ADDRESS%", address},
                {"%ORGANIZATION%", organizationName},
                {"%PURPOSE%", purposeOfUse},
                {"%NoP%", numberOfPeople},
                // ★要修正。dayOfWeekのようにならないか？
                {"%YoU%", dateTimePickerDateOfUse.Value.ToString("%y", Japanese)},
                {"%MoU%", dateTimePickerDateOfUse.Value.ToString("%M")},
                {"%DoU%", dateTimePickerDateOfUse.Value.ToString("%d")},
                {"%DoW%", dayOfWeek},
                {"%ST%", startTime},
                {"%SM%", startMinutes},
                {"%ET%", endTime},
                {"%EM%", endMinutes},
                {"%ROOM%", roomName},
                {"%AC%", airConditioner},
                {"%OE%", otherEquipment},
                {"%OEU%", auxiliaryEquipmentUsed},
                {"%EA%", exemptionApplication},
                {"%RFE%", reasonForApplyingForExemption},
            };

            string fileName = DateTime.Now.ToString("yyyy-MM-dd-HHmmss") + "-" + organizationName + ".docx";

            // ★Wordテンプレートを開くほうが良いかも
            // テンプレートファイルをコピー
            File.Copy("template\\_mousikomisyo.docx", "application forms\\" + fileName);

            // 現在のディレクトリを取得
            string currentDirectory = Directory.GetCurrentDirectory();

            // 現在のディレクトリからの相対パスを作成
            // Word ファイル
            string wordFile = Path.Combine(currentDirectory, "application forms\\" + fileName);

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

                Word.Find findObject = app.Selection.Find;

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

                // 表示する
                app.Visible = true;

                // 生成したWrod文書を上書き保存する
                app.ActiveDocument.Save();

                // 印刷設定
                object copies = "1";
                object pages = "";
                object range = Word.WdPrintOutRange.wdPrintAllDocument;
                object items = Word.WdPrintOutItem.wdPrintDocumentContent;
                object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                object oTrue = true;
                object oFalse = false;

                // 印刷する
                /*
                app.PrintOut(
                    Background: oTrue,
                    Append: oFalse,
                    Range: range,
                    Item: items,
                    Copies: copies,
                    Pages: pages,
                    PageType: pageType,
                    PrintToFile: oFalse,
                    Collate: oTrue,
                    ManualDuplexPrint: oFalse
                );
                */
        }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool inputConfirmation()
        {
            // 苗字または名前が未入力の場合
            if (textBoxLastName.Text == "")
            {
                MessageBox.Show("苗字を入力してください。");
                return false;
            }
            else if (textBoxFirstName.Text == "")
            {
                MessageBox.Show("名前を入力してください。");
                return false;
            }

            // 電話番号が未入力の場合
            if (textBoxTelephoneNumber.Text == "")
            {
                MessageBox.Show("電話番号を入力してください。");
                return false;
            }

            // 住所が未入力の場合
            if (textBoxAddress.Text == "")
            {
                MessageBox.Show("住所を入力してください。");
                return false;
            }

            // 団体名が未入力の場合
            if (textBoxOrganizationName.Text == "")
            {
                MessageBox.Show("団体名を入力してください。");
                return false;
            }
            // 使用目的が未入力の場合
            if (textBoxPurposeOfUse.Text == "")
            {
                MessageBox.Show("使用目的を入力してください。");
                return false;
            }
            // 使用人数が未入力の場合
            //★数字のみにする
            if (textBoxNumberOfPeople.Text == "")
            {
                MessageBox.Show("使用人数を入力してください。");
                return false;
            }
            // 開始時刻（時）が未選択の場合
            if (comboBoxStartTimeHour.Text == "")
            {
                MessageBox.Show("開始時刻（時）を選択してください。");
                return false;
            }
            // 開始時刻（分）が未選択の場合
            if (comboBoxStartTimeMinutes.Text == "")
            {
                MessageBox.Show("開始時刻（分）を選択してください。");
                return false;
            }
            // 終了時刻（時）が未選択の場合
            if (comboBoxEndTimeHour.Text == "")
            {
                MessageBox.Show("終了時刻（時）を選択してください。");
                return false;
            }
            // 終了時刻（分）が未選択の場合
            if (comboBoxEndTimeMinutes.Text == "")
            {
                MessageBox.Show("終了時刻（分）を選択してください。");
                return false;
            }

            // その他が選択されているが部屋名が未入力の場合
            if (checkBoxOtherRooms.Checked && textBoxOtherRooms.Text == "")
            {
                MessageBox.Show("その他の部屋名を入力してください。");
                return false;
            }

            return true;
        }

        private bool contentConfirmation(string dateOfUse, string applicationDate, string startTime, string startMinutes, string endTime, string endMinutes, string roomName, string auxiliaryEquipmentUsed)
        {
            // 使用日が申請日よりも前の場合
            //★たまに正常に動作しない？
            if (dateOfUse.CompareTo(applicationDate) == -1)
            {
                MessageBox.Show("申請日よりも後の日付を選択してください。");
                return false;
            }

            // 終了時刻が開始時刻と同じまたは終了時刻が開始時刻よりも前の場合
            DateTime now = DateTime.Now;
            DateTime startDate = new DateTime(now.Year, now.Month, now.Day, Int32.Parse(startTime), Int32.Parse(startMinutes), now.Second);
            DateTime endDate = new DateTime(now.Year, now.Month, now.Day, Int32.Parse(endTime), Int32.Parse(endMinutes), now.Second);
            
            switch (startDate.CompareTo(endDate))
            {
                case 0:
                case 1:
                    MessageBox.Show("終了時刻は開始時刻よりも後にしてください。");
                    return false;
            }

            // 使用する部屋が選択されていない場合
            if (roomName == "")
            {
                MessageBox.Show("使用する部屋を選択してください。");
                return false;
            }

            // 使用附属設備（その他）が選択されているが内容が未入力の場合
            if (checkBoxOtherEquipment.Checked && auxiliaryEquipmentUsed == "")
            {
                MessageBox.Show("その他の設備を入力してください。");
                return false;
            }

            return true;
        }
        private void buttonUserList_Click(object sender, EventArgs e)
        {
            this.Hide();

            FormUserList formUserList = new FormUserList(this);
            formUserList.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("使用者の登録が完了しました。");
        }
    }
}
