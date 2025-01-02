using System.Text.RegularExpressions;
using System.Windows.Forms;
using TranslationServer;
using System.Drawing;
using static System.Net.Mime.MediaTypeNames;
using System.Text;

namespace SourceTranlate
{

    public partial class Form1 : Form
    {

        private string[] fileLines;
        private Dictionary<string, TransItem> _transItems;

        public Form1()
        {
            InitializeComponent();
            Database.Initialize();

            /*
            var newDB = new List<Schema>();
            foreach (var item in Database.db)
            {
                if (item.Translation != null)
                {
                    newDB.Add(item);
                }
            }
            Database.db = newDB;
            Database.SaveDb();
            */

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Text = $"SourceTrans";
            _transItems = new Dictionary<string, TransItem>();
            listBox1.Items.Clear();

            string fileContent = "";

            //openFileDialog.InitialDirectory = Application.StartupPath;
            openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                string filePath = openFileDialog.FileName;

                fileLines = File.ReadAllLines(filePath, System.Text.Encoding.GetEncoding("iso-8859-1"));

                int cached = 0;

                if (fileLines.Length > 0)
                {
                    //fileLines = fileContent.Split(Environment.NewLine);
                    for (int i = 0; i < fileLines.Length; i++)
                    {
                        Regex regex = new Regex("\"(.*?)\"");

                        var matches = regex.Matches(fileLines[i]);

                        if (matches.Count > 0)
                        {
                            for (int ii = 0; ii < matches.Count; ii++)
                            {
                                if (SkipTranslation(matches[ii].Value)) { continue; }

                                var trans = GetTranslation(matches[ii].Value, true);

                                if (trans != null)
                                {
                                    AddTranslation($"{i}", trans, matches[ii].Value);
                                    listBox1.Items.Add($"{i}-{ii}: {trans}");
                                    cached++;
                                }
                                else
                                {
                                    listBox1.Items.Add($"{i}-{ii}: {matches[ii]}");
                                }

                            }
                            if (cached > 0)
                                this.Text = $"SourceTrans - Found {cached}";
                        }
                    }
                }

            }
        }

        private bool SkipTranslation(string value)
        {
            switch (value)
            {

                case "\"Georgia\"":
                    return true;

                case "\"Verdana\"":
                    return true;

                case "\"Tahoma\"":
                    return true;

                default:
                    return false;
                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Get the path of specified file
            string filePath = openFileDialog.FileName;

            foreach (var item in _transItems)
            {
                string line = fileLines[int.Parse(item.Value.line.Split("-")[0])];

                line = line.Replace(item.Value.original, item.Value.translation);
                fileLines[int.Parse(item.Value.line.Split("-")[0])] = line;
            }
            //File.Copy(filePath, filePath + ".bak", true);
            File.WriteAllLines(filePath, fileLines);
            MessageBox.Show("Done!");

        }

        private void AddTranslation(string line, string translation, string original)
        {
            if (_transItems == null) _transItems = new Dictionary<string, TransItem>();

            if (_transItems.ContainsKey(line)) _transItems.Remove(line);

            _transItems.Add(line, new TransItem(line, translation, original));

        }

        private String GetTranslation(string text, bool skipTranslate = false)
        {
            string md5 = Translator.CreateMD5(text);
            string md5Old = Translator.CreateMD5(md5, true);
            string cache = Database.GetTranslation(md5, text);
            string translation = "";

            if (cache == null)
            {
                cache = Database.GetTranslation(md5Old, text);

                if (cache == null && !skipTranslate)
                {
                    translation = TranslationServer.Translator.Yandex(text).Result;
                    translation = FixSpace(translation, text);
                    if (translation.Length > 0) Database.AddTranslation(md5, translation, text);
                    Database.SaveDb();
                    return translation;
                }
                else
                {
                    Database.RemoveTranslation(md5Old);
                    if (cache != null) Database.AddTranslation(md5, FixSpace(cache, text), text);
                    return cache;
                }
            }

            return cache;
        }

        private string FixSpace(string translation, string text)
        {
            if (translation != null & text.Substring(0, 2) == "\" ")
            {
                if (translation.Substring(0, 2) != "\" ")
                {
                    return translation.Insert(1, " ");
                }
            }
            return translation;
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var item = listBox1.Text.Split(": ");
            string line = item[0].Split("-")[0];
            string key = item[0];
            string original = string.Join(": ", item[1..]);

            var translation = GetTranslation(original);

            listBox1.Items[listBox1.SelectedIndex] = $"{key}: {translation}";

            label1.Text = "Translated: " + translation;

            AddTranslation(key, translation, original);

        }

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                var item = listBox1.Text.Split(": ");
                string line = item[0].Split("-")[0];
                string key = item[0];
                string original = string.Join(": ", item[1..]);
                string translation = original;

                if (_transItems.ContainsKey(key)) translation = _transItems[key].translation;
                if (_transItems.ContainsKey(key)) original = _transItems[key].original;

                if (e.Button == MouseButtons.Right)
                {
                    var d = InputBox("Custom Translation", original, ref translation);

                    if (d == DialogResult.OK)
                    {
                        string md5 = Translator.CreateMD5(original);
                        listBox1.Items[listBox1.SelectedIndex] = $"{key}: {translation}";
                        Database.RemoveTranslation(md5);
                        Database.AddTranslation(md5, translation, original);
                        AddTranslation(key, translation, original);
                    }
                }
            }
            catch { }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Database.SaveDb();
        }

    }
    public class TransItem
    {
        public string line;
        public string translation;
        public string original;

        public TransItem(string line, string translation, string original)
        {
            this.line = line;
            this.translation = translation;
            this.original = original;
        }
    }
}