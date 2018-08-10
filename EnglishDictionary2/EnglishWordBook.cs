namespace EnglishDictionary2
{
    using HtmlAgilityPack;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Drawing;
    using System.IO;
    using System.Net;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;
    using System.Linq;

    public class EnglishWordBook : Form
    {
        private HtmlAgilityPack.HtmlDocument hdHtml = null;
        private HtmlNode hnContentNode = null;
        private static int snWordCnt = 0;
        private IContainer components = null;
        private DataGridView dgResult;
        private System.Windows.Forms.TextBox tbWordList;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Button btnExcelExport;
        private System.Windows.Forms.Label lbWordList;
        private System.Windows.Forms.Label label1;
        private DataGridViewTextBoxColumn Number;
        private DataGridViewTextBoxColumn Vocabulary;
        private DataGridViewTextBoxColumn Pronunciation;
        private System.Windows.Forms.TextBox tbBack;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbFront;

        public EnglishWordBook()
        {
            this.InitializeComponent();
        }

        private void btnExcelExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "EnglishWordList_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.ExportToExcel(this.dgResult, dialog.FileName);
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            this.dgResult.DataSource = null;
            this.dgResult.Rows.Clear();
            this.dgResult.Refresh();
            snWordCnt = 0;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string[] strArray = Regex.Split(this.tbWordList.Text, Environment.NewLine);
            if (strArray.Length <= 0)
            {
                MessageBox.Show("단어 리스트를 입력해 주세요.");
                this.tbWordList.Focus();
            }
            else
            {
                int aiNum = 0;
                foreach (string str2 in strArray)
                {
                    aiNum++;
                    string asTargetURL = "http://endic.naver.com/search.nhn?query=" + str2 + "&searchOption=word&isOnlyViewEE=Y";
                    Vocabulary oneVocabulary = this.GetOneVocabulary(asTargetURL, aiNum);

                    if (!((oneVocabulary == null) || string.IsNullOrEmpty(oneVocabulary.NAME.Trim())))
                    {
                        snWordCnt++;
                        this.dgResult.Rows.Add(new object[] { snWordCnt, oneVocabulary.FRONT, oneVocabulary.BACK });
                        this.dgResult.Refresh();
                    }
                    else if (this.hnContentNode != null)
                    {
                        HtmlNode node = this.hnContentNode.SelectSingleNode("//div[contains(@class,'word_num')]/dl[1]/dt[1]/span[1]/a[contains(@href,'enenEntry.nhn?entryId=')][1]");
                        if (node != null)
                        {
                            HtmlAttribute attribute = node.Attributes["href"];
                            if ((attribute != null) && !string.IsNullOrEmpty(attribute.Value))
                            {
                                oneVocabulary = this.GetOneVocabulary("http://endic.naver.com" + attribute.Value, aiNum);
                                if (!((oneVocabulary == null) || string.IsNullOrEmpty(oneVocabulary.NAME.Trim())))
                                {
                                    snWordCnt++;
                                    this.dgResult.Rows.Add(new object[] { snWordCnt, oneVocabulary.FRONT, oneVocabulary.BACK });
                                    this.dgResult.Refresh();
                                }
                            }
                        }
                    }
                }
            }
        }

        private void dgResult_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            object frontObj = this.dgResult.Rows[e.RowIndex].Cells[1].Value;
            object backObj = this.dgResult.Rows[e.RowIndex].Cells[2].Value;

            if (!((frontObj == null) || string.IsNullOrEmpty(frontObj.ToString())))
            {
                this.tbFront.Text = frontObj.ToString();
            }
            else
            {
                this.tbFront.Text = "";
            }

            if (!((backObj == null) || string.IsNullOrEmpty(backObj.ToString())))
            {
                this.tbBack.Text = backObj.ToString();
            }
            else
            {
                this.tbBack.Text = "";
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void ExportToExcel(DataGridView oDgv, string filename)
        {
            int num2;
            int num = 0;
            object missing = System.Type.Missing;
            string[] strArray = new string[oDgv.ColumnCount];
            string[] strArray2 = new string[oDgv.ColumnCount];
            for (num2 = 0; num2 < oDgv.ColumnCount; num2++)
            {
                strArray[num2] = oDgv.Rows[0].Cells[num2].OwningColumn.HeaderText.ToString();
                strArray[num2] = oDgv.Rows[0].Cells[num2].OwningColumn.HeaderText.ToString();
                num = num2 + 0x41;
                strArray2[num2] = Convert.ToString((char) num);
            }
            try
            {
                Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application) Activator.CreateInstance(System.Type.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
                _Workbook workbook = application.Workbooks.Add(Missing.Value);
                Sheets worksheets = workbook.Worksheets;
                _Worksheet worksheet = (_Worksheet) worksheets.get_Item(1);
                for (num2 = 0; num2 < oDgv.ColumnCount; num2++)
                {
                    worksheet.get_Range(strArray2[num2] + "1", Missing.Value).set_Value(Missing.Value, strArray[num2]);
                }
                strArray2 = new string[oDgv.ColumnCount];
                for (int i = 0; i < (oDgv.RowCount - 1); i++)
                {
                    if (!string.IsNullOrEmpty(oDgv.Rows[i].Cells[0].Value.ToString()))
                    {
                        for (int j = 0; j < oDgv.ColumnCount; j++)
                        {
                            num = j + 0x41;
                            strArray2[j] = Convert.ToString((char) num);
                            worksheet.get_Range(strArray2[j] + Convert.ToString((int) (i + 2)), Missing.Value).set_Value(Missing.Value, oDgv.Rows[i].Cells[j].Value.ToString());
                        }
                    }
                }
                application.Visible = false;
                application.UserControl = false;
                workbook.SaveAs(filename, XlFileFormat.xlWorkbookNormal, missing, missing, missing, missing, XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                workbook.Close(false, missing, missing);
                Cursor.Current = Cursors.Default;
                MessageBox.Show("저장이 완료되었습니다.");
            }
            catch (Exception exception)
            {
                MessageBox.Show(("Error: " + exception.Message) + " Line: " + exception.Source, "Error");
            }
        }


        // 영문 뜻 텍스트 획득 함수
        private Vocabulary GetMeaningAndSentences(Vocabulary vocabulary, HtmlNodeCollection meaningListNodeCollection)
        {
            List<string> meanings = null;
            List<List<string>> sentencesList = null;
            List<string> sentences = null;

            if (vocabulary == null)
            {
                meanings = new List<string>();
                sentencesList = new List<List<string>>();
                sentences = null;
            }
            else
            {
                meanings = vocabulary.getMeanings();
                sentencesList = vocabulary.getSentencesList();
                sentences = null;
            }

            foreach (HtmlNode curListNode in meaningListNodeCollection)
            {                
                HtmlNode curNode = curListNode.SelectSingleNode("./*[1]");
                while (curNode != null)
                {
                    string newLine = string.Empty;

                    // 뜻("dt") 태그 인 경우
                    if (curNode.OriginalName.Equals("dt"))
                    {
                        HtmlNodeCollection meaningNodes = curNode.SelectNodes("./*/span[contains(@class, 'fnt')]");

                        string curMeaningStr = string.Empty;
                        if (meaningNodes != null)
                        {
                            foreach (HtmlNode meaningNode in meaningNodes)
                            {
                                curMeaningStr += this.MyReplace(meaningNode.InnerText).Replace("\r\n", "") + " ";
                            }
                            meanings.Add(curMeaningStr);
                        }
                        else
                        {
                            meanings.Add(this.MyReplace(curNode.InnerText).Replace("\r\n", ""));
                        }
                        sentencesList.Insert(meanings.Count - 1, new List<string>());
                    }
                    // 예문, 동의어, 반의어("dd") 태그 인 경우
                    else if (curNode.OriginalName.Equals("dd"))
                    {
                        sentences = sentencesList[meanings.Count - 1];
                        //string curSentenceStr = string.Empty;

                        HtmlNodeCollection sentenceNodes = curNode.SelectNodes("./*/*/span[contains(@class, 'fnt')]");
                        if (sentenceNodes != null)
                        {
                            foreach (HtmlNode sentenceNode in sentenceNodes)
                            {
                                //curSentenceStr += this.MyReplace(sentenceNode.InnerText).Replace("\r\n", "");
                                sentences.Add(this.MyReplace(sentenceNode.InnerText).Replace("\r\n", ""));
                            }
                        }
                        //sentences.Add(curSentenceStr);
                    }
                    curNode = curNode.NextSibling;
                }
            }

            if (vocabulary == null)
            {
                return new Vocabulary(meanings, sentencesList);
            }
            else
            {
                vocabulary.setMeanings(meanings);
                vocabulary.setSentencesList(sentencesList);
                return vocabulary;
            }
        }

        private EnglishDictionary2.Vocabulary GetOneVocabulary(string asTargetURL, int aiNum)
        {
            if (string.IsNullOrEmpty(asTargetURL))
            {
                return null;
            }
            this.hdHtml = this.LoadHtml(asTargetURL);

            this.hnContentNode = this.hdHtml.DocumentNode.SelectSingleNode("//*[@id='content']");
            if ( HtmlNodeUtil.isNullOrEmpty(this.hnContentNode) || (this.hnContentNode.InnerText.IndexOf("에 대한 검색결과가 없습니다.") >= 0))
            {
                return null;
            }

            // 
            HtmlNode ahnCurNode = this.hnContentNode.SelectSingleNode("//div[contains(@class, 'tit')]/h3");
            string name = string.Empty;
            if (!((ahnCurNode == null) || string.IsNullOrEmpty(ahnCurNode.InnerText)))
            {
                name = ahnCurNode.InnerText;
            }
            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            // 발음 기호 파싱
            String pronunciation = string.Empty;
            HtmlNode pronNode = this.hnContentNode.SelectSingleNode("//span[contains(@class, 'fnt_e16')]");

            if (!HtmlNodeUtil.isNullOrEmpty(pronNode))
            {
                // 강세 노드 탐색
                HtmlNode stressNode = pronNode.SelectSingleNode("./u[1]");

                // 강세가 있으면 강세 기호 삽입
                if( !HtmlNodeUtil.isNullOrEmpty(stressNode) )
                {
                    pronunciation = pronNode.InnerText.Insert(pronNode.InnerText.IndexOf(stressNode.InnerText), "'");
                }
                else
                {
                    pronunciation = pronNode.InnerText;
                }                
            }

            Vocabulary vocabulary = null;

            foreach (HtmlNode curNode in this.hnContentNode.SelectNodes("//div[@id='zoom_content']/div[contains(@class, 'box_wrap')]"))
            {
                //Console.WriteLine("box_wrap1 tag: " + curNode.InnerText);

                // etc 는 제외. 그외 품사나 PHRASE 는 포함.
                HtmlNode etcNode = curNode.SelectSingleNode("./h3[contains(@id, 'etc')]");
                if (!HtmlNodeUtil.isNullOrEmpty(etcNode))
                {
                    continue;
                }

                HtmlNodeCollection meaningListNodeCollection = curNode.SelectNodes("./dl[@class='list_a8']");
                if (meaningListNodeCollection != null && meaningListNodeCollection.Count > 0 )
                {
                    vocabulary = this.GetMeaningAndSentences(vocabulary, meaningListNodeCollection);
                }
            }
            
            if (vocabulary != null)
            {
                vocabulary.NAME = name;
                vocabulary.PRON = pronunciation;
            }

            return vocabulary;
        }

        private void InitializeComponent()
        {
            this.dgResult = new System.Windows.Forms.DataGridView();
            this.tbWordList = new System.Windows.Forms.TextBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnExcelExport = new System.Windows.Forms.Button();
            this.lbWordList = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tbFront = new System.Windows.Forms.TextBox();
            this.tbBack = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.Number = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Vocabulary = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Pronunciation = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgResult)).BeginInit();
            this.SuspendLayout();
            // 
            // dgResult
            // 
            this.dgResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgResult.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Number,
            this.Vocabulary,
            this.Pronunciation});
            this.dgResult.Location = new System.Drawing.Point(12, 247);
            this.dgResult.Name = "dgResult";
            this.dgResult.RowTemplate.Height = 23;
            this.dgResult.Size = new System.Drawing.Size(975, 305);
            this.dgResult.TabIndex = 4;
            this.dgResult.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgResult_CellClick);
            // 
            // tbWordList
            // 
            this.tbWordList.Location = new System.Drawing.Point(12, 30);
            this.tbWordList.Multiline = true;
            this.tbWordList.Name = "tbWordList";
            this.tbWordList.Size = new System.Drawing.Size(278, 202);
            this.tbWordList.TabIndex = 0;
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(864, 27);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(123, 38);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(864, 71);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(123, 38);
            this.btnReset.TabIndex = 2;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // btnExcelExport
            // 
            this.btnExcelExport.Location = new System.Drawing.Point(864, 115);
            this.btnExcelExport.Name = "btnExcelExport";
            this.btnExcelExport.Size = new System.Drawing.Size(123, 38);
            this.btnExcelExport.TabIndex = 3;
            this.btnExcelExport.Text = "Excel Export";
            this.btnExcelExport.UseVisualStyleBackColor = true;
            this.btnExcelExport.Click += new System.EventHandler(this.btnExcelExport_Click);
            // 
            // lbWordList
            // 
            this.lbWordList.AutoSize = true;
            this.lbWordList.Location = new System.Drawing.Point(12, 11);
            this.lbWordList.Name = "lbWordList";
            this.lbWordList.Size = new System.Drawing.Size(97, 12);
            this.lbWordList.TabIndex = 5;
            this.lbWordList.Text = "검색 단어 리스트";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Enabled = false;
            this.label1.Location = new System.Drawing.Point(294, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "앞면";
            // 
            // tbFront
            // 
            this.tbFront.Enabled = false;
            this.tbFront.Location = new System.Drawing.Point(296, 30);
            this.tbFront.Multiline = true;
            this.tbFront.Name = "tbFront";
            this.tbFront.Size = new System.Drawing.Size(278, 202);
            this.tbFront.TabIndex = 8;
            // 
            // tbBack
            // 
            this.tbBack.Enabled = false;
            this.tbBack.Location = new System.Drawing.Point(580, 30);
            this.tbBack.Multiline = true;
            this.tbBack.Name = "tbBack";
            this.tbBack.Size = new System.Drawing.Size(278, 202);
            this.tbBack.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Enabled = false;
            this.label3.Location = new System.Drawing.Point(578, 11);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 10;
            this.label3.Text = "뒷면";
            // 
            // Number
            // 
            this.Number.HeaderText = "번호";
            this.Number.MinimumWidth = 50;
            this.Number.Name = "Number";
            this.Number.Width = 50;
            // 
            // Vocabulary
            // 
            this.Vocabulary.HeaderText = "앞면";
            this.Vocabulary.MinimumWidth = 100;
            this.Vocabulary.Name = "Vocabulary";
            this.Vocabulary.Width = 400;
            // 
            // Pronunciation
            // 
            this.Pronunciation.HeaderText = "뒷면";
            this.Pronunciation.MinimumWidth = 100;
            this.Pronunciation.Name = "Pronunciation";
            this.Pronunciation.Width = 400;
            // 
            // EnglishWordBook
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(999, 564);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbBack);
            this.Controls.Add(this.tbFront);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbWordList);
            this.Controls.Add(this.btnExcelExport);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.tbWordList);
            this.Controls.Add(this.dgResult);
            this.Name = "EnglishWordBook";
            this.Text = "영단어정리기";
            ((System.ComponentModel.ISupportInitialize)(this.dgResult)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private HtmlAgilityPack.HtmlDocument LoadHtml(string asTargetUrl)
        {
            string str;
            using (WebClient client = new WebClient())
            {
                client.Encoding = Encoding.UTF8;
                str = client.DownloadString(asTargetUrl);
            }
            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(str);
            return document;
        }

        private string MyReplace(string sSource)
        {
            string[] strArray = new string[] { "                    ", "예문닫기", "\t", "??see also sea change", "?덈Ц?リ린", "??see also small change", "占?0" };
            foreach (string str in strArray)
            {
                sSource = sSource.Replace(str, "");
            }
            return sSource;
        }
    }
}

