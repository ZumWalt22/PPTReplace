using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace PPTAutoMakeTest
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        //置き換えログ保管用
        //StringBuilder sb = new StringBuilder();

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Power Point File 設定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSetPptFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "Power Point file|*.pptx|Power Point file 2003|*.ppt|All files(*.*)|*.*";
            ofd.Title = "Please select the Power point files.";
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;

            bool? rtn = ofd.ShowDialog();

            if (rtn != null && rtn == true)
            {
                labelPptFileName.Content = ofd.FileName;
            }
        }

        /// <summary>
        /// 置き換え文字列定義CSV File設定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSetTempFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "Temp file|*.csv|All files(*.*)|*.*";
            ofd.Title = "Temp select the Power point files.";
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;

            bool? rtn = ofd.ShowDialog();

            if (rtn != null && rtn == true)
            {
                labelTempFIleName.Content = ofd.FileName;
            }
        }

        /// <summary>
        /// テンプレートPPTファイルに文字を埋め込み
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Click(object sender, RoutedEventArgs e)
        {
            labelReplaceStatus.Content = "replace start";

            //テンプレートPPTファイル設定確認
            String pptTempFilePath = labelPptFileName.Content.ToString();
            if (String.IsNullOrEmpty(pptTempFilePath) || pptTempFilePath.Equals("-"))
            {
                labelReplaceStatus.Content = "テンプレートPowerPointファイルが設定されていません。";
                return;
            }

            //PPT保存ファイル名を取得
            String fileName = System.IO.Path.GetFileNameWithoutExtension(pptTempFilePath);
            fileName = replaceStr(fileName);
            //同一ファイル名がある場合は別名を作成
            if (fileName == System.IO.Path.GetFileNameWithoutExtension(pptTempFilePath))
            {
                fileName = System.IO.Path.GetFileNameWithoutExtension(pptTempFilePath) + "_Replace";
            }

            //ファイル保存先のフルパスを作成
            string pptGenerateFilePath = System.IO.Path.GetDirectoryName(pptTempFilePath) + "\\"
                + fileName
                + System.IO.Path.GetExtension(pptTempFilePath);

            //置き換え文字辞書を作成
            makeReplaceDic(labelTempFIleName.Content.ToString());

            //PPTテンプレート置き換えを実施
            doReplace(pptTempFilePath, pptGenerateFilePath);

            labelReplaceStatus.Content = "replace complete";
        }

        /// <summary>
        /// 置き換え文字辞書を作成
        /// </summary>
        private void makeReplaceDic(string replaceDicCsvFilePath)
        {
            //CSVファイル読み込み
            StreamReader sr = new StreamReader(
                  replaceDicCsvFilePath
                , Encoding.GetEncoding("Shift_JIS"));

            string tempStr = sr.ReadToEnd();

            sr.Close();

            //----行分解

            //置き換え文字列ペアを読み込み　
            string[] delimiter = { "\r\n", "\n" }; //改行で分割し置き換え文字列をリストに保管
            string[] replaceArray = tempStr.Split(delimiter, StringSplitOptions.None);

            //----カラム分解

            //行ごとの情報をKey、Valに分解し格納
            foreach (string replaceKeyValStr in replaceArray)
            {
                //shape.TextFrame.TextRange.Text.Replace();
                string[] delimiterCol = { @",""" };
                string[] replaceKeyVal = replaceKeyValStr.ToString().Split(delimiterCol, StringSplitOptions.None);

                if (replaceKeyVal.Length == 2)
                {
                    string key = replaceKeyVal[0].Replace(@"""", "");
                    string val = replaceKeyVal[1].Replace(@"""", "");

                    //辞書に追加
                    replaceKeyValDic.Add(key,val);
                }
            }
        }

        //変換情報の辞書情報（再帰的に使用されるためGlobalで宣言）
        Dictionary<String, String> replaceKeyValDic = new Dictionary<string, string>();

        /// <summary>
        /// 置き換え処理を実施
        /// </summary>
        /// <param name="pptFilePath">テンプレートPPTファイルパス</param>
        /// <param name="pptGenerateFilePath">生成PPT保存ファイルパス</param>
        private void doReplace(string pptFilePath,string pptGenerateFilePath)
        {
            List<string> notes = new List<string>();
            Microsoft.Office.Interop.PowerPoint.Application app = null;
            Microsoft.Office.Interop.PowerPoint.Presentation ppt = null;

            try
            {
                // PPTのインスタンス作成
                app = new Microsoft.Office.Interop.PowerPoint.Application();

                // PPTファイルオープン
                ppt = app.Presentations.Open(
                    pptFilePath,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse
               );

                // スライドのインデックスは１から　順にループする
                for (int i = 1; i <= ppt.Slides.Count; i++)
                {
                    //sb.AppendLine(" -------------------- Sheet " + ppt.Slides[i].SlideIndex.ToString());

                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in ppt.Slides[i].Shapes)
                    {
                        getShapeText(shape);
                    }
                }

                //生成PPTファイルの保存を実行
                ppt.SaveAs(pptGenerateFilePath,
                    PpSaveAsFileType.ppSaveAsDefault,
                    Microsoft.Office.Core.MsoTriState.msoFalse);
            }
            finally
            {
                // PPTファイルを閉じる
                if (ppt != null)
                {
                    ppt.Close();
                    ppt = null;
                }

                // PPTインスタンスを閉じる
                if (app != null)
                {
                    app.Quit();
                    app = null;
                }
            }
        }

        /// <summary>
        /// PPTオブジェクト内文字列を置き換え（再帰呼び出し）
        /// </summary>
        /// <param name="shape">PPTのShape</param>
        private void getShapeText(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            //sb.AppendLine(" -------------------- shape : " + shape.Id.ToString());
            //sb.AppendLine(" TYPE : " + shape.Type.ToString());
            //sb.AppendLine(" XY : " + shape.Top + " , " + shape.Left);
            //sb.AppendLine(" WH : " + shape.Width + " , " + shape.Height);

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                if (shape.TextFrame.TextRange.Text != "")
                {
                    //sb.AppendLine(" -------------------- shape : " + shape.Id.ToString());
                    //sb.AppendLine(" TYPE : " + shape.Type.ToString());
                    //sb.AppendLine(" XY : " + shape.Top + " , " + shape.Left);
                    //sb.AppendLine(" WH : " + shape.Width + " , " + shape.Height);
                    //sb.AppendLine(" Text : " + shape.TextFrame.TextRange.Text);

                    //PPT内の文字列置き換えを実施
                    shape.TextFrame.TextRange.Text = replaceStr(shape.TextFrame.TextRange.Text);
                }
            }

            // 構造が入れ子になっている場合を考慮し、再帰検索を実施
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape childShape in shape.GroupItems)
                {
                    //項目設定文字列を置き換え（再帰呼び出し）
                    getShapeText(childShape);
                }
            }

            //テーブル情報取得
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTable)
            {
                foreach (Row row in shape.Table.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        getShapeTextForTable(cell.Shape);
                    }
                }
            }
        }

        /// <summary>
        /// テーブル用文字列置き換え
        /// </summary>
        /// <param name="shape">PPTのShape</param>
        private void getShapeTextForTable(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                if (shape.TextFrame.TextRange.Text != "")
                {
                    shape.TextFrame.TextRange.Text = replaceStr(shape.TextFrame.TextRange.Text);
                }
            }
        }

        /// <summary>
        /// 文字列置き換え
        /// </summary>
        /// <param name="targetStr">置き換え対象文字列</param>
        /// <returns>置き換え後文字列</returns>
        private String replaceStr(String targetStr)
        {
            foreach (string replaceKeyValKey in replaceKeyValDic.Keys)
            {
                //PPTテンプレートに「[置き換え対象文字列]」の書式で設定したものを変換
                targetStr = targetStr.Replace("[" + replaceKeyValKey + "]", replaceKeyValDic[replaceKeyValKey]);
            }

            return targetStr;
        }
    }
}
