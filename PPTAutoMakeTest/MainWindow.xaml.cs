using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        StringBuilder sb = new StringBuilder();

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
                labelPptFIleName.Content = ofd.FileName;
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
            String pptFilePath = labelPptFIleName.Content.ToString().Replace("-","");
            if (String.IsNullOrEmpty(pptFilePath))
            {
                labelReplaceStatus.Content = "テンプレートPowerPointファイルが設定されていません。";
                return;
            }

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

                    sb.AppendLine(" -------------------- Sheet " + ppt.Slides[i].SlideIndex.ToString());

                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in ppt.Slides[i].Shapes)
                    {
                        getShapeText(shape);
                    }
                }

                //保存ファイル名を取得
                String fileName = System.IO.Path.GetFileNameWithoutExtension(labelPptFIleName.Content.ToString());
                fileName = replaceStr(fileName);
                //同一ファイル名がある場合は別名を作成
                if (fileName == System.IO.Path.GetFileNameWithoutExtension(labelPptFIleName.Content.ToString()))
                {
                    fileName = System.IO.Path.GetFileNameWithoutExtension(labelPptFIleName.Content.ToString()) + "_Replace";
                }

                //ファイル保存先のフルパスを作成
                string saveAsFile = System.IO.Path.GetDirectoryName(labelPptFIleName.Content.ToString()) + "\\"
                    + fileName
                    + System.IO.Path.GetExtension(labelPptFIleName.Content.ToString());

                //ファイルの別名保存を実行
                ppt.SaveAs(saveAsFile,
                    PpSaveAsFileType.ppSaveAsDefault,
                    Microsoft.Office.Core.MsoTriState.msoFalse);

                labelReplaceStatus.Content = "replace complete";
            }
            finally
            {
                // ファイルを閉じる
                if (ppt != null)
                {
                    ppt.Close();
                    ppt = null;
                }

                // PPTを閉じる
                if (app != null)
                {
                    app.Quit();
                    app = null;
                }
            }
        }

        private List<String> getReplaceStrList()
        {
            //CSVファイル読み込み
            StreamReader sr = new StreamReader(
                  labelTempFIleName.Content.ToString()
                , Encoding.GetEncoding("Shift_JIS"));

            string tempStr = sr.ReadToEnd();

            sr.Close();

            //置き換え文字列ペアを読み込み　
            string[] delimiter = { "\r\n", "\n" }; //改行で分割し置き換え文字列をリストに保管
            string[] replaceArray = tempStr.Split(delimiter, StringSplitOptions.None);

            return replaceArray.ToList<String>();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="shape"></param>
        private void getShapeText(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            sb.AppendLine(" -------------------- shape : " + shape.Id.ToString());
            sb.AppendLine(" TYPE : " + shape.Type.ToString());
            sb.AppendLine(" XY : " + shape.Top + " , " + shape.Left);
            sb.AppendLine(" WH : " + shape.Width + " , " + shape.Height);

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                if (shape.TextFrame.TextRange.Text != "")
                {
                    sb.AppendLine(" -------------------- shape : " + shape.Id.ToString());
                    sb.AppendLine(" TYPE : " + shape.Type.ToString());
                    sb.AppendLine(" XY : " + shape.Top + " , " + shape.Left);
                    sb.AppendLine(" WH : " + shape.Width + " , " + shape.Height);
                    sb.AppendLine(" Text : " + shape.TextFrame.TextRange.Text);

                    //PPT内の文字列置き換えを実施
                    shape.TextFrame.TextRange.Text = replaceStr(shape.TextFrame.TextRange.Text);
                }
            }

            // 構造が入れ子になっている場合を考慮し、再帰検索を実施
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape childShape in shape.GroupItems)
                {
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
        /// <param name="shape"></param>
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

        private String replaceStr(String targetStr)
        {
            //置き換えリストを取得
            List<String> replaceList = getReplaceStrList();

            foreach (string replaceKeyValStr in replaceList)
            {
                //shape.TextFrame.TextRange.Text.Replace();
                string[] delimiter = { @",""" };
                string[] replaceKeyVal = replaceKeyValStr.ToString().Split(delimiter, StringSplitOptions.None);

                if (replaceKeyVal.Length == 2)
                {
                    string key = replaceKeyVal[0].Replace(@"""", "");
                    string val = replaceKeyVal[1].Replace(@"""", "");

                    targetStr = targetStr.Replace("[" + key + "]", val);
                }
            }

            return targetStr;
        }
    }
}
