using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace SwimRanker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string getHTML(string gameCode)
        {
            // URLを受け取り、指定されたURLのHTMLを文字列にして返す

            WebClient client = new WebClient();

            string data = client.DownloadString(gameCode);

            return data;
        }

        private void getSwimInfo(string url, ref string sex, ref string style, ref string distance)
        {
            string[] swimInfo = url.Split('&');
            // 1:code=4818345, 2:sex=1, 3:event=1, 4:distance=2

            switch (swimInfo[2].Substring(swimInfo[2].Length - 1))
            {
                case "1":
                    sex = "男子";
                    break;
                case "2":
                    sex = "女子";
                    break;
            }

            switch (swimInfo[3].Substring(swimInfo[3].Length - 1))
            {
                case "1":
                    style = "自由形";
                    break;
                case "2":
                    style = "背泳ぎ";
                    break;
                case "3":
                    style = "平泳ぎ";
                    break;
                case "4":
                    style = "バタフライ";
                    break;
                case "5":
                    style = "個人メドレー";
                    break;
                case "6":
                    style = "フリーリレー";
                    break;
                case "7":
                    style = "メドレーリレー";
                    break;
            }

            switch (swimInfo[4].Substring(swimInfo[4].Length - 1))
            {
                case "2":
                    distance = "50";
                    break;
                case "3":
                    distance = "100";
                    break;
                case "4":
                    distance = "200";
                    break;
                case "5":
                    distance = "400";
                    break;
                case "6":
                    distance = "800";
                    break;
                case "7":
                    distance = "1500";
                    break;
            }
        }

        private void ParseHTMLToXlsx(string swimData, ref Excel.Application excelApp, ref Excel.Worksheet activeWorkSheet, bool ifInd, string style, string distance)
        {
            List<string> resultInfo = new List<string>();
            StringReader swimHtmlTxt = new StringReader(swimData);
            string swimStringData = swimHtmlTxt.ReadLine();


            //個人:    "名前", "所属", "学年", "距離", "種目", "予決", "記録"
            //リレー:  "所属", "第1泳者", "第2泳者", "第3泳者", "第4泳者", "距離", "種目","予決", "記録" 
            string yoketsu = "";
            string swimmerName = "", team = "", grade = "", result = "";
            string swim1 = "", swim2 = "", swim3 = "", swim4 = "";

            int startIndex = 0;
            int endIndex = 0;

            while (swimStringData.Trim() != "</body>")
            {
                if (swimStringData.Contains("予選") || swimStringData.Contains("決勝") || swimStringData.Contains("スイムオフ"))
                {

                    // 予決情報の抽出
                    startIndex = swimStringData.IndexOf(">") + 1;
                    endIndex = swimStringData.IndexOf("</") - 1;
                    yoketsu = swimStringData.Substring(startIndex, endIndex - startIndex + 1);

                    swimStringData = swimHtmlTxt.ReadLine();

                    while (swimStringData.Trim() != "<br>")
                    {
                        // <tr align="center">が選手データの切れ目
                        if (swimStringData.Trim() == "<tr align=\"center\">")
                        {
                            bool resFlag = true;
                            int k = 0;
                            // htmlをパースしてデータを保持
                            // 必要なデータ: 名前、所属、学種学年、記録
                            // 個人の場合　　名前: 5下、所属: 更に2下、学種学年: 2下、記録: 2下
                            // リレーの場合　1-4泳者: 5-8下、所属: 9、記録: 11
                            while (true)
                            {
                                k++;
                                swimStringData = swimHtmlTxt.ReadLine();
                                if (ifInd)
                                {
                                    if (k == 5)
                                    {
                                        // 名前
                                        swimmerName = swimStringData.Trim().Substring(0, swimStringData.Trim().Length - 6);
                                    }
                                    else if (k == 7)
                                    {
                                        // 所属
                                        startIndex = swimStringData.IndexOf(">") + 1;
                                        endIndex = swimStringData.IndexOf("</") - 1;
                                        team = swimStringData.Substring(startIndex, endIndex - startIndex + 1);
                                    }
                                    else if (k == 9)
                                    {
                                        // 学種学年
                                        endIndex = swimStringData.IndexOf("</") - 1;
                                        grade = swimStringData.Substring(0, endIndex);
                                    }
                                    else if (k == 11)
                                    {
                                        // 記録
                                        startIndex = swimStringData.IndexOf(">") + 1;
                                        endIndex = swimStringData.IndexOf("</") - 1;
                                        if (endIndex - startIndex + 1 < 0)
                                        {
                                            resFlag = false;
                                            break;
                                        }
                                        result = swimStringData.Substring(startIndex, endIndex - startIndex + 1);
                                        if (result.EndsWith("-"))
                                        {
                                            resFlag = false;
                                            break;
                                        }
                                    }
                                    else if (k > 12)
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    if (k == 5)
                                    {
                                        // 名前1
                                        endIndex = swimStringData.Trim().IndexOf("<");
                                        if (endIndex < 8)
                                        {
                                            resFlag = false;
                                            break;
                                        }
                                        swim1 = swimStringData.Trim().Substring(8, endIndex - 8);
                                    }
                                    else if (k == 6)
                                    {
                                        // 名前2
                                        endIndex = swimStringData.Trim().IndexOf("<");
                                        swim2 = swimStringData.Trim().Substring(8, endIndex - 8);
                                    }
                                    else if (k == 7)
                                    {
                                        // 名前3
                                        endIndex = swimStringData.Trim().IndexOf("<");
                                        swim3 = swimStringData.Trim().Substring(8, endIndex - 8);
                                    }
                                    else if (k == 8)
                                    {
                                        // 名前4
                                        endIndex = swimStringData.Trim().IndexOf("</");
                                        swim4 = swimStringData.Trim().Substring(8, endIndex - 8).Trim();
                                    }
                                    else if (k == 9)
                                    {
                                        // 所属
                                        startIndex = swimStringData.IndexOf(">") + 1;
                                        endIndex = swimStringData.IndexOf("</") - 1;
                                        team = swimStringData.Substring(startIndex, endIndex - startIndex + 1);
                                    }
                                    else if (k == 11)
                                    {
                                        // 記録
                                        startIndex = swimStringData.IndexOf(">") + 1;
                                        endIndex = swimStringData.IndexOf("</") - 1;
                                        if (endIndex - startIndex + 1 < 0)
                                        {
                                            resFlag = false;
                                            break;
                                        }
                                        result = swimStringData.Substring(startIndex, endIndex - startIndex + 1);
                                        if (result.EndsWith("-"))
                                        {
                                            resFlag = false;
                                            break;
                                        }

                                    }
                                    else if (k > 12)
                                    {
                                        break;
                                    }
                                }
                            }
                            // 1行ずつエクセルに追加
                            // 個人:    "名前", "所属", "学年", "距離", "種目", "予決", "記録",  
                            // リレー:  "所属", "第1泳者", "第2泳者", "第3泳者", "第4泳者", "距離", "種目", "予決", "記録"
                            resultInfo.Clear();
                            if (team.Contains("慶應義塾") || team.Contains("慶応") || team.Contains("レッドクイーン") || team.Contains("K E I O"))
                            {
                                if (ifInd)
                                {
                                    resultInfo.AddRange(new List<string> { swimmerName, team, grade, distance, style, yoketsu, result });
                                }
                                else
                                {
                                    resultInfo.AddRange(new List<string> { team, swim1, swim2, swim3, swim4, distance, style, yoketsu, result });
                                }

                                if (resFlag)
                                {
                                    // エクセルに追加
                                    int lastLine = activeWorkSheet.UsedRange.Rows.Count + 1;
                                    if (ifInd)
                                    {
                                        // 個人
                                        for (int l = 1; l <= resultInfo.Count(); l++)
                                        {
                                            Excel.Range range = activeWorkSheet.Cells[lastLine, l];
                                            range.NumberFormatLocal = "@";
                                            range.Value = resultInfo[l - 1];
                                        }
                                    }
                                    else
                                    {
                                        // リレー
                                        for (int l = 1; l <= resultInfo.Count(); l++)
                                        {
                                            Excel.Range range_r = activeWorkSheet.Cells[lastLine, l];
                                            range_r.NumberFormatLocal = "@";
                                            range_r.Value = resultInfo[l - 1];
                                        }
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }

                        }
                        swimStringData = swimHtmlTxt.ReadLine();
                    }
                }

                swimStringData = swimHtmlTxt.ReadLine();
            }
        }

        private string getGameName(string data)
        {
            // テキストファイルを受け取って大会名を抽出して文字列で返す
            string gameName = "";

            StringReader swimHtmlTxt = new StringReader(data);
            string swimStringData = swimHtmlTxt.ReadLine();

            // 1つめのtableまで読み飛ばす
            while (!swimStringData.StartsWith("<table")) swimStringData = swimHtmlTxt.ReadLine();
            // この５行下が大会名と長水路短水路
            for (int j = 0; j < 5; j++) swimStringData = swimHtmlTxt.ReadLine();
            // 一つ目の>の次から</までが大会名＋長短水路(最後の３文字)
            int startIndex = swimStringData.IndexOf(">") + 1;
            int endIndex = swimStringData.IndexOf("</") - 1;
            gameName = swimStringData.Substring(startIndex, endIndex - startIndex - 3).Trim();
            gameName = gameName.Replace(" ", "").Replace("　", "");
            // ファイル名禁止文字処理
            gameName = Regex.Replace(gameName, ":","_" );
            return gameName;
        }

        private List<string> getGameUrlList(string data)
        {
            // ダウンロードURLリストを生成してList<string>で返す
            List<string> urlList = new List<string>();
            int startIndex = 0;
            int endIndex = 0;
            string urlValue = "";

            // dataの中にあるarefタグの種目の結果を表示するURLを全取得
            StringReader swimHtmlTxt = new StringReader(data);
            string swimStringData = swimHtmlTxt.ReadLine();

            while (swimStringData.Trim() != "</body>")
            {
                if (swimStringData.Contains("/swims/ViewResult?"))
                { 
                    // urlListにその行の最初の"から次の"までを文字列にして追加
                    startIndex = swimStringData.IndexOf("\"") + 1;
                    endIndex = swimStringData.IndexOf("\"", startIndex) - 1;
                    
                    urlValue = swimStringData.Substring(startIndex, endIndex - startIndex + 1);

                    // 追加するときに頭にhttp://www.swim-record.comを追加
                    urlList.Add("http://www.swim-record.com" + urlValue);
                }
                // 次の行を読む
                swimStringData = swimHtmlTxt.ReadLine();
            }

            return urlList;
        }

        private void runningButton_Click(object sender, EventArgs e)
        {
            this.runningButton.Enabled = false;
            progressBarHTML.Value = 0;

            backgroundWorker1.RunWorkerAsync();

            //// 大会のURLを取得
            //string gamecode = code.Text;

            //// ダウンロード
            //string gamedata = getHTML("http://www.swim-record.com/swims/ViewResult?h=V1000&code=" + gamecode);

            //// 日付と大会名の取得
            //string gameName = getGameName(gamedata);

            //// 種目ごとのURLリストの取得
            //List<string> gameUrlList = new List<string>();
            //gameUrlList = getGameUrlList(gamedata);

            //// 保存ファイルの作成
            ////string excelName = savePath.ToString();
            //string excelName = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\" + gameName + ".xlsx"; //適当なフォルダに、日付_大会名.xlsxという名前にする
            //string[] excelHeaderInd = { "名前", "所属", "学年", "距離", "種目", "予決", "記録" };
            //string[] excelHeaderRelay = { "所属", "第1泳者", "第2泳者", "第3泳者", "第4泳者", "距離", "種目", "予決", "記録" };
            //Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = false;
            //Excel.Workbook wb = excelApp.Workbooks.Add();
            //Excel.Worksheet men = wb.Worksheets[1];
            //Excel.Worksheet women = wb.Worksheets.Add(After: men);
            //Excel.Worksheet men_relay = wb.Worksheets.Add(After: women);
            //Excel.Worksheet women_relay = wb.Worksheets.Add(After: men_relay);
            //men.Name = "men";
            //women.Name = "women";
            //men_relay.Name = "men_relay";
            //women_relay.Name = "women_relay";

            //// ヘッダの挿入(個人)
            //for (int i = 1; i <= excelHeaderInd.Length; i++)
            //{
            //    Excel.Range rng_m = men.Cells[1, i];
            //    Excel.Range rng_w = women.Cells[1, i];
            //    rng_m.Value = rng_w.Value = excelHeaderInd[i - 1];
            //}

            //// ヘッダの挿入(リレー)
            //for (int i = 1; i <= excelHeaderRelay.Length; i++)
            //{
            //    Excel.Range rng_m = men_relay.Cells[1, i];
            //    Excel.Range rng_w = women_relay.Cells[1, i];
            //    rng_m.Value = rng_w.Value = excelHeaderRelay[i - 1];
            //}

            //progressBarHTML.Minimum = 0;
            //progressBarHTML.Maximum = gameUrlList.Count;
            //progressBarHTML.Value = 0;

            //// gameUrlListのhtmlを一つずつ処理
            //foreach (string url in gameUrlList)
            //{
            //    // urlのデータをダウンロード
            //    string results = getHTML(url);

            //    // url: http://www.swim-record.com/swims/ViewResult/?h=V1100&code=4818435&sex=1&event=1&distance=2
            //    // 名前を分割して性別、距離、種目を取得
            //    // 性別、距離、種目の取得
            //    string sei = "";
            //    string style = "";
            //    string distance = "";
            //    getSwimInfo(url, ref sei, ref style, ref distance);

            //    // ここからhtmlをエクセルにパース                
            //    Excel.Worksheet activeWorksheet;
            //    bool indivFlag = (style.Contains("リレー") ? false : true);
            //    if (sei == "男子")
            //    {
            //        if (indivFlag)
            //        {
            //            // 男子個人
            //            men.Select(Type.Missing);
            //            activeWorksheet = men;
            //        }
            //        else
            //        {
            //            // 男子リレー
            //            men_relay.Select(Type.Missing);
            //            activeWorksheet = men_relay;
            //        }
            //    }
            //    else
            //    {
            //        if (indivFlag)
            //        {
            //            // 女子個人
            //            women.Select(Type.Missing);
            //            activeWorksheet = women;
            //        }
            //        else
            //        {
            //            // 女子リレー
            //            women_relay.Select(Type.Missing);
            //            activeWorksheet = women_relay;
            //        }
            //    }

            //    // HTMLをExcelにパース
            //    ParseHTMLToXlsx(results, ref excelApp, ref activeWorksheet, indivFlag, style, distance);

            //    // プログレスバーの更新
            //    progressBarHTML.Value++;
            //}


            //// 終了処理
            //// エクセル保存
            //wb.SaveAs(excelName);
            //wb.Close(false);
            //excelApp.Quit();

            //MessageBox.Show("完了");
            //runningButton.Enabled = true;

        }



        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            return;
        }

        private void code_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) runningButton_Click(sender, e);
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            // 大会コードを取得
            string gamecode = code.Text;

            // 大会コードの例外処理
            if (gamecode.Length != 7)
            {
                MessageBox.Show("大会コードの長さが不正です。\n大会コードは半角数字7文字です。");
                return;
            }

            // ダウンロード
            string gamedata = getHTML("http://www.swim-record.com/swims/ViewResult?h=V1000&code=" + gamecode);

            // 日付と大会名の取得
            string gameName = getGameName(gamedata);

            // 種目ごとのURLリストの取得
            List<string> gameUrlList = new List<string>();
            gameUrlList = getGameUrlList(gamedata);

            // 保存ファイルの作成
            //string excelName = savePath.ToString();
            string excelName = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\" + gameName + ".xlsx"; //適当なフォルダに、日付_大会名.xlsxという名前にする
            string[] excelHeaderInd = { "名前", "所属", "学年", "距離", "種目", "予決", "記録" };
            string[] excelHeaderRelay = { "所属", "第1泳者", "第2泳者", "第3泳者", "第4泳者", "距離", "種目", "予決", "記録" };
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel.Workbook wb = excelApp.Workbooks.Add();
            Excel.Worksheet men = wb.Worksheets[1];
            Excel.Worksheet women = wb.Worksheets.Add(After: men);
            Excel.Worksheet men_relay = wb.Worksheets.Add(After: women);
            Excel.Worksheet women_relay = wb.Worksheets.Add(After: men_relay);
            men.Name = "men";
            women.Name = "women";
            men_relay.Name = "men_relay";
            women_relay.Name = "women_relay";

            // ヘッダの挿入(個人)
            for (int i = 1; i <= excelHeaderInd.Length; i++)
            {
                Excel.Range rng_m = men.Cells[1, i];
                Excel.Range rng_w = women.Cells[1, i];
                rng_m.Value = rng_w.Value = excelHeaderInd[i - 1];
            }

            // ヘッダの挿入(リレー)
            for (int i = 1; i <= excelHeaderRelay.Length; i++)
            {
                Excel.Range rng_m = men_relay.Cells[1, i];
                Excel.Range rng_w = women_relay.Cells[1, i];
                rng_m.Value = rng_w.Value = excelHeaderRelay[i - 1];
            }

            int parseProgress = 0;

            // gameUrlListのhtmlを一つずつ処理
            foreach (string url in gameUrlList)
            {
                // urlのデータをダウンロード
                string results = getHTML(url);

                // url: http://www.swim-record.com/swims/ViewResult/?h=V1100&code=4818435&sex=1&event=1&distance=2
                // 名前を分割して性別、距離、種目を取得
                // 性別、距離、種目の取得
                string sei = "";
                string style = "";
                string distance = "";
                getSwimInfo(url, ref sei, ref style, ref distance);

                // ここからhtmlをエクセルにパース                
                Excel.Worksheet activeWorksheet;
                bool indivFlag = (style.Contains("リレー") ? false : true);
                if (sei == "男子")
                {
                    if (indivFlag)
                    {
                        // 男子個人
                        men.Select(Type.Missing);
                        activeWorksheet = men;
                    }
                    else
                    {
                        // 男子リレー
                        men_relay.Select(Type.Missing);
                        activeWorksheet = men_relay;
                    }
                }
                else
                {
                    if (indivFlag)
                    {
                        // 女子個人
                        women.Select(Type.Missing);
                        activeWorksheet = women;
                    }
                    else
                    {
                        // 女子リレー
                        women_relay.Select(Type.Missing);
                        activeWorksheet = women_relay;
                    }
                }

                // HTMLをExcelにパース
                ParseHTMLToXlsx(results, ref excelApp, ref activeWorksheet, indivFlag, style, distance);

                // プログレスバーの更新
                parseProgress++;
                backgroundWorker1.ReportProgress((int)(((double)parseProgress / (double)gameUrlList.Count) * 100f));
                // progressBarHTML.Value = parseProgress/gameUrlList.Count;
            }


            // 終了処理
            // エクセル保存
            wb.SaveAs(excelName);
            wb.Close(false);
            excelApp.Quit();

            MessageBox.Show(excelName + "に保存されました。");
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            runningButton.Enabled = true;
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBarHTML.Value = e.ProgressPercentage;
        }
    }
}
