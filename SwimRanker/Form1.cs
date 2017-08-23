using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SwimRanker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //西暦選択のダイアログ
            for (int i = DateTime.Now.Year; i > 2010; i--)
                yearSelect.Items.Add(i.ToString());
            yearSelect.SelectedItem = DateTime.Now.Year.ToString();
        }

        private void URLCreator(string gameCode, string sex, string dirPath)
        {
            int dist = 0;
            for (int ev = 1; ev < 8; ev++)
            {
                switch (ev.ToString())
                {
                    case "1":
                        // free
                        for (dist = 2; dist < 8; dist++)
                        {
                            if ((sex == "1" && dist == 6) || (sex == "2" && dist == 7)) continue;
                            getHTML(gameCode, sex, ev.ToString(), dist.ToString(), dirPath);
                        }
                        break;
                    case "2":
                    case "3":
                    case "4":
                        // back
                        // breast
                        // butterfly
                        for (dist = 3; dist < 5; dist++)
                        {
                            getHTML(gameCode, sex, ev.ToString(), dist.ToString(), dirPath);
                        }
                        break;
                    case "5":
                        // IM
                        for (dist = 4; dist < 6; dist++)
                        {
                            getHTML(gameCode, sex, ev.ToString(), dist.ToString(), dirPath);
                        }
                        break;
                    case "6":
                        // FR
                        for (dist = 5; dist < 7; dist++)
                        {
                            getHTML(gameCode, sex, ev.ToString(), dist.ToString(), dirPath);
                        }
                        break;
                    case "7":
                        // MR
                        getHTML(gameCode, sex, ev.ToString(), "5", dirPath);
                        break;
                }
            }
        }

        private void getHTML(string gameCode, string sex, string ev, string dist, string dirPath)
        {
            // 性別とgameCode、出力先パスを受け取り、HTMLを出力先フォルダパスに入れる

            string swimRecURL = @"http://www.swim-record.com/swims/ViewResult/?h=V1100&code=" + gameCode + @"&sex=" + sex + @"&event=" + ev + @"&distance=" + dist;
            string dataName = dirPath + @"\" + gameCode + "_" + sex + "_" + ev + "_" + dist + ".html";

            WebClient client = new WebClient();

            byte[] data = client.DownloadData(swimRecURL);
            File.WriteAllBytes(dataName, data);

            return;
        }

        private void getSwimInfo(string styleCode, string distanceCode, ref string style, ref string distance)
        {
            switch (styleCode)
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

            switch (distanceCode)
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

        private void ParseHTMLToXlsx(string swimData, string gameName, string poolLength, string stdResultString,
            ref Excel.Application excelApp, ref Excel.Worksheet activeWorkSheet, bool ifInd, string style, string distance)
        {
            List<string> resultInfo = new List<string>();
            // ParseHTMLToXlsx(swimData, gameName, poolLength, resultInfo, stdTime, excelApp);
            StreamReader swimHtmlTxt = new StreamReader(swimData, Encoding.GetEncoding("shift_jis"));
            string swimStringData = swimHtmlTxt.ReadLine();

            // 1つめのtableまで読み飛ばす
            while (!swimStringData.StartsWith("<table")) swimStringData = swimHtmlTxt.ReadLine();
            // この５行下が大会名と長水路短水路
            for (int j = 0; j < 5; j++) swimStringData = swimHtmlTxt.ReadLine();
            // 一つ目の>の次から</までが大会名＋長短水路(最後の３文字)
            int startIndex = swimStringData.IndexOf(">") + 1;
            int endIndex = swimStringData.IndexOf("</") - 1;
            gameName = swimStringData.Substring(startIndex, endIndex - startIndex - 3);
            poolLength = swimStringData.Substring(endIndex - 2, 3);

            //個人:    "名前", "所属", "学年", "距離", "種目", "記録", "長短", "大会", "予決"
            //リレー:  "所属", "第1泳者", "第2泳者", "第3泳者", "第4泳者", "距離", "種目", "記録", "長短", "大会", "予決" 
            string yoketsu = "";
            string swimmerName = "", team = "", grade = "", result = "";
            string swim1 = "", swim2 = "", swim3 = "", swim4 = "";


            float resultComp;

            while (swimStringData.Trim() != "</body>")
            {
                if (swimStringData.Contains("予選") || swimStringData.Contains("決勝"))
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
                                        // 予選でかつresultが標準記録より遅かったらループを次へ飛ばす
                                        if (yoketsu == "予選")
                                        {
                                            resultComp = (result.Contains(":")) ?
                                            (float)(60 * int.Parse(result.Substring(0, result.IndexOf(":")))) + float.Parse(result.Substring(result.IndexOf(":") + 1)) :
                                            float.Parse(result);
                                            if (resultComp > float.Parse(stdResultString))
                                            {
                                                // ループを予選まで抜ける
                                                resFlag = false;
                                                break;
                                            }
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

                                        // 予選でかつresultが標準記録より遅かったらループを次へ飛ばす
                                        if (yoketsu == "予選")
                                        {
                                            resultComp = (float)(60 * int.Parse(result.Substring(0, result.IndexOf(":")))) + float.Parse(result.Substring(result.IndexOf(":") + 1));
                                            if (resultComp > float.Parse(stdResultString))
                                            {
                                                // ループを予選まで抜ける
                                                resFlag = false;
                                                break;
                                            }


                                        }
                                    }
                                    else if (k > 12)
                                    {
                                        break;
                                    }
                                }
                            }
                            // 1行ずつエクセルに追加
                            // 個人:    "名前", "所属", "学年", "距離", "種目", "記録", "長短", "大会", "予決"
                            // リレー:  "所属", "第1泳者", "第2泳者", "第3泳者", "第4泳者", "距離", "種目", "記録", "長短", "大会", "予決" 
                            resultInfo.Clear();
                            if (ifInd)
                            {
                                resultInfo.AddRange(new List<string> { swimmerName, team, grade, distance, style, result, poolLength, gameName, yoketsu });
                            }
                            else
                            {
                                resultInfo.AddRange(new List<string> { team, swim1, swim2, swim3, swim4, distance, style, result, poolLength, gameName, yoketsu });
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
                        swimStringData = swimHtmlTxt.ReadLine();
                    }
                }

                swimStringData = swimHtmlTxt.ReadLine();
            }
        }

        private void runningButton_Click(object sender, EventArgs e)
        {
            // config.txtとstandards.txtのパスの読み取り
            // 基本的に./config/.に入っている
            // Settings.settingsに保存しておく
            string prefData = Properties.Settings.Default.configPath.ToString();
            string stdPath = Properties.Settings.Default.standardPath.ToString();

            runningButton.Enabled = false;

            // config.txtを読み込んで各都道府県の大会IDを取得し、配列に格納
            StreamReader str = new StreamReader(prefData);
            string[,] prefCode = new string[48, 2];

            for (int i = 1; i < 48; i++)
            {
                // 都道府県名と大会コードを配列に格納
                var line = str.ReadLine().Split(',');
                for (int j = 1; j < 3; j++)
                {
                    prefCode[i, j - 1] = line[j];
                }
            }
            // 標準記録のインポート
            StreamReader stdRead = new StreamReader(stdPath);
            string[] tempStd = { "", "" };
            Dictionary<string, string> stdTime = new Dictionary<string, string>();
            while (stdRead.Peek() >= 0)
            {
                tempStd = stdRead.ReadLine().Split(',');
                stdTime.Add(tempStd[0], tempStd[1]);
            }

            string defOutput = System.Environment.CurrentDirectory.ToString() + @"\output" + yearSelect.SelectedItem.ToString() + @"\";

            // 吐き出し先エクセルファイルの生成
            // sheet1:men, sheet2:men_relay, sheet3: women, sheet4: women_relay
            string excelName = defOutput + yearSelect.SelectedItem.ToString() + @"rank.xlsx";
            string[] excelHeaderInd = { "名前", "所属", "学年", "距離", "種目", "記録", "長短", "大会", "予決" };
            string[] excelHeaderRelay = { "所属", "第1泳者", "第2泳者", "第3泳者", "第4泳者", "距離", "種目", "記録", "長短", "大会", "予決" };
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

            progressBarHTML.Minimum = 0;
            progressBarHTML.Maximum = 32 * 48;
            progressBarHTML.Value = 0;

            // 変数宣言
            string gameCode = "";
            string dirPath = "";

            for (int i = 1; i < 48; i++)
            {

                // 大会コード生成
                gameCode = string.Format("{0:00}", i) + yearSelect.SelectedItem.ToString().Substring(2) + prefCode[i, 1].ToString();


                // データ保存パス
                // 最終的には相対パス
                dirPath = defOutput + string.Format("{0:00}", i);

                // データがすでに存在していたら次の県
                if (!Directory.Exists(dirPath))
                {
                    label1.Text = "ＤＬ中: " + prefCode[i, 0];
                    label1.Refresh();
                    Directory.CreateDirectory(dirPath);

                    for (int sex = 1; sex < 3; sex++)
                    {
                        // 除外処理
                        if ((!exMen.Checked && sex == 1) || (!exWomen.Checked && sex == 2))
                            URLCreator(gameCode, sex.ToString(), dirPath);
                    }
                }


                //dirPathの中にあるhtmlファイルを全部解析
                foreach (string swimData in Directory.GetFiles(dirPath, "*.html"))
                {
                    // swimData: dirPathの中に存在するファイル一覧
                    // fileName: ggggggg_se_st_d.html (ggggggg: 大会番号、se: 性別、st: 種目、d: 距離)
                    // 名前を分割して性別、距離、種目を取得
                    string fileName = System.IO.Path.GetFileName(swimData);
                    string[] swimInfo = fileName.Split('_');
                    string sei = (swimInfo[1] == "1") ? "男子" : "女子";
                    string style = "";
                    string distance = "";

                    // 距離と種目の取得
                    getSwimInfo(swimInfo[2], swimInfo[3].Substring(0, 1), ref style, ref distance);

                    // ラベル更新
                    label1.Text = "処理中: " + prefCode[i, 0] + "," + sei + distance + "m" + style;
                    label1.Refresh();

                    // 除外処理
                    if (sei == "女子" && exWomen.Checked)
                    {
                        progressBarHTML.Value++;
                        continue;
                    }
                    else if (sei == "男子" && exMen.Checked)
                    {
                        progressBarHTML.Value++;
                        continue;
                    }

                    // ここからhtmlをエクセルにパース
                    string gameName = "";   //大会名
                    string poolLength = ""; //長短水路
                    List<string> resultInfo = new List<string>();
                    string stdGetName = sei + distance + "m" + style;
                    string stdResultString = stdTime[stdGetName];

                    Excel.Worksheet activeWorksheet;
                    if (sei == "男子")
                    {
                        if (int.Parse(swimInfo[2]) < 6)
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
                        if (int.Parse(swimInfo[2]) < 6)
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
                    // trueなら個人、falseならリレー
                    bool ifInd = int.Parse(swimInfo[2]) < 6;
                    ParseHTMLToXlsx(swimData, gameName, poolLength, stdResultString, ref excelApp, ref activeWorksheet, ifInd, style, distance);

                    progressBarHTML.Value++;
                    if ((exMen.Checked || exWomen.Checked) && progressBarHTML.Value != 2 * i) progressBarHTML.Value++;
                }

            }
            // エクセル保存
            wb.SaveAs(excelName);
            wb.Close(false);
            excelApp.Quit();

            label1.Text = "完了";
            // debug
            MessageBox.Show("完了");
            runningButton.Enabled = true;
        }

        private void 大会データ変更ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // ファイル選択ダイアログが開く
            // ファイル選択ダイアログの起点は、現在の設定ファイル
            openFileDialog1.Title = "大会データのファイルを選択してください";
            openFileDialog1.InitialDirectory = System.Environment.CurrentDirectory.ToString();
            openFileDialog1.FileName = System.IO.Path.GetFileName(Properties.Settings.Default.configPath);

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.configPath = openFileDialog1.FileName;
            }
            Properties.Settings.Default.Save();
        }

        private void 標準記録変更ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // ファイル選択ダイアログが開く
            // ファイル選択ダイアログの起点は、現在の設定ファイル
            openFileDialog2.Title = "標準記録のファイルを選択してください";
            openFileDialog2.InitialDirectory = System.Environment.CurrentDirectory.ToString();
            openFileDialog2.FileName = System.IO.Path.GetFileName(Properties.Settings.Default.standardPath);

            // ダイアログを表示し、戻り値が [OK] の場合は、選択したファイルを表示する
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.standardPath = openFileDialog2.FileName;
            }
            Properties.Settings.Default.Save();
        }

        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            return;
        }
    }
}
