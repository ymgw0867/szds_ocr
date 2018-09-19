using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Leadtools;
using Leadtools.Codecs;
using SZDS_TIMECARD.Common;

namespace SZDS_TIMECARD.OCR
{
    public partial class frmOCR : Form
    {
        public frmOCR()
        {
            InitializeComponent();

            //// ＯＣＲモードがスキャナ読み取りのとき
            //if (Properties.Settings.Default.OCR_MODE == global.OCR_SCAN)
            //{
            //    // スキャナ画像出力先フォルダが登録されていないとき登録します
            //    if (!System.IO.Directory.Exists(Properties.Settings.Default.scanWinoutPath))
            //    {
            //        System.IO.Directory.CreateDirectory(Properties.Settings.Default.scanWinoutPath);
            //    }
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string jobname = string.Empty;
            string msg = string.Empty;

            if (rbtnTate.Checked)
            {
                jobname = Properties.Settings.Default.wrHands_Job_IP;
                msg = rbtnTate.Text;
            }
            else if (rbtnYoko.Checked)
            {
                jobname = Properties.Settings.Default.wrHands_Job_OUEN;
                msg = rbtnYoko.Text;
            }

            if (MessageBox.Show(msg + " のOCR変換を行います。よろしいですか", "OCR対象帳票確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                return;
            
            this.Hide();

            // マルチTiff画像をシングルtifに分解する(SCANフォルダ → TRAYフォルダ)
            if (!MultiTif(Properties.Settings.Default.scanPath, Properties.Settings.Default.trayPath))
            {
                this.Show();
                return;
            }

            // ＯＣＲ認識を実行します
            //DoOCR(Properties.Settings.Default.OCR_MODE);

            // 帳票ライブラリV8.0.3によるOCR認識実行
            wrhs803LibOCR(jobname);

            // フォームを閉じる
            this.Close();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     帳票認識ライブラリ V8.0.3 による認識処理実行
        /// </summary>
        ///----------------------------------------------------------------
        private void wrhs803LibOCR(string jobName)
        {
            // ファイル名のタイムスタンプを設定
            string fnm = string.Format("{0:0000}", DateTime.Today.Year) +
                         string.Format("{0:00}", DateTime.Today.Month) +
                         string.Format("{0:00}", DateTime.Today.Day) +
                         string.Format("{0:00}", DateTime.Now.Hour) +
                         string.Format("{0:00}", DateTime.Now.Minute) +
                         string.Format("{0:00}", DateTime.Now.Second);

            int sNum = 0;
            int ret = 0;

            try
            {
                // オーナーフォームを無効にする
                this.Enabled = false;

                // プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();

                // 処理する画像数を取得
                int t = System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif").Count();

                // 順番に認識処理を実行
                foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif"))
                {
                    // 画像数カウント
                    sNum++;

                    // プログレス表示
                    frmP.Text = "OCR認識中です ... " + sNum.ToString() + "/" + t.ToString();
                    frmP.progressValue = sNum * 100 / t;
                    frmP.ProgressStep();

                    // 標準パターンの読み込み
                    ret = FormRecog.OcrPatternLoad(Properties.Settings.Default.ocrPatternLoadPath);
                    
                    // パターン読み込みに成功したとき
                    if (ret > 0)
                    {
                        // 帳票認識ライブラリの制御内容を設定
                        FormRecog.OcrSetStatus(5, 1);   // 強制終了制御

                        // 認識結果出力イメージファイル
                        StringBuilder outimage = new StringBuilder(256);
                        outimage.Append(Properties.Settings.Default.wrOutPath + System.IO.Path.GetFileName(files));

                        // 認識結果出力テキストファイル
                        StringBuilder outtext = new StringBuilder(256);
                        outtext.Append(Properties.Settings.Default.wrOutPath + System.IO.Path.GetFileNameWithoutExtension(files) + ".csv");

                        // 認識結果 構造体
                        FormRecog.FORM_RECOG_DATA dt = new FormRecog.FORM_RECOG_DATA();

                        // 認識処理を開始
                        ret = FormRecog.OcrFormRecogStart(jobName, files, outimage, outtext, ref dt, false, false);

                        // 認識成功のとき
                        if (ret > 0)
                        {
                            // 認識結果のメモリ解放
                            ret = FormRecog.OcrFormStructFree(ref dt);

                            // 認識終了
                            ret = FormRecog.OcrFormRecogEnd();

                            //// PC毎の出力先フォルダがなければ作成する
                            //string rPath = Properties.Settings.Default.pcPath + _outPC + @"\";
                            //if (System.IO.Directory.Exists(rPath) == false)
                            //    System.IO.Directory.CreateDirectory(rPath);


                            // 出力先フォルダの確定 2017/01/25
                            string rPath = string.Empty;
                            if (rbtnTate.Checked)
                            {
                                // 勤怠データＩ／Ｐ票のとき
                                rPath = Properties.Settings.Default.dataPathIP;
                            }
                            else if (rbtnYoko.Checked)
                            {
                                // 応援移動票のとき
                                rPath = Properties.Settings.Default.dataPathOuen;
                            }

                            // 出力されたイメージファイルとテキストファイルのリネーム処理を行います
                            // READフォルダ → DATAフォルダ
                            string inCsvFile = Properties.Settings.Default.wrOutPath +
                                               Properties.Settings.Default.wrReaderOutFile;
                            string newFileName = rPath + fnm + sNum.ToString().PadLeft(3, '0');
                            wrhOutFileRename(inCsvFile, newFileName);
                        }
                        else
                        {
                            MessageBox.Show("OCR認識開始に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else
                    {
                        MessageBox.Show("OCR標準パターンの読み込みに失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                // 終了表示
                MessageBox.Show(sNum.ToString() + "件のOCR認識処理を行いました", "終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
               
                // TRAYフォルダの全てのtifファイルを削除します
                foreach (var files in System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     WinReaderFormOCRによるＯＣＲ処理を実行します </summary>
        /// <param name="ocrMode">
        ///     入力ファイル 1:スキャン, 2:画像</param>
        /// -------------------------------------------------------------------------
        //private void DoOCR(string ocrMode)
        //{
        //    // WinReaderFormOCRを起動
        //    WinReaderOCR(ocrMode);

        //    // ファイル名のタイムスタンプを設定
        //    string fnm = string.Format("{0:0000}", DateTime.Today.Year) +
        //                 string.Format("{0:00}", DateTime.Today.Month) +
        //                 string.Format("{0:00}", DateTime.Today.Day) +
        //                 string.Format("{0:00}", DateTime.Now.Hour) +
        //                 string.Format("{0:00}", DateTime.Now.Minute) +
        //                 string.Format("{0:00}", DateTime.Now.Second);

        //    // 連番を初期化
        //    dNo = 0;

        //    // ファイル分割処理
        //    LoadCsvDivide(ocrMode, fnm);
        //}

        // OCRファイル名連番 : WinReader
        int dNo = 0;

        // OCRファイル名（タイムスタンプ）: WinReader
        //string fnm = string.Empty;
              
        /// ----------------------------------------------------------------
        /// <summary>
        ///     WinReaderを起動してOCR処理を実施する</summary>
        /// <param name="OCRMODE">
        ///     スキャンモード 1:スキャナ, 2:イメージファイル</param>
        /// ----------------------------------------------------------------
        //private void WinReaderOCR(string OCRMODE)
        //{
        //    string JobName = string.Empty;

        //    // OCR認識モード毎のWinReaderJOB起動文字列
        //    if (OCRMODE == global.OCR_SCAN) // スキャナ
        //    {
        //        JobName = @"""" + Properties.Settings.Default.wrHands_Job_Scan + @"""" + " /H2";
        //    }
        //    else if (OCRMODE == global.OCR_IMAGE) // 画像
        //    {
        //        JobName = @"""" + Properties.Settings.Default.wrHands_Job_Image + @"""" + " /H2";
        //    }

        //    // WinReader実行ファイル
        //    string winReader_exe = Properties.Settings.Default.wrHands_Path + @"\" +
        //        Properties.Settings.Default.wrHands_Prg;

        //    // ProcessStartInfo の新しいインスタンスを生成する
        //    System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();

        //    // 起動するアプリケーションを設定する
        //    p.FileName = winReader_exe;

        //    // コマンドライン引数を設定する（WinReaderのJOB起動パラメーター）
        //    p.Arguments = JobName;

        //    // WinReaderを起動します
        //    System.Diagnostics.Process hProcess = System.Diagnostics.Process.Start(p);

        //    // WinReaderが終了するまで待機する
        //    hProcess.WaitForExit();
        //}

        /// -----------------------------------------------------------------
        /// <summary>
        ///     勤務報告書ＣＳＶデータを一枚ごとに分割する </summary>
        /// <param name="OCRMODE">
        ///     スキャンモード 1:スキャナ, 2:イメージファイル</param>
        /// -----------------------------------------------------------------
        //private void LoadCsvDivide(string OCRMODE, string fnm)
        //{
        //    string imgName = string.Empty;      // 画像ファイル名
        //    string firstFlg = global.FLGON;
        //    global.pblDenNum = 0;               // 枚数を0にセット
        //    string[] stArrayData;               // CSVファイルを１行単位で格納する配列
        //    string newFnm = string.Empty;       // 新ファイル名
        //    string inPath = string.Empty;       // ＯＣＲ認識モードごとの入力パス
        //    string inFilePath = string.Empty;   // ＯＣＲ認識モードごとの入力ファイル名
        //    //string dataMode = string.Empty;     // 勤務管理表種別（１：社員, ２：パート・アルバイト, ３：出向社員）

        //    // 入力ファイルパス
        //    if (OCRMODE == global.OCR_SCAN) // スキャナ
        //    {
        //        inPath = Properties.Settings.Default.scanWinoutPath;
        //        inFilePath = Properties.Settings.Default.scanWinoutPath + Properties.Settings.Default.winReaderOutFile;
        //    }
        //    else if (OCRMODE == global.OCR_IMAGE) // 画像
        //    {
        //        inPath = Properties.Settings.Default.imageWinoutPath;
        //        inFilePath = Properties.Settings.Default.imageWinoutPath + Properties.Settings.Default.winReaderOutFile;
        //    }

        //    // CSVデータの存在を確認します
        //    if (!System.IO.File.Exists(inFilePath)) return;

        //    // StreamReader の新しいインスタンスを生成する
        //    //入力ファイル
        //    System.IO.StreamReader inFile = new System.IO.StreamReader(inFilePath, Encoding.Default);

        //    // 読み込んだ結果をすべて格納するための変数を宣言する
        //    string stResult = string.Empty;
        //    string stBuffer;

        //    // 行番号
        //    int sRow = 0;

        //    // オーナーフォームを無効にする
        //    this.Enabled = false;

        //    // プログレスバーを表示する
        //    frmPrg frmP = new frmPrg();
        //    frmP.Owner = this;
        //    frmP.Show();

        //    // 勤務報告書枚数取得
        //    string[] t = System.IO.Directory.GetFiles(inPath, "*.tif");

        //    // 読み込みできる文字がなくなるまで繰り返す
        //    while (inFile.Peek() >= 0)
        //    {
        //        // ファイルを 1 行ずつ読み込む
        //        stBuffer = inFile.ReadLine();

        //        // カンマ区切りで分割して配列に格納する
        //        stArrayData = stBuffer.Split(',');

        //        //先頭に「*」か「#」があったら新たな伝票なのでCSVファイル作成
        //        if ((stArrayData[0] == "*"))
        //        {
        //            //最初の伝票以外のとき
        //            if (firstFlg == global.FLGOFF)
        //            {
        //                //ファイル書き出し
        //                outFileWrite(stResult, inPath + imgName, newFnm);
        //            }

        //            //伝票枚数カウント
        //            global.pblDenNum++;
        //            firstFlg = global.FLGOFF;

        //            // プログレス表示
        //            frmP.Text = "ＯＣＲデータロード中";
        //            frmP.progressValue = global.pblDenNum * 100 / t.Length;
        //            frmP.ProgressStep();

        //            // 伝票連番
        //            dNo++;

        //            // ファイル名
        //            newFnm = fnm + dNo.ToString().PadLeft(3, '0');

        //            // 画像ファイル名を取得
        //            imgName = stArrayData[1];

        //            //// 勤務報告書種別を取得（１：本社、２：静岡、３：大阪製造部）
        //            //dataMode = stArrayData[2];

        //            //文字列バッファをクリア
        //            stResult = string.Empty;

        //            // 文字列再構成（画像ファイル名を変更する）
        //            stBuffer = string.Empty;
        //            for (int i = 0; i < stArrayData.Length; i++)
        //            {
        //                if (stBuffer != string.Empty) stBuffer += ",";

        //                // 画像ファイル名を変更する
        //                if (i == 1) stArrayData[i] = newFnm + ".tif"; // 画像ファイル名を変更

        //                // フィールド結合
        //                string sta = stArrayData[i].Trim();
        //                stBuffer += sta;
        //            }

        //            sRow = 0;
        //        }
        //        else
        //        {
        //            sRow++;
        //        }

        //        // 読み込んだものを追加で格納する
        //        stResult += (stBuffer + Environment.NewLine);                
        //    }

        //    // いったんオーナーをアクティブにする
        //    this.Activate();

        //    // 進行状況ダイアログを閉じる
        //    frmP.Close();

        //    // オーナーのフォームを有効に戻す
        //    this.Enabled = true;

        //    // 後処理
        //    if (global.pblDenNum > 0)
        //    {
        //        // ファイル書き出し
        //        outFileWrite(stResult, inPath + imgName, newFnm);

        //        // 入力ファイルを閉じる
        //        inFile.Close();

        //        // 入力ファイル削除 : "txtout.csv"
        //        Utility.FileDelete(inPath, Properties.Settings.Default.winReaderOutFile);

        //        // 画像ファイル削除 : "WRH***.tif"
        //        Utility.FileDelete(inPath, "WRH*.tif");

        //        // 終了表示
        //        MessageBox.Show(global.pblDenNum.ToString() + "件の勤務報告書を処理しました", "終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //}

        /// -----------------------------------------------------------------
        /// <summary>
        ///     CSVファイルと画像ファイルの名前を日付スタンプに変更する </summary>
        /// <param name="readFilePath">
        ///     入力CSVファイル名(フルパス）</param>
        /// <param name="newFnm">
        ///     新ファイル名（フルパス・但し拡張子なし）</param>
        /// -----------------------------------------------------------------
        private void wrhOutFileRename(string readFilePath, string newFnm)
        {
            string imgName = string.Empty;      // 画像ファイル名
            string[] stArrayData;               // CSVファイルを１行単位で格納する配列
            string inFilePath = string.Empty;   // ＯＣＲ認識モードごとの入力ファイル名

            // CSVデータの存在を確認します
            if (!System.IO.File.Exists(readFilePath)) return;

            // StreamReader の新しいインスタンスを生成する
            //入力ファイル
            System.IO.StreamReader inFile = new System.IO.StreamReader(readFilePath, Encoding.Default);

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;
            string stBuffer;

            // 読み込みできる文字がなくなるまで繰り返す
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = inFile.ReadLine();

                // カンマ区切りで分割して配列に格納する
                stArrayData = stBuffer.Split(',');

                //先頭に「*」か「#」があったらヘッダー情報
                if ((stArrayData[0] == "*"))
                {
                    //文字列バッファをクリア
                    stResult = string.Empty;

                    // 文字列再構成（画像ファイル名を変更する）
                    stBuffer = string.Empty;
                    for (int i = 0; i < stArrayData.Length; i++)
                    {
                        if (stBuffer != string.Empty)
                        {
                            stBuffer += ",";
                        }

                        // 画像ファイル名を変更する
                        if (i == 1)
                        {
                            stArrayData[i] = System.IO.Path.GetFileName(newFnm) + ".tif"; // 画像ファイル名を変更
                        }

                        // フィールド結合
                        string sta = stArrayData[i].Trim();
                        stBuffer += sta;
                    }
                }

                // 読み込んだものを追加で格納する
                stResult += (stBuffer + Environment.NewLine);
            }

            // CSVファイル書き出し
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(newFnm + ".csv",
                                                    false, System.Text.Encoding.GetEncoding(932));
            outFile.Write(stResult);

            // 出力ファイルを閉じる
            outFile.Close();

            // 入力ファイルを閉じる
            inFile.Close();

            // 入力ファイル削除 : "txtout.csv"
            string inPath = System.IO.Path.GetDirectoryName(readFilePath);
            Utility.FileDelete(inPath, Properties.Settings.Default.wrReaderOutFile);

            // 画像ファイルをリネーム
            System.IO.File.Move(Properties.Settings.Default.wrOutPath + "WRH00001.tif", newFnm + ".tif");
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     分割ファイルを書き出す </summary>
        /// <param name="tempResult">
        ///     書き出す文字列</param>
        /// <param name="tempImgName">
        ///     元画像ファイルパス</param>
        /// <param name="outFileName">
        ///     新ファイル名</param>
        /// -------------------------------------------------------------------------------
        private void outFileWrite(string tempResult, string tempImgName, string outFileName)
        {
            //出力ファイル
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(Properties.Settings.Default.dataPathIP + outFileName + ".csv",
                                                    false, System.Text.Encoding.GetEncoding(932));
            // ファイル書き出し
            outFile.Write(tempResult);

            //ファイルクローズ
            outFile.Close();

            //画像ファイルをコピー
            System.IO.File.Copy(tempImgName, Properties.Settings.Default.dataPathIP + outFileName + ".tif");
        }

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    string [] tiffile = System.IO.Directory.GetFiles(Properties.Settings.Default.imagePath, "*.tif");

        //    if (tiffile.Length == 0)
        //    {
        //        MessageBox.Show("勤務報告書の画像がありません", "画像ＯＣＲ処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        return;
        //    }
        //    else
        //    {
        //        this.Hide();

        //        // 勤務報告書画像のＯＣＲ認識を実行します
        //        DoOCR(global.OCR_IMAGE);
        //        this.Close();
        //    }
        //}

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmOCR_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     マルチフレームの画像ファイルを頁ごとに分割する </summary>
        /// <param name="InPath">
        ///     画像ファイル入力パス</param>
        /// <param name="outPath">
        ///     分割後出力パス</param>
        /// <returns>
        ///     true:分割を実施, false:分割ファイルなし</returns>
        ///------------------------------------------------------------------------------
        private bool MultiTif(string InPath, string outPath)
        {
            //スキャン出力画像を確認
            if (System.IO.Directory.GetFiles(InPath, "*.tif").Count() == 0)
            {
                MessageBox.Show("ＯＣＲ変換処理対象の画像ファイルが指定フォルダ " + InPath + " に存在しません", "スキャン画像確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            // 出力先フォルダがなければ作成する
            if (System.IO.Directory.Exists(outPath) == false)
            {
                System.IO.Directory.CreateDirectory(outPath);
            }

            // 出力先フォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
            foreach (string files in System.IO.Directory.GetFiles(outPath, "*"))
            {
                System.IO.File.Delete(files);
            }

            RasterCodecs.Startup();
            RasterCodecs cs = new RasterCodecs();

            int _pageCount = 0;
            string fnm = string.Empty;

            // マルチTIFを分解して画像ファイルをTRAYフォルダへ保存する
            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                // 画像読み出す
                RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                // 頁数を取得
                int _fd_count = leadImg.PageCount;

                // 頁ごとに読み出す
                for (int i = 1; i <= _fd_count; i++)
                {
                    // ファイル名（日付時間部分）
                    string fName = string.Format("{0:0000}", DateTime.Today.Year) +
                            string.Format("{0:00}", DateTime.Today.Month) +
                            string.Format("{0:00}", DateTime.Today.Day) +
                            string.Format("{0:00}", DateTime.Now.Hour) +
                            string.Format("{0:00}", DateTime.Now.Minute) +
                            string.Format("{0:00}", DateTime.Now.Second);

                    // ファイル名設定
                    _pageCount++;
                    fnm = outPath + fName + string.Format("{0:000}", _pageCount) + ".tif";

                    // 画像保存
                    cs.Save(leadImg, fnm, RasterImageFormat.Tif, 0, i, i, 1, CodecsSavePageMode.Insert);
                }
            }

            // InPathフォルダの全てのtifファイルを削除する
            foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                System.IO.File.Delete(files);
            }

            return true;
        }

        private void frmOCR_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最少サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            //// コンボボックス
            //comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;

            //// コンボボックス
            //loadOutPcMst();
        }
    }
}
