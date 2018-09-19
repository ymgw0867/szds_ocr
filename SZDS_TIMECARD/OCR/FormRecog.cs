//********************************************************
// This is a part of the Mediadrive Source Code Samples.
// Copyright (C) 2007- Mediadrive Corporation.
// *********************************************************/
using System;
using System.Text;
using System.Runtime.InteropServices;

namespace SZDS_TIMECARD
{
    public class FormRecog
    {
        private FormRecog()
        {
        }

        /// <summary>
        /// ﾌｨｰﾙﾄﾞﾀｲﾌﾟの指定
        /// </summary>
        public enum FieldType : ushort
        {
            /// <summary>
            /// 印字ﾌｨｰﾙﾄﾞ
            /// </summary>
            FIELD_INJI              = 145,
            /// <summary>
            /// 手書きﾌｨｰﾙﾄﾞ
            /// </summary>
            FIELD_TEGAKI            = 146,
            /// <summary>
            ///  ﾏｰｸﾁｪｯｸﾌｨｰﾙﾄﾞ
            /// </summary>
            FIELD_MARK              = 147,
            /// <summary>
            /// ｲﾒｰｼﾞﾌｨｰﾙﾄﾞ
            /// </summary>
            FIELD_IMAGE             = 148,
            /// <summary>
            /// 活字ﾌｨｰﾙﾄﾞ
            /// </summary>
            FIELD_KATUJI            = 149,
            /// <summary>
            /// QRｺｰﾄﾞ、ﾊﾞｰｺｰﾄﾞﾌｨｰﾙﾄﾞ
            /// </summary>
            FIELD_QR_BARCODE        = 150
        }

        /// <summary>
        /// 文字属性
        /// </summary>
        [Flags]
        public enum MojiAttribute : ulong
        {
            /// <summary>
            /// 手書き/活字 記号
            /// </summary>
            ATR_HSYMBOL = 0x00000001,
            /// <summary>
            /// 手書き/活字 数字
            /// </summary>
            ATR_HNUMBER = 0x00000002,
            /// <summary>
            /// 手書き/活字 カタカナ
            /// </summary>
            ATR_HKATAKANA = 0x00000004,
            /// <summary>
            /// 手書き/活字 英大文字
            /// </summary>
            ATR_HALPHABET = 0x00000008,
            /// <summary>
            /// 手書き/活字 英小文字
            /// </summary>
            ATR_HALPHALOW = 0x00000010,
            /// <summary>
            /// 手書き/活字 ひらがな
            /// </summary>
            ATR_HHIRAGANA = 0x00000020,
            /// <summary>
            /// 手書き/活字 漢字
            /// </summary>
            ATR_HKANJI = 0x00000040,

            /// <summary>
            /// 手書き/活字 ユーザー1
            /// </summary>
            ATR_USER1 = 0x00001000,
            /// <summary>
            /// 手書き/活字 ユーザー2
            /// </summary>
            ATR_USER2 = 0x00002000,
            /// <summary>
            /// 手書き/活字 ユーザー3
            /// </summary>
            ATR_USER3 = 0x00004000,
            /// <summary>
            /// 手書き/活字 ユーザー4
            /// </summary>
            ATR_USER4 = 0x00008000,
            /// <summary>
            /// 手書き/活字 ユーザー5
            /// </summary>
            ATR_USER5 = 0x00010000,
            /// <summary>
            /// 手書き/活字 ユーザー6
            /// </summary>
            ATR_USER6 = 0x00020000,
            /// <summary>
            /// 手書き/活字 ユーザー7
            /// </summary>
            ATR_USER7 = 0x00040000,
            /// <summary>
            /// 手書き/活字 ユーザー8
            /// </summary>
            ATR_USER8 = 0x00080000,

            /// <summary>
            /// 印字記号
            /// </summary>
            ATR_PSYMBOL	= 0x00100000,
            /// <summary>
            /// 印字数字
            /// </summary>
            ATR_PNUMBER	= 0x00200000,
            /// <summary>
            /// 印字カタカナ
            /// </summary>
            ATR_PKATAKANA = 0x00400000,
            /// <summary>
            /// 印字英大文字
            /// </summary>
            ATR_PALPHABET = 0x00800000,

            /// <summary>
            /// 印字ユーザー1
            /// </summary>
            ATR_PUSER1 = 0x01000000,
            /// <summary>
            /// 印字ユーザー2
            /// </summary>
            ATR_PUSER2 = 0x02000000,
            /// <summary>
            /// 印字ユーザー3
            /// </summary>
            ATR_PUSER3 = 0x04000000,
            /// <summary>
            /// 印字ユーザー4
            /// </summary>
            ATR_PUSER4 = 0x08000000,
            /// <summary>
            /// 印字ユーザー5
            /// </summary>
            ATR_PUSER5 = 0x10000000,
            /// <summary>
            /// 印字ユーザー6
            /// </summary>
            ATR_PUSER6 = 0x20000000,
            /// <summary>
            /// 印字ユーザー7
            /// </summary>
            ATR_PUSER7 = 0x40000000,
            /// <summary>
            /// 印字ユーザー8
            /// </summary>
            ATR_PUSER8 = 0x80000000,

            ATR_PSUJI_EIGO = ATR_PNUMBER | ATR_PALPHABET,
            ATR_SUJI_EIGO = ATR_HNUMBER | ATR_HALPHABET,
            ATR_SIX_CHECK = ATR_HSYMBOL | ATR_HNUMBER | ATR_HKATAKANA | ATR_HALPHABET | ATR_HHIRAGANA | ATR_HKANJI,
            ATR_NUM_KA_KAN_HI_CHECK = ATR_HNUMBER | ATR_HKATAKANA | ATR_HHIRAGANA | ATR_HKANJI,
            ATR_KI_KA_KAN_HI_CHECK = ATR_HSYMBOL | ATR_HKATAKANA | ATR_HHIRAGANA | ATR_HKANJI,
            ATR_KA_KAN_HI_CHECK = ATR_HKATAKANA | ATR_HHIRAGANA | ATR_HKANJI,
            ATR_NUM_KA_EI_KAN_HI_CHECK = ATR_HNUMBER | ATR_HKATAKANA | ATR_HALPHABET | ATR_HHIRAGANA | ATR_HKANJI

        }
        //////////////////////////////////////////////////////////////////////////////////////

		/// <summary>
		/// 認識結果（１枠分）
		/// </summary>
		[StructLayout(LayoutKind.Sequential)]
			public struct CHAR_RECOG_INFO
		{
            /// <summary>
            /// 文字のX座標（ﾋﾟｸｾﾙ数）
            /// </summary>
			public ushort xMojiPoint;

            /// <summary>
            /// 文字のY座標（ﾋﾟｸｾﾙ数）
            /// </summary>
			public ushort yMojiPoint;

            /// <summary>
            /// 文字の幅（ﾋﾟｸｾﾙ数）
            /// </summary>
			public ushort MojiWidth;

            /// <summary>
            /// 文字の高さ（ﾋﾟｸｾﾙ数）
            /// </summary>
			public ushort MojiHeight;

            /// <summary>
            /// 認識結果格納ｴﾘｱ 第1候補(2Byte),第2候補,..第Can候補
            /// </summary>
			public IntPtr pData;
		}



        /// <summary>
        /// 認識結果保存用の構造体 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct FORM_RECOG_DATA
        {
            /// <summary>
            /// ﾌｨｰﾙﾄﾞ番号
            /// </summary>
            public ushort FieldNo;
            /// <summary>
            /// ﾌｨｰﾙﾄﾞﾀｲﾌﾟ
            /// </summary>
            public ushort Type;
            /// <summary>
            /// 文字属性
            /// </summary>
            public uint MojiAttr;
            /// <summary>
            /// ﾌｨｰﾙﾄﾞ枠数
            /// </summary>
            public ushort Waku;
            /// <summary>
            /// 候補数
            /// </summary>
            public byte Can;
            /// <summary>
            /// ﾘｼﾞｪｸﾄ情報 (ﾘｼﾞｪｸﾄが発生した最初の枠(1～)番号
            /// </summary>
            public byte RejectWaku;
            /// <summary>
            /// ﾃﾞｰﾀﾁｪｯｸ式の結果(TRUE:FALSE,TRUEの時は式が成り立たない)。
            /// </summary>
            public bool FlagDataCheck;
            /// <summary>
            /// 知識処理ｴﾗｰが発生した場合 負の値がｾｯﾄされる
            /// </summary>
            public short FlagAI;
            /// <summary>
            /// 行削除されているか(!=0)/有効か(=0)
            /// </summary>
            public byte FlagDead;
            /// <summary>
            /// ﾌｨｰﾙﾄﾞの左上のX座標(ﾋﾟｸｾﾙ数)
            /// </summary>
            public short XPoint;
            /// <summary>
            /// ﾌｨｰﾙﾄﾞの左上のY座標(ﾋﾟｸｾﾙ数)
            /// </summary>
            public short YPoint;
            /// <summary>
            /// ﾌｨｰﾙﾄﾞの幅(ﾋﾟｸｾﾙ数)
            /// </summary>
            public short Width;
            /// <summary>
            /// ﾌｨｰﾙﾄﾞの高(ﾋﾟｸｾﾙ数)
            /// </summary>
            public short Height;
            /// <summary>
            /// ﾏｰｸﾁｪｯｸの認識結果(最左端枠又は最上端枠がLSBで順次枠32までﾁｪｯｸされていればﾋﾞｯﾄはON)
            /// </summary>
            public uint MarkResult;
            /// <summary>
            /// 認識結果情報のﾎﾟｲﾝﾀ(枠数分)
            /// </summary>
            public IntPtr pRecogInfo;
            /// <summary>
            /// 手書き,印字ﾌｨｰﾙﾄﾞのFormat 形式
            /// </summary>
            public IntPtr pFormat;
            /// <summary>
            /// 活字ﾌｨｰﾙﾄﾞの文頭文字列
            /// </summary>
            public IntPtr pKatuji1;
            /// <summary>
            /// 活字ﾌｨｰﾙﾄﾞの行末文字列
            /// </summary>
            public IntPtr pKatuji2;
            /// <summary>
            /// 活字ﾌｨｰﾙﾄﾞの文末文字列
            /// </summary>
            public IntPtr pKatuji3;
            /// <summary>
            /// 活字ﾌｨｰﾙﾄﾞの削除文字群(2Byte)
            /// </summary>
            public IntPtr pKatujiDelMoji;
            /// <summary>
            /// 活字ﾌｨｰﾙﾄﾞの限定文字群(2Byteｺｰﾄﾞの2文字単位)
            /// </summary>
            public IntPtr pKatujiGentei;
            /// <summary>
            /// ﾏｰｸﾁｪｯｸﾌｨｰﾙﾄﾞON時のFormat 形式 
            /// </summary>
            public IntPtr pMarkOn;
            /// <summary>
            /// ﾏｰｸﾁｪｯｸﾌｨｰﾙﾄﾞOFF時のFormat 形式 
            /// </summary>
            public IntPtr pMarkOff;
            /// <summary>
            /// Formatされた認識結果
            /// </summary>
            public IntPtr pText;
            /// <summary>
            /// 1つ前の構造体のﾎﾟｲﾝﾀ
            /// </summary>
            public IntPtr pPrev;
            /// <summary>
            /// 1つ後の構造体のﾎﾟｲﾝﾀ
            /// </summary>
            public IntPtr pNext;

        }



        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// 帳票認識ライブラリの制御内容を設定します。
        /// </summary>
        /// <param name="status_no">[in]制御対象を指定します</param>
        /// <param name="status">[in]制御内容を指定します</param>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrSetStatus")]
        public static extern void OcrSetStatus(int status_no, int status);

        /// <summary>
        /// 初回認識時のマッチング率を設定します
        /// </summary>
        /// <param name="Rate">[in]マッチング率（％）を指定します。指定可能範囲は、１から１００です。</param>
        /// <returns>なし</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrSetMatchingRate")]
        public static extern void OcrSetMatchingRate(int Rate);

        /// <summary>
        /// 再認識時のマッチング率を設定します
        /// </summary>
        /// <param name="Rate">[in]マッチング率（％）を指定します。指定可能範囲は、１から１００です。</param>
        /// <returns>なし</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrSetR_MatchingRate")]
        public static extern void OcrSetR_MatchingRate(int Rate);

        /// <summary>
        /// 標準パターンファイルを読み込み、メモリに格納します。HASPの確認を行います。
        /// </summary>
        /// <param name="pPath">[in]標準パターンファイルの存在するフォルダ名を指定します。</param>
        /// <returns>成功すると、正の数を返し、エラーが発生すると負のエラーコードを返します。</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrPatternLoad")]
        public static extern int OcrPatternLoad(string pPath);

		/// <summary>
		/// OCRの初期化を行います。商品IDの確認を行います。
		/// </summary>
        /// <param name="pproductID">[in]商品ID文字列</param>
        /// <param name="pPath">[in]認識辞書のあるディレクトリ名</param>
		/// <returns>成功すると、正の数を返し、エラーが発生すると負のエラーコードを返します。※プロダクトIDは正規に配布された文字列である必要が有ります。</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrPatternLoadByLicense")]
		public static extern int OcrPatternLoadByLicense(string pproductID ,string pPath);

        /// <summary>
        /// 類似文字ファイルを読み込み、内部メモリに格納します。
        /// </summary>
        /// <param name="pFilename">[in]類似文字ファイル名を指定します</param>
        /// <returns>成功すると、正の数を返し、エラーが発生すると負のエラーコードを返します。</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrSimilarLoad")]
        public static extern int OcrSimilarLoad(string pFilename);

        /// <summary>
        /// 1ｲﾒｰｼﾞﾌｧｲﾙ(帳票単位)認識開始
        /// </summary>
        /// <param name="pJobfilename">[in]FormOCRにより作成されたJOBのﾌｧｲﾙ名を指定します。このﾌｧｲﾙは OcrPatternLoad() 関数で指定された path の下の \JOBフォルダに存在しなければなりません。</param>
        /// <param name="pImagefilename">[in]認識させたいｲﾒｰｼﾞﾌｧｲﾙ名をフルパスで指定します。</param>
        /// <param name="pOutImageName">[out]認識後出力された帳票ｲﾒｰｼﾞﾌｧｲﾙ名を保存するｴﾘｱです。JOBでｲﾒｰｼﾞ出力を行うように指定されていた場合、出力されたﾌｧｲﾙ名を格納します。フルパスで格納するため十分なｴﾘｱを確保して下さい。NULL を指定した場合や、“JOBでｲﾒｰｼﾞ出力をしない”と指定してある場合には格納されません。また、エラー発生時には格納されません。</param>
        /// <param name="pOutTextName">[out]認識結果をファイルにテキスト出力したときのﾃｷｽﾄﾌｧｲﾙ名を格納します。フルパスで格納するため十分なｴﾘｱを確保して下さい。NULL を指定した場合には格納されません。また、エラー発生時には格納されません。</param>
        /// <param name="pFormRecogData">[out]認識結果を格納する FORM_RECOG_DATA 構造体のﾎﾟｲﾝﾀです。</param>
        /// <param name="RetryFlag">[in]再認識を指定します。再認識時はTRUEを、通常はFALSEを指定します。TRUE を指定された場合、帳票マッチング時のレベルを下げます。これにより多少かすれた場合のイメージファイルでも認識可能となります。</param>
        /// <param name="LearnFlag">[in]活字学習ファイルの読み込みの制御をします。TRUEを設定すると活字学習ファイルの読み込み を行います。FALSEの場合は読み込みをしません。連続して同じ活字学習ファイルで認識を行う場合に最初の帳票認識時にTRUEを設定し以降はFALSEを設定することで認識処理が高速になります。</param>
        /// <returns>成功すると正の数を返し、エラーが発生すると負のエラーコードを返します。</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrFormRecogStart")]
        public static extern int OcrFormRecogStart(
            string pJobfilename,
            string pImagefilename,
            StringBuilder pOutImageName,
            StringBuilder pOutTextName,
            ref FORM_RECOG_DATA pFormRecogData,
            bool RetryFlag,
            bool LearnFlag);

        /// <summary>
        /// 認識直後の画像の傾き補正情報を取得します。
        /// </summary>
        /// <param name="rotate">[in/out]認識時に画像を回転した角度を返します。角度は時計回りで0,90,180,270度のいずれかの値になります。</param>
        /// <param name="degree">[in/out]帳票内の画像情報により傾き補正をした角度を返します。時計回り方向をプラスとして、－9.0～＋9.0程度の値になります。</param>
        /// <returns>なし</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrGetRotateInfo")]
        public static extern void OcrGetRotateInfo(ref int rotate, ref double degree);

        /// <summary>
        /// FORM_RECOG_DATA 構造体のメモリ解放を行います。
        /// </summary>
        /// <param name="pRecogData">[in/out]認識結果を格納する FORM_RECOG_DATA 構造体のﾎﾟｲﾝﾀです。</param>
        /// <returns>成功すると正の数を返し、エラーが発生すると負のエラーコードを返します。</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrFormStructFree")]
        public static extern int OcrFormStructFree(ref FORM_RECOG_DATA pRecogData);

        /// <summary>
        /// 認識のために確保した領域を解放します。
        /// </summary>
        /// <returns>成功すると正の数を返し、エラーが発生すると負のエラーコードを返します。</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrFormRecogEnd")]
        public static extern int OcrFormRecogEnd();

        /// <summary>
        /// 認識に使用したテンプレートファイル名と、フィールド数を出力します。
        /// </summary>
        /// <param name="TplName">[out]テンプレートファイル名(260バイト固定)を格納します。</param>
        /// <param name="fieldMax">[out]フィールド数を格納します。</param>
        /// <returns>成功すると1を返し、 エラーが発生すると0を返します。</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrGetTemplateFilenameUsedAfterRecognition")]
        public static extern int OcrGetTemplateFilenameUsedAfterRecognition(
            StringBuilder TplName,
            out int fieldMax
            );

        /// <summary>
        /// 認識結果に使用したテンプレートファイルのフィールド情報を読込、内部メモリに格納します。
        /// </summary>
        /// <param name="pTplFilePath">[in]テンプレートファイルフルパスを指定します。</param>
        /// <returns>成功すると正 または全フィールド数を返し、エラーが発生すると負 (MSG_ARGUMENT_ERR 、TMPL_F_OPEN_ERR、TMPL_DATA_ERR)のエラーコード または0を返します</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "OcrLoadFieldInfo")]
        public static extern int OcrLoadFieldInfo(
            string pTplFilePath
            );

        /// <summary>
        /// 指定したフィールド番号のフィールド名称を出力します。
        /// </summary>
        /// <param name="No">[in]フィールド番号を指定します。 1～999まで指定できます。</param>
        /// <param name="Name">[out]フィールド名称を格納します。</param>
        /// <param name="Buf_size">[in]Nameのバッファサイズを指定します。260バイト指定</param>
        /// <returns>成功すると正、または全フィールド数を返し、エラーが発生すると負 (MSG_ARGUMENT_ERR 、TMPL_DATA_ERR)のエラーコード または0を返します</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "OcrGetOutputFieldName")]
        public static extern int OcrGetOutputFieldName(
            int No,
            StringBuilder Name,
            int Buf_size
           );

        /// <summary>
        /// 指定したフィールド番号の「リジェクト文字の指定文字とリジェクト文字設定フラグ」を出力します。
        /// </summary>
        /// <param name="No">[in]フィールド番号を指定します。 1～999まで指定できます。</param>
        /// <param name="OutMoji">[out]「リジェクト文字の指定文字」を格納します。</param>
        /// <param name="RejectFlag">[out]リジェクト文字設定フラグを格納します。 「リジェクト文字指定」：1、「最有力候補」：0</param>
        /// <returns>成功すると正、または全フィールド数を返し、エラーが発生すると負 (MSG_ARGUMENT_ERR 、TMPL_DATA_ERR)のエラーコード または0を返します</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "OcrGetOutputRejectMoji")]
        public static extern int OcrGetOutputRejectMoji(
            int No,
            StringBuilder OutMoji,
            out int RejectFlag
           );

        /// <summary>
        /// 取得したフィールド情報を解放します。
        /// </summary>
        /// <returns>戻り値 常に1を返します</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrClearFieldInfo")]
        public static extern int OcrClearFieldInfo();



        /// <summary>
        /// フィールド内の認識文字を1文字毎にリジェクトの有無を示す文字列として出力します。
        /// </summary>
        /// <param name="FieldNo">[in]フィールド番号</param>
        /// <param name="nSize">[in]文字列のバッファサイズを指定します</param>
        /// <param name="pRejectText">[out]リジェクトの有無を示す文字列を格納します</param>
        /// <returns>成功すると正を返し、 エラーが発生すると負のエラーコードを返します。</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "OcrGetRejectPosition")]
        public static extern int OcrGetRejectPosition(
            ushort FieldNo,
            int nSize,
            StringBuilder pRejectText
            );

    }
}
