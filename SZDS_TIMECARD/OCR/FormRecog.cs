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
        /// ̨�������߂̎w��
        /// </summary>
        public enum FieldType : ushort
        {
            /// <summary>
            /// ��̨����
            /// </summary>
            FIELD_INJI              = 145,
            /// <summary>
            /// �菑��̨����
            /// </summary>
            FIELD_TEGAKI            = 146,
            /// <summary>
            ///  ϰ�����̨����
            /// </summary>
            FIELD_MARK              = 147,
            /// <summary>
            /// �Ұ��̨����
            /// </summary>
            FIELD_IMAGE             = 148,
            /// <summary>
            /// ����̨����
            /// </summary>
            FIELD_KATUJI            = 149,
            /// <summary>
            /// QR���ށA�ް����̨����
            /// </summary>
            FIELD_QR_BARCODE        = 150
        }

        /// <summary>
        /// ��������
        /// </summary>
        [Flags]
        public enum MojiAttribute : ulong
        {
            /// <summary>
            /// �菑��/���� �L��
            /// </summary>
            ATR_HSYMBOL = 0x00000001,
            /// <summary>
            /// �菑��/���� ����
            /// </summary>
            ATR_HNUMBER = 0x00000002,
            /// <summary>
            /// �菑��/���� �J�^�J�i
            /// </summary>
            ATR_HKATAKANA = 0x00000004,
            /// <summary>
            /// �菑��/���� �p�啶��
            /// </summary>
            ATR_HALPHABET = 0x00000008,
            /// <summary>
            /// �菑��/���� �p������
            /// </summary>
            ATR_HALPHALOW = 0x00000010,
            /// <summary>
            /// �菑��/���� �Ђ炪��
            /// </summary>
            ATR_HHIRAGANA = 0x00000020,
            /// <summary>
            /// �菑��/���� ����
            /// </summary>
            ATR_HKANJI = 0x00000040,

            /// <summary>
            /// �菑��/���� ���[�U�[1
            /// </summary>
            ATR_USER1 = 0x00001000,
            /// <summary>
            /// �菑��/���� ���[�U�[2
            /// </summary>
            ATR_USER2 = 0x00002000,
            /// <summary>
            /// �菑��/���� ���[�U�[3
            /// </summary>
            ATR_USER3 = 0x00004000,
            /// <summary>
            /// �菑��/���� ���[�U�[4
            /// </summary>
            ATR_USER4 = 0x00008000,
            /// <summary>
            /// �菑��/���� ���[�U�[5
            /// </summary>
            ATR_USER5 = 0x00010000,
            /// <summary>
            /// �菑��/���� ���[�U�[6
            /// </summary>
            ATR_USER6 = 0x00020000,
            /// <summary>
            /// �菑��/���� ���[�U�[7
            /// </summary>
            ATR_USER7 = 0x00040000,
            /// <summary>
            /// �菑��/���� ���[�U�[8
            /// </summary>
            ATR_USER8 = 0x00080000,

            /// <summary>
            /// �󎚋L��
            /// </summary>
            ATR_PSYMBOL	= 0x00100000,
            /// <summary>
            /// �󎚐���
            /// </summary>
            ATR_PNUMBER	= 0x00200000,
            /// <summary>
            /// �󎚃J�^�J�i
            /// </summary>
            ATR_PKATAKANA = 0x00400000,
            /// <summary>
            /// �󎚉p�啶��
            /// </summary>
            ATR_PALPHABET = 0x00800000,

            /// <summary>
            /// �󎚃��[�U�[1
            /// </summary>
            ATR_PUSER1 = 0x01000000,
            /// <summary>
            /// �󎚃��[�U�[2
            /// </summary>
            ATR_PUSER2 = 0x02000000,
            /// <summary>
            /// �󎚃��[�U�[3
            /// </summary>
            ATR_PUSER3 = 0x04000000,
            /// <summary>
            /// �󎚃��[�U�[4
            /// </summary>
            ATR_PUSER4 = 0x08000000,
            /// <summary>
            /// �󎚃��[�U�[5
            /// </summary>
            ATR_PUSER5 = 0x10000000,
            /// <summary>
            /// �󎚃��[�U�[6
            /// </summary>
            ATR_PUSER6 = 0x20000000,
            /// <summary>
            /// �󎚃��[�U�[7
            /// </summary>
            ATR_PUSER7 = 0x40000000,
            /// <summary>
            /// �󎚃��[�U�[8
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
		/// �F�����ʁi�P�g���j
		/// </summary>
		[StructLayout(LayoutKind.Sequential)]
			public struct CHAR_RECOG_INFO
		{
            /// <summary>
            /// ������X���W�i�߸�ِ��j
            /// </summary>
			public ushort xMojiPoint;

            /// <summary>
            /// ������Y���W�i�߸�ِ��j
            /// </summary>
			public ushort yMojiPoint;

            /// <summary>
            /// �����̕��i�߸�ِ��j
            /// </summary>
			public ushort MojiWidth;

            /// <summary>
            /// �����̍����i�߸�ِ��j
            /// </summary>
			public ushort MojiHeight;

            /// <summary>
            /// �F�����ʊi�[�ر ��1���(2Byte),��2���,..��Can���
            /// </summary>
			public IntPtr pData;
		}



        /// <summary>
        /// �F�����ʕۑ��p�̍\���� 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct FORM_RECOG_DATA
        {
            /// <summary>
            /// ̨���ޔԍ�
            /// </summary>
            public ushort FieldNo;
            /// <summary>
            /// ̨��������
            /// </summary>
            public ushort Type;
            /// <summary>
            /// ��������
            /// </summary>
            public uint MojiAttr;
            /// <summary>
            /// ̨���ޘg��
            /// </summary>
            public ushort Waku;
            /// <summary>
            /// ��␔
            /// </summary>
            public byte Can;
            /// <summary>
            /// ؼު�ď�� (ؼު�Ă����������ŏ��̘g(1�`)�ԍ�
            /// </summary>
            public byte RejectWaku;
            /// <summary>
            /// �ް��������̌���(TRUE:FALSE,TRUE�̎��͎������藧���Ȃ�)�B
            /// </summary>
            public bool FlagDataCheck;
            /// <summary>
            /// �m�������װ�����������ꍇ ���̒l����Ă����
            /// </summary>
            public short FlagAI;
            /// <summary>
            /// �s�폜����Ă��邩(!=0)/�L����(=0)
            /// </summary>
            public byte FlagDead;
            /// <summary>
            /// ̨���ނ̍����X���W(�߸�ِ�)
            /// </summary>
            public short XPoint;
            /// <summary>
            /// ̨���ނ̍����Y���W(�߸�ِ�)
            /// </summary>
            public short YPoint;
            /// <summary>
            /// ̨���ނ̕�(�߸�ِ�)
            /// </summary>
            public short Width;
            /// <summary>
            /// ̨���ނ̍�(�߸�ِ�)
            /// </summary>
            public short Height;
            /// <summary>
            /// ϰ������̔F������(�ō��[�g���͍ŏ�[�g��LSB�ŏ����g32�܂���������Ă�����ޯĂ�ON)
            /// </summary>
            public uint MarkResult;
            /// <summary>
            /// �F�����ʏ����߲��(�g����)
            /// </summary>
            public IntPtr pRecogInfo;
            /// <summary>
            /// �菑��,��̨���ނ�Format �`��
            /// </summary>
            public IntPtr pFormat;
            /// <summary>
            /// ����̨���ނ̕���������
            /// </summary>
            public IntPtr pKatuji1;
            /// <summary>
            /// ����̨���ނ̍s��������
            /// </summary>
            public IntPtr pKatuji2;
            /// <summary>
            /// ����̨���ނ̕���������
            /// </summary>
            public IntPtr pKatuji3;
            /// <summary>
            /// ����̨���ނ̍폜�����Q(2Byte)
            /// </summary>
            public IntPtr pKatujiDelMoji;
            /// <summary>
            /// ����̨���ނ̌��蕶���Q(2Byte���ނ�2�����P��)
            /// </summary>
            public IntPtr pKatujiGentei;
            /// <summary>
            /// ϰ�����̨����ON����Format �`�� 
            /// </summary>
            public IntPtr pMarkOn;
            /// <summary>
            /// ϰ�����̨����OFF����Format �`�� 
            /// </summary>
            public IntPtr pMarkOff;
            /// <summary>
            /// Format���ꂽ�F������
            /// </summary>
            public IntPtr pText;
            /// <summary>
            /// 1�O�̍\���̂��߲��
            /// </summary>
            public IntPtr pPrev;
            /// <summary>
            /// 1��̍\���̂��߲��
            /// </summary>
            public IntPtr pNext;

        }



        ////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// ���[�F�����C�u�����̐�����e��ݒ肵�܂��B
        /// </summary>
        /// <param name="status_no">[in]����Ώۂ��w�肵�܂�</param>
        /// <param name="status">[in]������e���w�肵�܂�</param>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrSetStatus")]
        public static extern void OcrSetStatus(int status_no, int status);

        /// <summary>
        /// ����F�����̃}�b�`���O����ݒ肵�܂�
        /// </summary>
        /// <param name="Rate">[in]�}�b�`���O���i���j���w�肵�܂��B�w��\�͈͂́A�P����P�O�O�ł��B</param>
        /// <returns>�Ȃ�</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrSetMatchingRate")]
        public static extern void OcrSetMatchingRate(int Rate);

        /// <summary>
        /// �ĔF�����̃}�b�`���O����ݒ肵�܂�
        /// </summary>
        /// <param name="Rate">[in]�}�b�`���O���i���j���w�肵�܂��B�w��\�͈͂́A�P����P�O�O�ł��B</param>
        /// <returns>�Ȃ�</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrSetR_MatchingRate")]
        public static extern void OcrSetR_MatchingRate(int Rate);

        /// <summary>
        /// �W���p�^�[���t�@�C����ǂݍ��݁A�������Ɋi�[���܂��BHASP�̊m�F���s���܂��B
        /// </summary>
        /// <param name="pPath">[in]�W���p�^�[���t�@�C���̑��݂���t�H���_�����w�肵�܂��B</param>
        /// <returns>��������ƁA���̐���Ԃ��A�G���[����������ƕ��̃G���[�R�[�h��Ԃ��܂��B</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrPatternLoad")]
        public static extern int OcrPatternLoad(string pPath);

		/// <summary>
		/// OCR�̏��������s���܂��B���iID�̊m�F���s���܂��B
		/// </summary>
        /// <param name="pproductID">[in]���iID������</param>
        /// <param name="pPath">[in]�F�������̂���f�B���N�g����</param>
		/// <returns>��������ƁA���̐���Ԃ��A�G���[����������ƕ��̃G���[�R�[�h��Ԃ��܂��B���v���_�N�gID�͐��K�ɔz�z���ꂽ������ł���K�v���L��܂��B</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrPatternLoadByLicense")]
		public static extern int OcrPatternLoadByLicense(string pproductID ,string pPath);

        /// <summary>
        /// �ގ������t�@�C����ǂݍ��݁A�����������Ɋi�[���܂��B
        /// </summary>
        /// <param name="pFilename">[in]�ގ������t�@�C�������w�肵�܂�</param>
        /// <returns>��������ƁA���̐���Ԃ��A�G���[����������ƕ��̃G���[�R�[�h��Ԃ��܂��B</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrSimilarLoad")]
        public static extern int OcrSimilarLoad(string pFilename);

        /// <summary>
        /// 1�Ұ��̧��(���[�P��)�F���J�n
        /// </summary>
        /// <param name="pJobfilename">[in]FormOCR�ɂ��쐬���ꂽJOB��̧�ٖ����w�肵�܂��B����̧�ق� OcrPatternLoad() �֐��Ŏw�肳�ꂽ path �̉��� \JOB�t�H���_�ɑ��݂��Ȃ���΂Ȃ�܂���B</param>
        /// <param name="pImagefilename">[in]�F�����������Ұ��̧�ٖ����t���p�X�Ŏw�肵�܂��B</param>
        /// <param name="pOutImageName">[out]�F����o�͂��ꂽ���[�Ұ��̧�ٖ���ۑ�����ر�ł��BJOB�ŲҰ�ޏo�͂��s���悤�Ɏw�肳��Ă����ꍇ�A�o�͂��ꂽ̧�ٖ����i�[���܂��B�t���p�X�Ŋi�[���邽�ߏ\���ȴر���m�ۂ��ĉ������BNULL ���w�肵���ꍇ��A�gJOB�ŲҰ�ޏo�͂����Ȃ��h�Ǝw�肵�Ă���ꍇ�ɂ͊i�[����܂���B�܂��A�G���[�������ɂ͊i�[����܂���B</param>
        /// <param name="pOutTextName">[out]�F�����ʂ��t�@�C���Ƀe�L�X�g�o�͂����Ƃ���÷��̧�ٖ����i�[���܂��B�t���p�X�Ŋi�[���邽�ߏ\���ȴر���m�ۂ��ĉ������BNULL ���w�肵���ꍇ�ɂ͊i�[����܂���B�܂��A�G���[�������ɂ͊i�[����܂���B</param>
        /// <param name="pFormRecogData">[out]�F�����ʂ��i�[���� FORM_RECOG_DATA �\���̂��߲���ł��B</param>
        /// <param name="RetryFlag">[in]�ĔF�����w�肵�܂��B�ĔF������TRUE���A�ʏ��FALSE���w�肵�܂��BTRUE ���w�肳�ꂽ�ꍇ�A���[�}�b�`���O���̃��x���������܂��B����ɂ�葽�������ꂽ�ꍇ�̃C���[�W�t�@�C���ł��F���\�ƂȂ�܂��B</param>
        /// <param name="LearnFlag">[in]�����w�K�t�@�C���̓ǂݍ��݂̐�������܂��BTRUE��ݒ肷��Ɗ����w�K�t�@�C���̓ǂݍ��� ���s���܂��BFALSE�̏ꍇ�͓ǂݍ��݂����܂���B�A�����ē��������w�K�t�@�C���ŔF�����s���ꍇ�ɍŏ��̒��[�F������TRUE��ݒ肵�ȍ~��FALSE��ݒ肷�邱�ƂŔF�������������ɂȂ�܂��B</param>
        /// <returns>��������Ɛ��̐���Ԃ��A�G���[����������ƕ��̃G���[�R�[�h��Ԃ��܂��B</returns>
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
        /// �F������̉摜�̌X���␳�����擾���܂��B
        /// </summary>
        /// <param name="rotate">[in/out]�F�����ɉ摜����]�����p�x��Ԃ��܂��B�p�x�͎��v����0,90,180,270�x�̂����ꂩ�̒l�ɂȂ�܂��B</param>
        /// <param name="degree">[in/out]���[���̉摜���ɂ��X���␳�������p�x��Ԃ��܂��B���v���������v���X�Ƃ��āA�|9.0�`�{9.0���x�̒l�ɂȂ�܂��B</param>
        /// <returns>�Ȃ�</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrGetRotateInfo")]
        public static extern void OcrGetRotateInfo(ref int rotate, ref double degree);

        /// <summary>
        /// FORM_RECOG_DATA �\���̂̃�����������s���܂��B
        /// </summary>
        /// <param name="pRecogData">[in/out]�F�����ʂ��i�[���� FORM_RECOG_DATA �\���̂��߲���ł��B</param>
        /// <returns>��������Ɛ��̐���Ԃ��A�G���[����������ƕ��̃G���[�R�[�h��Ԃ��܂��B</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrFormStructFree")]
        public static extern int OcrFormStructFree(ref FORM_RECOG_DATA pRecogData);

        /// <summary>
        /// �F���̂��߂Ɋm�ۂ����̈��������܂��B
        /// </summary>
        /// <returns>��������Ɛ��̐���Ԃ��A�G���[����������ƕ��̃G���[�R�[�h��Ԃ��܂��B</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrFormRecogEnd")]
        public static extern int OcrFormRecogEnd();

        /// <summary>
        /// �F���Ɏg�p�����e���v���[�g�t�@�C�����ƁA�t�B�[���h�����o�͂��܂��B
        /// </summary>
        /// <param name="TplName">[out]�e���v���[�g�t�@�C����(260�o�C�g�Œ�)���i�[���܂��B</param>
        /// <param name="fieldMax">[out]�t�B�[���h�����i�[���܂��B</param>
        /// <returns>���������1��Ԃ��A �G���[�����������0��Ԃ��܂��B</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrGetTemplateFilenameUsedAfterRecognition")]
        public static extern int OcrGetTemplateFilenameUsedAfterRecognition(
            StringBuilder TplName,
            out int fieldMax
            );

        /// <summary>
        /// �F�����ʂɎg�p�����e���v���[�g�t�@�C���̃t�B�[���h����Ǎ��A�����������Ɋi�[���܂��B
        /// </summary>
        /// <param name="pTplFilePath">[in]�e���v���[�g�t�@�C���t���p�X���w�肵�܂��B</param>
        /// <returns>��������Ɛ� �܂��͑S�t�B�[���h����Ԃ��A�G���[����������ƕ� (MSG_ARGUMENT_ERR �ATMPL_F_OPEN_ERR�ATMPL_DATA_ERR)�̃G���[�R�[�h �܂���0��Ԃ��܂�</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "OcrLoadFieldInfo")]
        public static extern int OcrLoadFieldInfo(
            string pTplFilePath
            );

        /// <summary>
        /// �w�肵���t�B�[���h�ԍ��̃t�B�[���h���̂��o�͂��܂��B
        /// </summary>
        /// <param name="No">[in]�t�B�[���h�ԍ����w�肵�܂��B 1�`999�܂Ŏw��ł��܂��B</param>
        /// <param name="Name">[out]�t�B�[���h���̂��i�[���܂��B</param>
        /// <param name="Buf_size">[in]Name�̃o�b�t�@�T�C�Y���w�肵�܂��B260�o�C�g�w��</param>
        /// <returns>��������Ɛ��A�܂��͑S�t�B�[���h����Ԃ��A�G���[����������ƕ� (MSG_ARGUMENT_ERR �ATMPL_DATA_ERR)�̃G���[�R�[�h �܂���0��Ԃ��܂�</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "OcrGetOutputFieldName")]
        public static extern int OcrGetOutputFieldName(
            int No,
            StringBuilder Name,
            int Buf_size
           );

        /// <summary>
        /// �w�肵���t�B�[���h�ԍ��́u���W�F�N�g�����̎w�蕶���ƃ��W�F�N�g�����ݒ�t���O�v���o�͂��܂��B
        /// </summary>
        /// <param name="No">[in]�t�B�[���h�ԍ����w�肵�܂��B 1�`999�܂Ŏw��ł��܂��B</param>
        /// <param name="OutMoji">[out]�u���W�F�N�g�����̎w�蕶���v���i�[���܂��B</param>
        /// <param name="RejectFlag">[out]���W�F�N�g�����ݒ�t���O���i�[���܂��B �u���W�F�N�g�����w��v�F1�A�u�ŗL�͌��v�F0</param>
        /// <returns>��������Ɛ��A�܂��͑S�t�B�[���h����Ԃ��A�G���[����������ƕ� (MSG_ARGUMENT_ERR �ATMPL_DATA_ERR)�̃G���[�R�[�h �܂���0��Ԃ��܂�</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "OcrGetOutputRejectMoji")]
        public static extern int OcrGetOutputRejectMoji(
            int No,
            StringBuilder OutMoji,
            out int RejectFlag
           );

        /// <summary>
        /// �擾�����t�B�[���h����������܂��B
        /// </summary>
        /// <returns>�߂�l ���1��Ԃ��܂�</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "OcrClearFieldInfo")]
        public static extern int OcrClearFieldInfo();



        /// <summary>
        /// �t�B�[���h���̔F��������1�������Ƀ��W�F�N�g�̗L��������������Ƃ��ďo�͂��܂��B
        /// </summary>
        /// <param name="FieldNo">[in]�t�B�[���h�ԍ�</param>
        /// <param name="nSize">[in]������̃o�b�t�@�T�C�Y���w�肵�܂�</param>
        /// <param name="pRejectText">[out]���W�F�N�g�̗L����������������i�[���܂�</param>
        /// <returns>��������Ɛ���Ԃ��A �G���[����������ƕ��̃G���[�R�[�h��Ԃ��܂��B</returns>
        [DllImport("FormRecog.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi, EntryPoint = "OcrGetRejectPosition")]
        public static extern int OcrGetRejectPosition(
            ushort FieldNo,
            int nSize,
            StringBuilder pRejectText
            );

    }
}
