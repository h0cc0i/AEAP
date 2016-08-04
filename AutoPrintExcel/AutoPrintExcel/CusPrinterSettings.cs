using System;
using System.Runtime.InteropServices;
using System.Drawing.Printing;
using System.ComponentModel;
namespace CustomprinterSettings
{
    ///

    /// Summary description for Class1.
    ///

    ///


    /*[Flags] private enum ACCESS_MASK

    {

    PRINTER_ACCESS_ADMINISTER = 4,

    PRINTER_ACCESS_USE = 8,

    PRINTER_ALL_ACCESS = 983052

    }*/
    public enum PageOrientation
    {
        DMORIENT_PORTRAIT = 1,
        DMORIENT_LANDSCAPE = 2
    }
    public enum PaperSize
    {

        // Paper selections
        DMPAPER_LETTER = 1, // Letter 8 1/2 x 11 in
        DMPAPER_LETTERSMALL = 2, // Letter Small 8 1/2 x 11 in
        DMPAPER_TABLOID = 3, // Tabloid 11 x 17 in
        DMPAPER_LEDGER = 4, // Ledger 17 x 11 in
        DMPAPER_LEGAL = 5, // Legal 8 1/2 x 14 in
        DMPAPER_STATEMENT = 6, // Statement 5 1/2 x 8 1/2 in
        DMPAPER_EXECUTIVE = 7, // Executive 7 1/4 x 10 1/2 in
        DMPAPER_A3 = 8, // A3 297 x 420 mm
        DMPAPER_A4 = 9, // A4 210 x 297 mm
        DMPAPER_A4SMALL = 10, // A4 Small 210 x 297 mm
        DMPAPER_A5 = 11, // A5 148 x 210 mm
        DMPAPER_B4 = 12, // B4 250 x 354
        DMPAPER_B5 = 13, // B5 182 x 257 mm
        DMPAPER_FOLIO = 14, // Folio 8 1/2 x 13 in
        DMPAPER_QUARTO = 15, // Quarto 215 x 275 mm
        DMPAPER_10X14 = 16, // 10x14 in
        DMPAPER_11X17 = 17, // 11x17 in
        DMPAPER_NOTE = 18, // Note 8 1/2 x 11 in
        DMPAPER_ENV_9 = 19, // Envelope #9 3 7/8 x 8 7/8
        DMPAPER_ENV_10 = 20, // Envelope #10 4 1/8 x 9 1/2
        DMPAPER_ENV_11 = 21, // Envelope #11 4 1/2 x 10 3/8
        DMPAPER_ENV_12 = 22, // Envelope #12 4 \276 x 11
        DMPAPER_ENV_14 = 23, // Envelope #14 5 x 11 1/2
        DMPAPER_CSHEET = 24, // C size sheet
        DMPAPER_DSHEET = 25, // D size sheet
        DMPAPER_ESHEET = 26, // E size sheet
        DMPAPER_ENV_DL = 27, // Envelope DL 110 x 220mm
        DMPAPER_ENV_C5 = 28, // Envelope C5 162 x 229 mm
        DMPAPER_ENV_C3 = 29, // Envelope C3 324 x 458 mm
        DMPAPER_ENV_C4 = 30, // Envelope C4 229 x 324 mm
        DMPAPER_ENV_C6 = 31, // Envelope C6 114 x 162 mm
        DMPAPER_ENV_C65 = 32, // Envelope C65 114 x 229 mm
        DMPAPER_ENV_B4 = 33, // Envelope B4 250 x 353 mm
        DMPAPER_ENV_B5 = 34, // Envelope B5 176 x 250 mm
        DMPAPER_ENV_B6 = 35, // Envelope B6 176 x 125 mm
        DMPAPER_ENV_ITALY = 36, // Envelope 110 x 230 mm
        DMPAPER_ENV_MONARCH = 37, // Envelope Monarch 3.875 x 7.5 in
        DMPAPER_ENV_PERSONAL = 38, // 6 3/4 Envelope 3 5/8 x 6 1/2 in
        DMPAPER_FANFOLD_US = 39, // US Std Fanfold 14 7/8 x 11 in
        DMPAPER_FANFOLD_STD_GERMAN = 40, // German Std Fanfold 8 1/2 x 12 in
        DMPAPER_FANFOLD_LGL_GERMAN = 41, // German Legal Fanfold 8 1/2 x 13 in
        DMPAPER_FIRST = DMPAPER_LETTER,
        DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN,

    }

    public enum PageDuplex
    {
        DMDUP_HORIZONTAL = 3,
        DMDUP_SIMPLEX = 1,
        DMDUP_VERTICAL = 2
    }

    public enum PaperSource
    {
        DMBIN_UPPER = 1,
        DMBIN_LOWER = 2,
        DMBIN_MIDDLE = 3,
        DMBIN_MANUAL = 4,
        DMBIN_ENVELOPE = 5,
        DMBIN_ENVMANUAL = 6,
        DMBIN_AUTO = 7,
        DMBIN_TRACTOR = 8,
        DMBIN_SMALLFMT = 9,
        DMBIN_LARGEFMT = 10,
        DMBIN_LARGECAPACITY = 11,
        DMBIN_CASSETTE = 14,
        DMBIN_FORMSOURCE = 15,
        DMRES_DRAFT = -1,
        DMRES_LOW = -2,
        DMRES_MEDIUM = -3,
        DMRES_HIGH = -4
    }

    #region "Data structure"
    public struct PrinterData
    {
        public PageOrientation Orientation;
        public PaperSize Size;
        public PaperSource source;
        public PageDuplex Duplex;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct PRINTER_DEFAULTS

    {

        public int pDatatype;

        public int pDevMode;

        public int DesiredAccess;

    }

    /*[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]

    public struct PRINTER_DEFAULTS

    {

    public int pDataType;

    public IntPtr pDevMode;

    public ACCESS_MASK DesiredAccess;

    }*/


    [StructLayout(LayoutKind.Sequential)]

    struct PRINTER_INFO_2

    {

        [MarshalAs(UnmanagedType.LPStr)]
        public string pServerName;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pPrinterName;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pShareName;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pPortName;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pDriverName;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pComment;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pLocation;

        public IntPtr pDevMode;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pSepFile;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pPrintProcessor;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pDatatype;

        [MarshalAs(UnmanagedType.LPStr)]
        public string pParameters;

        public IntPtr pSecurityDescriptor;

        public Int32 Attributes;

        public Int32 Priority;

        public Int32 DefaultPriority;

        public Int32 StartTime;

        public Int32 UntilTime;

        public Int32 Status;

        public Int32 cJobs;

        public Int32 AveragePPM;

    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE

    {

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]

        public string dmDeviceName;

        public short dmSpecVersion;

        public short dmDriverVersion;

        public short dmSize;

        public short dmDriverExtra;

        public int dmFields;

        public short dmOrientation;

        public short dmPaperSize;

        public short dmPaperLength;

        public short dmPaperWidth;

        public short dmScale;

        public short dmCopies;

        public short dmDefaultSource;

        public short dmPrintQuality;

        public short dmColor;

        public short dmDuplex;

        public short dmYResolution;

        public short dmTTOption;

        public short dmCollate;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]

        public string dmFormName;

        public short dmUnusedPadding;

        public short dmBitsPerPel;

        public int dmPelsWidth;

        public int dmPelsHeight;

        public int dmDisplayFlags;

        public int dmDisplayFrequency;

    }

    #endregion
    public class PrinterSettings
    {

        #region "Private Variables"

        private IntPtr hPrinter = new System.IntPtr();

        private PRINTER_DEFAULTS PrinterValues = new PRINTER_DEFAULTS();

        private PRINTER_INFO_2 pinfo = new PRINTER_INFO_2();

        private DEVMODE dm;

        private IntPtr ptrDM;

        private IntPtr ptrPrinterInfo;

        private int sizeOfDevMode = 0;

        private int lastError;

        private int nBytesNeeded;

        private long nRet;

        private int intError;

        private System.Int32 nJunk;

        private IntPtr yDevModeData;


        #endregion
        #region "Win API Def"

        [DllImport("kernel32.dll", EntryPoint = "GetLastError", SetLastError = false,

        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]

        private static extern Int32 GetLastError();

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true,

        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]

        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesA", SetLastError = true,

        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]

        private static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter,

        [MarshalAs(UnmanagedType.LPStr)] string pDeviceNameg,

        IntPtr pDevModeOutput, ref IntPtr pDevModeInput, int fMode);

        [DllImport("winspool.Drv", EntryPoint = "GetPrinterA", SetLastError = true, CharSet = CharSet.Ansi,

        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]

        private static extern bool GetPrinter(IntPtr hPrinter, Int32 dwLevel,

        IntPtr pPrinter, Int32 dwBuf, out Int32 dwNeeded);

        /*[DllImport("winspool.Drv", EntryPoint="OpenPrinterA", SetLastError=true, CharSet=CharSet.Ansi,

        ExactSpelling=true, CallingConvention=CallingConvention.StdCall)]

        static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter,

        out IntPtr hPrinter, ref PRINTER_DEFAULTS pd)


        [ DllImport( "winspool.drv",CharSet=CharSet.Unicode,ExactSpelling=false,

        CallingConvention=CallingConvention.StdCall )]

        public static extern long OpenPrinter(string pPrinterName,ref IntPtr phPrinter, int pDefault);*/


        /*[DllImport("winspool.Drv", EntryPoint="OpenPrinterA", SetLastError=true, CharSet=CharSet.Ansi,

        ExactSpelling=true, CallingConvention=CallingConvention.StdCall)]

        static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter,

        out IntPtr hPrinter, ref PRINTER_DEFAULTS pd);

        */

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi,

        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]

        private static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter,

        out IntPtr hPrinter, ref PRINTER_DEFAULTS pd);

        [DllImport("winspool.drv", CharSet = CharSet.Ansi, SetLastError = true)]

        private static extern bool SetPrinter(IntPtr hPrinter, int Level, IntPtr

        pPrinter, int Command);



        /*[DllImport("winspool.drv", CharSet=CharSet.Ansi, SetLastError=true)]

        private static extern bool SetPrinter(IntPtr hPrinter, int Level, IntPtr

        pPrinter, int Command);*/



        // Wrapper for Win32 message formatter.

        /*[DllImport("kernel32.dll", CharSet=System.Runtime.InteropServices.CharSet.Auto)]

        private unsafe static extern int FormatMessage( int dwFlags,

        ref IntPtr pMessageSource,

        int dwMessageID,

        int dwLanguageID,

        ref string lpBuffer,

        int nSize,

        IntPtr* pArguments);*/

        #endregion
        #region "Constants"
        private const int DM_DUPLEX = 0x1000;
        private const int DM_IN_BUFFER = 8;
        private const int DM_OUT_BUFFER = 2;
        private const int PRINTER_ACCESS_ADMINISTER = 0x4;
        private const int PRINTER_ACCESS_USE = 0x8;
        private const int STANDARD_RIGHTS_REQUIRED = 0xF0000;
        private const int PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED | PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE);
        #endregion
        #region "Function to change printer settings"
        public bool ChangePrinterSetting(string PrinterName, PrinterData PS)


        {


            if (((int)PS.Duplex < 1) || ((int)PS.Duplex > 3))

            {

                throw new ArgumentOutOfRangeException("nDuplexSetting", "nDuplexSetting is incorrect.");

            }

            else

            {

                dm = this.GetPrinterSettings(PrinterName);

                dm.dmDefaultSource = (short)PS.source;

                dm.dmOrientation = (short)PS.Orientation;

                dm.dmPaperSize = (short)PS.Size;

                dm.dmDuplex = (short)PS.Duplex;

                Marshal.StructureToPtr(dm, yDevModeData, true);

                pinfo.pDevMode = yDevModeData;

                pinfo.pSecurityDescriptor = IntPtr.Zero;

                /*update driver dependent part of the DEVMODE

                1 = DocumentProperties(IntPtr.Zero, hPrinter, sPrinterName, yDevModeData

                , ref pinfo.pDevMode, (DM_IN_BUFFER | DM_OUT_BUFFER));*/

                Marshal.StructureToPtr(pinfo, ptrPrinterInfo, true);

                lastError = Marshal.GetLastWin32Error();

                nRet = Convert.ToInt16(SetPrinter(hPrinter, 2, ptrPrinterInfo, 0));

                if (nRet == 0)

                {

                    //Unable to set shared printer settings.

                    lastError = Marshal.GetLastWin32Error();

                    //string myErrMsg = GetErrorMessage(lastError);

                    throw new Win32Exception(Marshal.GetLastWin32Error());


                }

                if (hPrinter != IntPtr.Zero)

                    ClosePrinter(hPrinter);

                return Convert.ToBoolean(nRet);

            }

        }

        private DEVMODE GetPrinterSettings(string PrinterName)

        {

            PrinterData PData = new PrinterData();

            DEVMODE dm;

            const int PRINTER_ACCESS_ADMINISTER = 0x4;

            const int PRINTER_ACCESS_USE = 0x8;

            const int PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED | PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE);





            PrinterValues.pDatatype = 0;

            PrinterValues.pDevMode = 0;

            PrinterValues.DesiredAccess = PRINTER_ALL_ACCESS;

            nRet = Convert.ToInt32(OpenPrinter(PrinterName, out hPrinter, ref PrinterValues));

            if (nRet == 0)

            {

                lastError = Marshal.GetLastWin32Error();

                throw new Win32Exception(Marshal.GetLastWin32Error());

            }

            GetPrinter(hPrinter, 2, IntPtr.Zero, 0, out nBytesNeeded);

            if (nBytesNeeded <= 0)

            {

                throw new System.Exception("Unable to allocate memory");

            }

            else

            {

                // Allocate enough space for PRINTER_INFO_2... {ptrPrinterIn fo = Marshal.AllocCoTaskMem(nBytesNeeded)};

                ptrPrinterInfo = Marshal.AllocHGlobal(nBytesNeeded);

                // The second GetPrinter fills in all the current settings, so all you // need to do is modify what you're interested in...

                nRet = Convert.ToInt32(GetPrinter(hPrinter, 2, ptrPrinterInfo, nBytesNeeded, out nJunk));

                if (nRet == 0)

                {

                    lastError = Marshal.GetLastWin32Error();

                    throw new Win32Exception(Marshal.GetLastWin32Error());

                }

                pinfo = (PRINTER_INFO_2)Marshal.PtrToStructure(ptrPrinterInfo, typeof(PRINTER_INFO_2));

                IntPtr Temp = new IntPtr();

                if (pinfo.pDevMode == IntPtr.Zero)

                {

                    // If GetPrinter didn't fill in the DEVMODE, try to get it by calling

                    // DocumentProperties...

                    IntPtr ptrZero = IntPtr.Zero;

                    //get the size of the devmode structure

                    sizeOfDevMode = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, ptrZero, ref ptrZero, 0);


                    ptrDM = Marshal.AllocCoTaskMem(sizeOfDevMode);

                    int i;

                    i = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, ptrDM,

                    ref ptrZero, DM_OUT_BUFFER);

                    if ((i < 0) || (ptrDM == IntPtr.Zero))

                    {

                        //Cannot get the DEVMODE structure.

                        throw new System.Exception("Cannot get DEVMODE data");

                    }

                    pinfo.pDevMode = ptrDM;

                }

                intError = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, IntPtr.Zero, ref Temp, 0);

                //IntPtr yDevModeData = Marshal.AllocCoTaskMem(i1);

                yDevModeData = Marshal.AllocHGlobal(intError);

                intError = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, yDevModeData, ref Temp, 2);

                dm = (DEVMODE)Marshal.PtrToStructure(yDevModeData, typeof(DEVMODE));

                //nRet = DocumentProperties(IntPtr.Zero, hPrinter, sPrinterName, yDevModeData

                // , ref yDevModeData, (DM_IN_BUFFER | DM_OUT_BUFFER));

                if ((nRet == 0) || (hPrinter == IntPtr.Zero))

                {

                    lastError = Marshal.GetLastWin32Error();

                    //string myErrMsg = GetErrorMessage(lastError);

                    throw new Win32Exception(Marshal.GetLastWin32Error());

                }

                return dm;

            }

            #endregion

        }
    }


}