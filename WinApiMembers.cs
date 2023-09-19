using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace WinApi
{
    internal static class WinApiMembers
    {
        #region 定数・列挙

        #region ウィンドウスタイル

        internal const int WS_VISIBLE = 0x10000000;
        internal const int WS_SYSMENU = 0x80000;
        internal const int WS_MINIMIZEBOX = 0x00020000;
        internal const int WS_MAXIMIZEBOX = 0x10000;
        internal const int WS_BORDER = 0x00800000;
        internal const int WS_DLGFRAME = 0x00400000;
        internal const int WS_THICKFRAME = 0x00040000;
        internal const int WS_CAPTION = 0x00C00000;
        internal const long WS_POPUP = 0x80000000L;

        #endregion

        #region 拡張ウィンドウスタイル

        internal const int WS_EX_TOOLWINDOW = 0x80;

        #endregion

        #region ウィンドウメッセージ

        internal const int WM_NULL = 0x0000;
        internal const int WM_CREATE = 0x0001;
        internal const int WM_DESTROY = 0x0002;
        internal const int WM_MOVE = 0x0003;
        internal const int WM_SIZE = 0x0005;
        internal const int WM_ACTIVATE = 0x0006;
        internal const int WM_SETFOCUS = 0x0007;
        internal const int WM_KILLFOCUS = 0x0008;
        internal const int WM_ENABLE = 0x000A;
        internal const int WM_SETREDRAW = 0x000B;
        internal const int WM_SETTEXT = 0x000C;
        internal const int WM_GETTEXT = 0x000D;
        internal const int WM_GETTEXTLENGTH = 0x000E;
        internal const int WM_PAINT = 0x000F;
        internal const int WM_CLOSE = 0x0010;
        internal const int WM_QUERYENDSESSION = 0x0011;
        internal const int WM_QUERYOPEN = 0x0013;
        internal const int WM_ENDSESSION = 0x0016;
        internal const int WM_QUIT = 0x0012;
        internal const int WM_ERASEBKGND = 0x0014;
        internal const int WM_SYSCOLORCHANGE = 0x0015;
        internal const int WM_SHOWWINDOW = 0x0018;
        internal const int WM_WININICHANGE = 0x001A;
        internal const int WM_SETTINGCHANGE = WM_WININICHANGE;
        internal const int WM_DEVMODECHANGE = 0x001B;
        internal const int WM_ACTIVATEAPP = 0x001C;
        internal const int WM_FONTCHANGE = 0x001D;
        internal const int WM_TIMECHANGE = 0x001E;
        internal const int WM_CANCELMODE = 0x001F;
        internal const int WM_SETCURSOR = 0x0020;
        internal const int WM_MOUSEACTIVATE = 0x0021;
        internal const int WM_CHILDACTIVATE = 0x0022;
        internal const int WM_QUEUESYNC = 0x0023;
        internal const int WM_GETMINMAXINFO = 0x0024;
        internal const int WM_PAINTICON = 0x0026;
        internal const int WM_ICONERASEBKGND = 0x0027;
        internal const int WM_NEXTDLGCTL = 0x0028;
        internal const int WM_SPOOLERSTATUS = 0x002A;
        internal const int WM_DRAWITEM = 0x002B;
        internal const int WM_MEASUREITEM = 0x002C;
        internal const int WM_DELETEITEM = 0x002D;
        internal const int WM_VKEYTOITEM = 0x002E;
        internal const int WM_CHARTOITEM = 0x002F;
        internal const int WM_SETFONT = 0x0030;
        internal const int WM_GETFONT = 0x0031;
        internal const int WM_SETHOTKEY = 0x0032;
        internal const int WM_GETHOTKEY = 0x0033;
        internal const int WM_QUERYDRAGICON = 0x0037;
        internal const int WM_COMPAREITEM = 0x0039;
        internal const int WM_GETOBJECT = 0x003D;
        internal const int WM_COMPACTING = 0x0041;
        internal const int WM_COMMNOTIFY = 0x0044;
        internal const int WM_WINDOWPOSCHANGING = 0x0046;
        internal const int WM_WINDOWPOSCHANGED = 0x0047;
        internal const int WM_POWER = 0x0048;
        internal const int WM_COPYDATA = 0x004A;
        internal const int WM_CANCELJOURNAL = 0x004B;
        internal const int WM_NOTIFY = 0x004E;
        internal const int WM_INPUTLANGCHANGEREQUEST = 0x0050;
        internal const int WM_INPUTLANGCHANGE = 0x0051;
        internal const int WM_TCARD = 0x0052;
        internal const int WM_HELP = 0x0053;
        internal const int WM_USERCHANGED = 0x0054;
        internal const int WM_NOTIFYFORMAT = 0x0055;
        internal const int WM_CONTEXTMENU = 0x007B;
        internal const int WM_STYLECHANGING = 0x007C;
        internal const int WM_STYLECHANGED = 0x007D;
        internal const int WM_DISPLAYCHANGE = 0x007E;
        internal const int WM_GETICON = 0x007F;
        internal const int WM_SETICON = 0x0080;
        internal const int WM_NCCREATE = 0x0081;
        internal const int WM_NCDESTROY = 0x0082;
        internal const int WM_NCCALCSIZE = 0x0083;
        internal const int WM_NCHITTEST = 0x0084;
        internal const int WM_NCPAINT = 0x0085;
        internal const int WM_NCACTIVATE = 0x0086;
        internal const int WM_GETDLGCODE = 0x0087;
        internal const int WM_SYNCPAINT = 0x0088;
        internal const int WM_NCMOUSEMOVE = 0x00A0;
        internal const int WM_NCLBUTTONDOWN = 0x00A1;
        internal const int WM_NCLBUTTONUP = 0x00A2;
        internal const int WM_NCLBUTTONDBLCLK = 0x00A3;
        internal const int WM_NCRBUTTONDOWN = 0x00A4;
        internal const int WM_NCRBUTTONUP = 0x00A5;
        internal const int WM_NCRBUTTONDBLCLK = 0x00A6;
        internal const int WM_NCMBUTTONDOWN = 0x00A7;
        internal const int WM_NCMBUTTONUP = 0x00A8;
        internal const int WM_NCMBUTTONDBLCLK = 0x00A9;
        internal const int WM_NCXBUTTONDOWN = 0x00AB;
        internal const int WM_NCXBUTTONUP = 0x00AC;
        internal const int WM_NCXBUTTONDBLCLK = 0x00AD;
        internal const int WM_INPUT = 0x00FF;
        internal const int WM_KEYFIRST = 0x0100;
        internal const int WM_KEYDOWN = 0x0100;
        internal const int WM_KEYUP = 0x0101;
        internal const int WM_CHAR = 0x0102;
        internal const int WM_DEADCHAR = 0x0103;
        internal const int WM_SYSKEYDOWN = 0x0104;
        internal const int WM_SYSKEYUP = 0x0105;
        internal const int WM_SYSCHAR = 0x0106;
        internal const int WM_SYSDEADCHAR = 0x0107;
        internal const int WM_UNICHAR = 0x0109;
        internal const int WM_KEYLAST = 0x0108;
        internal const int WM_IME_STARTCOMPOSITION = 0x010D;
        internal const int WM_IME_ENDCOMPOSITION = 0x010E;
        internal const int WM_IME_COMPOSITION = 0x010F;
        internal const int WM_IME_KEYLAST = 0x010F;
        internal const int WM_INITDIALOG = 0x0110;
        internal const int WM_COMMAND = 0x0111;
        internal const int WM_SYSCOMMAND = 0x0112;
        internal const int WM_TIMER = 0x0113;
        internal const int WM_HSCROLL = 0x0114;
        internal const int WM_VSCROLL = 0x0115;
        internal const int WM_INITMENU = 0x0116;
        internal const int WM_INITMENUPOPUP = 0x0117;
        internal const int WM_MENUSELECT = 0x011F;
        internal const int WM_MENUCHAR = 0x0120;
        internal const int WM_ENTERIDLE = 0x0121;
        internal const int WM_MENURBUTTONUP = 0x0122;
        internal const int WM_MENUDRAG = 0x0123;
        internal const int WM_MENUGETOBJECT = 0x0124;
        internal const int WM_UNINITMENUPOPUP = 0x0125;
        internal const int WM_MENUCOMMAND = 0x0126;
        internal const int WM_CHANGEUISTATE = 0x0127;
        internal const int WM_UPDATEUISTATE = 0x0128;
        internal const int WM_QUERYUISTATE = 0x0129;
        internal const int WM_CTLCOLOR = 0x0019;
        internal const int WM_CTLCOLORMSGBOX = 0x0132;
        internal const int WM_CTLCOLOREDIT = 0x0133;
        internal const int WM_CTLCOLORLISTBOX = 0x0134;
        internal const int WM_CTLCOLORBTN = 0x0135;
        internal const int WM_CTLCOLORDLG = 0x0136;
        internal const int WM_CTLCOLORSCROLLBAR = 0x0137;
        internal const int WM_CTLCOLORSTATIC = 0x0138;
        internal const int WM_MOUSEFIRST = 0x0200;
        internal const int WM_MOUSEMOVE = 0x0200;
        internal const int WM_LBUTTONDOWN = 0x0201;
        internal const int WM_LBUTTONUP = 0x0202;
        internal const int WM_LBUTTONDBLCLK = 0x0203;
        internal const int WM_RBUTTONDOWN = 0x0204;
        internal const int WM_RBUTTONUP = 0x0205;
        internal const int WM_RBUTTONDBLCLK = 0x0206;
        internal const int WM_MBUTTONDOWN = 0x0207;
        internal const int WM_MBUTTONUP = 0x0208;
        internal const int WM_MBUTTONDBLCLK = 0x0209;
        internal const int WM_MOUSEWHEEL = 0x020A;
        internal const int WM_XBUTTONDOWN = 0x020B;
        internal const int WM_XBUTTONUP = 0x020C;
        internal const int WM_XBUTTONDBLCLK = 0x020D;
        internal const int WM_MOUSELAST = 0x020D;
        internal const int WM_PARENTNOTIFY = 0x0210;
        internal const int WM_ENTERMENULOOP = 0x0211;
        internal const int WM_EXITMENULOOP = 0x0212;
        internal const int WM_NEXTMENU = 0x0213;
        internal const int WM_SIZING = 0x0214;
        internal const int WM_CAPTURECHANGED = 0x0215;
        internal const int WM_MOVING = 0x0216;
        internal const int WM_POWERBROADCAST = 0x0218;
        internal const int WM_DEVICECHANGE = 0x0219;
        internal const int WM_MDICREATE = 0x0220;
        internal const int WM_MDIDESTROY = 0x0221;
        internal const int WM_MDIACTIVATE = 0x0222;
        internal const int WM_MDIRESTORE = 0x0223;
        internal const int WM_MDINEXT = 0x0224;
        internal const int WM_MDIMAXIMIZE = 0x0225;
        internal const int WM_MDITILE = 0x0226;
        internal const int WM_MDICASCADE = 0x0227;
        internal const int WM_MDIICONARRANGE = 0x0228;
        internal const int WM_MDIGETACTIVE = 0x0229;
        internal const int WM_MDISETMENU = 0x0230;
        internal const int WM_ENTERSIZEMOVE = 0x0231;
        internal const int WM_EXITSIZEMOVE = 0x0232;
        internal const int WM_DROPFILES = 0x0233;
        internal const int WM_MDIREFRESHMENU = 0x0234;
        internal const int WM_IME_SETCONTEXT = 0x0281;
        internal const int WM_IME_NOTIFY = 0x0282;
        internal const int WM_IME_CONTROL = 0x0283;
        internal const int WM_IME_COMPOSITIONFULL = 0x0284;
        internal const int WM_IME_SELECT = 0x0285;
        internal const int WM_IME_CHAR = 0x0286;
        internal const int WM_IME_REQUEST = 0x0288;
        internal const int WM_IME_KEYDOWN = 0x0290;
        internal const int WM_IME_KEYUP = 0x0291;
        internal const int WM_MOUSEHOVER = 0x02A1;
        internal const int WM_MOUSELEAVE = 0x02A3;
        internal const int WM_NCMOUSELEAVE = 0x02A2;
        internal const int WM_WTSSESSION_CHANGE = 0x02B1;
        internal const int WM_TABLET_FIRST = 0x02c0;
        internal const int WM_TABLET_LAST = 0x02df;
        internal const int WM_CUT = 0x0300;
        internal const int WM_COPY = 0x0301;
        internal const int WM_PASTE = 0x0302;
        internal const int WM_CLEAR = 0x0303;
        internal const int WM_UNDO = 0x0304;
        internal const int WM_RENDERFORMAT = 0x0305;
        internal const int WM_RENDERALLFORMATS = 0x0306;
        internal const int WM_DESTROYCLIPBOARD = 0x0307;
        internal const int WM_DRAWCLIPBOARD = 0x0308;
        internal const int WM_PAINTCLIPBOARD = 0x0309;
        internal const int WM_VSCROLLCLIPBOARD = 0x030A;
        internal const int WM_SIZECLIPBOARD = 0x030B;
        internal const int WM_ASKCBFORMATNAME = 0x030C;
        internal const int WM_CHANGECBCHAIN = 0x030D;
        internal const int WM_HSCROLLCLIPBOARD = 0x030E;
        internal const int WM_QUERYNEWPALETTE = 0x030F;
        internal const int WM_PALETTEISCHANGING = 0x0310;
        internal const int WM_PALETTECHANGED = 0x0311;
        internal const int WM_HOTKEY = 0x0312;
        internal const int WM_PRINT = 0x0317;
        internal const int WM_PRINTCLIENT = 0x0318;
        internal const int WM_APPCOMMAND = 0x0319;
        internal const int WM_THEMECHANGED = 0x031A;
        internal const int WM_DWMCOMPOSITIONCHANGED = 0x031E;
        internal const int WM_HANDHELDFIRST = 0x0358;
        internal const int WM_HANDHELDLAST = 0x035F;
        internal const int WM_AFXFIRST = 0x0360;
        internal const int WM_AFXLAST = 0x037F;
        internal const int WM_PENWINFIRST = 0x0380;
        internal const int WM_PENWINLAST = 0x038F;
        internal const int WM_USER = 0x0400;
        internal const int WM_REFLECT = 0x2000;
        internal const int WM_APP = 0x8000;

        #endregion

        internal const int SIZE_RESTORED = 0;// ウィンドウがサイズ変更されました。ただし最小化または最大化ではありません。
        internal const int SIZE_MINIMIZED = 1;// ウィンドウが最小化されました。
        internal const int SIZE_MAXIMIZED = 2;// ウィンドウが最大化されました。
        internal const int SIZE_MAXSHOW = 3;// ある他のウィンドウが元のサイズに戻されたとき、すべてのポップアップウィンドウに送られます。
        internal const int SIZE_MAXHIDE = 4;// ある他のウィンドウが最大化されたとき、すべてのポップアップウィンドウに送られます。

        internal const UInt32 SWP_NOSIZE = 0x0001;
        internal const UInt32 SWP_NOMOVE = 0x0002;
        internal const UInt32 SWP_NOZORDER = 0x0004;
        internal const UInt32 SWP_NOREDRAW = 0x0008;
        internal const UInt32 SWP_NOACTIVATE = 0x0010;
        internal const UInt32 SWP_FRAMECHANGED = 0x0020;
        internal const UInt32 SWP_SHOWWINDOW = 0x0040;
        internal const UInt32 SWP_HIDEWINDOW = 0x0080;
        internal const UInt32 SWP_NOCOPYBITS = 0x0100;
        internal const UInt32 SWP_NOOWNERZORDER = 0x0200;
        internal const UInt32 SWP_NOSENDCHANGING = 0x0400;
        internal const UInt32 SWP_DRAWFRAME = SWP_FRAMECHANGED;
        internal const UInt32 SWP_NOREPOSITION = SWP_NOOWNERZORDER;
        internal const UInt32 SWP_DEFERERASE = 0x2000;
        internal const UInt32 SWP_ASYNCWINDOWPOS = 0x4000;

        internal const int DTT_COMPOSITED = 8192;
        internal const int DTT_GLOWSIZE = 2048;
        internal const int DTT_TEXTCOLOR = 1;

        internal const int BLACKNESS = 0x00000042; // 物理パレットのインデックス 0 に対応する色 (デフォルトは黒) で、コピー先の長方形を塗りつぶします。
        internal const int DSTINVERT = 0x550009;
        internal const int MERGECOPY = 0xC000CA;
        internal const int MERGEPAINT = 0x00BB0226; // コピー元の色を反転した色と、コピー先の色を、論理 OR 演算子で結合します。
        internal const int NOTSRCCOPY = 0x330008; // 色を反転して転送
        internal const int NOTSRCERASE = 0x1100A6;
        internal const int PATCOPY = 0x00F00021; // 指定したパターンをコピー先にコピーします。
        internal const int PATINVERT = 0x5A0049;
        internal const int PATPAINT = 0x00FB0A09; // 指定したパターンの色と、コピー元の色を反転した色を、論理 OR 演算子で結合し、さらにその結果を、コピー先の色と論理 OR 演算子で結合します。
        internal const int SRCAND = 0x8800C6; // 転送先の画像とAND演算して転送
        internal const int SRCCOPY = 0xCC0020; // そのまま転送
        internal const int SRCERASE = 0x440328;
        internal const int SRCINVERT = 0x660046; // 背景を反映して色を反転して転送
        internal const int SRCPAINT = 0x00EE0086; // 転送先の画像とOR演算して転送
        internal const int WHITENESS = 0x00FF0062; // 物理パレットのインデックス 1 に対応する色 (デフォルトは白) で、コピー先の長方形を塗りつぶします。

        internal const int WA_INACTIVE = 0;
        internal const int WA_ACTIVE = 1;
        internal const int WA_CLICKACTIVE = 2;

        internal const int HTCAPTION = 2;

        internal const int S_OK = 0;

        internal const int EM_SETRECT = 0xB3; // テキストを表示する領域を設定
        internal const int EM_GETLINECOUNT = 0xBA; // テキストの行数を取得する定数

        internal enum FileAttributesFlags : uint
        {
            FILE_ATTRIBUTE_ARCHIVE = 0x00000020,
            FILE_ATTRIBUTE_ENCRYPTED = 0x00004000,
            FILE_ATTRIBUTE_HIDDEN = 0x00000002,
            FILE_ATTRIBUTE_NORMAL = 0x00000080,
            FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000,
            FILE_ATTRIBUTE_OFFLINE = 0x00001000,
            FILE_ATTRIBUTE_READONLY = 0x00000001,
            FILE_ATTRIBUTE_SYSTEM = 0x00000004,
            FILE_ATTRIBUTE_TEMPORARY = 0x00000100
        }

        [Flags]
        internal enum ShellFileInfoFlags
        {
            SHGFI_ICON = 0x000000100,
            SHGFI_DISPLAYNAME = 0x000000200,
            SHGFI_TYPENAME = 0x000000400,
            SHGFI_ATTRIBUTES = 0x000000800,
            SHGFI_ICONLOCATION = 0x000001000,
            SHGFI_EXETYPE = 0x000002000,
            SHGFI_SYSICONINDEX = 0x000004000,
            SHGFI_LINKOVERLAY = 0x000008000,
            SHGFI_SELECTED = 0x000010000,
            SHGFI_ATTR_SPECIFIED = 0x000020000,
            SHGFI_LARGEICON = 0x000000000,
            SHGFI_SMALLICON = 0x000000001,
            SHGFI_OPENICON = 0x000000002,
            SHGFI_SHELLICONSIZE = 0x000000004,
            SHGFI_PIDL = 0x000000008,
            SHGFI_USEFILEATTRIBUTES = 0x000000010,
            SHGFI_ADDOVERLAYS = 0x000000020,
            SHGFI_OVERLAYINDEX = 0x000000040
        }

        internal enum ShellSpecialFolder
        {
            CSIDL_DESKTOP = 0x0000,
            CSIDL_INTERNET = 0x0001,
            CSIDL_PROGRAMS = 0x0002,
            CSIDL_CONTROLS = 0x0003,
            CSIDL_PRINTERS = 0x0004,
            CSIDL_PERSONAL = 0x0005,
            CSIDL_FAVORITES = 0x0006,
            CSIDL_STARTUP = 0x0007,
            CSIDL_RECENT = 0x0008,
            CSIDL_SENDTO = 0x0009,
            CSIDL_BITBUCKET = 0x000a,
            CSIDL_STARTMENU = 0x000b,
            CSIDL_DESKTOPDIRECTORY = 0x0010,
            CSIDL_DRIVES = 0x0011,
            CSIDL_NETWORK = 0x0012,
            CSIDL_NETHOOD = 0x0013,
            CSIDL_FONTS = 0x0014,
            CSIDL_TEMPLATES = 0x0015,
            CSIDL_COMMON_STARTMENU = 0x0016,
            CSIDL_COMMON_PROGRAMS = 0X0017,
            CSIDL_COMMON_STARTUP = 0x0018,
            CSIDL_COMMON_DESKTOPDIRECTORY = 0x0019,
            CSIDL_APPDATA = 0x001a,
            CSIDL_PRINTHOOD = 0x001b,
            CSIDL_ALTSTARTUP = 0x001d,
            CSIDL_COMMON_ALTSTARTUP = 0x001e,
            CSIDL_COMMON_FAVORITES = 0x001f,
            CSIDL_INTERNET_CACHE = 0x0020,
            CSIDL_COOKIES = 0x0021,
            CSIDL_HISTORY = 0x0022
        }

        internal enum HWND
        {
            HWND_TOP = 0, // 最前列に表示する
            HWND_BOTTOM = 1, // 最下層に表示する
            HWND_TOPMOST = -1, // 常に最前面に表示する
            HWND_NOTOPMOST = -2, // 常に最下層に表示する
        }

        internal enum RDW : int
        {
            INVALIDATE = 0x1,
            INTERNALPAINT = 0x2,
            ERASE = 0x4,
            VALIDATE = 0x8,
            NOINTERNALPAINT = 0x10,
            NOERASE = 0x20,
            NOCHILDREN = 0x40,
            ALLCHILDREN = 0x80,
            UPDATENOW = 0x100,
            ERASENOW = 0x200,
            FRAME = 0x400,
            NOFRAME = 0x800
        }

        internal enum SM
        {
            CXSCREEN = 0, // ディスプレイの幅
            CYSCREEN = 1, // 同、高さ
            CXVSCROLL = 2, // 垂直スクロールバーの矢印ビットマップの幅
            CYHSCROLL = 3, // 水平スクロールバーの矢印ビットマップの高さ
            CYCAPTION = 4, // キャプションの高さ
            CXBORDER = 5, // サイズ固定ウィンドウの境界線の、x方向の幅
            CYBORDER = 6, // 同、y方向の幅
            CXDLGFRAME = 7, // ダイアログフレームスタイルを持つウィンドウの枠線のx方向の幅
            CYDLGFRAME = 8, // 同、y方向の高さ
            CXHTHUMB = 10, // 水平スクロールバーのサムのビットマップの幅
            CYHTHUMB = 11, // 垂直スクロールバーのサムのビットマップの幅
            CYMENU = 15, // メニューバーの行の高さ
            CXFULLSCREEN = 16, // 最大化したときのクライアント領域の幅
            CYFULLSCREEN = 17, // 同、高さ
            CXMIN = 28, // ウィンドウの最小幅
            CYMIN = 29, // 同、高さ
            CXSIZE = 30, // タイトルバー内のビットマップの幅
            CYSIZE = 31, // 同、高さ
            CXSIZEFRAME = 32, // サイズ可変ウィンドウの境界線の、x方向の幅
            CYSIZEFRAME = 33, // 同、y方向の幅
            CXEDGE = 45,// 立体的なウィンドウの縁の幅と高さを取得します。SM_CXBORDER とSM_CYBORDER の 3D 版です。
            CYEDGE = 46,// 立体的なウィンドウの縁の幅と高さを取得します。SM_CXBORDER とSM_CYBORDER の 3D 版です。
            CXMINTRACK = 57, // ウィンドウの可能な最小幅
            CYMINTRACK = 58, // 同、高さ
            CXMAXTRACK = 59, // サイズ変更可能なウィンドウの最大幅
            CYMAXTRACK = 60, // 同、高さ
            CXMAXIMIZED = 61, // ディスプレイの最大幅
            CYMAXIMIZED = 62, // 同、高さ
            CLEANBOOT = 67 // Windowsを起動した方法 戻り値(0=通常起動、1=セーフモード、2=ネットワークのセーフモード)
        }

        [Flags]
        internal enum SHIL : int
        {
            SHIL_JUMBO = 0x0004,
            SHIL_EXTRALARGE = 0x0002
        }

        #endregion

        #region 静的変数

        internal static readonly Guid IID_IImageList = new("46EB5926-582E-4017-9FDF-E8998DAA0950");
        //internal static readonly Guid IID_IImageList2 = new Guid("192B9D83-50FC-457B-90A0-2B82A8B5DAE1");

        #endregion

        #region 構造体

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct SHFILEINFOW
        {
            internal IntPtr hIcon;
            internal int iIcon;
            internal uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            internal string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            internal string szTypeName;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            internal int left, top, right, bottom;
            internal RECT(int left, int top, int right, int bottom)
            {
                this.left = left;
                this.top = top;
                this.right = right;
                this.bottom = bottom;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct WINDOWPOS
        {
            internal IntPtr HWND;
            internal IntPtr HWNDAfterInsert;
            internal int x, y, cx, cy, flags;
        }

        internal struct NCCALCSIZE_PARAMS
        {
            internal RECT rgrc0, rgrc1, rgrc2;
            internal WINDOWPOS lppos;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct POINT
        {
            internal POINT(int x, int y)
            {
                this.x = x;
                this.y = y;
            }

            internal int x;
            internal int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct DTTOPTS
        {
            internal int dwSize;
            internal int dwFlags;
            internal int crText;
            internal int crBorder;
            internal int crShadow;
            internal int iTextShadowType;
            internal POINT ptShadowOffset;
            internal int iBorderSize;
            internal int iFontPropId;
            internal int iColorPropId;
            internal int iStateId;
            internal bool fApplyOverlay;
            internal int iGlowSize;
            internal int pfnDrawTextCallback;
            internal IntPtr lParam;
        }

        [StructLayout(LayoutKind.Sequential)]
        unsafe internal struct PAINTSTRUCT
        {
            internal IntPtr hdc;
            internal bool fErase;
            internal RECT rcPaint;
            internal bool fRestore;
            internal bool fIncUpdate;
            internal fixed byte rgbReserved[32];
        }

        internal struct POINTAPI
        {
            internal int x;
            internal int y;
        }

        internal struct MINMAXINFO
        {
            internal POINTAPI ptReserved;
            internal POINTAPI ptMaxSize;
            internal POINTAPI ptMaxPosition;
            internal POINTAPI ptMinTrackSize;
            internal POINTAPI ptMaxTrackSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct IMAGELISTDRAWPARAMS
        {
            internal int cbSize;
            internal IntPtr himl;
            internal int i;
            internal IntPtr hdcDst;
            internal int x;
            internal int y;
            internal int cx;
            internal int cy;
            internal int xBitmap;    // x offest from the upperleft of bitmap
            internal int yBitmap;    // y offset from the upperleft of bitmap
            internal int rgbBk;
            internal int rgbFg;
            internal int fStyle;
            internal int dwRop;
            internal int fState;
            internal int Frame;
            internal int crEffect;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct SHFILEINFO
        {
            internal IntPtr hIcon;
            internal int iIcon;
            internal int dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            internal string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            internal string szTypeName;
        }

        #endregion

        #region クラス

        [StructLayout(LayoutKind.Sequential)]
        internal class BITMAPINFO
        {
            internal int biSize;
            internal int biWidth;
            internal int biHeight;
            internal short biPlanes;
            internal short biBitCount;
            internal int biCompression;
            internal int biSizeImage;
            internal int biXPelsPerMeter;
            internal int biYPelsPerMeter;
            internal int biClrUsed;
            internal int biClrImportant;
            internal byte bmiColors_rgbBlue;
            internal byte bmiColors_rgbGreen;
            internal byte bmiColors_rgbRed;
            internal byte bmiColors_rgbReserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal class MARGINS
        {
            internal int cxLeftWidth, cxRightWidth, cyTopHeight, cyBottomHeight;
            internal MARGINS(int left, int top, int right, int bottom)
            {
                this.cxLeftWidth = left;
                this.cyTopHeight = top;
                this.cxRightWidth = right;
                this.cyBottomHeight = bottom;
            }
            internal MARGINS() { }
        }

        [StructLayout(LayoutKind.Sequential)]
        internal class DWM_BLURBEHIND
        {
            internal uint dwFlags;
            [MarshalAs(UnmanagedType.Bool)]
            internal bool fEnable;
            internal IntPtr hRegionBlur;
            [MarshalAs(UnmanagedType.Bool)]
            internal bool fTransitionOnMaximized;
            internal const uint DWM_BB_ENABLE = 0x00000001;
            internal const uint DWM_BB_BLURREGION = 0x00000002;
            internal const uint DWM_BB_TRANSITIONONMAXIMIZED = 0x00000004;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal class DWM_THUMBNAIL_PROPERTIES
        {
            internal uint dwFlags;
            internal RECT rcDestination;
            internal RECT rcSource;
            internal byte opacity;
            [MarshalAs(UnmanagedType.Bool)]
            internal bool fVisible;
            [MarshalAs(UnmanagedType.Bool)]
            internal bool fSourceClientAreaOnly;

            internal const uint DWM_TNP_RECTDESTINATION = 0x00000001;
            internal const uint DWM_TNP_RECTSOURCE = 0x00000002;
            internal const uint DWM_TNP_OPACITY = 0x00000004;
            internal const uint DWM_TNP_VISIBLE = 0x00000008;
            internal const uint DWM_TNP_SOURCECLIENTAREAONLY = 0x00000010;
        }

        #endregion

        #region インターフェース

        // interface COM IImageList
        [ComImportAttribute()]
        [GuidAttribute("46EB5926-582E-4017-9FDF-E8998DAA0950")]
        [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IImageList
        {
            [PreserveSig]
            int Add(IntPtr hbmImage, IntPtr hbmMask, ref int pi);

            [PreserveSig]
            int ReplaceIcon(int i, IntPtr hicon, ref int pi);

            [PreserveSig]
            int SetOverlayImage(int iImage, int iOverlay);

            [PreserveSig]
            int Replace(int i, IntPtr hbmImage, IntPtr hbmMask);

            [PreserveSig]
            int AddMasked(IntPtr hbmImage, int crMask, ref int pi);

            [PreserveSig]
            int Draw(ref IMAGELISTDRAWPARAMS pimldp);

            [PreserveSig]
            int Remove(int i);

            [PreserveSig]
            int GetIcon(int i, int flags, ref IntPtr picon);
        };

        #endregion

        #region APIメソッド

        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32", SetLastError = true)]
        internal static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = false)]
        internal static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern IntPtr SHGetFileInfo(string pszPath, int dwFileAttributes, ref SHFILEINFO psfi, int cbFileInfo, int uFlags);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        internal static extern IntPtr SHGetFileInfoW(string pszPath, FileAttributesFlags dwFileAttributes, ref SHFILEINFOW psfi, uint cbSizeFileInfo, ShellFileInfoFlags uFlags);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        internal static extern IntPtr SHGetFileInfoW(IntPtr idl, FileAttributesFlags dwFileAttributes, ref SHFILEINFOW psfi, uint cbSizeFileInfo, ShellFileInfoFlags uFlags);

        [DllImport("shlwapi.dll", EntryPoint = "PathFileExistsA")]
        internal static extern int PathFileExists(string lpszPath);

        [DllImport("Shell32.dll", CharSet = CharSet.Auto)]
        internal static extern int SHGetSpecialFolderLocation(IntPtr hwndOwner, ShellSpecialFolder folder, out IntPtr idl);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal extern static bool DestroyIcon(IntPtr handle);

        [DllImport("Comctl32.dll")]
        internal static extern bool ImageList_Destroy(IntPtr himl);

        [DllImport("gdiplus.dll", CharSet = CharSet.Unicode)]
        internal static extern int GdipLoadImageFromFile(string strFileName, ref IntPtr ipImage);

        [DllImport("user32.dll")]
        internal static extern int GetSystemMetrics(SM nIndex);

        [DllImport("User32.dll")]
        internal static extern bool ReleaseCapture();

        [DllImport("User32.dll")]
        internal static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("User32.dll")]
        internal static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);

        [DllImport("shell32.dll", EntryPoint = "ExtractIconEx", CharSet = CharSet.Auto)]
        internal static extern int ExtractIconEx([MarshalAs(UnmanagedType.LPTStr)] string file, int index, out IntPtr largeIconHandle, out IntPtr smallIconHandle, int icons);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);

        [DllImport("gdi32.dll")]
        internal static extern bool BitBlt(IntPtr hdcDst, int xDst, int yDst, int width, int height, IntPtr hdcSrc, int xSrc, int ySrc, int rasterOp);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        internal static extern uint GetShortPathName([MarshalAs(UnmanagedType.LPTStr)] string lpszLongPath, [MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpszShortPath, uint cchBuffer);

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern void DwmEnableBlurBehindWindow(IntPtr hWnd, DWM_BLURBEHIND pBlurBehind);

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern void DwmExtendFrameIntoClientArea(IntPtr hWnd, MARGINS pMargins);

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern bool DwmIsCompositionEnabled();

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern void DwmGetColorizationColor(out int pcrColorization, [MarshalAs(UnmanagedType.Bool)] out bool pfOpaqueBlend);

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern void DwmEnableComposition(bool bEnable);

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern IntPtr DwmRegisterThumbnail(IntPtr dest, IntPtr source);

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern void DwmUnregisterThumbnail(IntPtr hThumbnail);

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern void DwmUpdateThumbnailProperties(IntPtr hThumbnail, DWM_THUMBNAIL_PROPERTIES props);

        [DllImport("dwmapi.dll", PreserveSig = false)]
        internal static extern void DwmQueryThumbnailSourceSize(IntPtr hThumbnail, out Size size);

        [DllImport("User32.Dll")]
        internal static extern int GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("User32.Dll")]
        internal static extern bool SetWindowPos(IntPtr hWnd, HWND hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        [DllImport("gdi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern IntPtr CreateCompatibleDC(IntPtr hDC);

        [DllImport("gdi32.dll")]
        internal static extern IntPtr CreateDIBSection(IntPtr hdc, BITMAPINFO pbmi, uint iUsage, int ppvBits, IntPtr hSection, uint dwOffset);

        [DllImport("gdi32.dll", ExactSpelling = true)]
        internal static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);

        [DllImport("UxTheme.dll", CharSet = CharSet.Unicode)]
        internal static extern int DrawThemeTextEx(IntPtr hTheme, IntPtr hdc, int iPartId, int iStateId, string text, int iCharCount, int dwFlags, ref RECT pRect, ref DTTOPTS pOptions);

        [DllImport("gdi32.dll")]
        internal static extern bool BitBlt(IntPtr hdc, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, uint dwRop);

        [DllImport("gdi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool DeleteDC(IntPtr hdc);

        [DllImport("user32")]
        internal static extern int AdjustWindowRectEx(ref RECT lpRect, int dsStyle, int bMenu, int dwEsStyle);

        [DllImport("dwmapi.dll")]
        internal static extern int DwmDefWindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, out IntPtr result);

        [DllImport("User32.Dll")]
        internal static extern int ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        internal static extern IntPtr BeginPaint(IntPtr hWnd, out PAINTSTRUCT ps);

        [DllImport("user32.dll")]
        internal static extern IntPtr EndPaint(IntPtr hWnd, ref PAINTSTRUCT ps);

        [DllImport("user32.dll")]
        internal static extern bool UpdateWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

        [DllImport("shell32.dll", EntryPoint = "#727")]
        internal static extern int SHGetImageList(int iImageList, Guid riid, ref IImageList ppv);

        [DllImport("shell32.dll", EntryPoint = "#727")]
        internal static extern int SHGetImageList(SHIL iImageList, Guid riid, out IImageList ppv);

        //
        [DllImport("shell32.dll", EntryPoint = "#727")]
        internal static extern int SHGetImageList(SHIL iImageList, Guid riid, out IntPtr ppv);

        [DllImport("comctl32.dll", SetLastError = true)]
        internal static extern IntPtr ImageList_GetIcon(IntPtr himl, int i, int flags);

        #endregion

        #region メソッド

        internal static int RECTWIDTH(RECT rect)
        {
            return rect.right - rect.left;
        }

        internal static int RECTHEIGHT(RECT rect)
        {
            return rect.bottom - rect.top;
        }

        internal static int LoWord(int dwValue)
        {
            return dwValue & 0xFFFF;
        }

        internal static int HiWord(int dwValue)
        {
            return (dwValue >> 16) & 0xFFFF;
        }

        #endregion
    }
}