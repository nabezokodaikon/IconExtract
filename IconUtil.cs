using System.Drawing;
using System.Runtime.InteropServices;
using WinApi;
using static WinApi.WinApiMembers;

namespace IconExtract
{
    internal static class IconUtil
    {
        public static Image? GetSmallIconByFilePath(string filePath)
        {
            if (filePath == null)
            {
                throw new ArgumentNullException(nameof(filePath));
            }

            var sh = new WinApiMembers.SHFILEINFOW();
            var hSuccess = WinApiMembers.SHGetFileInfoW(filePath, 0, ref sh, (uint)Marshal.SizeOf(sh),
                                                           WinApiMembers.ShellFileInfoFlags.SHGFI_ICON |
                                                           WinApiMembers.ShellFileInfoFlags.SHGFI_SMALLICON);
            try
            {
                if (!hSuccess.Equals(IntPtr.Zero))
                {
                    using (var icon = Icon.FromHandle(sh.hIcon))
                    {
                        return icon.ToBitmap();
                    }
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                WinApiMembers.DestroyIcon(sh.hIcon);
            }
        }

        public static Image? GetExtraLargeIconByFilePath(string filePath)
        {
            if (filePath == null)
            {
                throw new ArgumentNullException(nameof(filePath));
            }

            var shinfo = new WinApiMembers.SHFILEINFO();
            var hSuccess = WinApiMembers.SHGetFileInfo(
                filePath,
                0,
                ref shinfo,
                (int)Marshal.SizeOf(typeof(WinApiMembers.SHFILEINFO)),
                (int)WinApiMembers.ShellFileInfoFlags.SHGFI_SYSICONINDEX);
            if (hSuccess.Equals(IntPtr.Zero))
            {
                return null;
            }

            var result = WinApiMembers.SHGetImageList(
                SHIL.SHIL_JUMBO,
                WinApiMembers.IID_IImageList,
                out IntPtr pimgList);

            try
            {
                if (result != WinApiMembers.S_OK)
                {
                    return null;
                }

                var hicon = WinApiMembers.ImageList_GetIcon(pimgList, shinfo.iIcon, 0);
                if (hicon.Equals(IntPtr.Zero))
                {
                    return null;
                }

                try
                {
                    using (var icon = Icon.FromHandle(hicon))
                    {
                        return icon.ToBitmap();
                    }
                }
                finally
                {
                    WinApiMembers.DestroyIcon(hicon);
                }
            }
            finally
            {
                WinApiMembers.ImageList_Destroy(pimgList);
            }
        }
    }
}
