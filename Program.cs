using System.Drawing.Imaging;
using System.Reflection;

namespace IconExtract
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine(@"
1: Icon size. ('s' or 'l')
2: EXE file path.
");
                Console.ReadLine();
                return;
            }

            var size = args[0];
            var filePath = args[1];

            var assembly = Assembly.GetExecutingAssembly();
            if (assembly == null) 
            {
                return;
            }

            var exportDirPath = Path.GetDirectoryName(assembly.Location);
            if (exportDirPath == null)
            {
                return;
            }

            if (!Directory.Exists(exportDirPath))
            { 
                Directory.CreateDirectory(exportDirPath);
            }

            Console.WriteLine(exportDirPath);

            var exportFileName = string.Format("{0}\\{1}.png", exportDirPath, Path.GetFileNameWithoutExtension(filePath));
            var exportFilePath = Path.Combine(exportDirPath, exportFileName);

            if (size == "s")
            {
                var icon = IconUtil.GetSmallIconByFilePath(filePath);
                if (icon != null)
                {
                    using (var fs = new FileStream(exportFilePath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
                    {
                        icon.Save(fs, ImageFormat.Png);
                    }
                }
            }
            else if (size == "l")
            {
                var icon = IconUtil.GetExtraLargeIconByFilePath(filePath);
                if (icon != null)
                {
                    using (var fs = new FileStream(exportFilePath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
                    {
                        icon.Save(fs, ImageFormat.Png);
                    }
                }
            }
            else
            {
                Console.WriteLine(@"
1: Icon size. ('s' or 'l')
2: EXE file path.
");
                Console.ReadLine();
            }
        }
    }
}