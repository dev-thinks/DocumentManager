using System.IO;

namespace DocumentManager.Core.Converters
{
    public class Helper
    {
        public static void ClearDirectory(string folderName)
        {
            DelDirectory(folderName);

            Directory.Delete(folderName);
        }

        private static void DelDirectory(string folderName)
        {
            var dir = new DirectoryInfo(folderName);

            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }

            foreach (DirectoryInfo di in dir.GetDirectories())
            {
                DelDirectory(di.FullName);
                di.Delete();
            }
        }
    }
}
