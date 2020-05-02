using System.IO;
using System.Text;

namespace BARS_GrupTest
{
    internal static class Helper
    {
        internal static string CurrentDirectopy { get; } = Directory.GetCurrentDirectory();

        internal static string GetJsonFromFileStream(FileStream stream)
        {
            byte[] output = new byte[stream.Length];
            stream.Read(output, 0, output.Length);
            string jsonFromFile = Encoding.Default.GetString(output);
            return jsonFromFile;
        }

        internal static string GetJsonFromFile(string FileName)
        {
            using (var stream = new FileStream(FileName, FileMode.Open, FileAccess.Read))
            {
                return GetJsonFromFileStream(stream);
            }
        }

        internal static void SaveJsonFromFromStream(FileStream stream, string value)
        {
            var bytes = Encoding.Default.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        internal static void SaveJsonToFile(string FileName, string value)
        {
            try
            {
                using (var stream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    SaveJsonFromFromStream(stream, value);
                }
            }
            catch { }
        }
    }
}
