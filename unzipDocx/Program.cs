using System.IO.Compression;
using System.Runtime.ConstrainedExecution;

namespace unzipDocx
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Write your .docx file path:");
            String path = Console.ReadLine();
            Console.WriteLine("Write the path where you want your file:");
            String destinationPath = Console.ReadLine();
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(path))
                {
                    Console.WriteLine(archive.ToString());
                    archive.ExtractToDirectory(destinationPath + "unzip", true);
                }
            }catch(Exception e)
            {
                Console.WriteLine($"failed {e}");
            }

        }
    }
}