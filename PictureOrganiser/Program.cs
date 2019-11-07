using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;

namespace PictureOrganiser
{
    public static class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"------ Organisation Mode enabled - {DateTime.Now} ------");


            var dob = new DateTime(2018, 11, 10);

            var monthsToProcess = GetMonthsToProcess(dob, 12);

            ProcessPictures(monthsToProcess);

            Console.WriteLine("press the any key to exit");
            Console.ReadKey();
        }

        private static void ProcessPictures(Dictionary<int, (DateTime startDate, DateTime EndDate)> monthsToProcess)
        {
            var sourceFolder = ConfigurationManager.AppSettings["SourcePath"];
            Console.WriteLine($"Files being read from {sourceFolder}");

            List<string> pictures = Directory.GetFiles(sourceFolder, "*.jpg", SearchOption.AllDirectories).ToList();
            pictures.AddRange(Directory.GetFiles(sourceFolder, "*.jpeg", SearchOption.AllDirectories).ToList());

            string[] movies = Directory.GetFiles(sourceFolder, "*.mp4", SearchOption.AllDirectories);

            Console.WriteLine($"Proccessing {pictures.Count} pictures");
            foreach (var picturePath in pictures)
            {
                var creationDate = GetCreationDate(picturePath);
                if (creationDate == null)
                {
                    Console.WriteLine($"could not find origin date for file {Path.GetFileName(picturePath)}");
                    continue;
                }

                var result = monthsToProcess.SingleOrDefault(month => month.Value.startDate <= creationDate && month.Value.EndDate >= creationDate);
                ProcessFile(picturePath, result.Key);
            }

            Console.WriteLine($"Proccessing {movies.Length} movies");
            foreach (var moviePath in movies)
            {
                var file = new FileInfo(moviePath);

                GetExtendedProperties(moviePath).TryGetValue("Media created", out var temp);


                string dateString = Regex.Replace(temp, @"[\u0009\u200F\u200E\u000D]", "");
                if (string.IsNullOrWhiteSpace(dateString))
                {
                    // fallback on date Modified )
                    GetExtendedProperties(moviePath).TryGetValue("Date modified", out temp);
                    dateString = Regex.Replace(temp, @"[\u0009\u200F\u200E\u000D]", "");
                }

                if (string.IsNullOrWhiteSpace(dateString))
                {
                    Console.WriteLine($"could not find origin date for file {Path.GetFileName(moviePath)}");
                    continue;
                }

                var creationDate = Convert.ToDateTime(dateString);
                var result = monthsToProcess.SingleOrDefault(month => month.Value.startDate <= creationDate && month.Value.EndDate >= creationDate);
                ProcessFile(moviePath, result.Key);
            }
        }

        private static void ProcessFile(string file, int month)
        {
            try
            {
                var destinationFolder = ConfigurationManager.AppSettings["DestinationFolder"];

                if (!Directory.Exists(destinationFolder))
                    throw new Exception($"destination folder ({destinationFolder}) does not exists");

                var path = Path.Combine(destinationFolder, month.ToString());

                Directory.CreateDirectory(path);

                var fileName = Path.GetFileName(file);

                File.Move(file, Path.Combine(path, fileName));
            }
            catch (IOException)
            {
                Console.WriteLine($"file {Path.GetFileName(file)} already in target folder");
            }

        }


        private static DateTime? GetCreationDate(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var img = BitmapFrame.Create(fs);
                var metaData = (BitmapMetadata)img.Metadata;
                //DateTaken
                string fileName = metaData.Title;
                string date = metaData.DateTaken;
                if (!string.IsNullOrWhiteSpace(date))
                {
                    return DateTime.Parse(date);
                }
                else
                {
                    var extendedProperties = GetExtendedProperties(path);
                    extendedProperties.TryGetValue("Date modified", out var temp);
                    var dateString = Regex.Replace(temp, @"[\u0009\u200F\u200E\u000D]", "");
                    if (!string.IsNullOrWhiteSpace(dateString))
                    {
                        return DateTime.Parse(dateString);
                    }
                }
            }

            return null;
        }

        private static Dictionary<string, string> GetExtendedProperties(string filePath)
        {
            var directory = Path.GetDirectoryName(filePath);
            var shellFolder = GetShell32NameSpaceFolder(directory);
            var fileName = Path.GetFileName(filePath);
            var folderitem = shellFolder.ParseName(fileName);
            var dictionary = new Dictionary<string, string>();
            var i = -1;
            while (++i < 320)
            {
                var header = shellFolder.GetDetailsOf(null, i);
                if (String.IsNullOrEmpty(header)) continue;
                var value = shellFolder.GetDetailsOf(folderitem, i);
                if (!dictionary.ContainsKey(header)) dictionary.Add(header, value);
            }

            Marshal.ReleaseComObject(shellFolder);
            return dictionary;
        }

        private static Shell32.Folder GetShell32NameSpaceFolder(Object folder)
        {
            Type shellAppType = Type.GetTypeFromProgID("Shell.Application");

            var shell = Activator.CreateInstance(shellAppType);
            return (Shell32.Folder)shellAppType.InvokeMember("NameSpace",
            System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { folder });
        }

        public static Dictionary<int, (DateTime startDate, DateTime EndDate)> GetMonthsToProcess(DateTime start, int totalMonths)
        {
            var listOfMonths = new Dictionary<int, (DateTime startDate, DateTime EndDate)>();

            for (int month = 0; month < totalMonths; month++)
            {
                listOfMonths.Add(month, (start.AddMonths(month), start.AddMonths(month + 1)));

            }

            return listOfMonths;
        }

    }
}
