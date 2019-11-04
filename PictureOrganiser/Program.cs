using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows.Media.Imaging;

namespace PictureOrganiser
{
    class Program
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

            string[] pictures = Directory.GetFiles(sourceFolder, "*.jpg");
            string[] movies = Directory.GetFiles(sourceFolder, "*.mp4");

            Console.WriteLine($"Proccessing {pictures.Length} pictures");
            foreach (var picturePath in pictures)
            {
                var creationDate = GetCreationDate(picturePath);

                var result = monthsToProcess.SingleOrDefault(month => month.Value.startDate <= creationDate && month.Value.EndDate >= creationDate);
                ProcessFile(picturePath, result.Key);
            }

            Console.WriteLine($"Proccessing {movies.Length} movies");
        }

        private static void ProcessFile(string file, int month)
        {
            var destinationFolder = ConfigurationManager.AppSettings["DestinationFolder"];

            if (!Directory.Exists(destinationFolder))
                throw new Exception($"destination folder ({destinationFolder}) does not exists");

            var path = Path.Combine(destinationFolder, month.ToString());

            Directory.CreateDirectory(path);

            var fileName = Path.GetFileName(file);

            File.Move(file, Path.Combine(path, fileName));
        }


        private static DateTime? GetCreationDate(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                BitmapSource img = BitmapFrame.Create(fs);
                BitmapMetadata md = (BitmapMetadata)img.Metadata;

                //DateTaken
                string fileName = md.Title;
                string date = md.DateTaken;
                if (!string.IsNullOrEmpty(date))
                {
                    return DateTime.Parse(date);
                }
            }

            return null;
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
