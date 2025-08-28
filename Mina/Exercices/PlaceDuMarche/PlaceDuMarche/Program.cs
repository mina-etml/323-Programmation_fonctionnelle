using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlaceDuMarche
{
    class WatermelonsSeller
    {
        public string Name { get; set; }
        public string Location { get; set; }
        public string Quantity { get; set; }
    }

    internal class Program
    {

        /*
            static void Main(string[] args)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                const string filePath = "C:\\Users\\pk60ftv\\Documents\\I323\\323-Programmation_fonctionnelle\\exos\\marché\\Place du marché.xlsx";

                List<string[]> locationInfos = ReadExcelTable(filePath, 0);
                List<string[]> productsInfos = ReadExcelTable(filePath, 1);

                //Trouver le nombre de vendeur de pêches
                int nbPeachSeller = FindNumberOfProductSeller(productsInfos, "Pêches");

                Console.WriteLine("Nombre de vendeurs de pêches: " + nbPeachSeller);

                //Trouver qui vend le plus de pastèques et ses informations

                int mostWatermelon = 0;
                var watermelonSellers = GetSellers(productsInfos, "Pastèques");
                WatermelonsSeller mostWatermelonSeller = new WatermelonsSeller();

                foreach (var seller in watermelonSellers)
                { 
                    int quantity = int.Parse(seller.Quantity);

                    if (quantity > mostWatermelon)
                    {
                        mostWatermelon = int.Parse(seller.Quantity);
                        mostWatermelonSeller = seller;
                    }
                }

                Console.WriteLine("C'est " + mostWatermelonSeller.Name + " qui a le plus de pastèques (stand " + mostWatermelonSeller.Location + ", " + mostWatermelonSeller.Quantity + " pièces)");


                Console.ReadLine();
            }
    */
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            const string filePath = @"C:\Users\pk60ftv\Documents\I323\323-Programmation_fonctionnelle\exos\marché\Place du marché.xlsx";

            var locationInfos = ReadExcelTable(filePath, 0);
            var productsInfos = ReadExcelTable(filePath, 1);

            var nbPeachSeller = FindNumberOfProductSeller(productsInfos, "Pêches");
            Console.WriteLine($"Nombre de vendeurs de pêches: {nbPeachSeller}");

            var watermelonSellers = GetSellers(productsInfos, "Pastèques");

            var mostWatermelonSeller = watermelonSellers
                .Where(s => int.TryParse(s.Quantity, out _))
                .OrderByDescending(s => int.Parse(s.Quantity))
                .FirstOrDefault();

            if (mostWatermelonSeller != null)
            {
                Console.WriteLine($"C'est {mostWatermelonSeller.Name} qui a le plus de pastèques (stand {mostWatermelonSeller.Location}, {mostWatermelonSeller.Quantity} pièces)");
            }

            Console.ReadLine();
        }

        /*
        static List<string[]> ReadExcelTable(string filePath, short pageNumber)
        {
            var info = new List<string[]>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[pageNumber];

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int row = 1; row <= rowCount; row++)
                {
                    var rowData = new List<string>();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Text;

                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            rowData.Add(cellValue);
                        }
                    }

                    info.Add(rowData.ToArray());
                }
            }
            return info;
        }
        */
        static List<string[]> ReadExcelTable(string filePath, short pageNumber)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[pageNumber];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                return Enumerable.Range(1, rowCount)
                    .Select(row =>
                        Enumerable.Range(1, colCount)
                            .Select(col => worksheet.Cells[row, col].Text)
                            .Where(cell => !string.IsNullOrWhiteSpace(cell))
                            .ToArray()
                    )
                    .ToList();
            }
        }

        /*
                static int FindNumberOfProductSeller(List<string[]> lists, string produtSaled)
                {
                    int nuberOfProductSeller = 0;

                    foreach(string[] row in lists)
                    {
                        foreach(string info in row)
                        {
                            if (info == produtSaled)
                            {
                                nuberOfProductSeller++;
                            }

                        }
                    }

                    return nuberOfProductSeller;            
                }

        */
        static int FindNumberOfProductSeller(List<string[]> rows, string productSaled)
            {
                return rows.Count(row => row.Contains(productSaled));
            }
        /*

                static List<WatermelonsSeller> GetSellers(List<string[]> rows, string produtSaled)
                {
                    var watermelonsSellers = new List<WatermelonsSeller>();

                    foreach (string[] row in rows)
                    {
                        if (row.Contains(produtSaled))
                        {
                            var seller = new WatermelonsSeller
                            {
                                Location = row[0],
                                Name = row[1],
                                Quantity = row[3]
                            };

                            watermelonsSellers.Add(seller);

                        }
                    }

                    return watermelonsSellers;
                }
            }
        }
        */

        static List<WatermelonsSeller> GetSellers(List<string[]> rows, string productSaled)
        {
            return rows
                .Where(row => row.Contains(productSaled))
                .Select(row => new WatermelonsSeller
                {
                    Location = row[0],
                    Name = row[1],
                    Quantity = row[3]
                })
                .ToList();
        }
    }
}

