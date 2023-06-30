﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace SwansonParser
{
    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var baseUrl = "https://www.swansonvitamins.com/ncat1/Vitamins+and+Supplements/ncat2/Multivitamins/ncat3/Multivitamins+with+Iron/q";
                var pageCount = 3;
                var products = ParseSwansonWebsite(baseUrl, pageCount);
                DisplayProductData(products);
                CreateExcelFile(products);
                Console.WriteLine("Excel file created successfully.");
        }
        public static List<Product> ParseSwansonWebsite(string baseUrl, int pageCount)
        {
            var allProducts = new List<Product>();

            for (int i = 1; i <= pageCount; i++)
            {
                var url = $"{baseUrl}?page={i}";

                using (var client = new HttpClient())
                {
                    var html = client.GetStringAsync(url).Result;
                    File.WriteAllText("page.html", html);

                    var pattern = @"adobeRecords"":(.+),""topProduct";
                    var matches = Regex.Matches(html, pattern);
                    Console.WriteLine(matches.Count);

                    if (matches.Count > 0)
                    {
                        var json = matches[0].Groups[1].Value;
                        var products = JsonSerializer.Deserialize<List<Product>>(json);
                        allProducts.AddRange(products);
                    }
                }
            }

            return allProducts;
        }
        public static void DisplayProductData(List<Product> products)
        {
            Console.WriteLine("Product Data:");
            Console.WriteLine("-------------");
            foreach (var product in products)
            {
                Console.WriteLine($"Number: {product.Number}");
                Console.WriteLine($"Title: {product.Title}");
                Console.WriteLine($"Vendor: {product.Vendor}");
                Console.WriteLine($"Price: {product.Price}");
                Console.WriteLine();
            }
        }
        public static void CreateExcelFile(List<Product> products)
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "products.xlsx");

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Products");
                worksheet.Cells[1, 1].Value = "Number";
                worksheet.Cells[1, 2].Value = "Title";
                worksheet.Cells[1, 3].Value = "Vendor";
                worksheet.Cells[1, 4].Value = "Price";

                for (int i = 0; i < products.Count; i++)
                {
                    var product = products[i];
                    var row = i + 2;

                    worksheet.Cells[row, 1].Value = product.Number;
                    worksheet.Cells[row, 2].Value = product.Title;
                    worksheet.Cells[row, 3].Value = product.Vendor;
                    worksheet.Cells[row, 4].Value = product.Price;
                }
                worksheet.Cells.AutoFitColumns();
                package.Save();
            }
        }
    }
}