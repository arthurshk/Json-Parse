using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using SwansonParser;
public class Program
{

    public static void Main(string[] args)
    {
        string String = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=Products;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        var swansonParser1 = new SwansonParser1(new ProductRepository(String));
        swansonParser1.RunSwansonParser();
    }
 
}