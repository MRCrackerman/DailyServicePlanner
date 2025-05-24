using System;
using ClosedXML.Excel;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Globalization;

class ExcelHandler
    {

        
        
        public ExcelHandler(){

        }

        public DataTable SaveToDataTable(string excelName){

            string excelPath = @".\Excels\" + excelName + ".xlsx";             
            DataTable table = new DataTable();

            Console.WriteLine("edo");

            using (var workbook = new XLWorkbook(excelPath))
                {

                    var worksheet = workbook.Worksheet(1); 
                    var firstRow = worksheet.FirstRowUsed();

                    foreach (var cell in firstRow.Cells())
                    {
                        table.Columns.Add(cell.Value.ToString());
                    }

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        DataRow dataRow = table.NewRow();
                       
                        for (int i = 0; i < row.Cells().Count(); i++)
                        {
                            dataRow[i] = row.Cell(i + 1).Value;  
                        }

                        table.Rows.Add(dataRow);
                    }
                    
                }
                                
                return table; 
                               
        }


        //tha dexete san isodo to antikimeno tis imeras kai tha ftianxei excel gia thn mera ekini
        public void CreateDateExcel(DayService serviceDay)
        {
            string dateName = serviceDay.date.Replace("/" , "");
            string[] parts = dateName.Split(" ");
            string filePath = @".\Excels\Month\" + serviceDay.fullDaySt + ".xlsx";

            //string filePath = @".\Excels\Month\" + parts[0] + ".xlsx";


            string templatepath = @".\Excels\Template.xlsx";
        
            using (var TemplateWorkbook = new XLWorkbook(templatepath))
            {

                using (var workbook = new XLWorkbook()){
                    
                    foreach (var sheet in TemplateWorkbook.Worksheets)
                    {
                        //var newSheet = workbook.Worksheets.Add(sheet.Name);
                        sheet.CopyTo(workbook , sheet.Name);
                    }

                    var worksheet = workbook.Worksheet(1);

                    // Add headers to the first row
                    worksheet.Cell(1, 1).Value = serviceDay.dayTitle;
                    worksheet.Cell(3, 2).Value = serviceDay.easRank + " " + serviceDay.easName;
                    worksheet.Cell(4, 2).Value = serviceDay.aydmRank + " " + serviceDay.aydmName;
                    worksheet.Cell(5, 2).Value = serviceDay.easRank + " " + serviceDay.easName;
                    worksheet.Cell(6, 2).Value = serviceDay.arxifilakasRank + " " + serviceDay.arxifilakasName;

                    worksheet.Cell(13, 2).Value = serviceDay.camera1;
                    Console.Write(serviceDay.camera1);
                    worksheet.Cell(14, 2).Value = serviceDay.camera2;
                    worksheet.Cell(15, 2).Value = serviceDay.camera3;
                    worksheet.Cell(23, 2).Value = serviceDay.per1a + " - " + serviceDay.per1b;
                    worksheet.Cell(24, 2).Value = serviceDay.per2a + " - " + serviceDay.per2b;
                    
                    worksheet.Cell(7, 2).Value = serviceDay.odigosEpif; // tha mpainei me diko tou excel
                    worksheet.Cell(10, 2).Value = serviceDay.esteiatoras;
                    worksheet.Cell(11, 2).Value = serviceDay.mageiras;
                    worksheet.Cell(16,2).Value = serviceDay.oksigonoRank + " " + serviceDay.oksigonoName;
                    
                    {
                        if(serviceDay.adeiaTYL.Count > 0){
                            int i = serviceDay.adeiaTYL.Count();

                            int k = 37;
                            for(int j = 0 ; j< i; j++){
                                worksheet.Cell(k + j, 2).Value = serviceDay.adeiaTYL[j];
                            }
                        }

                        if(serviceDay.adeiaBEB.Count > 0){
                            int i = serviceDay.adeiaBEB.Count();

                            int k = 46;
                            for(int j = 0 ; j< i; j++){
                                worksheet.Cell(k + j, 2).Value = serviceDay.adeiaBEB[j];
                            }
                        }
                        
                    }

                    {

                        int i = serviceDay.eksodosTYL.Count();

                        if(i>0){
                            int k = 37;
                            for(int j = 0 ; j< i; j++){
                                worksheet.Cell(k + j, 1).Value = serviceDay.eksodosTYL[j];
                            }
                        }

                        i = serviceDay.eksodosBEB.Count();

                        if(i>0){
                            int k = 46;
                            for(int j = 0 ; j< i; j++){
                                worksheet.Cell(k + j, 1).Value = serviceDay.eksodosBEB[j];
                            }
                        }
                        
                    }
                    
                    //ores
                    {
                        worksheet.Cell(13, 4).Value = serviceDay.oresk1;
                        worksheet.Cell(14, 4).Value = serviceDay.oresk2;
                        worksheet.Cell(15, 4).Value = serviceDay.oresk3;

                        worksheet.Cell(23, 4).Value = serviceDay.oresp1;
                        worksheet.Cell(24, 4).Value = serviceDay.oresp2;
                       
                    }
                    
                    
                    //worksheet.Cell(6, 2).Value = serviceDay.arxifilakasName;
                   //worksheet.Cell(6, 2).Value = serviceDay.arxifilakasName;


                    // Save the workbook to the specified file path
                    workbook.SaveAs(filePath);


                }
                // Add a new worksheet to the workbook
            }
        }

        public void WriteTable(DataTable table){
            foreach (DataRow row in table.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write(item + "\t");
                }
                Console.WriteLine();
            }
        }

    }

