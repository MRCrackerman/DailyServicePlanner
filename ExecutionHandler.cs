using System;
using System.Data;
using System.Collections.Generic;
using System.Windows;

//dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained -o "My App"

namespace _10Tim;

class ExecutionHandler{

    public  ExcelHandler excel;
    public DataTable eas;
    public DataTable aydm;
    public DataTable arxifilakas;
    public DataTable oksigono;
    public DataTable upiresies;
    public DataTable soldiers;
    public List<DayService> ListWithService;

    public ExecutionHandler()
    { 
        excel = new ExcelHandler();
        
    }

    public void CreateMonth()
    {
        InitTablesFromExcels();
        CheckExcels();
        CreateDayList();

        CreateDaysExcel();

        

        MessageBox.Show("Επίτυχης Δημιούργια Μήνα!");
    }


    private void InitTablesFromExcels()
    {
        ListWithService = new List<DayService>();
        
        eas = new DataTable();
        eas = excel.SaveToDataTable("ΕΑΣ");
        
        aydm = new DataTable();
        aydm = excel.SaveToDataTable("ΑΥΔΜ-ΒΕΒ");
        
        /*πως θα μπορούσε να λειτουργεί το try σε περίπτωση που δεν υπάρχει το excel
        try{
            aydm = excel.SaveToDataTable("Aydm");
        }
        catch(System.IO.FileNotFoundException ex){
            MessageBox.Show("To Excel ΑΥΔΜ βρέθηκε!!!");
        }

        or------------------------------------
        
        try{
            aydm = excel.SaveToDataTable("Aydm");
        }
        catch(System.IO.FileNotFoundException ex){
            CreateDummyTable(Aydm);
        }

        */
    
        arxifilakas = new DataTable();
        arxifilakas = excel.SaveToDataTable("ΑΡΧΙΦΥΛΑΚΑΣ");
        
        oksigono = new DataTable();
        oksigono = excel.SaveToDataTable("ΟΞΥΓΟΝΟ-ΒΕΒ");
        
        upiresies = new DataTable();
        upiresies = excel.SaveToDataTable("Υπηρεσίες");

        FormUpiresiesTable();
    }

    private void FormUpiresiesTable(){
        
        upiresies.Rows.RemoveAt(0);
        soldiers = new DataTable();
        soldiers.Columns.Add(upiresies.Columns[0].ColumnName, upiresies.Columns[0].DataType);
        
        foreach (DataRow row in upiresies.Rows)
        {
            soldiers.Rows.Add(row[0]);
        }

        //upiresies.Columns.RemoveAt(0);

        /*na emfanizo tous pinakes
        foreach (DataRow row in soldiers.Rows)
        {
            MessageBox.Show(row[0].ToString());
        }

        
        foreach (DataRow row in upiresies.Rows)
        {
            MessageBox.Show(row[0].ToString());
        }

        foreach (DataColumn column in upiresies.Columns){
            MessageBox.Show(upiresies.Rows[0][column].ToString());
        }
        */
    }

    private void CheckExcels()
    {
        //na tsekarei an ta excel einai idio megethos
        //an den einai na girnaei sfalma kai na termatizete to execution
    }

    private void CreateDayList()
    {
        int month = eas.Rows.Count;

        for (int i = 0; i < month; i++){

            DayService day = new DayService(eas.Rows[i] , aydm.Rows[i] , arxifilakas.Rows[i] , oksigono.Rows[i]);

            ServiceCalculator(day , i);

            ListWithService.Add(day);
        }
    }

    private void CreateDaysExcel()
    {
        //MessageBox.Show(ListWithService.Count.ToString());
        foreach(DayService day in ListWithService)
        {
            excel.CreateDateExcel(day);
        }
    }

    private void ServiceCalculator(DayService day , int i){
        
        //DataColumn imera = upiresies.Columns[i]; 
 
        foreach(DataRow imera in upiresies.Rows){
            
            string upiresia = imera[i + 3].ToString();
            string onoma = imera[1].ToString() + " " +  imera[2].ToString();
/*
            if (soldiers.Rows.Count > i) // Έλεγχος αν υπάρχει η γραμμή i
            {
                //onoma = soldiers.Rows[0][i].ToString(); // Διαβάζουμε την πρώτη στήλη της γραμμής i
                onoma = imera[1].ToString() + " " +  imera[2].ToString();
                //MessageBox.Show($"o stratiotis {onoma}");    //0
                ii=ii+1;
            }
*/
            //MessageBox.Show($"o stratiotis {onoma} exei upiresia {upiresia}"); //1
            if(upiresia == "Ε"){
                if(imera[0].ToString() == "ΒΕΒ"){
                    day.AddEksodouxoBEB(onoma);
                }
                else{
                    day.AddEksodouxoTYL(onoma); 
                }
                   
            } 
            else if(upiresia == "Α" || upiresia == "ΤΙΜ"){
                if(imera[0].ToString() == "ΒΕΒ"){
                    day.AddAdeiouxoBEB(onoma);   
                }
                else{
                    day.AddAdeiouxoTYL(onoma); 
                }
                   
            }else if(upiresia == "Κ1"){
                day.AddCamera1(onoma);
            }else if(upiresia == "Κ2"){
                day.AddCamera2(onoma);
            }else if(upiresia == "Κ3"){
                day.AddCamera3(onoma);
            }else if(upiresia == "Π1"){
                day.AddPeripolo1(onoma);
            }else if(upiresia == "Π2"){
                day.AddPeripolo2(onoma);
            }else if(upiresia == "Ο/Ε"){
                day.AddOdigo(onoma);
            }else if(upiresia == "Μ"){
                day.AddMageira(onoma);
            }else if(upiresia == "ΕΣΤ"){
                day.Addesteiatoras(onoma);
            }
        }

        day.ValidateDay();
    }
    
    /*
        Dictionary<string, DataRow> labeledTable = upiresies.AsEnumerable()
    .ToDictionary(row => row.Field<string>("Label")!, row => row);

        int tableSize = labeledTable
*/

    //Create empty table if no excel found
    //το δεύτερο try στα σχόλια
    //Σε περίπτωση που δεν βρεί ένα excel να φτίαχνει τον πίνακα με παύλες παντού έτσι ώστε να μην σκάει απλως
    private void CreateDummyTable(DataTable table){
        for (int i = 0; i < 4; i++)
        {
            table.Columns.Add("Column" + (i + 1), typeof(string));
        }

        // Γέμισμα του πίνακα με 30 γραμμές που περιέχουν '-'
        for (int i = 0; i < 30; i++)
        {
            DataRow row = table.NewRow();
            for (int j = 0; j < 4; j++)
            {
                row[j] = "-";
            }
            table.Rows.Add(row);
        }
    }
}