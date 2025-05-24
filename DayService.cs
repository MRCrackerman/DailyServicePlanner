using System;
using System.Data;
using System.Globalization;
using System.Collections.Generic;
using System.Windows;
using DocumentFormat.OpenXml.Office.Y2022.FeaturePropertyBag;
using DocumentFormat.OpenXml.Office.CoverPageProps;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

class DayService{

    //Σόρρυ για τα greeklish στα ονόματα άμα τον ξαναδεί κάποιός τον κώδικα
    public string date; //6/06/2024
    public string day; //ΠΕΜΠΤΗΣ
    public string fullDaySt; //"06 ΙΟΥΝΙΟΥ 2024"
    public string dayTitle; //"ΚΑΤΑΣΤΑΣΗ ΥΠΗΡΕΣΙΩΝ ΣΤΡΔΟΥ "ΜΑΧΑΙΡΑ" ΤΗΣ ΠΕΜΠΤΗΣ 06 ΙΟΥΝΙΟΥ 2024"
    public string easName;
    public string easRank;
    public string aydmName;
    public string aydmRank;
    public string arxifilakasName;
    public string arxifilakasRank;
    public string oksigonoRank;
    public string oksigonoName;

    //οι Υπηρεσίες
    public string camera1 = "";
    public string camera2 =  "";
    public string camera3 = "";
    public string per1a = ""; // πρώτο νούμερο α στρατιώτης
    public string per1b = ""; // πρώτο νούμερο β στρατιώτης
    public string per2a = ""; // δεύτερο νούμερο α στρατιώτης
    public string per2b = ""; // δεύτερο νούμερο β στρατιώτης
    public string mageiras = ""; // μάγειρας
    public string esteiatoras = ""; //εστειάτορας
    public List<string> eksodosBEB; 
    public List<string> adeiaBEB;
    public List<string> eksodosTYL; 
    public List<string> adeiaTYL;
    public string odigosEpif = ""; 
    public bool papaduan1 = true;
    public bool papaduan2 = true;
    public string oresk1 = "-";
    public string oresk2 = "-";
    public string oresk3 = "-";
    public string oresp1 = "-";
    public string oresp2 = "-";
    public string oresthal1 = "-";
    public string oresthal2 = "-";

    public DayService(){ }

    public DayService(DataRow eas, DataRow aydm, DataRow arxifilakas, DataRow oksigonas)
    {
        day = eas[0].ToString();
        date =eas[1].ToString();
        easName = (string)eas[3];
        easRank = eas[2].ToString();
        aydmName = aydm[3].ToString();
        aydmRank = aydm[2].ToString();
        arxifilakasName = arxifilakas[3].ToString();
        arxifilakasRank = arxifilakas[2].ToString();
        oksigonoName = oksigonas[3].ToString();
        oksigonoRank = oksigonas[2].ToString();

        DateToString();

        camera1 = "";
        camera2 = "";
        eksodosBEB = new List<string>(); 
        adeiaBEB = new List<string>(); 

        eksodosTYL = new List<string>(); 
        adeiaTYL = new List<string>(); 
    }

    public void ValidateDay(){

        if(camera3 == ""){
            oresk1 = "18:00-21:00,00:00-03:00";
            oresk2 = "21:00-00:00,03:00-06:00";
            if(day == "ΣΑΒΒΑΤΟ" || day == "ΚΥΡΙΑΚΗ"){
                oresk1 = "10:00-12:00," + oresk1;
                oresk2 = "12:00-14:00," + oresk2;
            }
            else{
                oresk1 = "15:00-16:30," + oresk1;
                oresk2 = "16:30-18:00 " + oresk2;
            }
        }
        else{
            oresk1 = "06:00-09:00,15:00-18:00,00:00-02:00";
            oresk2 = "09:00-12:00,18:00-21:00,02:00-04:00";
            oresk3 = "12:00-15:00,21:00-00:00,04:00-06:00";
        }

        oresp1 = "10:00-12:00,18:00-21:00,00:00-03:00";

        if(per2a != ""){
            oresp2 = "12:00-14:00,21:00-00:00,03:00-06:00";
        }

        if(esteiatoras == ""){
            esteiatoras = mageiras;
        }


    }

    private void DateToString(){
        date = date.Split(" ")[0];
        DateTime tdate = DateTime.ParseExact(date, "d/M/yyyy", CultureInfo.InvariantCulture);

        // Λίστα με τα ονόματα των μηνών στα ελληνικά
        string[] greekMonths = { 
    "ΙΑΝΟΥΑΡΙΟΥ", 
    "ΦΕΒΡΟΥΑΡΙΟΥ", 
    "ΜΑΡΤΙΟΥ", 
    "ΑΠΡΙΛΙΟΥ", 
    "ΜΑΪΟΥ", 
    "ΙΟΥΝΙΟΥ", 
    "ΙΟΥΛΙΟΥ", 
    "ΑΥΓΟΥΣΤΟΥ", 
    "ΣΕΠΤΕΜΒΡΙΟΥ", 
    "ΟΚΤΩΒΡΙΟΥ", 
    "ΝΟΕΜΒΡΙΟΥ", 
    "ΔΕΚΕΜΒΡΙΟΥ" 
};
        // Εξαγωγή της ημέρας, μήνα και έτους
        int THday = tdate.Day;
        int month = tdate.Month;
        int year = tdate.Year;

        // Δημιουργία της μορφοποιημένης ημερομηνίας
        fullDaySt = $"{THday} {greekMonths[month - 1]} {year}";
        //MessageBox.Show(fullDaySt);

        dayTitle = "ΚΑΤΑΣΤΑΣΗ ΥΠΗΡΕΣΙΩΝ ΣΤΡΔΟΥ 'ΜΑΧΑΙΡΑ' ΤΗN " + day + " " + fullDaySt;
    }

    public void AddEksodouxoBEB(string eksodouxos){
        eksodosBEB.Add(eksodouxos);
    }

    public void AddAdeiouxoBEB(string Adeiouxos){
        adeiaBEB.Add(Adeiouxos);
    }

    public void AddEksodouxoTYL(string eksodouxos){
        eksodosTYL.Add(eksodouxos);
    }

    public void AddAdeiouxoTYL(string Adeiouxos){
        adeiaTYL.Add(Adeiouxos);
    }

    public void AddCamera1(string Cam1){
        camera1 = Cam1;
    }

    public void AddCamera2(string Cam2){
        camera2 = Cam2;
    }
    
    public void AddCamera3(string Cam3){
        camera3 = Cam3;
    }

    public void AddPeripolo1(string per){
        
        if(papaduan1){
            Random random = new Random();
            int i = random.Next(0, 2);
            
            //MessageBox.Show(i.ToString());

            if(i == 0){
              per1a = per + " (ΠΕΡΙΠΟΛΑΡΧΗΣ)";  
            }
            else{
                per1b = per;
            }
            
            papaduan1 = !papaduan1;
            return;
        }
        
        

        if(per1a == ""){
            per1a = per+ " (ΠΕΡΙΠΟΛΑΡΧΗΣ)";
        }
        else{
            per1b = per;
        }



    }

    public void AddPeripolo2(string per){
        
        if(papaduan2){
            Random random = new Random();
            int i = random.Next(0, 2);
            
            

            if(i == 0){
              per2a = per + " (ΠΕΡΙΠΟΛΑΡΧΗΣ)";  
            }
            else{
                per2b = per;
            }
            
            papaduan2 = !papaduan2;
            return;
        }
        
        if(per2a == ""){
            per2a = per + " (ΠΕΡΙΠΟΛΑΡΧΗΣ)";
        }
        else{
            per2b = per;
        }
    }

    public void AddMageira(string mag){
        mageiras = mag;
    }

    public void AddOdigo(string od){
        odigosEpif = od;
    }

    public void Addesteiatoras(string est){
        esteiatoras = est;
    }
}