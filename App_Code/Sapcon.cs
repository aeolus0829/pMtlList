using System;

/// <summary>ed
/// Summary description for Sapcon
/// </summary>
public class Sapcon : SAPLogonCtrl.SAPLogonControlClass
{
    public string Name { get; set; }
    public string AppServerHost { get; set; }
    public string SystemID { get; set; }

    public  Sapcon(string clientNum)
	{          
        switch (clientNum)
        {
            case "800":
                Name = "SAPPRD01";
                AppServerHost = ApplicationServer = "192.168.0.16";
                SystemID = "PRD";
                Client = "800";
                break;
            case "620":
                Name = "SAPDEV01";
                AppServerHost = ApplicationServer = "192.168.0.15";
                SystemID = "DEV";
                Client = "620";
                break;
        }
        Language = "ZF";
        User = "DDIC";
        Password = "Ubn3dx";
        SystemNumber = int.Parse("00");       
	}
}