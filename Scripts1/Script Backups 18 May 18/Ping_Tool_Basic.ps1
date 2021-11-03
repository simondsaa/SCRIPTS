$ServerName = Get-Content "C:\Users\1252862141.adm\Desktop\Scripts\Ping_Tool_Basic.txt"  
  
foreach ($Server in $ServerName) {  
  
        if (test-Connection -ComputerName $Server -Count 2 -Quiet ) {   
          
            "$Server is Pinging "  
          
                    } else  
                      
                    {"$Server NOT pinging"  
              
                    }      
          
} 