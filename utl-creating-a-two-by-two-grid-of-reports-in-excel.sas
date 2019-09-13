Creating a two by two grid of reports in excel                                                                                   
                                                                                                                                 
As far as I know SAS cannot produce side by side reports and graphs in a binary execl worrkbook.                                 
                                                                                                                                 
SAS Forum                                                                                                                        
https://tinyurl.com/yyduabt7                                                                                                     
https://communities.sas.com/t5/Graphics-Programming/how-to-place-2-tables-horizontally-for-easy-comparison/m-p/587633            
                                                                                                                                 
other sde by side outputs mostly excel;                                                                                          
                                                                                                                                 
https://github.com/rogerjdeangelis/utl_table_graph_ppt/blob/master/utl_table_graph_ppt.sas                                       
https://github.com/rogerjdeangelis/utl_side_by_side_excel_reports                                                                
https://github.com/rogerjdeangelis/utl-excel-report-with-two-side-by-side-graphs-below_python                                    
https://github.com/rogerjdeangelis/utl_excel_add_to_sheet                                                                        
https://github.com/rogerjdeangelis/utl-excel-report-with-two-side-by-side-graphs-below_python                                    
                                                                                                                                 
*_                   _                                                                                                           
(_)_ __  _ __  _   _| |_                                                                                                         
| | '_ \| '_ \| | | | __|                                                                                                        
| | | | | |_) | |_| | |_                                                                                                         
|_|_| |_| .__/ \__,_|\__|                                                                                                        
        |_|                                                                                                                      
;                                                                                                                                
                                                                                                                                 
options validvarname=upcase;                                                                                                     
libname sd1 "d:/sd1";                                                                                                            
data sd1.have;                                                                                                                   
  length species $5;                                                                                                             
  set sashelp.fish (keep=species weight height width                                                                             
      where=(species in ('Parkki' ,'Pike' ,'Smelt' ,'Whitefish')));                                                              
run;quit;                                                                                                                        
                                                                                                                                 
/*                                                                                                                               
 SD1.HAVE total obs=48                                                                                                           
                                                                                                                                 
  SPECIES    WEIGHT     HEIGHT     WIDTH                                                                                         
                                                                                                                                 
   White      270.0     8.3804    4.2476                                                                                         
   White      270.0     8.1454    4.2485                                                                                         
   White      306.0     8.7780    4.6816                                                                                         
   White      540.0    10.7440    6.5620                                                                                         
   White      800.0    11.7612    6.5736                                                                                         
   White     1000.0    12.3540    6.5250                                                                                         
   Parkk       55.0     6.8475    2.3265                                                                                         
   Parkk       60.0     6.5772    2.3142                                                                                         
   Parkk       90.0     7.4052    2.6730                                                                                         
  ....                                                                                                                           
*/                                                                                                                               
                                                                                                                                 
*            _               _                                                                                                   
  ___  _   _| |_ _ __  _   _| |_                                                                                                 
 / _ \| | | | __| '_ \| | | | __|                                                                                                
| (_) | |_| | |_| |_) | |_| | |_                                                                                                 
 \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                
                |_|                                                                                                              
;                                                                                                                                
                                                                                                                                 
 d:/xls/utl-excel-grid-of-four-reports-in-one-sheet.xlsx                                                                         
                                                                                                                                 
                                                                                                                                 
  Start At D11 (species=White)                  Start At J11 (species=Parkk)                                                     
                                                                                                                                 
      +--------------------------------------+   +--------------------------------------+                                        
      |     D   |    E    |     F   |    G   |   |     J   |    K    |     L   |    M   |                                        
      +--------------------------------------+   +--------------------------------------+                                        
  11  | SPECIES |  WEIGHT |  HEIGHT |  WIDTH |   | SPECIES |  WEIGHT |  HEIGHT |  WIDTH |                                        
      +---------+---------+---------+--------+   +---------+---------+---------+--------+                                        
  12  | White   |   370   |   7.38  |    69  |   | Parkk   |    70   |   6.38  |    79  |                                        
      +---------+---------+---------+--------+   +---------+---------+---------+--------+                                        
  13  | White   |   250   |   8.14  |    58  |   | Parkk   |    84   |   5.14  |    98  |                                        
      +---------+---------+---------+--------+   +---------+---------+---------+--------+                                        
  14  | White   |   366   |   9.77  |    52  |   | Parkk   |    36   |   7.77  |    12  |                                        
      +---------+---------+---------+--------+   +---------+---------+---------+--------+                                        
  ...                                                                                                                            
                                                                                                                                 
  Start At D25 (species=Pike)                  Start At J25 (species=Smelt)                                                      
                                                                                                                                 
      +--------------------------------------+   +--------------------------------------+                                        
  24  | SPECIES |  WEIGHT |  HEIGHT |  WIDTH |   | SPECIES |  WEIGHT |  HEIGHT |  WIDTH |                                        
      +---------+---------+---------+--------+   +---------+---------+---------+--------+                                        
  25  | Pike    |   571   |   1.34  |    39  |   | Smelt   |   106   |   1.38  |    39  |                                        
      +---------+---------+---------+--------+   +---------+---------+---------+--------+                                        
  26  | Pike    |   471   |   2.64  |    58  |   | Smelt   |   304   |   1.14  |    56  |                                        
      +---------+---------+---------+--------+   +---------+---------+---------+--------+                                        
  27  | Pike    |   716   |   3.97  |    32  |   | Smelt   |   630   |   2.77  |    72  |                                        
      +---------+---------+---------+--------+   +---------+---------+---------+--------+                                        
                                                                                                                                 
                                                                                                                                 
 *                                                                                                                               
 _ __  _ __ ___   ___ ___  ___ ___                                                                                               
| '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                              
| |_) | | | (_) | (_|  __/\__ \__ \                                                                                              
| .__/|_|  \___/ \___\___||___/___/                                                                                              
|_|                                                                                                                              
;                                                                                                                                
                                                                                                                                 
%utl_submit_r64('                                                                                                                
library(haven);                                                                                                                  
library(XLConnect);                                                                                                              
have<-read_sas("d:/sd1/have.sas7bdat");                                                                                          
White<-have[have$SPECIES=="White",];                                                                                             
parkk <-have[have$SPECIES=="Parkk",];                                                                                            
pike<-have[have$SPECIES=="Pike",];                                                                                               
smelt<-have[have$SPECIES=="Smelt",];                                                                                             
wb <- loadWorkbook("d:/xls/utl-excel-grid-of-four-reports-in-one-sheet.xlsx", create = TRUE);                                    
createSheet(wb, name = "species");                                                                                               
writeWorksheet(wb, parkk , sheet = "species", startRow = 11,startCol = 10, header = TRUE);                                       
writeWorksheet(wb, White, sheet = "species", startRow = 11,startCol = 4, header = TRUE);                                         
writeWorksheet(wb, smelt, sheet = "species", startRow = 25,startCol = 10, header = TRUE);                                        
writeWorksheet(wb, pike, sheet = "species", startRow = 25,startCol = 4, header = TRUE);                                          
saveWorkbook(wb);                                                                                                                
');                                                                                                                              
                                                                                                                                 
                                                                                                                                 
                                                                                                                                 
