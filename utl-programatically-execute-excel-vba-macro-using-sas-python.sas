 Silently(lights out) open and execute excel vba macro using sas python and save updated workbook.                                    
                                                                                                                                      
 github                                                                                                                               
 https://tinyurl.com/u2k36h8                                                                                                          
 https://github.com/rogerjdeangelis/utl-programatically-execute-excel-vba-macro-using-sas-python                                      
                                                                                                                                      
 SAS Forum                                                                                                                            
 https://tinyurl.com/yx6hy8ot                                                                                                         
 https://communities.sas.com/t5/SAS-Programming/Write-and-execute-VBA-macro-in-SAS-EG/m-p/606061                                      
                                                                                                                                      
 see                                                                                                                                  
 https://tinyurl.com/y7lmf9y8                                                                                                         
 http://jacobjwalker.effectiveeducation.org/blog/2015/01/24/python-script-to-automate-refreshing-an-excel-spreadsheet/                
                                                                                                                                      
 For input xlsm workbook. YOU NEED TO DOWNLOAD IT IT HAS THE VBA MACRO                                                                
 https://tinyurl.com/yaloqcuo                                                                                                         
 https://github.com/rogerjdeangelis/utl_programatically_execute_excel_macro_using_wps_proc_python/blob/master/class_final.xlsm        
                                                                                                                                      
*_                   _                                                                                                                
(_)_ __  _ __  _   _| |_                                                                                                              
| | '_ \| '_ \| | | | __|                                                                                                             
| | | | | |_) | |_| | |_                                                                                                              
|_|_| |_| .__/ \__,_|\__|                                                                                                             
        |_|                                                                                                                           
;                                                                                                                                     
                                                                                                                                      
  * You need this. It works in 32 and 64 bit Win 7 ( I have python 3 and Win & 64bit sas 9,4m6)                                       
                                                                                                                                      
  pip install pypiwin32                                                                                                               
                                                                                                                                      
  The macro enable workbook contains                                                                                                  
                                                                                                                                      
   Sub sum_weight()                                                                                                                   
    Range("F21").Select                                                                                                               
    ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"                                                                                   
    Range("F22").Select                                                                                                               
   End Sub                                                                                                                            
                                                                                                                                      
   You need to download the workbook. It has the macro above                                                                          
   https://tinyurl.com/yaloqcuo                                                                                                       
   which I saved locally as d:/xls/class_final.xlsm                                                                                   
                                                                                                                                      
   d:/xls/class_final.xlsm                                                                                                            
                                                                                                                                      
      +----------------------------------------------------------------+                                                              
      |     A      |    B       |     C      |    D       |    E       |                                                              
      +----------------------------------------------------------------+                                                              
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |                                                              
      +------------+------------+------------+------------+------------+                                                              
   2  | ALFRED     |    M       |    14      |    69      |  112.5     |                                                              
      +------------+------------+------------+------------+------------+                                                              
       ...                                                                                                                            
      +------------+------------+------------+------------+------------+                                                              
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |                                                              
      +------------+------------+------------+------------+------------+                                                              
                                                                                                                                      
      [CLASS]                                                                                                                         
                                                                                                                                      
*            _               _                                                                                                        
  ___  _   _| |_ _ __  _   _| |_                                                                                                      
 / _ \| | | | __| '_ \| | | | __|                                                                                                     
| (_) | |_| | |_| |_) | |_| | |_                                                                                                      
 \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                     
                |_|                                                                                                                   
;                                                                                                                                     
                                                                                                                                      
 When you open excel you should see this                                                                                              
                                                                                                                                      
   d:/xls/class_final.xlsm                                                                                                            
                                                                                                                                      
      +----------------------------------------------------------------+                                                              
      |     A      |    B       |     C      |    D       |    E       |                                                              
      +----------------------------------------------------------------+                                                              
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |                                                              
      +------------+------------+------------+------------+------------+                                                              
   2  | ALFRED     |    M       |    14      |    69      |  112.5     |                                                              
      +------------+------------+------------+------------+------------+                                                              
       ...                                                                                                                            
      +------------+------------+------------+------------+------------+                                                              
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |                                                              
      +------------+------------+------------+------------+------------+                                                              
   21 |            |            |            |            |  1900.9    | calculated by vba sum_weight macro                           
      +------------+------------+------------+------------+------------+                                                              
                                                                                                                                      
      [CLASS]                                                                                                                         
                                                                                                                                      
*                                                                                                                                     
 _ __  _ __ ___   ___ ___  ___ ___                                                                                                    
| '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                                   
| |_) | | | (_) | (_|  __/\__ \__ \                                                                                                   
| .__/|_|  \___/ \___\___||___/___/                                                                                                   
|_|                                                                                                                                   
;                                                                                                                                     
                                                                                                                                      
* copy the excel work book to d:/xls;                                                                                                 
* make sure macros are enabled;                                                                                                       
* you may want to make a copy because the code changes the workbook;                                                                  
                                                                                                                                      
* you may need to delete this folder;                                                                                                 
* C:\Users\<username>\AppData\Local\Temp\gen_py\;                                                                                     
                                                                                                                                      
%utl_submit_py64_37("                                                                                                                 
import win32com.client;                                                                                                               
import os;                                                                                                                            
win32c = win32com.client.DispatchEx('Excel.Application');                                                                             
xl = win32com.client.Dispatch('Excel.Application');                                                                                   
wb = xl.Workbooks.open('d:/xls/class_final.xlsm');                                                                                    
xl.Visible = True;                                                                                                                    
xl.Run('sum_weight');                                                                                                                 
wb.Save();                                                                                                                            
xl.Quit();                                                                                                                            
");                                                                                                                                   
