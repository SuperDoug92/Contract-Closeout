function open_files(A, B, C) 
{ 
 var excel = new ActiveXObject("Excel.Application");
 excel.Visible=true;
 excel.DisplayAlerts = false;
 var wbA = excel.Workbooks.Open(document.getElementById(A).value);
 //var wbB = excel.Workbooks.Open(document.getElementById(B).value);
 //var wbC = excel.Workbooks.Open(document.getElementById(C).value);
 
 excel.EnableEvents = false;
 excel.ScreenUpdating = false;
 excel.Calculation = -4135 //xlCalculationManual enumeration;
 
 var wb_collection = [wbA]; //, wbB, wbC];

 //excel.Application.Run("'" + wbA.name + "'" + '!update_links');

 var CLIN_list = new VBArray(wbA.Sheets("Control Form").Range("B62:B141").value).toArray();
 for (var i=0; i < CLIN_list.length; i++) 
 {
  if (!CLIN_list[i])
  {
   CLIN_list.splice(i, 1);
   i--
  }
 }

 var clin_table = document.createElement('table');
 var CLIN_input_string = "<table>\
<tr>\
<th>CLIN</th>\
<th>LM Aero</th>\
<th>LMLS (Greenville)</th>\
<th>LM MST</th>\
<th>NGC</th>\
<th>BAE</th>\
</tr>";

 for (var i=0; i < CLIN_list.length; i++) 
 {
  CLIN_input_string.concat("<tr>\
<td>" + CLIN_list[i] + "</td>\
<td><input type=&quot;number&quot; name=" + CLIN_list[i] + " Aero" + "></td>\
<td><input type='number' name=" + CLIN_list[i] + " LMLS" + "></td>\
<td><input type='number' name=" + CLIN_list[i] + " MST" + "></td>\
<td><input type='number' name=" + CLIN_list[i] + " NGC" + "></td>\
<td><input type='number' name=" + CLIN_list[i] + " BAE" + "></td>\
</tr>"
  )
 }
 CLIN_input_string.concat("</table>");

 var fso = new ActiveXObject("Scripting.FileSystemObject");
 var path = fso.GetAbsolutePathName("innerhtml.html");
 path = path + "\\innerhtml.html"

 var inner_html_file = fso.CreateTextFile(path, true);
 
 inner_html_file.write(CLIN_input_string);
 
 update();
    
 

	
 


/*

 for (i = 0; i < CLIN_list.length+1; i++)
 { 
  if (CLIN_list[i] > 0)
  {
   var CLIN_list_count = i
  }
 }

 var decrement_range_start = wbA.Sheets("Fee & Decrement Table").Range("AJ14")
 
 for (i = 0; i < 80; i++){
   Sheets("Fee & Decrement Table").Cells(decrement_range_start.column+i


 Model Setup for VBA
 wbA.Sheets("CONTROL FORM").Activate
 wbA.Sheets("CONTROL FORM").OLEObjects("TextBox21").Object.Text = wbB.fullname
 wbA.Sheets("CONTROL FORM").OLEObjects("TextBox22").Object.Text = wbC.fullname
 
 excel.Application.Run("'" + wbA.name + "'" + '!Run_JPO');

 */






 for (var wb in wb_collection)
 {
  excel.Workbooks(wb_collection[wb].name).Close (false);
 }	
 excel.Application.Quit();
}

