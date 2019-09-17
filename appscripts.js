/**
*For printing headings of subjects
* 
* @customfunction
*/
function tpf() {
   var ss=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
   var c1;

  var col=ss.getLastColumn();
  var row=ss.getLastRow();
  //var col=17;

 for(var i=2;i<=col;i+=2)
 {
 
 var b=ss.getRange(5,i).getDisplayValue();
 ss.getRange(5,i+19).setValue(b);//Subject Heading Printing
 
 var a=ss.getRange(4,i).getDisplayValue();
 if(a==='')
 { a=ss.getRange(4,i+19-2).getDisplayValue();
   ss.getRange(4,i+19).setValue(a); }
 
 else
 {if(a==='Email-id')
    { c1=i-1;break;}
    else if(a==='Name')
    { c1=i-2;break;}
     else
       ss.getRange(4,i+19).setValue(a);}//Theory or Practical Printing

 }
  tpf2(row,c1);
 storecopy(); 
  
}
/**
*For printing roll numbers
* 
* @customfunction
*/

//fun for roll nos. printing
function tpf2(row,col) {
   var ss=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  var a=0,col,row;
 // var col=SpreadsheetApp.getActiveSheet().getLastColumn();
   //Logger.log(SpreadsheetApp.getActiveSheet().getLastRow());
   
//var row=SpreadsheetApp.getActiveSheet().getLastRow();
//var row=27;
//var col=17;
//Logger.clear();
for(var i=7;i<=row;i+=1)
 {
var a=0;
 for(var j=3;j<=col;j+=2)
 {
   var x=ss.getRange(i,j).getValue();
  // ss.getRange(i,20).setValue(x);//break;
   if(x<75)
   { a=1;
   var roll=ss.getRange(i,1).getValue();
     ss.getRange(i,20).setValue(roll);
     tpf3(i,row,col);
     break;

   }
 }
   //Logger.log('hello');
}
}

/**
*For printing percentage in each of subjects 
* @param i1 is the row number given by tpf2() function. It's a called function
* @customfunction
*/

function tpf3(i1,row,col) {
   var ss=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var a=0;

//var row=27;
//var col=17;
//var j=0;


for(var i=3;i<=col;i+=2)
 {
 
   var x=ss.getRange(i1,i).getValue();
     ss.getRange(i1,21+i-3).setValue(x);//j++;
 
 }

}

/**
*For wrapping in an array and passing it to another sheet. that is from sheet1 to sheet2.
* 
* @customfunction
*/

function storecopy() {

//Logger.clear();

var targetsheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
   var ss=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
//for date copy
var rng1=targetsheet.getRange(5,1,23-4);
var rng2=targetsheet.getRange(5,2,23-4);
var rng3=targetsheet.getRange(5,3,23-4);
var rng4=targetsheet.getRange(5,4,23-4);

targetsheet.getRange('A5').copyTo(rng1); 
targetsheet.getRange('B5').copyTo(rng2); 
targetsheet.getRange('C5').copyTo(rng3); 
targetsheet.getRange('D5').copyTo(rng4); 

var temp=ss.getRange('T7:AI26').getValues();
/*for(var i=0;i<=temp.length;i+=1)
 {
 
    targetsheet.getRange(1,1,20,16).setValue(temp[i]);
 
 }
*/   
 targetsheet.getRange('E1').setFormula('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1j-tUM2y9yb1YkasF1AeaBE5Es4AzyUuI54Zux5c7R3g/edit#gid=0","R4:AI26")');   //Formula is set for D1
//var lr=app.getLastRow();
//var rng=app.getRange(row, column,lr);
//app..getRange('D1').copyTo(rng); 

//Logger.log(temp[0]);
 sendemail();

}

/**
*Sending email.
* 
* @customfunction
*/


function sendemail() {

  var ss=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  var lr=ss.getLastRow();  var col=ss.getLastColumn();
  var templateText=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet3').getRange(1,1).getValue();
   //Logger.log(templateText);
  for(var i=5;i<=lr;i++)
  {
       var templateText=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet3').getRange(1,1).getValue();
       var currentDate=ss.getRange(i,1).getValue();
       var startTime=ss.getRange(i,2).getValue();
       var endTime=ss.getRange(i,3).getValue();
       var currentRoom=ss.getRange(i,4).getValue();
       var currentSubj=ss.getRange(2,4).getValue();
     
       var currentEmail=ss.getRange(i,5).getValue();
       var currentName=ss.getRange(i,6).getValue();
       var currentPrn=ss.getRange(i,7).getValue();
       var k=1,k1=1;
       var finalEmail;
       var type;
    for(var j=8;j<=22;j+=2)
   {
      var subjName=ss.getRange(2,j).getValue();
      var subjPer=ss.getRange(i,j).getValue();
      type=ss.getRange(1,j).getValue();
    if(type==='Theory') {   
      if(k==1)
        {templateText=templateText.replace('{subjName1}',subjName).replace('{subjPer1}',subjPer);  }
       
      else if(k==2)
         { templateText=templateText.replace('{subjName2}',subjName).replace('{subjPer2}',subjPer); } 
     
      else if(k==3)
           templateText=templateText.replace('{subjName3}',subjName).replace('{subjPer3}',subjPer);
     
       else if(k==4)
             templateText=templateText.replace('{subjName4}',subjName).replace('{subjPer4}',subjPer);
     
      templateText=templateText.replace('{type}',type);
     k++;
     }
     else
      {
      if(k1==1)
          {templateText=templateText.replace('{subjNames1}',subjName).replace('{subjPers1}',subjPer);  }
       
     else if(k1==2)
        { templateText=templateText.replace('{subjNames2}',subjName).replace('{subjPers2}',subjPer); } 
     
     else if(k1==3)
           templateText=templateText.replace('{subjNames3}',subjName).replace('{subjPers3}',subjPer);
     
     else if(k1==4)
          templateText=templateText.replace('{subjNames4}',subjName).replace('{subjPers4}',subjPer);
     
     templateText=templateText.replace('{types}',type);
     k1++;
      }  
    }  
    finalEmail=templateText.replace('{date}',currentDate).replace('{start}',startTime).replace('{end}',endTime).replace('{room}',currentRoom).replace('{name}',currentName).replace('{prn}',currentPrn);
    var subject='Information about Parents Teacher Meeting (PTM)';
    if(currentEmail!=='')
   MailApp.sendEmail(currentEmail, subject, finalEmail);
     //Logger.log(finalEmail);
} 
    
}

