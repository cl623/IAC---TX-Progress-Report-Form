function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
} 

function tag(text){
  var url = ScriptApp.getService().getUrl()
  var user = Session.getActiveUser().getEmail()
  let tag = /([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})/
  var email = text.match(tag)
  GmailApp.sendEmail(email[0], user, text)
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  var url = ScriptApp.getService().getUrl();
  return url;
}


function doGet(requestInfo) {
  console.log(requestInfo);
  if (requestInfo.parameter && requestInfo.parameter['page'] == 'Form') {
    if(requestInfo.parameter['clientID']==''){
      return HtmlService.createTemplateFromFile("Search").evaluate(); //Error page
    }
    else{
      return HtmlService.createTemplateFromFile("Form").evaluate();
    }
  }
  else if (requestInfo.parameter && requestInfo.parameter['page'] == 'Picker'){
    return HtmlService.createHtmlOutputFromFile('Picker').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
  return HtmlService.createTemplateFromFile("Search").evaluate();
}

////=========Database Functions===========
//
///*
//* Contains info to connect to database
//*/      
function databaseConnect(){
  var connection = PropertiesService.getScriptProperties().getProperties();
  
  return connection;
}
///*
//* Creates connection to database
//*/
function connect(connection){
  var connection = 'test-project-278417:us-east4:live-data-v1'
  var url = 'jdbc:google:mysql://test-project-278417:us-east4:live-data-v1/DatabaseProject'
  var user = 'root'
  var password = 'ci8sgfk8LfG4jo9P'
  var conn = Jdbc.getCloudSqlConnection(url, user, password);
  return conn;
  
}

/*
 *  Queries Database for client information
 */
function getInfo(clientId){
  var dbConnect = databaseConnect();
  var conn = connect(dbConnect);
  var stmt = conn.createStatement(),stmt2 = conn.createStatement(),stmt3 = conn.createStatement(),stmt4 = conn.createStatement();    
  var queryInfo = "select distinct concat(c.FirstName,' ',c.LastName) as ClientName,c.Gender,c.ClientId,c.DateOfBirth,DATE_FORMAT(NOW(), '%Y') - DATE_FORMAT(c.DateOfBirth, '%Y') - (DATE_FORMAT(NOW(), '00-%m-%d') < DATE_FORMAT(c.DateOfBirth, '00-%m-%d')) AS Age,lg1.LanguageDialect as PriLang,lg2.LanguageDialect as SecLang from client c left join person_language pl1 on pl1.ClientId=c.ClientId and pl1.Fluency='Primary' left join person_language pl2 on pl2.ClientId=c.ClientId and pl2.Fluency='Secondary' left join language lg1 on lg1.LanguageDialectId=pl1.LanguageDialectId left join language lg2 on lg2.LanguageDialectId=pl2.LanguageDialectId where c.ClientId="+clientId+";";
  var queryContacts = "select concat(FirstName,' ',LastName) as GuardianName,Address,City,State,ZipCode,HomeNumber,MobileNumber,EmergencyContact from client_contacts where ClientId="+clientId+";";
  var queryMedical = "select PcpName,PcpNumber from client_medical where ClientId="+clientId+";";
  var queryInsurance = "select i.NamePolicyHolder,i.MemberId,p.InsurancePlan,a.Payer from client_insurance i left join insurance_plan p on i.InsurancePlanId=p.InsurancePlanId left join payer a on p.PayerId=a.PayerId where i.ClientId="+clientId+";";
  var clientInfoData = stmt.executeQuery(queryInfo);
  
  
  
  var obj = {};
  while(clientInfoData.next()){
    obj['ClientName'] = clientInfoData.getString('ClientName');
    obj['ClientID'] = clientInfoData.getInt('ClientId');
    obj['ClientGen'] = clientInfoData.getString('Gender');
    obj['DOB'] = clientInfoData.getString('DateOfBirth');
    obj['Age'] = clientInfoData.getInt('Age');
  
    obj['PrimaryLangauge'] = clientInfoData.getString('PriLang');
    obj['SecondaryLanguage'] = clientInfoData.getString('SecLang');
    
    var clientContactsData = stmt2.executeQuery(queryContacts);
    var count = 1;
    while(clientContactsData.next()){
      obj['Guardian'+count+'Name'] = clientContactsData.getString('GuardianName');
      obj['Address'+count+''] = clientContactsData.getString('Address');
      obj['City'+count+''] = clientContactsData.getString('City');
      obj['State'+count+''] = clientContactsData.getString('State');
      obj['ZipCode'+count+''] = clientContactsData.getString('ZipCode');
      obj['Guardian'+count+'HP'] = clientContactsData.getString('HomeNumber');
      obj['Guardian'+count+'CP'] = clientContactsData.getString('MobileNumber');
      obj['Emergency'+count+'Name'] = clientContactsData.getString('EmergencyContact');
      count++
    }
//    var clientMedicalData = stmt3.executeQuery(queryMedical);
//    var count = 1;
//    while(clientMedicalData.next()){
//      obj['PcpName'] = clientMedicalData.getString('PcpName');
//      obj['PcpNumber'] = clientMedicalData.getString('PcpNumber');
//    }
    var clientInsuranceData = stmt4.executeQuery(queryInsurance);
    var count = 1;
    while(clientInsuranceData.next()){
      obj['InsuranceHolder'+count+'Name'] = clientInsuranceData.getString('NamePolicyHolder');
      obj['Insurance'+count+'ID'] = clientInsuranceData.getString('MemberId');
      obj['InsurancePlan'+count+'Name'] = clientInsuranceData.getString('InsurancePlan');
      obj['InsurancePayer'+count+'Name'] = clientInsuranceData.getString('Payer');
      count++
    }
  }
  stmt.close();
  stmt2.close();
  stmt3.close();
  stmt4.close();
  conn.close();
  var json = JSON.stringify(obj);
  return json;
}
function testDB(){
 Logger.log(getInfo('3')); 
}
//
////==================================

/*
 *  Gets client names from spreadsheet
 */

function getClientName() {
  var spreadsheetID = "1G0kpCNdmow74d0u7pGCbN7ubFE4lmMFbVMCnlWPbt0I";
  var clientNames = SpreadsheetApp.openById(spreadsheetID).getDataRange().getDisplayValues();
  return clientNames; 
}

function getMaladaptives(){
  var spreadsheet = '1aaQTY0EMtrNiuyhzchMABUvlHzc9fAkn4g5B087S2ZA';
  var maladaptives = SpreadsheetApp.openById(spreadsheet).getDataRange().getDisplayValues();
  return maladaptives
}

function testMal(){
  malList = getMaladaptives()
  for(x of malList){
    if(x[3] != ''){
      console.log('Madalaptive Behavior: ' + x[3] + '\nDefinition: ' + x[4])
    }
  }
}

/*
 *  Gets script user email
 */

function getUser(){
  return Session.getActiveUser().getEmail()
}

/*
 *  Deprecated function for getting client info. Now getting info from DATABASE
 */

//function getClientInfo(clientName, clientID) {
//  var clientName = clientName || "NA";
//  var clientID =  clientID||"NA";
//  Logger.log(clientName);
//  var clientData = [];
//  var spreadsheetID = "1G0kpCNdmow74d0u7pGCbN7ubFE4lmMFbVMCnlWPbt0I";
//  
//  //Need to make multiple variables to hold several sheets/pages (Statuses/Assignments, Client Info, Client Insurance)
//  
//  //Client Status/Assignment
//  var data = getClientName();
//  
//  //Loop for sheets that have multiplicities Eg. Insurance sheet
//  // data[i][3] Second Index is location of FULL NAME in the spreadsheet. CHANGE ACCORDINGLY
//  for (var i = 0; i < data.length; i++){
//    if (data[i][3] == clientName || data[i][0] == clientID){
//      clientData.push(data[i]);
//    }
//  }
//  
//  //Contact Info
//  var contacts = SpreadsheetApp.openById(spreadsheetID).getSheetByName("Sheet2").getDataRange().getDisplayValues();
//  for(var i=0; i < contacts.length;i++){
//    if(contacts[i][1] == clientName || contacts[i][0] == clientID){
//      clientData.push(contacts[i]);
//    }
//  }
//  Logger.log(clientData[1]);
//  
//  //Client Insurance
//  // ins[i][1] Second Index is location of FULL NAME in the spreadsheet. CHANGE ACCORDINGLY
//  //Sheet3 -> Insurance Sheet
//  var ins = SpreadsheetApp.openById(spreadsheetID).getSheetByName("Sheet3").getDataRange().getDisplayValues();
//  for(var i=0 ; i < ins.length; i++){
//    if(ins[i][1] == clientName  || ins[i][0] == clientID){
//      clientData.push(ins[i]);
//    }
//  }
//  return clientData;
//}


//======================= Save data to sheet ===========================

/*
 *  Get next empty row of spreadsheet
 *      @param id: id of spreadsheet
 */

function getEmptyRow(id){
  var ss = SpreadsheetApp.openById(id)
  var col = ss.getRange('A:A')
  var val = col.getValues();
  var nextRow = 0;
  while(val[nextRow] && val[nextRow][0] != ''){
    nextRow++;
  }
  return (nextRow+1);
}

/*
 *  Save data to spreasheet
 *    @param data: array to be saved pertaining to form info
 */

function sendFormData(data){
  var sheetID = '1HQnH3hY-0YuAUOM_eP6JxAVMemhZBCAVQZJ_GZMcVT8'
  var spreadsheet = SpreadsheetApp.openById(sheetID).getSheets()[0];
  var nextEmpty = getEmptyRow(sheetID);
  try{
    var range = spreadsheet.getRange(nextEmpty,1,1, data[0].length).setValues(data);
  }
  catch(e){
    return e;
  }
  return "Saved Successfully"
}
//======================== For Review ==============================

/*
 *  Report submitted for review.
 */

function emailReviewer(data, email, cont){
  let body = cont || 'EMAIL BODY'
  try{
  GmailApp.sendEmail(email, 'Funder Report Draft for ' + data[0][0], body)
  return 'Reviewer Emailed'
  }
  catch(e){
   return e 
  }
}

  /*
   *  Report rejected and comments are passed to writer
   */

function emailComments(comments, email, id){
  try{
    var body = 'Report for client ' + id + ' has been reviewed and rejected. Here are your reviewers comments: \n\n'
    var subject = 'Funder Report Rejected: Client #' + id 
    for(comment in comments){
      body += comment + ': ' + comments[comment] + '\n'
    }
   GmailApp.sendEmail(email, subject, body) 
   return "Comments Emailed"
  }
  catch(e){
   return e 
  }
}

//======================== Create Doc =====================================

/*
 *  Create Funder Report Template and fill with form data
 *    @param data: array with form input elements, username, review stage, timestamp
 *    @param skills: skill data from progress report parse
 *    @param signatureID: drive file ID for reviewer signature
 */


function makeTemplate(data, skills, signatureID){
  var file= DriveApp.getFileById('1_Bw_WeQnjnIdYI0teO2-00DKBVhwR4caqwq69SwyIn4').makeCopy(data[0][2] + " " + data[0][0] + " TX PLAN TEMPLATE")
  var doc = DocumentApp.openById(file.getId())
  var body= doc.getBody();
  
  let signature = DriveApp.getFileById(signatureID).getBlob()
  body.appendImage(signature)
  
  var crisisPlan = ['Assaultive Behavior','Self-Injurious Behavior', 'Fire Setting', 'Impulsive Behavior', 'Current Family Abuse Violence', 'Elopement/Bolting', 'Sexually Offending Behavior', 'Substance Abuse', 'Psychotic Symptoms', 'Coping with Significant Loss', 'Suicidality', 'Homicidality']
  var crisisRisks = ['Present','Ideation','Plan','Means','Prior']
  var crisisDone = false;
  
  var o = JSON.parse(data[0][4])
  
  for(x in o){
    if(x == 'Communication Skills Domain'){
      var range = body.findText("{"+x+"}");
      var ele = range.getElement();
      if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
        var offset = body.getChildIndex(ele.getParent());
        if(o['Physical Activity Skills Domain'] != undefined){
          var PASD = 'Physical Activity Skills Domain: ' + o['Physical Activity Skills Domain']
          body.insertListItem(offset + 1, PASD).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
        }
        if(o['Daily Living Skills Domain'] != undefined){
          var DLSD = 'Daily Living Skills Domain: ' + o['Daily Living Skills Domain']
          body.insertListItem(offset + 1, DLSD).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
        }
        if(o['Social Skills & Relationship Skills Domain'] != undefined){
          var SSRSD = 'Social Skills & Relationship Skills Domain: ' + o['Social Skills & Relationship Skills Domain']
          body.insertListItem(offset + 1, SSRSD).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
        }
        if(o[x] != undefined){
          var CDS = 'Communication Skills Domain: ' + o[x]
          body.insertListItem(offset + 1, CDS).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
        }
      }
    }
    else if(x == "Medical Condition/Diagnosis"){
      if(o[x] == undefined)
        body.replaceText("{"+x+"}", "According to "+ o['Name1'] +"’s parents, "+ o['Name1'] +" is a healthy child.") 
        else{
          body.replaceText("{"+x+"}", "According to "+ o['Name1'] +"’s parents, "+ o['Name1'] +" additionally suffers from the following medical conditions and medical diagnoses: " + o[x])
        }
    }
    else if(x == "Medications"){
      if(o[x] == undefined){
        body.replaceText("{"+x+"}", "According to "+ o['Name1'] +"’s parents, "+ o['Name1'] +" does not presently take any prescription or over the counter medications on a daily or regular basis.") 
      }else{
        var medications=[]
        if(Array.isArray(o[x])){
          for(var i = 0; i < o['Medications'].length; i++){
            medications.push([o['Medications'][i], o['Medications Dosage'][i], o['Medications Frequency'][i]])
          }
        }
        else{
          medications.push([o['Medications'], o['Medications Dosage'], o['Medications Frequency']])
        }
        var range = body.findText("{"+x+"}");
        var ele = range.getElement();
        if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
          var offset = body.getChildIndex(ele.getParent());
          body.insertTable(offset + 1, medications);
        }
        body.replaceText("{"+x+"}", "According to "+ o['Name1'] +"’s parents, "+ o['Name1'] +" presently takes the following prescription medications and dosages on a daily basis: ")
      }
    }
    else if(x == "Vitamins"){
      if(o[x] == undefined){
        body.replaceText("{"+x+"}", "According to "+ o['Name1'] +"’s parents, "+ o['Name1'] +" does not presently take any vitamins, minerals, or supplements on a daily or regular ") 
      }else{
        var vitamins=[]
        if(Array.isArray(o[x])){
          for(var i = 0; i < o['Vitamins'].length; i++){
            vitamins.push([o['Vitamins'][i], o['Vitamins Dosage'][i], o['Vitamins Frequency'][i]])
          }
        }
        else{
          vitamins.push([o['Vitamins'], o['Vitamins Dosage'], o['Vitamins Frequency']])
        }
        var range = body.findText("{"+x+"}");
        var ele = range.getElement();
        if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
          var offset = body.getChildIndex(ele.getParent());
          body.insertTable(offset + 1, vitamins);
        }
        body.replaceText("{"+x+"}", "According to "+ o['Name1'] +"’s parents, "+ o['Name1'] +" presently takes the following vitamins, minerals, supplements, and dosages on a daily basis: ")
      }
    }
    else if(x == "Allergies"){
      if(o[x] == undefined){
        body.replaceText("{"+x+"}", "According to "+ o['Name1'] +"’s parents, "+ o['Name1'] +" currently does not have any known allergies.") 
      }else{
        body.replaceText("{"+x+"}", "According to "+ o['Name1'] +"’s parents, "+ o['Name1'] +" currently has the following known allergies:" + o[x])
      }
    }
    else if(x == "Early Intervention"){
      if(o['OtherServicesNotEI'] == "Yes"){
        var EI=[['Treatment Service', 'Session Duration', 'Frequency', 'Location', 'Start Date', 'End Date']]
        var range = body.findText("{"+x+"}");
        var ele = range.getElement();
        var offset = body.getChildIndex(ele.getParent());
        for(var i = 0; i < o['OSEI Service'].length; i++){
          EI.push([o['OSEI Service'][i], o['OSEI Duration'][i], o['OSEI Frequency'][i], o['OSEI Location'][i], o['OSEI Start Date'][i], o['OSEI End Date'][i]])
        }        
        if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
          body.insertTable(offset + 1, EI);
        }
        
        body.insertParagraph(offset + 1, o["Name1"]+" currently receives the following treatment services provided by other by other qualified medical/mental health professionals.")      
        
      }
      if(o[x] == "Yes"){
        var range = body.findText("{"+x+"}");
        var ele = range.getElement();
        var offset = body.getChildIndex(ele.getParent());
        var EI=[['EI Service/Method', 'Provider', 'Location of Service', 'Frequency & Length', 'Intensity of Service', 'Duration of Service', 'Start Date', 'End Date']]
        for(var i = 0; i < o['PastEIService'].length; i++){
          EI.push([o['PastEIService'][i], o['PastEIProvider'][i], o['PastEILocation'][i], o['PastEIFrequency'][i], o['PastEIIntesity'][i], o['PastEIDuration'][i], o['PastEIStart'][i], o['PastEIEnd'][i]])
        }
        var range = body.findText("{"+x+"}");
        var ele = range.getElement();
        if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
          var offset = body.getChildIndex(ele.getParent());
          body.insertTable(offset + 1, EI);
        }
        
        body.findText("{"+x+"}").getElement().getParent()
        
        body.insertParagraph(offset + 1, o["Name1"] +" previously received early intervention (EI) treatment services from" + o[' previous ei facility name'] + 
                             ", located in " + o[' previous ei facility location'] + ", MA. {Pronoun3} specifically received the following early intervention treatment services:")
        
        
        //{Name1} was discharged EI treatment services on XX/XX/XXXX when he/she turned 3-years of age. X has not and currently does not receive any type of treatment services outside of the early intervention (EI)/school environment.
      }
      
      if(o[x] == "No" && o['OtherServicesNotEI'] == "No"){
        var range = body.findText("{"+x+"}");
        var ele = range.getElement();
        var offset = body.getChildIndex(ele.getParent());
        body.insertParagraph(offset + 1, o["Name1"]+" has not and currently does not receive any type of treatment services outside of the early intervention (EI)/school environment.")
      }
      body.replaceText("{" + o[x] + "}", '')
    }
    else if(x == "CEarlyIntervention"){
      var range = body.findText("{"+x+"}");
      var ele = range.getElement();
      var offset = body.getChildIndex(ele.getParent());
      if(o[x] == "Yes"){
        var EI=[['EI Service/Method', 'Provider', 'Location of Service', 'Frequency & Length', 'Intensity of Service', 'Duration of Service', 'Start Date', 'End Date']]
        for(var i = 0; i < o['CEI Service'].length; i++){
          EI.push([o['CEI Service'][i], o['CEI Provider'][i], o['CEI Location'][i], o['CEI Frequency'][i], o['CEI Intensity'][i], o['CEI Duration'][i], o['CEI Start'][i], o['CEI End'][i]])
        }
        
        if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
          body.insertTable(offset + 1, EI);
        }
        
        body.insertParagraph(offset + 1, "Currently, " + o["Name1"]+" is under three years of age and therefore is not eligible to attend school within the state of Massachusetts at this time. Presently, " + 
                             o["Name1"] + " ttends/receives early intervention treatment services from " + 
                             o['Current EI Facility Name'] + " Early Intervention Agency, which is located in " + o['Current EI Facility Location'] + ", MA.")      
        
      }
      if(o['attendingSchool'] == "Yes"){
        if(o['IEP504'] == 'notIEP504'){
          body.insertParagraph(offset + 1, "Presently, " + o["Name1"]+ " is not on an individualized education plan/program (IEP) at school and therefore does not receive any special education services or supports at this time.")
        }
        else if(o['IEP504'] == 'isIEP'){
          //          body.insertImage(offset + 1, DriveApp.getFileById(o['IEPFile']).getBlob())
          body.insertParagraph(offset + 1, "Presently, " + o["Name1"]+ " is on an individualized education plan/program (IEP) at school and accordingly receives the following special education services and supports:")
        }
        else if(o['IEP504'] == 'is504'){
          body.insertParagraph(offset + 1, "Other:" + o['504 Category Other'])
          body.insertParagraph(offset + 1, "Presentation:" + o['504 Category Presentation'])
          body.insertParagraph(offset + 1, "Setting:" + o['504 Category Setting'])
          body.insertParagraph(offset + 1, "Scheduling/Timing:" + o['504 Category Scheduling/Timing'])
          body.insertParagraph(offset + 1, "Response:" + o['504 Category Response'])
          body.insertParagraph(offset + 1, "Presently, " + o["Name1"]+ " is on a Section 504 Accommodation Plan at school and accordingly receives the following accommodations:")
        }
        else if(o['IEP504'] == 'isEval'){
          body.insertParagraph(offset + 1, "Presently, " + o["Name1"]+ " is being evaluated in order to determine whether or not " + o['Pronoun3'] + 
                               " is currently eligible to receive special education services and supports through the X Public School District. A Special Education Eligibility Determination Meeting has been scheduled for " +
                               o['Special Education Eligibility Determination meeting date'] + " During the scheduled Special Education Eligibility Determination Meeting, the results of all of the special education evaluations conducted with " + o["Name1"]+ 
                               " by the X Public School District’s Special Education Department will be reviewed thoroughly with " + o["Name1"]+ "’s educational treatment team. This information will then be used by " + o["Name1"]+
                               "’s educational treatment team in order to determine whether or not " + o["Name1"]+ " will be found eligible to receive special education services and supports via the X Public School District.")
        }
        if(o[x] == 'No' && o['attendingSchool'] == 'No'){
          body.insertParagraph(offset + 1, "Presently, " + o["Name1"]+ "  does not attend school and spends a majority of " + o['Pronoun1'] + 
                               " day within a private/public day care setting; within " + o['Pronoun1'] + " home; at relative’s/family friend’s home. Consequently, " + o['Name'] + 
                               "does not receive any type of early intervention (EI) or special education services or supports at this time.")
        }
        body.insertParagraph(offset + 1, "Currently, " + o["Name1"]+ " attends a " + o['Program Type'] +", "+ o['Private or Public School'] + " day school classroom, which is located in " + o['Location of School'] +", MA")
      }
      body.replaceText("{" + o[x] + "}", '')
    }
    else if(x == 'Current Communication Skills Domain'){
      var range = body.findText("{"+x+"}");
      var ele = range.getElement();
      if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
        var offset = body.getChildIndex(ele.getParent());
        if(o['Current Physical Activity Skills Domain'] != undefined){
          var PASD = 'Physical Activity Skills Domain: ' + o['Current Physical Activity Skills Domain']
          body.insertListItem(offset + 1, PASD).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
        }
        if(o['Current Daily Living Skills Domain'] != undefined){
          var DLSD = 'Daily Living Skills Domain: ' + o['Current Daily Living Skills Domain']
          body.insertListItem(offset + 1, DLSD).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
        }
        if(o['Current Social Skills & Relationship Skills Domain'] != undefined){
          var SSRSD = 'Social Skills & Relationship Skills Domain: ' + o['Current Social Skills & Relationship Skills Domain']
          body.insertListItem(offset + 1, SSRSD).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
        }
        if(o[x] != undefined){
          var CDS = 'Communication Skills Domain: ' + o[x]
          body.insertListItem(offset + 1, CDS).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
        }
      }
      body.replaceText("{"+x+"}", '')
    }
    else if(x == "Maladaptive Behaviors"){
      var range = body.findText("{"+x+"}");
      var ele = range.getElement();
      var Freq = [['Maladaptive Behavior', 'Frequency Score']]
      var Intensity = [['Maladaptive Behavior', 'Intensity Score']]
      var Duration = [['Maladaptive Behavior', 'Duration Score']]
      var Discrimmination = [['Maladaptive Behavior', 'Discrimination Score']]
      var HypothesizedFunc = [['Maladaptive Behavior', 'Hypothesized Function(s)']];
      if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
        
        /*
        *Current problem: function does not work if property has less than one answer/is not array. Must select more than one Maladaptive Behavior.
        *
        */if(Array.isArray(o['Maladaptive Behaviors'])){
          for(var j = o['Maladaptive Behaviors'].length-1; j > -1 ; j--){
            let mal = o[x][j]
            Freq.push([mal, o[mal + 'FreqScore']])
            Intensity.push([mal, o[mal + 'IntensityScore']])
            Duration.push([mal, o[mal + 'DurationScore']])
            Discrimmination.push([mal, o[mal + 'DiscriminationScore']])
            HypothesizedFunc.push([mal, o[mal + 'HF']])
            
            var offset = body.getChildIndex(ele.getParent());
            body.insertListItem(offset + 1, o['Maladaptive Behaviors'][j]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
            body.insertListItem(body.getChildIndex(body.findText("{Maladaptive2}").getElement().getParent()) + 1, o['Maladaptive Behaviors'][j]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
            body.insertParagraph(body.getChildIndex(body.findText("{Identified Antecedents & Consequences}").getElement().getParent()) + 1, '{'+ mal + 'Antecedent}')
            
            if(o.hasOwnProperty(mal+'Replace1')){
              for(var k = 4; k > -1 ; k--){
                body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Replace1'][k])
              }
              body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified Functionally Equivalent Replacement Skills to be Taught:")
            }
            if(o.hasOwnProperty(mal+'Access')){
              if(Array.isArray( o[mal+'Access'])){
                for(var k = o[mal+'Access'].length-1; k > -1 ; k--){
                  body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Access'][k]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
                }
              }
              else{
                body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Access']).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
              }
              body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Access To")
            }
            if(o.hasOwnProperty(mal+'Consequences')){
              if(Array.isArray( o[mal+'Consequences'])){
                for(var k = o[mal+'Consequences'].length; k > -1 ; k--){
                  body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Consequences'][k]);
                }
              }
              else{
                body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Consequences']);
              }
              body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified consequences that appear to be maintaining "+o['Name1']+"’s emission of non-compliance are as follows:")
            }
            if(o.hasOwnProperty(mal+'Antecedent')){
              if(Array.isArray( o[mal+'Antecedent'])){
                for(var k = o[mal+'Antecedent'].length-1; k > -1 ; k--){
                  body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Antecedent'][k]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
                }
              }
              else{
                body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Antecedent']).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
              }
              body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Antecedents that have been identified as initiating "+o['Name1']+"’s emission of non-compliance are as follows:")
              body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified Antecedents & Consequences: " + mal)
            }
            body.replaceText("{"+mal+"Antecedent}", '')
          }
        }
        else{
          let mal = o[x]
          Freq.push([mal, o[mal + 'FreqScore']])
          Intensity.push([mal, o[mal + 'IntensityScore']])
          Duration.push([mal, o[mal + 'DurationScore']])
          Discrimmination.push([mal, o[mal + 'DiscriminationScore']])
          HypothesizedFunc.push([mal, o[mal + 'HF']])
          
          var offset = body.getChildIndex(ele.getParent());
          body.insertListItem(offset + 1, mal).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
          body.insertListItem(body.getChildIndex(body.findText("{Maladaptive2}").getElement().getParent()) + 1, mal).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
          body.insertParagraph(body.getChildIndex(body.findText("{Identified Antecedents & Consequences}").getElement().getParent()) + 1, '{'+ mal + 'Antecedent}')
          
          if(o.hasOwnProperty(mal+'Replace1')){
            for(var k = 4; k > -1 ; k--){
              body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Replace1'][k]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
            }
            body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified Functionally Equivalent Replacement Skills to be Taught:")
          }
          if(o.hasOwnProperty(mal+'Access')){
            if(Array.isArray( o[mal+'Access'])){
              for(var k = o[mal+'Access'].length-1; k > -1 ; k--){
                body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Access'][k]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
              }
            }
            else{
              body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Access']).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
            }
            body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Access To")
          }
          if(o.hasOwnProperty(mal+'Consequences')){
            if(Array.isArray( o[mal+'Consequences'])){
              for(var k = o[mal+'Consequences'].length; k > -1 ; k--){
                body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Consequences'][k])
              }
            }
            else{
              body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Consequences'])
            }
            body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified consequences that appear to be maintaining "+o['Name1']+"’s emission of non-compliance are as follows:")
          }
          if(o.hasOwnProperty(mal+'Antecedent')){
            if(Array.isArray( o[mal+'Antecedent'])){
              for(var k = o[mal+'Antecedent'].length-1; k > -1 ; k--){
                body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Antecedent'][k]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
              }
            }
            else{
              body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Antecedent']).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);
            }
            body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Antecedents that have been identified as initiating "+o['Name1']+"’s emission of non-compliance are as follows:")
            body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified Antecedents & Consequences: " + mal)
          }
          body.replaceText("{"+mal+"Antecedent}", '')
        }
        body.replaceText("{Identified Antecedents & Consequences}", '')
      }
      var scores = body.findText("{MaladaptiveScores}");
      var elem = scores.getElement();
      var off = body.getChildIndex(elem.getParent());
      body.insertTable(off + 1, Discrimmination)
      body.insertTable(off + 1, Duration)
      body.insertTable(off + 1, Intensity)
      body.insertTable(off + 1, Freq)
      body.replaceText("{MaladaptiveScores}", '')
      body.replaceText("{Maladaptive2}", '')
      body.replaceText("{Maladaptive Behaviors}", '')
      
      off = body.getChildIndex(body.findText("{Hypothesized Functions}").getElement().getParent())
      body.insertTable(off + 1, HypothesizedFunc)
      body.replaceText("{Hypothesized Functions}", '')
      
    }
    else if(x == 'reinforcerMethod'){
      var range = body.findText("{"+x+"}");
      var ele = range.getElement();
      var offset = body.getChildIndex(ele.getParent());
      if(Array.isArray(o[x])){
        for(var i = o[x].length; i > -1 ; i--)
          if(o[x] != undefined){
            body.insertListItem(offset + 1, o[x][i]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);;
          }
      }
      else{
       body.insertListItem(offset + 1, o[x]).setNestingLevel(1).setIndentStart(72).setGlyphType(DocumentApp.GlyphType.BULLET);; 
      }
      body.replaceText("{"+x+"}", '')
    }
    else if(crisisPlan.indexOf(x) > -1 && !crisisDone){
      
      for(var j = crisisPlan.length-1; j > -1; j--){
        if(crisisPlan[j] in o){
          body.insertTable(body.getChildIndex(body.findText("{Client Crisis Plan}").getElement().getParent()) + 1, [
                           ['Step #1: Warning Signs that a Crisis May be Developing (Thoughts, Images, Mood, Situation, Behavior)'],
                           ['1. '+ o[crisisPlan[j]+' Step1_1']],
                           ['2. '+ o[crisisPlan[j]+' Step1_2']],
                           ['3. '+ o[crisisPlan[j]+' Step1_3']],
                           ['Step #2: Internal Coping Strategies/Things the Child/Client Can Do W/Out Contacting Another Person (Relaxation Techniques, Physical Activity)'],
                           ['1. '+ o[crisisPlan[j]+' Step2_1']],
                           ['2. '+ o[crisisPlan[j]+' Step2_2']],
                           ['3. '+ o[crisisPlan[j]+' Step2_3']],
                           ['Step #3: People & Social Settings that Provide a Distraction'],
                           ['1. '+ o[crisisPlan[j]+' Step3_1']],
                           ['2. '+ o[crisisPlan[j]+' Step3_2']],
                           ['3. '+ o[crisisPlan[j]+' Step3_3']],
                           ['Step #4: People Whom to Contact for Help'],
                           ['1. '+ o[crisisPlan[j]+' Step4_1']],
                           ['2. '+ o[crisisPlan[j]+' Step4_2']],
                           ['3. '+ o[crisisPlan[j]+' Step4_3']],
                           ['Step #4: People Whom to Contact for Help'],
                           ['1. '+ o[crisisPlan[j]+' Step5_1']],
                           ['2. '+ o[crisisPlan[j]+' Step5_2']],
                           ['3. '+ o[crisisPlan[j]+' Step5_3']],
                           ['Step #4: People Whom to Contact for Help'],
                           ['1. '+ o[crisisPlan[j]+' Step6_1']],
                           ['2. '+ o[crisisPlan[j]+' Step6_2']],
                           ['3. '+ o[crisisPlan[j]+' Step6_3']]
                           ])

        }
      }   
      crisisDone = true
    }
    else if(x == 'Hours of One-to-One Treatment'){
      var range = body.findText("{"+x+"}");
      var elem = range.getElement();
      var off = body.getChildIndex(elem.getParent());
      var table=[]
      var BHS = ['BHS-MH', 'BHS', 'BHS-Unicare']
      var HCPCS = ['BHS-MH', 'MBHP', 'Optum (Allways Medicaid)', 'THP-P']
      var CPT = ['Aetna', 'BCBS', 'Cigna', 'THP-C']
      var Hybrid = ['BHS', 'BHS-Unicare']
            //QUERY FOR CLIENT PAYER==========================================================================================================
      var payer= 'BHS'
      var hours = 0, superHours = 0;
      if(BHS.includes(payer)){
        hours = o['Hours of One-to-One Treatment']
        superHours = hours * 0.15
      }
      else{
        hours = o['Hours of One-to-One Treatment']
        superHours = hours * 0.2
      }
      if(HCPCS.includes(payer)){
        table = [
          ['Code Set', 'Procedure Code', 'Modifier', 'Hours', 'Units', 'Frequency', 'Provider', 'ABA Treatment Service'],
          ['HCPCS', 'H0031', 'U2', '3', '12', 'per 6-Months', 'BCBA/LABA', 'Reassessment'],
          ['HCPCS', 'H2019', 'U2', hours, hours * 4, 'per week', 'BT', 'One-to-One ABA Treatment Services'],
          ['HCPCS', 'H0032', 'U2', superHours, superHours * 1, 'per week', 'BCBA/LABA', 'Direct Case Supervision'],
          ['HCPCS', 'H2012', 'U2', '1', '1', 'per week', 'BCBA/LABA', 'Parent Training (Individual)'],
          ['HCPCS', 'H0031', 'U2', '1', '4', 'per week', 'BCBA/LABA', 'Treatment Planning']
        ]
        
        if(o['IEP Meeting Date'] != null){
         table.push(['HCPCS', 'H0031', 'U2', '2', '8', 'per 6-months', 'BCBA/LABA', 'Attending Client’s IEP Meeting']) 
        }
      }
      else if(CPT.includes(payer)){
        table = [
          ['Code Set', 'Procedure Code', 'Modifier', 'Hours', 'Units', 'Frequency', 'Provider', 'ABA Treatment Service'],
          ['CPT', '97151', '', '8', '32', 'per 6-Months', 'BCBA/LABA', 'Behavior Identification Reassessment'],
          ['CPT', '97153', '', hours, hours * 4, 'per week', 'BT', 'One-to-One Adaptive Behavioral Treatment'],
          ['CPT', '97155', '', superHours, superHours *4, 'per week', 'BCBA/LABA', 'Adaptive Behavior Treatment with Protocol Modification'],
          ['CPT', '97156', '', '1', '4', 'per week', 'BCBA/LABA', 'Family Adaptive Behavior Treatment Guidance'],
          ['CPT', '97157', '', '1', '4', 'per month', 'BCBA/LABA', 'Multiple-Family Group Adaptive Behavior Treatment Guidance']
           ]
      }
      else if(Hybrid.includes(payer)){
        table = [
          ['Code Set', 'Procedure Code', 'Modifier', 'Hours', 'Units', 'Frequency', 'Provider', 'ABA Treatment Service'],
          ['CPT', '97151', '', '3', '12', 'per 6-Months', 'BCBA/LABA', 'Behavior Identification Reassessment'],
          ['CPT', '97153', '', hours, hours * 4, 'per week', 'BT', 'One-to-One Adaptive Behavioral Treatment'],
          ['CPT', '97155', '', superHours, superHours * 4, 'per week', 'BCBA/LABA', 'Adaptive Behavior Treatment with Protocol Modification'],
          ['CPT', '97156', '', '1', '4', 'per week', 'BCBA/LABA', 'Parent Training (Individual)'],
          ['HCPCS', '97157', '', '1', '4', 'per month', 'BCBA/LABA', 'Family Training (Group)'],
          ['HCPCS', 'H0032', '', '1', '4', 'per week', 'BCBA/LABA', 'Treatment Planning'],
          ['HCPCS', 'H0032', '', '1', '4', 'per 6-months', 'BCBA/LABA', 'Attending IEP Meeting']
        ]
      }
      
      body.insertTable(off+1, table)
      body.replaceText("{"+x+"}", payer)
    }
    else
      if(o[x] != undefined)
        body.replaceText("{"+x+"}", o[x])
        }
  if(o.hasOwnProperty("IEP File ID")){
    var img = DriveApp.getFileById(o['IEP File ID']).getBlob()
    
    var range = body.findText("{IEP File}");
    var ele = range.getElement();
    var offset = body.getChildIndex(ele.getParent());
    
    body.insertImage(offset + 1, img)
  }
  body.replaceText("{IEP File}", '')
  
  if(o.hasOwnProperty("PSS0Time")){
    var cont = true
    var n = 0;
    var table = [['Session Time', 'Session Activity Outline']]
    while(cont){
      if(o.hasOwnProperty('PSS' + n + 'Time')){
        table.push([o['PSS' + n + 'Time'], o['PSS' + 0 + 'Outline']])
        n++;
      }
      else{
        cont = false;
      }
    }
    body.insertTable(body.getChildIndex(body.findText("{Proposed Session Schedule}").getElement().getParent()), table)
  }
  body.replaceText("{Proposed Session Schedule}", '')
  
  if(o.hasOwnProperty("Vineland File ID")){
    var img = DriveApp.getFileById(o['Vineland File ID']).getBlob()
    
    var range = body.findText("{Vineland File}");
    var ele = range.getElement();
    var offset = body.getChildIndex(ele.getParent());
    
    body.insertImage(offset + 1, img)
  }
  body.replaceText("{Vineland File}", '')
  
  for(var category of crisisPlan){
    for(var risk of crisisRisks){
      if(o.hasOwnProperty(category)){
        if(o[category].indexOf(risk) > -1 || o[crisisPlan[j]] == risk){
          body.replaceText(category + ' ' +  risk, "✅")
        }
        else{
          body.replaceText(category + ' ' +  risk, " ")
        }
      }
      else{
        body.replaceText(category + ' ' +  risk, " ")
      }
    } 
  }
  
  //PARSE DATA GOES HERE
  if(skills != null){
    var range = body.findText("{Parsed Report}");
    var ele = range.getElement();
    var offset = body.getChildIndex(ele.getParent());
    
    len = skills[0].length
    
    if(skills[0] !== undefined || skills[0].length != 0){
      for(var i = len-1; i > -1; i--){
        var x = skills[0][i]
        var title =(x.title != null) ? x.title[0] : ''; 
        var goalType =(x.goal_type != null) ? x.goal_type  : ''; 
        var goal =(x.goal != null) ? x.goal[0] : ''; 
        var masteryCriteria =(x.mastery_criteria != null) ? x.mastery_criteria[0] : ''; 
        var startDate =(x.start_date != null) ? x.start_date[0] : ''; 
        var baseline = (x.baseline != null) ? x.baseline[0] : ''; 
        var objectives_mastered = (x.objectives_mastered != null) ? x.objectives_mastered[0] : ''; 
        var totalMasteredTargets = (x.total_mastered_targets != null) ?  x.total_mastered_targets[0] : ''; 
        var graphOptions =(x.graph_options != null) ? x.graph_options[0] : ''; 
        var image =(x.images != null) ? x.images[0] : ''; 
        var image2 =(x.images != null) ? x.images[1] : ''; 
        
        var goalCategory = (o.hasOwnProperty(title+'goal_category')) ? o[title+'goal_category'] : '';
        var maintenanceCriteria = (o.hasOwnProperty(title+'maintenance_criteria')) ? o[title+'maintenance_criteria'] : '';
        var masteryTarget = (o.hasOwnProperty(title+'mastery_target')) ? o[title+'mastery_target'] : '';
        var initialBaseline = (o.hasOwnProperty(title+'initial_baseline')) ? o[title+'initial_baseline'] : '';
        var program = (o.hasOwnProperty(title+'program')) ? o[title+'program'] : '';
        var goalStatus = (o.hasOwnProperty(title+'goal_status')) ? o[title+'goal_status'] : '';
        var progressReason = (o.hasOwnProperty(title+'progress_reason')) ? o[title+'progress_reason'] : '';
        var continuedTreatment = (o.hasOwnProperty(title+'continued_treatment')) ? o[title+'continued_treatment'] : '';
        var recommendation = (o.hasOwnProperty(title+'recommendation')) ? o[title+'recommendation'] : '';
        
          var table = [
            ["Title:", x.title[0]],                                                         //1
            ["Goal Type:", x.goal_type],                                                    //2
            ["Goal Category:", goalCategory],                                 //3
            ["Goal:", goal],                                                               //4
            ["Mastery Criteria:", masteryCriteria],                                          //5
            ["Maintenance Mastery Criteria:", maintenanceCriteria],           //6
            ["Start Date:", startDate],                                                     //7
            ["Mastery Target Date:", masteryTarget],                          //8
            ["Initial Baseline:", initialBaseline],                                //9
            ["Baseline:", baseline],                                                   //10
            ["Objectives Mastered:", objectives_mastered],                             //11
            ["Total Mastered Targets:", totalMasteredTargets],                       //12
            ["Generalization Planning/Programming:", program],                      //13
            ["Graph Options:", graphOptions],                                         //14
            ['', ''],                                                                          //15
            ["Current Progress Status:", goalStatus],                              //16
            ["Reason for Not Meeting Goal:", progressReason],                      //17
            ["Recommended for Continued Treatment:", continuedTreatment],                      //18
            ["Reason for Not Recommending for Continued Treatment:", recommendation]//19  
          ]
        
        var t = body.insertTable(offset + 1, table)
        
        if(x.hasOwnProperty('images')){
          if(Array.isArray(x['images'])){
            for(img of x['images']){
              var image = DriveApp.getFileById(img).getBlob()
              t.getCell(14, 0).appendImage(image)}
          }
          else{
            var image = DriveApp.getFileById(x.images).getBlob();
            t.getCell(14, 0).insertImage(0, image)
          }
        }
        body.insertParagraph(offset + 1, '\n' + skills[0][i]['title'])
      }
    }
    body.replaceText("{Parsed Report}", '')
    
    var range2 = body.findText("{MAL Goals}");
    var ele2 = range2.getElement();
    var offset2 = body.getChildIndex(ele2.getParent());
    
    len = skills[1].length
    if(skills[1] !== undefined && skills[1].length != 0){
      for(var i = len-1; i > -1; i--){
        var x = skills[0][i]
        var title =(x.title != null) ? x.title[0] : ''; 
        var goalType =(x.goal_type != null) ? x.goal_type  : ''; 
        var goalCategory = (o.hasOwnProperty(title+'goal_category')) ? o[title+'goal_category'] : '';
        var maintenanceCriteria = (o.hasOwnProperty(title+'maintenance_criteria')) ? o[title+'maintenance_criteria'] : '';
        var masteryTarget = (o.hasOwnProperty(title+'mastery_target')) ? o[title+'mastery_target'] : '';
        var initialBaseline = (o.hasOwnProperty(title+'initial_baseline')) ? o[title+'initial_baseline'] : '';
        var program = (o.hasOwnProperty(title+'program')) ? o[title+'program'] : '';
        var goalStatus = (o.hasOwnProperty(title+'goal_status')) ? o[title+'goal_status'] : '';
        var progressReason = (o.hasOwnProperty(title+'progress_reason')) ? o[title+'progress_reason'] : '';
        var continuedTreatment = (o.hasOwnProperty(title+'continued_treatment')) ? o[title+'continued_treatment'] : '';
        var recommendation = (o.hasOwnProperty(title+'recommendation')) ? o[title+'recommendation'] : '';
        
          var table = [
            ["Title:", x.title[0]],                                                         //1
            ["Goal Type:", x.goal_type],                                                    //2
            ["Goal Category:", goalCategory],                                 //3
            ["Goal:", goal],                                                               //4
            ["Mastery Criteria:", masteryCriteria],                                          //5
            ["Maintenance Mastery Criteria:", maintenanceCriteria],           //6
            ["Start Date:", startDate],                                                     //7
            ["Mastery Target Date:", masteryTarget],                          //8
            ["Initial Baseline:", initialBaseline],                                //9
            ["Baseline:", baseline],                                                   //10
            ["Objectives Mastered:", objectives_mastered],                             //11
            ["Total Mastered Targets:", totalMasteredTargets],                       //12
            ["Generalization Planning/Programming:", program],                      //13
            ["Graph Options:", graphOptions],                                         //14
            ['', ''],                                                                          //15
            ["Current Progress Status:", goalStatus],                              //16
            ["Reason for Not Meeting Goal:", progressReason],                      //17
            ["Recommended for Continued Treatment:", continuedTreatment],                      //18
            ["Reason for Not Recommending for Continued Treatment:", recommendation]//19  
          ]
        var t = body.insertTable(offset2 + 1, table)

        if(x.hasOwnProperty('images')){
          if(Array.isArray(x['images'])){
            for(img of x['images']){
              var image = DriveApp.getFileById(img).getBlob()
              t.getCell(14, 0).insertImage(0, image)}
          }
          else{
            var image = DriveApp.getFileById(x.images).getBlob();
            t.getCell(14, 0).insertImage(0, image)
          }
        }
        body.insertParagraph(offset2 + 1, '\n' + skills[1][i]['title'])
      }
    }    
    var range3 = body.findText("{MAL Goals}");
    var ele3 = range3.getElement();
    var offset3 = body.getChildIndex(ele3.getParent());
    
    len = skills[2].length
    if(skills[2] !== undefined && skills[2].length != 0){
      for(var i = len-1; i > -1; i--){
        var x = skills[2][i]
        var title =(x.title != null) ? x.title[0] : ''; 
        var goalType =(x.goal_type != null) ? x.goal_type  : ''; 
        var goal =(x.goal != null) ? x.goal[0] : ''; 
        var masteryCriteria =(x.mastery_criteria != null) ? x.mastery_criteria[0] : ''; 
        var startDate =(x.start_date != null) ? x.start_date[0] : ''; 
        var baseline = (x.baseline != null) ? x.baseline[0] : ''; 
        var objectives_mastered = (x.objectives_mastered != null) ? x.objectives_mastered[0] : ''; 
        var totalMasteredTargets = (x.total_mastered_targets != null) ?  x.total_mastered_targets[0] : ''; 
        var graphOptions =(x.graph_options != null) ? x.graph_options[0] : ''; 
      
        var goalCategory = (o.hasOwnProperty(title+'goal_category')) ? o[title+'goal_category'] : '';
        var maintenanceCriteria = (o.hasOwnProperty(title+'maintenance_criteria')) ? o[title+'maintenance_criteria'] : '';
        var masteryTarget = (o.hasOwnProperty(title+'mastery_target')) ? o[title+'mastery_target'] : '';
        var initialBaseline = (o.hasOwnProperty(title+'initial_baseline')) ? o[title+'initial_baseline'] : '';
        var program = (o.hasOwnProperty(title+'program')) ? o[title+'program'] : '';
        
        
        var table = [
          ["Title:", title],                                                         //1
          ["Goal Type:", goalType],                                                    //2
          ["Goal Category:", goalCategory],                                      //3
          ["Goal:", goal],                                                           //4
          ["Mastery Criteria:", masteryCriteria],                                   //5
          ["Maintenance Mastery Criteria:", maintenanceCriteria],                //6
          //          ["Start Date:", startDate],                                               //7
          ["Mastery Target Date:", masteryTarget],                               //8
          ["Initial Baseline:", initialBaseline],                                //9
          //          ["Baseline:", baseline],                                                   //10
          //          ["Objectives Mastered:", objectives_mastered],                             //11
          //          ["Total Mastered Targets:", totalMasteredTargets],                       //12
          ["Generalization Planning/Programming:", program]                      //13
          //          ["Graph Options:", graphOptions],                                         //14
          //          ['', ''],                                                                          //15
          //          ["Current Progress Status:", o[title+'goal_status']],                              //16
          //          ["Reason for Not Meeting Goal:", o[title+'progress_reason']],                      //16
          //          ["Reason for Not Recommending for Continued Treatment:", o[title+'recommendation']]//17  
        ]
        
        //        var img = DriveApp.getFileById(x.images[0]).getBlob();
        //        var img2 = DriveApp.getFileById(x.images[1]).getBlob()
        
        var t = body.insertTable(offset3 + 1, table)
        //        t.getCell(14, 0).insertImage(0, img)
        //        t.getCell(14, 0).insertImage(0, img2)
        body.insertParagraph(offset3 + 1, '\n' + skills[2][i]['title'])
      }
    }
    body.replaceText("{MAL Goals}", '')
    
    var range4 = body.findText("{Parent Training Goals}");
    var ele4 = range4.getElement();
    var offset4 = body.getChildIndex(ele4.getParent());
    
    len = skills[3].length
    if(skills[3] !== undefined && skills[3].length != 0){
      for(var i = len-1; i > -1; i--){
        var x = skills[3][i]
        var title =(x.title != null) ? x.title[0] : ''; 
        var goalType =(x.goal_type != null) ? x.goal_type  : ''; 
        if(goalType != 'New Goal'){
          var goal =(x.goal != null) ? x.goal[0] : ''; 
          var masteryCriteria =(x.mastery_criteria != null) ? x.mastery_criteria[0] : ''; 
          var startDate =(x.start_date != null) ? x.start_date[0] : ''; 
          var baseline = (x.baseline != null) ? x.baseline[0] : ''; 
          var objectives_mastered = (x.objectives_mastered != null) ? x.objectives_mastered[0] : ''; 
          var totalMasteredTargets = (x.total_mastered_targets != null) ?  x.total_mastered_targets[0] : ''; 
          var graphOptions =(x.graph_options != null) ? x.graph_options[0] : ''; 
          var image =(x.images != null) ? x.images[0] : ''; 
          var image2 =(x.images != null) ? x.images[1] : ''; 
          
          var goalCategory = (o.hasOwnProperty(title+'goal_category')) ? o[title+'goal_category'] : '';
          var maintenanceCriteria = (o.hasOwnProperty(title+'maintenance_criteria')) ? o[title+'maintenance_criteria'] : '';
          var masteryTarget = (o.hasOwnProperty(title+'mastery_target')) ? o[title+'mastery_target'] : '';
          var initialBaseline = (o.hasOwnProperty(title+'initial_baseline')) ? o[title+'initial_baseline'] : '';
          var program = (o.hasOwnProperty(title+'program')) ? o[title+'program'] : '';
          var goalStatus = (o.hasOwnProperty(title+'goal_status')) ? o[title+'goal_status'] : '';
          var progressReason = (o.hasOwnProperty(title+'progress_reason')) ? o[title+'progress_reason'] : '';
          var continuedTreatment = (o.hasOwnProperty(title+'continued_treatment')) ? o[title+'continued_treatment'] : '';
          var recommendation = (o.hasOwnProperty(title+'recommendation')) ? o[title+'recommendation'] : '';
          
          var table = [
            ["Title:", x.title[0]],                                                         //1
            ["Goal Type:", x.goal_type],                                                    //2
            ["Goal Category:", goalCategory],                                 //3
            ["Goal:", goal],                                                               //4
            ["Mastery Criteria:", masteryCriteria],                                          //5
            ["Maintenance Mastery Criteria:", maintenanceCriteria],           //6
            ["Start Date:", startDate],                                                     //7
            ["Mastery Target Date:", masteryTarget],                          //8
            ["Initial Baseline:", initialBaseline],                                //9
            ["Baseline:", baseline],                                                   //10
            ["Objectives Mastered:", objectives_mastered],                             //11
            ["Total Mastered Targets:", totalMasteredTargets],                       //12
            ["Generalization Planning/Programming:", program],                      //13
            ["Graph Options:", graphOptions],                                         //14
            ['', ''],                                                                          //15
            ["Current Progress Status:", goalStatus],                              //16
            ["Reason for Not Meeting Goal:", progressReason],                      //17
            ["Recommended for Continued Treatment:", continuedTreatment],                      //18
            ["Reason for Not Recommending for Continued Treatment:", recommendation]//19  
          ]
          
          var t = body.insertTable(offset4 + 1, table)
          
          if(x.hasOwnProperty('images')){
            if(Array.isArray(x['images'])){
              for(img of x['images']){
                var image = DriveApp.getFileById(img).getBlob()
                t.getCell(14, 0).insertImage(0, image)}
            }
            else{
              var image = DriveApp.getFileById(x.images).getBlob();
              t.getCell(14, 0).insertImage(0, image)
            }
          }
          
          body.insertParagraph(offset4 + 1, '\n' + skills[1][i]['title'])
        }
      }
    }
    body.replaceText("{Parent Training Goals}", '')
  }
  
//  file.setOwner(data[0][7])
//
//  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW)
//  //Edit permissions for the doc
////  var editors = file.getEditors()
////  for (var i = 0; i < editors.length; i++) {
////     file.removeEditor(editors[i]);
////    };
//  
//  file.addViewer('jlison@innovativeautism.org')
////  if(data[0][7] != '')
////    doc.addViewer(data[0][7])
  
  return file.getUrl();
}

//=================================Load Data from Save State Sheet==========================================

/*
 *  Get saved funder report input
 *    @param clientID: string of client id
 *
 *    if data[7] != null, then the report has been passed to a reviewer
 */

function getFunderSheetData(clientID){
  var sheetID = '1HQnH3hY-0YuAUOM_eP6JxAVMemhZBCAVQZJ_GZMcVT8';
  var spreadsheet = SpreadsheetApp.openById(sheetID).getSheets()[0];
  var data = spreadsheet.getDataRange().getDisplayValues();
  var user = Session.getActiveUser().getEmail()
  for(var i=data.length-1; i > -1; i--){
    if (data[i][0] == clientID && (data[i][7] == user && data[i][7] != undefined)){
      Logger.log(data[i]);
      return [data[i], 'Reviewer'];
    }
    else if(data[i][0] == clientID){
     return [data[i], ''];
    }
    else{
      return 'No report on file.' 
    }
  }
}

/*
 *  Save progress report images to drive
 *    @param data: image file from parse in base64
 *    @param file: filename
 *    @param objName: file object name
 */

function saveToDrive(data, file, objName) {
  var email = Session.getActiveUser().getEmail();
  try {
    
    var dropbox = "My Dropbox";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    
    var contentType = data.substring(5,data.indexOf(';')),
        bytes = Utilities.base64Decode(data.substr(data.indexOf('base64,')+7)),
        blob = Utilities.newBlob(bytes, contentType, file);
    
    var fName = folder.createFolder([email].join(" ")).createFile(blob);
    
    return [fName.getName().split('.').slice(0, -1).join('.'), fName.getId(), objName];
    
  } catch (f) {
    return f.toString();
  }
}

/*
 *  Save signature to drive and return drive file id
 *    @param data: base64 encoded image
 */

function saveSignature(data){  
  try{
  var decoded = Utilities.base64Decode(data.substr(data.indexOf('base64,')+7));
  var blob = Utilities.newBlob(decoded, MimeType.PNG, "nameOfImage");
  
  var dropbox = "My Dropbox";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
  
  var signature = folder.createFile(blob);
  
  return signature.getId()
  }
  catch(e){
    return e 
  }
}

/*
 *  Parse vineland report and return the parsed summary
 */

function vinelandParse(){
//  var Doc = DriveApp.getFilesByName(fileName).next().getAs(contentType)
  var body = DocumentApp.openById('1BDY4-h-e9AA0x_WzlHKNq_2DHZxoxyuT').getBody().getText()
  var regex = /(?<=OVERALL SUMMARY\s+).*?(?=\s+Comprehensive Parent\/Caregiver Form Report)/gs
  var summary = body.match(regex)[0]
  Logger.log(summary)
  return summary
  } 

/*
 *  Parse progress report file from google drive as array of objects.
 *     @param fileName: string name of file in drive app.
 */

function beaconParse(fileName) {  
  var files = DriveApp.getFilesByName(fileName)
  while (files.hasNext()) {
    var file = files.next();
    var docID = file.getId();
  }
  var doc = DocumentApp.openById(docID);
  
  var text = doc.getBody().getText();
  var images = doc.getBody().getImages();
  
  //regex for major sections of the report
  //HEADER DATA IS WRITTEN WEIRD SO REGEX IS CONVOLUTED
  let headerData = /Report Name:\s+(?<reportname>[^\n]+)(?<reportdate>[0-9\.]+)?\s?\nClient First Name:\s+(?<ClientFName>\w+(?:\s\w+)?)\s\nClient Last Name:\s+(?<ClientLName>\w+(?:-\w+)?)\s+Skill Acquisition Progress\s+Start Date:\s+ (?<startDate>\d{2}\/\d{2}\/\d{4})\s\n+Skill Acquisition Progress\s+End Date:\s+ (?<endDate>\d{2}\/\d{2}\/\d{4})\s+Skill Acquisition Progress\s+Graph Type:\s+ (?<graphType>\w+\s+\w+)\s+Skill Acquisition Progress\s+View By:\s+ (?<viewBy>\w+\s+\w+)/;
  
  var header = headerData.exec(text)
  var imageList = []
  var now = new Date();
  for(var i = 0; i < images.length; i++){
    var prevSibling = images[i].getParent().getPreviousSibling();
    
    while(!/^\d+\. (?:\S\s)*/.test(prevSibling.getText())){
      prevSibling = prevSibling.getPreviousSibling();
    }
    Logger.log(prevSibling.getText());
    
    var goalTitle = prevSibling.getText();
    var email = Session.getActiveUser().getEmail();
    
    var dropbox = "Parsed Report Files";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    
    var userFolder, userFolders = folder.getFoldersByName([email].join(" "));
    if (userFolders.hasNext()) {
      userFolder = userFolders.next();
    } else {
      userFolder = folder.createFolder([email].join(" "));
    }
    

    
    var reportFolder, reportFolders = userFolder.getFoldersByName(now);
    if (reportFolders.hasNext()) {
      reportFolder = reportFolders.next();
    } else {
      reportFolder = userFolder.createFolder(now).setOwner('claggui@innovativeautism.org').addViewer(email);
    }
    var imageFile = reportFolder.createFile(images[i].getBlob().setName(goalTitle));
    var imageID = imageFile.getId();
    var image = {id: imageID, title: goalTitle}
    imageList.push(image);
  }
  
  let objectives = /\w+\s?\d?\:\s[^\n]*/g
  let goals = /(?<=\nGoal\: )[^\n]*/g
  let objective1 = /(?<=\nObjective 1\: )[^\n]*/g
  let objective2 = /(?<=\nObjective 2\: )[^\n]*/g
  let objective3 = /(?<=\nObjective 3\: )[^\n]*/g
  let objective4 = /(?<=\nObjective 4\: )[^\n]*/g
  let mastery = /(?<=\nMastery criteria\: )[^\n]*/g
  let maintenance = /(?<=\nMaintenance mastery criteria\: )[^\n]*/g
  let start = /(?<=\nStart Date\: )[^\n]*/g
  let base = /(?<=\nBaseline\: )[^\n]*/g
  let mastered = /(?<=\nObjectives mastered\: )[^\n]*/g
  let total = /(?<=\nTotal Mastered Targets\: )[^\n]*/g
  let options= /(?<=\nGraph Options\: )[^\n]*/g
  let title= /\d+\. \w+[^\n]*/g
  let parentTraining= /Parent/ig
  let Sections = /(\nProgress Data\s*)([\s\S]*?)(?=\nProgress Data|$)/g
  let SectionName = /(?<=\nProgress Data\s*)(\S+\s)+/gm
  var sName, SAP, BRP, RG, SAPObj, BRPObj, RGObj, skills = [], acquisitionGoals = [], brpGoals=[],recommendedGoals=[], ptGoals=[];
  let Tasks = /(\n\d+\.\s\w+\s*)([\s\S]*?)(?=\n+\d+\.\s*|$)/g
  
  var Body = text.match(Sections)
  for(var i=0; i < Body.length; i++){
    if(Body[i].match(SectionName)[0] == "Skill Acquisition Progress "){
      SAP = Body[i]
      SAPObj = SAP.match(Tasks)
    }
    if(Body[i].match(SectionName)[0] == "Beh Reduction Progress "){
      BRP = Body[i]
      BRPObj = BRP.match(Tasks)
    }
    if(Body[i].match(SectionName)[0] == "Recommended Goals "){
      RG = Body[i]
      RGObj = RG.match(Tasks)
    }
  }
  if(!SAPObj==''){
    var objID=[]
    for(var i=0; i< SAPObj.length; i++){
      for(var j=0; j< imageList.length; j++){
        if(imageList[j].title == SAPObj[i].match(title)[0])
        objID.push(imageList[j].id);
      }
      var objectiveObj = {
        title: SAPObj[i].match(title),
        goal: SAPObj[i].match(goals),
        goal_type: 'Current Goal',
        objective_1: SAPObj[i].match(objective1),
        objective_2: SAPObj[i].match(objective2),
        objective_3: SAPObj[i].match(objective3),
        objective_4: SAPObj[i].match(objective4),
        mastery_criteria: SAPObj[i].match(mastery),
        maintenance_mastery: SAPObj[i].match(maintenance),
        start_date: SAPObj[i].match(start),
        baseline: SAPObj[i].match(base),
        images: objID,
        objectives_mastered: SAPObj[i].match(mastered),
        total_mastered_targets: SAPObj[i].match(total),
        graph_options: SAPObj[i].match(options)
      }
      if(objectiveObj.title[0].match(parentTraining))
        ptGoals.push(objectiveObj);
      else
        acquisitionGoals.push(objectiveObj);
      objID=[]
    }
  }
  skills.push(acquisitionGoals)
  if(!BRPObj ==''){
    var objID=[]
    for(var i=0; i< BRPObj.length; i++){
      for(var j=0; j< imageList.length; j++){
        if(imageList[j].title == SAPObj[i].match(title)[0])
        objID.push(imageList[j].id);
      }
      var objectiveObj = {
        title: BRPObj[i].match(title),
        goal: BRPObj[i].match(goals),
        goal_type: 'Current Goal',
        objective_1: BRPObj[i].match(objective1),
        objective_2: BRPObj[i].match(objective2),
        objective_3: BRPObj[i].match(objective3),
        objective_4: SAPObj[i].match(objective4),
        mastery_criteria: BRPObj[i].match(mastery),
        maintenance_mastery: BRPObj[i].match(maintenance),
        start_date: BRPObj[i].match(start),
        baseline: BRPObj[i].match(base),
        images: objID,
        objectives_mastered: BRPObj[i].match(mastered),
        total_mastered_targets: BRPObj[i].match(total),
        graph_options: BRPObj[i].match(options)
      }
      if(objectiveObj.title[0].match(parentTraining))
        ptGoals.push(objectiveObj);
      else
        brpGoals.push(objectiveObj);
      objID=[]
    }
  }
  skills.push(brpGoals)
  if(!RGObj==''){
    for(var i=0; i< RGObj.length; i++){
      var objectiveObj = {
        title: RGObj[i].match(title),
        goal: RGObj[i].match(goals),
        goal_type: 'New Goal',
        objective_1: RGObj[i].match(objective1),
        objective_2: RGObj[i].match(objective2),
        objective_3: RGObj[i].match(objective3),
        objective_4: SAPObj[i].match(objective4),
        mastery_criteria: RGObj[i].match(mastery),
        maintenance_mastery: RGObj[i].match(maintenance),
        start_date: RGObj[i].match(start),
        baseline: RGObj[i].match(base),
        objectives_mastered: RGObj[i].match(mastered),
        total_mastered_targets: RGObj[i].match(total),
        graph_options: RGObj[i].match(options)
      }
      if(objectiveObj.title[0].match(parentTraining))
        ptGoals.push(objectiveObj);
      else
        recommendedGoals.push(objectiveObj);
    }
  }
  skills.push(recommendedGoals)
  skills.push(ptGoals)
  Logger.log(skills[0])
  return skills
}

/*
 *  Intermediate function; expendable.
 *
 *    convertDocuments && convertToGoogleDocs_
 *      @param file/fileName: drive app file to be converted to native gsuite file;
 *          docx -> doc
 *
 *    calls beaconparse to parse converted doc
 */

function convertDocuments(file) {  
  return convertToGoogleDocs_(file)  
}


// By Google Docs, we mean the native Google Docs format
function convertToGoogleDocs_(fileName) {
  
  var officeFile = DriveApp.getFilesByName(fileName).next();
  
  // Use the Advanced Drive API to upload the Excel file to Drive
  // convert = true will convert the file to the corresponding Google Docs format
  
  var uploadFile = JSON.parse(UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v2/files?uploadType=media&convert=true",
    {
      method: "POST",
      contentType: officeFile.getMimeType(),
      payload: officeFile.getBlob().getBytes(),
      headers: {
        "Authorization" : "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    }
  ).getContentText());
  
  // Remove the file extension from the original file name
  var googleFileName = fileName.split('.').slice(0, -1).join('.');
  
  // Update the name of the Google Sheet created from the Excel sheet
  DriveApp.getFileById(uploadFile.id).setName(googleFileName);
  return beaconParse(googleFileName);
}