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
  return HtmlService.createTemplateFromFile("Search").evaluate();
}

//=========TEST FUNCTIONS===========

function clientSearch(){
  console.log("Searching for client...");
}

function getUnreadEmails() {
  return GmailApp.getInboxUnreadCount();
}

//==================================

function getClientName() {
  var spreadsheetID = "1G0kpCNdmow74d0u7pGCbN7ubFE4lmMFbVMCnlWPbt0I";
  var clientNames = SpreadsheetApp.openById(spreadsheetID).getDataRange().getDisplayValues();
  return clientNames; 
}

function getUser(){
  return Session.getActiveUser().getEmail()
}

function getClientInfo(clientName, clientID) {
  var clientName = clientName || "NA";
  var clientID =  clientID||"NA";
  Logger.log(clientName);
  var clientData = [];
  var spreadsheetID = "1G0kpCNdmow74d0u7pGCbN7ubFE4lmMFbVMCnlWPbt0I";
  
  //Need to make multiple variables to hold several sheets/pages (Statuses/Assignments, Client Info, Client Insurance)
  
  //Client Status/Assignment
  var data = getClientName();
  
  //Loop for sheets that have multiplicities Eg. Insurance sheet
  // data[i][3] Second Index is location of FULL NAME in the spreadsheet. CHANGE ACCORDINGLY
  for (var i = 0; i < data.length; i++){
    if (data[i][3] == clientName || data[i][0] == clientID){
      clientData.push(data[i]);
    }
  }
  
  //Contact Info
  var contacts = SpreadsheetApp.openById(spreadsheetID).getSheetByName("Sheet2").getDataRange().getDisplayValues();
  for(var i=0; i < contacts.length;i++){
    if(contacts[i][1] == clientName || contacts[i][0] == clientID){
      clientData.push(contacts[i]);
    }
  }
  Logger.log(clientData[1]);
  
  //Client Insurance
  // ins[i][1] Second Index is location of FULL NAME in the spreadsheet. CHANGE ACCORDINGLY
  //Sheet3 -> Insurance Sheet
  var ins = SpreadsheetApp.openById(spreadsheetID).getSheetByName("Sheet3").getDataRange().getDisplayValues();
  for(var i=0 ; i < ins.length; i++){
    if(ins[i][1] == clientName  || ins[i][0] == clientID){
      clientData.push(ins[i]);
    }
  }
  return clientData;
}


function test(){
  var b = "Alexia Williams"
  var c = "000288"
  var a = getClientInfo(undefined,c);
  Logger.log(a[1][0]);
  
  //Logger.log(data);
}

//======================= Save data to sheet ===========================

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

function makeTemplate(data, skills){
  var file= DriveApp.getFileById('1_Bw_WeQnjnIdYI0teO2-00DKBVhwR4caqwq69SwyIn4').makeCopy(data[0][2] + " " + data[0][0] + " TX PLAN TEMPLATE")
  var body= DocumentApp.openById(file.getId()).getBody();
  
  var o = JSON.parse(data[0][4])
  var crisisPlan = ['Assaultive Behavior','Self-Injurious Behavior', 'Fire Setting', 'Impulsive Behavior', 'Current Family Abuse Violence', 'Elopement/Bolting', 'Sexually Offending Behavior', 'Substance Abuse', 'Psychotic Symptoms', 'Coping with Significant Loss', 'Suicidality', 'Homicidality']
  var crisisDone = false;
  
  for(x in o){
    if(x == 'Communication Skills Domain'){
      var range = body.findText("{"+x+"}");
      var ele = range.getElement();
      if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
        var offset = body.getChildIndex(ele.getParent());
        if(o['Physical Activity Skills Domain'] != undefined){
          var PASD = 'Physical Activity Skills Domain: ' + o['Physical Activity Skills Domain']
          body.insertListItem(offset + 1, PASD);
        }
        if(o['Daily Living Skills Domain'] != undefined){
          var DLSD = 'Daily Living Skills Domain: ' + o['Daily Living Skills Domain']
          body.insertListItem(offset + 1, DLSD);
        }
        if(o['Social Skills & Relationship Skills Domain'] != undefined){
          var SSRSD = 'Social Skills & Relationship Skills Domain: ' + o['Social Skills & Relationship Skills Domain']
          body.insertListItem(offset + 1, SSRSD);
        }
        if(o[x] != undefined){
          var CDS = 'Communication Skills Domain: ' + o[x]
          body.insertListItem(offset + 1, CDS);
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
        for(var i = 0; i < o['Medications'].length; i++){
          medications.push([o['Medications'][i], o['Medications Dosage'][i], o['Medications Frequency'][i]])
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
        for(var i = 0; i < o['Vitamins'].length; i++){
          vitamins.push([o['Vitamins'][i], o['Vitamins Dosage'][i], o['Vitamins Frequency'][i]])
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
          body.insertListItem(offset + 1, PASD);
        }
        if(o['Current Daily Living Skills Domain'] != undefined){
          var DLSD = 'Daily Living Skills Domain: ' + o['Current Daily Living Skills Domain']
          body.insertListItem(offset + 1, DLSD);
        }
        if(o['Current Social Skills & Relationship Skills Domain'] != undefined){
          var SSRSD = 'Social Skills & Relationship Skills Domain: ' + o['Current Social Skills & Relationship Skills Domain']
          body.insertListItem(offset + 1, SSRSD);
        }
        if(o[x] != undefined){
          var CDS = 'Communication Skills Domain: ' + o[x]
          body.insertListItem(offset + 1, CDS);
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
      var Discrimmination = [['Maladaptive Behavior', 'Discrimmination Score']]
      var HypothesizedFunc = [['Maladaptive Behavior', 'Hypothesized Function(s)']];
      if (ele.getParent().getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
        
        /*
        *Current problem: function does not work if property has less than one answer/is not array. Must select more than one Maladaptive Behavior.
        *
        */
        for(var j = o['Maladaptive Behaviors'].length-1; j > -1 ; j--){
          let mal = o[x][j]
          Freq.push([mal, o[mal + 'FreqScore']])
          Intensity.push([mal, o[mal + 'IntensityScore']])
          Duration.push([mal, o[mal + 'DurationScore']])
          Discrimmination.push([mal, o[mal + 'DiscriminationScore']])
          HypothesizedFunc.push([mal, o[mal + 'HF']])
          
          var offset = body.getChildIndex(ele.getParent());
          body.insertListItem(offset + 1, o['Maladaptive Behaviors'][j]);
          body.insertListItem(body.getChildIndex(body.findText("{Maladaptive2}").getElement().getParent()) + 1, o['Maladaptive Behaviors'][j]);
          body.insertParagraph(body.getChildIndex(body.findText("{Identified Antecedents & Consequences}").getElement().getParent()) + 1, '{'+ mal + 'Antecedent}')
          
          for(var k = o[mal+'Replace1'].length-1; k > -1 ; k--){
            body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Replace1'][k])
          }
          body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified Functionally Equivalent Replacement Skills to be Taught:")

          
          for(var k = o[mal+'Access'].length-1; k > -1 ; k--){
            body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Access'][k])
          }
          body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Access To")

          
          for(var k = o[mal+'Consequences'].length; k > -1 ; k--){
            body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Consequences'][k])
          }
          body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified consequences that appear to be maintaining "+o['Name1']+"’s emission of non-compliance are as follows:")

          for(var k = o[mal+'Antecedent'].length-1; k > -1 ; k--){
            body.insertListItem(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, o[mal+'Antecedent'][k])
          }
          body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Antecedents that have been identified as initiating "+o['Name1']+"’s emission of non-compliance are as follows:")
          body.insertParagraph(body.getChildIndex(body.findText("{"+mal+"Antecedent}").getElement().getParent()) + 1, "Identified Antecedents & Consequences: " + mal)
          
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
      for(var i = o[x].length; i > -1 ; i--)
        if(o[x] != undefined){
          body.insertListItem(offset + 1, o[x][i]);
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
    else
      if(o[x] != undefined)
        body.replaceText("{"+x+"}", o[x])
  }
  
  //PARSE DATA GOES HERE
  if(skills != ''){
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

        
        var table = [
          ["Title:", title],                                                         //1
          ["Goal Type:", goalType],                                                    //2
          ["Goal Category:", o[title+'goal_category']],                                      //3
          ["Goal:", goal],                                                           //4
          ["Mastery Criteria:", masteryCriteria],                                   //5
          ["Maintenance Mastery Criteria:", o[title+'maintenance_criteria']],                //6
          ["Start Date:", startDate],                                               //7
          ["Mastery Target Date:", o[title+'mastery_target']],                               //8
          ["Initial Baseline:", o[title+'initial_baseline']],                                //9
          ["Baseline:", baseline],                                                   //10
          ["Objectives Mastered:", objectives_mastered],                             //11
          ["Total Mastered Targets:", totalMasteredTargets],                       //12
          ["Generalization Planning/Programming:", o[title+'program']],                      //13
          ["Graph Options:", graphOptions],                                         //14
          ['', ''],                                                                          //15
          ["Current Progress Status:", o[title+'goal_status']],                              //16
          ["Reason for Not Meeting Goal:", o[title+'progress_reason']],                      //16
          ["Reason for Not Recommending for Continued Treatment:", o[title+'recommendation']]//17  
        ]
        
        var img = DriveApp.getFileById(image).getBlob();
        var img2 = DriveApp.getFileById(image2).getBlob()
        
        var t = body.insertTable(offset + 1, table)
        t.getCell(14, 0).insertImage(0, img)
        t.getCell(14, 0).insertImage(0, img2)
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
        var goal =(x.goal != null) ? x.goal[0] : ''; 
        var masteryCriteria =(x.mastery_criteria != null) ? x.mastery_criteria[0] : ''; 
        var startDate =(x.start_date != null) ? x.start_date[0] : ''; 
        var baseline = (x.baseline != null) ? x.baseline[0] : ''; 
        var objectives_mastered = (x.objectives_mastered != null) ? x.objectives_mastered[0] : ''; 
        var totalMasteredTargets = (x.total_mastered_targets != null) ?  x.total_mastered_targets[0] : ''; 
        var graphOptions =(x.graph_options != null) ? x.graph_options[0] : ''; 
        var image =(x.images != null) ? x.images[0] : ''; 
        var image2 =(x.images != null) ? x.images[1] : ''; 
        
        var table = [
          ["Title:", title],                                                         //1
          ["Goal Type:", goalType],                                                    //2
          ["Goal Category:", o[title+'goal_category']],                                      //3
          ["Goal:", goal],                                                           //4
          ["Mastery Criteria:", masteryCriteria],                                   //5
          ["Maintenance Mastery Criteria:", o[title+'maintenance_criteria']],                //6
          ["Start Date:", startDate],                                               //7
          ["Mastery Target Date:", o[title+'mastery_target']],                               //8
          ["Initial Baseline:", o[title+'initial_baseline']],                                //9
          ["Baseline:", baseline],                                                   //10
          ["Objectives Mastered:", objectives_mastered],                             //11
          ["Total Mastered Targets:", totalMasteredTargets],                       //12
          ["Generalization Planning/Programming:", o[title+'program']],                      //13
          ["Graph Options:", graphOptions],                                         //14
          ['', ''],                                                                          //15
          ["Current Progress Status:", o[title+'goal_status']],                              //16
          ["Reason for Not Meeting Goal:", o[title+'progress_reason']],                      //16
          ["Reason for Not Recommending for Continued Treatment:", o[title+'recommendation']]//17  
        ]
        
        var img = DriveApp.getFileById(x.images[0]).getBlob();
        var img2 = DriveApp.getFileById(x.images[1]).getBlob()
        
        var t = body.insertTable(offset2 + 1, table)
        t.getCell(14, 0).insertImage(0, img)
        t.getCell(14, 0).insertImage(0, img2)
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
        var image =(x.images != null) ? x.images[0] : ''; 
        var image2 =(x.images != null) ? x.images[1] : ''; 
        
        var table = [
          ["Title:", title],                                                         //1
          ["Goal Type:", goalType],                                                    //2
          ["Goal Category:", o[title+'goal_category']],                                      //3
          ["Goal:", goal],                                                           //4
          ["Mastery Criteria:", masteryCriteria],                                   //5
          ["Maintenance Mastery Criteria:", o[title+'maintenance_criteria']],                //6
//          ["Start Date:", startDate],                                               //7
          ["Mastery Target Date:", o[title+'mastery_target']],                               //8
          ["Initial Baseline:", o[title+'initial_baseline']],                                //9
//          ["Baseline:", baseline],                                                   //10
//          ["Objectives Mastered:", objectives_mastered],                             //11
//          ["Total Mastered Targets:", totalMasteredTargets],                       //12
          ["Generalization Planning/Programming:", o[title+'program']]                      //13
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
          
          var goalCategory = (o[title+'goal_category'] != null && o.hasOwnProperty(title+'goal_category')) ? o[title+'goal_category'] : '';
          var maintenanceCriteria = (o[title+'maintenance_criteria'] != null && o.hasOwnProperty(title+'maintenance_criteria')) ? o[title+'maintenance_criteria'] : '';
          var masteryTarget = (o[title+'mastery_target'] != null && o.hasOwnProperty(title+'mastery_target')) ? o[title+'mastery_target'] : '';
          var program = (o[title+'program'] != null && o.hasOwnProperty(title+'program')) ? o[title+'program'] : '';
          var goalStatus = (o[title+'goal_status'] != null && o.hasOwnProperty(title+'goal_status')) ? o[title+'goal_status'] : '';
          var progressReason = (o[title+'progress_reason'] != null && o.hasOwnProperty(title+'progress_reason')) ? o[title+'progress_reason'] : '';
          var continuedTreatment = (o[title+'continued_treatment'] != null && o.hasOwnProperty(title+'continued_treatment')) ? o[title+'continued_treatment'] : '';
          var recommendation = (o[title+'recommendation'] != null && o.hasOwnProperty(title+'recommendation')) ? o[title+'recommendation'] : '';
          
          var table = [
            ["Title:", x.title[0]],                                                         //1
            ["Goal Type:", x.goal_type],                                                    //2
            ["Goal Category:", goalCategory],                                 //3
            ["Goal:", goal],                                                               //4
            ["Mastery Criteria:", masteryCriteria],                                          //5
            ["Maintenance Mastery Criteria:", maintenanceCriteria],           //6
            ["Start Date:", startDate],                                                     //7
            ["Mastery Target Date:", masteryTarget],                          //8
            ["Initial Baseline:", "o[title+'initial_baseline']"],                                //9
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
            var img = DriveApp.getFileById(x.images[0]).getBlob();
            var img2 = DriveApp.getFileById(x.images[1]).getBlob()
            
            t.getCell(14, 0).insertImage(0, img)
            t.getCell(14, 0).insertImage(0, img2)
          }
          
          body.insertParagraph(offset2 + 1, '\n' + skills[1][i]['title'])
        }
      }
    }
    body.replaceText("{Parent Training Goals}", '')
  }
  
  return file.getUrl();
}

//=================================Load Data from Save State Sheet==========================================

function getFunderSheetData(clientID){
    var sheetID = '1HQnH3hY-0YuAUOM_eP6JxAVMemhZBCAVQZJ_GZMcVT8';
    var spreadsheet = SpreadsheetApp.openById(sheetID).getSheets()[0];
    var data = spreadsheet.getDataRange().getDisplayValues();
    
    for(var i=0; i < data.length; i++){
      if (data[i][0] == clientID){
        Logger.log(data[i]);
        return data[i];
      }
    }
}

function test2(){
  var body= DocumentApp.openById('1_Bw_WeQnjnIdYI0teO2-00DKBVhwR4caqwq69SwyIn4').getBody();
  var element;
    element = body.getChild(5)
//    Logger.log(element.asParagraph().getText())
        
//    Logger.log(element.getType())
    
  var searchType = DocumentApp.ElementType.PARAGRAPH;
  var searchHeading = DocumentApp.ParagraphHeading.HEADING2;
  var searchResult = null;

  while (searchResult = body.findElement(searchType, searchResult)) {
    var par = searchResult.getElement().asParagraph();
    if (par.getHeading() == searchHeading) {
      // Found one, update Logger.log and stop.
      var h = searchResult.getElement().asText().getText();
      
      if(h == 'ABA ASSESSMENT/TREATMENT PLAN TYPE:'){
         Logger.log(h)
      }
    }
  }

  
  
}

function saveToDrive(data, file, name, email) {
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

    var fName = folder.createFolder([email].join(" ")).createFile(blob).getName();

    return fName.split('.').slice(0, -1).join('.');;

  } catch (f) {
    return f.toString();
  }
}

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
  let headerData = /Report Name:\s+(?<reportname>(?:[\w]+\s+)+)(?<reportdate>[0-9\.]+)?\s?\nClient First Name:\s+(?<ClientFName>\w+(?:\s\w+)?)\s\nClient Last Name:\s+(?<ClientLName>\w+(?:-\w+)?)\s+Skill Acquisition Progress\s+Start Date:\s+ (?<startDate>\d{2}\/\d{2}\/\d{4})\s\n+Skill Acquisition Progress\s+End Date:\s+ (?<endDate>\d{2}\/\d{2}\/\d{4})\s+Skill Acquisition Progress\s+Graph Type:\s+ (?<graphType>\w+\s+\w+)\s+Skill Acquisition Progress\s+View By:\s+ (?<viewBy>\w+\s+\w+)/;
  
   var header = headerData.exec(text)
   var imageList = []
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
       
       var reportFolder, reportFolders = userFolder.getFoldersByName(`${header.groups.reportname}`);
      if (reportFolders.hasNext()) {
        reportFolder = reportFolders.next();
      } else {
        reportFolder = userFolder.createFolder(`${header.groups.reportname}`);
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

function testp(){
  Logger.log(convertDocuments('Rethink Report Example 4 05.28.20.docx')[0][0].title)  
  
}