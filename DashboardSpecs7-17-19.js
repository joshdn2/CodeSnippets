/*
Made by Josh D
07/2019

Takes the list of account on the Data tab, goes through the all agents view for each one. 
There, it finds the agents that have tracking changes turned on and it gets the job stats for the most recent batch. 
Then, it finds the agents that have active jobs and it gets the agent stats for each of those. 
All results are output to their corresponding sheets. 

Stable build 7-17-19
Added:
	Credits used for each agent
	Functions to run each line from data tab

*/
var spreadsheetID = "";
var spreadsheet = SpreadsheetApp.openById(spreadsheetID);
var tempSheet = spreadsheet.getSheetByName("Agent Status");
var changesSheet = spreadsheet.getSheetByName("Tracking Changes");
var downloadsSheet = spreadsheet.getSheetByName("Downloads");
var dataSheet = spreadsheet.getSheetByName("Data");

function runAccounts() {//clears all output sheets, makes a list from the departments and WSKs on Data sheet and runs getAgentInfo for each one 
  clearSheet("Agent Status"); 
  clearSheet("Tracking Changes");
  clearSheet("Downloads");
  var values = dataSheet.getRange("A2:C" + dataSheet.getLastRow()).getValues();
  var allStatus = [];
  var allHistory = [];
  for (var q = 0; q < values.length; q++ ){
    var wsk = values[q][0];
    var dept = values[q][1];
    var view = values[q][2];
    var info = getAgentInfo(wsk,dept,view);
    allStatus = allStatus.concat(info["Statuses"]);
    allHistory = allHistory.concat(info["History"]);
  }
  if (allHistory != undefined){
    var allDownloads = [];
    //check each index here to make sure it isn't undefined

    for (var k = 0; k < allHistory.length; k++){
      if(allHistory[k] != undefined){
      if (parseInt(allHistory[k]["Total Downloads"]) > 0){allDownloads.push(allHistory[k]);}
      }
    }
    setObjectsToPage(changesSheet,allHistory);
    setObjectsToPage(downloadsSheet,allDownloads);
  }
  setObjectsToPage(tempSheet,allStatus);
}

function runFirstView()//We have all the firstView, secondView, etc. because all of the views together take more than the 30 minute time limit in google scripts
{
  var row = 1;
  clearSheet("Agent Status"); 
  clearSheet("Tracking Changes");
  clearSheet("Downloads");
  var values1 = [];
  values1 = dataSheet.getRange("A" + (row+1) + ":C" + (row+1)).getValues();
  var wsk = values1[0][0];
  var dept = values1[0][1];
  var view = values1[0][2];
  var allStatus = [];
  var allHistory = [];
  var info = getAgentInfo(wsk,dept,view);
  allStatus = allStatus.concat(info["Statuses"]);
  allHistory = allHistory.concat(info["History"]);
  if (allHistory != undefined){
    var allDownloads = [];
    for (var k = 0; k < allHistory.length; k++){
      if(allHistory[k] != undefined){
      if (parseInt(allHistory[k]["Total Downloads"]) > 0){allDownloads.push(allHistory[k]);}
      }
    }
    setObjectsToPage(changesSheet,allHistory);
    setObjectsToPage(downloadsSheet,allDownloads);
  }
  setObjectsToPage(tempSheet,allStatus);
}

function runSecondView()
{
  var row = 2;
  var values1 = [];
  values1 = dataSheet.getRange("A" + (row+1) + ":C" + (row+1)).getValues();
  var wsk = values1[0][0];
  var dept = values1[0][1];
  var view = values1[0][2];
  var allStatus = [];
  var allHistory = [];
  var info = getAgentInfo(wsk,dept,view);
  allStatus = allStatus.concat(info["Statuses"]);
  allHistory = allHistory.concat(info["History"]);
  if (allHistory != undefined){
    var allDownloads = [];
    for (var k = 0; k < allHistory.length; k++){
      if(allHistory[k] != undefined){
      if (parseInt(allHistory[k]["Total Downloads"]) > 0){allDownloads.push(allHistory[k]);}
      }
    }
    setObjectsToPage(changesSheet,allHistory);
    setObjectsToPage(downloadsSheet,allDownloads);
  }
  setObjectsToPage(tempSheet,allStatus);
}

function runThirdView()
{
  var row = 3;
  var values1 = [];
  values1 = dataSheet.getRange("A" + (row+1) + ":C" + (row+1)).getValues();
  var wsk = values1[0][0];
  var dept = values1[0][1];
  var view = values1[0][2];
  var allStatus = [];
  var allHistory = [];
  var info = getAgentInfo(wsk,dept,view);
  allStatus = allStatus.concat(info["Statuses"]);
  allHistory = allHistory.concat(info["History"]);
  if (allHistory != undefined){
    var allDownloads = [];
    for (var k = 0; k < allHistory.length; k++){
      if(allHistory[k] != undefined){
      if (parseInt(allHistory[k]["Total Downloads"]) > 0){allDownloads.push(allHistory[k]);}
      }
    }
    setObjectsToPage(changesSheet,allHistory);
    setObjectsToPage(downloadsSheet,allDownloads);
  }
  setObjectsToPage(tempSheet,allStatus);
}

function runFourthView()
{
  var row = 4;
  var values1 = [];
  values1 = dataSheet.getRange("A" + (row+1) + ":C" + (row+1)).getValues();
  var wsk = values1[0][0];
  var dept = values1[0][1];
  var view = values1[0][2];
  var allStatus = [];
  var allHistory = [];
  var info = getAgentInfo(wsk,dept,view);
  allStatus = allStatus.concat(info["Statuses"]);
  allHistory = allHistory.concat(info["History"]);
  if (allHistory != undefined){
    var allDownloads = [];
    for (var k = 0; k < allHistory.length; k++){
      if(allHistory[k] != undefined){
      if (parseInt(allHistory[k]["Total Downloads"]) > 0){allDownloads.push(allHistory[k]);}
      }
    }
    setObjectsToPage(changesSheet,allHistory);
    setObjectsToPage(downloadsSheet,allDownloads);
  }
  setObjectsToPage(tempSheet,allStatus);
}

function runFifthView()
{
  var row = 5;
  var values1 = [];
  values1 = dataSheet.getRange("A" + (row+1) + ":C" + (row+1)).getValues();
  var wsk = values1[0][0];
  var dept = values1[0][1];
  var view = values1[0][2];
  var allStatus = [];
  var allHistory = [];
  var info = getAgentInfo(wsk,dept,view);
  allStatus = allStatus.concat(info["Statuses"]);
  allHistory = allHistory.concat(info["History"]);
  if (allHistory != undefined){
    var allDownloads = [];
    for (var k = 0; k < allHistory.length; k++){
      if(allHistory[k] != undefined){
      if (parseInt(allHistory[k]["Total Downloads"]) > 0){allDownloads.push(allHistory[k]);}
      }
    }
    setObjectsToPage(changesSheet,allHistory);
    setObjectsToPage(downloadsSheet,allDownloads);
  }
  setObjectsToPage(tempSheet,allStatus);
}

function clearSheet(name){//clears the results and borders on any given page
  var statusSheet = spreadsheet.getSheetByName(name);
  var lastRow = statusSheet.getLastRow();
  if (lastRow < 2) {return;}
  statusSheet.getRange(2, 1, lastRow - 1, statusSheet.getLastColumn() + 1).clearContent();
  statusSheet.getRange(2, 1, lastRow - 1, statusSheet.getLastColumn() + 1).setBorder(false,false,false,false,false,false);
}

function getAgentInfo(wsk,dept,view) {//takes department and WSK from runAccounts and gets the list of agents on the account, then the list of jobs from each active agent
  var statusList = [];//need to load to Agent Status page
  var trackHistoryList = [];//need to load to Tracking Changes and Downloads pages
  var getListURL = "https://api.mozenda.com/rest?WebServiceKey=" + wsk + "&Service=Mozenda10&Operation=View.GetItems&ViewID=" + view;
  var doTest = 0;
  do{
    var xml = UrlFetchApp.fetch(getListURL).getContentText();
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    doTest++;
  }while (doTest <=10 && !(root.getChild("Result") != null && root.getChild("Result").getValue() == "Success"));
  var namespace = root.getNamespace();
  var items = root.getChild("ItemList");
  var agents = items.getChildren("Item");
  for (var i = 0; i < agents.length; i++){
    var agentStatus = agents[i].getChild("Status").getValue();
    var agentID = agents[i].getChild("ItemID").getValue();
    var agentName = agents[i].getChild("Name").getValue();
    var configCheck = agents[i].getChild("Config").getValue();
    var lastRunTime = agents[i].getChild("LastRunTime").getValue();
    //create list of objects 
    if (agentStatus != "Ready"){statusList.push(getAgentStatus(wsk,agentID,dept,agentName,agentStatus));}//if the agent has active jobs
    if (configCheck.indexOf("TrackHistory") >= 0){trackHistoryList.push(trackHistory(wsk,agentID,dept,agentName,lastRunTime));}//if the agent has "Store Item History" selected
  }
  var obj = {"Statuses":statusList,"History":trackHistoryList}
  return obj;
}

function getLastColLetter(sheet)
{
  var lastNum = sheet.getLastColumn() - 1;
  return String.fromCharCode("A".charCodeAt(0) + lastNum);
}

function getHeader(lastColumn,sheet)
{
  var headerRange = "A1:" + lastColumn + "1";
  var result = sheet.getRange(headerRange).getValues()[0];
  return result;
}

function setObjectsToPage(sheet,objects)
{
  var lastColumn = getLastColLetter(sheet);
  var header = getHeader(lastColumn,sheet);
  var result = [];
  var undefinedTest = 0;
  for (var i = 0; i < objects.length; i++)
  {
    result[i] = [];
    if(objects[i] != undefined){
      for (var j = 0; j < header.length; j++)
      {
        if (objects[i].hasOwnProperty(header[j])){result[i][j] = objects[i][header[j]];}
        else {result[i][j] = "0";}
      }
    }
    else {undefinedTest++;}
  }
  //the new range part 1 number has to be the last row filled plus one
  //the new range part 2 number has to be the last row filled plus results.length
  //lastColumn stuff is still okay
  var rowToFill = sheet.getLastRow() + 1;
  if (result.length == 0){rowToFill = 1;}
  else
  {
  var range = "A" + rowToFill + ":" + lastColumn + (result.length + rowToFill - 1);
  //var range = "A2:" + lastColumn + (result.length + 1);
    sheet.getRange(range).setValues(result);
  }
}

function getAgentStatus(wsk,agentID,dept,agentName,agentStatus)
{
  var errored = 0;
  var paused = 0;
  var running = 0;
  var queued = 0;
  var postponed = 0;
  var getJobsURL = "https://api.mozenda.com/rest?WebServiceKey=" + wsk  + "&Service=Mozenda10&Operation=Agent.GetJobs&AgentID=" + agentID;
  var doTest = 0;
  var xml = null;
  var document = null;
  var root = null;
  do{
    try{
      xml = UrlFetchApp.fetch(getJobsURL).getContentText();
      document = XmlService.parse(xml);
      root = document.getRootElement();//root = ViewGetItemsResponse: need to go root/ItemList/Item for list of agents
    }
    catch(err)
    {
      Logger.log(err);
    }
    doTest++;

  }while (doTest <=10 && xml != null && document != null && root != null && !(root.getChild("Result") != null && root.getChild("Result").getValue() == "Success"));
  var namespace = root.getNamespace();
  var jobList = root.getChild("JobList");
  var jobsArray = [];
  jobsArray = jobList.getChildren("Job");
  //for each job in array, get Status and add increment
  var results = {};
  if (jobsArray != null){
    results["Department"] = dept;
    results["Agent ID"] = agentID;
    results["Agent Name"] = agentName;
    results["Agent Status"] = agentStatus;
    for(var i = 0; i < jobsArray.length; i++){
      var jobStatus = jobsArray[i].getChild("Status").getValue();
      var jobState = jobsArray[i].getChild("State").getValue();
      if (jobStatus == "Queued" && jobState == "Postponed") {jobStatus = "Postponed";}
      if (results.hasOwnProperty(jobStatus)) {results[jobStatus]++;}
      else{results[jobStatus] = 1;}
    }
  }
  return results;
}//new end of fxn

function trackHistory(wsk,agentID,dept,agentName,lastRunTime)
{
  var jobsURL = "https://api.mozenda.com/rest?WebServiceKey=" + wsk + "&Service=Mozenda10&Operation=Agent.GetJobs&AgentID=" + agentID + "&Job.State=Archived";
  var doTest = 0;
  var xml = null;
  var document = null;
  var root = null;
  do{
    try{
      xml = UrlFetchApp.fetch(jobsURL).getContentText();
      document = XmlService.parse(xml);
      root = document.getRootElement();//root = ViewGetItemsResponse: need to go root/ItemList/Item for list of agents
    }
    catch(err)
    {
      Logger.log(err);
    }
    doTest++;

  }while (doTest <=10 && xml != null && document != null && root != null && !(root.getChild("Result") != null && root.getChild("Result").getValue() == "Success"));
  var jList = root.getChild("JobList");
  var jobs = [];
  jobs = jList.getChildren("Job");
  var trackResults = {};
  trackResults["Department"] = dept;
  trackResults["Agent ID"] = agentID;
  trackResults["Agent Name"] = agentName;
  trackResults["Last Run Time"] = lastRunTime;
  if (jobs.length > 0){//if no jobs, takes us to the end of the fxn
  var created;// = jList[0].getChild("Created").getValue();
  var ended;
  if (jobs.length > 1){
    for ( var i = 0; i < jobs.length; i++){//to find the number of jobs to check
      created = new Date((jobs[i].getChild("Created").getValue()).replace(" ","T"));
      if (i+1 < jobs.length){ended = new Date((jobs[i+1].getChild("Ended").getValue()).replace(" ","T"));}
      if (ended.getTime() > created.getTime() && i+1 < jobs.length){created = new Date((jobs[i+1].getChild("Created").getValue()).replace(" ","T"));}
      else{
        var count = i+1;//number of jobs in the last batch - 1
        break;
      }
    }
    jobs.length = count;
  }//now we look at each job in "jobs" up to count (total jobs in last batch) and get the info we need
  var addedTotal = 0;
  var removedTotal = 0;
  var changedTotal = 0;
  var unchangedTotal = 0;
  var imagesTotal = 0;
  var filesTotal = 0;
  var totalFilesAdded = 0;
  var creditsUsed = 0;
  var jobIDs = [];
  for (var i = 0; i < jobs.length; i++) {jobIDs[i] = jobs[i].getChild("JobID").getValue();}
  var delim = jobIDs.join(",");
  var getAllJobs = "https://api.mozenda.com/rest?WebServiceKey=" + wsk + "&Service=Mozenda10&Operation=Job.Get&JobID=" + delim + "&IncludeJobStatistics=Yes";
  var doTest = 0;
  var xml = null;
  var document = null;
  var root = null;
  do{
    try{
      xml = UrlFetchApp.fetch(getAllJobs).getContentText();
      document = XmlService.parse(xml);
      root = document.getRootElement();//root = ViewGetItemsResponse: need to go root/ItemList/Item for list of agents
    }
    catch(err)
    {
      Logger.log(err);
    }
    doTest++;

  }while (doTest <=10 && xml != null && document != null && root != null && !(root.getChild("Result") != null && root.getChild("Result").getValue() == "Success"));
  var AllJobs = [];
  if (root.getChild("JobList") != null) {AllJobs = root.getChild("JobList").getChildren("Job");}
  else if (root.getChild("Job") != null) {AllJobs = [root.getChild("Job")];}
  for(var i = 0; i < AllJobs.length; i++){
    var statsList = AllJobs[i].getChild("JobStatisticCollection").getChildren("JobStatistic");
    for (var j = 0; j < statsList.length; j++){
      if (statsList[j].getChild("Group").getValue() == "Items"){
        //if added, removed, changed, refreshed(unchanged)
        if (statsList[j].getChild("Name").getValue() == "Added"){addedTotal += parseInt(statsList[j].getChild("Value").getValue());}
        else if (statsList[j].getChild("Name").getValue() == "Removed"){removedTotal += parseInt(statsList[j].getChild("Value").getValue());}
        else if (statsList[j].getChild("Name").getValue() == "Changed"){changedTotal += parseInt(statsList[j].getChild("Value").getValue());}
        else if (statsList[j].getChild("Name").getValue() == "Refreshed"){unchangedTotal += parseInt(statsList[j].getChild("Value").getValue());}
      }
      else if (statsList[j].getChild("Group").getValue() == "Downloaded"){
        if (statsList[j].getChild("Name").getValue() == "Images"){imagesTotal += parseInt(statsList[j].getChild("Value").getValue());}
        else if (statsList[j].getChild("Name").getValue() == "Files"){filesTotal += parseInt(statsList[j].getChild("Value").getValue());}
      }
      else if (statsList[j].getChild("Group").getValue() == "Files" && statsList[j].getChild("Name").getValue() == "Added"){totalFilesAdded += parseInt(statsList[j].getChild("Value").getValue());}
      else if (statsList[j].getChild("Group").getValue() == "CreditsUsed"){
        if (statsList[j].getChild("Name").getValue() == "WebPageCredits"){creditsUsed += parseInt(statsList[j].getChild("Value").getValue());}
        else if (statsList[j].getChild("Name").getValue() == "WebPagePremium"){creditsUsed += parseInt(statsList[j].getChild("Value").getValue());}
        else if (statsList[j].getChild("Name").getValue() == "FileDownloads"){creditsUsed += parseInt(statsList[j].getChild("Value").getValue());}
      }
    }//checks all stats
  }//for each job
  var totals = parseInt(imagesTotal) + parseInt(filesTotal);
  var dups = parseInt(imagesTotal) + parseInt(filesTotal) - parseInt(totalFilesAdded);
  trackResults["Items Added"] = addedTotal;
  trackResults["Items Removed"] = removedTotal;
  trackResults["Items Changed"] = changedTotal;
  trackResults["Items Unchanged"] = unchangedTotal; 
  trackResults["Credits Used"] = creditsUsed;
  trackResults["Images Downloaded"] = imagesTotal;
  trackResults["Files Downloaded"] = filesTotal;
  trackResults["Total Downloads"] = totals;
  trackResults["New Downloads"] = totalFilesAdded;
  trackResults["Duplicate Downloads"] = dups;
  }
  return trackResults;
}