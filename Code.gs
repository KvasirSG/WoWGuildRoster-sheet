function myFunction() {
  var spreadsheet = SpreadsheetApp.openById("111YiB6abs6jXGn2YmFLaw6HxKhfriyvn-GquRkATiuc");
  var sheet = SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]);
  var data = sheet.getDataRange().getValues();
  var realm_name = "REALM";
  var API_KEY = "Blizzard API KEY";
  for (var i = 1; i < data.length; i++) {
    var bob = i + 1;
    
 
    try {
      SpreadsheetApp.getActiveSheet().getRange('B'+bob).setValue(thistoonclass(data[i][0],realm_name,API_KEY));
      SpreadsheetApp.getActiveSheet().getRange('C'+bob).setValue(thistoonspec(data[i][0],realm_name,API_KEY));
      SpreadsheetApp.getActiveSheet().getRange('E'+bob).setValue(thistoonspecrole(data[i][0],realm_name,API_KEY));
      SpreadsheetApp.getActiveSheet().getRange('G'+bob).setValue(ilvl(data[i][0],realm_name,API_KEY));
      SpreadsheetApp.getActiveSheet().getRange('H'+bob).setValue(ilvleqpt(data[i][0],realm_name,API_KEY));
      var myhyperlink = "https://worldofwarcraft.com/en-gb/character/"+realm_name+"/"+thistoon(data[i][0],realm_name,API_KEY);
      SpreadsheetApp.getActiveSheet().getRange('J'+bob).setValue('=HYPERLINK("'+myhyperlink+'")');
      SpreadsheetApp.getActiveSheet().getRange('I'+bob).setValue(thistoonrank(data[i][0],realm_name,API_KEY));
      
    } catch(e) {
      Logger.log("Couldn't find data from: "+ data[i][0]);
    }
  
  } 
}
function ilvleqpt(toonName,realmName,API_KEY) {
//function ilvl(toonName,realmName) {

// Character information
var toonJSON = UrlFetchApp.fetch("https://eu.api.battle.net/wow/character/"+realmName+"/"+toonName+"?fields=items&locale=en_GB&apikey=" + API_KEY);
var toon = Utilities.jsonParse(toonJSON.getContentText());
  Utilities.sleep(500);

// Collated information we're going to output... 
var toonIlvl = new Array(
""+toon.items.averageItemLevelEquipped
)

var sheet = SpreadsheetApp.getActiveSheet();
Logger.log(toonIlvl);
return toonIlvl;
}

function ilvl(toonName,realmName,API_KEY) {
//function ilvl(toonName,realmName) {

// Character information
var toonJSON = UrlFetchApp.fetch("https://eu.api.battle.net/wow/character/"+realmName+"/"+toonName+"?fields=items&locale=en_GB&apikey="+API_KEY);
var toon = Utilities.jsonParse(toonJSON.getContentText());
   Utilities.sleep(500);

// Collated information we're going to output... 
var toonIlvl = new Array(
""+toon.items.averageItemLevel
)

var sheet = SpreadsheetApp.getActiveSheet();
Logger.log(toonIlvl);
return toonIlvl;
}

function thistoon(toonName,realmName,API_KEY) {
//function ilvl(toonName,realmName) {

// Character information
var toonJSON = UrlFetchApp.fetch("https://eu.api.battle.net/wow/character/"+realmName+"/"+toonName+"?locale=en_GB&apikey="+API_KEY);
var toon = Utilities.jsonParse(toonJSON.getContentText());
   Utilities.sleep(500);

// Collated information we're going to output... 
var toonIlvl = new Array(
""+toon.name
)

var sheet = SpreadsheetApp.getActiveSheet();
Logger.log(toonIlvl);
return toonIlvl;
}

function thistoonclass(toonName,realmName,API_KEY) {
//function ilvl(toonName,realmName) {

// Character information
var toonJSON = UrlFetchApp.fetch("https://eu.api.battle.net/wow/character/"+realmName+"/"+toonName+"?locale=en_GB&apikey="+API_KEY);
var id_to_class = {
	1 : "Warrior",
	2 : "Paladin",
	3 : "Hunter",
	4 : "Rogue",
	5 : "Priest",
	6 : "Death Knight",
	7 : "Shaman",
	8 : "Mage",
	9 : "Warlock",
	10 : "Monk",
	11 : "Druid",
    12 : "Demon Hunter", 
};
var toon = Utilities.jsonParse(toonJSON.getContentText());
   Utilities.sleep(500);

// Collated information we're going to output... 
var toonIlvl = new Array(
""+toon.class
)

var sheet = SpreadsheetApp.getActiveSheet();
Logger.log(toonIlvl);
return  id_to_class[toonIlvl];
}

function thistoonspec(toonName,realmName,API_KEY) {
//function ilvl(toonName,realmName) {

// Character information
var toonJSON = UrlFetchApp.fetch("https://eu.api.battle.net/wow/character/"+realmName+"/"+toonName+"?fields=talents&locale=en_GB&apikey="+API_KEY);
var toon = Utilities.jsonParse(toonJSON.getContentText());
   Utilities.sleep(500);

// Collated information we're going to output... 
var pritalent = toon.talents[0].spec.name;

var sheet = SpreadsheetApp.getActiveSheet();
Logger.log(pritalent);
return pritalent;
}
function thistoonspecrole(toonName,realmName,API_KEY) {
//function ilvl(toonName,realmName) {

// Character information
var toonJSON = UrlFetchApp.fetch("https://eu.api.battle.net/wow/character/"+realmName+"/"+toonName+"?fields=talents&locale=en_GB&apikey="+API_KEY);
var toon = Utilities.jsonParse(toonJSON.getContentText());
   Utilities.sleep(500);

// Collated information we're going to output... 
var pritalent = toon.talents[0].spec.role;

var sheet = SpreadsheetApp.getActiveSheet();
Logger.log(pritalent);
return pritalent;
}
function thistoonrank(toonName,realm_name,API_KEY) {
//function ilvl(toonName,realmName) {

// Character information
var guildJSON = UrlFetchApp.fetch("https://eu.api.battle.net/wow/guild/"+realm_name+"/guild%20name?fields=members&locale=en_GB&apikey="+API_KEY);
var guild = Utilities.jsonParse(guildJSON.getContentText());
   Utilities.sleep(500);

// Collated information we're going to output... 
// var memberrank = ;
var enteries = guild.members;
var ranks = ["Guild Master","Officer Council","Raid Leader","Raider","Pvper","Member","Social","Trial","Alt",];
  var members = {};
for (var i = 0; i < enteries.length; i++) {
  //Logger.log(enteries[i].character.name);
  members[enteries[i].character.name] = ranks[enteries[i].rank];
  //Logger.log(enteries[i].rank);
}
Logger.log(members[toonName]);
var sheet = SpreadsheetApp.getActiveSheet();
return members[toonName];
}
