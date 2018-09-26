console.time("save");
var fs=require('fs');
var capitalize=require('capitalize');
var json2xls = require('json2xls');
const Json2csvParser = require('json2csv').Parser;
const XLSX = require('xlsx');
var myData=require("./countries").data;
var cities=require("./cities").cities;
var countriesArray=Object.keys(myData);
//var citiesArray=Object.values(myData.data);

//Comment below two lines and Uncomment above and below to take large db city
var citiesFlatten=[];cities.forEach((cities)=>citiesFlatten.push(cities.city));

//var citiesFlatten = [].concat.apply([], citiesArray);
var countryCities=countriesArray.concat(citiesFlatten);
var pattern = new RegExp("\\b("+countryCities.join("|")+")\\b","ig");

//XL Read
console.log("Reading XL File...")
var workbook = XLSX.readFile('abcnews-date-text.xlsx');// ./assets is where your relative path directory where excel file is, if your excuting js file and excel file in same directory just igore that part
var sheet_name_list = workbook.SheetNames; // SheetNames is an ordered list of the sheets in the workbook
var data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); //if you have multiple sheets
var json=[];
var jsonObject={};
console.log("Checking Countries and Cities...");
  for(var key in data){
    var str=String(data[key]['headline_text']);
    var matchArray = str.match(pattern);
    if(!!matchArray){
        matchArray.forEach(function(item) {
            str=str.replace(item,capitalize.words(item));
        });
        jsonObject={
            publish_date:data[key]['publish_date'],
            headline_text:str.replace(/^\w/, c => c.toUpperCase())
        }
        json.push(jsonObject);
    }
  };
  console.log("Countries and Cities Capitalaized");
  console.log("Writing data to Excel...");
  var xls = json2xls(json);
  console.log("Converted to XL");
//   Uncomment Below lines to generate CSV format
//   const fields = ['publish_date','headline_text'];
//   const json2csvParser = new Json2csvParser({ fields });
//   const csv = json2csvParser.parse(json);
//   fs.writeFileSync('data.csv', csv, 'binary');
  fs.writeFileSync('output.xlsx', xls, 'binary');
  console.log("Success!!!");
  console.timeEnd("save");