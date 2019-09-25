var fulltext = `%FullText%`;

var keywords = [
    "data",
    "service",
    "saas",
    "revenue",
    "web"
];

var row = {};

for (var i = 0; i < keywords.length; i++){
    var inputReg = "\\b" + keywords[i] + "\\b";
    var myRegex = new RegExp(inputReg,"gi");
    var count = (fulltext.match(myRegex) || []).length;
    row[keywords[i]] = count;
}

console.log(row);//to console
M_SetFieldValues(row, true);//Mozenda custom function