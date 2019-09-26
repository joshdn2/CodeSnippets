/*
Josh D
This is a question I was asked at a career fair: 
Given a string of digits, how can you find the longest consecutive repeating number? 
8 minute time limit. 
This is what I came up with in 8 minutes and it works

*/
var arrayString = "122333444455";
if (arrayString){
    var array = arrayString.split("");
    var highestCount = 0;
    var mostRepeats;
    var tempRepeatsCheck;
    var count = 1;
    for (var i = 0; i < array.length; i++){
        //for each one, check the next one to see if they match
        if(array[i] == array[i+1]){//if they match, count++
            count++;
        }
        else {//if array[i] doesn't match the next value
            if(count > highestCount){
                mostRepeats = array[i];
                highestCount = count;
            }
            count = 1;
        }
    }
    console.log(mostRepeats);
}
else {console.log("No string");}