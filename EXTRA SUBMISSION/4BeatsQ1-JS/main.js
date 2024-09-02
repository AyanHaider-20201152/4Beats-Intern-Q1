const XLSX = require('xlsx');
const {By, Builder, until, Key, WebDriverWait} = require('selenium-webdriver');

const date = new Date();
const weekday = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
var current_day = weekday[date.getDay()]; 


console.log("Current Day: "+current_day);

let filename = "4BeatsQ1.xlsx";
var workbook = XLSX.readFile(filename, {cellStyles:true});

var worksheet = workbook.Sheets[current_day];

// numberKeywords needs to be same as number ok keywords present in Excel Sheet
let numberKeywords = 10;
keyword_iterate(numberKeywords);



async function keyword_iterate(numKeywords) {
    for (let i = 3; i < numKeywords+3; i++) {
        var keyword = worksheet["C" + String(i)]['v'];
        console.log();
        console.log("Searching for keyword: "+keyword);
    
        let driver = new Builder().forBrowser('chrome').build();
        await search(driver, keyword, String(i)).then();
    }
}


async function search(driver, keyword, row) {

    await driver.get('https://www.google.com/');
    await driver.findElement(By.name("q")).sendKeys(keyword);

    // .click() chains are required to ensure that the page fully loads between operations
    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();

    let options = await driver.findElements(By.className("lnnVSe"));

    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();
    

    
    var longest;
    var len_longest;
    var shortest;
    var len_shortest;

    await options[0].getAttribute("aria-label").then(text =>{ 
        longest = text;
        len_longest = text.length;
        shortest = text;
        len_shortest = text.length;
    });
    
    for (let i = 1; i < 10; i++) {
        await options[i].getAttribute("aria-label").then(text =>{ 

            // If there are multiple options with the same length, they will ALL be saved
            if (text.length > len_longest){
                longest = text;
                len_longest = text.length;
            }
                
            else if (text.length == longest.length)
                longest += ", " + text;
    
            else if (text.length < len_shortest){
                shortest = text;
                len_shortest = text.length;
            }
    
            else if (text.length == shortest.length){
                shortest += ", " + text;
            }
        });
    }

    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();
    await driver.findElement(By.name("q")).click();
    
    await driver.quit();

    console.log("Longest option: "+longest);
    console.log("Shortest option: "+shortest);

    XLSX.utils.sheet_add_aoa(worksheet, [[longest]], { origin: ("D" + row) });
    XLSX.utils.sheet_add_aoa(worksheet, [[shortest]], { origin: ("E" + row) });
    XLSX.writeFileXLSX(workbook, filename);
}



    
   
   
    