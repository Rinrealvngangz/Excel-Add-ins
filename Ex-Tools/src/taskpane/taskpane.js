const {exception} = require('../validates/exception');
let IS_SUB_ARRAY = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = test;
    document.getElementById("clear").onclick = clear;
  }
});

async function test(){
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    let productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    await context.sync();
});
}

export async function run() {
      let textJson = document.getElementById("data-json").value;
  try {
     await Excel.run(async (context) => {
      let dataJson = JSON.parse(textJson);
      document.getElementById("valid-json").style.display = "none";
     if(!Array.isArray(dataJson)){
       dataJson = [dataJson].flat();//convert {} -> []
     }
      let data = fillJsonObjectToData(dataJson);
      const range = context.workbook.getSelectedRange();
      // Read the range address
      range.load("address");
      let rangeSize = range.getResizedRange(data.length-1, data[0].length-1);
        // Update the fill color first row
      let firstRow = rangeSize.getRow();
      firstRow.format.fill.color = "#f4a460";
      firstRow.format.autofitRows();
      rangeSize.values = data;
      await context.sync();
    });
   
  }catch(error){
    exception(error,textJson);
  }
}
//  Structure so that display table Excel
//  let data = [["name","age","contact/address","contact/sdt"],["tuan",20,"bui thi xuan","0949238337"]];
    function fillJsonObjectToData(dataJson){
        let props =[];
        let resultData =[];
        let dictionary = {};
        dataJson.forEach((el,index) =>{
        let nodeRoots =[];//only one element
        const valueIndex = {
           el: el,
           index: index
        }// dictionary:{"name":["rin","tuan"] , "age":[22,20]}
        mapKeyToValue(dictionary,dataJson.length,valueIndex,nodeRoots);
      });
        dictionaryToStructureExcel(resultData,dictionary,props);
        resultData.unshift(props);
        return resultData;
    }
      //props[]:key, dataProps[]:values, el:{} ,nodeRoots:save key of sub object so that check isExist
    function mapKeyToValue(dictionary,length,valueIndex,nodeRoots){
      for( let key in valueIndex.el){
        const dataProps = createTemp(length);
        if(valueIndex.el.hasOwnProperty(key)){
              //check object sub
              if(valueIndex.el[key] !== null && valueIndex.el[key].constructor === ({}).constructor){
                createSubObjectJson(dictionary,length,valueIndex,nodeRoots,key);
              }
             else if(Array.isArray(valueIndex.el[key])){      
              createSubJsonArray(dictionary,length,valueIndex,nodeRoots,key);
              }
             else{
              createDictionary(dictionary,valueIndex,nodeRoots,dataProps,key);
             } 
        }
      }
    }
    function dictionaryToStructureExcel(data,dictionary,props){
      for( let key in dictionary){
          if(props.indexOf(key) === -1){
            props.push(key);
          }
      }
      for(let i =0 ; i<props.length ;i++){
        const values = [];
        for(let j =0 ;j<props.length; j++){
          values.push(dictionary[props[j]][i]);
        }
        data.push(values);
      }
    }
    //[null,null,null,....n]
    function createTemp(length){
      let temps = [];
        for(var i = 0; i < length; i++) {
            temps.push(null);
        }
        return temps;
    }

    function clear(){
      document.getElementById("data-json").value = "";
      document.getElementById("valid-json").style.display = "none";
    }

    function createDictionary(dictionary,valueIndex,nodeRoots,dataProps,key){
   //check is exist key dictionary by key sub(save nodeRoots) 
      if(nodeRoots.length >0){
        if(!dictionary.hasOwnProperty(nodeRoots[0]+ "/" + key)){
        dictionary[nodeRoots[0]+ "/" + key] = dataProps;
      }
      dictionary[nodeRoots[0] + "/" + key][valueIndex.index] = valueIndex.el[key];
      }
      else{
        if(!dictionary.hasOwnProperty(key)){
            dictionary[key] = dataProps;//{"name":[null,null,...n]}
        }
        dictionary[key][valueIndex.index] = valueIndex.el[key];// replace null to value of key
      }
    }

    function createSubJsonArray(dictionary,length,valueIndex,nodeRoots,key){
      IS_SUB_ARRAY = true;
      Object.assign({}, valueIndex.el[key]);
      nodeRoots[0] = key;
      let valueSub = {
        el:valueIndex.el[key],
        index: valueIndex.index
      }
      mapKeyToValue(dictionary,length,valueSub,nodeRoots);
      nodeRoots.shift();
      IS_SUB_ARRAY =false;
    }

    function createSubObjectJson(dictionary,length,valueIndex,nodeRoots,key){
      if(nodeRoots.length >0){
        nodeRoots[0] = nodeRoots[0] +"/"+ key; // ["contact/address"]
     }else{
      nodeRoots[0] =key;// ["contact"];
     }
     let valueSub = {
       el:valueIndex.el[key],
       index: valueIndex.index
     }
      mapKeyToValue(dictionary,length,valueSub,nodeRoots);
      if(IS_SUB_ARRAY){
       let rootArray = nodeRoots[0].split('/')[0];
           nodeRoots[0] = rootArray;                 
     }else{
       nodeRoots.shift();
     }
    }