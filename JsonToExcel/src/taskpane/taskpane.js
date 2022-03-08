/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      let textJson = document.getElementById("data-json").value;
      const dataJson = JSON.parse(textJson);
      let data = fillJsonObjectToData(dataJson);
      const range = context.workbook.getSelectedRange();
      // Read the range address
      range.load("address");
      let rangeSize = range.getResizedRange(data.length-1, data[0].length-1);
      let firstRow = rangeSize.getRow();
        // Update the fill color
      firstRow.format.fill.color = "#f4a460";
  
      rangeSize.values = data;
      await context.sync();
    });
   
  } catch (error) {
    console.error(error);
  }
}
//      let data = [ ["name","age","contact/address","contact/sdt"],["tuan",20,"bui thi xuan","0949238337"]];
    function fillJsonObjectToData(dataJson){
      let props =[];
      let resultData =[];
      let dictionary = {};
      dataJson.forEach((el,index) =>{
       
        //let dataProps =[];
        let nodeRoots =[];
        const valueIndex = {
           el: el,
           index: index
        }
           // dictionary:{"name":["rin","tuan"] , "age":[22,20]}
        mapKeyToValue(dictionary,dataJson.length,valueIndex,nodeRoots);
      //  resultData.push(dataProps);
      });
      dictionaryToStructureExcel(resultData,dictionary,props);
      resultData.unshift(props);
          return resultData;
    }
      //props[]:key, dataProps[]:values, el:{} , nameKey:format name key
    function mapKeyToValue(dictionary,length,valueIndex,nodeRoots){
      for( let key in valueIndex.el){
        const dataProps = createTemp(length);
        if(valueIndex.el.hasOwnProperty(key)){
              if(valueIndex.el[key] !== null && valueIndex.el[key].constructor === ({}).constructor){
                if(nodeRoots.length >0){
                   nodeRoots[0] = nodeRoots[0] +"/"+ key; // ["contact/address"]
                }else{
                 nodeRoots.push(key);// ["contact"];
                }
                let valueSub = {
                  el:valueIndex.el[key],
                  index: valueIndex.index
                }
                 mapKeyToValue(dictionary,dataProps,valueSub,nodeRoots);
                 nodeRoots.shift();
             }
             else{
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
        }
      }
    }

    function dictionaryToStructureExcel(data,dictionary,props){
      for( let key in dictionary){//dictionary:{"name":["rin,"trien"] , "age":[22,20]}
          if(props.indexOf(key) === -1){
            props.push(key);
          }
      }
      for(let i =0 ; i<props.length ;i++){
        const values = [];
        for(let j =0 ;j<props.length; j++){
          values.push(dictionary[props[j]][i]); // ["rin",22]
        }
        data.push(values);
      }
    }

    function createTemp(length){
      let temps = [];
        for(var i = 0; i < length; i++) {
            temps.push(null);
        }
        return temps;
    }

