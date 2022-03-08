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

      dataJson.forEach(el =>{
        let dataProps =[];
        let nodeRoots =[];
       mapKeyToValue(props,dataProps,el,nodeRoots);
        resultData.push(dataProps);
      });
      resultData.unshift(props);
          return resultData;
    }
      //props[]:key, dataProps[]:values, el:{} , nameKey:format name key
    function mapKeyToValue(props,dataProps,el,nodeRoots){
      for( let key in el){
        if(el.hasOwnProperty(key)){
              if(el[key] !== null && el[key].constructor === ({}).constructor){
                if(nodeRoots.length >0){
                   nodeRoots[0] = nodeRoots[0] +"/"+ key; // ["contact/address"]
                }else{
                 nodeRoots.push(key);// ["contact"];
                }
                 mapKeyToValue(props,dataProps,el[key],nodeRoots);
                 nodeRoots.shift();
             }
             else{
              if(props.indexOf(key) === -1){
                if(nodeRoots.length >0){
                  if(props.indexOf(nodeRoots[0]+ "/" + key)=== -1){
                    props.push(nodeRoots[0] + "/" + key);//contact/sdt
                  }
                }else{
                  props.push(key);
                }
            }        
              dataProps.push(el[key]);
             } 
        }
      }
    }

