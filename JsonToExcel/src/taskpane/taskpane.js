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
      let data = fillJsonToData(dataJson);
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

  // let data = [ ["name","age"],["tuan",20],["rin",10]];
    function fillJsonToData(dataJson){
      let props =[];
      let resultData =[];
      dataJson.forEach(el =>{
        let dataProps =[];
       
        for( let key in el){
              if(el.hasOwnProperty(key)){
                  if(props.indexOf(key) === -1 ){
                    props.push(key);
                  }
                  dataProps.push(el[key]);
              }
        }
        resultData.push(dataProps);
      });
      resultData.unshift(props);
          return resultData;
    }

