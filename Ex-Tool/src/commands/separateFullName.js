// coding standard
// coding convention
// clean coding
// refactoring code legacy code
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let basePosition =0;

 async function getRange() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRanges();
      range.load("address");  
      await context.sync();
      const address = getAddress(range.address);
      getTextFromRange(address);
    });
  } catch (error) {
    console.log(error);
  }
}
 function getAddress(range) {
  const arrRange = range.split("!");
  if(arrRange[1].includes(":")){
  basePosition = arrRange[1].split(":")[0]
  }else{
    basePosition = arrRange[1];
  }
  console.log(basePosition);
  return arrRange[1];
}

 async function getTextFromRange(address) {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem("Sheet1");
      let range = sheet.getRange(address);
      range.load("text");
      await context.sync();
      let result = range.text;
      handleTextName(result);
    });
  } catch (error) {
    console.log(error);
  }
}


const handleTextName = (arrayName) => {
  let dataArr = [];
  arrayName.forEach((element) => seperateFullName(element[0].trim(), dataArr));
  rangeForData(dataArr);
};

const seperateFullName = (fullName, dataArr) => {
  let data = [];
  const errorName = ["",""];
  let arrLastName = fullName.split(" ");
  if (arrLastName.length >= 2) {
    for (let i = 0; i <= arrLastName.length; i++) {
      if (isValid(arrLastName[i])) {
        if (i == arrLastName.length - 1) {
          filterFullNameToRange(arrLastName, data, dataArr);
        }
      } else {
        dataArr.push(errorName);
      }
    }
  } else {
    dataArr.push(errorName);
  }
  console.log(dataArr);
  //rangeForData(dataArr);
};

function filterFullNameToRange(arrLastName, data, dataArr) {
  let firstName = arrLastName.splice(-1);
  data.push(arrLastName.join(" "), firstName[0]);
  dataArr.push(data);
}


function removeAscent(str) {
  if (str === null || str === undefined) return str;
  str = str.toLowerCase();
  str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
  str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
  str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
  str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
  str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
  str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
  str = str.replace(/đ/g, "d");
  return str;
}
function isValid(string) {
  var re = /^[a-zA-Z!@#\$%\^\&*\)\(+=._-]{2,}$/g; // regex here
  return re.test(removeAscent(string))
}

async function rangeForData(valuesRange) {
  try {
    await new Promise((resolve, reject) => {
      Office.context.document.bindings.addFromPromptAsync(
        Office.BindingType.Matrix,
        { id: "currencyRange", promptText: "Select where to display the data" },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject();
          }
        }
      );
    });
    await Excel.run(async (context) => {
      let binding = context.workbook.bindings.getItem("currencyRange");
      let range = binding.getRange();
      console.log(valuesRange);
      range.load("address");
      
      let resizeRange = range.getResizedRange(valuesRange.length - 1, valuesRange[0].length - 1);
       const positionInvalids = getPositionInvalid(valuesRange);
       const arrFormatPositionInvalids = handlingPositionAddressInvalid(positionInvalids);
      await filterPositionInvalidForNotify(arrFormatPositionInvalids);
      resizeRange.getCell().format.horizontalAlignment = Excel.HorizontalAlignment.center;
      resizeRange.values = valuesRange;
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}

function getPositionInvalid(valueRange){

  let result =[];
  valueRange.forEach((e,index)=>{
      // e = ["abc",""] => sai
      // e[0] là họ
      // e[1] là tên
      if(e[1] ===""){
        result.push(index)
      } 
  })

  return result;
}

function handlingPositionAddressInvalid(positionInvalids) {
  let arrFormatPositionInvalids = [];
  if (positionInvalids.length > 0) {
    positionInvalids.forEach((e) => {
      let positionInvalid = basePosition.substring(1);
      let stringPosition = basePosition.substring(0, 1);
      let totalPositionInvalid = parseInt(positionInvalid) + e;
      let result = `${stringPosition}${totalPositionInvalid}`;
      arrFormatPositionInvalids.push(result);
    });
  }
  return arrFormatPositionInvalids;
}

async function notifyByPositionInvalid(address) {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem("Sheet1");
      let range = sheet.getRange(address);
      range.load("address");
        range.dataValidation.prompt = {
          message: "Họ và tên không hợp lệ",
          showPrompt: true, // The default is 'false'.
          title: "Sai định dạng",
        };
      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}
 function filterPositionInvalidForNotify(arrFormatPositionInvalids){
  arrFormatPositionInvalids.forEach(async (e) => await notifyByPositionInvalid(e));
}