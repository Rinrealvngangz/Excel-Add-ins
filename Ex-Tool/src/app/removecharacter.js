
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      document.getElementById("btnRemove").onclick = removeCharacterUnwanted;
    }
  });
  
  let addressGlobal = ""
  async function removeCharacterUnwanted(){
  let characterUnwanted = document.getElementById("inputCharacter").value;
    if(characterUnwanted===""){
         document.getElementById("notifyInvalid").hidden = false;
    }else{
             document.getElementById("notifyInvalid").hidden = true;
  
        let arrStrCharacter = [];
        let arrayDataSelecteds = await ReturnArrayDataFromCells();
        arrayDataSelecteds.map((element) => {
          element.map((strCharacter) => {
            strCharacter.includes(characterUnwanted)
              ? (strCharacter = recursionString(characterUnwanted, strCharacter))
              : strCharacter;
            arrStrCharacter.push([strCharacter]);
          });
        });
        await fillData(arrStrCharacter);
  
    }
  
  }
  
  function characterWanted(characterUnwanted, strCharacter) {
    strCharacter= strCharacter.replace(characterUnwanted,'');
    strCharacter = recursionString(characterUnwanted, strCharacter)
    return strCharacter;
  }
  function recursionString(characterUnwanted, strCharacter){
    strCharacter= strCharacter.includes(characterUnwanted) 
    ? strCharacter= characterWanted(characterUnwanted, strCharacter) 
     : strCharacter;
      return strCharacter;  
  }
  
  async function ReturnArrayDataFromCells() {
    try {
      const result= await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRanges();
        range.load("address");
        await context.sync();
      var address = range.address;
      var addressDetail = filterAddress(address);
      let  arrContentFromAddressDetail= await getContentInAddress(addressDetail);
      return arrContentFromAddressDetail;  
      });
      return result;
    } catch (error) {
      console.log(error);
    }
  }
  
  
  // range.address => sheet!address
  // filterAdress fuction return address
  function filterAddress(address){
   const arrRange = address.split("!");
     addressGlobal = arrRange[1];
   return arrRange[1];
  }
  
  async function getContentInAddress(addressDetail) {
      try {
       const result= await Excel.run(async (context) => {
          let sheet = context.workbook.worksheets.getItem("Sheet1");
          let range = sheet.getRange(addressDetail);
          range.load("text");
          await context.sync();
          let result = range.text;
          return result 
        });
          return result;
      } catch (error) {
        console.log(error);
      }
  }
  async function fillData(valuesRange) {
    try {
      await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem("Sheet1");
        arrCellSelected = addressGlobal.split(':');
        arrCellSelected.length > 1 ? addressGlobal= arrCellSelected[0] : addressGlobal;
        let range = sheet.getRange(addressGlobal);
        console.log(addressGlobal);
        console.log(valuesRange);
        range.load("address");
        let resizeRange = range.getResizedRange(valuesRange.length - 1, valuesRange[0].length - 1);
        // resizeRange.getCell().format.horizontalAlignment = Excel.HorizontalAlignment.center;
        // resizeRange.values = [["Nguyễn văn èo  "], ["Phạm Anh quốc  "]];
       resizeRange.values = valuesRange;
        await context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }
  
  async function rangeForData(valuesRange) {
    try {
      await new Promise((resolve, reject) => {
        Office.context.document.bindings.addFromPromptAsync(
          Office.BindingType.Matrix,
          { id: "removeCharacterUnwanted", promptText: "Select where to display the data" },
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
        let binding = context.workbook.bindings.getItem("removeCharacterUnwanted");
        let range = binding.getRange();
        console.log(valuesRange);
        range.load("address");
        let resizeRange = range.getResizedRange(valuesRange.length - 1, valuesRange[0].length - 1);
        resizeRange.values = valuesRange;
        await context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }
  
  
  
  