
  let rate =0;
  const API_KEY ="a5a4757e1c-26fa375a73-rblve6";
  const sympolCurrency={
    USD:"$",
    JPY:"¥",
    AUD:"A$",
    VND:"đ",
    EUR:"€",
    GBP:"£",
    CAD:"C$",
    CHF:"CHf",
    CNH:"CN¥",
    HKD:"HK$",
    NZD:"$"

  }
  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      document.getElementById("btn-exec").onclick = currencyExchange;
    }
  });

  async function currencyExchange(){
      let from = document.getElementById("frmcurrency").value;//VND
      let to   = document.getElementById("tocurrency").value;//USD
      let valuesFromRange =  getSelectedRange().then(async (valuesRange) =>{
      let values = await getValuesRange(valuesRange);  //sheet1!A1:B5
       return new Promise((resolve ,reject)=>{ resolve(values)})
       });    
       valuesFromRange.then((val) => {
       processRange(from,to,val.values).then(async (valuesExchange) =>{
        await setValuesRange(valuesExchange,to);
       })  
      })
  }

  async function getSelectedRange(){
   return await Excel.run(async (context) => {
      let range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      return range.address;//sheet1!A1:B5
  });
  }
  //sheet1!A1:B5
  async function getValuesRange(range){
    return await Excel.run(async (context)=>{
      const [worksheets,...address]  = range.split('!');//
      let sheet = context.workbook.worksheets.getItem(worksheets);//sheet1
      let valuesRange = sheet.getRange(address[0]);//A1:B5
      valuesRange.load("values");
      await context.sync();
      return valuesRange;// obj:{ values:[[200,300] ,[200],[200]]}
    })
  }

  async function setValuesRange(valuesRange,symbol){
   // write(JSON.stringify(valuesRange));
    await rangeForData(valuesRange);
  }
  
  function formatValueInput(input,toExchange){//2000
    let output;
     if(typeof input == "number"){
       output = convert(input);
     }else{
       let numberArr = input.replace(/\D+/g, ' ').trim().split(' ');
       output = hasNumber(input) && numberArr.length == 1 ?
       convert(parseInt(numberArr[0])) : input;
     }
     if(output !=""){
      output+= " "+sympolCurrency[toExchange];
     }
      
     return output;
  }

  function hasNumber(numberStr){
    return /\d/.test(numberStr);
  }

  function processRange(fromExchange,toExchange, arrCurrency){
      return new Promise((resolve,reject)=>{
        getRate(fromExchange,toExchange,1).then(()=>{
          let result = arrCurrency.map(el => el = processCurrency(el,toExchange));//[200]
          resolve(result); 
        }).catch(err => reject(err.error))
      }) 
  }

  function processCurrency(itemsFromRange,toExchange){
      return itemsFromRange.map(el => el = formatValueInput(el,toExchange));      
  }

  async function rangeForData(valuesRange) {
    try {
        await new Promise((resolve, reject) => {
            Office.context.document.bindings.addFromPromptAsync(
                Office.BindingType.Matrix,
                { id: "currencyRange",promptText:"Select where to display the data" },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                    } else {
                        reject();
                    }
                }
            )
        })
        await Excel.run(async (context) => {
            let binding = context.workbook.bindings.getItem("currencyRange");
            let range = binding.getRange();
            range.load("address");
          //  write(JSON.stringify(valuesRange));
            let resizeRange =  range.getResizedRange(valuesRange.length-1, valuesRange[0].length-1);
           // resizeRange.getCell(0,0).format.horizontalAlignment = Excel.HorizontalAlignment.center;
            resizeRange.values = valuesRange;
            await context.sync();
        });
    }
    catch (error) {
      write(error);
    }
}

  function getRate(from,to,amount){
    try{
      const options = {method: 'GET', headers: {Accept: 'application/json'}};
      const currencyExchange = fetch(`https://api.fastforex.io/convert?from=${from}&to=${to}&amount=${amount}&api_key=${API_KEY}`, options)
            .then(response => response.json())
            .then(response =>{return response})
            .catch(err =>  write(err.error));
    return currencyExchange.then(val => {
      rate =  val.result.rate;
    })
    }catch(err){
      write(err.error);
  }}

  function convert(amount){
    return rate * amount;
  }
  function write(message){
    document.getElementById('testData').innerText += message;
}