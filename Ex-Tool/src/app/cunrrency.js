/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
  async function currency(event){
   await getSelectedRange().then(async (valuesRange) =>{
    await getValuesRange(valuesRange);
   });
    console.log(event.source);
    event.completed();
  }

  async function getSelectedRange(){
   return await Excel.run(async (context) => {
      let range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      return range.address;
  });
  }

  async function getValuesRange(range){
    await Excel.run(async (context)=>{
      const [worksheets,...address]  = range.split('!');
      let sheet = context.workbook.worksheets.getItem(worksheets);
      let valuesRange = sheet.getRange(address[0]);
      valuesRange.load("values");
      await context.sync();
      console.log(JSON.stringify(valuesRange.values));
    })
  }
