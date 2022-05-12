
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      document.getElementById("btnFillColor").onclick = run;
    }
  });
let colorHexStar;
let colorHexEnd;

 async function run() {
try {
   await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address","rowCount", "columnCount", "cellCount"]); 
    // Read the range address
    await context.sync();
    const propertiesToGet = range.getCellProperties({ address: true }); 
    await context.sync(); 
    var arrAddress = []; 
  for (let iRow = 0; iRow < range.rowCount; iRow++){
       for (let iCol = 0; iCol < range.columnCount; iCol++){
            const cellAddress = propertiesToGet.value[iRow][iCol]; 
        arrAddress.push(cellAddress.address.slice(cellAddress.address.lastIndexOf("!") + 1));
        } 
    }  
    const rowCount = range.rowCount;
    const scaleColors = chroma.scale([colorHexStar,colorHexEnd])
                        .mode('lch').colors(rowCount);
    const [worksheets,...address]  = range.address.split('!');
    if(range.columnCount >1){
       const filterNumberAddress = groupNumberAddress(arrAddress);
       const rowAddress = GetRowAddress(filterNumberAddress);
       write(rowAddress.length);
       rowAddress.forEach((row,i)=>{
         write(row);
            fillColorCellByAddress(row,worksheets,scaleColors[i]);     
        })
    }else{
        arrAddress.forEach(async(el,i) =>{
           await fillColorCellByAddress(el,worksheets,scaleColors[i]);
        })
    }
    });
  
}catch(error){
  exception(error,textJson);
}
}
    async function fillColorCellByAddress(address,worksheets,color){
       await Excel.run(async (context)=>{
            let sheet = context.workbook.worksheets.getItem(worksheets);//sheet1
            let range = sheet.getRange(address);//A1:B5
            range.load("address");
            range.format.fill.color = color;
            await context.sync();
    })
    }
    $("#mycolorStart").colorpicker({
            defaultPalette: 'web',
            history: false
    });
    
    $("#mycolorEnd").colorpicker({
            defaultPalette: 'web',
            history: false
    });
    $("#mycolorStart").on("change.color", function(event, color){
        colorHexStar =color;
    });
    
    $("#mycolorEnd").on("change.color", function(event, color){
        colorHexEnd =color;
    });
    
const write = (message) =>{
    $("#testData1").append(message);
}

function groupNumberAddress(array){
   const grouped = array.reduce((r, v, i, a) => {
        let item = a[i - 1];
        if(item != undefined){
           item = parseInt(item.match(/\d+/)[0]);
        }
        if (parseInt(v.match(/\d+/)[0]) === item ) {
            r[r.length - 1].push(v);
        } else {
            r.push(parseInt(v.match(/\d+/)[0]) === parseInt(a[i + 1].match(/\d+/)[0]) ? [v] : v);
        }
        return r;
    }, []);
    return grouped;
}

function GetRowAddress(array){
  const rowAddress =[];
   array.forEach(el=>{
     let n = el.length;
     let cell = el[0]+":"+el[n-1];//A1:C1
     rowAddress.push(cell);
   })
  
   return rowAddress;
}