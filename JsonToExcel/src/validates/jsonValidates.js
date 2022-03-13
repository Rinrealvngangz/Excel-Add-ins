const textArea = document.querySelector('textarea');
textArea.addEventListener("input",bindTextArea);
function bindTextArea(e){
  if(e.target.value == ""){
      document.getElementById("valid-json").style.display = "none";
  }else{
      let txtJson =e.target.value;
      try{
          JSON.parse(txtJson);
          document.getElementById("valid-json").style.display = "none";
      }catch(error){
          document.getElementById("valid-json").style.display = "block";      
          const [first, ...rest] = `${error}`.split(':');
          const formatError  = rest.join(':');
          document.getElementById("text-error").innerHTML = formatError;      
      }
  }
}