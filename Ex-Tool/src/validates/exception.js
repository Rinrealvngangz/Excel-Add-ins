
export function exception(error,textJson){
    document.getElementById("valid-json").style.display = "block";
    if(textJson == ""){
      document.getElementById("text-error").innerHTML ="Require add Json"
    }else{
       const [first, ...rest] = `${error}`.split(':');
       const formatError  = rest.join(':');
      document.getElementById("text-error").innerHTML = formatError;
    }
}