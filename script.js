
function sumofNumber(){
    const num1 = parseInt(document.getElementById("num1").value);
    const num2 = parseInt(document.getElementById("num2").value);
    var result = num1+num2;
    console.log(result);
    document.getElementById("result").innerText = result;
}
