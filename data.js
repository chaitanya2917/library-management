export let s;
s=55;
console.log(s);
console.log('hey beech');
window.height = 10;
document.getElementById('submitData').addEventListener('click', (e) => {
    var heightInput = document.getElementById("userId");
    var weightInput = document.getElementById("userPassword");
    window.height = heightInput.value; // Assign the height value to the global variable
    
    var weight = weightInput.value;
    if (window.height && weight) {
        var userHeight = parseFloat(window.height); 
        var userWeight = parseFloat(weight); 
        console.log("Height: " + userHeight + " cm");
        console.log("Weight: " + userWeight + " kg");
        heightInput.value = '';
        weightInput.value = '';
    } else {
        alert("Please fill in both height and weight fields.");
    }
});

console.log(window.height);
console.log("print")
