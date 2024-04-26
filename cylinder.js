import { s} from './data.js';
console.log(window.height);
console.log(height)
console.log("hey bro")
function fetchFillLevel() {
    return height; 
}

function updateCylinderFillLevel() {
    const fillLevelElement = document.getElementById("fill-level");
    const inputFillLevel = fetchFillLevel();
    const validatedFillLevel = Math.min(Math.max(inputFillLevel, 0), 100);

    // Calculate the height based on the input fill level.
    const fillHeight = (validatedFillLevel) + "%";

    fillLevelElement.style.height = fillHeight;

    // Change the fill color for levels below 60.
    if (validatedFillLevel < 85) {
        fillLevelElement.style.background = orange; // Change to your preferred color
    } else {
        fillLevelElement.style.background = "red"; // Default fill color
    }
}

updateCylinderFillLevel();


