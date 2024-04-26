
import { readDistance } from './ultrasonic.js';
export function isIRLimitExceeded() {
    const isExceeded = readDistance();
    console.log("irsensor reading:", isExceeded)
    if (isExceeded > 80) {
        return 1; 
    } else {
        return 0; 
    }
}
