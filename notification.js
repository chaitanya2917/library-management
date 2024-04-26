// import { readDistance } from './ultrasonic.js';
document.addEventListener("DOMContentLoaded", function () {
    const notificationList = document.getElementById("notification-list");
    function addNotification(message) {
        const notificationItem = document.createElement("li");
        notificationItem.classList.add("notification-item");
        const notificationTime = document.createElement("p");
        notificationTime.classList.add("notification-time");
        notificationTime.textContent = getCurrentTime();
        const notificationMessage = document.createElement("p");
        notificationMessage.textContent = message;
        notificationItem.appendChild(notificationTime);
        notificationItem.appendChild(notificationMessage);
        notificationList.appendChild(notificationItem);
    }
    function getCurrentTime() {
        const now = new Date();
        const hours = now.getHours().toString().padStart(2, '0');
        const minutes = now.getMinutes().toString().padStart(2, '0');
        const time = `${hours}:${minutes}`;
        return time;
    }

    // Function to check waste level and add a notification
    function checkWasteLevel() {
        const wasteLevel = 90; // Replace this with your actual waste level reading logic
        if (wasteLevel >= 80) {
            const message = `Waste level reached ${wasteLevel}% collect the garbage`;
            addNotification(message);
        }
    }

    checkWasteLevel();

    // setInterval(checkWasteLevel, 1000);
});
