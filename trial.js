// Create a pie chart using Chart.js
const pieChartCanvas = document.getElementById('pieChart');
const pieChartData = {
    labels: ['Label 1', 'Label 2', 'Label 3'],
    datasets: [{
        data: [30, 40, 30],
        backgroundColor: ['#ff5733', '#33ff57', '#5733ff'],
    }],
};

const pieChart = new Chart(pieChartCanvas, {
    type: 'doughnut',
    data: pieChartData,
});
