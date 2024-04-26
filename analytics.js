// Updated category labels for waste management analytics
const categoryLabels = [
    "Plastic Waste",
    "Organic Waste",
    "Paper Waste",
    "Glass Waste",
    "Metal Waste",
    "E-Waste",
    "Hazardous Waste",
    "Construction Waste",
    "Biodegradable Waste",
    "Non-Biodegradable Waste",
    "Medical Waste",
    "Other Waste",
];

document.addEventListener("DOMContentLoaded", function () {
    const ctxBar = document.getElementById("barChart").getContext("2d");
    const ctxPie = document.getElementById("pieChart").getContext("2d");

    // Historical waste management data for different categories
    const category1Data = [10, 20, 15, 25, 30, 18, 12, 22, 27, 17, 28, 32];
    const category2Data = [18, 22, 27, 15, 21, 25, 30, 20, 12, 28, 32, 17];
    const category3Data = [12, 25, 30, 17, 28, 32, 18, 22, 27, 15, 21, 20];
    const category4Data = [32, 18, 22, 27, 15, 21, 25, 30, 20, 12, 28, 17];

    const barData = {
        labels: categoryLabels,
        datasets: [
            {
                label: categoryLabels[0],
                data: category1Data,
                backgroundColor: "rgba(75, 192, 192, 0.2)",
                borderColor: "rgba(75, 192, 192, 1)",
                borderWidth: 1,
            },
            {
                label: categoryLabels[1],
                data: category2Data,
                backgroundColor: "rgba(255, 99, 132, 0.2)",
                borderColor: "rgba(255, 99, 132, 1)",
                borderWidth: 1,
            },
            {
                label: categoryLabels[2],
                data: category3Data,
                backgroundColor: "rgba(54, 162, 235, 0.2)",
                borderColor: "rgba(54, 162, 235, 1)",
                borderWidth: 1,
            },
            {
                label: categoryLabels[3],
                data: category4Data,
                backgroundColor: "rgba(255, 206, 86, 0.2)",
                borderColor: "rgba(255, 206, 86, 1)",
                borderWidth: 1,
            },
        ],
    };

    const pieData = {
        labels: categoryLabels,
        datasets: [
            {
                data: [50, 30, 15, 25, 10, 20, 35, 45, 18, 12, 22, 27],
                backgroundColor: [
                    "#FF5733",
                    "#33FF33",
                    "#3366FF",
                    "#FF9900",
                    "#FF3366",
                    "#66FF33",
                    "#3366CC",
                    "#CC9900",
                    "#6633FF",
                    "#FF3399",
                    "#99FF33",
                    "#33CCFF",
                ],
            },
        ],
    };

    const barConfig = {
        type: "bar",
        data: barData,
        options: {
            scales: {
                y: {
                    beginAtZero: true,
                },
            },
        },
    };

    const pieConfig = {
        type: "doughnut",
        data: pieData,
    };

    const barChart = new Chart(ctxBar, barConfig);
    const pieChart = new Chart(ctxPie, pieConfig);
});
