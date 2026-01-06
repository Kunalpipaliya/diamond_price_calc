const RAP_TABLES = [
    { min: 0.30, max: 0.39, data: { "D": { "IF": 31, "VVS1": 24, "VVS2": 21, "VS1": 19, "VS2": 17, "SI1": 15, "SI2": 14, "SI3": 13, "I1": 10 }, "E": { "IF": 27, "VVS1": 22, "VVS2": 20, "VS1": 18, "VS2": 16, "SI1": 14, "SI2": 13, "SI3": 12, "I1": 9 } } },
    { min: 0.40, max: 0.49, data: { "D": { "IF": 35, "VVS1": 28, "VVS2": 24, "VS1": 22, "VS2": 20, "SI1": 18, "SI2": 16, "SI3": 15, "I1": 11 } } },
    { min: 0.50, max: 0.69, data: { "D": { "IF": 55, "VVS1": 44, "VVS2": 34, "VS1": 28, "VS2": 25, "SI1": 22, "SI2": 18, "SI3": 16, "I1": 12 } } },
    { min: 0.70, max: 0.89, data: { "D": { "IF": 70, "VVS1": 56, "VVS2": 44, "VS1": 38, "VS2": 33, "SI1": 29, "SI2": 25, "SI3": 23, "I1": 15 } } },
    { min: 0.90, max: 0.99, data: { "D": { "IF": 104, "VVS1": 89, "VVS2": 67, "VS1": 57, "VS2": 49, "SI1": 40, "SI2": 32, "SI3": 28, "I1": 20 } } },
    {
        min: 1.00, max: 1.49, data: {
            "D": { "IF": 160, "VVS1": 128, "VVS2": 102, "VS1": 87, "VS2": 73, "SI1": 55, "SI2": 44, "SI3": 39, "I1": 36, "I2": 25, "I3": 16 },
            "E": { "IF": 125, "VVS1": 111, "VVS2": 93, "VS1": 79, "VS2": 66, "SI1": 51, "SI2": 41, "SI3": 36, "I1": 33, "I2": 24, "I3": 15 },
            "F": { "IF": 107, "VVS1": 97, "VVS2": 84, "VS1": 72, "VS2": 60, "SI1": 48, "SI2": 38, "SI3": 33, "I1": 31, "I2": 23, "I3": 14 },
            "G": { "IF": 82, "VVS1": 77, "VVS2": 70, "VS1": 62, "VS2": 54, "SI1": 44, "SI2": 36, "SI3": 31, "I1": 29, "I2": 22, "I3": 13 }
        }
    },
    {
        min: 1.50, max: 1.99, data: {
            "D": { "IF": 210, "VVS1": 187, "VVS2": 154, "VS1": 134, "VS2": 120, "SI1": 96, "SI2": 78, "SI3": 69, "I1": 57, "I2": 35, "I3": 18 },
            "E": { "IF": 188, "VVS1": 173, "VVS2": 143, "VS1": 122, "VS2": 110, "SI1": 89, "SI2": 71, "SI3": 63, "I1": 54, "I2": 33, "I3": 17 }
        }
    },
    {
        min: 2.00, max: 2.99, data: {
            "D": { "IF": 330, "VVS1": 275, "VVS2": 235, "VS1": 205, "VS2": 175, "SI1": 141, "SI2": 113, "SI3": 95, "I1": 80, "I2": 41, "I3": 19 },
            "E": { "IF": 270, "VVS1": 245, "VVS2": 210, "VS1": 190, "VS2": 160, "SI1": 132, "SI2": 105, "SI3": 88, "I1": 76, "I2": 39, "I3": 18 }
        }
    }
];

function runCalculation() {
    const weight = parseFloat(document.getElementById('weightInput').value);
    const color = document.getElementById('colorInput').value;
    const clarity = document.getElementById('clarityInput').value;

    if (!weight || weight < 0.30 || weight > 2.99) {
        alert("Enter weight between 0.30 and 2.99 ct");
        return;
    }

    const tableMatch = RAP_TABLES.find(t => weight >= t.min && weight <= t.max);

    if (tableMatch && tableMatch.data[color] && tableMatch.data[color][clarity]) {
        const ratePerCarat = tableMatch.data[color][clarity] * 100;
        const totalValue = ratePerCarat * weight;

        document.getElementById('perCaratResult').innerText = "$" + ratePerCarat.toLocaleString();
        document.getElementById('totalValueResult').innerText = "$" + totalValue.toLocaleString(undefined, { minimumFractionDigits: 2 });
        document.getElementById('resultArea').style.display = 'block';
    } else {
        alert("Combination not found in the current price table.");
    }
}