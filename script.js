const fs = require('fs');
const XLSX = require('xlsx');

// Sample JSON data
const jsonData = {
    name: "Alice",
    age: 25,
    address: {
        street: "123 Elm St",
        city: "Wonderland"
    },
    phones: [
        { type: "home", number: "123-456-7890" },
        { type: "work", number: "987-654-3210" }
    ]
};

// Function to flatten JSON data
function flattenJson(data, parentKey = '', sep = '_') {
    let items = {};
    for (let [k, v] of Object.entries(data)) {
        let newKey = parentKey ? `${parentKey}${sep}${k}` : k;
        if (typeof v === 'object' && !Array.isArray(v)) {
            Object.assign(items, flattenJson(v, newKey, sep));
        } else if (Array.isArray(v)) {
            v.forEach((item, i) => {
                Object.assign(items, flattenJson(item, `${newKey}${sep}${i}`, sep));
            });
        } else {
            items[newKey] = v;
        }
    }
    return items;
}

// Flatten the JSON data
const flatData = flattenJson(jsonData);

// Convert to worksheet
const worksheet = XLSX.utils.json_to_sheet([flatData]);

// Create a new workbook and append the worksheet
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// Write to Excel file
XLSX.writeFile(workbook, 'output_nodejs.xlsx');

console.log("Excel file created: output_nodejs.xlsx");
