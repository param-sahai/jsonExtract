const fs = require('fs');
const XLSX = require('xlsx');

// Sample JSON data
const jsonData = {
    users: [
        {
            user_id: 1,
            name: "John Doe",
            accounts: [
                {
                    account_id: 101,
                    plan: "Unlimited",
                    devices: [
                        { device_id: 1001, device_name: "iPhone 12" },
                        { device_id: 1002, device_name: "Galaxy S21" }
                    ]
                },
                {
                    account_id: 102,
                    plan: "Basic",
                    devices: [
                        { device_id: 1003, device_name: "Pixel 5" }
                    ]
                }
            ]
        },
        {
            user_id: 2,
            name: "Jane Smith",
            accounts: [
                {
                    account_id: 103,
                    plan: "Family",
                    devices: [
                        { device_id: 1004, device_name: "iPhone 11" }
                    ]
                }
            ]
        }
    ]
};

// Function to flatten JSON data
function flattenJson(data, parentKey = '', sep = '_') {
    let items = {};
    for (let [k, v] of Object.entries(data)) {
        let newKey = parentKey ? `${parentKey}${sep}${k}` : k;

        if (v && typeof v === 'object' && !Array.isArray(v)) {
            Object.assign(items, flattenJson(v, newKey, sep));
        } else if (Array.isArray(v)) {
            v.forEach((item, i) => {
                if (item && typeof item === 'object') {
                    Object.assign(items, flattenJson(item, `${newKey}${sep}${i}`, sep));
                } else {
                    items[`${newKey}${sep}${i}`] = item;
                }
            });
        } else {
            items[newKey] = v !== undefined ? v : null;
        }
    }
    return items;
}

// Flatten the JSON data
const flattenedData = jsonData.users.map(user => flattenJson(user));

// Convert to worksheet
const worksheet = XLSX.utils.json_to_sheet(flattenedData);

// Create a new workbook and append the worksheet
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// Write to Excel file
XLSX.writeFile(workbook, 'output_nodejs.xlsx');

console.log("Excel file created: output_nodejs.xlsx");
