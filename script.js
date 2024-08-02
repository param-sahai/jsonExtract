const fs = require('fs');
const XLSX = require('xlsx');

// Sample complex nested JSON structure (Replace this with your actual JSON data)
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
        // This is where a potential issue could arise (e.g., missing data)
        {
            user_id: 2,
            name: "Jane Smith",
            accounts: null  // Imagine 'accounts' is unexpectedly null
        }
    ]
};

// Function to flatten JSON data with error handling
function flattenJson(data, parentKey = '', sep = '_') {
    let items = {};

    if (!data || typeof data !== 'object') {
        return { [parentKey]: data }; // Return as is if it's not an object
    }

    for (let [k, v] of Object.entries(data)) {
        let newKey = parentKey ? `${parentKey}${sep}${k}` : k;

        if (v && typeof v === 'object' && !Array.isArray(v)) {
            Object.assign(items, flattenJson(v, newKey, sep));
        } else if (Array.isArray(v)) {
            v.forEach((item, i) => {
                if (item && typeof item === 'object') {
                    Object.assign(items, flattenJson(item, `${newKey}${sep}${i}`, sep));
                } else {
                    items[`${newKey}${sep}${i}`] = item !== undefined ? item : null;
                }
            });
        } else {
            items[newKey] = v !== undefined ? v : null;
        }
    }

    return items;
}

// Handle the JSON structure safely and ensure processing
let flattenedData;

try {
    // Safely access 'users' and handle undefined, null, or missing data
    flattenedData = (Array.isArray(jsonData.users) ? jsonData.users : []).map(user => flattenJson(user));
} catch (error) {
    console.error("An error occurred while flattening the JSON data:", error);
    flattenedData = [];  // Fallback to an empty array if an error occurs
}

// Convert the flattened data to a worksheet
const worksheet = XLSX.utils.json_to_sheet(flattenedData);

// Create a new workbook and append the worksheet
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// Write the Excel file
XLSX.writeFile(workbook, 'output_nodejs.xlsx');

console.log("Excel file created: output_nodejs.xlsx");
