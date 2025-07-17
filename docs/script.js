import * as XLSX from 'https://cdn.sheetjs.com/xlsx-0.20.0/package/xlsx.mjs';

let workbook, selectedSheet, jsonData = [];
let apiKeyVisible = false;

const apiKeyInput = document.getElementById("apiKey");
const eyeBtn = document.getElementById("toggleApiKey");
eyeBtn.innerHTML = "🔓";

eyeBtn.addEventListener("click", () => {
  apiKeyVisible = !apiKeyVisible;
  apiKeyInput.type = apiKeyVisible ? "text" : "password";
  eyeBtn.innerHTML = apiKeyVisible ? "🔒" : "🔓";
});

const dropArea = document.getElementById("drop-area");
const fileInput = document.getElementById("fileElem");
const tablePreview = document.getElementById("tablePreview");
const addressSelect = document.getElementById("addressSelect");
const addressColumnContainer = document.getElementById("addressColumnContainer");
const geocodeBtn = document.getElementById("geocodeBtn");
const resetBtn = document.getElementById("resetBtn");
const downloadBtn = document.getElementById("downloadBtn");
const actionButtons = document.getElementById("action-buttons");

["dragenter", "dragover", "dragleave", "drop"].forEach(event => {
  dropArea.addEventListener(event, e => {
    e.preventDefault();
    e.stopPropagation();
  }, false);
});

dropArea.addEventListener("drop", handleDrop, false);
fileInput.addEventListener("change", handleFile, false);
geocodeBtn.addEventListener("click", geocodeAddresses);
resetBtn.addEventListener("click", resetApp);
downloadBtn.addEventListener("click", downloadExcel);

function handleDrop(e) {
  const files = e.dataTransfer.files;
  if (files.length > 0) readExcel(files[0]);
}

function handleFile(e) {
  const file = e.target.files[0];
  if (file) readExcel(file);
}

function readExcel(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    selectedSheet = workbook.Sheets[sheetName];
    jsonData = XLSX.utils.sheet_to_json(selectedSheet);
    populateColumnSelector(Object.keys(jsonData[0]));
    previewTable(jsonData);
    addressColumnContainer.classList.remove("hidden");
    actionButtons.classList.remove("hidden");
    downloadBtn.classList.add("hidden"); // Hide initially
  };
  reader.readAsArrayBuffer(file);
}

function populateColumnSelector(columns) {
  addressSelect.innerHTML = "";
  columns.forEach(col => {
    const option = document.createElement("option");
    option.value = col;
    option.textContent = col;
    addressSelect.appendChild(option);
  });
}

function previewTable(data) {
  tablePreview.innerHTML = "";
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  Object.keys(data[0]).forEach(key => {
    const th = document.createElement("th");
    th.textContent = key;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  data.forEach(row => {
    const tr = document.createElement("tr");
    Object.values(row).forEach(val => {
      const td = document.createElement("td");
      td.textContent = val;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  tablePreview.appendChild(table);
}

async function geocodeAddresses() {
  const apiKey = apiKeyInput.value.trim();
  const addressCol = addressSelect.value;

  if (!apiKey || !addressCol) {
    alert("Please provide API Key and select Address column.");
    return;
  }

  for (let row of jsonData) {
    const address = row[addressCol];
    if (!address) continue;

    const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`);
    const result = await response.json();

    if (result.status === "OK" && result.results.length > 0) {
      const location = result.results[0].geometry.location;
      row["Latitude"] = location.lat;
      row["Longitude"] = location.lng;
    } else {
      row["Latitude"] = "N/A";
      row["Longitude"] = "N/A";
    }
  }

  previewTable(jsonData);
  downloadBtn.classList.remove("hidden"); // Show after geocoding
}

function downloadExcel() {
  const newSheet = XLSX.utils.json_to_sheet(jsonData);
  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWb, newSheet, "Geocoded");
  XLSX.writeFile(newWb, "geocoded_results.xlsx");
}

function resetApp() {
  workbook = null;
  selectedSheet = null;
  jsonData = [];
  tablePreview.innerHTML = "";
  addressSelect.innerHTML = "";
  fileInput.value = "";
  addressColumnContainer.classList.add("hidden");
  actionButtons.classList.add("hidden");
}
