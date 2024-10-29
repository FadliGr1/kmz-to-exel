const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const convertBtn = document.getElementById("convertBtn");
const downloadBtn = document.getElementById("downloadBtn");
const loadingOverlay = document.getElementById("loadingOverlay");
const loadingText = document.getElementById("loadingText");
const resultDisplay = document.getElementById("resultDisplay");
const previewModal = new bootstrap.Modal(document.getElementById('previewModal'));
const previewTabs = document.getElementById('previewTabs');
const previewTabContent = document.getElementById('previewTabContent');
let convertedData = {};
let selectedFile = null;
let convertedZip = null;

// Event Listeners
dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", handleFileDrop);
fileInput.addEventListener("change", handleFileSelect);
convertBtn.addEventListener("click", convertToExcel);
downloadBtn.addEventListener("click", downloadConvertedFiles);

// File handling functions
function handleFileDrop(e) {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  selectedFile = e.dataTransfer.files[0];
  updateUI();
}

function handleFileSelect(e) {
  selectedFile = e.target.files[0];
  updateUI();
}

function updateUI() {
  if (selectedFile) {
    dropZone.textContent = `Selected file: ${selectedFile.name}`;
    convertBtn.disabled = false;
  }
}

// Text animation function
async function animateText(text, duration) {
  loadingText.innerHTML = "";
  const chars = text.split("");
  const delayPerChar = duration / chars.length;
  for (let char of chars) {
    await new Promise((resolve) => {
      setTimeout(() => {
        const span = document.createElement("span");
        span.textContent = char;
        span.style.animation = "fadeIn 0.5s forwards";
        loadingText.appendChild(span);
        resolve();
      }, delayPerChar);
    });
  }
}

// Main conversion function
async function convertToExcel() {
  if (!selectedFile) return;

  const startTime = Date.now();
  loadingOverlay.style.display = "flex";
  resultDisplay.style.display = "none";
  downloadBtn.style.display = "none";
  convertedData = {}; // Reset converted data

  // Start text animation
  animateText("Mengkonversi", 2000);

  try {
    const zip = new JSZip();
    const kmzContent = await zip.loadAsync(selectedFile);
    const outputZip = new JSZip();

    for (let [path, zipEntry] of Object.entries(kmzContent.files)) {
      if (path.endsWith(".kml")) {
        const kmlContent = await zipEntry.async("text");
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(kmlContent, "text/xml");
        const folders = xmlDoc.getElementsByTagName("Folder");

        await processFolder(folders, "", outputZip);
      }
    }

    convertedZip = await outputZip.generateAsync({ type: "blob" });

    // Ensure minimum loading time
    const elapsedTime = Date.now() - startTime;
    if (elapsedTime < 5000) {
      await new Promise((resolve) => setTimeout(resolve, 5000 - elapsedTime));
    }

    displayResult(outputZip);
  } catch (error) {
    console.error("Error converting file:", error);
    alert("An error occurred while converting the file. Please try again.");
  } finally {
    loadingOverlay.style.display = "none";
  }
}

// Process folder function
async function processFolder(folders, parentPath, outputZip) {
  for (let folder of folders) {
    const folderName = folder.getElementsByTagName("name")[0]?.textContent || "Unnamed Folder";
    const currentPath = parentPath ? `${parentPath}/${folderName}` : folderName;

    const placemarks = folder.getElementsByTagName("Placemark");
    const subFolders = folder.getElementsByTagName("Folder");

    if (placemarks.length > 0 && subFolders.length === 0) {
      const data = [];
      for (let placemark of placemarks) {
        const name = placemark.getElementsByTagName("name")[0]?.textContent || "";
        const coordinates = placemark.getElementsByTagName("coordinates")[0]?.textContent || "";
        const [longitude, latitude] = coordinates.split(",");
        data.push({
          Name: name,
          Latitude: latitude,
          Longitude: longitude,
        });
      }

      // Save data for preview
      convertedData[currentPath] = data;

      const worksheet = XLSX.utils.json_to_sheet(data);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Placemarks");
      const excelBuffer = XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array",
      });
      outputZip.file(`${currentPath}.xlsx`, excelBuffer);
    }

    if (subFolders.length > 0) {
      await processFolder(subFolders, currentPath, outputZip);
    }
  }
}

// Display results function
function displayResult(outputZip) {
  const fileList = Object.keys(outputZip.files).filter(
    (filename) => !outputZip.files[filename].dir
  );
  let resultHTML = "<h3>Converted Files:</h3><ul>";
  fileList.forEach((filename) => {
    resultHTML += `<li onclick="showPreview('${filename.replace(/\.xlsx$/, '')}')">${filename}</li>`;
  });
  resultHTML += "</ul>";
  resultDisplay.innerHTML = resultHTML;
  resultDisplay.style.display = "block";
  downloadBtn.style.display = "block";
}

// Preview function
function showPreview(path) {
    const data = convertedData[path];
    if (!data) return;
  
    // Clear existing tabs and content
    previewTabs.innerHTML = '';
    previewTabContent.innerHTML = '';
  
    // Create table view tab
    const tableTab = document.createElement('li');
    tableTab.className = 'nav-item';
    tableTab.innerHTML = `
      <button class="nav-link active" data-bs-toggle="tab" data-bs-target="#tableView" type="button">
        Table View
      </button>
    `;
    previewTabs.appendChild(tableTab);
  
    // Create map view tab
    const mapTab = document.createElement('li');
    mapTab.className = 'nav-item';
    mapTab.innerHTML = `
      <button class="nav-link" data-bs-toggle="tab" data-bs-target="#mapView" type="button">
        Map View
      </button>
    `;
    previewTabs.appendChild(mapTab);
  
    // Create table view content
    const tableContent = document.createElement('div');
    tableContent.className = 'tab-pane fade show active';
    tableContent.id = 'tableView';
    
    let tableHTML = '<table class="preview-table"><thead><tr>';
    const headers = Object.keys(data[0]);
    headers.forEach(header => {
      tableHTML += `<th>${header}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';
    
    data.forEach(row => {
      tableHTML += '<tr>';
      headers.forEach(header => {
        tableHTML += `<td>${row[header]}</td>`;
      });
      tableHTML += '</tr>';
    });
    tableHTML += '</tbody></table>';
    
    tableContent.innerHTML = tableHTML;
    previewTabContent.appendChild(tableContent);
  
    // Create map view content
    const mapContent = document.createElement('div');
    mapContent.className = 'tab-pane fade';
    mapContent.id = 'mapView';
    
    // Create map container
    const mapContainer = document.createElement('div');
    mapContainer.id = 'map';
    mapContainer.style.height = '400px';
    mapContent.appendChild(mapContainer);
    previewTabContent.appendChild(mapContent);
  
    // Show modal
    previewModal.show();
  
    // Initialize map after modal is shown
    previewModal._element.addEventListener('shown.bs.modal', function () {
      // Find first valid coordinates for initial center
      let initialLat = -6.2088;  // Default to Jakarta coordinates
      let initialLng = 106.8456;
      let hasValidCoordinates = false;
  
      for (const point of data) {
        const lat = parseFloat(point.Latitude);
        const lng = parseFloat(point.Longitude);
        if (!isNaN(lat) && !isNaN(lng)) {
          initialLat = lat;
          initialLng = lng;
          hasValidCoordinates = true;
          break;
        }
      }
  
      // Initialize the map with center and zoom
      const map = L.map('map', {
        center: [initialLat, initialLng],
        zoom: hasValidCoordinates ? 13 : 5  // Zoom level 13 if we have coordinates, 5 if using default
      });
      
      // Add OpenStreetMap tiles
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '© OpenStreetMap contributors'
      }).addTo(map);
  
      // Create bounds for auto-zoom
      const bounds = L.latLngBounds();
      let hasMarkers = false;
      
      // Add markers for each point
      data.forEach(point => {
        if (point.Latitude && point.Longitude) {
          const lat = parseFloat(point.Latitude);
          const lng = parseFloat(point.Longitude);
          
          if (!isNaN(lat) && !isNaN(lng)) {
            // Create marker with popup
            const marker = L.marker([lat, lng])
              .bindPopup(`<b>${point.Name}</b><br>Lat: ${lat}<br>Lng: ${lng}`)
              .addTo(map);
            
            // Extend bounds to include this point
            bounds.extend([lat, lng]);
            hasMarkers = true;
          }
        }
      });
  
      // Fit map to bounds if we have markers
      if (hasMarkers) {
        map.fitBounds(bounds, { 
          padding: [50, 50],
          maxZoom: 15  // Prevent excessive zoom on single point
        });
      }
  
      // Fix map display issues when shown in modal
      setTimeout(() => {
        map.invalidateSize();
      }, 100);
  
      // Update map when tab is shown
      const mapTabButton = document.querySelector('button[data-bs-target="#mapView"]');
      mapTabButton.addEventListener('shown.bs.tab', function () {
        map.invalidateSize();
        if (hasMarkers) {
          map.fitBounds(bounds, { 
            padding: [50, 50],
            maxZoom: 15
          });
        }
      });
    }, { once: true }); // Remove listener after first execution
  }

// Download function
function downloadConvertedFiles() {
  if (convertedZip) {
    const url = window.URL.createObjectURL(convertedZip);
    const a = document.createElement("a");
    a.href = url;
    a.download = "kmz_converted.zip";
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
  }
}