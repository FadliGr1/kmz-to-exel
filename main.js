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
const settingsModal = new bootstrap.Modal(document.getElementById('settingsModal'));
const startConversionBtn = document.getElementById('startConversion');
const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB
const MIN_FILE_SIZE = 1024; // 1KB
const fileInfoModal = new bootstrap.Modal(document.getElementById('fileInfoModal'));
let exportSettings = {};
let convertedData = {};
let mergedFileName = '';
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
convertBtn.addEventListener("click", () => {
    if (selectedFile) {
      validateAndShowFileInfo(selectedFile);
    }
  });
  
  downloadBtn.addEventListener("click", () => {
      if (convertedZip) {
          downloadConvertedFiles();
      }
  });
startConversionBtn.addEventListener("click", () => {
    // Ambil pengaturan dari form
    exportSettings = {
      format: document.querySelector('input[name="outputFormat"]:checked').value,
      columns: Array.from(document.querySelectorAll('#columnSelection input:checked')).map(cb => cb.value),
      filename: document.getElementById('outputFilename').value || 'converted_data',
      mode: document.querySelector('input[name="outputMode"]:checked').value
    };
    
    settingsModal.hide();
    convertToExcel();
  });


// File handling functions
async function handleFileDrop(e) {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (await validateAndShowFileInfo(file)) {
      selectedFile = file;
      updateUI();
    }
  }

  async function handleFileSelect(e) {
    const file = e.target.files[0];
    if (await validateAndShowFileInfo(file)) {
      selectedFile = file;
      updateUI();
    }
  }

  async function validateAndShowFileInfo(file) {
    const warnings = [];
    let isValid = true;
  
    // Reset file info
    document.getElementById('fileInfoName').textContent = file.name;
    document.getElementById('fileInfoSize').textContent = formatFileSize(file.size);
  
    // Validasi dasar
    if (!file.name.toLowerCase().endsWith('.kmz')) {
      warnings.push('File must be in KMZ format');
      isValid = false;
    }
  
    if (file.size > MAX_FILE_SIZE) {
      warnings.push('File size exceeds 50MB limit');
      isValid = false;
    }
  
    if (file.size < MIN_FILE_SIZE) {
      warnings.push('File seems too small to be a valid KMZ');
      isValid = false;
    }
  
    // Validasi struktur KMZ
    try {
      const zip = new JSZip();
      const kmzContent = await zip.loadAsync(file);
      let hasKML = false;
      let folderCount = 0;
      let placemarkCount = 0;
  
      for (let [path, zipEntry] of Object.entries(kmzContent.files)) {
        if (path.endsWith('.kml')) {
          hasKML = true;
          const kmlContent = await zipEntry.async("text");
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(kmlContent, "text/xml");
          
          // Count folders and placemarks
          folderCount = xmlDoc.getElementsByTagName("Folder").length;
          placemarkCount = xmlDoc.getElementsByTagName("Placemark").length;
  
          // Validate KML structure
          if (placemarkCount === 0) {
            warnings.push('No placemarks found in the file');
          }
  
          document.getElementById('fileInfoStructure').textContent = 
            `${folderCount} folders found`;
          document.getElementById('fileInfoPlacemarks').textContent = 
            `${placemarkCount} placemarks found`;
        }
      }
  
      if (!hasKML) {
        warnings.push('No KML file found inside KMZ');
        isValid = false;
      }
  
    } catch (error) {
      console.error('Error validating KMZ:', error);
      warnings.push('Invalid KMZ file structure');
      isValid = false;
    }
  
    // Tampilkan warnings
    const warningsContainer = document.getElementById('fileWarnings');
    if (warnings.length > 0) {
      warningsContainer.classList.remove('d-none');
      warningsContainer.innerHTML = warnings.map(warn => 
        `<div class="warning-item"><i class="bi bi-exclamation-triangle"></i>${warn}</div>`
      ).join('');
    } else {
      warningsContainer.classList.add('d-none');
    }
  
    // Show modal with validation results
    fileInfoModal.show();
  
    // Setup proceed button
    const proceedBtn = document.getElementById('proceedConversion');
    proceedBtn.disabled = !isValid;
    proceedBtn.onclick = () => {
      fileInfoModal.hide();
      if (isValid) {
        settingsModal.show();
      }
    };
  
    return isValid;
  }
  
  // Helper function untuk format ukuran file
  function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  function updateUI() {
    if (selectedFile) {
      dropZone.innerHTML = `
        <div class="drop-zone-content animate__animated animate__fadeIn">
          <i class="bi bi-file-earmark-check big-icon"></i>
          <p class="file-name mb-3">${selectedFile.name}</p>
          <div class="file-meta text-muted mb-3">
            <small>
              <i class="bi bi-hdd"></i>
              ${formatFileSize(selectedFile.size)}
            </small>
          </div>
          <button class="btn btn-outline-primary btn-sm">
            <i class="bi bi-arrow-repeat"></i>
            Change File
          </button>
        </div>
      `;
      convertBtn.disabled = false;
      convertBtn.classList.add('animate__animated', 'animate__fadeIn');
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
    convertedData = {};
  
    try {
      const zip = new JSZip();
      const kmzContent = await zip.loadAsync(selectedFile);
      const outputZip = new JSZip();
      let mergedData = [];
      const loadingTexts = [
        "Reading KMZ file...",
        "Processing coordinates...",
        "Generating Excel files...",
        "Finalizing conversion..."
      ];

      // Log untuk melihat isi kmzContent
      console.log('KMZ Content:', kmzContent.files);

      for (const text of loadingTexts) {
        await animateText(text, 1000);
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
  
      for (let [path, zipEntry] of Object.entries(kmzContent.files)) {
        if (path.endsWith(".kml")) {
          const kmlContent = await zipEntry.async("text");
          // Log untuk melihat isi KML
          console.log('KML Content:', kmlContent);

          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(kmlContent, "text/xml");
          const folders = xmlDoc.getElementsByTagName("Folder");
          
          // Log untuk melihat jumlah folder
          console.log('Found folders:', folders.length);
  
          if (exportSettings.mode === 'merged') {
            mergedData = mergedData.concat(await processFolder(folders, "", null, true));
          } else {
            await processFolder(folders, "", outputZip, false);
          }
        }
      }
  
      // Log untuk melihat hasil konversi
      console.log('Merged Data:', mergedData);
      console.log('Output Zip Files:', outputZip.files);
  
      if (exportSettings.mode === 'merged') {
        // Simpan data merged untuk preview
        mergedFileName = exportSettings.filename || 'merged_data';
        convertedData[mergedFileName] = mergedData;

        // Buat ZIP baru untuk file merged
        const mergedZip = new JSZip();

        if (exportSettings.format === 'xlsx') {
          const workbook = XLSX.utils.book_new();
          const worksheet = XLSX.utils.json_to_sheet(mergedData);
          XLSX.utils.book_append_sheet(workbook, worksheet, "All Data");
          const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
          // Masukkan file Excel ke dalam ZIP
          mergedZip.file(`${mergedFileName}.xlsx`, excelBuffer);
        } else {
          // Untuk CSV
          const worksheet = XLSX.utils.json_to_sheet(mergedData);
          const csvContent = XLSX.utils.sheet_to_csv(worksheet);
          // Masukkan file CSV ke dalam ZIP
          mergedZip.file(`${mergedFileName}.csv`, csvContent);
        }

        // Generate ZIP final
        convertedZip = await mergedZip.generateAsync({ type: "blob" });
      } else {
        convertedZip = await outputZip.generateAsync({ type: "blob" });
      }
  
      const elapsedTime = Date.now() - startTime;
      if (elapsedTime < 5000) {
        await new Promise(resolve => setTimeout(resolve, 5000 - elapsedTime));
      }
  
      displayResult(exportSettings.mode === 'merged' ? mergedZip : outputZip);
    } catch (error) {
      console.error("Error converting file:", error);
      alert("An error occurred while converting the file. Please try again.");
    } finally {
      loadingOverlay.style.display = "none";
    }
}
// Process folder function
async function processFolder(folders, parentPath, outputZip, returnData = false) {
    let allData = [];
    
    console.log('Processing folders:', folders.length);
    
    for (let folder of folders) {
      const folderName = folder.getElementsByTagName("name")[0]?.textContent || "Unnamed Folder";
      const currentPath = parentPath ? `${parentPath}/${folderName}` : folderName;
  
      console.log('Processing folder:', folderName);

      const placemarks = folder.getElementsByTagName("Placemark");
      const subFolders = folder.getElementsByTagName("Folder");
  
      console.log('Found placemarks:', placemarks.length);
      console.log('Found subfolders:', subFolders.length);

      if (placemarks.length > 0 && subFolders.length === 0) {
        const data = [];
        for (let placemark of placemarks) {
          const rowData = {};
          if (exportSettings.columns.includes('Name')) {
            rowData.Name = placemark.getElementsByTagName("name")[0]?.textContent || "";
          }
          if (exportSettings.columns.includes('Latitude') || exportSettings.columns.includes('Longitude')) {
            const coordinates = placemark.getElementsByTagName("coordinates")[0]?.textContent || "";
            const [longitude, latitude] = coordinates.split(",");
            if (exportSettings.columns.includes('Latitude')) rowData.Latitude = latitude;
            if (exportSettings.columns.includes('Longitude')) rowData.Longitude = longitude;
          }
          data.push(rowData);
        }
  
        console.log('Processed data:', data);

        // Selalu simpan data untuk preview
        convertedData[currentPath] = data;
  
        if (returnData) {
          data.forEach(row => {
            row.Folder = currentPath;
          });
          allData = allData.concat(data);
        } else {
          if (exportSettings.format === 'xlsx') {
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data), "Placemarks");
            const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
            outputZip.file(`${currentPath}.xlsx`, excelBuffer);
          } else {
            const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(data));
            outputZip.file(`${currentPath}.csv`, csvContent);
          }
        }
      }
  
      if (subFolders.length > 0) {
        const subData = await processFolder(subFolders, currentPath, outputZip, returnData);
        if (returnData) {
          allData = allData.concat(subData);
        }
      }
    }
  
    console.log('Returning data:', allData);
    return returnData ? allData : null;
}

  function createWorksheet(data) {
    // Pastikan data memiliki format yang konsisten
    const worksheet = XLSX.utils.json_to_sheet(data);
    
    // Atur lebar kolom
    const columnWidths = {
      Folder: { wch: 40 },
      Name: { wch: 30 },
      Latitude: { wch: 15 },
      Longitude: { wch: 15 }
    };
    
    worksheet['!cols'] = Object.keys(data[0]).map(key => columnWidths[key] || { wch: 12 });
    
    return worksheet;
  }

// Display results function
function displayResult(outputZip) {
    const fileList = Object.keys(outputZip.files).filter(
        (filename) => !outputZip.files[filename].dir
    );

    let resultHTML = `
        <div class="result-header animate__animated animate__fadeIn">
            <div class="result-header-content">
                <div class="result-title">
                    <i class="bi bi-check-circle-fill text-success me-2"></i>
                    Converted Files
                    <span class="file-count">${fileList.length} files</span>
                </div>
                <div class="result-info">
                    <small class="text-muted">
                        <i class="bi bi-info-circle me-1"></i>
                        Click on file to preview content
                    </small>
                </div>
            </div>
        </div>
        <ul class="converted-files-list">
    `;

    fileList.forEach((filename, index) => {
        const fileIcon = exportSettings.format === 'xlsx' ? 'file-earmark-excel' : 'file-earmark-text';
        const fileExtension = exportSettings.format === 'xlsx' ? 'Excel' : 'CSV';
        const previewPath = filename.replace(/\.(xlsx|csv)$/, '');
        
        resultHTML += `
            <li class="file-item animate__animated animate__fadeInUp animate__delay-${index}s" 
                onclick="showPreview('${previewPath}')">
                <div class="file-info">
                    <i class="bi bi-${fileIcon} file-icon"></i>
                    <div class="file-details">
                        <span class="file-name">${filename}</span>
                        <span class="file-type">${fileExtension} File</span>
                    </div>
                </div>
                <i class="bi bi-eye preview-icon"></i>
            </li>
        `;
    });
    
    resultHTML += `</ul>`;
    
    // Update result display
    resultDisplay.innerHTML = resultHTML;
    resultDisplay.style.display = "block";

    // Update download button
    const downloadBtnText = document.getElementById('downloadBtnText');
    if (downloadBtnText) {
        downloadBtnText.innerHTML = `Download ${fileList.length > 1 ? 'Files' : 'File'} <small class="download-meta">${exportSettings.format.toUpperCase()}</small>`;
    }
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
        <i class="bi bi-table"></i>
        Table View
      </button>
    `;
    previewTabs.appendChild(tableTab);
  
    // Create map view tab
    const mapTab = document.createElement('li');
    mapTab.className = 'nav-item';
    mapTab.innerHTML = `
      <button class="nav-link" data-bs-toggle="tab" data-bs-target="#mapView" type="button">
        <i class="bi bi-geo-alt"></i>
        Map View
      </button>
    `;
    previewTabs.appendChild(mapTab);
  
    // Create table view content
    const tableContent = document.createElement('div');
    tableContent.className = 'tab-pane fade show active animate__animated animate__fadeIn';
    tableContent.id = 'tableView';
    
    let tableHTML = `
        <div class="preview-header mb-3">
          <h6 class="preview-title">
            <i class="bi bi-file-earmark-text me-2"></i>
            ${path}
          </h6>
          <span class="preview-count">
            <i class="bi bi-list-ul me-1"></i>
            ${data.length} records
          </span>
        </div>
      <div class="table-responsive">
        
        <table class="preview-table">
          <thead>
            <tr>
    `;
    const headers = Object.keys(data[0]);
    headers.forEach(header => {
      tableHTML += `<th>${header}</th>`;
    });
    tableHTML += '</tr></thead><tbody>';
    
    data.forEach((row, index) => {
      tableHTML += `<tr class="animate__animated animate__fadeIn" style="animation-delay: ${index * 50}ms">`;
      headers.forEach(header => {
        tableHTML += `<td>${row[header]}</td>`;
      });
      tableHTML += '</tr>';
    });
    tableHTML += `
          </tbody>
        </table>
      </div>
    `;
    
    tableContent.innerHTML = tableHTML;
    previewTabContent.appendChild(tableContent);
  
    // Create map view content
    const mapContent = document.createElement('div');
    mapContent.className = 'tab-pane fade animate__animated animate__fadeIn';
    mapContent.id = 'mapView';
    
    mapContent.innerHTML = `
      <div class="map-header mb-3">
        <h6 class="preview-title">
          <i class="bi bi-geo me-2"></i>
          Location Preview
        </h6>
        <span class="preview-count">
          <i class="bi bi-geo-alt me-1"></i>
          ${data.length} points
        </span>
      </div>
      <div id="map"></div>
    `;
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
        zoom: hasValidCoordinates ? 13 : 5
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
            // Create custom popup content
            const popupContent = `
              <div class="popup-content">
                <h4>${point.Name || 'Unnamed Location'}</h4>
                <div class="coordinates">
                  <div><i class="bi bi-geo-alt"></i> Lat: ${lat}</div>
                  <div><i class="bi bi-geo"></i> Lng: ${lng}</div>
                </div>
              </div>
            `;

            // Create marker with popup
            const marker = L.marker([lat, lng])
              .bindPopup(popupContent)
              .addTo(map);
            
            bounds.extend([lat, lng]);
            hasMarkers = true;
          }
        }
      });
  
      // Fit map to bounds if we have markers
      if (hasMarkers) {
        map.fitBounds(bounds, { 
          padding: [50, 50],
          maxZoom: 15
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
    }, { once: true });
}

// Download function
function downloadConvertedFiles() {
    if (convertedZip) {
        const url = window.URL.createObjectURL(convertedZip);
        const a = document.createElement("a");
        a.href = url;
        
        // Dapatkan nama file yang diinput
        const customFileName = document.getElementById('outputFilename').value;
        const baseFileName = customFileName || 'converted_data';
        
        if (exportSettings.mode === 'merged') {
            // Untuk mode merged, tetap gunakan ZIP
            const zipFileName = `${baseFileName}.zip`;
            a.download = zipFileName;
        } else {
            // Untuk mode separate
            a.download = `${baseFileName}.zip`;
        }
        
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
    }
}