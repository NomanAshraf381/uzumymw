// --- CONFIGURATION ---

let globalData = [];
let headers = [];

const COL_DESIGNATION = 1;
const COL_NAME = 2;
const COL_BELT = 3;

const DEFAULT_AVATAR = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%23a0aec0'%3E%3Cpath d='M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 3c1.66 0 3 1.34 3 3s-1.34 3-3 3-3-1.34-3-3 1.34-3 3-3zm0 14.2c-2.5 0-4.71-1.28-6-3.22.03-1.99 4-3.08 6-3.08 1.99 0 5.97 1.09 6 3.08-1.29 1.94-3.5 3.22-6 3.22z'/%3E%3C/svg%3E";

// Supported image formats
const IMAGE_FORMATS = ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg'];

// --- DATE FORMATTER ---
function formatValue(value) {
    if (typeof value === 'number' && value > 25000 && value < 60000) {
        const date = new Date(Math.round((value - 25569) * 86400 * 1000));
        return date.toLocaleDateString('en-GB', {
            day: '2-digit',
            month: 'short',
            year: 'numeric'
        });
    }
    return value;
}

// --- LOAD EXCEL ---
window.onload = function () {
    fetch(EXCEL_FILE_PATH)
        .then(response => {
            if (!response.ok) throw new Error("File not found at " + EXCEL_FILE_PATH);
            return response.arrayBuffer();
        })
        .then(data => {
            let workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
            let worksheet = workbook.Sheets[workbook.SheetNames[0]];
            let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

            if (jsonData.length > 0) {
                processExcelData(jsonData);
            }
        })
        .catch(error => {
            document.getElementById('loading-msg').innerHTML =
                `<span style="color:red;">Error: Could not load the Excel file.<br> 
                Make sure the file is named <b>${EXCEL_FILE_PATH}</b> and you are using a local server.</span>`;
            console.error(error);
        });
};

// --- PROCESS DATA ---
function processExcelData(jsonData) {
    let headerRowIndex = 0;
    let maxCols = 0;

    for (let i = 0; i < Math.min(20, jsonData.length); i++) {
        let filledColsCount = (jsonData[i] || []).filter(cell => cell !== "").length;
        if (filledColsCount > maxCols) {
            maxCols = filledColsCount;
            headerRowIndex = i;
        }
    }

    headers = jsonData[headerRowIndex];
    globalData = jsonData.slice(headerRowIndex + 1);

    document.getElementById('loading-msg').style.display = 'none';
    document.getElementById('list-section').style.display = 'block';

    renderList(globalData);
}

// --- LIST VIEW ---
function renderList(dataToRender) {
    let container = document.getElementById('list-container');
    container.innerHTML = '';

    dataToRender.forEach((row) => {
        if (!row || row.length === 0 || (!row[COL_NAME] && !row[COL_BELT])) return;

        let itemDiv = document.createElement('div');
        itemDiv.className = 'list-item';

        // Thumbnail
        let thumb = document.createElement('img');
        thumb.className = 'list-thumb';
        thumb.width = 100;
        thumb.height = 100;
        thumb.loading = 'lazy';
        thumb.alt = '';
        loadPhoto(thumb, row[COL_BELT]);

        // Name + designation
        let info = document.createElement('div');
        info.className = 'item-info';
        info.innerHTML = `
            <div class="item-main">${row[COL_NAME] || 'Unknown'}</div>
            <div class="item-sub">${row[COL_DESIGNATION] || 'N/A'}</div>
        `;

        // Belt badge
        let belt = document.createElement('div');
        belt.className = 'item-belt';
        belt.textContent = `Belt: ${row[COL_BELT] || 'N/A'}`;

        itemDiv.append(thumb, info, belt);
        itemDiv.onclick = () => showDetails(row);
        container.appendChild(itemDiv);
    });
}

// --- SEARCH ---
function filterList() {
    let query = document.getElementById('search').value.toLowerCase();

    let filteredData = globalData.filter(row => {
        let name = (row[COL_NAME] || '').toString().toLowerCase();
        let beltNo = (row[COL_BELT] || '').toString().toLowerCase();
        return name.includes(query) || beltNo.includes(query);
    });

    renderList(filteredData);
}

// --- DETAIL VIEW ---
function showDetails(row) {
    document.getElementById('list-section').style.display = 'none';
    document.getElementById('detail-section').style.display = 'block';

    let beltNo = row[COL_BELT] || 'N/A';

    document.getElementById('detail-name').innerText = row[COL_NAME] || 'Unknown';

    document.getElementById('detail-belt').innerHTML =
        `<strong>Belt No:</strong> ${beltNo} 
         &nbsp;|&nbsp; 
         <strong>Designation:</strong> ${row[COL_DESIGNATION] || 'N/A'}`;

    // --- IMAGE LOADING (MULTI-FORMAT SUPPORT) ---
    let imgElement = document.getElementById('profile-image');
    let basePath = `assets/pictures/${beltNo}`;

    let index = 0;

    function tryNextImage() {
        if (index >= IMAGE_FORMATS.length) {
            imgElement.src = DEFAULT_AVATAR;
            imgElement.onerror = null;
            return;
        }

        imgElement.src = `${basePath}.${IMAGE_FORMATS[index]}`;
        index++;
    }

    imgElement.onerror = tryNextImage;
    tryNextImage();

    // --- DETAILS ---
    let detailContainer = document.getElementById('detail-container');
    detailContainer.innerHTML = '';

    headers.forEach((headerName, i) => {
        let cellValue = row[i];

        if (cellValue !== undefined && cellValue.toString().trim() !== "") {
            let card = document.createElement('div');
            card.className = 'detail-card';

            card.innerHTML = `
                <div class="detail-label">${headerName || 'Field'}</div>
                <div class="detail-value">${formatValue(cellValue)}</div>
            `;

            detailContainer.appendChild(card);
        }
    });
}

// --- BACK TO LIST ---
function showList() {
    document.getElementById('detail-section').style.display = 'none';
    document.getElementById('list-section').style.display = 'block';
}

// --- PHOTO LOADER (tries each format, then falls back to default avatar) ---
function loadPhoto(imgElement, beltNo) {
    let index = 0;

    function tryNext() {
        if (!beltNo || beltNo === 'N/A' || index >= IMAGE_FORMATS.length) {
            imgElement.onerror = null;
            imgElement.src = DEFAULT_AVATAR;
            return;
        }
        imgElement.src = `assets/pictures/${beltNo}.${IMAGE_FORMATS[index]}`;
        index++;
    }

    imgElement.onerror = tryNext;
    tryNext();
}