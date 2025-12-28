document.addEventListener('DOMContentLoaded', () => {
    fetchData();
});

async function fetchData() {
    try {
        const response = await fetch('/api/data');
        const data = await response.json();

        if (data.error) {
            throw new Error(data.error);
        }

        renderTable(data);
    } catch (error) {
        console.error('Error:', error);
        document.getElementById('loading').textContent = "Failed to load data. " + error;
        document.getElementById('loading').style.color = '#ef4444';
    }
}

function renderTable(data) {
    const container = document.getElementById('table-container');
    const loading = document.getElementById('loading');
    loading.style.display = 'none';

    if (!data.records || data.records.length === 0) {
        container.innerHTML = "<div style='padding:1rem; text-align:center'>No records found or Excel is empty.</div>";
        return;
    }

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');

    // Header Row 1: Static + Months
    const tr1 = document.createElement('tr');

    // Static Headers
    // Note: 'Rating' column contains Industry data for Quant
    ['Name', 'ISIN', 'Industry/Rating'].forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        th.rowSpan = 2; // Merge vertical
        th.style.verticalAlign = "middle";
        tr1.appendChild(th);
    });

    // Month Headers
    data.months.forEach(m => {
        const th = document.createElement('th');
        th.textContent = m;
        th.colSpan = 3; // Qty, Val, Pct
        th.className = 'month-header';
        tr1.appendChild(th);
    });
    thead.appendChild(tr1);

    // Header Row 2: Sub-headers
    const tr2 = document.createElement('tr');
    data.months.forEach(() => {
        ['Qty', 'Val (L)', '% Net'].forEach(sub => {
            const th = document.createElement('th');
            th.textContent = sub;
            th.className = 'sub-header d-num';
            tr2.appendChild(th);
        });
    });
    thead.appendChild(tr2);

    // Data Rows
    data.records.forEach(rec => {
        const tr = document.createElement('tr');
        // Check for Total row
        if (rec.Name === 'Total') {
            tr.classList.add('total-row');
        }

        // Static Info
        [rec.Name, rec.ISIN, rec.Rating].forEach(val => {
            const td = document.createElement('td');
            td.textContent = val;
            tr.appendChild(td);
        });

        // Dynamic Data
        data.months.forEach(m => {
            const mData = rec.Months[m] || { Quantity: 0, Value: 0, Pct: 0 };

            // Quantity
            const tdQty = document.createElement('td');
            tdQty.textContent = formatIndianNumber(mData.Quantity, 0);
            tdQty.className = 'd-num';
            tr.appendChild(tdQty);

            // Value
            const tdVal = document.createElement('td');
            tdVal.textContent = formatIndianNumber(mData.Value, 2);
            tdVal.className = 'd-num';
            tr.appendChild(tdVal);

            // Pct
            const tdPct = document.createElement('td');
            // Data comes as float (0.1034), display as 10.34%
            tdPct.textContent = (mData.Pct * 100).toFixed(2) + '%';
            tdPct.className = 'd-num d-pct';
            tr.appendChild(tdPct);
        });

        tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    container.appendChild(table);
}

function formatIndianNumber(num, decimals) {
    if (num === null || num === undefined) return '0';
    let n = parseFloat(num);
    if (isNaN(n)) return '0';

    // Indian locale
    return n.toLocaleString('en-IN', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
    });
}
