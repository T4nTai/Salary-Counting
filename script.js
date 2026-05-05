function fmtDate(d) {
	const yyyy = d.getFullYear();
	const mm = String(d.getMonth() + 1).padStart(2, '0');
	const dd = String(d.getDate()).padStart(2, '0');
	return `${yyyy}-${mm}-${dd}`;
}

function fmtDateDDMMYY(d) {
	const dd = String(d.getDate()).padStart(2, '0');
	const mm = String(d.getMonth() + 1).padStart(2, '0');
	const yy = String(d.getFullYear()).slice(-2);
	return `${dd}/${mm}/${yy}`;
}

function parseDateDDMMYY(dateStr) {
	// Parse dd/mm/yy format
	const parts = String(dateStr).split('/');
	if (parts.length !== 3) return null;
	let dd = parseInt(parts[0], 10);
	let mm = parseInt(parts[1], 10);
	let yy = parseInt(parts[2], 10);
	// Handle 2-digit year: 00-30 -> 2000-2030, 31-99 -> 1931-1999
	const yyyy = yy > 30 ? 1900 + yy : 2000 + yy;
	const d = new Date(yyyy, mm - 1, dd);
	if (d.getFullYear() !== yyyy || d.getMonth() !== mm - 1 || d.getDate() !== dd) {
		return null; // Invalid date
	}
	const iso = `${String(yyyy).padStart(4, '0')}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`;
	return iso;
}

function listDates(start, end) {
	const a = [];
	const cur = new Date(start.getTime());
	while (cur <= end) {
		a.push(fmtDate(cur));
		cur.setDate(cur.getDate() + 1);
	}
	return a;
}

function downloadCSV(filename, content) {
	const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.href = url;
	a.download = filename;
	document.body.appendChild(a);
	a.click();
	a.remove();
	URL.revokeObjectURL(url);
}

function downloadXLSX(filename, sheetName, aoa) {
	const ws = XLSX.utils.aoa_to_sheet(aoa);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Sheet1');
	const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
	const blob = new Blob([wbout], { type: 'application/octet-stream' });
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.href = url;
	a.download = filename;
	document.body.appendChild(a);
	a.click();
	a.remove();
	URL.revokeObjectURL(url);
}

function generateExcelMatrixTemplate() {
	// Create template with 3 rows and 15 empty columns for user to fill
	const numCols = 15;
	const emptyCols = Array(numCols).fill('');
	const row1 = ['Ngày', ...emptyCols];
	const row2 = ['Ca', ...emptyCols];
	const row3 = ['Ngày lễ', ...emptyCols];
	const aoa = [row1, row2, row3];
	return aoa;
}

function parseCSV(text) {
	const lines = text.trim().split(/\r?\n/).map(l => l.trim()).filter(l => l.length);
	if (lines.length === 0) return [];
	const header = lines[0].split(',').map(h => h.trim().toLowerCase());
	const dateIdx = header.indexOf('date');
	const hoursIdx = header.indexOf('hours');
	const nightIdx = header.indexOf('nighthours');
	const holIdx = header.indexOf('isholiday');
	if (dateIdx === -1 || hoursIdx === -1) return [];
	const data = [];
	for (let i = 1; i < lines.length; i++) {
		const cols = lines[i].split(',');
		const date = cols[dateIdx].trim();
		const hours = parseFloat((cols[hoursIdx] || '').trim()) || 0;
		const nightHours = parseFloat((cols[nightIdx] || '').trim()) || 0;
		const isHoliday = ((cols[holIdx] || '').trim() === '1') ? 1 : 0;
		data.push({ date, hours, nightHours, isHoliday });
	}
	return data;
}

function parseXLSX(arrayBuffer) {
	const wb = XLSX.read(arrayBuffer, { type: 'array' });
	const first = wb.SheetNames[0];
	const ws = wb.Sheets[first];
	// read raw rows to detect matrix layout
	const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
	// if matrix format: row0 starts with 'Ngày' and row1 starts with 'Ca'
	if (raw.length >= 2 && String(raw[0][0]).toLowerCase().trim() === 'ngày' && String(raw[1][0]).toLowerCase().trim() === 'ca') {
		const rows = [];
		const firstRow = raw[0];
		const secondRow = raw[1];
		const thirdRow = raw[2] || [];
		for (let c = 1; c < firstRow.length; c++) {
			const dateRaw = firstRow[c];
			const shift = secondRow[c] || '';
			const isHol = Number(thirdRow[c] || 0) ? 1 : 0;
			const dateISO = parseDateDDMMYY(dateRaw);
			if (!dateISO) continue; // Skip invalid dates
			const date = dateISO;
			let start = '', end = '';
			if (shift && typeof shift === 'string' && shift.indexOf('-') !== -1) {
				const parts = shift.split('-').map(s => s.trim());
				start = parts[0];
				end = parts[1] || '';
			}
			// compute hours/nightHours same logic as row-based parse
			let hours = 0, nightHours = 0;
			if (date && start && end) {
				const dtStart = new Date(String(date) + 'T' + start);
				let dtEnd = new Date(String(date) + 'T' + end);
				if (dtEnd <= dtStart) dtEnd.setDate(dtEnd.getDate() + 1);
				hours = (dtEnd - dtStart) / 3600000;
				const night1Start = new Date(String(date) + 'T22:00:00');
				const night1End = new Date(String(date) + 'T23:59:59');
				const nextDate = new Date(String(date));
				nextDate.setDate(nextDate.getDate() + 1);
				const nd = nextDate.toISOString().slice(0,10);
				const night2Start = new Date(nd + 'T00:00:00');
				const night2End = new Date(nd + 'T06:00:00');
				function overlap(a,b,c,d){
					const s = Math.max(a.getTime(), c.getTime());
					const e = Math.min(b.getTime(), d.getTime());
					return Math.max(0, (e - s)/3600000);
				}
				nightHours += overlap(dtStart, dtEnd, night1Start, new Date(night1End.getTime()+1000));
				nightHours += overlap(dtStart, dtEnd, night2Start, night2End);
			}
			rows.push({ date: String(dateRaw), hours: parseFloat(hours.toFixed(2)), nightHours: parseFloat(nightHours.toFixed(2)), isHoliday: isHol });
		}
		return rows;
	}
	// otherwise treat as row-based Date/StartTime/EndTime table
	const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
	const rows = [];
	json.forEach(r => {
		const date = r['Date'] || r['date'] || '';
		const start = r['StartTime'] || r['starttime'] || '';
		const end = r['EndTime'] || r['endtime'] || '';
		const isHol = (r['IsHoliday'] || r['isholiday'] || 0) ? 1 : 0;
		let hours = 0, nightHours = 0;
		if (date && start && end) {
			const dtStart = new Date(date + 'T' + start);
			let dtEnd = new Date(date + 'T' + end);
			if (dtEnd <= dtStart) dtEnd.setDate(dtEnd.getDate() + 1);
			hours = (dtEnd - dtStart) / 3600000;
			const night1Start = new Date(date + 'T22:00:00');
			const night1End = new Date(date + 'T23:59:59');
			const nextDate = new Date(date);
			nextDate.setDate(nextDate.getDate() + 1);
			const nd = nextDate.toISOString().slice(0,10);
			const night2Start = new Date(nd + 'T00:00:00');
			const night2End = new Date(nd + 'T06:00:00');
			function overlap(a,b,c,d){
				const s = Math.max(a.getTime(), c.getTime());
				const e = Math.min(b.getTime(), d.getTime());
				return Math.max(0, (e - s)/3600000);
			}
			nightHours += overlap(dtStart, dtEnd, night1Start, new Date(night1End.getTime()+1000));
			nightHours += overlap(dtStart, dtEnd, night2Start, night2End);
		}
		rows.push({ date, hours: parseFloat(hours.toFixed(2)), nightHours: parseFloat(nightHours.toFixed(2)), isHoliday: isHol });
	});
	return rows;
}


function computePayroll(rows, hourlyRate) {
	const detailed = [];
	let total = 0;
	rows.forEach(r => {
		const hours = Number(r.hours) || 0;
		const nightHours = Number(r.nightHours) || 0;
		const isHol = Number(r.isHoliday) || 0;
		if (isHol) {
			const dayTotal = hours * hourlyRate * 3; // holiday 300%
			detailed.push({ date: r.date, hours, nightHours, isHoliday: 1, basePay: 0, nightExtra: 0, dayTotal });
			total += dayTotal;
			return;
		}
		const basePay = hours * hourlyRate;
		const nightExtra = nightHours * hourlyRate * 0.3; // +30% for night hours
		const dayTotal = basePay + nightExtra;
		detailed.push({ date: r.date, hours, nightHours, isHoliday: 0, basePay, nightExtra, dayTotal });
		total += dayTotal;
	});
	return { total, detailed };
}

function renderResults(payroll) {
	const area = document.getElementById('resultsArea');
	area.innerHTML = '';
	if (!payroll || !payroll.detailed) return;
	const table = document.createElement('table');
	const thead = document.createElement('thead');
	thead.innerHTML = '<tr><th>Ngày</th><th>Giờ</th><th>Giờ đêm</th><th>Ngày lễ</th><th>Lương cơ bản</th><th>Phụ cấp đêm</th><th>Tổng ngày</th></tr>';
	table.appendChild(thead);
	const tbody = document.createElement('tbody');
	payroll.detailed.forEach(d => {
		const tr = document.createElement('tr');
		tr.innerHTML = `<td>${d.date}</td><td>${d.hours}</td><td>${d.nightHours || ''}</td><td>${d.isHoliday ? 'Có' : ''}</td><td>${Math.round(d.basePay).toLocaleString()}</td><td>${Math.round(d.nightExtra || 0).toLocaleString()}</td><td>${Math.round(d.dayTotal).toLocaleString()}</td>`;
		tbody.appendChild(tr);
	});
	table.appendChild(tbody);
	area.appendChild(table);
	const tot = document.createElement('div');
	tot.className = 'total';
	tot.textContent = 'Tổng trả: ' + Math.round(payroll.total).toLocaleString() + ' VND';
	area.appendChild(tot);
	const btn = document.createElement('button');
	btn.textContent = 'Tải kết quả CSV';
	btn.addEventListener('click', () => {
		const rows = ['Date,Hours,NightHours,IsHoliday,BasePay,NightPay,TotalPay'];
		payroll.detailed.forEach(d => rows.push(`${d.date},${d.hours},${d.nightHours || 0},${d.isHoliday ? 1 : 0},${Math.round(d.basePay)},${Math.round(d.nightExtra || 0)},${Math.round(d.dayTotal)}`));
		rows.push(`,,,,,,Tổng,${Math.round(payroll.total)}`);
		downloadCSV('payroll_results.csv', rows.join('\n'));
	});
	area.appendChild(btn);
}

let importedRows = [];
document.getElementById('exportExcelMatrix').addEventListener('click', () => {
	const aoa = generateExcelMatrixTemplate();
	downloadXLSX('timesheet_template.xlsx', 'Timesheet', aoa);
});

document.getElementById('importFile').addEventListener('change', (ev) => {
	const f = ev.target.files[0];
	if (!f) return;
	const name = f.name.toLowerCase();
	if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
		const reader = new FileReader();
		reader.onload = function(e) {
			const ab = e.target.result;
			try {
				const rows = parseXLSX(ab);
				if (!rows || rows.length === 0) return alert('Không đọc được file Excel hoặc không có dòng dữ liệu');
				importedRows = rows;
				document.getElementById('summary').textContent = `Đã nhập ${rows.length} dòng (Excel).`;
			} catch (err) {
				console.error(err);
				alert('Lỗi khi đọc Excel');
			}
		};
		reader.readAsArrayBuffer(f);
	} else {
		const reader = new FileReader();
		reader.onload = function(e) {
			const text = e.target.result;
			const rows = parseCSV(text);
			if (rows.length === 0) return alert('Không đọc được CSV. Đảm bảo có header Date,Hours hoặc Date,Hours,NightHours,IsHoliday');
			importedRows = rows;
			document.getElementById('summary').textContent = `Đã nhập ${rows.length} dòng (CSV).`;
		};
		reader.readAsText(f, 'utf-8');
	}
});

document.getElementById('calculate').addEventListener('click', () => {
	const rate = parseFloat(document.getElementById('hourlyRate').value) || 0;
	if (!importedRows || importedRows.length === 0) return alert('Chưa nhập CSV dữ liệu');
	const payroll = computePayroll(importedRows, rate);
	renderResults(payroll);
});

