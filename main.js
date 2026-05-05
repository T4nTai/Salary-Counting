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
	const parts = String(dateStr).split('/');
	if (parts.length !== 3) return null;
	const dd = parseInt(parts[0], 10);
	const mm = parseInt(parts[1], 10);
	const yy = parseInt(parts[2], 10);
	const yyyy = yy > 30 ? 1900 + yy : 2000 + yy;
	const d = new Date(yyyy, mm - 1, dd);
	if (d.getFullYear() !== yyyy || d.getMonth() !== mm - 1 || d.getDate() !== dd) return null;
	return `${String(yyyy).padStart(4, '0')}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`;
}

function convertExcelDate(excelNum) {
	const excelEpoch = new Date(1900, 0, 1);
	const date = new Date(excelEpoch.getTime() + (excelNum - 1) * 24 * 60 * 60 * 1000);
	const dd = String(date.getDate()).padStart(2, '0');
	const mm = String(date.getMonth() + 1).padStart(2, '0');
	const yy = String(date.getFullYear()).slice(-2);
	return `${dd}/${mm}/${yy}`;
}

function listDates(start, end) {
	const dates = [];
	const cur = new Date(start.getTime());
	while (cur <= end) {
		dates.push(fmtDate(cur));
		cur.setDate(cur.getDate() + 1);
	}
	return dates;
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
	const numCols = 15;
	const emptyCols = Array(numCols).fill('');
	return [['Ngày', ...emptyCols], ['Ca', ...emptyCols], ['Ngày lễ', ...emptyCols]];
}

function getPayrollPeriod(yearMonth) {
	if (!yearMonth) return null;
	const [year, month] = yearMonth.split('-').map(Number);
	let startMonth = month - 1;
	let startYear = year;
	if (startMonth === 0) {
		startMonth = 12;
		startYear = year - 1;
	}
	return {
		startDate: new Date(startYear, startMonth - 1, 21),
		endDate: new Date(year, month - 1, 20),
	};
}

function generatePayrollTemplate(yearMonth) {
	const period = getPayrollPeriod(yearMonth);
	if (!period) return generateExcelMatrixTemplate();
	const dates = [];
	const cur = new Date(period.startDate.getTime());
	while (cur <= period.endDate) {
		dates.push(new Date(cur));
		cur.setDate(cur.getDate() + 1);
	}
	return [
		['Ngày', ...dates.map(d => fmtDateDDMMYY(d))],
		['Ca', ...Array(dates.length).fill('')],
		['Ngày lễ', ...Array(dates.length).fill('')],
	];
}

function parseTimeParts(timeStr) {
	const parts = String(timeStr || '').trim().split(':');
	if (parts.length < 2) return null;
	const hours = parseInt(parts[0], 10);
	const minutes = parseInt(parts[1], 10);
	const seconds = parts.length >= 3 ? parseInt(parts[2], 10) : 0;
	if ([hours, minutes, seconds].some(n => Number.isNaN(n))) return null;
	if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59 || seconds < 0 || seconds > 59) return null;
	return { hours, minutes, seconds };
}

function buildDateTime(dateISO, timeStr) {
	const parts = parseTimeParts(timeStr);
	if (!dateISO || !parts) return null;
	const dt = new Date(`${dateISO}T00:00:00`);
	dt.setHours(parts.hours, parts.minutes, parts.seconds, 0);
	return dt;
}

function overlapHours(startA, endA, startB, endB) {
	const start = Math.max(startA.getTime(), startB.getTime());
	const end = Math.min(endA.getTime(), endB.getTime());
	return Math.max(0, (end - start) / 3600000);
}

function computeShiftMetrics(dateISO, start, end) {
	if (!dateISO || !start || !end) return { hours: 0, nightHours: 0 };
	const dtStart = buildDateTime(dateISO, start);
	let dtEnd = buildDateTime(dateISO, end);
	if (!dtStart || !dtEnd) return { hours: 0, nightHours: 0 };
	if (dtEnd <= dtStart) dtEnd.setDate(dtEnd.getDate() + 1);
	const hours = (dtEnd - dtStart) / 3600000;
	const night1Start = buildDateTime(dateISO, '21:30:00');
	const night1End = new Date(`${dateISO}T23:59:59.999`);
	const nextDate = new Date(`${dateISO}T00:00:00`);
	nextDate.setDate(nextDate.getDate() + 1);
	const nextDateISO = fmtDate(nextDate);
	const night2Start = new Date(`${nextDateISO}T00:00:00`);
	const night2End = buildDateTime(nextDateISO, '06:00:00');
	const nightHours = overlapHours(dtStart, dtEnd, night1Start, night1End) + overlapHours(dtStart, dtEnd, night2Start, night2End);
	return { hours: parseFloat(hours.toFixed(2)), nightHours: parseFloat(nightHours.toFixed(2)) };
}

function parseCSV(text) {
	const lines = text.trim().split(/\r?\n/).map(l => l.trim()).filter(Boolean);
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
		data.push({
			date: (cols[dateIdx] || '').trim(),
			hours: parseFloat((cols[hoursIdx] || '').trim()) || 0,
			nightHours: parseFloat((cols[nightIdx] || '').trim()) || 0,
			isHoliday: ((cols[holIdx] || '').trim() === '1') ? 1 : 0,
		});
	}
	return data;
}

function parseXLSX(arrayBuffer) {
	const wb = XLSX.read(arrayBuffer, { type: 'array' });
	const first = wb.SheetNames[0];
	const ws = wb.Sheets[first];
	const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
	if (raw.length >= 2 && String(raw[0][0]).toLowerCase().trim() === 'ngày' && String(raw[1][0]).toLowerCase().trim() === 'ca') {
		const rows = [];
		const firstRow = raw[0];
		const secondRow = raw[1];
		const thirdRow = raw[2] || [];
		for (let c = 1; c < firstRow.length; c++) {
			let dateRaw = firstRow[c];
			if (typeof dateRaw === 'number') dateRaw = convertExcelDate(dateRaw);
			const dateISO = parseDateDDMMYY(dateRaw);
			if (!dateISO) continue;
			const shift = secondRow[c] || '';
			let start = '';
			let end = '';
			if (shift && typeof shift === 'string' && shift.includes('-')) {
				const parts = shift.split('-').map(s => s.trim());
				start = parts[0];
				end = parts[1] || '';
			}
			const isHol = Number(thirdRow[c] || 0) ? 1 : 0;
			const metrics = computeShiftMetrics(dateISO, start, end);
			rows.push({ date: dateISO, hours: metrics.hours, nightHours: metrics.nightHours, isHoliday: isHol });
		}
		return rows;
	}
	const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
	const rows = [];
	json.forEach(r => {
		const date = r['Date'] || r['date'] || '';
		const start = r['StartTime'] || r['starttime'] || '';
		const end = r['EndTime'] || r['endtime'] || '';
		const isHol = (r['IsHoliday'] || r['isholiday'] || 0) ? 1 : 0;
		if (!date || !start || !end) return;
		const metrics = computeShiftMetrics(date, start, end);
		rows.push({ date, hours: metrics.hours, nightHours: metrics.nightHours, isHoliday: isHol });
	});
	return rows;
}

function computePayroll(rows, hourlyRate) {
	const detailed = [];
	let total = 0;
	rows.forEach(r => {
		const totalHours = Number(r.hours) || 0;
		const rawNightHours = Number(r.nightHours) || 0;
		const nightHours = Math.ceil(rawNightHours);
		const isHol = Number(r.isHoliday) || 0;
		let dayHours = Math.max(0, totalHours - rawNightHours);
		dayHours = Math.max(0, dayHours - 0.5);
		if (isHol) {
			const totalHoursPaid = dayHours + nightHours;
			const basePay = totalHoursPaid * hourlyRate * 3;
			detailed.push({ date: r.date, totalHours: parseFloat(totalHours.toFixed(2)), dayHours: parseFloat(dayHours.toFixed(2)), nightHours, isHoliday: 1, basePay, nightExtra: 0, dayTotal: basePay });
			total += basePay;
			return;
		}
		const totalHoursPaid = dayHours + nightHours;
		const basePay = totalHoursPaid * hourlyRate;
		const nightExtra = nightHours * hourlyRate * 0.3;
		const dayTotal = basePay + nightExtra;
		detailed.push({ date: r.date, totalHours: parseFloat(totalHours.toFixed(2)), dayHours: parseFloat(dayHours.toFixed(2)), nightHours, isHoliday: 0, basePay, nightExtra, dayTotal });
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
	thead.innerHTML = '<tr><th>Ngày</th><th>Giờ ngày</th><th>Giờ đêm</th><th>Ngày lễ</th><th>Lương cơ bản</th><th>Phụ cấp đêm</th><th>Tổng ngày</th></tr>';
	table.appendChild(thead);
	const tbody = document.createElement('tbody');
	payroll.detailed.forEach(d => {
		const tr = document.createElement('tr');
		tr.innerHTML = `<td>${d.date}</td><td>${d.dayHours}</td><td>${d.nightHours || ''}</td><td>${d.isHoliday ? 'Có' : ''}</td><td>${Math.round(d.basePay).toLocaleString()}</td><td>${Math.round(d.nightExtra || 0).toLocaleString()}</td><td>${Math.round(d.dayTotal).toLocaleString()}</td>`;
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
		const rows = ['Date,DayHours,NightHours,IsHoliday,BasePay,NightPay,TotalPay'];
		payroll.detailed.forEach(d => rows.push(`${d.date},${d.dayHours},${d.nightHours || 0},${d.isHoliday ? 1 : 0},${Math.round(d.basePay)},${Math.round(d.nightExtra || 0)},${Math.round(d.dayTotal)}`));
		rows.push(`,,,,,,Tổng,${Math.round(payroll.total)}`);
		downloadCSV('payroll_results.csv', rows.join('\n'));
	});
	area.appendChild(btn);
}

let activeInputMode = 'file';
let importedRows = [];

function setSummary(text) {
	const summary = document.getElementById('summary');
	if (summary) summary.textContent = text;
}

function getPayrollDatesForMonth(yearMonth) {
	const period = getPayrollPeriod(yearMonth);
	if (!period) return [];
	const dates = [];
	const cur = new Date(period.startDate.getTime());
	while (cur <= period.endDate) {
		dates.push(fmtDate(cur));
		cur.setDate(cur.getDate() + 1);
	}
	return dates;
}

function updateDirectRowPreview(tr) {
	if (!tr) return;
	const date = tr.querySelector('[data-field="date"]')?.value || '';
	const start = tr.querySelector('[data-field="start"]')?.value || '';
	const end = tr.querySelector('[data-field="end"]')?.value || '';
	const preview = tr.querySelector('[data-field="preview"]');
	const metrics = computeShiftMetrics(date, start, end);
	if (preview) preview.textContent = metrics ? `${metrics.hours.toFixed(2)} / ${metrics.nightHours.toFixed(2)}` : '';
}

function createDirectRow(row = {}) {
	const tr = document.createElement('tr');
	tr.innerHTML = `
		<td><input type="date" data-field="date" value="${row.date || ''}"></td>
		<td><input type="time" data-field="start" value="${row.start || ''}"></td>
		<td><input type="time" data-field="end" value="${row.end || ''}"></td>
		<td>
			<select data-field="holiday">
				<option value="0" ${Number(row.isHoliday) ? '' : 'selected'}>Không</option>
				<option value="1" ${Number(row.isHoliday) ? 'selected' : ''}>Có</option>
			</select>
		</td>
		<td data-field="preview" class="preview-cell"></td>
		<td><button type="button" class="ghost-button remove-row">Xóa</button></td>
	`;
	const update = () => updateDirectRowPreview(tr);
	tr.querySelectorAll('input, select').forEach(el => {
		el.addEventListener('input', update);
		el.addEventListener('change', update);
	});
	tr.querySelector('.remove-row')?.addEventListener('click', () => tr.remove());
	updateDirectRowPreview(tr);
	return tr;
}

function renderDirectSheet() {
	const body = document.getElementById('directSheetBody');
	if (!body) return;
	body.innerHTML = '';
	const dates = getPayrollDatesForMonth(document.getElementById('payrollMonth').value);
	const rows = dates.length ? dates.map(date => ({ date, start: '', end: '', isHoliday: 0 })) : Array.from({ length: 15 }, () => ({ date: '', start: '', end: '', isHoliday: 0 }));
	rows.forEach(row => body.appendChild(createDirectRow(row)));
	setSummary(dates.length ? `Đã tạo bảng nhập trực tiếp cho ${dates.length} ngày của kỳ lương.` : 'Chọn tháng lương rồi nhấn "Tạo bảng" để sinh bảng nhập trực tiếp.');
}

function collectDirectRows() {
	const rows = [];
	document.querySelectorAll('#directSheetBody tr').forEach(tr => {
		const date = tr.querySelector('[data-field="date"]')?.value || '';
		const start = tr.querySelector('[data-field="start"]')?.value || '';
		const end = tr.querySelector('[data-field="end"]')?.value || '';
		const isHoliday = Number(tr.querySelector('[data-field="holiday"]')?.value || 0) ? 1 : 0;
		const metrics = computeShiftMetrics(date, start, end);
		rows.push({
			date,
			hours: metrics ? metrics.hours : 0,
			nightHours: metrics ? metrics.nightHours : 0,
			isHoliday,
		});
	});
	return rows;
}

function setActiveMode(mode) {
	activeInputMode = mode;
	const filePanel = document.getElementById('filePanel');
	const sheetPanel = document.getElementById('sheetPanel');
	const fileBtn = document.getElementById('fileModeBtn');
	const sheetBtn = document.getElementById('sheetModeBtn');
	if (filePanel) filePanel.classList.toggle('hidden', mode !== 'file');
	if (sheetPanel) sheetPanel.classList.toggle('hidden', mode !== 'sheet');
	if (fileBtn) fileBtn.classList.toggle('active', mode === 'file');
	if (sheetBtn) sheetBtn.classList.toggle('active', mode === 'sheet');
	if (mode === 'sheet') renderDirectSheet();
}

document.getElementById('fileModeBtn').addEventListener('click', () => setActiveMode('file'));
document.getElementById('sheetModeBtn').addEventListener('click', () => setActiveMode('sheet'));
document.getElementById('buildDirectSheet').addEventListener('click', () => {
	setActiveMode('sheet');
	renderDirectSheet();
});
document.getElementById('addDirectRow').addEventListener('click', () => {
	setActiveMode('sheet');
	const body = document.getElementById('directSheetBody');
	if (body) body.appendChild(createDirectRow());
});
document.getElementById('resetDirectSheet').addEventListener('click', () => renderDirectSheet());

document.getElementById('payrollMonth').addEventListener('change', () => {
	if (activeInputMode === 'sheet') renderDirectSheet();
});

document.getElementById('exportExcelMatrix').addEventListener('click', () => {
	const yearMonth = document.getElementById('payrollMonth').value;
	let aoa;
	if (yearMonth) {
		aoa = generatePayrollTemplate(yearMonth);
		const period = getPayrollPeriod(yearMonth);
		alert(`✅ Template tháng lương từ ${fmtDateDDMMYY(period.startDate)} đến ${fmtDateDDMMYY(period.endDate)}`);
	} else {
		aoa = generateExcelMatrixTemplate();
		alert('⚠️ Chưa chọn tháng. Template 15 cột trống');
	}
	downloadXLSX('timesheet_template.xlsx', 'Timesheet', aoa);
});

document.getElementById('importFile').addEventListener('change', (ev) => {
	const f = ev.target.files[0];
	if (!f) return;
	const name = f.name.toLowerCase();
	if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
		const reader = new FileReader();
		reader.onload = function(e) {
			try {
				const rows = parseXLSX(e.target.result);
				if (!rows || rows.length === 0) {
					alert('❌ Không đọc được file Excel.\n\nNguyên nhân có thể:\n• File không có dữ liệu\n• Hàng 1 không chứa "Ngày"\n• Hàng 2 không chứa "Ca"\n• Ngày không đúng format dd/mm/yy\n\nMở browser console (F12) để xem chi tiết.');
					return;
				}
				importedRows = rows;
				setSummary(`✅ Đã nhập ${rows.length} dòng (Excel).`);
			} catch (err) {
				console.error('Excel parse error:', err);
				alert('❌ Lỗi khi đọc Excel:\n' + err.message + '\n\nMở browser console (F12) để xem chi tiết.');
			}
		};
		reader.readAsArrayBuffer(f);
	} else {
		const reader = new FileReader();
		reader.onload = function(e) {
			const rows = parseCSV(e.target.result);
			if (rows.length === 0) {
				alert('❌ Không đọc được CSV.\n\nCần có header: Date,Hours\n\nHoặc: Date,Hours,NightHours,IsHoliday');
				return;
			}
			importedRows = rows;
			setSummary(`✅ Đã nhập ${rows.length} dòng (CSV).`);
		};
		reader.readAsText(f, 'utf-8');
	}
});

document.getElementById('calculate').addEventListener('click', () => {
	const rate = parseFloat(document.getElementById('hourlyRate').value) || 0;
	const rows = activeInputMode === 'sheet' ? collectDirectRows() : importedRows;
	if (!rows || rows.length === 0) {
		alert(activeInputMode === 'sheet' ? 'Chưa có dữ liệu trong bảng nhập trực tiếp' : 'Chưa nhập dữ liệu từ file');
		return;
	}
	if (activeInputMode === 'sheet') {
		importedRows = rows;
		setSummary(`✅ Đã nhập ${rows.length} dòng (nhập trực tiếp).`);
	}
	const payroll = computePayroll(rows, rate);
	renderResults(payroll);
});

window.addEventListener('DOMContentLoaded', () => {
	const today = new Date();
	const year = today.getFullYear();
	const month = String(today.getMonth() + 1).padStart(2, '0');
	const monthInput = document.getElementById('payrollMonth');
	if (monthInput) monthInput.value = `${year}-${month}`;
	setActiveMode('file');
});
