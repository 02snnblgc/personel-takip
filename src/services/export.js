import jsPDF from 'jspdf';
import 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { db } from '../db/database';

// Türkçe karakter için Roboto fontu base64 olarak import edilmeli.
// fonts/roboto.js dosyasına fontconverter ile üretilmiş base64 içeriği koyun.
import { robotoBase64 } from './fonts/roboto';

async function getReportData(startDate, endDate) {
  const records = await db.records
    .where('date').between(startDate, endDate, true, true)
    .toArray();
  const personnel = await db.personnel.toArray();
  const pMap = Object.fromEntries(personnel.map(p => [p.id, p]));
  const leaves = await db.leaves.toArray();

  return records.map(r => {
    const p = pMap[r.personnelId] || {};
    const leave = leaves.find(l =>
      l.personnelId === r.personnelId && l.startDate <= r.date && l.endDate >= r.date
    );
    return {
      name: p.name || '-',
      sicil: p.sicil || '-',
      department: p.department || '-',
      date: r.date,
      checkIn: r.checkIn || '-',
      checkOut: r.checkOut || '-',
      status: leave ? 'İzinli' :
              r.status === 'open' ? 'Açık' :
              r.status === 'early' ? 'Erken' :
              r.status === 'late' ? 'Geç' : 'Tamam'
    };
  });
}

export async function exportPDF(startDate, endDate) {
  const data = await getReportData(startDate, endDate);
  const doc = new jsPDF();

  // Türkçe karakter desteği
  doc.addFileToVFS('Roboto-Regular.ttf', robotoBase64);
  doc.addFont('Roboto-Regular.ttf', 'Roboto', 'normal');
  doc.setFont('Roboto');

  doc.setFontSize(14);
  doc.text('Personel Devam Raporu', 14, 20);
  doc.setFontSize(9);
  doc.text(`Tarih: ${startDate} — ${endDate}`, 14, 28);

  doc.autoTable({
    head: [['Ad Soyad', 'Sicil', 'Departman', 'Tarih', 'Giriş', 'Çıkış', 'Durum']],
    body: data.map(d => [d.name, d.sicil, d.department, d.date, d.checkIn, d.checkOut, d.status]),
    startY: 34,
    styles: { font: 'Roboto', fontSize: 8 },
    headStyles: { fillColor: [168, 213, 186] }
  });

  doc.save(`rapor_${startDate}_${endDate}.pdf`);
}

export async function exportExcel(startDate, endDate) {
  const data = await getReportData(startDate, endDate);
  const rows = data.map(d => ({
    'Ad Soyad': d.name,
    'Sicil': d.sicil,
    'Departman': d.department,
    'Tarih': d.date,
    'Giriş': d.checkIn,
    'Çıkış': d.checkOut,
    'Durum': d.status
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0] || {}).map(k => ({ wch: 16 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Devam');
  XLSX.writeFile(wb, `rapor_${startDate}_${endDate}.xlsx`);
}