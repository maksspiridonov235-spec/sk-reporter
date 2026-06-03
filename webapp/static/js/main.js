// ── Drag & Drop ───────────────────────────────────────────────────────────

const uploadZone = document.getElementById('inp-reports').closest('.upload-zone');
uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.style.borderColor = '#1a56a0'; });
uploadZone.addEventListener('dragleave', () => { uploadZone.style.borderColor = '#aac4e8'; });
uploadZone.addEventListener('drop', e => {
  e.preventDefault();
  uploadZone.style.borderColor = '#aac4e8';
  const input = document.getElementById('inp-reports');
  const dt = new DataTransfer();
  for (const f of e.dataTransfer.files) dt.items.add(f);
  input.files = dt.files;
  uploadReports(input);
});
