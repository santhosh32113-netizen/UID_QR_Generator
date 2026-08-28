const keyFields = new Set(['Drone ID', 'Ser No', 'Command', 'Corps', 'Brigade', 'Unit', 'Drone Name', 'Type']);
let editingRecord = null;
function renderFleetWithEdit(query = '') {
  const filters = [['Command', 'command-filter'], ['Corps', 'corps-filter'], ['Division', 'division-filter'], ['Brigade', 'brigade-filter']];
  const rows = fleetData.filter((row) => Object.values(row).join(' ').toLowerCase().includes(query.toLowerCase()) && filters.every(([field, id]) => {
    const selected = document.getElementById(id)?.value || 'all';
    return selected === 'all' || row[field] === selected;
  }));
  document.getElementById('result-count').textContent = `Showing ${rows.length} record${rows.length === 1 ? '' : 's'}`;
  document.getElementById('fleet-table').innerHTML = '<div class="table-head"><span>DRONE ID</span><span>DRONE NAME</span><span>UNIT</span><span>TYPE</span><span>STATUS</span><span></span></div>' + rows.map((row) => `<div class="table-row"><strong>${row['Drone ID']}</strong><span>${row['Drone Name']}</span><small>${row.Unit}</small><span>${row.Type}</span><span class="status-badge ${row.Serv === 'Svc' || row.Serv === 'Ser' ? '' : 'warn'}">${row.Serv}</span>${window.userRole === 'admin' ? `<button class="details-button" data-details-id="${row['Drone ID']}" title="View all fields" aria-label="View all fields">View</button><button class="edit-button" data-edit-id="${row['Drone ID']}" title="Edit non-key fields" aria-label="Edit asset">Edit</button><button class="delete-button" data-delete-id="${row['Drone ID']}" title="Delete asset" aria-label="Delete asset">×</button>` : ''}</div>`).join('');
  document.querySelectorAll('[data-details-id]').forEach((button) => button.addEventListener('click', () => openDetails(button.dataset.detailsId)));
  document.querySelectorAll('[data-edit-id]').forEach((button) => button.addEventListener('click', () => openEdit(button.dataset.editId)));
  document.querySelectorAll('[data-delete-id]').forEach((button) => button.addEventListener('click', deleteAsset));
}
function openEdit(droneId) {
  editingRecord = fleetData.find((row) => row['Drone ID'] === droneId);
  if (!editingRecord) return;
  document.getElementById('edit-drone-id').textContent = droneId;
  document.getElementById('edit-form').querySelectorAll('[name]').forEach((input) => { input.value = editingRecord[input.name] || ''; });
  document.getElementById('edit-dialog').showModal();
}
function openDetails(droneId) {
  const record = fleetData.find((row) => row['Drone ID'] === droneId);
  if (!record) return;
  const content = document.getElementById('details-content');
  content.replaceChildren();
  Object.entries(record).forEach(([field, value]) => {
    const label = document.createElement('dt');
    label.textContent = field;
    const detail = document.createElement('dd');
    detail.textContent = value == null || value === '' ? '-' : String(value);
    content.append(label, detail);
  });
  document.getElementById('details-dialog').showModal();
}
document.getElementById('edit-cancel').addEventListener('click', () => document.getElementById('edit-dialog').close());
document.getElementById('edit-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const changes = Object.fromEntries(new FormData(event.currentTarget));
  if (changes.Serv === 'Svc') changes.Serv = 'Ser';
  if (changes.Serv === 'Unsvc') changes.Serv = 'Unser';
  const response = await fetch('/api/assets/edit', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ 'Drone ID': editingRecord['Drone ID'], changes }) });
  if (!response.ok) { document.getElementById('edit-message').textContent = (await response.json()).error; return; }
  if (changes.Serv === 'Ser') changes.Serv = 'Svc';
  if (changes.Serv === 'Unser') changes.Serv = 'Unsvc';
  Object.assign(editingRecord, changes);
  document.getElementById('edit-dialog').close();
  renderOverview(); renderService(); renderGuidance(); renderEWProfiles(); renderProcFund(); renderFleetWithEdit(document.getElementById('fleet-search').value);
});
const originalRenderFleet = renderFleet;
renderFleet = renderFleetWithEdit;
renderProcFund();
renderFleetWithEdit();