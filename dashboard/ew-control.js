const ewRow = document.querySelector('#anti-ew-rows .anti-ew-row');
if (ewRow) {
  document.getElementById('add-anti-ew')?.classList.add('hidden');
  const heading = ewRow.closest('.multi-entry').querySelector('.multi-heading');
  heading.textContent = 'EW resilience';
  const mode = document.createElement('select');
  mode.id = 'ew-mode';
  mode.className = 'ew-mode';
  mode.innerHTML = '<option value="none">Non-EW resilient</option><option value="resilient">EW resilient</option>';
  const technology = ewRow.querySelector('[data-multi-new="anti-ew"]');
  const valueSelect = ewRow.querySelector('select[data-multi="anti-ew"]');
  valueSelect.classList.add('hidden');
  ewRow.querySelector('.remove-row').classList.add('hidden');
  technology.placeholder = 'Enter EW resilience technology';

  function updateEwMode() {
    const resilient = mode.value === 'resilient';
    valueSelect.value = resilient ? '__new__' : 'Nil';
    technology.value = resilient ? technology.value : '';
    technology.classList.toggle('hidden', !resilient);
    technology.required = resilient;
  }

  mode.addEventListener('change', updateEwMode);
  ewRow.parentElement.insertBefore(mode, ewRow);
  updateEwMode();
}
