const auth = { role: '', password: '' };
const originalFetch = window.fetch.bind(window);
window.fetch = (url, init = {}) => {
  if (String(url).startsWith('/api/')) init.headers = { ...(init.headers || {}), 'X-Role': auth.role, 'X-Password': auth.password };
  return originalFetch(url, init).then(async (response) => {
    return response;
  });
};
function applyRole() {
  const admin = auth.role === 'admin';
  window.userRole = auth.role;
  window.userRole = auth.role;
  document.getElementById('login-screen').classList.toggle('hidden', Boolean(auth.role));
  document.querySelector('.app-shell').classList.toggle('hidden', !auth.role);
  document.querySelectorAll('.admin-only').forEach((item) => item.classList.toggle('hidden', !admin));
  document.getElementById('role-badge').textContent = admin ? 'ADMIN VIEW' : 'USER VIEW';
  document.getElementById('view-label').textContent = admin ? 'ADMIN VIEW' : 'USER VIEW';
  document.getElementById('overview-view').classList.toggle('hidden', !admin);
  document.getElementById('add-asset-view').classList.toggle('hidden', admin);
  if (admin) switchView('overview'); else switchView('add-asset');
  renderFleet(document.getElementById('fleet-search').value);
}
document.getElementById('login-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const role = document.getElementById('login-role').value;
  const password = document.getElementById('login-password').value;
  const message = document.getElementById('login-message');
  message.textContent = '';
  try {
    const response = await originalFetch('/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }, body: JSON.stringify({ role, password }) });
    if (!response.ok) { message.textContent = 'Invalid role or password.'; return; }
    auth.role = role; auth.password = password; applyRole();
  } catch (error) {
    message.textContent = 'Unable to connect to the KUIN-G server. Restart KUIN-G.exe.';
  }
});
document.getElementById('logout-button').addEventListener('click', () => {
  auth.role = ''; auth.password = '';
  window.location.replace(`/index.html?session=${Date.now()}`);
});
document.getElementById('password-button').addEventListener('click', () => document.getElementById('password-dialog').showModal());
document.getElementById('password-cancel').addEventListener('click', () => document.getElementById('password-dialog').close());
document.getElementById('password-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const newPassword = document.getElementById('new-password').value;
  const confirmPassword = document.getElementById('confirm-password').value;
  const message = document.getElementById('password-message');
  if (newPassword !== confirmPassword) { message.textContent = 'Passwords do not match.'; return; }
  const response = await fetch('/api/password', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ new_password: newPassword }) });
  if (!response.ok) { message.textContent = (await response.json()).error; return; }
  auth.password = newPassword;
  message.textContent = 'Password updated.';
  setTimeout(() => document.getElementById('password-dialog').close(), 500);
});
applyRole();