const state = { date: new Date(), selected: null, events: [] };
const $ = id => document.getElementById(id);
let syncTimer;
let notificationTimer;

function key(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function toast(msg) {
  $('toast').textContent = msg;
  $('toast').classList.remove('hidden');
  setTimeout(() => $('toast').classList.add('hidden'), 2800);
}
function show(id, on = true) { $(id).classList.toggle('hidden', !on); }
function escapeText(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
async function refresh() {
  state.events = (await allEvents()).filter(e => !e.deletedAt);
  render();
  checkScheduledNotifications();
}
function notificationSupported() {
  return 'Notification' in window && 'serviceWorker' in navigator;
}
function updateNotificationUI() {
  const status = $('notificationStatus');
  const button = $('notificationBtn');
  if (!notificationSupported()) {
    status.textContent = '이 브라우저에서는 일정 알림을 지원하지 않습니다.';
    status.className = 'status error';
    button.disabled = true;
    return;
  }
  const permission = Notification.permission;
  status.textContent = permission === 'granted' ? '휴대폰 일정 알림: 켜짐' : permission === 'denied' ? '휴대폰 설정에서 Cellendar 알림을 허용해야 합니다.' : '휴대폰 일정 알림: 꺼짐';
  status.className = `status ${permission === 'granted' ? 'ok' : permission === 'denied' ? 'error' : ''}`;
  button.textContent = permission === 'granted' ? '알림이 켜져 있습니다' : '휴대폰 일정 알림 켜기';
  button.disabled = permission === 'granted';
}
async function enableNotifications() {
  if (!notificationSupported()) return updateNotificationUI();
  const permission = await Notification.requestPermission();
  updateNotificationUI();
  if (permission === 'granted') {
    toast('일정 알림을 켰습니다.');
    await showSystemNotification('Cellendar 알림 설정 완료', '일정의 알림시각이 되면 휴대폰에 알려드립니다.', 'notification-enabled');
    checkScheduledNotifications();
  }
}
async function showSystemNotification(title, body, tag) {
  if (!notificationSupported() || Notification.permission !== 'granted') return;
  const registration = await navigator.serviceWorker.ready;
  registration.active?.postMessage({ type: 'SHOW_NOTIFICATION', title, body, tag });
}
async function checkScheduledNotifications() {
  if (!notificationSupported() || Notification.permission !== 'granted') return;
  const now = new Date();
  const today = key(now);
  const currentMinutes = now.getHours() * 60 + now.getMinutes();
  const notified = JSON.parse(localStorage.getItem('notifiedEvents') || '{}');
  let changed = false;
  for (const event of state.events) {
    if (event.date !== today || !event.alarmTime) continue;
    const [hour, minute] = event.alarmTime.split(':').map(Number);
    const scheduledMinutes = hour * 60 + minute;
    const notificationKey = `${event.id}|${event.date}|${event.alarmTime}`;
    if (currentMinutes >= scheduledMinutes && currentMinutes - scheduledMinutes <= 10 && !notified[notificationKey]) {
      await showSystemNotification(event.title || 'Cellendar 일정', event.memo || `${event.alarmTime} 예정 일정입니다.`, `event-${event.id}`);
      notified[notificationKey] = new Date().toISOString();
      changed = true;
    }
  }
  if (changed) localStorage.setItem('notifiedEvents', JSON.stringify(notified));
}
function render() {
  const y = state.date.getFullYear(), m = state.date.getMonth();
  $('monthTitle').textContent = `${y}년 ${m + 1}월`;
  const first = new Date(y, m, 1);
  const start = new Date(y, m, 1 - first.getDay());
  const today = key(new Date());
  $('calendarGrid').innerHTML = '';
  for (let i = 0; i < 42; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const dk = key(d);
    const events = state.events.filter(e => e.date === dk);
    const cell = document.createElement('button');
    cell.className = `day ${d.getMonth() !== m ? 'other ' : ''}${d.getDay() === 0 ? 'sun ' : d.getDay() === 6 ? 'sat ' : ''}${dk === today ? 'today' : ''}`;
    cell.innerHTML = `<div class="num">${d.getDate()}</div>${events.slice(0, 2).map(e => `<div class="chip" style="background:${e.color}">${escapeText(e.title)}</div>`).join('')}${events.length > 2 ? `<div class="more">+${events.length - 2}</div>` : ''}`;
    cell.onclick = () => openDay(d);
    $('calendarGrid').appendChild(cell);
  }
}
function openDay(d) {
  state.selected = key(d);
  $('dayWeekday').textContent = ['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일'][d.getDay()];
  $('dayTitle').textContent = `${d.getMonth() + 1}월 ${d.getDate()}일`;
  const list = state.events.filter(e => e.date === state.selected);
  $('dayEvents').innerHTML = list.length ? list.map(e => `<div class="event-row"><i class="event-bar" style="background:${e.color}"></i><div class="event-body"><strong>${escapeText(e.title)}</strong>${e.memo ? `<p>${escapeText(e.memo)}</p>` : ''}</div><button data-edit="${e.id}">수정</button></div>`).join('') : '<div class="empty">등록된 일정이 없습니다.</div>';
  $('dayEvents').querySelectorAll('[data-edit]').forEach(b => b.onclick = () => openForm(state.events.find(e => e.id === b.dataset.edit)));
  show('daySheet');
}
function openForm(e = null) {
  $('eventForm').reset();
  $('eventId').value = e?.id || '';
  $('eventDate').value = e?.date || state.selected || key(new Date());
  $('eventTitle').value = e?.title || '';
  $('eventMemo').value = e?.memo || '';
  $('eventColor').value = e?.color || '#4f8ef7';
  $('eventAlarm').value = e?.alarmTime || '';
  $('formTitle').textContent = e ? '일정 수정' : '일정 추가';
  show('deleteBtn', !!e);
  show('eventModal');
}
async function saveForm(ev) {
  ev.preventDefault();
  const old = state.events.find(e => e.id === $('eventId').value);
  const now = new Date().toISOString();
  await putEvent({ ...(old || {}), id: old?.id || crypto.randomUUID(), date: $('eventDate').value, title: $('eventTitle').value.trim(), memo: $('eventMemo').value.trim(), color: $('eventColor').value, alarmTime: $('eventAlarm').value, updatedAt: now, deletedAt: null, source: 'local' });
  show('eventModal', false);
  await refresh();
  openDay(new Date(state.selected + 'T12:00:00'));
  autoSync();
}
async function autoSync() {
  if (!navigator.onLine || !await account()) return;
  doSync(false);
}
async function doSync(interactive = true) {
  $('syncDot').className = '';
  setStatus('동기화 중…');
  try {
    const result = await syncNow(interactive);
    if (result) {
      $('syncDot').className = 'ok';
      setStatus(`동기화 완료 · 일정 ${result.total}개`, 'ok');
      await refresh();
      toast('동기화했습니다.');
    }
  } catch (error) {
    $('syncDot').className = 'error';
    setStatus(error.message, 'error');
    if (interactive) toast(error.message);
  }
}
function setStatus(text, type = '') {
  $('connectionStatus').textContent = text;
  $('connectionStatus').className = 'status ' + type;
}
async function updateAccountUI() {
  const user = await account();
  show('welcomeView', !user);
  show('calendarView', !!user);
  $('accountInfo').innerHTML = user ? `<strong>${escapeText(user.name || 'Microsoft 사용자')}</strong><br><small>${escapeText(user.username || '')}</small>` : '로그인되지 않았습니다.';
  if (user) {
    const last = localStorage.getItem('lastSync');
    setStatus(last ? `마지막 동기화: ${new Date(last).toLocaleString('ko-KR')}` : '첫 동기화가 필요합니다.');
    $('accountBtn').title = user.name || user.username;
  }
}
function bind() {
  $('helpBtn').onclick = () => show('helpPanel');
  $('closeHelpBtn').onclick = () => show('helpPanel', false);
  $('prevBtn').onclick = () => { state.date = new Date(state.date.getFullYear(), state.date.getMonth() - 1, 1); render(); };
  $('nextBtn').onclick = () => { state.date = new Date(state.date.getFullYear(), state.date.getMonth() + 1, 1); render(); };
  $('todayBtn').onclick = () => { state.date = new Date(); render(); };
  $('closeDayBtn').onclick = () => show('daySheet', false);
  $('addBtn').onclick = () => openForm();
  $('closeFormBtn').onclick = () => show('eventModal', false);
  $('eventForm').onsubmit = saveForm;
  $('deleteBtn').onclick = async () => {
    if (!confirm('이 일정을 삭제할까요?')) return;
    await removeEvent($('eventId').value);
    show('eventModal', false);
    await refresh();
    show('daySheet', false);
    autoSync();
  };
  $('loginBtn').onclick = async () => { try { await login(); await updateAccountUI(); doSync(true); } catch (error) { toast(error.message); } };
  $('syncBtn').onclick = () => doSync(true);
  $('panelSyncBtn').onclick = () => doSync(true);
  $('accountBtn').onclick = () => { updateAccountUI(); show('accountPanel'); };
  $('notificationBtn').onclick = enableNotifications;
  $('closeAccountBtn').onclick = () => show('accountPanel', false);
  $('logoutBtn').onclick = async () => { await logout(); location.reload(); };
  $('templateBtn').onclick = () => { const a = document.createElement('a'); a.href = 'template/캘린더(자동).xlsx'; a.download = '캘린더(자동).xlsx'; a.click(); };
  document.querySelectorAll('.sheet').forEach(sheet => sheet.addEventListener('click', event => { if (event.target === sheet) show(sheet.id, false); }));
}
async function init() {
  await initStorage();
  await refresh();
  bind();
  try {
    await initAuth();
    await updateAccountUI();
    if (await account()) doSync(false);
  } catch (error) {
    show('welcomeView');
    show('calendarView', false);
    toast(error.message);
  }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
  updateNotificationUI();
  notificationTimer = setInterval(checkScheduledNotifications, 30000);
  syncTimer = setInterval(() => { if (document.visibilityState === 'visible') autoSync(); }, CFG.syncIntervalMs);
}
document.addEventListener('DOMContentLoaded', init);
