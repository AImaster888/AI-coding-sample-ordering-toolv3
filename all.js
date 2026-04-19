/* ============================================================
   公司點餐系統 — all.js
   技術架構：
     - Google Identity Services (GIS) 負責 OAuth 2.0 身分驗證
     - Google Sheets API v4（REST）負責試算表讀寫
     - 純原生 JS，無框架依賴
   初始化流程：
     1. all.js 載入時設定 window.onGoogleLibraryLoad
     2. GSI 腳本（accounts.google.com/gsi/client）載入完成後
        自動呼叫 window.onGoogleLibraryLoad → initAuth()
============================================================ */

/* ============================================================
   ⚠️ 設定區域 — 使用前請修改這兩個必填值
============================================================ */
const CONFIG = {
  // GCP Console → API 和服務 → 憑證 → OAuth 2.0 用戶端 ID
  // 應用程式類型選「網頁應用程式」，並在「已授權的 JavaScript 來源」
  // 加入此工具的部署網址（本機開發時加 http://localhost 或 file://）
  CLIENT_ID: '875020439363-gljqn23a9m36c4nbupt776d5e4hhporo.apps.googleusercontent.com',

  // Google Sheets 網址中 /d/ 和 /edit 之間的長串 ID
  // 例：https://docs.google.com/spreadsheets/d/【這段文字】/edit
  SPREADSHEET_ID: '1l9SHfAySYgt2Lve6ym2R9o4vrIyRsk9PqsF1x8KBw_0',

  // OAuth 存取範圍：試算表讀寫 + 取得使用者基本資料
  SCOPES: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
  ].join(' '),
};

/* ============================================================
   工作表名稱常數（與試算表中的分頁名稱對應）
============================================================ */
const SHEETS = {
  TODAY_CONFIG: 'TodayConfig',
  MENU:         'Menu',
  USERS:        'Users',
  ORDERS:       'Orders',
};

/* Sheets API 基礎端點 */
const SHEETS_BASE = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}`;

/* ============================================================
   全域應用程式狀態（單一資料來源）
============================================================ */
const state = {
  accessToken:      null,  // Google OAuth 存取令牌
  currentUser:      null,  // { name, email, role }
  todayRestaurants: [],    // 今日開放的餐廳名稱陣列
  allMenuItems:     [],    // 所有菜單項目 [{ restaurant, name, price, category }]
  allRestaurants:   [],    // 所有不重複的餐廳名稱（給管理員設定用）
  todayOrders:      [],    // 今日訂單陣列 [{ time, email, restaurant, item, price, note }]
  cart:             [],    // 購物車（暫存，尚未送出）[{ restaurant, name, unitPrice, qty, note }]
};

/* OAuth Token Client（Google Identity Services 物件） */
let tokenClient = null;

/* ============================================================
   GSI 函式庫就緒回呼
   window.onGoogleLibraryLoad 是 GSI 提供的特殊鉤子，
   函式庫載入完成後會自動呼叫此函數。
============================================================ */
window.onGoogleLibraryLoad = function () {
  initAuth();
};

/* ============================================================
   初始化 OAuth 用戶端並顯示登入介面
   ⚠️ 登入按鈕與未授權登出按鈕必須在這裡就綁好，
      因為 initEventListeners() 要等進入 App 後才執行。
============================================================ */
function initAuth() {
  // 建立 OAuth 2.0 Token Client
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.CLIENT_ID,
    scope:     CONFIG.SCOPES,
    // 取得令牌後的回呼函數
    callback:  handleTokenResponse,
  });

  // 登入按鈕：點擊後向 Google 請求 Access Token
  document.getElementById('login-btn').addEventListener('click', handleLogin);

  // 未授權頁的「換帳號」按鈕
  document.getElementById('unauthorized-logout-btn').addEventListener('click', logout);

  // 隱藏載入遮罩，顯示登入頁面
  hideLoading();
  showSection('login');
}

/* ============================================================
   點擊「登入」按鈕：向 Google 請求 Access Token
   （若使用者已有快取的 token，可能不顯示選帳號視窗）
============================================================ */
function handleLogin() {
  tokenClient.requestAccessToken({ prompt: 'consent' });
}

/* ============================================================
   OAuth Token 回應處理
   成功取得令牌後：查詢使用者資料 → 核對 Users 表 → 載入 App
============================================================ */
async function handleTokenResponse(response) {
  if (response.error) {
    console.error('[OAuth 錯誤]', response);
    showToast('登入失敗，請重試', 'error');
    return;
  }

  state.accessToken = response.access_token;
  showLoading('驗證身分中…');

  try {
    // 取得 Google 使用者資訊（email、name 等）
    const userInfo = await fetchUserInfo();

    // 比對 Users 工作表，確認是否有存取權限
    const userRecord = await checkUserAuthorization(userInfo.email);

    if (!userRecord) {
      // Email 不在授權名單 → 顯示未授權頁面
      hideLoading();
      showSection('unauthorized');
      document.getElementById('unauthorized-email').textContent = userInfo.email;
      return;
    }

    // 設定目前登入使用者
    state.currentUser = {
      name:  userRecord.name,
      email: userInfo.email,
      role:  userRecord.role, // '管理員' 或 '一般成員'
    };

    // 載入主應用程式
    await loadApp();

  } catch (err) {
    hideLoading();
    console.error('[初始化失敗]', err);
    showToast(`初始化失敗：${err.message}`, 'error');
  }
}

/* ============================================================
   呼叫 Google userinfo 端點取得使用者基本資料
============================================================ */
async function fetchUserInfo() {
  return apiGet('https://www.googleapis.com/oauth2/v3/userinfo');
}

/* ============================================================
   比對 Users 工作表，回傳使用者資料或 null（未授權）
============================================================ */
async function checkUserAuthorization(email) {
  // 讀取 Users 工作表（跳過第一列標題，取 姓名/Email/權限）
  const rows = await readSheet(`${SHEETS.USERS}!A2:C`);

  for (const row of rows) {
    const [name, rowEmail, role] = row;
    if (rowEmail && rowEmail.trim().toLowerCase() === email.toLowerCase()) {
      return { name: name || '', email: rowEmail.trim(), role: role || '一般成員' };
    }
  }

  return null; // 找不到對應的 Email
}

/* ============================================================
   載入主應用程式：同步拉取三張表，解析資料後渲染 UI
============================================================ */
async function loadApp() {
  showLoading('載入餐廳資料中…');

  try {
    // 同時發出三個 API 請求，加快速度
    const [todayRows, menuRows, orderRows] = await Promise.all([
      readSheet(`${SHEETS.TODAY_CONFIG}!A2:A`),
      readSheet(`${SHEETS.MENU}!A2:D`),
      readSheet(`${SHEETS.ORDERS}!A2:F`),
    ]);

    // 解析今日開放餐廳
    state.todayRestaurants = todayRows
      .map(r => (r[0] || '').trim())
      .filter(Boolean);

    // 解析所有菜單項目
    state.allMenuItems = menuRows
      .filter(r => r.length >= 2 && r[0] && r[1])
      .map(r => ({
        restaurant: (r[0] || '').trim(),
        name:       (r[1] || '').trim(),
        price:      parseInt(r[2]) || 0,
        category:   (r[3] || '').trim(),
      }));

    // 從菜單中提取所有不重複餐廳名稱（管理員設定用）
    state.allRestaurants = [...new Set(state.allMenuItems.map(i => i.restaurant))].filter(Boolean);

    // 解析今日訂單（只取今天的記錄）
    const today = getTodayString();
    state.todayOrders = orderRows
      .filter(r => isTodayRow(r[0]))
      .map(r => ({
        time:       r[0] || '',
        email:      r[1] || '',
        restaurant: r[2] || '',
        item:       r[3] || '',
        price:      parseInt(r[4]) || 0,
        note:       r[5] || '',
      }));

    // 初始化 UI 事件監聽
    initEventListeners();

    // 渲染應用程式介面
    renderApp();

    hideLoading();
    showSection('app');

  } catch (err) {
    hideLoading();
    throw err;
  }
}

/* ============================================================
   一次性掛載應用程式內的事件監聽器（登入/未授權相關已在 initAuth 綁定）
============================================================ */
function initEventListeners() {
  // App 內的登出按鈕
  document.getElementById('app-logout-btn').addEventListener('click', logout);

  // 主頁籤切換
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => showTab(btn.dataset.tab));
  });

  // 菜單卡片點餐按鈕（事件委派：綁在容器上，避免 innerHTML 後失效）
  document.getElementById('menu-cards').addEventListener('click', handleMenuCardClick);

  // 訂單頁刷新按鈕
  document.getElementById('refresh-orders-btn').addEventListener('click', refreshOrders);

  // 一鍵複製按鈕
  document.getElementById('copy-orders-btn').addEventListener('click', copyOrdersToClipboard);

  // 管理員：儲存今日餐廳設定
  document.getElementById('save-setup-btn').addEventListener('click', saveTodayConfig);

  // 管理員：清空今日訂單
  document.getElementById('clear-orders-btn').addEventListener('click', clearAllOrders);

  // 一般使用者：清除自己的今日訂單
  document.getElementById('cancel-my-orders-btn').addEventListener('click', cancelMyOrders);

  // 管理員：新增品項列
  document.getElementById('add-menu-item-row-btn').addEventListener('click', addMenuItemRow);

  // 管理員：確認新增餐廳
  document.getElementById('save-new-restaurant-btn').addEventListener('click', addNewRestaurant);

  // 管理員：刪除現有餐廳（事件委派）
  document.getElementById('restaurant-manage-list').addEventListener('click', handleRestaurantManageClick);

  // 購物車：清空 / 確認送出
  document.getElementById('clear-cart-btn').addEventListener('click', clearCart);
  document.getElementById('confirm-order-btn').addEventListener('click', confirmOrder);

  // 購物車項目刪除（事件委派）
  document.getElementById('cart-items').addEventListener('click', handleCartClick);

  // 管理員刪除個別訂單（事件委派）
  document.getElementById('orders-list').addEventListener('click', handleOrderDeleteClick);
}

/* ============================================================
   渲染整個應用程式 UI
============================================================ */
function renderApp() {
  const { currentUser } = state;
  const isAdmin = currentUser.role === '管理員';

  // 顯示使用者姓名與權限
  document.getElementById('user-display-name').textContent = currentUser.name || currentUser.email;
  const roleEl = document.getElementById('user-display-role');
  roleEl.textContent = isAdmin ? '★ 管理員' : '一般成員';
  roleEl.classList.toggle('is-admin', isAdmin);

  // 根據角色顯示/隱藏管理員專屬元素
  document.querySelectorAll('.admin-only').forEach(el => {
    el.style.display = isAdmin ? '' : 'none';
  });
  // 清空訂單區塊同步（在查詢訂單頁籤內）
  const clearSection = document.getElementById('clear-orders-section');
  if (clearSection) clearSection.style.display = isAdmin ? '' : 'none';

  // 渲染各頁籤
  renderMenuTab();
  renderOrdersTab();
  if (isAdmin) renderSetupTab();
}

/* ============================================================
   今日點餐頁籤 — 渲染餐廳選擇器與第一家餐廳的菜單
============================================================ */
function renderMenuTab() {
  const { todayRestaurants } = state;
  const tabsEl  = document.getElementById('restaurant-tabs');
  const cardsEl = document.getElementById('menu-cards');

  if (todayRestaurants.length === 0) {
    tabsEl.innerHTML  = '';
    cardsEl.innerHTML = buildEmptyState('🍽️', '今日尚未設定開放餐廳', '請等候管理員設定今日可點餐的餐廳');
    return;
  }

  // 渲染餐廳選擇標籤（pill 樣式）
  tabsEl.innerHTML = todayRestaurants.map((name, i) =>
    `<button class="restaurant-tab${i === 0 ? ' active' : ''}"
             data-restaurant="${escapeAttr(name)}">
       ${getRestaurantEmoji(name)} ${escapeHtml(name)}
     </button>`
  ).join('');

  // 餐廳標籤點擊切換（事件委派）
  tabsEl.addEventListener('click', function handleRestaurantTab(e) {
    const btn = e.target.closest('.restaurant-tab');
    if (!btn) return;
    document.querySelectorAll('.restaurant-tab').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    renderMenuCards(btn.dataset.restaurant);
  });

  // 預設顯示第一家餐廳的菜單
  renderMenuCards(todayRestaurants[0]);
}

/* ============================================================
   渲染指定餐廳的菜單卡片（依分類分組）
============================================================ */
function renderMenuCards(restaurantName) {
  const cardsEl = document.getElementById('menu-cards');
  const items   = state.allMenuItems.filter(i => i.restaurant === restaurantName);

  if (items.length === 0) {
    cardsEl.innerHTML = buildEmptyState('😅', '此餐廳暫無菜單資料', '請聯繫管理員在 Menu 工作表新增餐點');
    return;
  }

  // 將菜單依「分類」欄位分組
  const groups = {};
  items.forEach(item => {
    const cat = item.category || '其他';
    if (!groups[cat]) groups[cat] = [];
    groups[cat].push(item);
  });

  cardsEl.innerHTML = Object.entries(groups).map(([cat, catItems]) =>
    `<div class="category-section">
       <div class="category-header">${getCategoryEmoji(cat)} ${escapeHtml(cat)}</div>
       <div class="cards-grid">
         ${catItems.map(buildMenuCard).join('')}
       </div>
     </div>`
  ).join('');
}

/* ============================================================
   建立單一菜單卡片 HTML
   使用 data-* 屬性儲存資料（避免 innerHTML 中的 XSS 風險）
============================================================ */
function buildMenuCard(item) {
  return `
    <div class="menu-card"
         data-restaurant="${escapeAttr(item.restaurant)}"
         data-name="${escapeAttr(item.name)}"
         data-price="${item.price}"
         data-category="${escapeAttr(item.category)}">
      <div class="card-header">
        <div class="card-name">${escapeHtml(item.name)}</div>
        <div class="card-price">NT$&nbsp;${item.price}</div>
      </div>
      <!-- 數量選擇器 -->
      <div class="card-qty-row">
        <span class="qty-label">數量</span>
        <div class="qty-ctrl">
          <button class="qty-btn qty-minus" type="button" aria-label="減少">−</button>
          <input
            type="number"
            class="qty-input"
            value="1" min="1" max="20"
            aria-label="數量"
            readonly
          />
          <button class="qty-btn qty-plus" type="button" aria-label="增加">＋</button>
        </div>
        <span class="qty-subtotal">NT$&nbsp;${item.price}</span>
      </div>
      <input
        type="text"
        class="card-note"
        placeholder="備註（如：不要辣、少冰）"
        maxlength="50"
        aria-label="備註"
      />
      <button class="order-btn" type="button">＋ 加入點餐清單</button>
    </div>`;
}

/* ============================================================
   菜單卡片點擊處理（事件委派統一入口）
   同時處理：數量 +/− 按鈕 與 點餐按鈕
============================================================ */
async function handleMenuCardClick(e) {
  const card = e.target.closest('.menu-card');
  if (!card) return;

  // ── 數量減少 ──
  if (e.target.closest('.qty-minus')) {
    const input = card.querySelector('.qty-input');
    const val   = Math.max(1, parseInt(input.value) - 1);
    input.value = val;
    updateSubtotal(card, val);
    return;
  }

  // ── 數量增加 ──
  if (e.target.closest('.qty-plus')) {
    const input = card.querySelector('.qty-input');
    const val   = Math.min(20, parseInt(input.value) + 1);
    input.value = val;
    updateSubtotal(card, val);
    return;
  }

  // ── 加入購物車按鈕 ──
  const btn = e.target.closest('.order-btn');
  if (!btn || btn.disabled) return;

  const qty       = parseInt(card.querySelector('.qty-input')?.value) || 1;
  const unitPrice = parseInt(card.dataset.price) || 0;
  const note      = (card.querySelector('.card-note')?.value || '').trim();

  // 加入購物車（本地狀態，尚未寫入 Sheet）
  state.cart.push({
    restaurant: card.dataset.restaurant,
    name:       card.dataset.name,
    unitPrice,
    qty,
    note,
  });

  // 重設數量為 1，清空備註
  const qtyInput  = card.querySelector('.qty-input');
  const noteInput = card.querySelector('.card-note');
  if (qtyInput)  { qtyInput.value = 1; updateSubtotal(card, 1); }
  if (noteInput) noteInput.value = '';

  // 按鈕短暫反饋
  btn.disabled    = true;
  btn.textContent = '✅ 已加入！';
  setTimeout(() => {
    btn.disabled    = false;
    btn.textContent = '＋ 加入點餐清單';
  }, 1200);

  showToast(`已加入：${card.dataset.name}${qty > 1 ? ` ×${qty}` : ''}`, 'success');
  renderCart();
}

/* 更新卡片小計顯示 */
function updateSubtotal(card, qty) {
  const unitPrice = parseInt(card.dataset.price) || 0;
  const subtotalEl = card.querySelector('.qty-subtotal');
  if (subtotalEl) subtotalEl.textContent = `NT\u00a0${unitPrice * qty}`;
}

/* ============================================================
   購物車：即時渲染使用者目前的點餐清單（未送出）
============================================================ */
function renderCart() {
  const section    = document.getElementById('cart-section');
  const itemsEl    = document.getElementById('cart-items');
  const totalEl    = document.getElementById('cart-total-amount');
  const confirmBtn = document.getElementById('confirm-order-btn');

  if (state.cart.length === 0) {
    section.style.display = 'none';
    return;
  }

  section.style.display = '';

  // 依餐廳分組顯示
  const byRestaurant = {};
  state.cart.forEach((item, idx) => {
    if (!byRestaurant[item.restaurant]) byRestaurant[item.restaurant] = [];
    byRestaurant[item.restaurant].push({ ...item, idx });
  });

  itemsEl.innerHTML = Object.entries(byRestaurant).map(([restaurant, items]) => `
    <div class="cart-group">
      <div class="cart-group-title">${getRestaurantEmoji(restaurant)} ${escapeHtml(restaurant)}</div>
      ${items.map(item => `
        <div class="cart-item">
          <div class="cart-item-info">
            <span class="cart-item-name">${escapeHtml(item.name)}${item.qty > 1 ? ` <em>×${item.qty}</em>` : ''}</span>
            ${item.note ? `<span class="cart-item-note">📝 ${escapeHtml(item.note)}</span>` : ''}
          </div>
          <div class="cart-item-right">
            <span class="cart-item-price">NT$ ${item.unitPrice * item.qty}</span>
            <button class="cart-remove-btn" data-idx="${item.idx}" title="移除">✕</button>
          </div>
        </div>`).join('')}
    </div>`).join('');

  const total = state.cart.reduce((s, i) => s + i.unitPrice * i.qty, 0);
  totalEl.textContent = `NT$ ${total.toLocaleString()}`;
}

/* 購物車：移除單一項目（事件委派入口） */
function handleCartClick(e) {
  const btn = e.target.closest('.cart-remove-btn');
  if (!btn) return;
  const idx = parseInt(btn.dataset.idx);
  if (!isNaN(idx)) {
    state.cart.splice(idx, 1);
    renderCart();
  }
}

/* 購物車：清空全部 */
function clearCart() {
  if (state.cart.length === 0) return;
  if (!confirm('確定要清空目前的點餐清單嗎？')) return;
  state.cart = [];
  renderCart();
}

/* ============================================================
   確認送出訂單：使用者檢視購物車後確認，才寫入 Orders 工作表
============================================================ */
async function confirmOrder() {
  if (state.cart.length === 0) return;

  // 組出確認文字讓使用者再確認一次
  const total   = state.cart.reduce((s, i) => s + i.unitPrice * i.qty, 0);
  const listText = state.cart.map(i =>
    `• ${i.name}${i.qty > 1 ? ` ×${i.qty}` : ''}  NT$${i.unitPrice * i.qty}${i.note ? `（${i.note}）` : ''}`
  ).join('\n');

  const ok = confirm(
    `確認送出以下訂單？\n\n${listText}\n\n合計：NT$ ${total.toLocaleString()}\n\n送出後將無法修改。`
  );
  if (!ok) return;

  const btn       = document.getElementById('confirm-order-btn');
  btn.disabled    = true;
  btn.textContent = '送出中…';

  try {
    const now  = formatDateTime(new Date());
    const rows = state.cart.map(item => [
      now,
      state.currentUser.email,
      item.restaurant,
      item.qty > 1 ? `${item.name} ×${item.qty}` : item.name,
      item.unitPrice * item.qty,
      item.note,
    ]);

    // 一次寫入所有購物車項目
    await appendRows(`${SHEETS.ORDERS}!A:F`, rows);

    // 同步更新本地 todayOrders
    rows.forEach((r, i) => {
      state.todayOrders.push({
        time:       r[0],
        email:      r[1],
        restaurant: r[2],
        item:       r[3],
        price:      r[4],
        note:       r[5],
      });
    });

    // 清空購物車
    state.cart = [];
    renderCart();
    renderOrdersTab();

    showToast(`🎉 訂單已送出！共 ${rows.length} 項，合計 NT$ ${total.toLocaleString()}`, 'success');

  } catch (err) {
    showToast(`送出失敗：${err.message}`, 'error');
  } finally {
    btn.disabled    = false;
    btn.textContent = '✅ 確認送出訂單';
  }
}

/* ============================================================
   送出訂單：寫入一筆資料到 Orders 工作表
============================================================ */
async function submitOrder(item, note, btn) {
  // 立即停用按鈕，防止重複點擊
  btn.disabled     = true;
  btn.textContent  = '送出中…';

  const now      = formatDateTime(new Date());
  const qty      = item.qty || 1;
  // 數量 > 1 時在餐點名稱後加 ×N 標示
  const itemName = qty > 1 ? `${item.name} ×${qty}` : item.name;
  const row = [
    now,                       // 點餐時間
    state.currentUser.email,   // 訂購人 Email
    item.restaurant,           // 餐廳名稱
    itemName,                  // 餐點內容（含數量）
    item.price,                // 金額（已乘數量）
    note,                      // 備註
  ];

  try {
    // 以 append 模式寫入（不覆蓋既有資料）
    await appendRows(`${SHEETS.ORDERS}!A:F`, [row]);

    // 同步更新本地狀態
    state.todayOrders.push({
      time:       now,
      email:      state.currentUser.email,
      restaurant: item.restaurant,
      item:       itemName,
      price:      item.price,
      note:       note,
    });

    // 清空備註輸入框，數量重設為 1
    const cardEl = btn.closest('.menu-card');
    if (cardEl) {
      const noteInput = cardEl.querySelector('.card-note');
      const qtyInput  = cardEl.querySelector('.qty-input');
      if (noteInput) noteInput.value = '';
      if (qtyInput)  { qtyInput.value = 1; updateSubtotal(cardEl, 1); }
    }

    // 按鈕變綠表示成功
    btn.classList.add('success');
    btn.textContent = `✅ 已點 ${qty > 1 ? qty + ' 份' : ''}！`;
    showToast(`🎉 成功點餐：${itemName}（NT$ ${item.price}）`, 'success');

    // 3 秒後恢復按鈕狀態
    setTimeout(() => {
      btn.disabled     = false;
      btn.textContent  = '🛒 點我下單';
      btn.classList.remove('success');
    }, 3000);

    // 同步更新確認訂單頁籤的資料
    renderOrdersTab();

  } catch (err) {
    btn.disabled     = false;
    btn.textContent  = '🛒 點我下單';
    showToast(`點餐失敗：${err.message}`, 'error');
  }
}

/* ============================================================
   確認訂單頁籤 — 顯示今日所有訂單（依餐廳分組）
============================================================ */
function renderOrdersTab() {
  const { todayOrders, currentUser } = state;
  const listEl      = document.getElementById('orders-list');
  const copyBtn     = document.getElementById('copy-orders-btn');
  const cancelMyBtn = document.getElementById('cancel-my-orders-btn');
  const isAdmin     = currentUser.role === '管理員';

  // 一般成員只看自己的訂單；管理員看全部
  const visibleOrders = isAdmin
    ? todayOrders
    : todayOrders.filter(o => o.email === currentUser.email);

  const myOrders = todayOrders.filter(o => o.email === currentUser.email);

  if (visibleOrders.length === 0) {
    listEl.innerHTML      = buildEmptyState('📭', '今日尚無訂單', '快去點餐吧！大家在等你喔 😋');
    copyBtn.style.display = 'none';
    cancelMyBtn.style.display = 'none';
    return;
  }

  copyBtn.style.display     = isAdmin ? '' : 'none'; // 複製功能僅管理員需要
  cancelMyBtn.style.display = myOrders.length > 0 ? '' : 'none';

  // 依餐廳分組（以 visibleOrders 為準）
  const byRestaurant = {};
  visibleOrders.forEach(o => {
    if (!byRestaurant[o.restaurant]) byRestaurant[o.restaurant] = [];
    byRestaurant[o.restaurant].push(o);
  });

  const totalCount  = visibleOrders.length;
  const totalAmount = visibleOrders.reduce((s, o) => s + o.price, 0);

  const summaryLabel = isAdmin ? '今日全部' : '我的訂單';
  listEl.innerHTML = `
    <div class="orders-summary">
      📊 ${summaryLabel} <strong>${totalCount}</strong> 筆 ／ 合計
      <strong>NT$ ${totalAmount.toLocaleString()}</strong>
    </div>
    ${Object.entries(byRestaurant).map(([restaurant, orders]) => `
      <div class="orders-group">
        <div class="orders-group-title">
          ${getRestaurantEmoji(restaurant)} ${escapeHtml(restaurant)}
          <span class="orders-group-count">${orders.length} 筆</span>
        </div>
        ${orders.map((o, i) => `
          <div class="order-item">
            <div class="order-item-main">
              <div class="order-email">${escapeHtml(o.email)}</div>
              <div class="order-detail">
                <span class="order-food">${escapeHtml(o.item)}</span>
                <span class="order-price">NT$ ${o.price}</span>
              </div>
              ${o.note ? `<div class="order-note">📝 ${escapeHtml(o.note)}</div>` : ''}
            </div>
            ${isAdmin ? `
              <button class="order-delete-btn"
                      data-time="${escapeAttr(o.time)}"
                      data-email="${escapeAttr(o.email)}"
                      data-item="${escapeAttr(o.item)}"
                      title="刪除此筆訂單">🗑️</button>` : ''}
          </div>`).join('')}
      </div>`).join('')}`;
}

/* ============================================================
   一鍵複製：產生可貼到 LINE 的純文字摘要
============================================================ */
function copyOrdersToClipboard() {
  const { todayOrders } = state;
  const today = getTodayString();

  // 依餐廳分組
  const byRestaurant = {};
  todayOrders.forEach(o => {
    if (!byRestaurant[o.restaurant]) byRestaurant[o.restaurant] = [];
    byRestaurant[o.restaurant].push(o);
  });

  const totalAmount = todayOrders.reduce((s, o) => s + o.price, 0);

  let text = `📋 今日點餐確認（${today}）\n`;
  text    += '================================\n';

  Object.entries(byRestaurant).forEach(([restaurant, orders]) => {
    text += `\n🍽️ ${restaurant}\n`;
    orders.forEach(o => {
      text += `  • ${o.email} — ${o.item}  NT$${o.price}`;
      if (o.note) text += `  [備註：${o.note}]`;
      text += '\n';
    });
  });

  text += '\n================================\n';
  text += `📊 總計：${todayOrders.length} 筆訂單 ／ NT$ ${totalAmount.toLocaleString()}`;

  // 優先使用現代 Clipboard API，失敗時降級使用 execCommand
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text)
      .then(() => showToast('已複製！可直接貼到 LINE 📋', 'success'))
      .catch(() => fallbackCopy(text));
  } else {
    fallbackCopy(text);
  }
}

/* 降級複製方法（for 不支援 Clipboard API 的環境） */
function fallbackCopy(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.cssText = 'position:fixed;top:0;left:0;opacity:0';
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  try {
    document.execCommand('copy');
    showToast('已複製！可直接貼到 LINE 📋', 'success');
  } catch {
    showToast('複製失敗，請手動複製', 'error');
  }
  document.body.removeChild(ta);
}

/* ============================================================
   管理員：設定餐廳 — 渲染三個子區塊
============================================================ */
function renderSetupTab() {
  renderTodayCheckboxes();
  renderRestaurantManagement();
  initNewRestaurantForm();
}

/* A. 今日開放餐廳 checkbox 清單 */
function renderTodayCheckboxes() {
  const { allRestaurants, todayRestaurants } = state;
  const el = document.getElementById('restaurant-checkboxes');

  if (allRestaurants.length === 0) {
    el.innerHTML = `<p style="color:var(--text-muted);font-size:.9rem;padding:12px 0">
      菜單中目前沒有餐廳，請先在下方「新增餐廳」。</p>`;
    return;
  }

  el.innerHTML = allRestaurants.map(name => {
    const isChecked = todayRestaurants.includes(name);
    const labelId   = `rest-${btoa(unescape(encodeURIComponent(name)))}`;
    return `
      <label class="restaurant-checkbox-label${isChecked ? ' checked' : ''}" for="${labelId}">
        <input type="checkbox" id="${labelId}" value="${escapeAttr(name)}" ${isChecked ? 'checked' : ''} />
        <span class="checkbox-mark"></span>
        <span class="checkbox-text">${getRestaurantEmoji(name)} ${escapeHtml(name)}</span>
      </label>`;
  }).join('');

  el.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', function () {
      this.closest('.restaurant-checkbox-label').classList.toggle('checked', this.checked);
    });
  });
}

/* B. 現有餐廳清單（含刪除按鈕） */
function renderRestaurantManagement() {
  const { allRestaurants } = state;
  const el = document.getElementById('restaurant-manage-list');

  if (allRestaurants.length === 0) {
    el.innerHTML = `<p style="color:var(--text-muted);font-size:.9rem;padding:8px 0">尚無餐廳資料。</p>`;
    return;
  }

  el.innerHTML = allRestaurants.map(name => {
    const itemCount = state.allMenuItems.filter(i => i.restaurant === name).length;
    return `
      <div class="manage-restaurant-row">
        <span class="manage-restaurant-name">
          ${getRestaurantEmoji(name)} ${escapeHtml(name)}
          <span class="manage-restaurant-count">${itemCount} 項</span>
        </span>
        <button class="btn btn-sm btn-cancel delete-restaurant-btn"
                data-restaurant="${escapeAttr(name)}">🗑️ 刪除</button>
      </div>`;
  }).join('');
}

/* C. 新增餐廳表單初始化（加入第一列品項輸入） */
function initNewRestaurantForm() {
  document.getElementById('new-restaurant-name').value = '';
  const itemsEl = document.getElementById('new-menu-items');
  itemsEl.innerHTML = '';
  addMenuItemRow(); // 預設一列
}

/* 動態新增一列品項輸入 */
function addMenuItemRow() {
  const el  = document.getElementById('new-menu-items');
  const row = document.createElement('div');
  row.className = 'menu-item-row';
  row.innerHTML = `
    <input type="text"  class="form-input item-name"     placeholder="品名（例：大麥克）" maxlength="30" />
    <input type="number" class="form-input item-price"   placeholder="單價" min="1" max="9999" />
    <input type="text"  class="form-input item-category" placeholder="分類（例：主食）" maxlength="20" />
    <button class="item-remove-btn" type="button" title="移除此列">✕</button>`;

  // 移除此列按鈕（至少保留一列）
  row.querySelector('.item-remove-btn').addEventListener('click', () => {
    if (document.querySelectorAll('.menu-item-row').length > 1) {
      row.remove();
    } else {
      showToast('至少需要一項餐點', 'error');
    }
  });

  el.appendChild(row);
}

/* 餐廳管理刪除按鈕點擊（事件委派） */
function handleRestaurantManageClick(e) {
  const btn = e.target.closest('.delete-restaurant-btn');
  if (!btn) return;
  deleteRestaurant(btn.dataset.restaurant);
}

/* ============================================================
   管理員：儲存今日餐廳設定到 TodayConfig 工作表
============================================================ */
async function saveTodayConfig() {
  const checked  = document.querySelectorAll('#restaurant-checkboxes input:checked');
  const selected = Array.from(checked).map(cb => [cb.value]);

  if (selected.length === 0) {
    const ok = confirm('⚠️ 您未選取任何餐廳，確定要清空今日設定嗎？\n清空後員工將看不到可點餐的選項。');
    if (!ok) return;
  }

  const btn       = document.getElementById('save-setup-btn');
  btn.disabled    = true;
  btn.textContent = '儲存中…';

  try {
    // 先清空現有設定（保留標題列）
    await clearRange(`${SHEETS.TODAY_CONFIG}!A2:A`);

    // 寫入新的餐廳清單
    if (selected.length > 0) {
      await updateRange(`${SHEETS.TODAY_CONFIG}!A2`, selected);
    }

    // 更新本地狀態
    state.todayRestaurants = selected.map(r => r[0]);

    // 重新渲染點餐頁籤，讓改動立即生效
    renderMenuTab();

    showToast(`✅ 已儲存今日餐廳設定（共 ${selected.length} 家）`, 'success');

  } catch (err) {
    showToast(`儲存失敗：${err.message}`, 'error');
  } finally {
    btn.disabled    = false;
    btn.textContent = '💾 儲存今日設定';
  }
}

/* ============================================================
   管理員：刪除整家餐廳（移除 Menu 工作表中該餐廳的所有品項）
============================================================ */
async function deleteRestaurant(name) {
  const itemCount = state.allMenuItems.filter(i => i.restaurant === name).length;
  const confirmed = confirm(
    `確定要刪除「${name}」嗎？\n\n將一併移除其 ${itemCount} 項菜單品項，且無法復原。`
  );
  if (!confirmed) return;

  try {
    showLoading('刪除中…');
    const allRows = await readSheet(`${SHEETS.MENU}!A2:D`);

    // 過濾掉該餐廳的所有列
    const keep = allRows.filter(r => (r[0] || '').trim() !== name);

    await clearRange(`${SHEETS.MENU}!A2:D`);
    if (keep.length > 0) await updateRange(`${SHEETS.MENU}!A2`, keep);

    // 更新本地狀態
    state.allMenuItems     = keep.map(r => ({
      restaurant: r[0] || '', name: r[1] || '',
      price: parseInt(r[2]) || 0, category: r[3] || '',
    }));
    state.allRestaurants   = [...new Set(state.allMenuItems.map(i => i.restaurant))].filter(Boolean);
    state.todayRestaurants = state.todayRestaurants.filter(r => r !== name);

    hideLoading();
    renderSetupTab();
    renderMenuTab();
    showToast(`已刪除餐廳「${name}」及其所有品項`, 'success');

  } catch (err) {
    hideLoading();
    showToast(`刪除失敗：${err.message}`, 'error');
  }
}

/* ============================================================
   管理員：新增餐廳（連同品項一起寫入 Menu 工作表）
============================================================ */
async function addNewRestaurant() {
  const restaurantName = document.getElementById('new-restaurant-name').value.trim();
  if (!restaurantName) { showToast('請填寫餐廳名稱', 'error'); return; }

  // 收集所有品項列
  const rows = [];
  document.querySelectorAll('.menu-item-row').forEach(row => {
    const name     = row.querySelector('.item-name')?.value.trim();
    const price    = row.querySelector('.item-price')?.value.trim();
    const category = row.querySelector('.item-category')?.value.trim() || '其他';
    if (name && price) {
      rows.push([restaurantName, name, parseInt(price), category]);
    }
  });

  if (rows.length === 0) {
    showToast('請至少填寫一項餐點（品名與單價必填）', 'error');
    return;
  }

  const btn       = document.getElementById('save-new-restaurant-btn');
  btn.disabled    = true;
  btn.textContent = '新增中…';

  try {
    await appendRows(`${SHEETS.MENU}!A:D`, rows, 'USER_ENTERED');

    // 更新本地狀態
    rows.forEach(r => {
      state.allMenuItems.push({
        restaurant: r[0], name: r[1], price: r[2], category: r[3],
      });
    });
    if (!state.allRestaurants.includes(restaurantName)) {
      state.allRestaurants.push(restaurantName);
    }

    initNewRestaurantForm();   // 清空表單
    renderSetupTab();          // 重新渲染設定頁
    renderMenuTab();           // 更新點餐頁

    showToast(`已新增餐廳「${restaurantName}」（共 ${rows.length} 項品項）`, 'success');

  } catch (err) {
    showToast(`新增失敗：${err.message}`, 'error');
  } finally {
    btn.disabled    = false;
    btn.textContent = '✅ 確認新增餐廳';
  }
}

/* ============================================================
   管理員：清空今日所有訂單（保留 Orders 工作表的標題列）
============================================================ */
async function clearAllOrders() {
  // 第一層：說明危險性，請使用者主動輸入確認字串
  const input = prompt(
    '⚠️ 危險操作警告\n\n' +
    '此動作將永久刪除今日所有人的點餐記錄，且無法復原。\n\n' +
    '請在下方輸入「確認清空」以繼續：'
  );

  // 使用者取消或輸入錯誤
  if (input === null) return; // 按取消
  if (input.trim() !== '確認清空') {
    showToast('輸入錯誤，已取消清空操作', 'error');
    return;
  }

  const btn       = document.getElementById('clear-orders-btn');
  btn.disabled    = true;
  btn.textContent = '清空中…';

  try {
    // 清除 Orders 工作表第 2 列以後的所有資料（保留第 1 列標題）
    await clearRange(`${SHEETS.ORDERS}!A2:F`);

    // 清除本地狀態
    state.todayOrders = [];

    // 重新渲染確認訂單頁籤
    renderOrdersTab();

    showToast('已成功清空今日所有訂單記錄 🗑️', 'success');

  } catch (err) {
    showToast(`清空失敗：${err.message}`, 'error');
  } finally {
    btn.disabled    = false;
    btn.textContent = '🗑️ 清空今日所有點餐記錄';
  }
}

/* ============================================================
   使用者：清除自己今日的所有訂單，以便重新選餐
   作法：讀取全表 → 移除自己今日的列 → 清空後整表重寫
   （Sheets API 不支援條件刪除列，需整表覆寫）
============================================================ */
async function cancelMyOrders() {
  const myEmail = state.currentUser.email;
  const myCount = state.todayOrders.filter(o => o.email === myEmail).length;

  const confirmed = confirm(
    `確定要清除你（${myEmail}）今日的 ${myCount} 筆訂單嗎？\n\n清除後可以重新點餐。`
  );
  if (!confirmed) return;

  const btn       = document.getElementById('cancel-my-orders-btn');
  btn.disabled    = true;
  btn.textContent = '處理中…';

  try {
    // 讀取 Orders 工作表所有資料（含其他人、其他日期的歷史訂單）
    const allRows = await readSheet(`${SHEETS.ORDERS}!A2:F`);
    const today   = getTodayString();

    // 保留：不是「今天 + 自己」的所有列
    const keep = allRows.filter(r => {
      const isToday = isTodayRow(r[0]);
      const isMine  = r[1] && r[1].trim().toLowerCase() === myEmail.toLowerCase();
      return !(isToday && isMine); // 排除掉自己今天的訂單
    });

    // 清空整個資料區（保留標題列第 1 行）
    await clearRange(`${SHEETS.ORDERS}!A2:F`);

    // 若還有剩餘資料，整批寫回
    if (keep.length > 0) {
      await updateRange(`${SHEETS.ORDERS}!A2`, keep);
    }

    // 同步更新本地狀態（移除自己今日訂單）
    state.todayOrders = state.todayOrders.filter(o => o.email !== myEmail);

    renderOrdersTab();
    showToast(`已清除你的 ${myCount} 筆訂單，可以重新點餐了 🍽️`, 'success');

  } catch (err) {
    showToast(`清除失敗：${err.message}`, 'error');
  } finally {
    btn.disabled    = false;
    btn.textContent = '✕ 清除我的訂單';
  }
}

/* ============================================================
   管理員：刪除單筆訂單（點擊 🗑️ 按鈕的事件委派入口）
============================================================ */
function handleOrderDeleteClick(e) {
  const btn = e.target.closest('.order-delete-btn');
  if (!btn) return;
  deleteOrder({
    time:  btn.dataset.time,
    email: btn.dataset.email,
    item:  btn.dataset.item,
  }, btn);
}

/* ============================================================
   管理員：刪除單筆訂單
   作法：讀全表 → 比對 time + email + item 移除那列 → 整表重寫
============================================================ */
async function deleteOrder(target, btn) {
  const confirmed = confirm(
    `確定要刪除以下訂單嗎？\n\n訂購人：${target.email}\n餐點：${target.item}\n時間：${target.time}`
  );
  if (!confirmed) return;

  btn.disabled    = true;
  btn.textContent = '…';

  try {
    // 讀取全部訂單（含其他日期）
    const allRows = await readSheet(`${SHEETS.ORDERS}!A2:F`);

    // 找出要刪除的那列（time + email + item 三欄比對）
    let deleted = false;
    const keep = allRows.filter(r => {
      if (!deleted
          && r[0] === target.time
          && r[1] === target.email
          && r[3] === target.item) {
        deleted = true;   // 只刪第一筆符合的
        return false;
      }
      return true;
    });

    if (!deleted) {
      showToast('找不到該筆訂單，可能已被刪除', 'error');
      btn.disabled    = false;
      btn.textContent = '🗑️';
      return;
    }

    // 清空後整表重寫
    await clearRange(`${SHEETS.ORDERS}!A2:F`);
    if (keep.length > 0) {
      await updateRange(`${SHEETS.ORDERS}!A2`, keep);
    }

    // 同步本地狀態
    const today = getTodayString();
    state.todayOrders = keep
      .filter(r => isTodayRow(r[0]))
      .map(r => ({
        time: r[0], email: r[1], restaurant: r[2],
        item: r[3], price: parseInt(r[4]) || 0, note: r[5] || '',
      }));

    renderOrdersTab();
    showToast(`已刪除 ${target.email} 的訂單：${target.item}`, 'success');

  } catch (err) {
    showToast(`刪除失敗：${err.message}`, 'error');
    btn.disabled    = false;
    btn.textContent = '🗑️';
  }
}

/* ============================================================
   切換主頁籤顯示
============================================================ */
function showTab(tabName) {
  // 更新頁籤按鈕的 active 樣式
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tabName);
  });

  // 顯示對應的內容區塊，隱藏其他
  document.querySelectorAll('.tab-content').forEach(el => {
    el.style.display = 'none';
  });
  const target = document.getElementById(`tab-${tabName}`);
  if (target) target.style.display = '';

  // 切換到「確認訂單」時，從試算表重新拉取最新資料
  if (tabName === 'orders') refreshOrders();
}

/* ============================================================
   從試算表重新拉取今日訂單（刷新按鈕用）
============================================================ */
async function refreshOrders() {
  try {
    const rows  = await readSheet(`${SHEETS.ORDERS}!A2:F`);
    const today = getTodayString();

    state.todayOrders = rows
      .filter(r => isTodayRow(r[0]))
      .map(r => ({
        time:       r[0] || '',
        email:      r[1] || '',
        restaurant: r[2] || '',
        item:       r[3] || '',
        price:      parseInt(r[4]) || 0,
        note:       r[5] || '',
      }));

    renderOrdersTab();
    showToast('訂單已刷新 ✅', 'success');

  } catch (err) {
    showToast(`刷新失敗：${err.message}`, 'error');
  }
}

/* ============================================================
   登出：撤銷 Access Token 並返回登入頁
============================================================ */
function logout() {
  if (state.accessToken) {
    google.accounts.oauth2.revoke(state.accessToken, () => {});
  }
  state.accessToken      = null;
  state.currentUser      = null;
  state.todayRestaurants = [];
  state.allMenuItems     = [];
  state.allRestaurants   = [];
  state.todayOrders      = [];
  showSection('login');
}

/* ============================================================
   Google Sheets API — 資料讀取
   range 格式：'TodayConfig!A2:A'  或  'Menu!A2:D'
============================================================ */
async function readSheet(range) {
  const url  = `${SHEETS_BASE}/values/${encodeURIComponent(range)}`;
  const data = await apiGet(url);
  return data.values || [];
}

/* ============================================================
   Google Sheets API — 新增列（Append，不覆蓋既有資料）
   valueInputOption 預設用 RAW：防止 Sheets 把日期字串自動轉換成
   日期型別，讀回來格式不同導致日期過濾失效。
============================================================ */
async function appendRows(range, values, valueInputOption = 'RAW') {
  const url = `${SHEETS_BASE}/values/${encodeURIComponent(range)}:append`
            + `?valueInputOption=${valueInputOption}&insertDataOption=INSERT_ROWS`;
  return apiPost(url, { values });
}

/* ============================================================
   Google Sheets API — 更新指定範圍（PUT 覆寫）
============================================================ */
async function updateRange(range, values) {
  const url = `${SHEETS_BASE}/values/${encodeURIComponent(range)}`
            + '?valueInputOption=USER_ENTERED';
  return apiPut(url, { values });
}

/* ============================================================
   Google Sheets API — 清除指定範圍的所有內容
============================================================ */
async function clearRange(range) {
  const url = `${SHEETS_BASE}/values/${encodeURIComponent(range)}:clear`;
  return apiPost(url, {});
}

/* ============================================================
   HTTP 輔助函數：統一加入 Authorization Header
============================================================ */
async function apiGet(url) {
  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${state.accessToken}` },
  });
  if (!resp.ok) await throwApiError(resp);
  return resp.json();
}

async function apiPost(url, body) {
  const resp = await fetch(url, {
    method:  'POST',
    headers: {
      Authorization:  `Bearer ${state.accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!resp.ok) await throwApiError(resp);
  return resp.json();
}

async function apiPut(url, body) {
  const resp = await fetch(url, {
    method:  'PUT',
    headers: {
      Authorization:  `Bearer ${state.accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!resp.ok) await throwApiError(resp);
  return resp.json();
}

/* 解析 API 錯誤訊息並拋出 */
async function throwApiError(resp) {
  let message = `HTTP ${resp.status}`;
  try {
    const body = await resp.json();
    message = body?.error?.message || message;
  } catch { /* 無法解析 JSON，使用預設訊息 */ }

  // 401 通常代表 Token 過期，引導使用者重新登入
  if (resp.status === 401) {
    state.accessToken = null;
    showToast('登入已過期，請重新登入', 'error');
    showSection('login');
  }

  throw new Error(message);
}

/* ============================================================
   UI 工具函數
============================================================ */

/* 顯示指定區塊，隱藏其餘區塊 */
function showSection(name) {
  ['login', 'unauthorized', 'app'].forEach(s => {
    const el = document.getElementById(`${s}-section`);
    if (el) el.style.display = (s === name) ? '' : 'none';
  });
}

/* 顯示全螢幕載入遮罩 */
function showLoading(msg = '載入中…') {
  const overlay = document.getElementById('loading-overlay');
  const msgEl   = document.getElementById('loading-message');
  if (overlay) overlay.style.display = 'flex';
  if (msgEl)   msgEl.textContent     = msg;
}

/* 隱藏全螢幕載入遮罩 */
function hideLoading() {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) overlay.style.display = 'none';
}

/* 顯示 Toast 提示訊息（type: 'success' | 'error' | 'info'） */
function showToast(message, type = 'info') {
  const toast    = document.getElementById('toast');
  toast.textContent = message;
  toast.className   = `toast toast-${type} show`;

  clearTimeout(toast._hideTimer);
  toast._hideTimer = setTimeout(() => {
    toast.classList.remove('show');
  }, 3500);
}

/* 產生空狀態的 HTML 字串 */
function buildEmptyState(icon, title, desc) {
  return `
    <div class="empty-state">
      <div class="empty-icon">${icon}</div>
      <h3>${escapeHtml(title)}</h3>
      <p>${escapeHtml(desc)}</p>
    </div>`;
}

/* ============================================================
   日期時間工具
============================================================ */

/* 回傳今天日期字串，格式：'2026/01/19' */
function getTodayString() {
  const d  = new Date();
  const y  = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}/${mo}/${dd}`;
}

/*
 * 判斷試算表中某一列的日期欄位是否為今天。
 * 以寬鬆方式比對：先 startsWith，若失敗再嘗試 Date 解析，
 * 兼容 Sheets 因地區設定轉換成不同格式的日期字串。
 */
function isTodayRow(dateStr) {
  if (!dateStr) return false;
  const today = getTodayString(); // 'YYYY/MM/DD'
  if (String(dateStr).startsWith(today)) return true;

  // 嘗試 Date 解析（兼容 Sheets 轉成 'M/D/YYYY' 等格式）
  try {
    const parsed = new Date(String(dateStr).replace(/\//g, '-'));
    if (isNaN(parsed)) return false;
    const y  = parsed.getFullYear();
    const mo = String(parsed.getMonth() + 1).padStart(2, '0');
    const dd = String(parsed.getDate()).padStart(2, '0');
    return `${y}/${mo}/${dd}` === today;
  } catch {
    return false;
  }
}

/* 回傳日期時間字串，格式：'2026/01/19 12:05' */
function formatDateTime(date) {
  const y  = date.getFullYear();
  const mo = String(date.getMonth() + 1).padStart(2, '0');
  const d  = String(date.getDate()).padStart(2, '0');
  const h  = String(date.getHours()).padStart(2, '0');
  const mi = String(date.getMinutes()).padStart(2, '0');
  return `${y}/${mo}/${d} ${h}:${mi}`;
}

/* ============================================================
   Emoji 輔助：根據餐廳/分類名稱挑選合適的表情符號
============================================================ */
function getRestaurantEmoji(name) {
  const rules = [
    ['排骨|梁社漢',    '🍱'],
    ['嵐|飲料|珍珠|奶茶', '🧋'],
    ['壽司|生魚',      '🍣'],
    ['拉麵|牛肉麵|麵', '🍜'],
    ['披薩',           '🍕'],
    ['漢堡',           '🍔'],
    ['火鍋|麻辣',      '🫕'],
    ['燒肉|烤肉',      '🥩'],
    ['咖啡',           '☕'],
    ['三明治|潛艇堡',  '🥪'],
    ['便當|飯|定食',   '🍱'],
  ];
  for (const [pattern, emoji] of rules) {
    if (new RegExp(pattern).test(name)) return emoji;
  }
  return '🍽️';
}

function getCategoryEmoji(cat) {
  const rules = [
    ['主食|飯|定食|便當', '🍱'],
    ['麵',               '🍜'],
    ['湯',               '🍲'],
    ['小菜|沙拉',        '🥗'],
    ['飲品|飲料|奶茶',   '🧋'],
    ['點心|甜點|蛋糕',   '🍰'],
    ['素食',             '🥬'],
    ['海鮮',             '🦐'],
  ];
  for (const [pattern, emoji] of rules) {
    if (new RegExp(pattern).test(cat)) return emoji;
  }
  return '🍽️';
}

/* ============================================================
   安全性工具：防止 XSS
============================================================ */

/* 轉義 HTML 內容中的特殊字元 */
function escapeHtml(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/* 轉義 HTML 屬性值中的特殊字元 */
function escapeAttr(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
