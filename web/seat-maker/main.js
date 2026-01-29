// main.js

// ダッシュボードGAS（本番環境）
const API_URL = 'https://script.google.com/macros/s/AKfycbzgVIzNjfW_UHihZS5bwrWM7xix0U4dnodZSlq7nPC8eGXGu_Fj6haCzivxiARVDPGL/exec';

// ========== グローバル変数 ==========
let participants = [];
let tables = {};
let waitingZone = [];
let paMembers = [];
let mcMembers = [];
let selectedCard = null;

// テーブルマスター優先順位
const MASTER_ROLES = ['代表世話人', '会計長', '筆頭副代表世話人', '副会計', '副代表世話人', '世話人'];

// テーブルマトリックス関連
let layoutCandidates = [];
let selectedLayout = null;
let currentLayout = null;

// 印刷関連
let currentPaperSize = 'a3';

// ========== データ管理関数 ==========
function initTables() {
  tables = {};
  waitingZone = [];
  currentLayout = null;
}

function setParticipants(rawParticipants) {
  console.log('受信データ（1件目）:', rawParticipants[0]);

  participants = rawParticipants.map(p => {
    const affiliation = String(p.affiliation || p.所属 || '').trim();

    // ★修正: APIから受け取ったcategoryをそのまま使用
    // category が無い場合のみ従来ロジックでフォールバック
    let category = String(p.category || '').trim();
    if (!category) {
      const isGuest = !affiliation;
      const isOtherVenue = affiliation && affiliation !== '福岡飯塚';
      category = isGuest ? 'ゲスト' : (isOtherVenue ? '他会場' : '会員');
    }

    return {
      id: p.id || p.ID || Math.random().toString(36),
      name: p.name || p.氏名 || '名前なし',
      category: category,
      role: p.role || p.役割 || '',
      team: p.team || p.チーム || '',
      affiliation: affiliation,
      business: p.business || p.営業内容 || '',
      assignedTable: 'waiting'
    };
  });

  // PA・司会は専用ゾーンへ
  paMembers = participants.filter(p => p.role === 'PA');
  mcMembers = participants.filter(p => p.role === '事務局長');

  // 待機ゾーンには PA・司会を除いたメンバー
  waitingZone = participants.filter(p => p.role !== 'PA' && p.role !== '事務局長');
}

function generateTables(count, minCap, maxCap) {
  tables = {};
  for (let i = 0; i < count; i++) {
    const tableId = String.fromCharCode(65 + i);
    tables[tableId] = {
      id: tableId,
      master: null,
      members: Array(maxCap - 1).fill(null),
      minCap,
      maxCap
    };
  }
}

function getTableCount(tableId) {
  const table = tables[tableId];
  const memberCount = table.members.filter(m => m !== null).length;
  return (table.master ? 1 : 0) + memberCount;
}

function findEmptySeatIndex(tableId) {
  const table = tables[tableId];
  for (let i = 0; i < table.members.length; i++) {
    if (table.members[i] === null) {
      return i;
    }
  }
  return -1;
}

// ========== カード生成 ==========
function createPersonCard(person, fromTable, fromSeat) {
  const card = document.createElement('div');
  card.className = 'person-card';

  const isMaster = MASTER_ROLES.includes(person.role);
  const isGuest = person.category === 'ゲスト';
  const isOtherVenue = person.category === '他会場';

  if (isGuest) {
    card.classList.add('guest');
  } else if (isOtherVenue) {
    card.classList.add('other-venue');
  } else {
    card.classList.add('jikai');
  }

  card.dataset.personId = person.id;
  card.dataset.fromTable = fromTable;
  card.dataset.fromSeat = fromSeat;
  card.draggable = true;

  if (fromSeat === 'master') {
    const title = document.createElement('div');
    title.className = 'person-title';
    title.textContent = '👑 テーブルマスター';
    card.appendChild(title);
  } else if (isOtherVenue) {
    const title = document.createElement('div');
    title.className = 'person-title';
    title.textContent = `${person.affiliation}会場`;
    card.appendChild(title);
  } else if (isGuest) {
    const title = document.createElement('div');
    title.className = 'person-title';
    title.textContent = 'ゲスト';
    card.appendChild(title);
  }

  const name = document.createElement('div');
  name.className = 'person-name';

  if (isGuest || isOtherVenue) {
    name.textContent = `${person.name} 様`;
  } else {
    name.textContent = person.name;
  }

  card.appendChild(name);

  card.addEventListener('dragstart', handleDragStart);
  card.addEventListener('dragend', handleDragEnd);
  card.addEventListener('click', handleCardClick);

  return card;
}

// ========== バリデーション ==========
function validateDrop(personId, toTable, toSeat, toZone, fromTable, fromSeat) {
  const person = participants.find(p => p.id === personId);
  if (!person) return { valid: false, message: '参加者が見つかりません' };

  // テーブルマスター席は役割があればOK
  if (toSeat === 'master' && (!person.role || person.role.trim() === '')) {
    return { valid: false, message: 'テーブルマスター席には役割のある人のみ配置できます' };
  }

  if (toZone === 'waiting') {
    return { valid: true };
  }

  if (toTable && toSeat === 'member') {
    const table = tables[toTable];
    if (getTableCount(toTable) >= table.maxCap) {
      return { valid: false, message: `テーブル${toTable}は満席です` };
    }
  }

  return { valid: true };
}

function movePerson(personId, fromTable, fromSeat, toTable, toSeat, toZone, dropZone) {
  const person = participants.find(p => p.id === personId);
  if (!person) return;

  // 移動先の重複チェック
  if (toZone === 'waiting') {
    const alreadyExists = waitingZone.some(p => p.id === personId);
    if (alreadyExists) {
      console.warn('既に待機ゾーンに存在:', personId);
      return;
    }
  }

  if (toZone === 'pa') {
    const alreadyExists = paMembers.some(p => p.id === personId);
    if (alreadyExists) {
      console.warn('既にPA枠に存在:', personId);
      return;
    }
  }

  if (toZone === 'mc') {
    const alreadyExists = mcMembers.some(p => p.id === personId);
    if (alreadyExists) {
      console.warn('既に司会枠に存在:', personId);
      return;
    }
  }

  // 移動元からの削除処理
  if (fromTable && fromTable !== 'waiting' && fromTable !== 'pa' && fromTable !== 'mc') {
    const table = tables[fromTable];
    if (fromSeat === 'master') {
      table.master = null;
    } else {
      const idx = table.members.findIndex(m => m && m.id === personId);
      if (idx !== -1) {
        table.members[idx] = null;
      }
    }
  } else if (fromTable === 'waiting') {
    const idx = waitingZone.findIndex(p => p.id === personId);
    if (idx !== -1) waitingZone.splice(idx, 1);
  } else if (fromTable === 'pa') {
    const idx = paMembers.findIndex(p => p.id === personId);
    if (idx !== -1) paMembers.splice(idx, 1);
  } else if (fromTable === 'mc') {
    const idx = mcMembers.findIndex(p => p.id === personId);
    if (idx !== -1) mcMembers.splice(idx, 1);
  }

  // 移動先への配置処理
  if (toZone === 'waiting') {
    waitingZone.push(person);
    person.assignedTable = 'waiting';
  } else if (toZone === 'pa') {
    paMembers = [person]; // 1人だけ
    person.assignedTable = 'pa';
  } else if (toZone === 'mc') {
    mcMembers = [person]; // 1人だけ
    person.assignedTable = 'mc';
  } else if (toTable) {
    const table = tables[toTable];
    if (toSeat === 'master') {
      table.master = person;
    } else {
      const emptyIdx = findEmptySeatIndex(toTable);
      if (emptyIdx !== -1) {
        table.members[emptyIdx] = person;
      }
    }
    person.assignedTable = toTable;
  }
}
function swapPersons(personId1, fromTable1, fromSeat1, personId2, fromTable2, fromSeat2) {
  const person1 = participants.find(p => p.id === personId1);
  const person2 = participants.find(p => p.id === personId2);
  if (!person1 || !person2) return;

  const table1 = tables[fromTable1];
  const table2 = tables[fromTable2];

  if (!table1 || !table2) return;

  const idx1 = fromSeat1 === 'member'
    ? table1.members.findIndex(m => m && m.id === personId1)
    : null;

  const idx2 = fromSeat2 === 'member'
    ? table2.members.findIndex(m => m && m.id === personId2)
    : null;

  if (fromSeat1 === 'master' && fromSeat2 === 'master') {
    const tmp = table1.master;
    table1.master = table2.master;
    table2.master = tmp;
  } else if (fromSeat1 === 'member' && fromSeat2 === 'member') {
    if (idx1 === -1 || idx2 === -1) return;
    const tmp = table1.members[idx1];
    table1.members[idx1] = table2.members[idx2];
    table2.members[idx2] = tmp;
  } else if (fromSeat1 === 'master' && fromSeat2 === 'member') {
    if (idx2 === -1) return;
    const tmp = table1.master;
    table1.master = table2.members[idx2];
    table2.members[idx2] = tmp;
  } else if (fromSeat1 === 'member' && fromSeat2 === 'master') {
    if (idx1 === -1) return;
    const tmp = table2.master;
    table2.master = table1.members[idx1];
    table1.members[idx1] = tmp;
  }

  if (table1.master) table1.master.assignedTable = fromTable1;
  table1.members.forEach(m => { if (m) m.assignedTable = fromTable1; });
  if (table2.master) table2.master.assignedTable = fromTable2;
  table2.members.forEach(m => { if (m) m.assignedTable = fromTable2; });
}

// ========== 自動配席 ==========
function autoSeat(mode) {
  if (!confirm(`${mode === 'team' ? '同チーム優先' : '営業相性優先'}で自動配席しますか？\n現在の配席はリセットされます。`)) {
    return;
  }

  Object.values(tables).forEach(table => {
    table.master = null;
    table.members.fill(null);
  });

  waitingZone = participants.filter(p => p.role !== 'PA' && p.role !== '事務局長');

  const masters = [];
  MASTER_ROLES.forEach(role => {
    const roleMasters = waitingZone.filter(p => p.role === role);
    masters.push(...roleMasters);
  });

  const guests = waitingZone.filter(p => p.category === 'ゲスト');
  const otherVenue = waitingZone.filter(p => p.category === '他会場');
  const regularMembers = waitingZone.filter(p =>
    !MASTER_ROLES.includes(p.role) &&
    p.category !== 'ゲスト' &&
    p.category !== '他会場'
  );

  const tableIds = Object.keys(tables).sort();

  tableIds.forEach((tableId, idx) => {
    if (masters[idx]) {
      tables[tableId].master = masters[idx];
      masters[idx].assignedTable = tableId;
    }
  });

  distributeGuests(guests, tableIds);
  distributeOtherVenue(otherVenue, tableIds);

  const remainingMasters = masters.slice(tableIds.length);
  const allRegularMembers = [...regularMembers, ...remainingMasters];

  if (mode === 'team') {
    autoSeatByTeam(allRegularMembers, tableIds);
  } else {
    autoSeatByBusiness(allRegularMembers, tableIds);
  }

  balanceTables(tableIds);

  waitingZone = [];

　// ← ここに追加
  document.getElementById('sync-btn').style.display = 'block';

  renderAll();
  initDragDrop();
  alert('自動配席が完了しました！');
}

function distributeGuests(guests, tableIds) {
  const guestTables = tableIds.slice(0, 6);
  let guestIndex = 0;

  for (let round = 0; round < 2; round++) {
    guestTables.forEach(tableId => {
      if (guestIndex < guests.length) {
        const table = tables[tableId];
        if (getTableCount(tableId) < table.maxCap) {
          const emptyIdx = findEmptySeatIndex(tableId);
          if (emptyIdx !== -1) {
            table.members[emptyIdx] = guests[guestIndex];
            guests[guestIndex].assignedTable = tableId;
            guestIndex++;
          }
        }
      }
    });
  }

  let tableIndex = 0;
  while (guestIndex < guests.length) {
    const tableId = guestTables[tableIndex % guestTables.length];
    const table = tables[tableId];

    if (getTableCount(tableId) < table.maxCap) {
      const emptyIdx = findEmptySeatIndex(tableId);
      if (emptyIdx !== -1) {
        table.members[emptyIdx] = guests[guestIndex];
        guests[guestIndex].assignedTable = tableId;
        guestIndex++;
      }
    }

    tableIndex++;
    if (tableIndex > guestTables.length * 10) break;
  }
}

function distributeOtherVenue(otherVenue, tableIds) {
  const shuffled = [...otherVenue].sort(() => Math.random() - 0.5);
  let venueIndex = 0;

  tableIds.forEach(tableId => {
    if (venueIndex < shuffled.length) {
      const table = tables[tableId];
      if (getTableCount(tableId) < table.maxCap) {
        const emptyIdx = findEmptySeatIndex(tableId);
        if (emptyIdx !== -1) {
          table.members[emptyIdx] = shuffled[venueIndex];
          shuffled[venueIndex].assignedTable = tableId;
          venueIndex++;
        }
      }
    }
  });

  let tableIndex = 0;
  while (venueIndex < shuffled.length) {
    const tableId = tableIds[tableIndex % tableIds.length];
    const table = tables[tableId];

    if (getTableCount(tableId) < table.maxCap) {
      const emptyIdx = findEmptySeatIndex(tableId);
      if (emptyIdx !== -1) {
        table.members[emptyIdx] = shuffled[venueIndex];
        shuffled[venueIndex].assignedTable = tableId;
        venueIndex++;
      }
    }

    tableIndex++;
    if (tableIndex > tableIds.length * 10) break;
  }
}

function autoSeatByTeam(members, tableIds) {
  const teamGroups = {};
  members.forEach(person => {
    const team = person.team || '未所属';
    if (!teamGroups[team]) teamGroups[team] = [];
    teamGroups[team].push(person);
  });

  let currentTableIdx = 0;

  Object.values(teamGroups).forEach(teamMembers => {
    teamMembers.forEach(person => {
      let attempts = 0;
      while (attempts < tableIds.length) {
        const tableId = tableIds[currentTableIdx];
        const table = tables[tableId];

        if (getTableCount(tableId) < table.maxCap) {
          const emptyIdx = findEmptySeatIndex(tableId);
          if (emptyIdx !== -1) {
            table.members[emptyIdx] = person;
            person.assignedTable = tableId;
            break;
          }
        }

        currentTableIdx = (currentTableIdx + 1) % tableIds.length;
        attempts++;
      }

      currentTableIdx = (currentTableIdx + 1) % tableIds.length;
    });
  });
}

function autoSeatByBusiness(members, tableIds) {
  const businessGroups = {};
  members.forEach(person => {
    const business = person.business || '未記入';
    if (!businessGroups[business]) businessGroups[business] = [];
    businessGroups[business].push(person);
  });

  const businessTypes = Object.keys(businessGroups);
  let currentTableIdx = 0;

  businessTypes.forEach(businessType => {
    const businessMembers = businessGroups[businessType];

    businessMembers.forEach(person => {
      let attempts = 0;
      while (attempts < tableIds.length) {
        const tableId = tableIds[currentTableIdx];
        const table = tables[tableId];

        if (getTableCount(tableId) < table.maxCap) {
          const emptyIdx = findEmptySeatIndex(tableId);
          if (emptyIdx !== -1) {
            table.members[emptyIdx] = person;
            person.assignedTable = tableId;
            break;
          }
        }

        currentTableIdx = (currentTableIdx + 1) % tableIds.length;
        attempts++;
      }

      currentTableIdx = (currentTableIdx + 1) % tableIds.length;
    });
  });
}

function balanceTables(tableIds) {
  tableIds.forEach(tableId => {
    const table = tables[tableId];
    let currentCount = getTableCount(tableId);

    while (currentCount < table.minCap) {
      let maxTableId = null;
      let maxCount = 0;

      tableIds.forEach(tid => {
        const count = getTableCount(tid);
        if (count > maxCount && count > tables[tid].minCap) {
          maxCount = count;
          maxTableId = tid;
        }
      });

      if (!maxTableId) break;

      const fromTable = tables[maxTableId];
      let personToMove = null;
      for (let i = fromTable.members.length - 1; i >= 0; i--) {
        if (fromTable.members[i] !== null) {
          personToMove = fromTable.members[i];
          fromTable.members[i] = null;
          break;
        }
      }

      if (personToMove) {
        const emptyIdx = findEmptySeatIndex(tableId);
        if (emptyIdx !== -1) {
          table.members[emptyIdx] = personToMove;
          personToMove.assignedTable = tableId;
          currentCount++;
        } else {
          break;
        }
      } else {
        break;
      }
    }
  });
}

// ========== 初期化 ==========
window.addEventListener('DOMContentLoaded', () => {
  initTables();

  document.getElementById('load-btn').addEventListener('click', loadFromSheet);
  document.getElementById('load-archive-btn').addEventListener('click', loadFromArchive);
  document.getElementById('reset-btn').addEventListener('click', resetSeats);
  document.getElementById('generateBtn').addEventListener('click', generateSeating);

  // 差分モーダルのボタン
  document.getElementById('diffCancelBtn').addEventListener('click', () => {
    document.getElementById('diffModal').style.display = 'none';
  });
  document.getElementById('diffApplyBtn').addEventListener('click', applyArchiveAssignments);
  
  // 📤 座席反映ボタン
document.getElementById('sync-btn').addEventListener('click', async () => {
  const button = document.getElementById('sync-btn');
  button.disabled = true;
  button.textContent = '反映中...';

  try {
    // LIFFから userId を取得
    let userId = '';
    if (typeof liff !== 'undefined' && liff.isLoggedIn()) {
      const profile = await liff.getProfile();
      userId = profile.userId;
    }

    // userId がなければダミーを送る（テスト用）
    if (!userId) {
      userId = 'test_user_' + Date.now();
    }

    // 配席データを収集
    const assignments = collectAssignments();

    if (assignments.length === 0) {
      alert('配席データがありません。先に配席を行ってください。');
      return;
    }

    // 確認ダイアログ
    if (!confirm(`${assignments.length}名の座席情報を反映します。よろしいですか？`)) {
      return;
    }

    const payload = {
      eventKey: window.currentEventKey || '',
      userId: userId,
      assignments: assignments,
      layout: currentLayout || []  // テーブル配列パターン
    };

    // Note: Content-Typeヘッダーを省略してCORSプリフライトを回避
    // GAS側でJSON.parse(e.postData.contents)で解析可能
    const res = await fetch(`${API_URL}?action=syncSeats`, {
      method: 'POST',
      body: JSON.stringify(payload)
    });

    if (!res.ok) throw new Error('HTTPエラー: ' + res.status);

    const json = await res.json();

    if (json.success) {
      alert(`反映完了しました！\n\n${json.archivedCount}名の座席情報を保存しました。`);
    } else {
      alert('エラー: ' + (json.error || 'unknown'));
    }

  } catch (err) {
    console.error(err);
    alert('通信エラー: ' + err.message);
  } finally {
    button.disabled = false;
    button.textContent = '座席反映';
  }
});

/**
 * 現在の配席状態からassignments配列を生成
 */
function collectAssignments() {
  const assignments = [];
  const addedIds = new Set();  // ★重複防止用

  // 各テーブルから収集
  Object.keys(tables).forEach(tableId => {
    const table = tables[tableId];

    // マスター
    if (table.master) {
      const person = participants.find(p => p.id === table.master.id);
      if (person && !addedIds.has(person.id)) {  // ★重複チェック
        addedIds.add(person.id);
        assignments.push({
          id: person.id,
          name: person.name,
          category: person.category,
          affiliation: person.affiliation || '',
          role: person.role || '',
          team: person.team || '',
          table: tableId,
          seat: 0
        });
      }
    }

    // メンバー
    table.members.forEach((member, idx) => {
      if (member) {
        const person = participants.find(p => p.id === member.id);
        if (person && !addedIds.has(person.id)) {  // ★重複チェック
          addedIds.add(person.id);
          assignments.push({
            id: person.id,
            name: person.name,
            category: person.category,
            affiliation: person.affiliation || '',
            role: person.role || '',
            team: person.team || '',
            table: tableId,
            seat: idx + 1
          });
        }
      }
    });
  });

  // PA
  paMembers.forEach((person, idx) => {
    if (!addedIds.has(person.id)) {  // ★重複チェック
      addedIds.add(person.id);
      assignments.push({
        id: person.id,
        name: person.name,
        category: person.category,
        affiliation: person.affiliation || '',
        role: person.role || '',
        team: person.team || '',
        table: 'PA',
        seat: idx
      });
    }
  });

  // MC
  mcMembers.forEach((person, idx) => {
    if (!addedIds.has(person.id)) {  // ★重複チェック
      addedIds.add(person.id);
      assignments.push({
        id: person.id,
        name: person.name,
        category: person.category,
        affiliation: person.affiliation || '',
        role: person.role || '',
        team: person.team || '',
        table: 'MC',
        seat: idx
      });
    }
  });

  return assignments;
}


  
  document.getElementById('autoBusinessBtnModal').addEventListener('click', () => {
    document.getElementById('autoSeatModal').style.display = 'none';
    autoSeat('business');
  });

  document.getElementById('minCapacity').addEventListener('change', updateConfigSummary);
  document.getElementById('maxCapacity').addEventListener('change', updateConfigSummary);
  document.getElementById('tableCount').addEventListener('change', () => {
    selectedLayout = null;
    rebuildLayoutCandidates();
    updateConfigSummary();
  });

  rebuildLayoutCandidates();
});

// ========== スプレッドシート読込 ==========
async function loadFromSheet() {
  const statusEl = document.getElementById('status');
  const summaryEl = document.getElementById('summary');
  const previewEl = document.getElementById('preview');
  const configSection = document.getElementById('configSection');

  statusEl.textContent = '読み込み中...';
  summaryEl.innerHTML = '';
  previewEl.innerHTML = '';

  try {
    const res = await fetch(`${API_URL}?action=getSeatingParticipants`);
    if (!res.ok) throw new Error('HTTPエラー: ' + res.status);

    const json = await res.json();
    if (!json.success) throw new Error(json.error || 'unknown error');

    // イベント情報をグローバルに保持
    window.currentEventKey = json.eventKey || '';
    window.currentEventName = json.eventName || '';

    setParticipants(json.participants || []);
    statusEl.textContent = '読み込み完了！';

    const total = participants.length;
    const members = participants.filter(p => p.category === '会員').length;
    const guests = participants.filter(p => p.category === 'ゲスト').length;
    const others = participants.filter(p => p.category === '他会場').length;
    const masters = participants.filter(p => MASTER_ROLES.includes(p.role)).length;

    summaryEl.innerHTML = `
      <ul class="summary-list">
        <li>総参加者数：<strong>${total}</strong> 名</li>
        <li>会員：<strong>${members}</strong> 名</li>
        <li>ゲスト：<strong>${guests}</strong> 名</li>
        <li>他会場：<strong>${others}</strong> 名</li>
        <li>テーブルマスター：<strong>${masters}</strong> 名</li>
        <li id="tableCountLi" style="display: none;">テーブル数：<strong id="tableCountValue">0</strong> 卓</li>
      </ul>
    `;

    const first10 = participants.slice(0, 10);
    if (first10.length === 0) {
      previewEl.innerHTML = '<p>参加者が存在しません。</p>';
    } else {
      const rows = first10.map(p => {
        const isMaster = MASTER_ROLES.includes(p.role);
        const badge = p.category === 'ゲスト' ? 'guest' : isMaster ? 'master' : p.category === '他会場' ? 'other' : 'member';
        const badgeLabel = isMaster ? 'TM' : p.category === 'ゲスト' ? 'G' : p.category === '他会場' ? '他' : 'M';
        return `<div class="person-row person-${badge}"><span class="person-badge">${badgeLabel}</span><span class="person-name">${p.name}</span><span class="person-meta">${p.affiliation || ''} / ${p.team || ''}</span></div>`;
      }).join('');
      previewEl.innerHTML = rows;
    }

    configSection.style.display = 'block';
    rebuildLayoutCandidates();
    updateConfigSummary();

    // 「続きから読み込む」ボタンを表示
    document.getElementById('load-archive-btn').style.display = 'inline-block';
    // リセットボタンを表示
    document.getElementById('reset-btn').style.display = 'inline-block';
  } catch (err) {
    console.error(err);
    statusEl.textContent = '読み込みに失敗しました：' + err.message;
  }
}

// ========== アーカイブ読み込み ==========
// 差分データを一時保存
let archiveDiff = null;
let archiveData = null;

async function loadFromArchive() {
  const eventKey = window.currentEventKey;
  if (!eventKey) {
    alert('先にスプレッドシートからデータを読み込んでください。');
    return;
  }

  if (participants.length === 0) {
    alert('参加者データがありません。先にスプレッドシートから読み込んでください。');
    return;
  }

  const statusEl = document.getElementById('status');
  statusEl.textContent = 'アーカイブを読み込み中...';

  try {
    const res = await fetch(`${API_URL}?action=getSeatingArchive&eventKey=${encodeURIComponent(eventKey)}`);
    if (!res.ok) throw new Error('HTTPエラー: ' + res.status);

    const json = await res.json();
    if (!json.success) throw new Error(json.error || 'アーカイブの取得に失敗しました');

    if (!json.assignments || json.assignments.length === 0) {
      alert('保存された配席がありません。');
      statusEl.textContent = '読み込み完了';
      return;
    }

    archiveData = json;

    // 差分計算
    archiveDiff = calculateDiff(participants, json.assignments);

    // 差分モーダル表示
    showDiffModal(archiveDiff, json);
    statusEl.textContent = '読み込み完了';

  } catch (err) {
    console.error(err);
    alert('アーカイブ読み込みエラー: ' + err.message);
    statusEl.textContent = 'エラー: ' + err.message;
  }
}

function calculateDiff(currentParticipants, archivedAssignments) {
  const currentIds = new Set(currentParticipants.map(p => p.id));
  const archivedIds = new Set(archivedAssignments.map(a => a.id));

  // 配席済み（両方に存在）
  const matched = archivedAssignments.filter(a => currentIds.has(a.id));

  // 新規参加者（現在のみ存在）
  const newParticipants = currentParticipants.filter(p => !archivedIds.has(p.id));

  // 欠席者（アーカイブのみ存在）
  const absentees = archivedAssignments.filter(a => !currentIds.has(a.id));

  return { matched, newParticipants, absentees };
}

function showDiffModal(diff, archiveJson) {
  const modal = document.getElementById('diffModal');
  const savedAtEl = document.getElementById('diffSavedAt');
  const matchedCountEl = document.getElementById('diffMatchedCount');
  const newCountEl = document.getElementById('diffNewCount');
  const absentCountEl = document.getElementById('diffAbsentCount');
  const detailsEl = document.getElementById('diffDetails');

  // 保存日時
  savedAtEl.textContent = archiveJson.confirmedAt || '-';

  // カウント
  matchedCountEl.textContent = diff.matched.length;
  newCountEl.textContent = diff.newParticipants.length;
  absentCountEl.textContent = diff.absentees.length;

  // 詳細リスト
  let detailsHtml = '';

  if (diff.newParticipants.length > 0) {
    detailsHtml += `
      <div class="diff-section">
        <div class="diff-section-title new">新規追加（待機ゾーンへ）</div>
        <ul class="diff-section-list">
          ${diff.newParticipants.map(p => `<li>${p.name}（${p.category}）</li>`).join('')}
        </ul>
      </div>
    `;
  }

  if (diff.absentees.length > 0) {
    detailsHtml += `
      <div class="diff-section">
        <div class="diff-section-title absent">欠席（配席から除外）</div>
        <ul class="diff-section-list">
          ${diff.absentees.map(a => `<li>${a.name}</li>`).join('')}
        </ul>
      </div>
    `;
  }

  detailsEl.innerHTML = detailsHtml;
  modal.style.display = 'flex';
}

function applyArchiveAssignments() {
  if (!archiveDiff || !archiveData) {
    alert('アーカイブデータがありません。');
    return;
  }

  const modal = document.getElementById('diffModal');
  modal.style.display = 'none';

  const assignments = archiveData.assignments;
  const minCap = parseInt(document.getElementById('minCapacity').value);
  const maxCap = parseInt(document.getElementById('maxCapacity').value);

  // テーブルIDを抽出してテーブル数を決定
  const tableIds = [...new Set(
    assignments
      .map(a => a.table)
      .filter(t => t && !['PA', 'MC', 'waiting'].includes(t))
  )].sort();

  const tableCount = tableIds.length;

  // テーブル数を設定フォームに反映
  document.getElementById('tableCount').value = tableCount;

  // テーブルを生成
  generateTables(tableCount, minCap, maxCap);

  // 各参加者をリセット
  participants.forEach(p => {
    p.assignedTable = 'waiting';
  });
  paMembers = [];
  mcMembers = [];
  waitingZone = [];

  // アーカイブの配席を適用
  assignments.forEach(a => {
    const person = participants.find(p => p.id === a.id);
    if (!person) return; // 欠席者はスキップ

    person.assignedTable = a.table;

    if (a.table === 'PA') {
      paMembers.push(person);
    } else if (a.table === 'MC') {
      mcMembers.push(person);
    } else if (a.table === 'waiting') {
      // 待機ゾーン
    } else if (tables[a.table]) {
      if (a.seat === 0) {
        tables[a.table].master = person;
      } else {
        const emptyIdx = findEmptySeatIndex(a.table);
        if (emptyIdx !== -1) {
          tables[a.table].members[emptyIdx] = person;
        }
      }
    }
  });

  // 新規参加者は待機ゾーンへ
  archiveDiff.newParticipants.forEach(p => {
    if (p.role === 'PA') {
      paMembers.push(p);
      p.assignedTable = 'PA';
    } else if (p.role === '事務局長') {
      mcMembers.push(p);
      p.assignedTable = 'MC';
    } else {
      waitingZone.push(p);
      p.assignedTable = 'waiting';
    }
  });

  // 待機ゾーンに残っている人を追加（配席されなかった人）
  participants.forEach(p => {
    if (p.assignedTable === 'waiting' &&
        !waitingZone.includes(p) &&
        p.role !== 'PA' &&
        p.role !== '事務局長') {
      waitingZone.push(p);
    }
  });

  // アーカイブからレイアウトを復元（保存されていれば）
  document.getElementById('tableCount').value = tableCount;

  if (archiveData.layout && archiveData.layout.length > 0) {
    // 保存されたレイアウトを使用
    currentLayout = archiveData.layout;
    selectedLayout = archiveData.layout;
  } else {
    // レイアウトがない場合は候補から選択
    rebuildLayoutCandidates();
    if (tableCount >= 6 && layoutCandidates.length > 0) {
      selectedLayout = layoutCandidates[0];
      currentLayout = layoutCandidates[0];
    } else {
      currentLayout = [tableCount];
      selectedLayout = currentLayout;
    }
  }

  // UIを更新
  document.getElementById('mainContainer').style.display = 'grid';
  document.getElementById('configSection').style.display = 'none';
  document.getElementById('sync-btn').style.display = 'inline-block';
  document.getElementById('print-preview-btn').style.display = 'inline-block';
  document.getElementById('preview').style.display = 'none';

  renderAll();
  initDragDrop();

  // ステータス更新
  document.getElementById('status').textContent =
    `アーカイブから読み込み完了（配席済み: ${archiveDiff.matched.length}名, 新規: ${archiveDiff.newParticipants.length}名）`;

  // 後片付け
  archiveDiff = null;
  archiveData = null;
}

function updateConfigSummary() {
  const minCap = parseInt(document.getElementById('minCapacity').value);
  const maxCap = parseInt(document.getElementById('maxCapacity').value);
  const tableCount = parseInt(document.getElementById('tableCount').value);
  const summaryEl = document.getElementById('configSummary');
  const generateBtn = document.getElementById('generateBtn');

  const totalParticipants = participants.length - paMembers.length - mcMembers.length;
  const minRequired = minCap * tableCount;
  const maxCapacity = maxCap * tableCount;
  let canGenerate = totalParticipants >= minRequired && totalParticipants <= maxCapacity;

  summaryEl.innerHTML = `総参加者数：<strong>${totalParticipants}</strong>人（PA・司会除く）<br>収容可能人数：<strong>${minRequired}〜${maxCapacity}</strong>人`;

  if (canGenerate) {
    summaryEl.style.color = '#065f46';
  } else {
    summaryEl.style.color = '#dc2626';
    summaryEl.innerHTML += totalParticipants < minRequired
      ? '<br>⚠️ テーブル数が多すぎます'
      : '<br>⚠️ テーブル数が足りません';
  }

  const requireLayout = tableCount >= 6;
  if (requireLayout && !selectedLayout) {
    generateBtn.disabled = true;
  } else {
    generateBtn.disabled = !canGenerate;
  }
}

function rebuildLayoutCandidates() {
  const section = document.getElementById('layoutSection');
  const container = document.getElementById('layoutOptions');
  if (!section || !container) return;

  const tableCount = parseInt(document.getElementById('tableCount').value);
  container.innerHTML = '';
  layoutCandidates = [];
  selectedLayout = null;

  if (isNaN(tableCount) || tableCount < 6) {
    section.style.display = 'none';
    return;
  }

  const rawLayouts = [];

  function dfs(remaining, rows, current) {
    if (rows > 5) return;
    if (remaining === 0) {
      if (rows >= 2) {
        rawLayouts.push(current.slice());
      }
      return;
    }

    [3, 2].forEach(size => {
      if (remaining - size < 0) return;
      if (rows === 0 && size !== 3) return;

      current.push(size);
      dfs(remaining - size, rows + 1, current);
      current.pop();
    });
  }

  dfs(tableCount, 0, []);

  if (!rawLayouts.length) {
    section.style.display = 'none';
    return;
  }

  const scored = rawLayouts.map(arr => {
    const rows = arr.length;
    const maxCols = Math.max(...arr);
    const minCols = Math.min(...arr);
    const diff = maxCols - minCols;
    const rowsFrom3 = Math.abs(rows - 3);
    const score = rowsFrom3 * 3 + diff * 2;
    return { arr, score, rows, maxCols, diff };
  });

  scored.sort((a, b) => a.score - b.score);

  const used = new Set();
  const picked = [];
  for (const s of scored) {
    const key = s.arr.join('-');
    if (used.has(key)) continue;
    used.add(key);
    picked.push(s.arr);
    if (picked.length === 3) break;
  }

  layoutCandidates = picked;
  renderLayoutOptions(tableCount);
  section.style.display = 'block';
}

function renderLayoutOptions(tableCount) {
  const container = document.getElementById('layoutOptions');
  if (!container) return;
  container.innerHTML = '';

  if (!layoutCandidates.length) return;

  const labelChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

  layoutCandidates.forEach((layout, idx) => {
    const card = document.createElement('div');
    card.className = 'layout-card';
    card.dataset.index = idx;

    const badge = document.createElement('div');
    badge.className = 'layout-badge';
    badge.textContent = String.fromCharCode(65 + idx) + '案';
    card.appendChild(badge);

    const mini = document.createElement('div');
    mini.className = 'layout-mini';

    let tableIndex = 0;
    layout.forEach(rowSize => {
      const row = document.createElement('div');
      row.className = 'layout-row';

      for (let i = 0; i < rowSize; i++) {
        const dot = document.createElement('div');
        dot.className = 'layout-dot';
        const label = labelChars[tableIndex] || '?';
        dot.textContent = label;
        row.appendChild(dot);
        tableIndex++;
      }

      mini.appendChild(row);
    });

    card.appendChild(mini);

    const label = document.createElement('div');
    label.className = 'layout-label';
    label.textContent = `${layout.join(' × ')}（${tableCount}卓）`;
    card.appendChild(label);

    const check = document.createElement('div');
    check.className = 'layout-check';
    check.textContent = 'クリックで選択';
    card.appendChild(check);

    card.addEventListener('click', () => {
      selectedLayout = layout;
      currentLayout = layout.slice();
      document.querySelectorAll('.layout-card').forEach(c => c.classList.remove('selected'));
      card.classList.add('selected');
      updateConfigSummary();
    });

    container.appendChild(card);
  });
}

function generateSeating() {
  const tableCount = parseInt(document.getElementById('tableCount').value);
  const minCap = parseInt(document.getElementById('minCapacity').value);
  const maxCap = parseInt(document.getElementById('maxCapacity').value);

  generateTables(tableCount, minCap, maxCap);

  document.getElementById('configSection').style.display = 'none';
  document.getElementById('mainContainer').style.display = 'grid';

  // ボタン表示
  document.getElementById('sync-btn').style.display = 'inline-block';
  document.getElementById('print-preview-btn').style.display = 'inline-block';

  renderAll();
  initDragDrop();

  // 自動配席選択モーダルを表示
  document.getElementById('autoSeatModal').style.display = 'flex';
}

function resetSeats() {
  if (!confirm('配席をすべてリセットしますか？')) return;

  initTables();

  waitingZone = participants.filter(p => p.role !== 'PA' && p.role !== '事務局長');
  participants.forEach(p => {
    if (p.role !== 'PA' && p.role !== '事務局長') {
      p.assignedTable = 'waiting';
    }
  });

  document.getElementById('configSection').style.display = 'block';
  document.getElementById('mainContainer').style.display = 'none';

  // ボタン非表示
  document.getElementById('sync-btn').style.display = 'none';
  document.getElementById('print-preview-btn').style.display = 'none';

  rebuildLayoutCandidates();
  updateConfigSummary();
  renderAll();
  initDragDrop();
}

// ========== レンダリング ==========
function renderAll() {
  // テーブル数の表示更新
  const tableCountLi = document.getElementById('tableCountLi');
  const tableCountValue = document.getElementById('tableCountValue');
  if (tableCountLi && tableCountValue && Object.keys(tables).length > 0) {
    tableCountValue.textContent = Object.keys(tables).length;
    tableCountLi.style.display = 'list-item';
  }

  renderTables();
  renderWaitingZone();
  renderPAZone();
  renderMCZone();
}

function renderTables() {
  const grid = document.getElementById('tablesGrid');
  if (!grid) return;
  grid.innerHTML = '';

  const tableIds = Object.keys(tables).sort();
  
  if (!currentLayout || currentLayout.length === 0) {
    const row = document.createElement('div');
    row.className = 'table-row';
    
    tableIds.forEach(tableId => {
      row.appendChild(createTableCard(tableId));
    });
    
    grid.appendChild(row);
    return;
  }

  let tableIndex = 0;
  currentLayout.forEach(rowSize => {
    const row = document.createElement('div');
    row.className = 'table-row';

    for (let i = 0; i < rowSize && tableIndex < tableIds.length; i++) {
      const tableId = tableIds[tableIndex];
      row.appendChild(createTableCard(tableId));
      tableIndex++;
    }

    grid.appendChild(row);
  });
}

function createTableCard(tableId) {
  const table = tables[tableId];
  const card = document.createElement('div');
  card.className = 'table-card';
  card.dataset.table = table.id;

  const totalCount = getTableCount(table.id);
  const memberCount = (table.master ? 1 : 0) + table.members.filter(m => m && m.category === '会員').length;
  const guestCount = table.members.filter(m => m && m.category === 'ゲスト').length;
  const otherCount = table.members.filter(m => m && m.category === '他会場').length;

  const header = document.createElement('div');
  header.className = 'table-header';
  header.innerHTML = `<div class="table-title">👥 テーブル ${table.id}</div><div class="table-count">${totalCount}/${table.maxCap}人</div>`;
  card.appendChild(header);

  const seatsGrid = document.createElement('div');
  seatsGrid.className = 'seats-grid';

  const masterSeat = document.createElement('div');
  masterSeat.className = 'seat master';
  masterSeat.dataset.table = table.id;
  masterSeat.dataset.seat = 'master';

  if (table.master) {
    masterSeat.classList.add('occupied');
    masterSeat.appendChild(createPersonCard(table.master, table.id, 'master'));
  } else {
    const label = document.createElement('div');
    label.className = 'master-label';
    label.textContent = '👑 TM';
    masterSeat.appendChild(label);
  }
  masterSeat.addEventListener('click', e => handleSeatClick(e, table.id, 'master', null));
  seatsGrid.appendChild(masterSeat);

  for (let i = 0; i < table.members.length; i++) {
    const seat = document.createElement('div');
    seat.className = 'seat';
    seat.dataset.table = table.id;
    seat.dataset.seat = 'member';
    seat.dataset.index = i;
    
    if (table.members[i]) {
      seat.classList.add('occupied');
      seat.appendChild(createPersonCard(table.members[i], table.id, 'member'));
    }
    
    seat.addEventListener('click', e => handleSeatClick(e, table.id, 'member', seat));
    seatsGrid.appendChild(seat);
  }

  card.appendChild(seatsGrid);

  const summary = document.createElement('div');
  summary.className = 'table-summary';
  summary.innerHTML = `<span>会員:${memberCount}</span><span>ゲスト:${guestCount}</span><span>他会場:${otherCount}</span>`;
  card.appendChild(summary);

  return card;
}

function renderWaitingZone() {
  const list = document.getElementById('waitingList');
  if (!list) return;
  list.innerHTML = '';

  const maxSlots = 5;

  for (let i = 0; i < maxSlots; i++) {
    const slot = document.createElement('div');
    slot.className = 'waiting-slot';
    slot.dataset.zone = 'waiting';

    const person = waitingZone[i];
    if (person) {
      slot.appendChild(createPersonCard(person, 'waiting', ''));
    } else {
      const placeholder = document.createElement('div');
      placeholder.className = 'waiting-placeholder';
      placeholder.textContent = '空き枠';
      slot.appendChild(placeholder);
    }

    list.appendChild(slot);
  }
}

function renderPAZone() {
  const list = document.getElementById('paList');
  if (!list) return;
  list.innerHTML = '';

  // 1スロット固定
  const slot = document.createElement('div');
  slot.className = 'special-slot';
  slot.dataset.zone = 'pa';

  const person = paMembers[0]; // 1人だけ
  if (person) {
    slot.appendChild(createPersonCard(person, 'pa', ''));
  } else {
    const placeholder = document.createElement('div');
    placeholder.className = 'special-placeholder';
    placeholder.textContent = '空き枠';
    slot.appendChild(placeholder);
  }

  list.appendChild(slot);
}
function renderMCZone() {
  const list = document.getElementById('mcList');
  if (!list) return;
  list.innerHTML = '';

  // 1スロット固定
  const slot = document.createElement('div');
  slot.className = 'special-slot';
  slot.dataset.zone = 'mc';

  const person = mcMembers[0]; // 1人だけ
  if (person) {
    slot.appendChild(createPersonCard(person, 'mc', ''));
  } else {
    const placeholder = document.createElement('div');
    placeholder.className = 'special-placeholder';
    placeholder.textContent = '空き枠';
    slot.appendChild(placeholder);
  }

  list.appendChild(slot);
}
// ========== タッチ / クリック操作 ==========
function handleSeatClick(e, toTable, toSeat, dropZone) {
  if (e.target.closest('.person-card')) return;
  if (selectedCard) {
    const personId = selectedCard.dataset.personId;
    const fromTable = selectedCard.dataset.fromTable;
    const fromSeat = selectedCard.dataset.fromSeat;
    const check = validateDrop(personId, toTable, toSeat, null, fromTable, fromSeat);
    if (!check.valid) {
      alert(check.message);
      clearSelection();
      return;
    }
    movePerson(personId, fromTable, fromSeat, toTable, toSeat, null, dropZone);
    clearSelection();
    renderAll();
    initDragDrop();
  }
}

function handleCardClick(e) {
  e.stopPropagation();
  const card = e.currentTarget;
  if (selectedCard === card) {
    clearSelection();
    return;
  }
  if (selectedCard) {
    const personId = selectedCard.dataset.personId;
    const fromTable = selectedCard.dataset.fromTable;
    const fromSeat = selectedCard.dataset.fromSeat;
    const toTable = card.dataset.fromTable;
    const toSeat = card.dataset.fromSeat;

    if (fromTable === 'waiting' && toTable === 'waiting') {
      clearSelection();
      return;
    }

    const person = participants.find(p => p.id === personId);
    if (toSeat === 'master' && (!person.role || person.role.trim() === '')) {
      alert('テーブルマスター席には役割のある人のみ配置できます');
      clearSelection();
      return;
    }

    if (fromTable !== 'waiting' && toTable !== 'waiting') {
      swapPersons(personId, fromTable, fromSeat, card.dataset.personId, toTable, toSeat);
    } else if (fromTable === 'waiting' && toTable !== 'waiting') {
      const check = validateDrop(personId, toTable, toSeat, null, fromTable, fromSeat);
      if (!check.valid) {
        alert(check.message);
        clearSelection();
        return;
      }
      movePerson(personId, fromTable, fromSeat, toTable, toSeat, null, null);
    } else if (fromTable !== 'waiting' && toTable === 'waiting') {
      movePerson(personId, fromTable, fromSeat, null, null, 'waiting', document.getElementById('waitingList'));
    }

    clearSelection();
    renderAll();
    initDragDrop();
  } else {
    selectedCard = card;
    card.classList.add('selected');
  }
}

function clearSelection() {
  if (selectedCard) {
    selectedCard.classList.remove('selected');
    selectedCard = null;
  }
}

function initDragDrop() {
  // 待機ゾーン + PA枠 + 司会枠のスロット全てにイベント設定
  document.querySelectorAll('.seat, .waiting-slot, .special-slot').forEach(zone => {
    zone.addEventListener('dragover', handleDragOver);
    zone.addEventListener('drop', handleDrop);
    zone.addEventListener('dragleave', handleDragLeave);
  });

  // 以下は既存のまま
  document.querySelectorAll('.person-card').forEach(card => {
    card.addEventListener('dragstart', handleDragStart);
    card.addEventListener('dragend', handleDragEnd);
    card.addEventListener('click', handleCardClick);
    card.addEventListener('dragover', handleDragOver);
    card.addEventListener('drop', handleCardDrop);
    card.addEventListener('dragleave', handleDragLeave);
  });
}

function handleDragStart(e) {
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('personId', e.target.dataset.personId);
  e.dataTransfer.setData('fromTable', e.target.dataset.fromTable || '');
  e.dataTransfer.setData('fromSeat', e.target.dataset.fromSeat || '');
  e.target.classList.add('dragging');
}

function handleDragEnd(e) {
  e.target.classList.remove('dragging');
}

function handleDragOver(e) {
  e.preventDefault();
  e.currentTarget.classList.add('drag-over');
}

function handleDragLeave(e) {
  e.currentTarget.classList.remove('drag-over');
}

function handleCardDrop(e) {
  e.preventDefault();
  e.stopPropagation();
  e.stopImmediatePropagation();
  e.currentTarget.classList.remove('drag-over');

  const draggedPersonId = e.dataTransfer.getData('personId');
  const fromTable = e.dataTransfer.getData('fromTable') || null;
  const fromSeat = e.dataTransfer.getData('fromSeat') || null;

  const targetCard = e.currentTarget;
  const targetPersonId = targetCard.dataset.personId;
  const toTable = targetCard.dataset.fromTable || null;
  const toSeat = targetCard.dataset.fromSeat || null;

  if (!draggedPersonId || !targetPersonId || draggedPersonId === targetPersonId) return;

  const draggedIsWaiting = fromTable === 'waiting';
  const targetIsWaiting = toTable === 'waiting';

  if (draggedIsWaiting && targetIsWaiting) return;

  if (!draggedIsWaiting && targetIsWaiting) {
    movePerson(draggedPersonId, fromTable, fromSeat, null, null, 'waiting', document.getElementById('waitingList'));
    renderAll();
    initDragDrop();
    return;
  }

  if (draggedIsWaiting && !targetIsWaiting) {
    const check = validateDrop(draggedPersonId, toTable, toSeat, null, fromTable, fromSeat);
    if (!check.valid) {
      alert(check.message);
      return;
    }
    movePerson(targetPersonId, toTable, toSeat, null, null, 'waiting', document.getElementById('waitingList'));
    movePerson(draggedPersonId, fromTable, fromSeat, toTable, toSeat, null, null);
    renderAll();
    initDragDrop();
    return;
  }

  const person = participants.find(p => p.id === draggedPersonId);
  if (toSeat === 'master' && (!person.role || person.role.trim() === '')) {
    alert('テーブルマスター席には役割のある人のみ配置できます');
    return;
  }

  swapPersons(draggedPersonId, fromTable, fromSeat, targetPersonId, toTable, toSeat);
  renderAll();
  initDragDrop();
}

function handleDrop(e) {
  e.preventDefault();
  e.stopPropagation();
  const dropZone = e.currentTarget;
  dropZone.classList.remove('drag-over');

  const personId = e.dataTransfer.getData('personId');
  const fromTable = e.dataTransfer.getData('fromTable') || null;
  const fromSeat = e.dataTransfer.getData('fromSeat') || null;
  const toTable = dropZone.dataset.table || null;
  const toSeat = dropZone.dataset.seat || null;
  const toZone = dropZone.dataset.zone || null;

  const hasCard = dropZone.querySelector('.person-card');
  if (hasCard) {
    return;
  }

  // PA枠・司会枠への移動制限（役割が入っていればOK）
  if (toZone === 'pa' || toZone === 'mc') {
    const person = participants.find(p => p.id === personId);
    if (!person) return;
    
    if (!person.role || person.role.trim() === '') {
      alert('PA枠・司会枠には役割のある人のみ配置できます');
      return;
    }
  }

  const check = validateDrop(personId, toTable, toSeat, toZone, fromTable, fromSeat);
  if (!check.valid) {
    alert(check.message);
    return;
  }

  movePerson(personId, fromTable, fromSeat, toTable, toSeat, toZone, dropZone);
  renderAll();
  initDragDrop();
}


// ===============================
// 印刷プレビュー機能
// ===============================

// モーダル開閉
document.getElementById('print-preview-btn').addEventListener('click', async () => {
  generatePrintPreview();
  document.getElementById('printModal').style.display = 'flex';
});

document.getElementById('printModalClose').addEventListener('click', () => {
  document.getElementById('printModal').style.display = 'none';
});

// 用紙サイズ切り替え
document.querySelectorAll('.btn-paper-size').forEach(btn => {
  btn.addEventListener('click', (e) => {
    document.querySelectorAll('.btn-paper-size').forEach(b => b.classList.remove('active'));
    e.target.classList.add('active');
    currentPaperSize = e.target.dataset.size;
    generatePrintPreview();
  });
});

// 印刷実行（PDF生成方式）
document.getElementById('executePrintBtn').addEventListener('click', async () => {
  const button = document.getElementById('executePrintBtn');
  button.disabled = true;
  button.textContent = '生成中...';
  
  try {
    const element = document.querySelector('.print-page');
    const { jsPDF } = window.jspdf;
    
    // A4もA3も縦向き
    const format = currentPaperSize === 'a4' ? 'a4' : 'a3';
    
    const pdf = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: format,
      compress: true
    });
    
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    
    // 余白設定（A3のみ15mm、A4は5mm）
    const margin = currentPaperSize === 'a3' ? 15 : 5;
    const contentWidth = pageWidth - (margin * 2);
    const contentHeight = pageHeight - (margin * 2);
    
    // コンテンツ領域に合わせてキャプチャ（scale調整）
    const canvas = await html2canvas(element, {
      scale: 3,  // 高解像度化
      useCORS: true,
      logging: false,
      width: element.scrollWidth,
      height: element.scrollHeight
    });
    
    const imgData = canvas.toDataURL('image/png');
    
    // アスペクト比を維持しつつ、コンテンツ領域に収める
    const imgAspect = canvas.width / canvas.height;
    const contentAspect = contentWidth / contentHeight;
    
    let imgWidth, imgHeight;
    if (imgAspect > contentAspect) {
      // 横長：幅に合わせる
      imgWidth = contentWidth;
      imgHeight = contentWidth / imgAspect;
    } else {
      // 縦長：高さに合わせる
      imgHeight = contentHeight;
      imgWidth = contentHeight * imgAspect;
    }
    
    // 中央配置
    const x = (pageWidth - imgWidth) / 2;
    const y = (pageHeight - imgHeight) / 2;
    
    pdf.addImage(imgData, 'PNG', x, y, imgWidth, imgHeight);
    
    // PDFを新しいタブで開く
    const pdfBlob = pdf.output('blob');
    const pdfUrl = URL.createObjectURL(pdfBlob);
    window.open(pdfUrl, '_blank');
    
  } catch (err) {
    console.error(err);
    alert('PDF生成に失敗しました: ' + err.message);
  } finally {
    button.disabled = false;
    button.textContent = '印刷する';
  }
});

// ===============================
// 印刷プレビュー生成
// ===============================
function generatePrintPreview() {
  const previewArea = document.getElementById('printPreviewArea');
  
  const mcPerson = document.querySelector('#mcList .person-card');
  const paPerson = document.querySelector('#paList .person-card');
  
  const mcName = mcPerson ? mcPerson.querySelector('.person-name')?.textContent || '' : '';
  const paName = paPerson ? paPerson.querySelector('.person-name')?.textContent || '' : '';
  
  let tableCardSizeClass = 'large';
  if (currentPaperSize === 'a4') {
    tableCardSizeClass = 'medium';
  }
  
  let html = `
    <div class="print-page ${currentPaperSize}">
      <div class="print-stage">ステージ</div>
      
      <div class="print-special-box">
        <div class="print-special-label">🎙️ 司会</div>
        <div class="print-special-card">${mcName}</div>
      </div>
      
      <div class="print-pa-box">
        <div class="print-special-label">🎤 PA</div>
        <div class="print-special-card">${paName}</div>
      </div>
      
      <div class="print-tables-grid">
  `;
  
  const tableRows = document.querySelectorAll('.table-row');
  
  tableRows.forEach(row => {
    html += '<div class="print-table-row">';
    
    const tables = row.querySelectorAll('.table-card');
    tables.forEach(table => {
      const tableTitle = table.querySelector('.table-title')?.textContent || '';
      const seats = table.querySelectorAll('.seat');
      
      html += `
        <div class="print-table-card ${tableCardSizeClass}">
          <div class="print-table-header">${tableTitle}</div>
          <div class="print-seats-grid">
      `;
      
      seats.forEach((seat, idx) => {
        const personCard = seat.querySelector('.person-card');
        
        if (personCard) {
          const nameEl = personCard.querySelector('.person-name');
          const name = nameEl ? nameEl.textContent : '';
          
          let seatClass = '';
          let label = '';
          
          if (personCard.classList.contains('jikai')) {
            seatClass = 'jikai';
          } else if (personCard.classList.contains('guest')) {
            seatClass = 'guest';
            label = 'ゲスト';
          } else if (personCard.classList.contains('other-venue')) {
            seatClass = 'other-venue';
            const titleEl = personCard.querySelector('.person-title');
            if (titleEl) {
              label = titleEl.textContent.replace('会場', '') + '会場';
            }
          }
          
          if (idx === 0) {
            html += `
              <div class="print-seat master">
                <div class="print-seat-label">👑 テーブルマスター</div>
                <div>${name}</div>
              </div>
            `;
          } else if (label) {
            html += `
              <div class="print-seat ${seatClass}">
                <div class="print-seat-label">${label}</div>
                <div>${name}</div>
              </div>
            `;
          } else {
            html += `<div class="print-seat ${seatClass}">${name}</div>`;
          }
        }
      });
      
      html += `
          </div>
        </div>
      `;
    });
    
    html += '</div>';
  });
  
  html += `
      </div>
    </div>
  `;
  
  previewArea.innerHTML = html;
}