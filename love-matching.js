/**
 * Google Apps Script - 月老配對活動自動化系統
 * 功能：自動配對男女志願，並輸出配對結果到 Google 試算表
 * 作者：林俞翔
 * 備註：需在 Google Apps Script 環境下運行
 */

function onFormSubmit(e) {
  runMatching(); // 觸發主配對函數
}

function runMatching() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0]; // 表單回應工作表
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  const boys = {};
  const girls = {};
  const allPlayers = {}; // 記錄配對狀態

  // 建立資料結構
  rows.forEach(row => {
    const gender = row[1];
    const id = row[2];
    const prefsRaw = [row[3], row[4], row[5]].filter(v => v !== "" && v !== null && v !== undefined);

    const key = (gender === "男" ? "B" : "G") + id;
    allPlayers[key] = { matched: null };

    if (gender === "男") boys[key] = prefsRaw.map(x => "G" + x);
    if (gender === "女") girls[key] = prefsRaw.map(x => "B" + x);
  });

  // 計算人氣
  const popularity = {};
  for (let b in boys) boys[b].forEach(g => { if(allPlayers[g]) popularity[g] = (popularity[g] || 0) + 1; });
  for (let g in girls) girls[g].forEach(b => { if(allPlayers[b]) popularity[b] = (popularity[b] || 0) + 1; });

  // 人氣高的人先分配
  const playersByPopularity = Object.keys(popularity).sort((a, b) => popularity[b] - popularity[a]);

  playersByPopularity.forEach(target => {
    if (!allPlayers[target] || allPlayers[target].matched) return;

    const prefs = target.startsWith("B") ? boys[target] : girls[target];
    if (!prefs) return;

    // 找喜歡 target 且未配對的人
    const candidates = [];
    for (let other in allPlayers) {
      if (allPlayers[other].matched) continue;
      const otherPrefs = other.startsWith("B") ? boys[other] : girls[other];
      if (!otherPrefs) continue;
      if (otherPrefs.includes(target)) candidates.push(other);
    }

    // 從被選者的志願決定配對
    for (let i = 0; i < prefs.length; i++) {
      const preferred = prefs[i];
      if (!allPlayers[preferred]) continue;
      if (candidates.includes(preferred)) {
        allPlayers[target].matched = preferred;
        allPlayers[preferred].matched = target;
        const rankInfo = `(${target} 的第${i+1}志願, ${preferred} 的第${(preferred.startsWith("B") ? boys[preferred].indexOf(target) : girls[preferred].indexOf(target))+1}志願)`;
        allPlayers[target].rankInfo = rankInfo;
        allPlayers[preferred].rankInfo = rankInfo;
        break;
      }
    }
  });

  // 輸出結果
  let outSheet = ss.getSheetByName("配對結果");
  if (!outSheet) outSheet = ss.insertSheet("配對結果");
  outSheet.clear();
  outSheet.appendRow(["自己", "配對到的人", "順位"]);

  // 先列男生 B1~B60，再列女生 G1~G60
  const sortedKeys = Object.keys(allPlayers).sort((a, b) => {
    if (a.startsWith("B") && b.startsWith("G")) return -1;
    if (a.startsWith("G") && b.startsWith("B")) return 1;
    return parseInt(a.slice(1)) - parseInt(b.slice(1));
  });

  sortedKeys.forEach(p => {
    if (allPlayers[p].matched) {
      outSheet.appendRow([p, allPlayers[p].matched, allPlayers[p].rankInfo]);
    } else {
      outSheet.appendRow([p, "未配對", ""]);
    }
  });
// ===== 自動上色 =====
const lastRow = outSheet.getLastRow();
const range = outSheet.getRange(2, 1, lastRow - 1, 3); // 從第2列到最後列，3欄
const bgColors = [];

for (let i = 2; i <= lastRow; i++) {
  const rank = outSheet.getRange(i, 3).getValue();
  if (!rank || rank === "") {
    bgColors.push(["#D3D3D3", "#D3D3D3", "#D3D3D3"]); // 灰色 → 未配對
  } else if (rank.includes("第1志願") && rank.match(/第1志願/g).length === 2) {
    bgColors.push(["#C6EFCE", "#C6EFCE", "#C6EFCE"]); // 綠色 → 雙方都是第1志願
  } else {
    bgColors.push(["#FFF2CC", "#FFF2CC", "#FFF2CC"]); // 黃色 → 第2或第3志願
  }
}

range.setBackgrounds(bgColors);


}