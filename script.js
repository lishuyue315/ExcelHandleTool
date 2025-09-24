const areaMapping = {
  "北海道": "北海道",
  "青森県": "東北",
  "岩手県": "東北",
  "宮城県": "東北",
  "秋田県": "東北",
  "山形県": "東北",
  "福島県": "東北",
  "茨城県": "関東",
  "栃木県": "関東",
  "群馬県": "関東",
  "埼玉県": "関東",
  "千葉県": "関東",
  "東京都": "関東",
  "神奈川県": "関東",
  "新潟県": "中部",
  "富山県": "中部",
  "石川県": "中部",
  "福井県": "中部",
  "山梨県": "中部",
  "長野県": "中部",
  "岐阜県": "中部",
  "静岡県": "中部",
  "愛知県": "中部",
  "三重県": "近畿",
  "滋賀県": "近畿",
  "京都府": "近畿",
  "大阪府": "近畿",
  "兵庫県": "近畿",
  "奈良県": "近畿",
  "和歌山県": "近畿",
  "鳥取県": "中国",
  "島根県": "中国",
  "岡山県": "中国",
  "広島県": "中国",
  "山口県": "中国",
  "徳島県": "四国",
  "香川県": "四国",
  "愛媛県": "四国",
  "高知県": "四国",
  "福岡県": "九州",
  "佐賀県": "九州",
  "長崎県": "九州",
  "熊本県": "九州",
  "大分県": "九州",
  "宮崎県": "九州",
  "鹿児島県": "九州",
  "沖縄県": "沖縄"
};

function processFile() {
  const fileInput = document.getElementById("fileInput");
  const file = fileInput.files[0];
  if (!file) return alert("ファイルを選択してください。");

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const header = json[0];
    const rows = json.slice(1);

    const resultMap = {};

    for (const row of rows) {
      const ken = row[0];
      const cho = row[1];
      if (!ken || !cho) continue;
      const key = `${ken}-${cho}`;
      resultMap[key] = (resultMap[key] || 0) + 1;
    }

    const kenMap = {}; // 県別の町出現数
    for (const key in resultMap) {
      const [ken, cho] = key.split("-");
      if (!kenMap[ken]) kenMap[ken] = [];
      kenMap[ken].push({ cho, count: resultMap[key] });
    }

    // エリア列を追加
    const output = [["エリア", "県", "町", "出現数", "県内占比", "県内順位"]];

    for (const ken in kenMap) {
      const list = kenMap[ken];
      const total = list.reduce((sum, item) => sum + item.count, 0);

      // 排序
      list.sort((a, b) => b.count - a.count);

      let lastCount = null;
      let lastRank = 0;
      let currentIndex = 1;

      list.forEach((item, index) => {
        if (item.count === lastCount) {
          item.rank = lastRank;
        } else {
          item.rank = currentIndex;
          lastRank = currentIndex;
          lastCount = item.count;
        }
        currentIndex++;
      });

      // 県からエリアを取得
      const area = areaMapping[ken] || "不明";
      
      list.forEach(item => {
        const ratio = (item.count / total * 100).toFixed(2) + "%";
        output.push([area, ken, item.cho, item.count, ratio, item.rank]);
      });
    }

    const ws = XLSX.utils.aoa_to_sheet(output);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "結果");
    XLSX.writeFile(wb, "processed.xlsx");
  };

  reader.readAsArrayBuffer(file);
}

function analyzeEStats() {
  const fileInput = document.getElementById("eStatsInput");
  const resultDiv = document.getElementById("eStatsResult");
  const tableBody = document.getElementById("eStatsTable");
  const placeholderImage = document.getElementById("placeholderImage");
  const resetBtn = document.getElementById("resetEBtn");
  const runBtn = document.getElementById("runEBtn");

  if (!fileInput.files.length) {
    alert("ファイルを選択してください！");
    return;
  }

  // 隐藏占位图片并显示结果表格
  placeholderImage.style.maxHeight = "0";
  placeholderImage.style.opacity = "0";
  
  // 显示重置按钮
  resetBtn.classList.remove("hidden");
  runBtn.classList.add("hidden");

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    // 原有的数据处理逻辑保持不变...
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    // 结果格納 { MAWB番号: { totalE, osakaE, shigaE } }
    const stats = {};

    jsonData.forEach(row => {
      const mawb = row["MAWB番号"];
      const hawb = row["HAWB番号"];
      const addr = row["收件人地址"] || "";

      if (!mawb) return;

      if (!stats[mawb]) {
        stats[mawb] = { totalE: 0, osakaE: 0, shigaE: 0 };
      }

      if (hawb && hawb.startsWith("E")) {
        stats[mawb].totalE++;
        if (addr.startsWith("大阪府")) stats[mawb].osakaE++;
        if (addr.startsWith("滋賀県")) stats[mawb].shigaE++;
      }
    });

    // 清空旧结果
    tableBody.innerHTML = "";

    // 渲染每个 MAWB 的结果（竖列交替）
    Object.entries(stats).forEach(([mawb, values]) => {
      const diff = values.totalE - values.osakaE - values.shigaE;
      const row = document.createElement("tr");
      const rowValues = [ mawb, values.totalE, values.osakaE, values.shigaE, diff ];

      rowValues.forEach((val, colIndex) => {
        const cell = document.createElement("td");
        cell.className = `px-4 py-2 border text-center ${colIndex % 2 === 0 ? "bg-gray-50" : "bg-white"}`;

        if (colIndex === 0) {
          cell.className += " font-semibold text-green-700 max-w-[200px] truncate";
          cell.title = val;
        }

        cell.textContent = val;
        row.appendChild(cell);
      });

      tableBody.appendChild(row);
    });

    // 统计总和
    let totalE = 0, totalOsaka = 0, totalShiga = 0;
    Object.values(stats).forEach(v => {
      totalE += v.totalE;
      totalOsaka += v.osakaE;
      totalShiga += v.shigaE;
    });
    const totalDiff = totalE - totalOsaka - totalShiga;

    // 添加合計行（放最下面）
    const totalRow = document.createElement("tr");
    const totalValues = [ "合計", totalE, totalOsaka, totalShiga, totalDiff ];
    totalValues.forEach((val, colIndex) => {
      const cell = document.createElement("td");
      cell.className = `text-center px-4 py-2 border font-bold text-gray-800 ${colIndex % 2 === 0 ? "bg-gray-100" : "bg-gray-50"}`;
      cell.textContent = val;
      totalRow.appendChild(cell);
    });
    tableBody.appendChild(totalRow);

    // 显示结果表格（带动画）
    setTimeout(() => {
      resultDiv.style.maxHeight = "1000px";
      resultDiv.style.opacity = "1";
    }, 300);
  };

  reader.readAsArrayBuffer(file);
}

// 添加重置函数
function resetEView() {
  const resultDiv = document.getElementById("eStatsResult");
  const placeholderImage = document.getElementById("placeholderImage");
  const resetBtn = document.getElementById("resetEBtn");
  const runBtn = document.getElementById("runEBtn");

  // 隐藏结果表格
  resultDiv.style.maxHeight = "0";
  resultDiv.style.opacity = "0";
  
  // 显示重置按钮
  resetBtn.classList.add("hidden");
  runBtn.classList.remove("hidden");

  // 显示占位图片
  setTimeout(() => {
    placeholderImage.style.maxHeight = "18rem";
    placeholderImage.style.opacity = "1";
  }, 300);
}

  
async function processMapping() {
  const mappingFileInput = document.getElementById("mappingFile");
  const targetFileInput = document.getElementById("targetFile");

  if (!mappingFileInput.files[0] || !targetFileInput.files[0]) {
    alert("請上傳文件1和文件2！");
    return;
  }

  // 1) 讀取 mapping 文件（文件1）
  const mappingBuf = await mappingFileInput.files[0].arrayBuffer();
  const mappingWb = XLSX.read(mappingBuf);
  const mappingSheet = mappingWb.Sheets[mappingWb.SheetNames[0]];
  const mappingAoa = XLSX.utils.sheet_to_json(mappingSheet, { header: 1, defval: "" });

  // 建立映射 Map：把 B 列 (index 1) -> A 列 (index 0)
  const mappingMap = new Map();
  mappingAoa.forEach((row, i) => {
    const key = row[1] !== undefined && row[1] !== null ? String(row[1]).trim() : "";
    const val = row[0] !== undefined && row[0] !== null ? row[0] : "";
    if (key !== "") mappingMap.set(key, val);
  });
  console.log("mappingMap size:", mappingMap.size);

  // 2) 讀取 target 文件（文件2） - 直接操作 worksheet
  const targetBuf = await targetFileInput.files[0].arrayBuffer();
  const targetWb = XLSX.read(targetBuf);
  const sheetName = targetWb.SheetNames[0];
  const ws = targetWb.Sheets[sheetName];

  // 确保有 !ref
  if (!ws || !ws['!ref']) {
    alert("目标工作表缺少数据或无法读取！");
    return;
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  console.log("原始范围:", range);

  const startRow = range.s.r; 
  const endRow = range.e.r;

  // 逐行读取 G 列（c = 6），写到 A 列（c = 0）
  for (let R = startRow; R <= endRow; R++) {
    const gAddr = XLSX.utils.encode_cell({ r: R, c: 6 }); // G列 index 6
    const gCell = ws[gAddr];
    const key = gCell && gCell.v !== undefined && gCell.v !== null ? String(gCell.v).trim() : "";

    if (!key) continue; // 没 key 就跳过

    if (mappingMap.has(key)) {
      const mappedVal = mappingMap.get(key);

      const aAddr = XLSX.utils.encode_cell({ r: R, c: 0 }); // A列 index 0
      // 判断类型：若是数字则写数字，否则写字符串
      if (mappedVal !== null && mappedVal !== undefined && mappedVal !== "" && !isNaN(mappedVal) && typeof mappedVal !== "string") {
        ws[aAddr] = { t: "n", v: Number(mappedVal) };
      } else if (mappedVal !== null && mappedVal !== undefined && !isNaN(String(mappedVal)) && String(mappedVal).trim() !== "") {
        // 当映射值是字符串但可以转成数字，也写成数字
        const maybeNum = Number(String(mappedVal).trim());
        if (!Number.isNaN(maybeNum)) {
          ws[aAddr] = { t: "n", v: maybeNum };
        } else {
          ws[aAddr] = { t: "s", v: String(mappedVal) };
        }
      } else {
        ws[aAddr] = { t: "s", v: String(mappedVal === undefined ? "" : mappedVal) };
      }
    }
  }

  // 3) 更新 !ref，确保 A 列被包含（如果原来 !ref 的起始列 > 0）
  const newMinC = Math.min(range.s.c, 0);
  const newRange = { s: { r: range.s.r, c: newMinC }, e: { r: range.e.r, c: range.e.c } };
  ws['!ref'] = XLSX.utils.encode_range(newRange);
  console.log("更新後範圍:", ws['!ref']);

  // 4) 导出：直接把修改过的 worksheet 写回到新的 workbook 并下载
  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, ws, sheetName);
  XLSX.writeFile(outWb, "mapped_result.xlsx");
}

async function buildCodesFromSheet2() {
  const fileInput = document.getElementById("dedupFile");
  const prefixInput = document.getElementById("codePrefix");
  const resultDiv = document.getElementById("dedupResult");
  const tableBody = document.getElementById("dedupTable");
  const placeholder = document.getElementById("dedupPlaceholder");
  const resetBtn = document.getElementById("resetDedupBtn");
  const runBtn = document.getElementById("runDedupBtn");
  const meta = document.getElementById("dedupMeta");

  if (!fileInput.files.length) {
    alert("ファイルを選択してください！");
    return;
  }

  const prefixRaw = (prefixInput.value || "").trim();
  if (!/^\d+$/.test(prefixRaw)) {
    alert("前缀请输入纯数字（例：23）");
    return;
  }

  // UI 切换
  placeholder.classList.add("hidden");
  resetBtn.classList.remove("hidden");
  runBtn.classList.add("hidden");

  // 读取 Excel
  const buf = await fileInput.files[0].arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  // 选择工作表：
  // 1) 如果只有一个表，直接用它；
  // 2) 如果有多个表，挑选 H 列（索引7）非空值最多的那个
  const H_INDEX = 7;
  let chosenSheet = null;
  let maxHCount = -1;

  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) return;
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    let count = 0;
    for (let r = 0; r < aoa.length; r++) {
      const v = aoa[r] && aoa[r][H_INDEX] != null ? String(aoa[r][H_INDEX]).trim() : "";
      if (v !== "") count++;
    }
    if (count > maxHCount) {
      maxHCount = count;
      chosenSheet = ws;
    }
  });

  // 兜底：实在没选到，就用第一个
  if (!chosenSheet) chosenSheet = wb.Sheets[wb.SheetNames[0]];
  if (!chosenSheet || !chosenSheet['!ref']) {
    alert("无法读取工作表数据，请检查文件内容。");
    resetBtn.classList.add("hidden");
    runBtn.classList.remove("hidden");
    placeholder.classList.remove("hidden");
    return;
  }

  // 读取 H 列数据
  const aoa = XLSX.utils.sheet_to_json(chosenSheet, { header: 1, defval: "" });
  const values = [];
  for (let r = 0; r < aoa.length; r++) {
    const cellVal = aoa[r] && aoa[r][H_INDEX] != null ? String(aoa[r][H_INDEX]).trim() : "";
    if (cellVal !== "") values.push(cellVal);
  }

  // 去重（保持首次出现顺序）
  const seen = new Set();
  const uniqueList = [];
  for (const v of values) {
    if (!seen.has(v)) {
      seen.add(v);
      uniqueList.push(v);
    }
  }

  // 序号位宽：至少2位
  const width = Math.max(2, String(uniqueList.length).length);

  // 渲染表格
  tableBody.innerHTML = "";
  uniqueList.forEach((val, idx) => {
    const serial = String(idx + 1).padStart(width, "0");
    const code = `${prefixRaw}-${serial}`;

    const tr = document.createElement("tr");
    const cells = [code, val];

    cells.forEach((c, i) => {
      const td = document.createElement("td");
      td.className = `px-4 py-2 border text-center ${i % 2 === 0 ? "bg-gray-50" : "bg-white"}`;
      if (i === 1) td.className += " text-left";
      td.textContent = c;
      tr.appendChild(td);
    });

    tableBody.appendChild(tr);
  });

  meta.textContent = `去重总数：${uniqueList.length}　|　编号位数：${width}　|　前缀：${prefixRaw}`;

  // 显示结果（固定高度 + 可滚动），柔和淡入
  resultDiv.classList.remove("hidden");
  requestAnimationFrame(() => {
    resultDiv.style.opacity = "1";
  });
}
