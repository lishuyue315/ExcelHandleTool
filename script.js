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

  if (!fileInput.files.length) {
    alert("ファイルを選択してください！");
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    // 結果格納 { MAWB番号: { totalE, osakaE, shigaE } }
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


    resultDiv.classList.remove("hidden");
  };

  reader.readAsArrayBuffer(file);
}

  