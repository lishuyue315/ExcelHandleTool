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
  
      const output = [["県", "町", "出現数", "県内占比", "県内順位"]];
  
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
  
        list.forEach(item => {
          const ratio = (item.count / total * 100).toFixed(2) + "%";
          output.push([ken, item.cho, item.count, ratio, item.rank]);
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
    const input = document.getElementById("eStatsInput");
    const file = input.files[0];
    if (!file) return alert("ファイルを選択してください。");
  
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" }); // header: true by default
  
      let totalE = 0;
      let osakaE = 0;
  
      for (const row of json) {
        const airwayNo = String(row["HOUSE AIR WAYBILL NO."] || "").trim();
        const address = String(row["輸入者住所"] || "").trim();
  
        if (airwayNo.startsWith("E")) {
          totalE++;
          if (address.startsWith("Osaka")) {
            osakaE++;
          }
        }
      }
  
      // 填入结果
      document.getElementById("totalECount").textContent = totalE;
      document.getElementById("osakaECount").textContent = osakaE;
      document.getElementById("diffCount").textContent = totalE - osakaE;
      document.getElementById("eStatsResult").classList.remove("hidden");
    };
  
    reader.readAsArrayBuffer(file);
  }
  