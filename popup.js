document.addEventListener('DOMContentLoaded', () => {
  // 绑定事件
  document.getElementById('startRow').addEventListener('input', checkAndPreview);
  document.getElementById('minRows').addEventListener('input', checkAndPreview);
  document.getElementById('exportExcel').addEventListener('click', () => doExport('excel'));
  document.getElementById('exportJson').addEventListener('click', () => doExport('json'));

  // 初始检查当前页面的表格信息
  checkAndPreview();
});

function getFilterParams() {
  const startRow = parseInt(document.getElementById('startRow').value, 10) || 0;
  const minRows = parseInt(document.getElementById('minRows').value, 10) || 1;
  return { startRow: Math.max(0, startRow), minRows: Math.max(1, minRows) };
}

function setStatus(text) {
  document.getElementById('status').innerText = text;
}

// 检查并预览符合条件的表格数量
async function checkAndPreview() {
  const { startRow, minRows } = getFilterParams();
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab || !tab.id) return;

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: parseTablesFromPage,
    args: [startRow, minRows]
  }, (results) => {
    const infoPanel = document.getElementById('infoPanel');
    const exportExcelBtn = document.getElementById('exportExcel');
    const exportJsonBtn = document.getElementById('exportJson');

    if (!results || !results[0] || !results[0].result) {
      infoPanel.innerText = '无法获取当前页面信息';
      exportExcelBtn.disabled = true;
      exportJsonBtn.disabled = true;
      return;
    }

    const { totalTables, qualifiedTables } = results[0].result;

    infoPanel.innerHTML = `
      网页共找到 <b>${totalTables}</b> 张表格<br>
      符合条件 (从第 <b>${startRow}</b> 行起至少 <b>${minRows}</b> 行)：<b>${qualifiedTables.length}</b> 张
    `;

    const hasData = qualifiedTables.length > 0;
    exportExcelBtn.disabled = !hasData;
    exportJsonBtn.disabled = !hasData;
  });
}

// 执行导出逻辑
async function doExport(format) {
  const { startRow, minRows } = getFilterParams();
  setStatus('正在导出数据...');
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: parseTablesFromPage,
    args: [startRow, minRows]
  }, (results) => {
    if (!results || !results[0] || !results[0].result) {
      setStatus('导出失败，无法提取数据');
      return;
    }

    const { qualifiedTables } = results[0].result;
    if (qualifiedTables.length === 0) {
      setStatus('没有可供导出的表格');
      return;
    }

    if (format === 'excel') {
      exportToExcel(qualifiedTables);
    } else {
      exportToJson(qualifiedTables);
    }
    setStatus(`已成功导出 ${qualifiedTables.length} 个表格`);
  });
}

// 页面内部提取和过滤表格的函数 (运行在目标页面上下文)
function parseTablesFromPage(startRow, minRows) {
  const tables = document.querySelectorAll('table');
  const qualifiedTables = [];

  tables.forEach((table, index) => {
    const rows = Array.from(table.rows);

    // 从第 startRow 行开始截取数据
    if (rows.length > startRow) {
      const slicedRows = rows.slice(startRow);

      // 如果截取后的有效行数达到 minRows 要求，则保留该表
      if (slicedRows.length >= minRows) {
        const tableData = slicedRows.map(row => {
          return Array.from(row.cells).map(cell => cell.innerText.trim());
        });

        qualifiedTables.push({
          sheetName: `Table_${index + 1}`,
          data: tableData
        });
      }
    }
  });

  return {
    totalTables: tables.length,
    qualifiedTables: qualifiedTables
  };
}

// 导出为 Excel
function exportToExcel(tablesData) {
  const wb = XLSX.utils.book_new();
  tablesData.forEach(item => {
    const ws = XLSX.utils.aoa_to_sheet(item.data);
    XLSX.utils.book_append_sheet(wb, ws, item.sheetName.substring(0, 31));
  });
  XLSX.writeFile(wb, `filtered_tables_${Date.now()}.xlsx`);
}

// 导出为 JSON
function exportToJson(tablesData) {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(tablesData, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  downloadAnchor.setAttribute("download", `filtered_tables_${Date.now()}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
}