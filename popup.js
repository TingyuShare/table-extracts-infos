document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('startRow').addEventListener('input', checkAndPreview);
  document.getElementById('minRows').addEventListener('input', checkAndPreview);
  document.getElementById('enablePagination').addEventListener('change', (e) => {
    document.getElementById('paginationSettings').style.display = e.target.checked ? 'block' : 'none';
  });

  document.getElementById('exportExcel').addEventListener('click', () => startProcess('excel'));
  document.getElementById('exportJson').addEventListener('click', () => startProcess('json'));

  checkAndPreview();
});

function getFilterParams() {
  const startRow = parseInt(document.getElementById('startRow').value, 10) || 0;
  const minRows = parseInt(document.getElementById('minRows').value, 10) || 1;
  const enablePagination = document.getElementById('enablePagination').checked;
  const nextBtnSelector = document.getElementById('nextBtnSelector').value.trim();
  const maxPages = parseInt(document.getElementById('maxPages').value, 10) || 10;
  const pageDelay = parseInt(document.getElementById('pageDelay').value, 10) || 2;

  return {
    startRow: Math.max(0, startRow),
    minRows: Math.max(1, minRows),
    enablePagination,
    nextBtnSelector,
    maxPages,
    pageDelay
  };
}

function setStatus(text) {
  document.getElementById('status').innerText = text;
}

// 预览当前页情况
async function checkAndPreview() {
  const params = getFilterParams();
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab || !tab.id) return;

  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: parseSinglePageTables,
    args: [params.startRow, params.minRows]
  }, (results) => {
    const infoPanel = document.getElementById('infoPanel');
    if (!results || !results[0] || !results[0].result) {
      infoPanel.innerText = '无法获取当前页面信息';
      return;
    }

    const { totalTables, qualifiedTables } = results[0].result;
    infoPanel.innerHTML = `
      当前页共 <b>${totalTables}</b> 张表格<br>
      符合条件：<b>${qualifiedTables.length}</b> 张表格
    `;
  });
}

// 启动导出主逻辑
async function startProcess(format) {
  const params = getFilterParams();
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!params.enablePagination) {
    // 仅抓取单页
    setStatus('正在抓取当前页...');
    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: parseSinglePageTables,
      args: [params.startRow, params.minRows]
    }, (results) => {
      const qualifiedTables = results[0].result.qualifiedTables;
      outputData(qualifiedTables, format);
    });
  } else {
    // 自动多页抓取逻辑
    if (!params.nextBtnSelector) {
      setStatus('请输入有效的“下一页”选择器');
      return;
    }
    setStatus('正在启动多页抓取...');

    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: autoCrawlPagination,
      args: [params]
    }, (results) => {
      if (!results || !results[0] || !results[0].result) {
        setStatus('翻页抓取失败');
        return;
      }
      const aggregatedTables = results[0].result;
      outputData(aggregatedTables, format);
    });
  }
}

// 多页自动点击抓取函数 (在目标页面运行)
async function autoCrawlPagination(params) {
  const { startRow, minRows, nextBtnSelector, maxPages, pageDelay } = params;
  let currentPage = 1;
  const mergedTablesMap = {}; // 用来按表格序号（Table_1, Table_2...）合并多页数据

  const delay = (ms) => new Promise(res => setTimeout(res, ms));

  // 辅助寻找“下一页”按钮
  function findNextButton(selector) {
    let el = document.querySelector(selector);
    if (el) return el;

    // 备用兜底策略：如果 selector 没找到，按文本匹配包含“下一页”或“>”的按钮/链接
    const elements = Array.from(document.querySelectorAll('a, button, li, span'));
    return elements.find(e => {
      const text = e.innerText.trim();
      return text === '下一页' || text === '下页' || text === '>' || text === 'Next »';
    });
  }

  while (currentPage <= maxPages) {
    console.log(`正在抓取第 ${currentPage} 页...`);

    // 1. 提取当前页表格数据
    const tables = document.querySelectorAll('table');
    tables.forEach((table, index) => {
      const rows = Array.from(table.rows);
      if (rows.length > startRow) {
        const slicedRows = rows.slice(startRow);
        if (slicedRows.length >= minRows) {
          const tableData = slicedRows.map(row => 
            Array.from(row.cells).map(cell => cell.innerText.trim())
          );

          const key = `Table_${index + 1}`;
          if (!mergedTablesMap[key]) {
            mergedTablesMap[key] = { sheetName: key, data: [] };
          }

          // 如果不是第一页，过滤重复的表头（可选逻辑：跳过翻页后的第1行表头）
          let dataToAdd = tableData;
          if (currentPage > 1 && mergedTablesMap[key].data.length > 0) {
            // 比对当前页第一行与上一页表头是否一致，若一致则不重复添加表头
            const prevHeader = JSON.stringify(mergedTablesMap[key].data[0]);
            const currHeader = JSON.stringify(tableData[0]);
            if (prevHeader === currHeader) {
              dataToAdd = tableData.slice(1);
            }
          }

          mergedTablesMap[key].data.push(...dataToAdd);
        }
      }
    });

    // 2. 尝试寻找并点击“下一页”
    const nextBtn = findNextButton(nextBtnSelector);

    // 没有下一页按钮，或按钮处于不可用(disabled)状态时结束
    if (!nextBtn || nextBtn.classList.contains('disabled') || nextBtn.hasAttribute('disabled')) {
      console.log('未检测到可用的下一页按钮，翻页结束。');
      break;
    }

    if (currentPage >= maxPages) break;

    // 3. 点击下一页并等待加载
    nextBtn.click();
    currentPage++;
    await delay(pageDelay * 1000);
  }

  return Object.values(mergedTablesMap);
}

// 提取单页数据的原逻辑 (在目标页面运行)
function parseSinglePageTables(startRow, minRows) {
  const tables = document.querySelectorAll('table');
  const qualifiedTables = [];

  tables.forEach((table, index) => {
    const rows = Array.from(table.rows);
    if (rows.length > startRow) {
      const slicedRows = rows.slice(startRow);
      if (slicedRows.length >= minRows) {
        const tableData = slicedRows.map(row => 
          Array.from(row.cells).map(cell => cell.innerText.trim())
        );

        qualifiedTables.push({
          sheetName: `Table_${index + 1}`,
          data: tableData
        });
      }
    }
  });

  return { totalTables: tables.length, qualifiedTables };
}

// 文件导出处理
function outputData(tablesData, format) {
  if (!tablesData || tablesData.length === 0) {
    setStatus('未抓取到符合条件的表格数据');
    return;
  }

  if (format === 'excel') {
    const wb = XLSX.utils.book_new();
    tablesData.forEach(item => {
      const ws = XLSX.utils.aoa_to_sheet(item.data);
      XLSX.utils.book_append_sheet(wb, ws, item.sheetName.substring(0, 31));
    });
    XLSX.writeFile(wb, `paginated_tables_${Date.now()}.xlsx`);
  } else {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(tablesData, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", `paginated_tables_${Date.now()}.json`);
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
  }

  setStatus(`完成！合并抓取了 ${tablesData.length} 张表格。`);
}