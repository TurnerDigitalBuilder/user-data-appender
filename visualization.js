// Data table and presentation utilities
const DataTable = (function() {
  'use strict';

  let columns = [];
  let rows = [];
  let tableElement = null;
  let tableHead = null;
  let tableBody = null;
  let tableWrapper = null;
  let summaryElement = null;
  let emptyStateElement = null;

  function normalizeRow(row) {
    if (!Array.isArray(row)) {
      return Array.from({ length: columns.length }, () => '');
    }
    return columns.map((_, idx) => formatCell(row[idx]));
  }

  function formatCell(value) {
    if (value === undefined || value === null) return '';
    return value.toString();
  }

  function updateSummary() {
    if (!summaryElement) return;
    if (!hasData()) {
      summaryElement.textContent = 'Upload a CSV or Excel file to begin.';
      return;
    }
    const rowLabel = rows.length === 1 ? 'row' : 'rows';
    const colLabel = columns.length === 1 ? 'column' : 'columns';
    summaryElement.textContent = `${rows.length} ${rowLabel} • ${columns.length} ${colLabel}`;
  }

  function updateEmptyState() {
    if (!tableWrapper || !emptyStateElement) return;
    if (!hasData()) {
      tableWrapper.classList.add('empty');
      emptyStateElement.style.display = 'flex';
    } else {
      tableWrapper.classList.remove('empty');
      emptyStateElement.style.display = 'none';
    }
  }

  function render() {
    if (!tableHead || !tableBody) return;

    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    if (!hasData()) {
      updateSummary();
      updateEmptyState();
      return;
    }

    const headerRow = document.createElement('tr');
    columns.forEach(column => {
      const th = document.createElement('th');
      th.textContent = column;
      headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);

    rows.forEach(row => {
      const tr = document.createElement('tr');
      columns.forEach((_, idx) => {
        const td = document.createElement('td');
        td.textContent = formatCell(row[idx]);
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });

    updateSummary();
    updateEmptyState();
  }

  function hasData() {
    return columns.length > 0 && rows.length > 0;
  }

  function normalizeColumnName(name) {
    return (name || '')
      .toString()
      .replace(/[^a-z0-9]/gi, '')
      .toLowerCase();
  }

  return {
    initialize: function() {
      tableElement = document.getElementById('dataTable');
      tableHead = tableElement ? tableElement.querySelector('thead') : null;
      tableBody = tableElement ? tableElement.querySelector('tbody') : null;
      tableWrapper = document.getElementById('tableWrapper');
      summaryElement = document.getElementById('tableSummary');
      emptyStateElement = document.getElementById('tableEmptyState');
      this.clear();
    },

    clear: function() {
      columns = [];
      rows = [];
      if (tableHead) tableHead.innerHTML = '';
      if (tableBody) tableBody.innerHTML = '';
      updateSummary();
      updateEmptyState();
    },

    loadData: function(newColumns, newRows) {
      columns = Array.isArray(newColumns)
        ? newColumns.map((col, idx) => {
            const label = (col === undefined || col === null || col === '')
              ? `Column ${idx + 1}`
              : col.toString();
            return label;
          })
        : [];
      rows = Array.isArray(newRows) ? newRows.map(normalizeRow) : [];
      render();
    },

    hasData,

    getRowCount: function() {
      return rows.length;
    },

    getColumnCount: function() {
      return columns.length;
    },

    getColumns: function() {
      return columns.slice();
    },

    getRows: function() {
      return rows.map(row => row.slice());
    },

    getCellValue: function(rowIndex, columnIndex) {
      if (!rows[rowIndex] || columnIndex < 0 || columnIndex >= columns.length) {
        return '';
      }
      return formatCell(rows[rowIndex][columnIndex]);
    },

    ensureColumn: function(columnName, preferredIndex) {
      const normalized = columnName.toLowerCase();
      let existingIndex = columns.findIndex(col => col.toLowerCase() === normalized);
      if (existingIndex !== -1) {
        return existingIndex;
      }

      const targetIndex = Math.max(0, Math.min(
        typeof preferredIndex === 'number' ? preferredIndex : columns.length,
        columns.length
      ));

      columns.splice(targetIndex, 0, columnName);
      rows.forEach(row => {
        row.splice(targetIndex, 0, '');
      });
      return targetIndex;
    },

    setColumnValues: function(columnIndex, values) {
      rows.forEach((row, idx) => {
        row[columnIndex] = formatCell(values[idx]);
      });
      render();
    },

    applyFieldValues: function(fieldLabels, valueLookup, options = {}) {
      if (!Array.isArray(fieldLabels) || fieldLabels.length === 0) {
        return;
      }
      const mode = options.mode === 'index' ? 'index' : 'end';
      let insertionIndex = mode === 'index'
        ? Math.max(0, Math.min(options.index ?? columns.length, columns.length))
        : columns.length;

      fieldLabels.forEach(label => {
        const targetIndex = this.ensureColumn(label, insertionIndex);
        const values = valueLookup[label] || [];
        rows.forEach((row, rowIdx) => {
          row[targetIndex] = formatCell(values[rowIdx]);
        });
        if (mode === 'index' && targetIndex >= insertionIndex) {
          insertionIndex = targetIndex + 1;
        }
      });

      render();
    },

    findEmailColumn: function() {
      if (!columns.length) return -1;
      const preferredMatches = ['email', 'mail', 'userprincipalname', 'signinname', 'upn'];
      const normalizedColumns = columns.map(normalizeColumnName);

      for (const match of preferredMatches) {
        const index = normalizedColumns.findIndex(col => col.includes(match));
        if (index !== -1) return index;
      }

      return normalizedColumns.findIndex(col => /mail|email/.test(col));
    },

    getDataForExport: function() {
      return {
        columns: this.getColumns(),
        rows: this.getRows()
      };
    }
  };
})();
