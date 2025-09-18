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
  let deleteButton = null;
  let deleteButtonBaseLabel = '';
  let selectAllCheckbox = null;
  let sortState = { columnIndex: null, direction: 'asc' };
  const selectedRows = new Set();
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });

  function resetSortState() {
    sortState = { columnIndex: null, direction: 'asc' };
  }

  function notifyChange(type) {
    document.dispatchEvent(new CustomEvent('datatable:change', {
      detail: { type }
    }));
  }

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
    let summary = `${rows.length} ${rowLabel} • ${columns.length} ${colLabel}`;
    if (selectedRows.size > 0) {
      summary += ` • ${selectedRows.size} selected`;
    }
    summaryElement.textContent = summary;
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

  function updateDeleteButtonState() {
    if (!deleteButton) return;
    const baseLabel = deleteButtonBaseLabel || deleteButton.textContent.trim();
    const count = selectedRows.size;
    deleteButton.disabled = count === 0;
    deleteButton.textContent = count > 0 ? `${baseLabel} (${count})` : baseLabel;
  }

  function updateSelectAllState() {
    if (!selectAllCheckbox) return;
    const total = rows.length;
    const selectedCount = selectedRows.size;
    selectAllCheckbox.checked = total > 0 && selectedCount === total;
    selectAllCheckbox.indeterminate = selectedCount > 0 && selectedCount < total;
  }

  function toggleRowSelection(row, shouldSelect) {
    if (!row) return;
    if (shouldSelect) {
      selectedRows.add(row);
    } else {
      selectedRows.delete(row);
    }

    const index = rows.indexOf(row);
    if (index !== -1 && tableBody && tableBody.rows[index]) {
      const tr = tableBody.rows[index];
      const checkbox = tr.querySelector('input[type="checkbox"]');
      tr.classList.toggle('row-selected', selectedRows.has(row));
      if (checkbox) {
        checkbox.checked = selectedRows.has(row);
      }
    }

    updateSelectAllState();
    updateDeleteButtonState();
    updateSummary();
  }

  function handleSelectAll(checked) {
    if (checked) {
      rows.forEach(row => selectedRows.add(row));
    } else {
      selectedRows.clear();
    }

    if (tableBody) {
      Array.from(tableBody.rows).forEach((tr, idx) => {
        const row = rows[idx];
        const isSelected = Boolean(row) && checked;
        tr.classList.toggle('row-selected', isSelected);
        const checkbox = tr.querySelector('input[type="checkbox"]');
        if (checkbox) {
          checkbox.checked = isSelected;
        }
      });
    }

    updateSelectAllState();
    updateDeleteButtonState();
    updateSummary();
  }

  function handleSort(columnIndex) {
    if (columnIndex < 0 || columnIndex >= columns.length) return;

    const nextDirection =
      sortState.columnIndex === columnIndex && sortState.direction === 'asc'
        ? 'desc'
        : 'asc';
    sortState = { columnIndex, direction: nextDirection };

    const directionFactor = nextDirection === 'asc' ? 1 : -1;
    rows = rows
      .map((row, idx) => ({ row, idx }))
      .sort((a, b) => {
        const valueA = formatCell(a.row[columnIndex]);
        const valueB = formatCell(b.row[columnIndex]);
        const comparison = collator.compare(valueA, valueB);
        if (comparison !== 0) {
          return directionFactor * comparison;
        }
        return directionFactor * (a.idx - b.idx);
      })
      .map(item => item.row);

    render();
  }

  function removeSelectedRows() {
    if (!selectedRows.size) return;
    rows = rows.filter(row => !selectedRows.has(row));
    selectedRows.clear();
    render();
    updateSelectAllState();
    updateDeleteButtonState();
    notifyChange('data');
  }

  function render() {
    if (!tableHead || !tableBody) return;

    const scrollTop = tableWrapper ? tableWrapper.scrollTop : 0;

    selectAllCheckbox = null;
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    if (!hasData()) {
      updateSummary();
      updateEmptyState();
      updateDeleteButtonState();
      updateSelectAllState();
      if (tableWrapper) {
        tableWrapper.scrollTop = 0;
      }
      return;
    }

    const headerRow = document.createElement('tr');

    const selectHeader = document.createElement('th');
    selectHeader.classList.add('select-column');
    const selectInput = document.createElement('input');
    selectInput.type = 'checkbox';
    selectInput.addEventListener('click', event => event.stopPropagation());
    selectInput.addEventListener('change', () => handleSelectAll(selectInput.checked));
    selectHeader.appendChild(selectInput);
    headerRow.appendChild(selectHeader);
    selectAllCheckbox = selectInput;

    columns.forEach((column, columnIndex) => {
      const th = document.createElement('th');
      th.classList.add('sortable');

      const headerContent = document.createElement('div');
      headerContent.className = 'header-content';

      const headerLabel = document.createElement('span');
      headerLabel.className = 'header-label';
      headerLabel.textContent = column;

      const indicator = document.createElement('span');
      indicator.className = 'sort-indicator';
      if (sortState.columnIndex === columnIndex) {
        th.classList.add(sortState.direction === 'asc' ? 'sorted-asc' : 'sorted-desc');
        indicator.textContent = sortState.direction === 'asc' ? '▲' : '▼';
      } else {
        indicator.textContent = '↕';
      }

      headerContent.appendChild(headerLabel);
      headerContent.appendChild(indicator);
      th.appendChild(headerContent);
      th.addEventListener('click', event => {
        event.preventDefault();
        handleSort(columnIndex);
      });

      headerRow.appendChild(th);
    });

    tableHead.appendChild(headerRow);

    rows.forEach(row => {
      const tr = document.createElement('tr');

      const selectCell = document.createElement('td');
      selectCell.classList.add('select-column');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = selectedRows.has(row);
      checkbox.addEventListener('click', event => event.stopPropagation());
      checkbox.addEventListener('change', () => toggleRowSelection(row, checkbox.checked));
      selectCell.appendChild(checkbox);
      tr.appendChild(selectCell);

      if (selectedRows.has(row)) {
        tr.classList.add('row-selected');
      }

      columns.forEach((_, idx) => {
        const td = document.createElement('td');
        td.textContent = formatCell(row[idx]);
        tr.appendChild(td);
      });

      tr.addEventListener('click', event => {
        const target = event.target;
        if (!target) return;
        if (target.tagName === 'A' || target.tagName === 'BUTTON' || target.type === 'checkbox') {
          return;
        }
        toggleRowSelection(row, !selectedRows.has(row));
      });

      tableBody.appendChild(tr);
    });

    updateSummary();
    updateEmptyState();
    updateSelectAllState();
    updateDeleteButtonState();

    if (tableWrapper) {
      tableWrapper.scrollTop = scrollTop;
    }
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
      deleteButton = document.getElementById('deleteSelectedButton');
      if (deleteButton) {
        deleteButtonBaseLabel = deleteButton.dataset.label || deleteButton.textContent.trim();
        deleteButton.addEventListener('click', removeSelectedRows);
      }
      this.clear();
    },

    clear: function() {
      columns = [];
      rows = [];
      selectedRows.clear();
      resetSortState();
      render();
      notifyChange('data');
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
      selectedRows.clear();
      resetSortState();
      render();
      notifyChange('data');
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
      notifyChange('data');
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
      notifyChange('data');
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
    },

    deleteSelectedRows: removeSelectedRows
  };
})();
