// Microsoft Graph API integration for profile enrichment
const GraphAPI = (function() {
  'use strict';

  const graphBaseUrl = 'https://graph.microsoft.com/v1.0';
  const fieldDefinitions = [
    { key: 'displayName', label: 'Display Name' },
    { key: 'jobTitle', label: 'Job Title' },
    { key: 'department', label: 'Department' },
    { key: 'officeLocation', label: 'Office Location' },
    { key: 'companyName', label: 'Company Name' },
    { key: 'mail', label: 'Mail' },
    { key: 'userPrincipalName', label: 'User Principal Name' },
    { key: 'businessPhones', label: 'Business Phones' },
    { key: 'mobilePhone', label: 'Mobile Phone' },
    { key: 'preferredLanguage', label: 'Preferred Language' },
    { key: 'givenName', label: 'Given Name' },
    { key: 'surname', label: 'Surname' }
  ];

  let accessToken = '';
  let loginDomain = '';

  function decodeJwtPayload(token) {
    if (!token || typeof token !== 'string') {
      return null;
    }

    const parts = token.split('.');
    if (parts.length < 2) {
      return null;
    }

    try {
      const base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      const padded = base64 + '='.repeat((4 - (base64.length % 4)) % 4);
      const decoded = atob(padded);
      return JSON.parse(decoded);
    } catch (error) {
      console.warn('Unable to decode access token payload:', error);
      return null;
    }
  }

  function extractDomainFromToken(token) {
    const payload = decodeJwtPayload(token);
    if (!payload) {
      return '';
    }

    const domainSource = payload.preferred_username || payload.upn || payload.email || '';
    if (typeof domainSource !== 'string') {
      return '';
    }

    const atIndex = domainSource.lastIndexOf('@');
    if (atIndex === -1) {
      return '';
    }

    return domainSource.slice(atIndex + 1).toLowerCase();
  }

  function updateDetectedDomainHint() {
    const element = document.getElementById('detectedDomainHint');
    if (!element) return;

    if (loginDomain) {
      element.textContent = `Detected sign-in domain: @${loginDomain}`;
    } else {
      element.textContent = 'Detected sign-in domain: not detected yet.';
    }
  }

  function buildLookupContext(emailValue) {
    const rawEmail = typeof emailValue === 'string' ? emailValue : '';
    const normalizedEmail = rawEmail.trim();

    if (!normalizedEmail) {
      return {
        rawEmail,
        normalizedEmail: '',
        identifiers: []
      };
    }

    const identifiers = [];
    identifiers.push(normalizedEmail);

    const normalizedDomain = loginDomain ? loginDomain.replace(/^@/, '').toLowerCase() : '';
    const separatorIndex = normalizedEmail.indexOf('@');
    const localPart = separatorIndex !== -1 ? normalizedEmail.slice(0, separatorIndex) : normalizedEmail;

    if (normalizedDomain && localPart) {
      const swappedIdentifier = `${localPart}@${normalizedDomain}`;
      const lowerSwapped = swappedIdentifier.toLowerCase();
      if (!identifiers.some(value => value.toLowerCase() === lowerSwapped)) {
        identifiers.push(swappedIdentifier);
      }
    }

    return {
      rawEmail,
      normalizedEmail,
      identifiers
    };
  }

  function getSelectedFields() {
    const checkboxContainer = document.getElementById('fieldCheckboxes');
    if (!checkboxContainer) return [];

    return Array.from(checkboxContainer.querySelectorAll('input[type="checkbox"]:checked'))
      .map(input => input.value);
  }

  function updateStatus(elementId, message, type = 'info') {
    const element = document.getElementById(elementId);
    if (!element) return;

    if (!message) {
      element.textContent = '';
      element.className = 'status-message';
      element.style.display = 'none';
      return;
    }

    element.textContent = message;
    element.className = `status-message status-${type}`;
    element.style.display = 'flex';
  }

  function showLoading(message) {
    const overlay = document.getElementById('loadingOverlay');
    const text = document.getElementById('loadingText');
    if (overlay) overlay.classList.add('active');
    if (text) text.textContent = message || 'Processing...';
  }

  function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.classList.remove('active');
  }

  function parseCsvFile(file) {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        complete: results => {
          if (results.errors && results.errors.length > 0) {
            reject(new Error(results.errors[0].message));
            return;
          }
          const data = Array.isArray(results.data) ? results.data : [];
          resolve(cleanParsedRows(data));
        },
        error: error => reject(error)
      });
    });
  }

  function parseExcelFile(file) {
    return file.arrayBuffer().then(buffer => {
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      if (!firstSheet) {
        throw new Error('The workbook does not contain any sheets.');
      }
      const worksheet = workbook.Sheets[firstSheet];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      return cleanParsedRows(rows);
    });
  }

  function cleanParsedRows(rows) {
    return rows
      .map(row => Array.isArray(row) ? row.map(cell => (cell === null || cell === undefined ? '' : cell)) : [])
      .filter(row => row.some(cell => cell !== '' && cell !== null && cell !== undefined));
  }

  function handleFileUpload(event) {
    const file = event.target.files ? event.target.files[0] : null;
    if (!file) return;

    const extension = file.name.split('.').pop().toLowerCase();
    const supported = ['csv', 'xlsx', 'xls'];
    if (!supported.includes(extension)) {
      updateStatus('loadStatus', 'Unsupported file type. Upload a CSV or Excel file.', 'error');
      event.target.value = '';
      return;
    }
    const uploadArea = document.getElementById('fileUploadArea');
    if (uploadArea) uploadArea.classList.remove('has-file');

    updateStatus('loadStatus', 'Loading file...', 'info');
    showLoading('Reading uploaded file...');

    const parser = extension === 'csv' ? parseCsvFile : parseExcelFile;

    parser(file)
      .then(rows => {
        if (!rows.length) {
          throw new Error('The file did not contain any data.');
        }
        const [headerRow, ...dataRows] = rows;
        if (!headerRow || !headerRow.length) {
          throw new Error('The first row of the file must contain headers.');
        }

        const columns = headerRow.map((cell, idx) => {
          const label = cell === null || cell === undefined || cell === '' ? `Column ${idx + 1}` : cell;
          return label.toString();
        });

        const filteredRows = dataRows
          .map(row => columns.map((_, idx) => {
            const value = row[idx];
            return value === null || value === undefined ? '' : value;
          }))
          .filter(row => row.some(cell => cell !== '' && cell !== null && cell !== undefined));

        if (!filteredRows.length) {
          throw new Error('No data rows were found after the header.');
        }

        DataTable.loadData(columns, filteredRows);
        updateStatus('loadStatus', `${file.name} loaded successfully (${filteredRows.length} rows).`, 'success');
        if (uploadArea) uploadArea.classList.add('has-file');
        setSectionCollapsed('loadSection', true);
        updateStatus('fetchStatus', '');
        updateDownloadButtons();
        updateFetchButtonState();
      })
      .catch(error => {
        console.error('Error loading file:', error);
        updateStatus('loadStatus', error.message || 'Unable to read the uploaded file.', 'error');
      })
      .finally(() => hideLoading());
  }

  function handlePasteLoad() {
    const textarea = document.getElementById('pasteEmails');
    if (!textarea) return;

    const raw = textarea.value.trim();
    if (!raw) {
      updateStatus('loadStatus', 'Paste one or more email addresses to load the table.', 'error');
      return;
    }

    const emails = raw
      .split(/\s|,|;|\n|\r/)
      .map(value => value.trim())
      .filter(value => value.length > 0);

    if (!emails.length) {
      updateStatus('loadStatus', 'No valid email addresses were detected.', 'error');
      return;
    }

    const uniqueEmails = Array.from(new Set(emails));
    const rows = uniqueEmails.map(email => [email]);
    DataTable.loadData(['Email'], rows);
    updateStatus('loadStatus', `Loaded ${rows.length} email${rows.length === 1 ? '' : 's'} from pasted list.`, 'success');
    const uploadArea = document.getElementById('fileUploadArea');
    if (uploadArea) uploadArea.classList.remove('has-file');
    setSectionCollapsed('loadSection', true);
    updateStatus('fetchStatus', '');
    updateDownloadButtons();
    updateFetchButtonState();
  }

  function updateAppendModeUI() {
    const insertInput = document.getElementById('insertIndex');
    const selectedMode = document.querySelector('input[name="appendMode"]:checked');
    if (!insertInput) return;

    if (selectedMode && selectedMode.value === 'index') {
      insertInput.disabled = false;
    } else {
      insertInput.disabled = true;
      insertInput.value = '';
    }
  }

  function getAppendMode() {
    const selected = document.querySelector('input[name="appendMode"]:checked');
    if (!selected || selected.value === 'end') {
      return { mode: 'end', index: DataTable.getColumnCount() };
    }

    const insertInput = document.getElementById('insertIndex');
    if (!insertInput || !insertInput.value) {
      return { mode: 'invalid' };
    }

    const parsed = Number.parseInt(insertInput.value, 10);
    if (!Number.isFinite(parsed) || parsed < 1) {
      return { mode: 'invalid' };
    }

    return { mode: 'index', index: Math.min(parsed - 1, DataTable.getColumnCount()) };
  }

  function updateDownloadButtons() {
    const hasData = DataTable.hasData();
    const csvButton = document.getElementById('downloadCsvButton');
    const excelButton = document.getElementById('downloadExcelButton');
    if (csvButton) csvButton.disabled = !hasData;
    if (excelButton) excelButton.disabled = !hasData;
  }

  function updateFetchButtonState() {
    const fetchButton = document.getElementById('fetchButton');
    if (!fetchButton) return;

    const hasToken = accessToken.length > 0;
    const hasTableData = DataTable.hasData();
    const hasFields = getSelectedFields().length > 0;
    fetchButton.disabled = !(hasToken && hasTableData && hasFields);
  }

  function downloadCsv() {
    const exportData = DataTable.getDataForExport();
    if (!exportData.columns.length || !exportData.rows.length) return;

    const csv = Papa.unparse({ fields: exportData.columns, data: exportData.rows });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `profile-data-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  function downloadExcel() {
    const exportData = DataTable.getDataForExport();
    if (!exportData.columns.length || !exportData.rows.length) return;

    const worksheet = XLSX.utils.aoa_to_sheet([
      exportData.columns,
      ...exportData.rows
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Profile Data');
    XLSX.writeFile(workbook, `profile-data-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function formatFieldValue(fieldKey, rawValue) {
    if (rawValue === undefined || rawValue === null) {
      return '';
    }
    if (Array.isArray(rawValue)) {
      return rawValue.filter(Boolean).join('; ');
    }
    if (typeof rawValue === 'object') {
      return JSON.stringify(rawValue);
    }
    return rawValue.toString();
  }

  async function fetchUserProfile(email, fields) {
    const selectFields = fields.join(',');
    const endpoint = `${graphBaseUrl}/users/${encodeURIComponent(email)}?$select=${selectFields}`;
    const response = await fetch(endpoint, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    if (response.status === 404) {
      return null;
    }

    if (!response.ok) {
      const errorBody = await response.json().catch(() => ({}));
      const message = errorBody.error && errorBody.error.message
        ? errorBody.error.message
        : `Request failed with status ${response.status}`;
      throw new Error(message);
    }

    return response.json();
  }

  async function fetchProfileData() {
    updateStatus('fetchStatus', '');

    if (!accessToken) {
      updateStatus('fetchStatus', 'Provide a Microsoft Graph access token to fetch profile data.', 'error');
      return;
    }

    if (!DataTable.hasData()) {
      updateStatus('fetchStatus', 'Load a CSV, Excel file, or paste emails before fetching profile data.', 'error');
      return;
    }

    const selectedFields = getSelectedFields();
    if (!selectedFields.length) {
      updateStatus('fetchStatus', 'Select at least one profile field to append.', 'error');
      return;
    }

    const appendMode = getAppendMode();
    if (appendMode.mode === 'invalid') {
      updateStatus('fetchStatus', 'Enter a valid column position (1 or greater) to insert new data.', 'error');
      return;
    }

    const emailColumnIndex = DataTable.findEmailColumn();
    if (emailColumnIndex === -1) {
      updateStatus('fetchStatus', 'Unable to find an email column. Ensure a column named "Email" or "UserPrincipalName" exists.', 'error');
      return;
    }

    const rowCount = DataTable.getRowCount();
    if (!rowCount) {
      updateStatus('fetchStatus', 'No rows available to process.', 'error');
      return;
    }

    showLoading(`Retrieving Microsoft 365 profiles (0 of ${rowCount})...`);

    const fieldLabels = selectedFields.map(field => {
      const definition = fieldDefinitions.find(item => item.key === field);
      return definition ? definition.label : field;
    });

    const valuesByField = fieldLabels.reduce((acc, label) => {
      acc[label] = new Array(rowCount).fill('');
      return acc;
    }, {});

    const errors = [];

    try {
      for (let index = 0; index < rowCount; index++) {
        const emailCellValue = DataTable.getCellValue(index, emailColumnIndex);
        const lookupContext = buildLookupContext(emailCellValue);
        const normalizedEmail = lookupContext.normalizedEmail;
        const displayEmail = (lookupContext.rawEmail && lookupContext.rawEmail.trim()) || normalizedEmail;

        if (!normalizedEmail) {
          errors.push(`Row ${index + 1}: missing email address.`);
          continue;
        }

        const currentPosition = index + 1;
        showLoading(`Retrieving Microsoft 365 profiles (${currentPosition} of ${rowCount})...`);
        updateStatus('fetchStatus', `Fetching profile ${currentPosition} of ${rowCount} (${displayEmail})...`, 'info');

        try {
          const lookupIdentifiers = lookupContext.identifiers;
          if (!lookupIdentifiers.length) {
            errors.push(`${displayEmail}: unable to determine lookup identifier.`);
            continue;
          }

          let profile = null;
          let lastError = null;

          for (const identifier of lookupIdentifiers) {
            try {
              const result = await fetchUserProfile(identifier, selectedFields);
              if (result) {
                profile = result;
                break;
              }
            } catch (error) {
              lastError = error;
            }
          }

          if (!profile) {
            const attempts = lookupIdentifiers.join(', ');
            if (lastError) {
              errors.push(`${displayEmail}: ${lastError.message} (tried ${attempts})`);
            } else {
              errors.push(`${displayEmail}: user not found (tried ${attempts})`);
            }
            continue;
          }

          selectedFields.forEach(fieldKey => {
            const definition = fieldDefinitions.find(item => item.key === fieldKey);
            const label = definition ? definition.label : fieldKey;
            valuesByField[label][index] = formatFieldValue(fieldKey, profile[fieldKey]);
          });
        } catch (error) {
          errors.push(`${displayEmail}: ${error.message}`);
        }
      }

      DataTable.applyFieldValues(fieldLabels, valuesByField, appendMode);

      if (errors.length) {
        const errorSummary = errors.slice(0, 5).join(' | ');
        const details = errors.length > 5 ? `${errorSummary} | ...` : errorSummary;
        updateStatus('fetchStatus', `Completed with ${errors.length} issue${errors.length === 1 ? '' : 's'}. Details: ${details}`, 'warning');
        console.warn('Profile fetch issues:', errors);
      } else {
        updateStatus('fetchStatus', 'Profile data appended successfully.', 'success');
      }

      updateDownloadButtons();
    } catch (error) {
      console.error('Error fetching profile data:', error);
      updateStatus('fetchStatus', error.message || 'Failed to fetch profile data.', 'error');
    } finally {
      hideLoading();
      updateFetchButtonState();
    }
  }

  function attachEventListeners() {
    const tokenInput = document.getElementById('graphToken');
    if (tokenInput) {
      tokenInput.addEventListener('input', event => {
        accessToken = event.target.value.trim();
        loginDomain = accessToken ? extractDomainFromToken(accessToken) : '';
        updateDetectedDomainHint();
        updateFetchButtonState();
      });
    }

    const fileInput = document.getElementById('dataFile');
    if (fileInput) {
      fileInput.addEventListener('change', handleFileUpload);
    }

    const pasteButton = document.getElementById('pasteLoadButton');
    if (pasteButton) {
      pasteButton.addEventListener('click', handlePasteLoad);
    }

    const appendModeRadios = document.querySelectorAll('input[name="appendMode"]');
    appendModeRadios.forEach(radio => {
      radio.addEventListener('change', () => {
        updateAppendModeUI();
        updateFetchButtonState();
      });
    });

    const checkboxContainer = document.getElementById('fieldCheckboxes');
    if (checkboxContainer) {
      checkboxContainer.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
        checkbox.addEventListener('change', updateFetchButtonState);
      });
    }

    const fetchButton = document.getElementById('fetchButton');
    if (fetchButton) {
      fetchButton.addEventListener('click', fetchProfileData);
    }

    const csvButton = document.getElementById('downloadCsvButton');
    if (csvButton) {
      csvButton.addEventListener('click', downloadCsv);
    }

    const excelButton = document.getElementById('downloadExcelButton');
    if (excelButton) {
      excelButton.addEventListener('click', downloadExcel);
    }

    document.addEventListener('datatable:change', () => {
      updateDownloadButtons();
      updateFetchButtonState();
    });
  }

  return {
    initialize: function() {
      attachEventListeners();
      updateAppendModeUI();
      updateDownloadButtons();
      updateDetectedDomainHint();
      updateFetchButtonState();
    },

    updateDownloadButtons,
    updateFetchButtonState
  };
})();
