/**
 * FIELD TRIALS MANAGEMENT MODULE
 */

(function (webapi, $) {

  // Make sure validateLoginSession exists
  if (typeof window.validateLoginSession !== "function") {
    window.validateLoginSession = function (data, textStatus, jqXHR, onOk) {
      if (typeof onOk === "function") onOk(data, textStatus, jqXHR);
    };
  }

  /**
   * GLOBAL safeAjax for Power Pages
   * Always injects anti-forgery token automatically.
   */
  function safeAjax(options) {
    var d = $.Deferred();

    if (!(window.shell && typeof shell.getTokenDeferred === "function")) {
      d.reject({
        status: 401,
        statusText: "shell.getTokenDeferred unavailable on this page"
      });
      return d.promise();
    }

    shell.getTokenDeferred().done(function (token) {

      const finalOptions = {
        type: options.type,
        url: options.url,
        data: options.data || null,
        contentType: "application/json; charset=utf-8",
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json; charset=utf-8",
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          "__RequestVerificationToken": token
        }
      };

      $.ajax(finalOptions)
        .done((data, textStatus, jqXHR) => {
          validateLoginSession(data, textStatus, jqXHR, () => {
            d.resolve(data);
          });
        })
        .fail((err) => d.reject(err));

    }).fail(() => d.reject());

    return d.promise();
  }

  // Expose global safeAjax
  webapi.safeAjax = safeAjax;

})(window.webapi = window.webapi || {}, jQuery);


/* ============================================================
   FIELD TRIALS MANAGER MODULE
============================================================ */
window.CURRENT_TSR_REPORT_ID = null;

const FieldTrialsManager = (function () {
  "use strict";

  const CONFIG = {
    API_BASE: "/_api",
    ENTITY_SET: "cr650_tsefieldtrialses",
    TRIALS_PER_PAGE: 100,
    CHART_COLORS: {
      planned: "#F59E0B",    // Amber
      ongoing: "#3B82F6",    // Blue
      completed: "#22C55E"   // Green
    }
  };

  let state = {
    trials: [],
    filtered: [],
    chart: null,
    editing: false,
    currentId: null,
    initialized: false
  };

  /* ------------------------------
     API WRAPPER (always safeAjax)
  ------------------------------ */
  function api(options) {
    if (!window.webapi || !window.webapi.safeAjax) {
      throw new Error("safeAjax is not available on this page.");
    }
    return window.webapi.safeAjax(options);
  }

  /* ------------------------------
     Initialization
  ------------------------------ */
  function init() {
    if (state.initialized) return;
    state.initialized = true;

    console.log("Initializing Field Trials Manager…");

    // Check dependencies
    if (typeof Chart === "undefined") {
      console.error("Chart.js not loaded");
      return;
    }

    if (typeof $ === "undefined") {
      console.error("jQuery not loaded");
      return;
    }

    // Attach form handler with prevention
    attachFormHandler();

    populateMonthFilter();
    loadData();
  }

  /* ------------------------------
     Form Handler - Triple Prevention
  ------------------------------ */
  function attachFormHandler() {
    const form = document.getElementById("trialForm");
    if (!form) {
      console.warn("Form not found");
      return;
    }

    // Remove form action
    form.removeAttribute("action");

    // Prevent default submission
    form.onsubmit = function (e) {
      e.preventDefault();
      e.stopPropagation();
      return false;
    };

    // Attach submit listener
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      e.stopPropagation();
      saveTrial(e);
      return false;
    }, true);

    console.log("✓ Form handler attached");
  }

  /* ------------------------------
     Load Data
  ------------------------------ */
  function loadData() {
    const loadingState = document.getElementById("tableLoadingState");
    if (loadingState) loadingState.style.display = "block";

    // Build base query
    let query =
      CONFIG.API_BASE + "/" + CONFIG.ENTITY_SET + "?" +
      "$select=cr650_tsefieldtrialsid,cr650_name,cr650_customername,cr650_startdate,cr650_closingdate," +
      "cr650_product,cr650_lob,cr650_comments,cr650_status,cr650_submittedby&" +
      "$orderby=cr650_startdate desc&$top=" + CONFIG.TRIALS_PER_PAGE;

    // Add user filter ONLY if user is available
    const currentUser = window.PORTAL_USER_NAME;
    if (currentUser && currentUser !== 'Portal User' && currentUser !== '') {
      const escapedUser = currentUser.replace(/'/g, "''");
      query += "&$filter=cr650_submittedby eq '" + escapedUser + "'";
      console.log("Filtering trials for user:", currentUser);
    } else {
      console.warn("No user context - loading all trials");
    }

    api({ type: "GET", url: query })
      .then((res) => {
        state.trials = res.value || [];
        state.filtered = state.trials;
        renderTable();
        renderSummary();
        renderChart();
      })
      .fail((err) => {
        console.error("Load failed:", err);
        alert("Failed to load field trials. Please refresh the page.");
      })
      .always(() => {
        if (loadingState) loadingState.style.display = "none";
      });
  }

  /* ------------------------------
     Rendering – Table
  ------------------------------ */
  function renderTable() {
    const tbody = document.getElementById("trialsTableBody");
    if (!tbody) return;

    tbody.innerHTML = "";

    // Update results count
    const resultCount = document.getElementById("tableResultsCount");
    if (resultCount) {
      resultCount.innerHTML =
        "Showing <strong>" + state.filtered.length + "</strong> of <strong>" + state.trials.length + "</strong> trials";
    }

    if (state.filtered.length === 0) {
      tbody.innerHTML =
        '<tr class="empty-row">' +
        '  <td colspan="9" class="empty-state">' +
        '    <div class="empty-icon">' +
        '      <i class="fas fa-inbox"></i>' +
        '    </div>' +
        '    <p class="empty-text">No field trials found</p>' +
        '    <p class="empty-subtext">Try adjusting your filters or add a new trial</p>' +
        '  </td>' +
        '</tr>';
      return;
    }

    state.filtered.forEach((t, i) => {
      const row = document.createElement("tr");

      const customerName = escapeHtml(t.cr650_customername) || '-';
      const resolutionText = escapeHtml(t.cr650_comments) || 'No comment provided';
      const resolutionDisplay = escapeHtml(t.cr650_comments) || '-';

      row.innerHTML =
        '<td>' + (i + 1) + '</td>' +
        '<td>' + customerName + '</td>' +
        '<td>' + escapeHtml(t.cr650_product) + '</td>' +
        '<td>' + lobLabel(t.cr650_lob) + '</td>' +  // ⭐ NEW: LOB column
        '<td class="resolution-cell" title="' + resolutionText + '">' + resolutionDisplay + '</td>' +
        '<td>' + formatDate(t.cr650_startdate) + '</td>' +
        '<td>' + formatDate(t.cr650_closingdate) + '</td>' +
        '<td>' + statusBadge(t.cr650_status) + '</td>' +
        '<td class="actions-cell">' +
        '  <button onclick="FieldTrialsManager.editTrial(\'' + t.cr650_tsefieldtrialsid + '\')" class="btn-edit" title="Edit">' +
        '    <i class="fas fa-edit"></i>' +
        '  </button>' +
        '  <button onclick="FieldTrialsManager.deleteTrial(\'' + t.cr650_tsefieldtrialsid + '\')" class="btn-delete" title="Delete">' +
        '    <i class="fas fa-trash"></i>' +
        '  </button>' +
        '</td>';

      tbody.appendChild(row);
    });
  }

  /* ------------------------------
     Rendering – Summary
  ------------------------------ */
  function renderSummary() {
    const planned = state.filtered.filter(t => t.cr650_status === 1).length;
    const ongoing = state.filtered.filter(t => t.cr650_status === 2).length;
    const completed = state.filtered.filter(t => t.cr650_status === 3).length;

    const plannedCount = document.getElementById("plannedTrialsCount");
    const ongoingCount = document.getElementById("ongoingTrialsCount");
    const completedCount = document.getElementById("completedTrialsCount");
    const totalCount = document.getElementById("totalTrialsCount");

    if (plannedCount) plannedCount.textContent = planned;
    if (ongoingCount) ongoingCount.textContent = ongoing;
    if (completedCount) completedCount.textContent = completed;
    if (totalCount) totalCount.textContent = state.filtered.length;
  }

  /* ------------------------------
     Rendering – Chart
  ------------------------------ */
  function renderChart() {
    const ctx = document.getElementById("trialsStatusChart");
    if (!ctx) return;

    const planned = state.filtered.filter(t => t.cr650_status === 1).length;
    const ongoing = state.filtered.filter(t => t.cr650_status === 2).length;
    const completed = state.filtered.filter(t => t.cr650_status === 3).length;

    // Destroy existing chart
    if (state.chart) {
      state.chart.destroy();
    }

    // Create new chart
    state.chart = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['Planned / Pending', 'On-going', 'Completed'],
        datasets: [{
          data: [planned, ongoing, completed],
          backgroundColor: [
            CONFIG.CHART_COLORS.planned,
            CONFIG.CHART_COLORS.ongoing,
            CONFIG.CHART_COLORS.completed
          ],
          borderWidth: 2,
          borderColor: '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              padding: 20,
              font: {
                size: 14,
                weight: 600
              },
              usePointStyle: true,
              pointStyle: 'circle'
            }
          },
          tooltip: {
            callbacks: {
              label: function (context) {
                const label = context.label || '';
                const value = context.parsed || 0;
                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                return label + ': ' + value + ' (' + percentage + '%)';
              }
            }
          }
        }
      }
    });
  }

  /* ------------------------------
     Filters
  ------------------------------ */
  function populateMonthFilter() {
    const monthFilter = document.getElementById("monthFilter");
    if (!monthFilter) return;

    // Clear existing options (keep "All Months")
    while (monthFilter.options.length > 1) {
      monthFilter.remove(1);
    }

    // Get unique months from trials
    const months = new Set();
    state.trials.forEach(t => {
      if (t.cr650_startdate) {
        const date = new Date(t.cr650_startdate);
        const monthYear = date.toLocaleString('en-US', { month: 'long', year: 'numeric' });
        months.add(monthYear);
      }
    });

    // Sort and add options
    Array.from(months).sort((a, b) => {
      const dateA = new Date(a);
      const dateB = new Date(b);
      return dateB - dateA;
    }).forEach(month => {
      const option = document.createElement('option');
      option.value = month;
      option.textContent = month;
      monthFilter.appendChild(option);
    });
  }

  function applyFilters() {
    const searchInput = document.getElementById("searchInput");
    const monthFilter = document.getElementById("monthFilter");
    const statusFilter = document.getElementById("statusFilter");
    const lobFilter = document.getElementById("lobFilter");  // ⭐ NEW

    const searchTerm = searchInput ? searchInput.value.toLowerCase() : '';
    const selectedMonth = monthFilter ? monthFilter.value : '';
    const selectedStatus = statusFilter ? statusFilter.value : '';
    const selectedLob = lobFilter ? lobFilter.value : '';  // ⭐ NEW

    state.filtered = state.trials.filter(t => {
      // Search filter
      if (searchTerm) {
        const customerName = (t.cr650_customername || '').toLowerCase();
        const product = (t.cr650_product || '').toLowerCase();
        const comments = (t.cr650_comments || '').toLowerCase();

        if (!customerName.includes(searchTerm) &&
          !product.includes(searchTerm) &&
          !comments.includes(searchTerm)) {
          return false;
        }
      }

      // Month filter (based on starting date)
      if (selectedMonth && t.cr650_startdate) {
        const trialDate = new Date(t.cr650_startdate);
        const trialMonth = trialDate.toLocaleString('en-US', { month: 'long', year: 'numeric' });
        if (trialMonth !== selectedMonth) {
          return false;
        }
      }

      // Status filter
      if (selectedStatus) {
        if (t.cr650_status !== parseInt(selectedStatus)) {
          return false;
        }
      }

      // LOB filter ⭐ NEW
      if (selectedLob) {
        if (t.cr650_lob !== parseInt(selectedLob)) {
          return false;
        }
      }

      return true;
    });

    renderTable();
    renderSummary();
    renderChart();
  }

  /* ------------------------------
     Save Trial
  ------------------------------ */
  function saveTrial(event) {
    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }

    const saveBtn = document.getElementById("saveTrialBtn");
    if (saveBtn && saveBtn.disabled) return;

    const data = collectForm();
    if (!data) return;

    // Disable button
    if (saveBtn) {
      saveBtn.disabled = true;
      saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
    }

    if (state.editing && state.currentId) {
      // Update existing
      const url = CONFIG.API_BASE + "/" + CONFIG.ENTITY_SET + "(" + state.currentId + ")";
      api({ type: "PATCH", url: url, data: JSON.stringify(data) })
        .then(() => {
          closeModal();
          loadData();
          alert("Field trial updated successfully!");
        })
        .fail((err) => {
          console.error("Update failed:", err);
          alert("Failed to update trial. Please try again.");
        })
        .always(() => resetSaveButton());
    } else {
      // Create new
      const url = CONFIG.API_BASE + "/" + CONFIG.ENTITY_SET;
      api({ type: "POST", url: url, data: JSON.stringify(data) })
        .then(() => {
          closeModal();
          loadData();
          alert("Field trial added successfully!");
        })
        .fail((err) => {
          console.error("Create failed:", err);
          alert("Failed to add trial. Please try again.");
        })
        .always(() => resetSaveButton());
    }
  }

  /* ------------------------------
     Delete Trial
  ------------------------------ */
  function deleteTrial(id) {
    if (!confirm("Are you sure you want to delete this field trial?")) {
      return;
    }

    const url = CONFIG.API_BASE + "/" + CONFIG.ENTITY_SET + "(" + id + ")";
    api({ type: "DELETE", url: url })
      .then(() => {
        loadData();
        alert("Field trial deleted successfully!");
      })
      .fail((err) => {
        console.error("Delete failed:", err);
        alert("Failed to delete trial. Please try again.");
      });
  }

  /* ------------------------------
     Export to Excel
  ------------------------------ */
  function exportToExcel() {
    if (typeof XLSX === 'undefined') {
      alert('Excel export library not loaded. Please refresh the page.');
      return;
    }

    // Prepare data for export
    const exportData = state.filtered.map((t, i) => {
      return {
        'S.N.': i + 1,
        'Customer': t.cr650_customername || '-',
        'Product': t.cr650_product || '-',
        'Line of Business': lobLabelText(t.cr650_lob),
        'Resolution / Comment': t.cr650_comments || '-',
        'Starting Date': formatDate(t.cr650_startdate),
        'Closing Date': formatDate(t.cr650_closingdate),
        'Status': statusLabel(t.cr650_status),
        'Submitted By': t.cr650_submittedby || '-'
      };
    });

    // Create workbook and worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);

    // Set column widths
    ws['!cols'] = [
      { wch: 6 },   // S.N.
      { wch: 25 },  // Customer
      { wch: 20 },  // Product
      { wch: 15 },  // LOB
      { wch: 40 },  // Resolution
      { wch: 12 },  // Starting Date
      { wch: 12 },  // Closing Date
      { wch: 15 },  // Status
      { wch: 20 }   // Submitted By
    ];

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, "Field Trials");

    // Generate filename with current date
    const filename = "Field_Trials_Export_" + new Date().toISOString().split('T')[0] + ".xlsx";

    // Download file
    XLSX.writeFile(wb, filename);
  }

  /* ------------------------------
     Modal Management
  ------------------------------ */
  function openAddForm() {
    state.editing = false;
    state.currentId = null;

    const form = document.getElementById("trialForm");
    if (form) form.reset();

    const modalTitle = document.getElementById("modalTitle");
    if (modalTitle) {
      modalTitle.innerHTML = '<i class="fas fa-plus-circle"></i> Add New Field Trial';
    }

    const modal = document.getElementById("trialModal");
    if (modal) modal.style.display = "flex";

    toggleDates();
  }

  function editTrial(id) {
    state.editing = true;
    state.currentId = id;

    const t = state.trials.find(x => x.cr650_tsefieldtrialsid === id);
    if (!t) {
      alert("Trial not found");
      return;
    }

    // Populate form fields
    document.getElementById("customerName").value = t.cr650_customername || '';
    document.getElementById("product").value = t.cr650_product || '';
    document.getElementById("lob").value = t.cr650_lob || '';  // ⭐ NEW
    document.getElementById("status").value = t.cr650_status || '';
    document.getElementById("resolution").value = t.cr650_comments || '';

    if (t.cr650_startdate) {
      document.getElementById("startingDate").value = formatDateInput(t.cr650_startdate);
    }

    if (t.cr650_closingdate) {
      document.getElementById("closingDate").value = formatDateInput(t.cr650_closingdate);
    }

    toggleDates();

    const modalTitle = document.getElementById("modalTitle");
    if (modalTitle) {
      modalTitle.innerHTML = '<i class="fas fa-edit"></i> Edit Field Trial';
    }

    const modal = document.getElementById("trialModal");
    if (modal) modal.style.display = "flex";
  }

  function closeModal() {
    const modal = document.getElementById('trialModal');
    if (modal) modal.style.display = 'none';

    // Reset state
    state.editing = false;
    state.currentId = null;

    const form = document.getElementById('trialForm');
    if (form) form.reset();

    resetSaveButton();
  }

  /* ------------------------------
     Date Field Toggle Logic
     Business Rules:
     - Planned (1) or On-going (2): Starting Date required, Closing Date hidden
     - Completed (3): Both dates required
  ------------------------------ */
  function toggleDates() {
    const statusSelect = document.getElementById('status');
    const startingDateGroup = document.getElementById('startingDateGroup');
    const closingDateGroup = document.getElementById('closingDateGroup');
    const startingDateInput = document.getElementById('startingDate');
    const closingDateInput = document.getElementById('closingDate');

    if (!statusSelect || !startingDateGroup || !closingDateGroup) return;

    const status = parseInt(statusSelect.value);

    if (status === 1 || status === 2) {
      // Planned or On-going
      startingDateGroup.style.display = 'block';
      closingDateGroup.style.display = 'none';

      startingDateInput.required = true;
      closingDateInput.required = false;

      // Set default starting date to today if empty
      if (!startingDateInput.value) {
        startingDateInput.value = new Date().toISOString().split('T')[0];
      }
    } else if (status === 3) {
      // Completed
      startingDateGroup.style.display = 'block';
      closingDateGroup.style.display = 'block';

      startingDateInput.required = true;
      closingDateInput.required = true;

      // Set default dates if empty
      if (!startingDateInput.value) {
        startingDateInput.value = new Date().toISOString().split('T')[0];
      }
      if (!closingDateInput.value) {
        closingDateInput.value = new Date().toISOString().split('T')[0];
      }
    } else {
      // No status selected
      startingDateGroup.style.display = 'none';
      closingDateGroup.style.display = 'none';

      startingDateInput.required = false;
      closingDateInput.required = false;
    }
  }

  /* ------------------------------
     Utilities
  ------------------------------ */
  const STATUS_MAP = {
    1: "Planned / Pending",
    2: "On-going",
    3: "Completed"
  };

  function collectForm() {
    const customer = document.getElementById("customerName").value.trim();
    const product = document.getElementById("product").value.trim();
    const statusValue = document.getElementById("status").value;

    if (!customer || !product) {
      alert("Customer and product are required.");
      return null;
    }

    if (!statusValue) {
      alert("Please select a status.");
      return null;
    }

    const status = parseInt(statusValue);
    const startingDate = document.getElementById("startingDate").value;

    // Validate starting date for all statuses
    if (!startingDate) {
      alert("Starting date is required.");
      return null;
    }

    const lobValue = document.getElementById("lob").value;

    if (!lobValue) {
      alert("Please select Line of Business.");
      return null;
    }

    const data = {
      cr650_name: customer + " - " + product,
      cr650_customername: customer,
      cr650_product: product,
      cr650_startdate: startingDate,
      cr650_lob: parseInt(lobValue),
      cr650_status: status,
      cr650_comments: document.getElementById("resolution").value || '',
      cr650_submittedby: window.PORTAL_USER_NAME || 'Unknown User'
    };

    // Add closing date only if status is Completed
    if (status === 3) {
      const closingDate = document.getElementById("closingDate").value;
      if (!closingDate) {
        alert("Closing date is required for completed trials.");
        return null;
      }
      data.cr650_closingdate = closingDate;
    }

    return data;
  }

  function formatDate(d) {
    if (!d) return "-";
    return new Date(d).toLocaleDateString("en-US");
  }

  function formatDateInput(d) {
    if (!d) return "";
    return new Date(d).toISOString().split('T')[0];
  }

  function escapeHtml(text) {
    if (!text) return "";
    var map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, function (m) { return map[m]; });
  }

  function statusLabel(v) {
    return STATUS_MAP[v] || "-";
  }

  function statusBadge(v) {
    if (v === 1) {
      return '<span class="status planned">Planned / Pending</span>';
    } else if (v === 2) {
      return '<span class="status ongoing">On-going</span>';
    } else if (v === 3) {
      return '<span class="status completed">Completed</span>';
    }
    return '<span class="status">-</span>';
  }

  function refreshData() {
    console.log("Refreshing data...");
    loadData();
  }

  function resetSaveButton() {
    const saveBtn = document.getElementById("saveTrialBtn");
    if (saveBtn) {
      saveBtn.disabled = false;
      saveBtn.innerHTML = '<i class="fas fa-save"></i> Save Trial';
    }
  }

  function lobLabel(value) {
    const labels = {
      1: '<span class="lob-badge petromin">Petromin</span>',
      2: '<span class="lob-badge gulf">Gulf</span>',
      3: '<span class="lob-badge commercial">Commercial</span>'
    };
    return labels[value] || '<span class="lob-badge">-</span>';
  }

  function lobLabelText(value) {
    const labels = {
      1: 'Petromin',
      2: 'Gulf',
      3: 'Commercial'
    };
    return labels[value] || '-';
  }

  /* ------------------------------
     Public API
  ------------------------------ */
  return {
    init: init,
    loadData: loadData,
    applyFilters: applyFilters,
    saveTrial: saveTrial,
    editTrial: editTrial,
    openAddForm: openAddForm,
    closeModal: closeModal,
    toggleDates: toggleDates,
    refreshData: refreshData,
    deleteTrial: deleteTrial,
    exportToExcel: exportToExcel
  };

})();

/* ------------------------------
   Autostart
------------------------------ */
if (document.readyState === 'loading') {
  document.addEventListener("DOMContentLoaded", function () {
    FieldTrialsManager.init();
  });
} else {
  FieldTrialsManager.init();
}