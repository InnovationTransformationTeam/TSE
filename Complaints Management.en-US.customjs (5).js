/**
 * COMPLAINTS MANAGEMENT MODULE – BULLETPROOF VERSION
 * - Triple-layer form submission prevention
 * - Changed submit button to type="button"
 * - Added explicit form action prevention
 * - Uses ONLY Power Pages safeAjax
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
   COMPLAINTS MANAGER MODULE
   Cleaned, optimized, using ONLY safeAjax
============================================================ */

const ComplaintsManager = (function () {
  "use strict";

  const CONFIG = {
    API_BASE: "/_api",
    ENTITY_SET: "cr650_tsecomplaintses",
    COMPLAINTS_PER_PAGE: 100,
    CHART_COLORS: {
      open: "#EF4444",
      closed: "#22C55E"
    }
  };

  let state = {
    complaints: [],
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

    console.log("Initializing Complaints Manager…");

    // Check dependencies
    if (typeof Chart === "undefined") {
      console.error("Chart.js not loaded");
      return;
    }

    if (typeof $ === "undefined") {
      console.error("jQuery not loaded");
      return;
    }

    // CRITICAL: Attach form handler with triple-layer prevention
    attachFormHandler();
    
    populateMonthFilter();
    loadData();
  }

  /* ------------------------------
     Form Handler - TRIPLE PREVENTION LAYER
  ------------------------------ */
  function attachFormHandler() {
    const form = document.getElementById("complaintForm");
    if (!form) {
      console.warn("Form not found");
      return;
    }

    // LAYER 1: Remove form action attribute
    form.removeAttribute("action");
    
    // LAYER 2: Set form to not submit
    form.onsubmit = function(e) {
      e.preventDefault();
      e.stopPropagation();
      return false;
    };
    
    // LAYER 3: Attach submit event listener
    form.addEventListener("submit", function(e) {
      e.preventDefault();
      e.stopPropagation();
      saveComplaint(e);
      return false;
    }, true); // Use capture phase
    
    console.log("✓ Form handler attached with triple-layer prevention");
  }

  /* ------------------------------
     Load Data
  ------------------------------ */
  function loadData() {
    // Show loading state
    const loadingState = document.getElementById("tableLoadingState");
    if (loadingState) loadingState.style.display = "block";

    // Build base query
    let query =
      CONFIG.API_BASE + "/" + CONFIG.ENTITY_SET + "?" +
      "$select=cr650_tsecomplaintsid,cr650_name,cr650_complaintdate,cr650_closingdate," +
      "cr650_product,cr650_complaintnature,cr650_resolution,cr650_status,cr650_customername,cr650_lob,cr650_submittedby&" +
      "$orderby=cr650_complaintdate desc&$top=" + CONFIG.COMPLAINTS_PER_PAGE;

    // Add user filter ONLY if user is available
    const currentUser = window.PORTAL_USER_NAME;
    if (currentUser && currentUser !== 'Portal User' && currentUser !== '') {
      const escapedUser = currentUser.replace(/'/g, "''");
      query += "&$filter=cr650_submittedby eq '" + escapedUser + "'";
      console.log("Filtering complaints for user:", currentUser);
    } else {
      console.warn("No user context - loading all complaints");
    }

    api({ type: "GET", url: query })
      .then((res) => {
        state.complaints = res.value || [];
        state.filtered = state.complaints;
        renderTable();
        renderSummary();
        renderChart();
      })
      .fail((err) => {
        console.error("Load failed:", err);
        alert("Failed to load complaints. Please refresh the page.");
      })
      .always(() => {
        if (loadingState) loadingState.style.display = "none";
      });
  }

  /* ------------------------------
     Rendering – Table
  ------------------------------ */
  function renderTable() {
    const tbody = document.getElementById("complaintsTableBody");
    if (!tbody) return;

    tbody.innerHTML = "";

    // Update results count
    const resultCount = document.getElementById("tableResultsCount");
    if (resultCount) {
      resultCount.innerHTML = 
        "Showing <strong>" + state.filtered.length + "</strong> of <strong>" + state.complaints.length + "</strong> complaints";
    }

    if (state.filtered.length === 0) {
      tbody.innerHTML = 
        '<tr class="empty-row">' +
        '  <td colspan="10" class="empty-state">' +
        '    <div class="empty-icon">' +
        '      <i class="fas fa-inbox"></i>' +
        '    </div>' +
        '    <p class="empty-text">No complaints found</p>' +
        '    <p class="empty-subtext">Try adjusting your filters or add a new complaint</p>' +
        '  </td>' +
        '</tr>';
      return;
    }

    state.filtered.forEach((c, i) => {
      const row = document.createElement("tr");
      
      const resolutionText = escapeHtml(c.cr650_resolution) || 'No resolution provided';
      const resolutionDisplay = escapeHtml(c.cr650_resolution) || '-';
      
      row.innerHTML = 
        '<td>' + (i + 1) + '</td>' +
        '<td>' + formatDate(c.cr650_complaintdate) + '</td>' +
        '<td>' + escapeHtml(c.cr650_customername) + '</td>' +
        '<td>' + escapeHtml(c.cr650_product) + '</td>' +
        '<td>' + natureLabel(c.cr650_complaintnature) + '</td>' +
        '<td>' + lobLabel(c.cr650_lob) + '</td>' +  // ⭐ NEW: LOB column
        '<td class="resolution-cell" title="' + resolutionText + '">' + resolutionDisplay + '</td>' +
        '<td>' + formatDate(c.cr650_closingdate) + '</td>' +
        '<td>' + statusBadge(c.cr650_status) + '</td>' +
        '<td class="actions-cell">' +
        '  <button onclick="ComplaintsManager.editComplaint(\'' + c.cr650_tsecomplaintsid + '\')" class="btn-edit" title="Edit">' +
        '    <i class="fas fa-edit"></i>' +
        '  </button>' +
        '  <button onclick="ComplaintsManager.deleteComplaint(\'' + c.cr650_tsecomplaintsid + '\')" class="btn-delete" title="Delete">' +
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
    const open = state.filtered.filter(c => c.cr650_status === 1).length;
    const closed = state.filtered.filter(c => c.cr650_status === 2).length;

    const openCount = document.getElementById("openComplaintsCount");
    const closedCount = document.getElementById("closedComplaintsCount");
    const totalCount = document.getElementById("totalComplaintsCount");

    if (openCount) openCount.textContent = open;
    if (closedCount) closedCount.textContent = closed;
    if (totalCount) totalCount.textContent = state.filtered.length;
  }

  /* ------------------------------
     Rendering – Chart
  ------------------------------ */
  function renderChart() {
    const canvas = document.getElementById("complaintsStatusChart");
    if (!canvas) return;

    if (state.chart) state.chart.destroy();

    const open = state.filtered.filter(c => c.cr650_status === 1).length;
    const closed = state.filtered.filter(c => c.cr650_status === 2).length;

    state.chart = new Chart(canvas.getContext("2d"), {
      type: "doughnut",
      data: {
        labels: ["Open", "Closed"],
        datasets: [{
          data: [open, closed],
          backgroundColor: [CONFIG.CHART_COLORS.open, CONFIG.CHART_COLORS.closed],
          borderWidth: 0
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          legend: {
            position: 'bottom'
          }
        }
      }
    });
  }

  /* ------------------------------
     Add / Edit / Save - NO FORM SUBMISSION
  ------------------------------ */
  function saveComplaint(event) {
    // CRITICAL: Prevent any form submission
    if (event) {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
    }

    const data = collectForm();
    if (!data) return false;

    // Disable save button to prevent double-click
    const saveBtn = document.getElementById("saveComplaintBtn");
    if (saveBtn) {
      saveBtn.disabled = true;
      saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';
    }

    if (state.editing) {
      updateComplaint(state.currentId, data);
    } else {
      createComplaint(data);
    }
    
    return false;
  }

  function createComplaint(data) {
    api({
      type: "POST",
      url: CONFIG.API_BASE + "/" + CONFIG.ENTITY_SET,
      data: JSON.stringify(data)
    })
    .done(function() {
      alert("Complaint created successfully.");
      loadData();
      closeModal();
    })
    .fail(function(e) {
      console.error("Create failed:", e);
      console.error("Response:", e.responseJSON);
      alert("Failed to create complaint. Please try again.");
    })
    .always(function() {
      resetSaveButton();
    });
  }

  function updateComplaint(id, data) {
    api({
      type: "PATCH",
      url: CONFIG.API_BASE + "/" + CONFIG.ENTITY_SET + "(" + id + ")",
      data: JSON.stringify(data)
    })
      .done(function() {
        alert("Complaint updated successfully.");
        loadData();
        closeModal();
      })
      .fail(function(e) {
        console.error("Update failed:", e);
        console.error("Response:", e.responseJSON);
        alert("Failed to update complaint. Please try again.");
      })
      .always(function() {
        resetSaveButton();
      });
  }

  function deleteComplaint(id) {
    if (!confirm("Are you sure you want to delete this complaint?")) return;

    api({
      type: "DELETE",
      url: CONFIG.API_BASE + "/" + CONFIG.ENTITY_SET + "(" + id + ")"
    })
    .done(function() {
      alert("Complaint deleted successfully.");
      loadData();
    })
    .fail(function(e) {
      console.error("Delete failed:", e);
      console.error("Response:", e.responseJSON);
      alert("Failed to delete complaint. Please try again.");
    })
    .always(function() {
      resetSaveButton();
    });
  }

  /* ------------------------------
     Filters
  ------------------------------ */
  function populateMonthFilter() {
    const select = document.getElementById('monthFilter');
    if (!select) return;

    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];

    const currentMonth = new Date().getMonth();

    select.innerHTML = '<option value="">All Months</option>';
    months.forEach((month, index) => {
      const option = document.createElement('option');
      option.value = index + 1;
      option.textContent = month;
      if (index === currentMonth) {
        option.selected = true;
      }
      select.appendChild(option);
    });
  }

  function applyFilters() {
    const monthFilter = document.getElementById("monthFilter");
    const statusFilter = document.getElementById("statusFilter");
    const searchInput = document.getElementById("searchInput");

    const lobFilter = document.getElementById("lobFilter");  // ⭐ NEW
    
    const month = monthFilter ? monthFilter.value : "";
    const status = statusFilter ? statusFilter.value : "";
    const lob = lobFilter ? lobFilter.value : "";  // ⭐ NEW
    const search = searchInput ? searchInput.value.toLowerCase() : "";
    state.filtered = state.complaints.filter(c => {
      // Month filter
      if (month) {
        const m = new Date(c.cr650_complaintdate).getMonth() + 1;
        if (m !== parseInt(month)) return false;
      }

      // Status filter
      if (status && c.cr650_status !== parseInt(status)) return false;
      if (lob && c.cr650_lob !== parseInt(lob)) return false;

      // Search filter
      if (search) {
        const text = (c.cr650_customername || '') + ' ' + 
                     (c.cr650_product || '') + ' ' + 
                     (c.cr650_resolution || '');
        if (text.toLowerCase().indexOf(search) === -1) return false;
      }

      return true;
    });

    renderTable();
    renderSummary();
    renderChart();
  }

  /* ------------------------------
     Excel Export (XLSX)
  ------------------------------ */
  function exportToExcel() {
    if (state.filtered.length === 0) {
      alert("No data to export");
      return;
    }

    // Check if XLSX library is loaded
    if (typeof XLSX === 'undefined') {
      alert('Excel export library not loaded. Please refresh the page.');
      return;
    }

    // Prepare data for export
    const exportData = state.filtered.map((c, i) => {
      return {
        'S.N.': i + 1,
        'Date': formatDate(c.cr650_complaintdate),
        'Customer': c.cr650_customername || '-',
        'Product': c.cr650_product || '-',
        'Nature': natureLabel(c.cr650_complaintnature),
        'Line of Business': lobLabelText(c.cr650_lob),
        'Resolution': c.cr650_resolution || '-',
        'Closing Date': formatDate(c.cr650_closingdate),
        'Status': c.cr650_status === 1 ? 'Open' : 'Closed',
        'Submitted By': c.cr650_submittedby || '-'
      };
    });

    // Create workbook and worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);

    // Set column widths
    ws['!cols'] = [
      { wch: 6 },   // S.N.
      { wch: 12 },  // Date
      { wch: 25 },  // Customer
      { wch: 20 },  // Product
      { wch: 20 },  // Nature
      { wch: 15 },  // Line of Business
      { wch: 40 },  // Resolution
      { wch: 12 },  // Closing Date
      { wch: 10 },  // Status
      { wch: 20 }   // Submitted By
    ];

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, "Complaints");

    // Generate filename with current date
    const filename = "Complaints_Export_" + new Date().toISOString().split('T')[0] + ".xlsx";

    // Download file
    XLSX.writeFile(wb, filename);
  }

  /* ------------------------------
     Modal Management
  ------------------------------ */
  function openAddForm() {
    state.editing = false;
    state.currentId = null;

    const form = document.getElementById("complaintForm");
    if (form) form.reset();

    const modalTitle = document.getElementById("modalTitle");
    if (modalTitle) {
      modalTitle.innerHTML = '<i class="fas fa-plus-circle"></i> Add New Complaint';
    }

    const modal = document.getElementById("complaintModal");
    if (modal) modal.style.display = "flex";

    // Set default date to today
    const dateInput = document.getElementById("complaintDate");
    if (dateInput && !dateInput.value) {
      dateInput.value = new Date().toISOString().split('T')[0];
    }

    toggleClosingDate();
  }

  function editComplaint(id) {
    state.editing = true;
    state.currentId = id;

    const c = state.complaints.find(x => x.cr650_tsecomplaintsid === id);
    if (!c) {
      alert("Complaint not found");
      return;
    }

    // Populate form fields
    document.getElementById("customerName").value = c.cr650_customername || '';
    document.getElementById("product").value = c.cr650_product || '';
    document.getElementById("complaintNature").value = c.cr650_complaintnature || '';
    document.getElementById("lob").value = c.cr650_lob || '';  // ⭐ NEW
    document.getElementById("status").value = c.cr650_status || 1;
    document.getElementById("resolution").value = c.cr650_resolution || '';
    document.getElementById("complaintDate").value = formatDateInput(c.cr650_complaintdate);

    if (c.cr650_closingdate) {
      document.getElementById("closingDate").value = formatDateInput(c.cr650_closingdate);
    }

    toggleClosingDate();

    const modalTitle = document.getElementById("modalTitle");
    if (modalTitle) {
      modalTitle.innerHTML = '<i class="fas fa-edit"></i> Edit Complaint';
    }

    const modal = document.getElementById("complaintModal");
    if (modal) modal.style.display = "flex";
  }

  function closeModal() {
    const modal = document.getElementById('complaintModal');
    if (modal) modal.style.display = 'none';

    // Reset state
    state.editing = false;
    state.currentId = null;

    const form = document.getElementById('complaintForm');
    if (form) form.reset();

    resetSaveButton();
  }

  function toggleClosingDate() {
    const statusSelect = document.getElementById('status');
    const closingDateGroup = document.getElementById('closingDateGroup');
    const closingDateInput = document.getElementById('closingDate');

    if (statusSelect && closingDateGroup) {
      const status = parseInt(statusSelect.value);
      closingDateGroup.style.display = status === 2 ? 'block' : 'none';

      if (status === 2) {
        closingDateInput.required = true;
        // Set default closing date to today if empty
        if (!closingDateInput.value) {
          closingDateInput.value = new Date().toISOString().split('T')[0];
        }
      } else {
        closingDateInput.required = false;
      }
    }
  }

  /* ------------------------------
     Utilities
  ------------------------------ */
  const NATURE_MAP = {
    1: "Leakage",
    2: "Wrong Product",
    3: "Product Performance",
    4: "Application Damage",
    5: "Other"
  };

  function collectForm() {
    const cust = document.getElementById("customerName").value.trim();
    const product = document.getElementById("product").value.trim();
    const natureValue = document.getElementById("complaintNature").value;
    const statusValue = document.getElementById("status").value;
    const complaintDate = document.getElementById("complaintDate").value;

    if (!cust || !product) {
      alert("Customer and product are required.");
      return null;
    }

    if (!natureValue) {
      alert("Please select the nature of complaint.");
      return null;
    }

    const lobValue = document.getElementById("lob").value;
    
    if (!lobValue) {
      alert("Please select Line of Business.");
      return null;
    }

    if (!complaintDate) {
      alert("Complaint date is required.");
      return null;
  }

    const data = {
      cr650_name: cust + " - " + product,
      cr650_customername: cust,
      cr650_product: product,
      cr650_complaintdate: toIsoDate(complaintDate),
      cr650_complaintnature: natureValue,
      cr650_lob: parseInt(lobValue),
      cr650_status: parseInt(statusValue),
      cr650_resolution: document.getElementById("resolution").value || '',
    };

    // 🔑 ONLY set submittedby on CREATE
    if (!state.editing) {
      data.cr650_submittedby = window.PORTAL_USER_NAME || 'Unknown User';
    }
        

    // Add closing date only if status is Closed
    if (parseInt(statusValue) === 2) {
      const closingDate = document.getElementById("closingDate").value;
      if (!closingDate) {
        alert("Closing date is required for closed complaints.");
        return null;
      }
      data.cr650_closingdate = toIsoDate(closingDate);
    }

    console.log("Complaint payload:", JSON.stringify(data, null, 2));

    return data;
  }

  function formatDate(d) {
    if (!d) return "-";
    return new Date(d).toLocaleDateString("en-US");
  }

 function toIsoDate(d) {
  return d || null;  // Just return the date string as-is!
}


  function formatDateInput(d) {
    if (!d) return "";
    return new Date(d).toISOString().split('T')[0];
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

  function escapeHtml(text) {
    if (!text) return "";
    var map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, function(m) { return map[m]; });
  }

  function natureLabel(v) {
    return v || "-";
  }

  function statusBadge(v) {
    return v === 1
      ? '<span class="status open">Open</span>'
      : '<span class="status closed">Closed</span>';
  }

  function refreshData() {
    console.log("Refreshing data...");
    loadData();
  }

  function resetSaveButton() {
    const saveBtn = document.getElementById("saveComplaintBtn");
    if (saveBtn) {
      saveBtn.disabled = false;
      saveBtn.innerHTML = '<i class="fas fa-save"></i> Save Complaint';
    }
  }

  /* ------------------------------
     Public API
  ------------------------------ */
  return {
    init: init,
    loadData: loadData,
    applyFilters: applyFilters,
    saveComplaint: saveComplaint,
    editComplaint: editComplaint,
    openAddForm: openAddForm,
    closeModal: closeModal,
    toggleClosingDate: toggleClosingDate,
    refreshData: refreshData,
    deleteComplaint: deleteComplaint,
    exportToExcel: exportToExcel
  };

})();

/* ------------------------------
   Autostart
------------------------------ */
if (document.readyState === 'loading') {
  document.addEventListener("DOMContentLoaded", function() {
    ComplaintsManager.init();
  });
} else {
  ComplaintsManager.init();
}