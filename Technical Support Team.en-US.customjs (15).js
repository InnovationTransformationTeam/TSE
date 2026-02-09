/**
 * TSR Power Pages – Production-Ready v3.5 (CORRECTED)
 * All training field bugs fixed
 */

(function (window, document, $) {
  'use strict';
  if (!$) {
    console.error('TSR App: jQuery is required.');
    return;
  }

  /* ==========================
     Web API transport (shell token + session validation)
  ========================== */
  (function (webapi, $) {
    if (typeof window.validateLoginSession !== 'function') {
      window.validateLoginSession = function (data, textStatus, jqXHR, onOk) {
        if (typeof onOk === 'function') onOk(data, textStatus, jqXHR);
      };
    }

    function safeAjax(ajaxOptions) {
      var d = $.Deferred();

      if (!(window.shell && typeof shell.getTokenDeferred === 'function')) {
        d.reject({ status: 401, statusText: 'shell.getTokenDeferred unavailable on this page' });
        return d.promise();
      }

      shell.getTokenDeferred().done(function (token) {
        ajaxOptions = ajaxOptions || {};
        ajaxOptions.headers = $.extend({
          'Accept': 'application/json',
          'Content-Type': 'application/json; charset=utf-8',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          '__RequestVerificationToken': token
        }, ajaxOptions.headers || {});

        $.ajax(ajaxOptions)
          .done(function (data, textStatus, jqXHR) {
            validateLoginSession(data, textStatus, jqXHR, function () {
              d.resolve(data, textStatus, jqXHR);
            });
          })
          .fail(d.reject);
      }).fail(function () { d.rejectWith(this, arguments); });

      return d.promise();
    }

    webapi.safeAjax = safeAjax;
  })(window.webapi = window.webapi || {}, jQuery);

  /* ==========================
     CONFIGURATION
  ========================== */
  const CONFIG = {
    API_BASE: '/_api',
    ENTITY_SET: 'cr650_technicalsupportreportses',
    ENTITY_ID_FIELD: 'cr650_technicalsupportreportsid',
    CONTACT_LOOKUP: 'cr650_contact',
    MAX_RETRIES: 3,
    DEBOUNCE_DELAY: 300,
    USE_UPSERT: false,
    STATUS: { Draft: 1, Submitted: 2 },
    LOB: { Petromin: 1, Gulf: 2, Commercial: 3 },
    FIELD_MAP: {
      customerCourtesyVisits: 'cr650_customercourtesyvisits',
      productEquivalent: 'cr650_productequivalent',
      internalTraining: 'cr650_internaltraining',
      onsiteTraining: 'cr650_onsitetraining',
      businessOpportunity: 'cr650_businessopportunity',
      complaintHandled: 'cr650_complainthandled',
      fieldTrial: 'cr650_fieldtrial',
      lubricationSurvey: 'cr650_lubricationsurvey',
      dvrEndorsement: 'cr650_dvrendorsement',
      totalActivities: 'cr650_totalactivities'
    },
    ALT_KEY: {
      contact: 'cr650_contact',
      month: 'cr650_reportmonth',
      lob: 'cr650_lineofbusiness'
    }
  };

  /* ==========================
     PORTAL USER CONTEXT
  ========================== */
  (function setPortalUser() {
    function cleanGuid(g) {
      if (!g) return null;
      const s = String(g).trim();
      // Reject empty, Liquid template literals, and placeholder values
      if (!s || s.includes('{{') || s === 'null' || s === 'undefined') return null;
      return s.replace(/[{}]/g, '');
    }

    function cleanLiquid(val) {
      if (!val) return null;
      const s = String(val).trim();
      if (!s || s.includes('{{') || s === 'null' || s === 'undefined') return null;
      return s;
    }

    // 1. Contact ID — check multiple sources in priority order
    const contactId =
      cleanGuid(window.PORTAL_CONTACT_ID) ||
      cleanGuid(document.getElementById('liquid-contact-id')?.value) ||
      cleanGuid(document.querySelector('meta[name="microsoft-powerpages-userid"]')?.content) ||
      cleanGuid(document.querySelector('input[name="ContactId"]')?.value);

    window.PORTAL_CONTACT_ID = contactId;

    // 2. Username — check Liquid hidden input, then DOM .username element
    const liquidName = cleanLiquid(document.getElementById('liquid-user-name')?.value);
    const domName = document.querySelector('.username')?.textContent?.trim();

    if (!window.PORTAL_USER_NAME || window.PORTAL_USER_NAME === 'Portal User') {
      window.PORTAL_USER_NAME = liquidName || domName || window.PORTAL_USER_NAME || 'Portal User';
    }

    // 3. Email
    const liquidEmail = cleanLiquid(document.getElementById('liquid-user-email')?.value);
    if (!window.PORTAL_USER_EMAIL) {
      window.PORTAL_USER_EMAIL = liquidEmail || null;
    }

    if (!window.PORTAL_CONTACT_ID) {
      console.log('TSR: Contact ID not available — records will use username-based filtering');
    } else {
      console.log('Portal contact:', window.PORTAL_CONTACT_ID);
    }

    console.log('Portal user:', window.PORTAL_USER_NAME);
  })();

  /* ==========================
     UTILS
  ========================== */
  const Utils = {
    cleanGuid: (g) => (g ? String(g).replace(/[{}]/g, '') : null),
    yyyyMmToDateOnly: (m) => (/^\d{4}-\d{2}$/.test(m) ? `${m}-01` : null),
    nzInt: (v) => {
      const n = parseInt(v, 10);
      return Number.isFinite(n) && n >= 0 ? n : 0;
    },
    // FIXED: computeTotal now uses split training fields
    computeTotal: (d) => [
      'customerCourtesyVisits', 'productEquivalent',
      'internalTraining', 'onsiteTraining',
      'businessOpportunity', 'complaintHandled', 'fieldTrial',
      'lubricationSurvey', 'dvrEndorsement'
    ].reduce((s, f) => s + Utils.nzInt(d[f]), 0),
    nowIso: () => new Date().toISOString(),
    currentMonth: () => new Date().toISOString().slice(0, 7),
    debounce(fn, ms) {
      let t; return function () { clearTimeout(t); t = setTimeout(() => fn.apply(this, arguments), ms); };
    },
    formatMonth(yyyyMm) {
      if (!yyyyMm) return 'Not selected';
      const dt = new Date(`${yyyyMm}-01T00:00:00`);
      return dt.toLocaleDateString('en', { month: 'long', year: 'numeric' });
    },
    // ADDED: Safe OData string escaping for user names with apostrophes
    escapeODataString(str) {
      if (!str) return str;
      return str.replace(/'/g, "''");
    }
  };

  function guidLit(id) { return `guid'${id}'`; }

  /* ==========================
     NOTIFY
  ========================== */
  const Notify = {
    show(msg, type = 'info', title = '', duration = 5000) {
      const container = document.getElementById('notificationContainer');
      if (!container) { alert((title ? title + '\n' : '') + msg.replace(/<[^>]*>?/gm, '')); return; }
      const el = document.createElement('div');
      el.className = `notification ${type}`;
      const icons = { success: 'fas fa-check-circle', error: 'fas fa-exclamation-circle', warning: 'fas fa-exclamation-triangle', info: 'fas fa-info-circle' };
      el.innerHTML = `
        <div class="notification-icon"><i class="${icons[type] || icons.info}"></i></div>
        <div class="notification-content">
          ${title ? `<div class="notification-title">${title}</div>` : ''}
          <div class="notification-message">${msg}</div>
        </div>
        <div class="notification-close" aria-label="Close">✕</div>`;
      el.querySelector('.notification-close').addEventListener('click', () => el.remove());
      container.appendChild(el);
      if (duration > 0) setTimeout(() => el.remove(), duration);
    },
    success(m, t) { this.show(m, 'success', t); },
    error(m, t) { this.show(m, 'error', t); },
    warning(m, t) { this.show(m, 'warning', t); },
    info(m, t) { this.show(m, 'info', t); }
  };

  /* ==========================
     APP STATE
  ========================== */
  const AppState = {
    petromin: { isDraftComplete: false, data: {} },
    gulf: { isDraftComplete: false, data: {} },
    commercial: { isDraftComplete: false, data: {} },
    officiallySubmitted: false,
    currentReportingPeriod: null,
    currentStep: 'petromin',
    offlineQueue: [],
    totalSubmittedReports: 0
  };

  /* ==========================
     DATAVERSE API (with user-specific filtering)
  ========================== */
  const DataverseAPI = {
    buildPayload(record) {
      const payload = {
        cr650_reportmonth: Utils.yyyyMmToDateOnly(record.month),
        cr650_lineofbusiness: record.lob,
        cr650_status: record.status || CONFIG.STATUS.Submitted,
        cr650_submissiondate: record.status === CONFIG.STATUS.Submitted ? Utils.nowIso() : null,
        cr650_submittedby: window.PORTAL_USER_NAME || 'Unknown User'
      };

      const contactId = Utils.cleanGuid(window.PORTAL_CONTACT_ID);
      if (contactId) {
        payload[`${CONFIG.CONTACT_LOOKUP}@odata.bind`] = `/contacts(${contactId})`;
      }

      Object.keys(CONFIG.FIELD_MAP).forEach((uiField) => {
        if (uiField === 'totalActivities') return;
        payload[CONFIG.FIELD_MAP[uiField]] = Utils.nzInt(record[uiField]);
      });
      payload[CONFIG.FIELD_MAP.totalActivities] = Utils.computeTotal(record);

      return payload;
    },

    create(record) {
      return webapi.safeAjax({
        type: 'POST',
        url: `${CONFIG.API_BASE}/${CONFIG.ENTITY_SET}`,
        data: JSON.stringify(this.buildPayload(record))
      });
    },

    update(id, record) {
      return webapi.safeAjax({
        type: 'PATCH',
        url: `${CONFIG.API_BASE}/${CONFIG.ENTITY_SET}(${id})`,
        data: JSON.stringify(this.buildPayload(record))
      });
    },

    findExisting(month, lob) {
      const dateOnly = Utils.yyyyMmToDateOnly(month);
      if (!dateOnly) {
        console.warn('findExisting: invalid month format', month);
        return $.Deferred().resolve(null).promise();
      }
      const select = [CONFIG.ENTITY_ID_FIELD, 'cr650_reportmonth', 'cr650_lineofbusiness'].join(',');

      let filter = `cr650_reportmonth eq ${dateOnly} and cr650_lineofbusiness eq ${lob}`;

      // FIXED: Escape username for OData query (handles apostrophes)
      const userName = window.PORTAL_USER_NAME;
      if (userName && userName !== 'Portal User' && userName !== 'Unknown User') {
        filter += ` and cr650_submittedby eq '${Utils.escapeODataString(userName)}'`;
      }

      const url = `${CONFIG.API_BASE}/${CONFIG.ENTITY_SET}?$select=${select}&$top=1&$filter=${encodeURIComponent(filter)}`;
      console.log('Find existing query:', url);

      return webapi.safeAjax({ type: 'GET', url }).then(r => {
        const items = r?.value || [];
        return items.length ? items[0][CONFIG.ENTITY_ID_FIELD] : null;
      });
    },

    submitOrUpdate(record) {
      return this.findExisting(record.month, record.lob).then((id) => {
        return id ? this.update(id, record) : this.create(record);
      });
    },

    getMyReports(top = 500, monthFilter = null) {
      const select = [
        CONFIG.ENTITY_ID_FIELD, 'cr650_reportmonth', 'cr650_lineofbusiness', 'cr650_status', 'cr650_submissiondate',
        'cr650_customercourtesyvisits', 'cr650_productequivalent',
        'cr650_internaltraining', 'cr650_onsitetraining',  // FIXED: Both training fields
        'cr650_businessopportunity', 'cr650_complainthandled', 'cr650_fieldtrial',
        'cr650_lubricationsurvey', 'cr650_dvrendorsement', 'cr650_totalactivities',
        'cr650_submittedby'
      ].join(',');

      let url = `${CONFIG.API_BASE}/${CONFIG.ENTITY_SET}?$select=${select}&$orderby=cr650_reportmonth desc&$top=${top}`;

      const filters = [];

      // FIXED: Escape username for OData query
      const userName = window.PORTAL_USER_NAME;
      if (userName && userName !== 'Portal User' && userName !== 'Unknown User') {
        filters.push(`cr650_submittedby eq '${Utils.escapeODataString(userName)}'`);
        console.log('Filtering dashboard by submittedby:', userName);
      }

      if (monthFilter && monthFilter !== '__all__') {
        filters.push(`cr650_reportmonth eq ${Utils.yyyyMmToDateOnly(monthFilter)}`);
      }

      if (filters.length) {
        url += `&$filter=${encodeURIComponent(filters.join(' and '))}`;
      }

      console.log('Dashboard query URL:', url);
      return webapi.safeAjax({ type: 'GET', url });
    }
  };

  /* ==========================
     AUTO-CALC MANAGER
  ========================== */
  const AutoCalcManager = {
    // Track pending auto-calc requests per form type
    _pending: { petromin: 0, gulf: 0, commercial: 0 },

    isLoading(formType) {
      return this._pending[formType] > 0;
    },

    isAnyLoading() {
      return Object.values(this._pending).some(v => v > 0);
    },

    fillComplaints(formType, month) {
      const lobMap = {
        petromin: CONFIG.LOB.Petromin,
        gulf: CONFIG.LOB.Gulf,
        commercial: CONFIG.LOB.Commercial
      };
      const lob = lobMap[formType];
      const input = document.getElementById(formType + 'ComplaintHandled');

      if (!input || !month) return $.Deferred().resolve().promise();

      this._pending[formType]++;
      input.placeholder = 'Loading...';

      // Calculate start and end of month
      const startDate = Utils.yyyyMmToDateOnly(month);
      const year = parseInt(month.split('-')[0]);
      const monthNum = parseInt(month.split('-')[1]);
      const lastDay = new Date(year, monthNum, 0).getDate();
      const endDate = `${month}-${lastDay.toString().padStart(2, '0')}`;

      let filter = `cr650_lob eq ${lob} and cr650_complaintdate ge ${startDate} and cr650_complaintdate le ${endDate}`;

      // Add user filter if available
      const userName = window.PORTAL_USER_NAME;
      if (userName && userName !== 'Portal User' && userName !== 'Unknown User') {
        filter += ` and cr650_submittedby eq '${Utils.escapeODataString(userName)}'`;
      }

      const url = `${CONFIG.API_BASE}/cr650_tsecomplaintses?$select=cr650_status&$filter=${encodeURIComponent(filter)}&$top=1000`;

      return webapi.safeAjax({ type: 'GET', url })
        .done((res) => {
          const data = res.value || [];
          const total = data.length;
          const open = data.filter(c => c.cr650_status === 1).length;
          const closed = data.filter(c => c.cr650_status === 2).length;

          input.value = total;
          input.readOnly = true;
          input.style.backgroundColor = '#F9FAFB';
          input.style.fontWeight = '600';

          this.addStatusBadge(input, 'complaints', { total, open, closed });
        })
        .fail((err) => {
          console.error(`Failed to load complaints for ${formType}:`, err);
          input.value = 0;
          input.placeholder = '0';
        })
        .always(() => {
          this._pending[formType]--;
        });
    },

    fillFieldTrials(formType, month) {
      const lobMap = {
        petromin: CONFIG.LOB.Petromin,
        gulf: CONFIG.LOB.Gulf,
        commercial: CONFIG.LOB.Commercial
      };
      const lob = lobMap[formType];
      const input = document.getElementById(formType + 'FieldTrial');

      if (!input || !month) return $.Deferred().resolve().promise();

      this._pending[formType]++;
      input.placeholder = 'Loading...';

      // Calculate start and end of month
      const startDate = Utils.yyyyMmToDateOnly(month);
      const year = parseInt(month.split('-')[0]);
      const monthNum = parseInt(month.split('-')[1]);
      const lastDay = new Date(year, monthNum, 0).getDate();
      const endDate = `${month}-${lastDay.toString().padStart(2, '0')}`;

      let filter = `cr650_lob eq ${lob} and cr650_startdate ge ${startDate} and cr650_startdate le ${endDate}`;

      // Add user filter if available
      const userName = window.PORTAL_USER_NAME;
      if (userName && userName !== 'Portal User' && userName !== 'Unknown User') {
        filter += ` and cr650_submittedby eq '${Utils.escapeODataString(userName)}'`;
      }

      const url = `${CONFIG.API_BASE}/cr650_tsefieldtrialses?$select=cr650_status&$filter=${encodeURIComponent(filter)}&$top=1000`;

      return webapi.safeAjax({ type: 'GET', url })
        .done((res) => {
          const data = res.value || [];
          const total = data.length;
          const planned = data.filter(t => t.cr650_status === 1).length;
          const ongoing = data.filter(t => t.cr650_status === 2).length;
          const completed = data.filter(t => t.cr650_status === 3).length;

          input.value = total;
          input.readOnly = true;
          input.style.backgroundColor = '#F9FAFB';
          input.style.fontWeight = '600';

          this.addStatusBadge(input, 'trials', { total, planned, ongoing, completed });
        })
        .fail((err) => {
          console.error(`Failed to load trials for ${formType}:`, err);
          input.value = 0;
          input.placeholder = '0';
        })
        .always(() => {
          this._pending[formType]--;
        });
    },

    addStatusBadge(input, type, data) {
      const existing = input.parentElement.querySelector('.auto-calc-status');
      if (existing) existing.remove();

      const badge = document.createElement('div');
      badge.className = 'auto-calc-status';
      badge.style.cssText = 'margin-top: 8px; font-size: 12px; color: #64748b;';

      if (type === 'complaints') {
        badge.innerHTML =
          '<span style="color: #F59E0B;">●</span> Open: ' + data.open + ' | ' +
          '<span style="color: #10B981;">●</span> Closed: ' + data.closed + ' | ' +
          '<a href="/Technical-Support-Team/Complaints-Management/" style="color: #1B8B47; text-decoration: none; margin-left: 8px; font-weight: 500;">View Details →</a>';
      } else {
        badge.innerHTML =
          '<span style="color: #3B82F6;">●</span> Planned: ' + data.planned + ' | ' +
          '<span style="color: #F59E0B;">●</span> On-going: ' + data.ongoing + ' | ' +
          '<span style="color: #10B981;">●</span> Done: ' + data.completed + ' | ' +
          '<a href="/Technical-Support-Team/Field-Trials-Management" style="color: #1B8B47; text-decoration: none; margin-left: 8px; font-weight: 500;">View Details →</a>';
      }

      input.parentElement.insertBefore(badge, input.nextSibling);
    },

    autoFillForMonth(formType, month) {
      if (!month) return;
      console.log(`Auto-filling ${formType} for ${month}`);
      this.fillComplaints(formType, month);
      this.fillFieldTrials(formType, month);
    }
  };

  /* ==========================
     FORM MANAGER
  ========================== */
  const FormManager = {
    collectData(formType) {
      const prefix = formType === 'petromin' ? '#petromin' : formType === 'gulf' ? '#gulf' : '#commercial';
      // FIXED: Using split training fields
      const fields = [
        'month', 'customerCourtesyVisits', 'productEquivalent',
        'internalTraining', 'onsiteTraining',
        'businessOpportunity', 'complaintHandled', 'fieldTrial', 'lubricationSurvey', 'dvrEndorsement'
      ];
      const data = {};
      for (const f of fields) {
        const id = `${prefix}${f.charAt(0).toUpperCase()}${f.slice(1)}`;
        const el = document.querySelector(id);
        data[f] = el?.value || '';
      }
      return data;
    },

    validate(formType, data) {
      if (!data?.month) {
        Notify.error(`Select the reporting month for ${formType}.`, 'Missing Month');
        return false;
      }
      // FIXED: Using split training fields
      const numeric = [
        'customerCourtesyVisits', 'productEquivalent',
        'internalTraining', 'onsiteTraining',
        'businessOpportunity', 'complaintHandled', 'fieldTrial', 'lubricationSurvey', 'dvrEndorsement'
      ];

      for (const f of numeric) {
        const val = data[f];
        if (val === '' || val === null || val === undefined || isNaN(Number(val)) || Number(val) < 0) {
          const label = f.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase());
          Notify.error(`Enter a valid non-negative number for "${label}". Zero is allowed.`, 'Invalid Value');
          const inputId = `${formType}${f.charAt(0).toUpperCase()}${f.slice(1)}`;
          const input = document.getElementById(inputId);
          if (input) { input.classList.add('error'); input.focus(); }
          return false;
        }
      }
      return true;
    },

    // FIXED: checkCompletion now uses split training fields
    checkCompletion(formType) {
      const d = this.collectData(formType);
      if (!d) return false;
      const req = [
        'month', 'customerCourtesyVisits', 'productEquivalent',
        'internalTraining', 'onsiteTraining',  // FIXED: Was 'customerTraining'
        'businessOpportunity', 'complaintHandled', 'fieldTrial', 'lubricationSurvey', 'dvrEndorsement'
      ];
      return req.every(k => {
        const value = d[k];
        if (k === 'month') return String(value).trim() !== '';
        return value !== '' && value !== null && value !== undefined && !isNaN(Number(value)) && Number(value) >= 0;
      });
    },

    saveDraft(formType) {
      // Block save if auto-calc fields are still loading (would save stale/zero values)
      if (AutoCalcManager.isLoading(formType)) {
        Notify.warning('Please wait — complaints and field trials are still loading.', 'Auto-Calc In Progress');
        return;
      }

      const spinner = document.getElementById(formType + 'Spinner');
      spinner?.classList.remove('hidden');

      const data = this.collectData(formType);
      if (!this.validate(formType, data)) {
        spinner?.classList.add('hidden');
        return;
      }

      const lobMap = { petromin: CONFIG.LOB.Petromin, gulf: CONFIG.LOB.Gulf, commercial: CONFIG.LOB.Commercial };
      data.lob = lobMap[formType];
      data.status = CONFIG.STATUS.Draft;

      DataverseAPI.submitOrUpdate(data)
        .done(() => {
          AppState[formType].data = data;
          AppState[formType].isDraftComplete = true;
          WorkflowManager.updateProgress();
          Notify.success(`${formType.charAt(0).toUpperCase() + formType.slice(1)} draft saved to Dataverse.`, 'Draft Saved');

          if (formType === 'petromin' && WorkflowManager.canNavigateTo('gulf')) {
            setTimeout(() => WorkflowManager.navigateTo('gulf'), 600);
          } else if (formType === 'gulf' && WorkflowManager.canNavigateTo('commercial')) {
            setTimeout(() => WorkflowManager.navigateTo('commercial'), 600);
          } else if (formType === 'commercial' && WorkflowManager.canNavigateTo('review')) {
            setTimeout(() => WorkflowManager.navigateTo('review'), 600);
          }
        })
        .fail((err) => {
          console.error('Draft save failed:', err);
          Notify.error('Failed to save draft to Dataverse.', 'Save Failed');
        })
        .always(() => {
          spinner?.classList.add('hidden');
        });
    },

    loadDrafts() {
      const checks = [];

      ['petromin', 'gulf', 'commercial'].forEach(ft => {
        const lobMap = { petromin: CONFIG.LOB.Petromin, gulf: CONFIG.LOB.Gulf, commercial: CONFIG.LOB.Commercial };
        const lob = lobMap[ft];

        const monthInput = document.querySelector(`#${ft}Month`);
        const selectedMonth = (monthInput && monthInput.value) ? monthInput.value : Utils.currentMonth();

        const promise = DataverseAPI.findExisting(selectedMonth, lob)
          .then(recordId => {
            if (!recordId) return null;

            return webapi.safeAjax({
              type: 'GET',
              url: `${CONFIG.API_BASE}/${CONFIG.ENTITY_SET}(${recordId})`
            });
          })
          .done(record => {
            if (!record) return;

            const prefix = ft === 'petromin' ? '#petromin' : ft === 'gulf' ? '#gulf' : '#commercial';
            const fieldMap = {
              cr650_reportmonth: 'month',
              cr650_customercourtesyvisits: 'customerCourtesyVisits',
              cr650_productequivalent: 'productEquivalent',
              cr650_internaltraining: 'internalTraining',
              cr650_onsitetraining: 'onsiteTraining',
              cr650_businessopportunity: 'businessOpportunity',
              cr650_complainthandled: 'complaintHandled',
              cr650_fieldtrial: 'fieldTrial',
              cr650_lubricationsurvey: 'lubricationSurvey',
              cr650_dvrendorsement: 'dvrEndorsement'
            };

            // Auto-calculated fields should only be populated by AutoCalcManager (live from source tables)
            const autoCalcFields = ['cr650_complainthandled', 'cr650_fieldtrial'];

            Object.keys(fieldMap).forEach(dataverseField => {
              const uiField = fieldMap[dataverseField];
              const value = record[dataverseField];

              if (value !== null && value !== undefined) {
                let displayValue = value;
                if (dataverseField === 'cr650_reportmonth') {
                  displayValue = value.slice(0, 7);
                }

                const id = `${prefix}${uiField.charAt(0).toUpperCase()}${uiField.slice(1)}`;
                const el = document.querySelector(id);
                if (el) {
                  // Skip auto-calc fields unless the record is already submitted (read-only view)
                  if (autoCalcFields.includes(dataverseField) && record.cr650_status !== CONFIG.STATUS.Submitted) {
                    return;
                  }
                  el.value = displayValue;
                  if (record.cr650_status === CONFIG.STATUS.Submitted) {
                    el.readOnly = true;
                    el.style.backgroundColor = '#F3F4F6';
                    el.style.cursor = 'not-allowed';
                  }
                }
              }
            });

            // ✅ CRITICAL FIX: Only mark as draft complete if NOT submitted
            const isComplete = this.checkCompletion(ft);
            const isSubmitted = record.cr650_status === CONFIG.STATUS.Submitted;

            AppState[ft].isDraftComplete = isComplete && !isSubmitted;

            console.log(`${ft}: Complete=${isComplete}, Submitted=${isSubmitted}, DraftComplete=${AppState[ft].isDraftComplete}`);

            return record;
          })
          .fail(err => {
            console.log(`No draft found for ${ft}:`, err);
            return null;
          });

        checks.push(promise);
      });

      // Wait for all three to complete, then check submission status
      Promise.all(checks).then(records => {
        const submittedCount = records.filter(r => r?.cr650_status === CONFIG.STATUS.Submitted).length;

        if (submittedCount === 3) {
          AppState.officiallySubmitted = true;
          WorkflowManager.updateProgress();
          const displayMonth = document.querySelector('#petrominMonth')?.value || Utils.currentMonth();
          Notify.info(`Reports for ${Utils.formatMonth(displayMonth)} are already submitted.`, 'Already Submitted');
        }
      });
    },

    showSubmissionStatus(formType, month) {
      const lobMap = {
        petromin: CONFIG.LOB.Petromin,
        gulf: CONFIG.LOB.Gulf,
        commercial: CONFIG.LOB.Commercial
      };
      const lob = lobMap[formType];

      DataverseAPI.findExisting(month, lob)
        .then(recordId => {
          if (!recordId) return;

          return webapi.safeAjax({
            type: 'GET',
            url: `${CONFIG.API_BASE}/${CONFIG.ENTITY_SET}(${recordId})?$select=cr650_status,cr650_submissiondate`
          });
        })
        .done(record => {
          if (!record || record.cr650_status !== CONFIG.STATUS.Submitted) return;

          const formEl = document.getElementById(`${formType}Content`);
          const existingAlert = formEl?.querySelector('.submission-alert');
          if (existingAlert) existingAlert.remove();

          const alert = document.createElement('div');
          alert.className = 'submission-alert';
          alert.style.cssText = `
            background: #FEF3C7;
            border-left: 4px solid #F59E0B;
            padding: 16px;
            margin-bottom: 24px;
            border-radius: 8px;
          `;
          alert.innerHTML = `
            <div style="display: flex; align-items: center; gap: 12px;">
              <i class="fas fa-lock" style="color: #F59E0B; font-size: 20px;"></i>
              <div>
                <div style="font-weight: 600; color: #92400E; margin-bottom: 4px;">
                  This report has been submitted
                </div>
                <div style="font-size: 14px; color: #78350F;">
                  Submitted on ${new Date(record.cr650_submissiondate).toLocaleDateString()} - Fields are read-only
                </div>
              </div>
            </div>
          `;

          const panelHeader = formEl?.querySelector('.panel-header');
          if (panelHeader?.nextSibling) {
            formEl.insertBefore(alert, panelHeader.nextSibling);
          }
        })
        .catch(err => {
          console.log(`Could not check submission status for ${formType}:`, err);
        });
    },

    lockAllFields() {
      ['petromin', 'gulf', 'commercial'].forEach(formType => {
        const prefix = formType === 'petromin' ? '#petromin' : formType === 'gulf' ? '#gulf' : '#commercial';
        const fields = [
          'month', 'customerCourtesyVisits', 'productEquivalent',
          'internalTraining', 'onsiteTraining',
          'businessOpportunity', 'complaintHandled', 'fieldTrial', 'lubricationSurvey', 'dvrEndorsement'
        ];

        fields.forEach(f => {
          const id = `${prefix}${f.charAt(0).toUpperCase()}${f.slice(1)}`;
          const el = document.querySelector(id);
          if (el) {
            el.readOnly = true;
            el.style.backgroundColor = '#F3F4F6';
            el.style.cursor = 'not-allowed';
          }
        });

        console.log(`${formType} fields locked`);
      });
    },



    reset() {
      if (!confirm('Reset all forms and local data?')) return;

      AppState.petromin = { isDraftComplete: false, data: {} };
      AppState.gulf = { isDraftComplete: false, data: {} };
      AppState.commercial = { isDraftComplete: false, data: {} };
      AppState.officiallySubmitted = false;
      AppState.currentReportingPeriod = null;

      document.querySelectorAll('form').forEach(f => f.reset());

      localStorage.removeItem('tsr_offline_queue');

      WorkflowManager.setDefaultMonth();
      WorkflowManager.updateProgress();
      WorkflowManager.navigateTo('petromin');
      Notify.info('Forms reset. Start a new reporting period.', 'Reset');
    }
  };

  /* ==========================
     WORKFLOW MANAGER
  ========================== */
  const WorkflowManager = {
    stepOrder: ['petromin', 'gulf', 'commercial', 'review', 'dashboard'],

    setDefaultMonth() {
      const current = Utils.currentMonth();
      document.querySelectorAll('input[type="month"]').forEach(i => { if (!i.value) i.value = current; });
      AppState.currentReportingPeriod = current;
    },

    canNavigateTo(step) {
      switch (step) {
        case 'petromin': return true;
        case 'gulf': return AppState.petromin.isDraftComplete;
        case 'commercial': return AppState.petromin.isDraftComplete && AppState.gulf.isDraftComplete;
        case 'review': return AppState.petromin.isDraftComplete && AppState.gulf.isDraftComplete && AppState.commercial.isDraftComplete;
        case 'dashboard': return true;
        default: return false;
      }
    },

    navigateTo(step) {
      if (!this.canNavigateTo(step)) {
        const msg = step === 'gulf' ? 'Complete Petromin first.' : step === 'commercial' ? 'Complete Petromin and Gulf first.' : 'Step not available yet.';
        Notify.warning(msg, 'Navigation Restricted');
        return;
      }
      document.querySelectorAll('.report-panel').forEach(p => p.classList.remove('active'));
      document.getElementById(`${step}Content`)?.classList.add('active');

      document.querySelectorAll('.workflow-step').forEach(s => s.classList.remove('active'));
      document.querySelector(`[data-step="${step}"]`)?.classList.add('active');

      AppState.currentStep = step;
      this.updateNavigationButtons();

      if (step === 'petromin' || step === 'gulf' || step === 'commercial') {
        const monthInput = document.querySelector(`#${step}Month`);
        if (monthInput?.value) {
          FormManager.showSubmissionStatus(step, monthInput.value);
        }
      }


      if (step === 'review') {
        this.updateReviewSummary();

        // Force-check button status after brief delay
        setTimeout(() => {
          const submitBtn = document.getElementById('submitAllButton');
          if (submitBtn) {
            let allComplete = true;
            ['petromin', 'gulf', 'commercial'].forEach(ft => {
              if (!AppState[ft].isDraftComplete) {
                allComplete = false;
              }
            });

            if (AppState.officiallySubmitted) {
              submitBtn.style.display = 'none'; // Hide if already submitted
            } else {
              submitBtn.style.display = 'inline-flex';
              submitBtn.disabled = !allComplete;
              if (allComplete) {
                console.log('✓ Submit button ready to click');
              }
            }
          }
        }, 100);
      }
      if (step === 'dashboard') DashboardManager.load();
    },

    navigateNext() {
      const i = this.stepOrder.indexOf(AppState.currentStep);
      if (i < this.stepOrder.length - 1) this.navigateTo(this.stepOrder[i + 1]);
    },

    navigateBack() {
      const i = this.stepOrder.indexOf(AppState.currentStep);
      if (i > 0) this.navigateTo(this.stepOrder[i - 1]);
    },

    updateNavigationButtons() {
      const back = document.getElementById('backButton');
      const next = document.getElementById('nextButton');
      if (back) back.disabled = AppState.currentStep === 'petromin';

      if (next) {
        const i = this.stepOrder.indexOf(AppState.currentStep);
        if (i === this.stepOrder.length - 1) {
          next.style.display = 'none';
        } else {
          next.style.display = 'inline-flex';
          next.disabled = !this.canNavigateTo(this.stepOrder[i + 1]);
        }
      }
    },

    updateProgress() {
      AppState.petromin.isDraftComplete = FormManager.checkCompletion('petromin');
      AppState.gulf.isDraftComplete = FormManager.checkCompletion('gulf');
      AppState.commercial.isDraftComplete = FormManager.checkCompletion('commercial');

      this.updateWorkflowSteps();

      let completeCount = 0;
      if (AppState.petromin.isDraftComplete) completeCount++;
      if (AppState.gulf.isDraftComplete) completeCount++;
      if (AppState.commercial.isDraftComplete) completeCount++;

      const progress = Math.round((completeCount / 3) * 100);

      const progressBar = document.getElementById('overallProgressBar');
      const progressPercentage = document.getElementById('overallPercentage');
      const progressStatus = document.getElementById('progressStatus');
      const progressDetailsText = document.getElementById('progressDetailsText');

      if (progressBar) {
        progressBar.style.width = `${progress}%`;
        if (progress === 100) {
          progressBar.classList.add('complete');
        } else {
          progressBar.classList.remove('complete');
        }
      }

      if (progressPercentage) {
        progressPercentage.textContent = `${progress}%`;
      }

      if (progressDetailsText) {
        progressDetailsText.textContent = `${completeCount} of 3 reports completed`;
      }

      if (progressStatus) {
        if (AppState.officiallySubmitted) {
          progressStatus.textContent = 'Successfully submitted';
          progressStatus.classList.add('complete');
        } else if (completeCount === 3) {
          progressStatus.textContent = 'Ready to submit';
          progressStatus.style.color = '#2B6E3E';
        } else if (completeCount > 0) {
          progressStatus.textContent = `${completeCount} of 3 reports completed`;
          progressStatus.style.color = '#F59E0B';
        } else {
          progressStatus.textContent = '0 of 3 reports completed';
          progressStatus.style.color = '#666';
        }
      }

      const submitBtn = document.getElementById('submitAllButton');
      if (submitBtn) {
        const canSubmit = completeCount === 3 && !AppState.officiallySubmitted;
        submitBtn.disabled = !canSubmit;

        if (canSubmit) {
          submitBtn.classList.add('pulse-ready');
        } else {
          submitBtn.classList.remove('pulse-ready');
        }
      }

      this.updateNavigationButtons();
    },

    onMonthChange(sourceInputId) {
      // Sync all month inputs to the same value
      const sourceInput = sourceInputId ? document.getElementById(sourceInputId) : null;
      const selectedMonth = sourceInput?.value || Utils.currentMonth();

      ['petrominMonth', 'gulfMonth', 'commercialMonth'].forEach(id => {
        const el = document.getElementById(id);
        if (el && el.value !== selectedMonth) {
          el.value = selectedMonth;
        }
      });

      // Reset submission state when month changes
      AppState.officiallySubmitted = false;

      // Re-enable form fields, but preserve auto-calculated fields (complaints & field trials)
      const autoCalcIds = ['petrominComplaintHandled', 'gulfComplaintHandled', 'commercialComplaintHandled',
        'petrominFieldTrial', 'gulfFieldTrial', 'commercialFieldTrial'];
      document.querySelectorAll('input[type="number"], input[type="month"]').forEach(el => {
        if (autoCalcIds.includes(el.id)) return;
        el.readOnly = false;
        el.style.backgroundColor = '';
        el.style.cursor = '';
      });

      // Re-check if new month is already submitted
      FormManager.loadDrafts();

      // Auto-fill complaints and field trials for each form
      ['petromin', 'gulf', 'commercial'].forEach(ft => {
        const monthInput = document.querySelector(`#${ft}Month`);
        if (monthInput && monthInput.value) {
          AutoCalcManager.autoFillForMonth(ft, monthInput.value);
        }
      });

      // Update UI
      this.updateProgress();
    },

    displayUserInfo() {
      const userInfo = document.getElementById('userInfo');
      const userName = window.PORTAL_USER_NAME || 'Guest';

      if (userInfo) {
        userInfo.textContent = `Welcome, ${userName}`;
      }
    },

    updateWorkflowSteps() {
      const steps = {
        petromin: document.querySelector('[data-step="petromin"]'),
        gulf: document.querySelector('[data-step="gulf"]'),
        commercial: document.querySelector('[data-step="commercial"]'),
        review: document.querySelector('[data-step="review"]'),
        dashboard: document.querySelector('[data-step="dashboard"]')
      };

      ['petromin', 'gulf', 'commercial'].forEach(step => {
        const element = steps[step];
        if (!element) return;

        element.classList.remove('completed', 'disabled', 'active');

        if (AppState[step].isDraftComplete) {
          element.classList.add('completed');
          const num = element.querySelector('.step-number');
          const check = element.querySelector('.step-complete');
          if (num) num.classList.add('hidden');
          if (check) check.classList.remove('hidden');
        } else if (AppState.currentStep === step) {
          element.classList.add('active');
        } else if (!this.canNavigateTo(step)) {
          element.classList.add('disabled');
        }
      });

      const canReview = AppState.petromin.isDraftComplete &&
        AppState.gulf.isDraftComplete &&
        AppState.commercial.isDraftComplete;
      steps.review?.classList.toggle('disabled', !canReview);

      if (AppState.currentStep === 'dashboard') {
        steps.dashboard?.classList.add('active');
      }
    },

    // FIXED: Review summary now shows split training fields
    updateReviewSummary() {
      ['petromin', 'gulf', 'commercial'].forEach(ft => {
        const d = FormManager.collectData(ft) || {};
        const monthEl = document.getElementById(`${ft}SummaryMonth`);
        if (monthEl) monthEl.textContent = Utils.formatMonth(d.month);

        const list = [
          ['Customer Courtesy Visits', Utils.nzInt(d.customerCourtesyVisits)],
          ['Technical Enquiries / Proposals', Utils.nzInt(d.productEquivalent)],
          ['Internal Technical Sessions', Utils.nzInt(d.internalTraining)],
          ['Onsite Customer Trainings', Utils.nzInt(d.onsiteTraining)],
          ['Business Opportunities', Utils.nzInt(d.businessOpportunity)],
          ['Complaints Handled', Utils.nzInt(d.complaintHandled)],
          ['Product Field Trials', Utils.nzInt(d.fieldTrial)],
          ['Lubrication Surveys', Utils.nzInt(d.lubricationSurvey)],
          ['DVR Endorsements', Utils.nzInt(d.dvrEndorsement)]
        ];

        const activitiesEl = document.getElementById(`${ft}SummaryActivities`);
        if (activitiesEl) {
          activitiesEl.innerHTML = list.map(([label, val]) =>
            `<div class="summary-item"><span>${label}</span><span>${val}</span></div>`
          ).join('');
        }

        const totalEl = document.getElementById(`${ft}SummaryTotal`);
        if (totalEl) totalEl.textContent = Utils.computeTotal(d);

        const statusEl = document.getElementById(`${ft}Status`);
        if (statusEl) {
          if (AppState.officiallySubmitted) {
            statusEl.innerHTML = '<i class="fas fa-check-circle"></i> Officially Submitted';
            statusEl.className = 'status-indicator submitted';
          } else if (AppState[ft].isDraftComplete) {
            statusEl.innerHTML = '<i class="fas fa-check"></i> Draft Ready';
            statusEl.className = 'status-indicator complete';
          } else {
            statusEl.innerHTML = '<i class="fas fa-edit"></i> Draft In Progress';
            statusEl.className = 'status-indicator draft';
          }
        }
      });

      // ADD THIS INSTEAD (in updateReviewSummary, after the forEach loop)
      const reviewContainer = document.getElementById('reviewContent');
      if (reviewContainer && AppState.officiallySubmitted) {
        const existingBanner = reviewContainer.querySelector('.submission-status-banner');
        if (existingBanner) existingBanner.remove();

        const banner = document.createElement('div');
        banner.className = 'submission-status-banner';
        banner.style.cssText = `
          background: linear-gradient(135deg, #10B981 0%, #059669 100%);
          color: white;
          padding: 20px;
          border-radius: 12px;
          margin-bottom: 24px;
          box-shadow: 0 4px 6px rgba(16, 185, 129, 0.2);
        `;
        banner.innerHTML = `
          <div style="display: flex; align-items: center;">
            <div>
              <div style="font-size: 18px; font-weight: 600; margin-bottom: 4px;">
                <i class="fas fa-check-circle"></i> Reports Successfully Submitted
              </div>
              <div style="font-size: 14px; opacity: 0.9;">
                This month's reports have been officially submitted and are now read-only
              </div>
            </div>
          </div>
        `;
        reviewContainer.insertBefore(banner, reviewContainer.firstChild);
      }

    }
  };

  /* ==========================
     SUBMISSION MANAGER
  ========================== */
  const SubmissionManager = {
    submitAll() {
      const btn = document.getElementById('submitAllButton');
      const spinner = document.getElementById('submitSpinner');
      if (btn?.disabled) return;

      // ✅ CHECK 0: Block if auto-calc is still loading
      if (AutoCalcManager.isAnyLoading()) {
        Notify.warning('Please wait — complaints and field trials are still loading.', 'Auto-Calc In Progress');
        return;
      }

      // ✅ CHECK 1: All drafts must be complete
      if (!AppState.petromin.isDraftComplete || !AppState.gulf.isDraftComplete || !AppState.commercial.isDraftComplete) {
        Notify.error('Complete all drafts first.', 'Incomplete Drafts');
        return;
      }

      // ✅ CHECK 2: Prevent double submission
      if (AppState.officiallySubmitted) {
        Notify.warning('Reports have already been submitted for this month.', 'Already Submitted');
        return;
      }

      const petrominData = FormManager.collectData('petromin');
      const gulfData = FormManager.collectData('gulf');
      const commercialData = FormManager.collectData('commercial');

      if (!FormManager.validate('petromin', petrominData) ||
        !FormManager.validate('gulf', gulfData) ||
        !FormManager.validate('commercial', commercialData)) {
        return;
      }

      // Auto-sync months if they somehow drifted apart
      if (!(petrominData.month === gulfData.month && petrominData.month === commercialData.month)) {
        const resolvedMonth = petrominData.month || gulfData.month || commercialData.month;
        ['petrominMonth', 'gulfMonth', 'commercialMonth'].forEach(id => {
          const el = document.getElementById(id);
          if (el) el.value = resolvedMonth;
        });
        petrominData.month = resolvedMonth;
        gulfData.month = resolvedMonth;
        commercialData.month = resolvedMonth;
        Notify.info('Reporting month has been aligned across all forms.', 'Month Synchronized');
      }

      const currentMonth = Utils.currentMonth();
      if (petrominData.month > currentMonth) {
        Notify.error('Cannot submit future months.', 'Invalid Month');
        return;
      }

      spinner?.classList.remove('hidden');
      if (btn) btn.disabled = true;

      petrominData.lob = CONFIG.LOB.Petromin;
      gulfData.lob = CONFIG.LOB.Gulf;
      commercialData.lob = CONFIG.LOB.Commercial;

      // ✅ CRITICAL: Set status explicitly
      petrominData.status = CONFIG.STATUS.Submitted;
      gulfData.status = CONFIG.STATUS.Submitted;
      commercialData.status = CONFIG.STATUS.Submitted;

      const attemptAll = (tries = 0) => Promise.all([
        DataverseAPI.submitOrUpdate(petrominData),
        DataverseAPI.submitOrUpdate(gulfData),
        DataverseAPI.submitOrUpdate(commercialData)
      ]).catch(err => {
        if (tries < CONFIG.MAX_RETRIES && (err?.status >= 500 || err?.status === 0)) {
          return new Promise(res => setTimeout(res, Math.pow(2, tries) * 1000))
            .then(() => attemptAll(tries + 1));
        }
        throw err;
      });

      attemptAll().then(() => {
        AppState.officiallySubmitted = true;
        AppState.totalSubmittedReports += 3;

        // Lock all fields immediately
        FormManager.lockAllFields();

        WorkflowManager.updateProgress();

        Notify.success('All reports submitted to Dataverse.', 'Submission Complete');
        setTimeout(() => WorkflowManager.navigateTo('dashboard'), 1200);
      }).catch((error) => {
        console.error('Submission error:', error);
        if (error.status === 0 || error.status >= 500) {
          this.addToOfflineQueue(petrominData, gulfData, commercialData);
          Notify.warning('Network/server issue. Saved offline and will retry when online.', 'Offline Queue');
        } else {
          Notify.error(this.extractError(error), 'Submission Error');
        }
      }).finally(() => {
        spinner?.classList.add('hidden');
        if (btn) btn.disabled = false;
      });
    },

    extractError(error) {
      try {
        if (error.responseJSON?.error?.message) return error.responseJSON.error.message;
        if (typeof error === 'string') return error;
        if (error.statusText) return error.statusText;
      } catch { }
      return 'Unexpected error';
    },

    addToOfflineQueue(petrominData, gulfData, commercialData) {
      try {
        const queue = JSON.parse(localStorage.getItem('tsr_offline_queue') || '[]');
        queue.push({
          id: Date.now(),
          timestamp: Utils.nowIso(),
          petromin: petrominData,
          gulf: gulfData,
          commercial: commercialData
        });
        localStorage.setItem('tsr_offline_queue', JSON.stringify(queue));
        AppState.offlineQueue = queue;
      } catch (e) {
        console.error('Offline queue save failed', e);
      }
    },

    processOfflineQueue() {
      const queue = JSON.parse(localStorage.getItem('tsr_offline_queue') || '[]');
      if (!queue.length) return;

      const removeFromQueue = (itemId) => {
        const updated = JSON.parse(localStorage.getItem('tsr_offline_queue') || '[]').filter(q => q.id !== itemId);
        localStorage.setItem('tsr_offline_queue', JSON.stringify(updated));
        AppState.offlineQueue = updated;
      };

      console.log(`Processing ${queue.length} offline submissions...`);
      queue.forEach(item => {
        item.petromin.lob = CONFIG.LOB.Petromin;
        item.gulf.lob = CONFIG.LOB.Gulf;
        item.commercial.lob = CONFIG.LOB.Commercial;

        Promise.all([
          DataverseAPI.submitOrUpdate(item.petromin),
          DataverseAPI.submitOrUpdate(item.gulf),
          DataverseAPI.submitOrUpdate(item.commercial)
        ]).then(() => {
          removeFromQueue(item.id);
          AppState.totalSubmittedReports += 3;
          WorkflowManager.updateProgress();
          Notify.success('Offline submission processed.', 'Queue');
        }).catch((e) => {
          console.error('Offline processing failed', e);
          // Remove items with permanent (4xx) errors — only keep retryable server/network errors
          if (e?.status && e.status >= 400 && e.status < 500) {
            removeFromQueue(item.id);
            Notify.error(`Offline submission from ${new Date(item.timestamp).toLocaleDateString()} failed permanently and was removed. Please re-enter the data.`, 'Queue Error');
          }
        });
      });
    }
  };

  /* ==========================
     DASHBOARD MANAGER (FULLY FIXED)
  ========================== */
  const DashboardManager = {
    load() {
      const statusEl = document.getElementById('dashboardStatusText');
      if (statusEl) statusEl.textContent = 'Loading your dashboard data...';

      DataverseAPI.getMyReports(500)
        .done((resp) => {
          const items = resp?.value || [];
          this.processData(items);

          const submittedItems = items.filter(item => item.cr650_status === CONFIG.STATUS.Submitted);
          AppState.totalSubmittedReports = submittedItems.length;
          WorkflowManager.updateProgress();

          if (statusEl) {
            const userInfo = window.PORTAL_USER_NAME ? `for ${window.PORTAL_USER_NAME}` : '';
            statusEl.textContent = items.length ?
              `Showing ${items.length} reports ${userInfo}.` :
              `No reports found ${userInfo}.`;
          }
        })
        .fail((err) => {
          console.error('Dashboard load error:', err);
          Notify.error('Failed to load dashboard data.', 'Dashboard Error');
          if (statusEl) statusEl.textContent = 'Failed to load data.';
        });
    },

    // FIXED: Now uses split training fields throughout
    buildMonthBuckets(items) {
      const buckets = {};
      items.forEach(r => {
        const month = (r.cr650_reportmonth || '').slice(0, 7);
        if (!month) return;
        if (!buckets[month]) buckets[month] = this.getEmptyMetrics();

        buckets[month].courtesyVisits += Utils.nzInt(r.cr650_customercourtesyvisits);
        buckets[month].productEquivalent += Utils.nzInt(r.cr650_productequivalent);
        buckets[month].internalTraining += Utils.nzInt(r.cr650_internaltraining);
        buckets[month].onsiteTraining += Utils.nzInt(r.cr650_onsitetraining);
        buckets[month].opportunities += Utils.nzInt(r.cr650_businessopportunity);
        buckets[month].complaints += Utils.nzInt(r.cr650_complainthandled);
        buckets[month].fieldTrials += Utils.nzInt(r.cr650_fieldtrial);
        buckets[month].surveys += Utils.nzInt(r.cr650_lubricationsurvey);
        buckets[month].dvr += Utils.nzInt(r.cr650_dvrendorsement);

        // FIXED: Total calculation now uses split training fields
        buckets[month].totalActivities =
          buckets[month].courtesyVisits +
          buckets[month].productEquivalent +
          buckets[month].internalTraining +
          buckets[month].onsiteTraining +
          buckets[month].opportunities +
          buckets[month].complaints +
          buckets[month].fieldTrials +
          buckets[month].surveys +
          buckets[month].dvr;
      });
      return { buckets, months: Object.keys(buckets).sort() };
    },

    // FIXED: Now includes split training fields
    getEmptyMetrics() {
      return {
        courtesyVisits: 0,
        productEquivalent: 0,
        internalTraining: 0,  // FIXED: Was 'training'
        onsiteTraining: 0,    // FIXED: Added
        opportunities: 0,
        complaints: 0,
        fieldTrials: 0,
        surveys: 0,
        dvr: 0,
        totalActivities: 0
      };
    },

    aggregateAll(items) {
      const totals = this.getEmptyMetrics();
      const monthData = this.buildMonthBuckets(items);
      Object.values(monthData.buckets).forEach(m => {
        Object.keys(totals).forEach(k => { totals[k] += m[k]; });
      });
      const months = monthData.months;
      const latest = months[months.length - 1];
      const changes = this.calculateMonthOverMonth(latest, months, monthData.buckets);
      return { totals, changes };
    },

    calculateMonthOverMonth(selected, months, buckets) {
      const idx = months.indexOf(selected);
      let prev = null;
      for (let i = idx - 1; i >= 0; i--) {
        if (buckets[months[i]]?.totalActivities > 0) {
          prev = months[i];
          break;
        }
      }
      if (!prev) return this.getEmptyMetrics();
      const cur = buckets[selected] || this.getEmptyMetrics();
      const prv = buckets[prev] || this.getEmptyMetrics();
      const delta = {};
      Object.keys(cur).forEach(k => delta[k] = cur[k] - prv[k]);
      return delta;
    },

    populateMonthSlicer(months) {
      const slicer = document.getElementById('monthSlicer');
      if (!slicer) return;
      const current = slicer.value;
      slicer.innerHTML = '<option value="__all__">All months</option>';
      months.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = Utils.formatMonth(m);
        slicer.appendChild(opt);
      });
      if (current && (current === '__all__' || months.includes(current))) {
        slicer.value = current;
      }
    },

    processData(items) {
      const { buckets, months } = this.buildMonthBuckets(items);
      this.populateMonthSlicer(months);

      const selected = document.getElementById('monthSlicer')?.value || '__all__';
      let metrics, changes;
      if (selected === '__all__') {
        const agg = this.aggregateAll(items);
        metrics = agg.totals;
        changes = agg.changes;
      } else {
        metrics = buckets[selected] || this.getEmptyMetrics();
        changes = this.calculateMonthOverMonth(selected, months, buckets);
      }

      this.updateMetrics(metrics, changes);
    },

    // FIXED: Now maps to correct HTML element IDs for split training
    updateMetrics(metrics, changes) {
      const map = {
        totalCourtesyVisits: metrics.courtesyVisits,
        totalInternalTraining: metrics.internalTraining,  // FIXED: New element
        totalOnsiteTraining: metrics.onsiteTraining,      // FIXED: New element
        totalOpportunities: metrics.opportunities,
        totalComplaints: metrics.complaints,
        totalFieldTrials: metrics.fieldTrials,
        totalSurveys: metrics.surveys,
        totalDVR: metrics.dvr,
        totalActivities: metrics.totalActivities
      };
      for (const [id, val] of Object.entries(map)) {
        const el = document.getElementById(id);
        if (el) el.textContent = Utils.nzInt(val).toLocaleString();
      }

      // FIXED: Change indicators for split training fields
      const deltaMap = {
        courtesyVisitsChange: changes.courtesyVisits,
        internalTrainingChange: changes.internalTraining,  // FIXED
        onsiteTrainingChange: changes.onsiteTraining,      // FIXED
        opportunitiesChange: changes.opportunities,
        complaintsChange: changes.complaints,
        fieldTrialsChange: changes.fieldTrials,
        surveysChange: changes.surveys,
        dvrChange: changes.dvr,
        totalActivitiesChange: changes.totalActivities
      };
      for (const [id, ch] of Object.entries(deltaMap)) {
        const el = document.getElementById(id);
        if (!el) continue;
        if (ch === undefined || ch === 0) {
          el.textContent = 'No change';
          el.className = 'kpi-change neutral';
        } else if (ch > 0) {
          el.textContent = `+${ch} vs previous`;
          el.className = 'kpi-change positive';
        } else {
          el.textContent = `${ch} vs previous`;
          el.className = 'kpi-change negative';
        }
      }
    },

    refresh() {
      Notify.info('Refreshing...', 'Refresh');
      this.load();
    },

    showReports() {
      DataverseAPI.getMyReports(200).done((resp) => {
        const items = resp?.value || [];
        if (!items.length) {
          Notify.info('No submitted reports yet.', 'Reports');
          return;
        }
        let html = '<div style="max-height:400px; overflow-y:auto;">';
        items.forEach(r => {
          const month = (r.cr650_reportmonth || '').slice(0, 7);
          const lob = r.cr650_lineofbusiness === CONFIG.LOB.Petromin ? 'Petromin'
            : r.cr650_lineofbusiness === CONFIG.LOB.Gulf ? 'Gulf' : 'Commercial';
          const status = r.cr650_status === CONFIG.STATUS.Submitted ? 'Submitted' : 'Draft';
          const total = Utils.nzInt(r.cr650_totalactivities);
          const dt = r.cr650_submissiondate ? new Date(r.cr650_submissiondate).toLocaleDateString() : 'N/A';
          html += `<div style="border:1px solid #e2e8f0; border-radius:8px; padding:16px; margin:10px 0; background:#fff;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
              <strong>${lob} - ${Utils.formatMonth(month)}</strong>
              <span style="color:${status === 'Submitted' ? '#16A34A' : '#64748b'}; font-weight:600;">${status}</span>
            </div>
            <div style="font-size:14px; color:#64748b;">
              <div>Total Activities: <strong>${total}</strong></div>
              <div>Submitted: ${dt}</div>
            </div>
          </div>`;
        });
        html += '</div>';
        Notify.show(html, 'info', 'Your Reports History', 0);
      }).fail(() => Notify.error('Failed to load reports history.', 'Error'));
    },

    // FIXED: Export now includes split training columns
    exportData() {
      const sel = document.getElementById('monthSlicer')?.value || '__all__';
      DataverseAPI.getMyReports(1000, sel === '__all__' ? null : sel)
        .done((resp) => {
          const items = resp?.value || [];
          if (!items.length) {
            Notify.warning('No data to export.', 'Export');
            return;
          }
          // FIXED: Headers now have two training columns
          const headers = [
            'Month', 'Line of Business', 'Status', 'Submission Date',
            'Customer Courtesy Visits', 'Technical Enquiries / Proposals',
            'Internal Technical Sessions', 'Onsite Customer Trainings',  // FIXED: Split columns
            'Business Opportunities', 'Complaints Handled', 'Product Field Trials',
            'Lubrication Surveys', 'DVR Endorsements', 'Total Activities'
          ];
          // FIXED: Row data now uses correct field names
          const rows = items.map(it => [
            (it.cr650_reportmonth || '').slice(0, 7),
            it.cr650_lineofbusiness === CONFIG.LOB.Petromin ? 'Petromin' : it.cr650_lineofbusiness === CONFIG.LOB.Gulf ? 'Gulf' : 'Commercial',
            it.cr650_status === CONFIG.STATUS.Submitted ? 'Submitted' : 'Draft',
            it.cr650_submissiondate ? new Date(it.cr650_submissiondate).toLocaleDateString() : '',
            it.cr650_customercourtesyvisits || 0,
            it.cr650_productequivalent || 0,
            it.cr650_internaltraining || 0,   // FIXED
            it.cr650_onsitetraining || 0,     // FIXED
            it.cr650_businessopportunity || 0,
            it.cr650_complainthandled || 0,
            it.cr650_fieldtrial || 0,
            it.cr650_lubricationsurvey || 0,
            it.cr650_dvrendorsement || 0,
            it.cr650_totalactivities || 0
          ]);
          let csv = headers.join(',') + '\n';
          rows.forEach(r => {
            csv += r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',') + '\n';
          });

          const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
          const link = document.createElement('a');
          link.href = URL.createObjectURL(blob);
          link.download = sel === '__all__' ? 'tsr-my-reports.csv' : `tsr-${sel}-report.csv`;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          Notify.success('Export complete.', 'Export');
        })
        .fail(() => Notify.error('Export failed.', 'Export Error'));
    }
  };

  /* ==========================
     EVENT HANDLERS
  ========================== */
  const EventHandlers = {
    _monthChangeTimer: null,

    init() {
      const debounced = Utils.debounce(() => WorkflowManager.updateProgress(), CONFIG.DEBOUNCE_DELAY);

      document.querySelectorAll('input[type="number"], input[type="month"]').forEach(input => {
        input.addEventListener('input', debounced);
        input.addEventListener('change', function () {
          debounced();
          // If month changed, sync all months and reload data
          if (this.type === 'month') {
            const inputId = this.id;
            clearTimeout(EventHandlers._monthChangeTimer);
            EventHandlers._monthChangeTimer = setTimeout(() => WorkflowManager.onMonthChange(inputId), 500);
          }
        });
        input.addEventListener('input', function () {
          this.classList.remove('error');
          const err = document.getElementById(this.id + '-error');
          if (err) err.style.display = 'none';
        });
      });

      document.querySelectorAll('.executive-workflow-step').forEach(step => {
        step.addEventListener('click', function () {
          if (!this.classList.contains('disabled')) {
            WorkflowManager.navigateTo(this.dataset.step);
          }
        });
      });

      const slicer = document.getElementById('monthSlicer');
      slicer?.addEventListener('change', () => DashboardManager.load());

      document.addEventListener('keydown', function (e) {
        if (e.altKey && e.key === 'ArrowRight') {
          e.preventDefault();
          WorkflowManager.navigateNext();
        }
        if (e.altKey && e.key === 'ArrowLeft') {
          e.preventDefault();
          WorkflowManager.navigateBack();
        }
        if (e.ctrlKey && e.key === 's') {
          e.preventDefault();
          const st = AppState.currentStep;
          if (st === 'petromin' || st === 'gulf' || st === 'commercial') {
            FormManager.saveDraft(st);
          }
        }
      });

      window.addEventListener('online', () => {
        Notify.success('Online. Processing offline queue...', 'Online');
        SubmissionManager.processOfflineQueue();
      });
      window.addEventListener('offline', () => {
        Notify.warning('Offline. Submissions will be queued.', 'Offline');
      });

      window.addEventListener('beforeunload', function (e) {
        if ((AppState.petromin.isDraftComplete ||
          AppState.gulf.isDraftComplete ||
          AppState.commercial.isDraftComplete) &&
          !AppState.officiallySubmitted) {
          const msg = 'You have unsent drafts. Leave the page?';
          e.returnValue = msg;
          return msg;
        }
      });
    }
  };

  /* ==========================
     APP INITIALIZATION
  ========================== */
  const App = {
    init() {
      const overlay = document.getElementById('loadingOverlay');
      overlay?.classList.remove('hidden');

      try {
        this.displayUserInfo();
        WorkflowManager.setDefaultMonth();
        FormManager.loadDrafts();
        // Auto-fill complaints and field trials on initial load
        setTimeout(() => {
          ['petromin', 'gulf', 'commercial'].forEach(ft => {
            const monthInput = document.querySelector(`#${ft}Month`);
            if (monthInput && monthInput.value) {
              AutoCalcManager.autoFillForMonth(ft, monthInput.value);
            }
          });
        }, 1500);
        EventHandlers.init();
        WorkflowManager.updateProgress();
        DashboardManager.load();
        WorkflowManager.navigateTo('petromin');
        if (navigator.onLine) SubmissionManager.processOfflineQueue();

        this.initializeProgressBar();

        Notify.info('System ready. Dashboard is always accessible.', 'Welcome');
      } catch (e) {
        console.error('Initialization error:', e);
        Notify.error('Initialization failed. Refresh the page.', 'Initialization Error');
      } finally {
        overlay?.classList.add('hidden');
      }
    },

    initializeProgressBar() {
      const progressBar = document.getElementById('overallProgressBar');
      if (progressBar && progressBar.parentElement) {
        const container = progressBar.parentElement;
        container.style.backgroundColor = '#e2e8f0';
        container.style.height = '8px';
        container.style.borderRadius = '4px';
        container.style.overflow = 'hidden';
        container.style.position = 'relative';

        progressBar.style.position = 'absolute';
        progressBar.style.left = '0';
        progressBar.style.top = '0';
        progressBar.style.height = '100%';
        progressBar.style.backgroundColor = '#1B8B47';
        progressBar.style.transition = 'width 0.3s ease';
        progressBar.style.width = '0%';
      }
    },

    displayUserInfo() {
      const el = document.getElementById('userInfo');
      if (el && window.PORTAL_USER_NAME) {
        el.textContent = `Welcome, ${window.PORTAL_USER_NAME}`;
      }

      const userName = window.PORTAL_USER_NAME || 'User';
      const initials = userName.split(' ').map(n => n[0]).join('').toUpperCase().slice(0, 2);
      const avatar = document.getElementById('userAvatar');
      if (avatar) avatar.textContent = initials;
    }
  };

  /* ==========================
     GLOBAL EXPORTS
  ========================== */
  window.saveDraft = (ft) => FormManager.saveDraft(ft);
  window.submitAllReports = () => SubmissionManager.submitAll();
  window.navigateNext = () => WorkflowManager.navigateNext();
  window.navigateBack = () => WorkflowManager.navigateBack();
  window.updateProgress = () => WorkflowManager.updateProgress();
  window.resetAllForms = () => FormManager.reset();
  window.refreshDashboard = () => DashboardManager.refresh();
  window.exportDashboard = () => DashboardManager.exportData();
  window.showReports = () => DashboardManager.showReports();
  window.goNextStep = () => WorkflowManager.navigateNext();
  window.goPrevStep = () => WorkflowManager.navigateBack();

  $(document).ready(function () { App.init(); });

  window.refreshAutoCalc = () => {
    ['petromin', 'gulf', 'commercial'].forEach(ft => {
      const monthInput = document.querySelector(`#${ft}Month`);
      if (monthInput && monthInput.value) {
        AutoCalcManager.autoFillForMonth(ft, monthInput.value);
      }
    });
  };

})(window, document, window.jQuery);
