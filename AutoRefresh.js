'use strict';

(() => {
  const DEFAULT_INTERVAL_SEC = 15;
  let refreshIntervalId = null;
  let activeDatasourceIdList = [];
  let uniqueDataSources = [];

  document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync()
      .then(() => {
        getSettings();
        tableau.extensions.settings.addEventListener(
          tableau.TableauEventType.SettingsChanged,
          (evt) => updateExtensionBasedOnSettings(evt.newSettings)
        );

        if (tableau.extensions.settings.get("configured") !== "1") {
          configure();
        }
      })
      .catch((err) => console.error("Extension init error:", err));
  });

  function getSettings() {
    const settings = tableau.extensions.settings.getAll();
    if (settings.selectedDatasources) {
      activeDatasourceIdList = JSON.parse(settings.selectedDatasources);
    }

    const interval = settings.intervalkey || DEFAULT_INTERVAL_SEC;
    if (activeDatasourceIdList.length) {
      toggleUI(true);
      updateUI("interval", interval);
      updateUI("datasourceCount", activeDatasourceIdList.length);
      setupRefreshInterval(Number(interval));
    } else {
      toggleUI(false);
    }
  }

  function configure() {
    const baseUrl = window.location.origin;
    const popupUrl = `${baseUrl}/AutoRefreshDialog.html`;

    tableau.extensions.ui
      .displayDialogAsync(popupUrl, DEFAULT_INTERVAL_SEC.toString(), { height: 500, width: 500 })
      .then((closePayload) => {
        toggleUI(true);

        // Try parsing payload safely
        let interval = Number(closePayload);
        if (isNaN(interval)) {
          try { interval = Number(JSON.parse(closePayload)); } catch {}
        }

        if (!interval || isNaN(interval)) {
          interval = DEFAULT_INTERVAL_SEC;
        }

        updateUI("interval", interval);
        setupRefreshInterval(interval);
      })
      .catch((error) => {
        if (error.errorCode === tableau.ErrorCodes.DialogClosedByUser) {
          console.info("Configuration dialog closed by user.");
        } else {
          console.error("Dialog error:", error.message);
        }
      });
  }

  function setupRefreshInterval(intervalSec) {
    clearInterval(refreshIntervalId);

    const collectUniqueDataSources = async () => {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const uniqueIds = new Set();
      uniqueDataSources = [];

      await Promise.all(
        dashboard.worksheets.map((ws) =>
          ws.getDataSourcesAsync().then((datasources) => {
            datasources.forEach((ds) => {
              if (activeDatasourceIdList.includes(ds.id) && !uniqueIds.has(ds.id)) {
                uniqueIds.add(ds.id);
                uniqueDataSources.push(ds);
              }
            });
          })
        )
      );

      updateUI("uniqueCount", uniqueDataSources.length);
    };

    const refreshDataSources = async () => {
      if (uniqueDataSources.length === 0) {
        console.warn("No valid datasources selected for refresh.");
        return;
      }
      try {
        await Promise.all(uniqueDataSources.map((ds) => ds.refreshAsync()));
        updateNextRefreshTime(intervalSec);
        updateUI("errorMessage", "");
        const errBox = document.getElementById("errorMessage");
        if (errBox) errBox.classList.add("d-none");
      } catch (err) {
        console.error("Error refreshing datasources:", err);
        const errBox = document.getElementById("errorMessage");
        if (errBox) {
          errBox.textContent = "⚠️ Error refreshing datasources. Check connection or permissions.";
          errBox.classList.remove("d-none");
        }
      }
    };

    collectUniqueDataSources().then(() => {
      refreshDataSources();
      refreshIntervalId = setInterval(refreshDataSources, intervalSec * 1000);
    });
  }

  function updateNextRefreshTime(intervalSec) {
    const nextRefresh = new Date(Date.now() + intervalSec * 1000);
    updateUI("nextrefresh", nextRefresh.toLocaleTimeString());
  }

  function updateExtensionBasedOnSettings(settings) {
    if (settings.selectedDatasources) {
      activeDatasourceIdList = JSON.parse(settings.selectedDatasources);
      updateUI("datasourceCount", activeDatasourceIdList.length);
    }
    if (settings.intervalkey) {
      updateUI("interval", settings.intervalkey);
      setupRefreshInterval(Number(settings.intervalkey));
    }
  }

  function toggleUI(isActive) {
    const activeEl = document.getElementById("active");
    const inactiveEl = document.getElementById("inactive");

    if (activeEl) activeEl.classList.toggle("d-none", !isActive);
    if (inactiveEl) inactiveEl.classList.toggle("d-none", isActive);
  }

  function updateUI(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  }

  window.AutoRefresh = window.AutoRefresh || {};
  window.AutoRefresh.configure = configure;
})();
