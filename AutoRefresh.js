'use strict';

(() => {
  const DEFAULT_INTERVAL_SEC = 15;
  let refreshIntervalId = null;
  let activeDatasourceIdList = [];
  let uniqueDataSources = [];

  document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync({ configure })
      .then(() => {
        getSettings();
        tableau.extensions.settings.addEventListener(
          tableau.TableauEventType.SettingsChanged,
          (evt) => updateExtensionBasedOnSettings(evt.newSettings)
        );
        if (tableau.extensions.settings.get("configured") !== "1") {
          configure();
        }
      });
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
    const popupUrl = `${window.location.origin}/AutoRefreshDialog.html`;
    tableau.extensions.ui
      .displayDialogAsync(popupUrl, DEFAULT_INTERVAL_SEC.toString(), { height: 500, width: 500 })
      .then((closePayload) => {
        toggleUI(true);
        updateUI("interval", closePayload);
        setupRefreshInterval(Number(closePayload));
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
      try {
        await Promise.all(uniqueDataSources.map((ds) => ds.refreshAsync()));
        updateNextRefreshTime(intervalSec);
      } catch (err) {
        console.error("Error refreshing datasources:", err);
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
  }

  function toggleUI(isActive) {
    document.getElementById("active").classList.toggle("d-none", !isActive);
    document.getElementById("inactive").classList.toggle("d-none", isActive);
  }

  function updateUI(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  }
})();
