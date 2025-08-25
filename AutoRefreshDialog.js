'use strict';

(function () {
  const datasourcesSettingsKey = 'selectedDatasources';
  const intervalkey = 'intervalkey';
  const configured = 'configured';
  let selectedDatasources = [];

  $(document).ready(function () {
    tableau.extensions.initializeDialogAsync().then(function (openPayload) {
      $('#interval').val(openPayload);
      $('#closeButton').click(closeDialog);

      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const visibleDatasources = [];

      if (tableau.extensions.settings.get(configured) == 1) {
        $('#interval').val(tableau.extensions.settings.get(intervalkey));
      } else {
        $('#interval').val(60);
      }

      selectedDatasources = parseSettingsForActiveDataSources();

      dashboard.worksheets.forEach(function (worksheet) {
        worksheet.getDataSourcesAsync().then(function (datasources) {
          datasources.forEach(function (datasource) {
            const isActive = selectedDatasources.includes(datasource.id);
            if (!visibleDatasources.includes(datasource.id)) {
              addDataSourceItemToUI(datasource, isActive);
              visibleDatasources.push(datasource.id);
            }
          });
        });
      });
    });
  });

  function parseSettingsForActiveDataSources() {
    let activeDatasourceIdList = [];
    const settings = tableau.extensions.settings.getAll();
    if (settings.selectedDatasources) {
      activeDatasourceIdList = JSON.parse(settings.selectedDatasources);
    }
    return activeDatasourceIdList;
  }

  function updateDatasourceList(id) {
    const idx = selectedDatasources.indexOf(id);
    if (idx < 0) {
      selectedDatasources.push(id);
    } else {
      selectedDatasources.splice(idx, 1);
    }
  }

  function addDataSourceItemToUI(datasource, isActive) {
    const containerDiv = $('<div />');
    $('<input />', {
      type: 'checkbox',
      id: datasource.id,
      value: datasource.name,
      checked: isActive,
      click: function () { updateDatasourceList(datasource.id); }
    }).appendTo(containerDiv);

    $('<label />', {
      for: datasource.id,
      text: datasource.name
    }).appendTo(containerDiv);

    $('#datasources').append(containerDiv);
  }

  function closeDialog() {
    tableau.extensions.settings.set(datasourcesSettingsKey, JSON.stringify(selectedDatasources));
    tableau.extensions.settings.set(intervalkey, $('#interval').val());
    tableau.extensions.settings.set(configured, 1);
    tableau.extensions.settings.saveAsync().then(() => {
      tableau.extensions.ui.closeDialog($('#interval').val());
    });
  }
})();
