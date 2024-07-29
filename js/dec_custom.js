// #region DEC.Namespaces
var DEC = {}; // Application Namespace
DEC.Core = {}; // Core Functions
DEC.DOM = {}; // DOM Functions
DEC.Utilities = {}; // Utilities Functions
DEC.Xrm = {}; // Xrm Functions
// #endregion

// #region DEC.Utilities
/**
 * Utilities - Has Value
 * Returns true if a parameter is not undefined, not null and not an empty string, otherwise returns false
 * @param {any} parameter Parameter to check
 */
DEC.Utilities.HasValue = function (parameter) {
    if (parameter !== undefined && parameter !== null && parameter !== '') { return true; } else { return false; }
}

/**
 * Utilities - Download File
 * Download a file
 */
DEC.Utilities.DownloadFile = function (blob, fileName) {
    try {
        var customLink = document.createElement('a');
        customLink.href = URL.createObjectURL(blob);
        customLink.download = fileName;
        customLink.click();
    }
    catch (e) { }
}

DEC.Utilities.SplitText = function (text, separators) {
    let firstSeparator = separators[0];
    for (let i = 1; i < separators.length; i++) { text = text.split(separators[i]).join(firstSeparator); }
    text = text.split(firstSeparator);
    return text;
}
// #endregion

// #region DEC.Xrm
/**
 * Xrm - Get Xrm Object
 */
DEC.Xrm.GetXrmObject = function () {
    try {
        if (typeof parent !== 'undefined') { return parent.Xrm; } else { return undefined; }
    } catch { return undefined; }
}

/**
 * Xrm - Is Demo Mode
 */
DEC.Xrm.IsDemoMode = function () {
    return typeof DEC.Xrm.GetXrmObject() === 'undefined';
}

/**
 * Xrm - Is Instance Mode
 */
DEC.Xrm.IsInstanceMode = function () {
    return typeof DEC.Xrm.GetXrmObject() !== 'undefined';
}

/**
 * Xrm - Get Client Url
 */
DEC.Xrm.GetClientUrl = function () {
    if (DEC.Xrm.IsInstanceMode()) { return DEC.Xrm.GetXrmObject().Utility.getGlobalContext().getClientUrl(); }
    if (DEC.Xrm.IsDemoMode()) { return 'https://democall'; }
}

/**
 * Xrm - Get Context
 */
DEC.Xrm.GetContext = function () {
    var context = 'Demo';
    if (DEC.Xrm.IsInstanceMode()) { context = DEC.Xrm.GetClientUrl(); }
    return '<small>(' + context + ')</small>';
}

/**
 * Xrm - Get Version
 */
DEC.Xrm.GetVersion = function () {
    var currentVersion = '';
    if (DEC.Xrm.IsInstanceMode()) { currentVersion = DEC.Xrm.GetXrmObject().Utility.getGlobalContext().getVersion(); }
    if (DEC.Xrm.IsDemoMode()) { currentVersion = '9.2.0.0'; }

    if (!DEC.Utilities.HasValue(currentVersion)) { return ''; }
    var versionArray = currentVersion.split('.');
    if (versionArray.length < 2) { return ''; }
    return versionArray[0] + '.' + versionArray[1];
}
// #endregion

// #region DEC.DOM
DEC.DOM.FillSelectDisabled = function (id, headerText, values, valueProperty, textProperty, dataPropertyName, dataProperty) {
    $('#' + id).empty();
    let headerOption = $('<option>', { value: '', text: headerText + ': ' + values.length });
    headerOption.prop('selected', 'selected');
    headerOption.prop('disabled', 'disabled');
    $('#' + id).append(headerOption);

    values.forEach(function (item) {
        let option = $('<option>', { value: item[valueProperty], text: item[textProperty] });
        option.prop('disabled', 'disabled');
        if (DEC.Utilities.HasValue(dataPropertyName)) {
            option.attr('data-' + dataPropertyName, item[dataProperty]);
        }
        $('#' + id).append(option);
    });
}
// #endregion


// #region DEC.Core
DEC.Core.CleanEmails = function () {
    let sourceTextAreaId = 'txtEmailsToClean';
    let targetSelectId = 'selCleanedEmails';
    let nextBtnsIds = ['btnLoadUsers', 'btnLoadUsersOnlyDataverse', 'btnLoadUsersOnlyEntra'];

    let btn = $('#btnCleanEmails');
    let originalText = btn.text();
    $('#' + sourceTextAreaId).prop('disabled', 'disabled');
    btn.text('Please wait...').prop('disabled', 'disabled');
    $('#' + targetSelectId).empty().prop('disabled', 'disabled');
    nextBtnsIds.forEach(function (item) { $('#' + item).prop('disabled', 'disabled'); });

    setTimeout(() => {
        let content = $('#' + sourceTextAreaId).val();
        let firstSplit = DEC.Utilities.SplitText(content, [',', ' ', ';', '\t', '\n', '\r\n']);
        let hash = new Set();
        firstSplit.forEach(function (item) { if (item.includes('@')) { hash.add(item.trim().toLowerCase()); } });
        let arrSet = Array.from(hash);
        arrSet.sort();

        let emailAddresses = [];
        arrSet.forEach(function (item) { emailAddresses.push({ 'email': item }); });

        DEC.DOM.FillSelectDisabled(targetSelectId, 'Email founds', emailAddresses, 'email', 'email');
        $('#' + targetSelectId).prop('disabled', '');
        nextBtnsIds.forEach(function (item) { $('#' + item).prop('disabled', ''); });
        $('#' + sourceTextAreaId).prop('disabled', '');
        btn.text(originalText).prop('disabled', '');
    }, 300);
}

DEC.Core.GetAllResults = async function (fetchUrl, pageSize, tableName) {
    let allResults = [];
    let fetchNextLink = false;
    do {
        const res = await fetch(fetchUrl, {
            method: 'GET',
            headers: {
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Content-Type': 'application/json; charset=utf-8',
                'Accept': 'application/json',
                'Prefer': 'odata.include-annotations=*,odata.maxpagesize=' + pageSize
            }
        });

        const results = await res.json();
        allResults.push(...results.value);
        $('#spnResults').text('Current ' + tableName + ': ' + allResults.length);
        if (results.hasOwnProperty('@odata.nextLink')) { fetchUrl = results['@odata.nextLink']; fetchNextLink = true; } else { fetchNextLink = false; }
    } while (fetchNextLink);
    $('#spnResults').text('');
    return allResults;
}

DEC.Core.LoadUsers = async function (loadDataverseUsers, loadEntraUsers) {
    let btnsIds = ['btnLoadUsers', 'btnLoadUsersOnlyDataverse', 'btnLoadUsersOnlyEntra'];

    let dataverseUsersId = 'selDataverseUsers';
    let entraUsersId = 'selEntraUsers';
    let nextStepId = 'btnCheckEmails';

    let originalTexts = [];

    [$('#' + btnsIds[0]).text(), $('#' + btnsIds[1]).text(), $('#' + btnsIds[2]).text()];
    btnsIds.forEach(function (item) { originalTexts.push($('#' + item).text()); $('#' + item).text('Plase wait...').prop('disabled', 'disabled'); });

    $('#' + dataverseUsersId).prop('disabled', 'disabled');
    $('#' + entraUsersId).prop('disabled', 'disabled');
    $('#' + nextStepId).prop('disabled', 'disabled');

    let dataverseUsers = [];
    let entraUsers = [];

    if (DEC.Xrm.IsDemoMode() == true) {
        if (loadDataverseUsers == true) {
            dataverseUsers.push({ 'id': 1, 'email': 'user1@contoso.com', 'fullname': 'User One', 'enabled': true });
            dataverseUsers.push({ 'id': 2, 'email': 'user2@contoso.com', 'fullname': 'User Two', 'enabled': false });
        }
        if (loadEntraUsers == true) {
            entraUsers.push({ 'id': 5, 'email': 'user1@contoso.com', 'fullname': 'User One', 'enabled': true });
            entraUsers.push({ 'id': 7, 'email': 'user2@contoso.com', 'fullname': 'User Two', 'enabled': false });
            entraUsers.push({ 'id': 8, 'email': 'user3@contoso.com', 'fullname': 'User Three', 'enabled': true });
            entraUsers.push({ 'id': 8, 'email': 'user4@contoso.com', 'fullname': 'User Foud', 'enabled': false });
        }
    }

    if (DEC.Xrm.IsInstanceMode() == true) {
        if (loadDataverseUsers == true) {
            let url = DEC.Xrm.GetXrmObject().Utility.getGlobalContext().getClientUrl() + '/api/data/v' + DEC.Xrm.GetVersion() + '/systemusers?$select=fullname,domainname,isdisabled';
            let results = await DEC.Core.GetAllResults(url, 900, 'Dataverse Users');
            results.forEach(function (item) {
                dataverseUsers.push({ 'id': item.systemuserid, 'email': item.domainname.toLowerCase(), 'fullname': item.fullname, 'enabled': !item.isdisabled });
            });
        }
        if (loadEntraUsers == true) {
            let url = DEC.Xrm.GetXrmObject().Utility.getGlobalContext().getClientUrl() + '/api/data/v' + DEC.Xrm.GetVersion() + '/aadusers?$select=displayname,mail,accountenabled&$filter=mail ne null';
            let results = await DEC.Core.GetAllResults(url, 900, 'Entra Users');
            results.forEach(function (item) {
                entraUsers.push({ 'id': item.aaduserid, 'email': item.mail.toLowerCase(), 'fullname': item.displayname, 'enabled': item.accountenabled });
            });
        }
    }

    if (loadDataverseUsers == true) {
        let dataverseUsersFormatted = [];
        dataverseUsers.forEach(function (item) {
            let enabledText = 'Enabled';
            if (item.enabled == false) { enabledText = 'Disabled'; }
            dataverseUsersFormatted.push({ 'email': item.email, 'enabled': item.enabled, 'name': item.email + ' - ' + item.fullname + ' - ' + enabledText, });
        });
        DEC.DOM.FillSelectDisabled(dataverseUsersId, 'Dataverse Users found', dataverseUsersFormatted, 'email', 'name', 'enabled', 'enabled');
    }
    if (loadEntraUsers == true) {
        let entraUsersFormatted = [];
        entraUsers.forEach(function (item) {
            let enabledText = 'Enabled';
            if (item.enabled == false) { enabledText = 'Disabled'; }
            entraUsersFormatted.push({ 'email': item.email, 'enabled': item.enabled, 'name': item.email + ' - ' + item.fullname + ' - ' + enabledText, });
        });
        DEC.DOM.FillSelectDisabled(entraUsersId, 'Entra Users found', entraUsersFormatted, 'email', 'name', 'enabled', 'enabled');
    }
    //$('#' + dataverseUsersId).prop('disabled', '');
    //$('#' + entraUsersId).prop('disabled', '');
    btnsIds.forEach(function (item, index) { $('#' + item).text(originalTexts[index]).prop('disabled', ''); });
    $('#' + nextStepId).prop('disabled', '');
}

DEC.Core.CheckEmails = function (sourceButton) {

    let sourceEmailsId = 'selCleanedEmails';
    let sourceDataverseUsersId = 'selDataverseUsers';
    let sourceEntraUsersId = 'selEntraUsers';

    let targetSelectsIds = ['selDataverseUsersEnabled', 'selDataverseUsersDisabled', 'selEntraUsersEnabled', 'selEntraUsersDisabled', 'selEntraUsersEnabledNoDataverse', 'selEntraUsersDisabledNoDataverse', 'selMissingEmails'];
    targetSelectsIds.forEach(function (item) { $('#' + item).empty(); $('#' + item).prop('disabled', 'disabled'); });
    let btnId = 'btnCheckEmails';
    let btn = $('#' + btnId);
    let originalText = btn.text();
    btn.text('Please wait...').prop('disabled', 'disabled');

    setTimeout(() => {
        let emails = [];
        let dataverseUsers = [];
        let entraUsers = [];
        $('#' + sourceEmailsId + ' option').not(':first').each(function () { emails.push($(this).val()); });
        $('#' + sourceDataverseUsersId + ' option').not(':first').each(function () { dataverseUsers.push({ 'email': $(this).val(), 'text': $(this).text(), 'enabled': $(this).data('enabled') }); });
        $('#' + sourceEntraUsersId + ' option').not(':first').each(function () { entraUsers.push({ 'email': $(this).val(), 'text': $(this).text(), 'enabled': $(this).data('enabled') }); });

        let dataverseUsersEnabled = [];
        let entraUsersEnabled = [];
        let dataverseUsersDisabled = [];
        let entraUsersDisabled = [];
        let entraUsersEnabledNoDataverse = [];
        let entraUsersDisabledNoDataverse = [];
        let missingEmails = [];

        emails.forEach(function (item) {
            let resultDataverseUsers = $(dataverseUsers).filter(function (i, n) { return n.email == item; });
            let resultEntraUsers = $(entraUsers).filter(function (i, n) { return n.email == item; });

            if (resultDataverseUsers.length == 0 && resultEntraUsers.length == 0) { missingEmails.push(item); }

            if (resultDataverseUsers.length == 0 && resultEntraUsers.length > 0) {
                resultEntraUsers.each(function (i, n) { if (n.enabled == true) { entraUsersEnabledNoDataverse.push(n); } else { entraUsersDisabledNoDataverse.push(n); } });
            }

            if (resultDataverseUsers.length > 0) {
                resultDataverseUsers.each(function (i, n) { if (n.enabled == true) { dataverseUsersEnabled.push(n); } else { dataverseUsersDisabled.push(n); } });
            }

            if (resultEntraUsers.length > 0) {
                resultEntraUsers.each(function (i, n) { if (n.enabled == true) { entraUsersEnabled.push(n); } else { entraUsersDisabled.push(n); } });
            }
        });

        DEC.DOM.FillSelectDisabled('selDataverseUsersEnabled', 'Dataverse Users', dataverseUsersEnabled, 'email', 'text');
        DEC.DOM.FillSelectDisabled('selDataverseUsersDisabled', 'Dataverse Users', dataverseUsersDisabled, 'email', 'text');
        DEC.DOM.FillSelectDisabled('selEntraUsersEnabled', 'Entra Users', entraUsersEnabled, 'email', 'text');
        DEC.DOM.FillSelectDisabled('selEntraUsersDisabled', 'Entra Users', entraUsersDisabled, 'email', 'text');
        DEC.DOM.FillSelectDisabled('selEntraUsersEnabledNoDataverse', 'Entra Users', entraUsersEnabledNoDataverse, 'email', 'text');
        DEC.DOM.FillSelectDisabled('selEntraUsersDisabledNoDataverse', 'Entra Users', entraUsersDisabledNoDataverse, 'email', 'text');

        let missingEmailsArr = [];
        missingEmails.forEach(function (item) { missingEmailsArr.push({ 'email': item }); });
        DEC.DOM.FillSelectDisabled('selMissingEmails', 'Missing Emails', missingEmailsArr, 'email', 'email');
        targetSelectsIds.forEach(function (item) { $('#' + item).prop('disabled', ''); });
        btn.text(originalText).prop('disabled', '');
    }, 300);
}

DEC.Core.DownloadContent = function (dropdown, entraHeader) {
    let count = $('#' + dropdown + ' option').not(':first').length;
    if (count == 0) { alert('No Values'); }
    else {
        let content = [];
        if (entraHeader == true) {
            content.push('version:v1.0');
            content.push('Member object ID or user principal name [memberObjectIdOrUpn] Required');
        }
        $('#' + dropdown + ' option').not(':first').each(function (index, element) {
            content.push(element.value);
        });


        let fileContent = content.join('\r\n', content);
        let now = new Date().toISOString();
        now = now.substring(0, now.length - 5);
        now = now.replaceAll('-', '_').replaceAll(':', '_').replaceAll('T', '_');


        var saveFile = new Blob([fileContent], { type: 'application/csv;charset=utf-8' });
        let filename = dropdown.substring(3).toLowerCase() + '_' + now + '.csv';
        DEC.Utilities.DownloadFile(saveFile, filename)
    }
}

// #endregion
$(document).ready(function () {
    $('#span_context').html(DEC.Xrm.GetContext());
});