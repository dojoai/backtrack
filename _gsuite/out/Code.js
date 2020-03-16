"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var BaseRunner = /** @class */ (function () {
    function BaseRunner() {
    }
    BaseRunner.prototype.runGmail = function (userEmail, timer, state) {
        return FLAG.FINISHED;
    };
    BaseRunner.prototype.run = function () {
        var _a, _b;
        var settings = getDomainSettings();
        var state = getState('Gmail');
        log(state);
        var timer = new Timer(state.resource, state.timeElapsed);
        var flag = getFlag();
        timer.startTimer();
        var activeUser = activeUsers.moveToNext();
        while (activeUser && !timer.isTimerExpired()) {
            timer.setUser(activeUser);
            flag = settings.getMessages && state.resource == "Gmail" ? this.runGmail(activeUser, timer, state) : FLAG.FINISHED;
            if (flag == FLAG.FINISHED)
                state.resource = "Calendar";
            if (!timer.isTimerExpired()) {
                if (settings.getEvents && state.resource == "Calendar") {
                    flag = runCalendar(activeUser, timer, state);
                    if (flag == FLAG.FINISHED)
                        state.resource = "Gmail";
                }
            }
            if (flag == FLAG.FINISHED) {
                var msg_1 = "Done " + activeUser;
                log(msg_1);
                toast(msg_1, 5);
                timer.showProgress(msg_1);
                (_b = (_a = getActiveUserRange(activeUser)) === null || _a === void 0 ? void 0 : _a.getCell(1, DOMAIN_USERS_SHEET.COL_PROGRESS)) === null || _b === void 0 ? void 0 : _b.setValue("DONE");
                activeUser = activeUsers.moveToNext();
            }
            else {
                log('Not finished, flag: ' + flag);
            }
        }
        // if all done for ALL users
        var msg;
        if (flag == FLAG.FINISHED && !timer.isTimerExpired()) {
            msg = 'DONE in ' + timer.getTimeFormatted(timer.totalElapsedTime);
            log(msg);
            timer.setTask('');
            timer.stopTimer(msg);
        }
        else {
            msg = 'Paused, total elapsed: ' + timer.getTimeFormatted(timer.totalElapsedTime);
            SpreadsheetApp.getActive().toast(msg);
            log(msg);
            timer.stopTimer();
            createTrigger();
        }
        setFlag(flag);
        return flag;
    };
    return BaseRunner;
}());
var BaseUserSheet = /** @class */ (function () {
    function BaseUserSheet() {
    }
    /**
     *
    * Creates combined metrics file from the template for all users where the corresponding column is empty
    */
    BaseUserSheet.prototype.createMissingUserSpreadsheets = function () {
        var _this = this;
        console.time('createMissingUserSpreadsheets');
        var shu = SpreadsheetApp.getActive().getSheetByName(SHEET_DOMAIN_USERS);
        var users = shu.getDataRange().getValues().slice(DOMAIN_USERS_SHEET.FIRST_DATA_ROW - 1);
        users.forEach(function (row, i) {
            var url;
            if (row[2] == "") {
                var email = row[0];
                shu.getRange(DOMAIN_USERS_SHEET.FIRST_DATA_ROW + i, DOMAIN_USERS_SHEET.COL_SHEET_URL, 1, 1).setValue('please wait, working...');
                SpreadsheetApp.flush(); // show progress
                console.time('createFile');
                url = _this.createFile(email);
                console.timeEnd('createFile');
                shu.getRange(DOMAIN_USERS_SHEET.FIRST_DATA_ROW + i, DOMAIN_USERS_SHEET.COL_SHEET_URL, 1, 1).setValue(url);
                SpreadsheetApp.flush(); // show progress
            }
        });
        console.timeEnd('createMissingUserSpreadsheets');
    };
    return BaseUserSheet;
}());
function listMockEvents(minCount, userEmail, maxResults, query, pageToken) {
    var events = listEvents(userEmail, maxResults, query, pageToken);
    var realEventCount = events.items.length;
    if (realEventCount == 0) {
        log('cannot mock events, there are no real events for set criteria');
        return events;
    }
    if (minCount > maxResults) {
        minCount = maxResults;
    }
    if (realEventCount < minCount) {
        var fillUpItem = events.items[0];
        // fill up to count
        for (var i = 0; i < minCount - realEventCount; i++) {
            events.items.push(fillUpItem);
        }
    }
    return events;
}
function getMockThreads(minCount, userEmail, maxResults, query, pageToken) {
    var _a, _b;
    var threads = getThreads(userEmail, maxResults, query, pageToken);
    var realThreadCount = (_b = (_a = threads) === null || _a === void 0 ? void 0 : _a.threads) === null || _b === void 0 ? void 0 : _b.length;
    if (!realThreadCount) {
        log('cannot mock threads, there are no real threads for set criteria');
        return threads;
    }
    if (minCount > maxResults) {
        minCount = maxResults;
    }
    if (realThreadCount < minCount) {
        var fillUpThrad = threads.threads[0];
        // fill up to count
        for (var i = 0; i < minCount - realThreadCount; i++) {
            threads.threads.push(fillUpThrad);
        }
    }
    return threads;
}
/// <reference path="../Common/Mock.ts" />
var CalendarQuery = /** @class */ (function () {
    function CalendarQuery(appSettings) {
        var _a, _b;
        this.appSettings = appSettings;
        this.before = new Date(appSettings.before);
        this.after = new Date(appSettings.after);
        if (this.before < this.after)
            throw "Before date must follow After date";
        this.startWorkday = getDecimalHours((_a = appSettings.startTime, (_a !== null && _a !== void 0 ? _a : '00:00')));
        this.endWorkday = getDecimalHours((_b = appSettings.endTime, (_b !== null && _b !== void 0 ? _b : '24:00')));
    }
    CalendarQuery.prototype.toString = function () {
        var timeMin = this.after.getTime();
        var timeMax = this.before.getTime();
        var min = Utilities.formatDate(new Date(timeMin), Session.getScriptTimeZone(), "yyyy-MM-dd'T'hh:mm:ssZ");
        var max = Utilities.formatDate(new Date(timeMax), Session.getScriptTimeZone(), "yyyy-MM-dd'T'hh:mm:ssZ");
        var q = this.appSettings.internalOnlyEvents ? '&q=' + this.appSettings.domain : '';
        return q + "&timeMin=" + min + "&timeMax=" + max;
    };
    return CalendarQuery;
}());
function runCalendar(userEmail, timer, state) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t, _u, _v, _w, _x, _y, _z, _0, _1;
    setFlag(FLAG.RUNNING);
    timer.setTask('Calendar');
    var flag = FLAG.FINISHED;
    try {
        var query = getCalendarQuery();
        log('fetching events from calendar API');
        var events = TESTING ?
            listMockEvents(400, userEmail, MAX_EVENTS, query.toString()) :
            listEvents(userEmail, MAX_THREADS, query.toString());
        var counts = {}, people = {}, daily = {}, volumes = [], everyone = {};
        var recent_1 = {};
        log('event items: ' + events.items.length);
        var event_1;
        for (var i = (_a = state.startingEventIndex, (_a !== null && _a !== void 0 ? _a : 0)); i < events.items.length && !timer.isTimerExpired(); i++) {
            timer.displayElapsedTime();
            try {
                event_1 = events.items[i];
                var guests = (_d = (_c = (_b = event_1.attendees) === null || _b === void 0 ? void 0 : _b.filter(function (attendee) { return !attendee.self; })) === null || _c === void 0 ? void 0 : _c.map(function (atendee) { var _a; return (_a = atendee) === null || _a === void 0 ? void 0 : _a.email; }), (_d !== null && _d !== void 0 ? _d : []));
                var creators = [(_f = (_e = event_1.creator) === null || _e === void 0 ? void 0 : _e.email, (_f !== null && _f !== void 0 ? _f : 'not available'))];
                // Calendar tab
                // log('Calendar tab, items ' + creators.length)
                for (var j = 0; j < creators.length; j++) {
                    for (var k = 0; k < guests.length; k++)
                        incrementCount(counts, creators[j], guests[k]);
                }
                // People tab
                var attendees = creators;
                // log('People tab, items ' + attendees.length)
                for (var j = 0; j < guests.length; j++) {
                    var guest = guests[j];
                    if (attendees.indexOf(guest) == -1)
                        attendees.push(guest);
                }
                for (var j = 0; j < attendees.length; j++) {
                    for (var k = 0; k < attendees.length; k++)
                        if (j != k)
                            incrementCount(people, attendees[j], attendees[k]);
                    everyone[attendees[j]] = true;
                }
                attendees.forEach(function (attendee) {
                    var _a, _b, _c, _d;
                    if (attendee != userEmail) {
                        recent_1[attendee] = recent_1[attendee] || [];
                        var meetingDate = new Date(Date.parse((_d = (_b = (_a = event_1.start) === null || _a === void 0 ? void 0 : _a.date, (_b !== null && _b !== void 0 ? _b : (_c = event_1.start) === null || _c === void 0 ? void 0 : _c.dateTime)), (_d !== null && _d !== void 0 ? _d : ''))));
                        recent_1[attendee].push(meetingDate);
                    }
                });
                // Daily tab
                var isAllDay = ((_g = event_1.start) === null || _g === void 0 ? void 0 : _g.date) != null;
                if (!isAllDay && attendees.length > 1) {
                    // log('Daily tab')
                    var start = Date.parse((_l = (_j = (_h = event_1.start) === null || _h === void 0 ? void 0 : _h.date, (_j !== null && _j !== void 0 ? _j : (_k = event_1.start) === null || _k === void 0 ? void 0 : _k.dateTime)), (_l !== null && _l !== void 0 ? _l : '')));
                    var end = Date.parse((_p = (_o = (_m = event_1.end) === null || _m === void 0 ? void 0 : _m.date, (_o !== null && _o !== void 0 ? _o : event_1.end.dateTime)), (_p !== null && _p !== void 0 ? _p : '')));
                    var date = Utilities.formatDate(new Date(start), Session.getScriptTimeZone(), "yyyyMMdd");
                    var hours = (end - start) / (60 * 60 * 1000);
                    if (!daily[date])
                        daily[date] = { busy: 0, meetings: [] };
                    daily[date].busy += hours;
                    daily[date].meetings.push({ start: getDecimalHours(new Date(start)), end: getDecimalHours(new Date(end)) });
                }
                // Volumes tab
                if ((_q = event_1.organizer) === null || _q === void 0 ? void 0 : _q.self) {
                    // log('Volumes tab')
                    volumes.push({
                        atendees: attendees,
                        date: new Date((_u = (_s = (_r = event_1.start) === null || _r === void 0 ? void 0 : _r.date, (_s !== null && _s !== void 0 ? _s : (_t = event_1.start) === null || _t === void 0 ? void 0 : _t.dateTime)), (_u !== null && _u !== void 0 ? _u : ''))),
                        duration: getDuration(new Date((_w = (_v = event_1.start) === null || _v === void 0 ? void 0 : _v.date, (_w !== null && _w !== void 0 ? _w : (_x = event_1.start) === null || _x === void 0 ? void 0 : _x.dateTime))), new Date((_1 = (_z = (_y = event_1.end) === null || _y === void 0 ? void 0 : _y.date, (_z !== null && _z !== void 0 ? _z : (_0 = event_1.end) === null || _0 === void 0 ? void 0 : _0.dateTime)), (_1 !== null && _1 !== void 0 ? _1 : '')))),
                    });
                }
            }
            catch (e) {
                log(e);
                state.timeElapsed = 0;
                timer.stopTimer(e);
                throw e;
            }
        }
        state.timeElapsed = timer.totalElapsedTime;
        if (timer.isTimerExpired()) {
            state.startingEventIndex = i;
            saveState(state);
            createTrigger();
            flag = FLAG.SLEEPING;
        }
        else {
            log('save calendar data');
            state.startingEventIndex = 0;
            var outpuSheet = new OutputSheet();
            outpuSheet.saveCalendar(counts);
            outpuSheet.savePeople(people, everyone);
            outpuSheet.saveDaily(daily, query.after.getTime(), query.before.getTime(), query.startWorkday, query.endWorkday);
            outpuSheet.saveVolume(volumes);
            outpuSheet.saveGmail(state);
            outpuSheet.saveRecent(recent_1);
            flag = FLAG.FINISHED;
        }
        log(state);
        return flag;
    }
    catch (e) {
        log(e);
        return FLAG.FINISHED;
    }
}
function getDuration(start, end) {
    if (!start || !end || typeof start.getTime != "function" || typeof end.getTime != "function")
        return 0;
    return (end.getTime() - start.getTime()) / 3600000;
}
/**
 *
 * @param emailList repalces aeach of the string in the list which does not contain a domain, with "EXTERNAL"
 */
function anonymize(emailList, domain) {
    if (domain === void 0) { domain = getDomainSettings().domain; }
    return emailList.map(function (item) { return domain && item.indexOf(domain) >= 0 ? item : 'EXTERNAL'; });
}
function getCalendarQuery() {
    var appSettings = getDomainSettings();
    if (!appSettings.before || !appSettings.after)
        throw 'Before/after date not set';
    return new CalendarQuery(appSettings);
}
var HANDLER_FUNCTION = "restart";
var DELAY_MIN = 1;
var Timer = /** @class */ (function () {
    function Timer(task, totalElapsedTime) {
        if (totalElapsedTime === void 0) { totalElapsedTime = 0; }
        this.task = task;
        this.totalElapsedTime = 0;
        this.maxScriptRunTimeMin = parseInt(getDomainSettings().maxScriptRunTimeMin);
        log('constructor maxScriptRuntimeMin=' + this.maxScriptRunTimeMin);
        this.totalElapsedTime = totalElapsedTime;
    }
    Timer.prototype.setTask = function (task) {
        this.task = task;
    };
    Timer.prototype.setUser = function (userEmail) {
        var range = getActiveUserRange(userEmail);
        this.progressCell = range.getCell(1, DOMAIN_USERS_SHEET.COL_PROGRESS);
    };
    Timer.prototype.showProgress = function (message) {
        if (this.progressCell) {
            this.progressCell.setValue((this.task ? this.task + ': ' : '') + message);
            SpreadsheetApp.flush();
        }
        else {
            console.log(message);
        }
    };
    Timer.prototype.displayElapsedTime = function () {
        var _a;
        this.showProgress(this.getTimeFormatted((_a = this.elapsedTime, (_a !== null && _a !== void 0 ? _a : 0))) + '/' + this.getTimeFormatted(this.totalElapsedTime) + ' elapsed');
    };
    Timer.prototype.startTimer = function () {
        log('start Timer');
        this.previousTime = new Date().getTime();
        this.elapsedTime = 0;
        this.showProgress('Timer started');
    };
    Timer.prototype.isTimerExpired = function () {
        if (this.previousTime == undefined || this.elapsedTime == undefined || this.totalElapsedTime == undefined)
            throw 'Timer has not been started';
        var now = new Date().getTime();
        this.elapsedTime += now - this.previousTime;
        this.totalElapsedTime += now - this.previousTime;
        this.previousTime = now;
        var isExpired = this.elapsedTime > this.maxScriptRunTimeMin * 60 * 1000;
        return isExpired;
    };
    Timer.prototype.getTimeFormatted = function (millisec) {
        var totSeconds = Math.floor(millisec / 1000);
        var seconds = totSeconds % 60;
        var minutes = Math.floor(totSeconds / 60);
        return Utilities.formatString('%02d:%02d', minutes, seconds);
    };
    /**
     *
     * @param message stops the timer
     * @returns total elapsed time
     */
    Timer.prototype.stopTimer = function (message) {
        var _a;
        log('stop Timer, time elapsed: ' + this.getTimeFormatted((_a = this.elapsedTime, (_a !== null && _a !== void 0 ? _a : 0))) + ', total time: ' + this.getTimeFormatted(this.totalElapsedTime));
        this.previousTime = undefined;
        this.elapsedTime = undefined;
        if (message)
            this.showProgress(message);
    };
    return Timer;
}());
function restart() {
    log('restart Timer');
    deleteTrigger();
    new Runner().run();
}
function createTrigger() {
    deleteTrigger();
    log('create trigger');
    ScriptApp.newTrigger(HANDLER_FUNCTION).timeBased().everyMinutes(DELAY_MIN).create();
}
function deleteTrigger() {
    var triggers = ScriptApp.getProjectTriggers();
    log('delete trigger ' + triggers.length);
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() == HANDLER_FUNCTION)
            ScriptApp.deleteTrigger(triggers[i]);
    }
}
function saveState(state) {
    var _a, _b, _c;
    var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_INPROGRESS);
    // the header row
    var values = [[state.query, (_a = state.nextPageToken, (_a !== null && _a !== void 0 ? _a : '')), (_b = state.startingThreadIndex, (_b !== null && _b !== void 0 ? _b : '')), state.timeElapsed, state.resource, (_c = state.startingEventIndex, (_c !== null && _c !== void 0 ? _c : ''))]];
    // the rest lines
    OutputSheet.appendCounts(values, state.counts);
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    if (sheet.getLastRow() > values.length) {
        sheet.getRange(values.length + 1, 1, sheet.getLastRow() - values.length, values[0].length).clear();
    }
}
function getState(defaultResource) {
    if (defaultResource === void 0) { defaultResource = 'Gmail'; }
    var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_INPROGRESS);
    var data = sheet.getDataRange().getValues();
    var headerRow = data[0];
    var state = {
        resource: headerRow[INPROGRESS_SHEET.COL_RESOURCE - 1] ? headerRow[INPROGRESS_SHEET.COL_RESOURCE - 1] : defaultResource,
        query: (headerRow[INPROGRESS_SHEET.COL_QUERY - 1] ? headerRow[INPROGRESS_SHEET.COL_QUERY - 1] : ""),
        nextPageToken: (headerRow[INPROGRESS_SHEET.COL_NEXT_PAGE_TOKEN - 1] ? headerRow[INPROGRESS_SHEET.COL_NEXT_PAGE_TOKEN - 1] : ""),
        startingThreadIndex: (headerRow[INPROGRESS_SHEET.COL_STARTING_THREAD_INDEX - 1] ? headerRow[INPROGRESS_SHEET.COL_STARTING_THREAD_INDEX - 1] : 0),
        startingEventIndex: (headerRow[INPROGRESS_SHEET.COL_STARTING_EVENT_INDEX - 1] ? headerRow[INPROGRESS_SHEET.COL_STARTING_EVENT_INDEX - 1] : 0),
        timeElapsed: headerRow[INPROGRESS_SHEET.COL_TIME_ELAPSED - 1] ? headerRow[INPROGRESS_SHEET.COL_TIME_ELAPSED - 1] : 0,
        counts: {}
    };
    for (var i = 1; i < data.length; i++) {
        var sender = data[i][INPROGRESS_SHEET.COL_SENDER - 1];
        var receiver = data[i][INPROGRESS_SHEET.COL_RECEIVER - 1];
        var count = data[i][INPROGRESS_SHEET.COL_COUNT - 1];
        incrementCount(state.counts, sender, receiver, count);
    }
    return state;
}
/* eslint-disable no-unused-vars */
function getUsers(maxResults) {
    var allUsers = [];
    var nextPageToken;
    do {
        var response = getApiResponse("admin/directory/v1/users?maxResults=" + maxResults + "&domain=" + getDomainSettings().domain + (nextPageToken ? "&pageToken=" + nextPageToken : ""));
        var usersResponse = JSON.parse(response.getContentText());
        nextPageToken = usersResponse.nextPageToken;
        log('new users ' + usersResponse.users.length + ' nextPageToken: ' + usersResponse.nextPageToken);
        allUsers = allUsers.concat(usersResponse.users);
    } while (nextPageToken);
    return allUsers;
}
function getMessages(userEmail, maxResults, threadId, pageToken) {
    var response = getApiResponse("gmail/v1/users/" + userEmail + "/threads/" + threadId + "?format=metadata&metadataHeaders=To&metadataHeaders=From&metadataHeaders=Cc&metadataHeaders=Bcc&maxResults=" + maxResults + (pageToken ? "&pageToken=" + pageToken : ""), userEmail);
    var messages = JSON.parse(response.getContentText());
    return messages.messages;
}
function getThreads(userEmail, maxResults, query, pageToken) {
    var response = getApiResponse("gmail/v1/users/" + userEmail + "/threads?maxResults=" + maxResults + (query ? "&q=" + query : "") + (pageToken ? "&pageToken=" + pageToken : ""), userEmail);
    var threads = JSON.parse(response.getContentText());
    return threads;
}
function listEvents(userEmail, maxResults, query, pageToken) {
    var response = getApiResponse("calendar/v3/calendars/" + userEmail + "/events?maxResults=" + maxResults + query + (pageToken ? "&pageToken=" + pageToken : ""), userEmail);
    var events = JSON.parse(response.getContentText());
    return events;
}
function reset() {
    getService(getDomainSettings()).reset();
}
function getService(domainSettings, user) {
    return OAuth2.createService('Domain:' + domainSettings.domainAdmin)
        .setAuthorizationBaseUrl(domainSettings.token_uri)
        .setTokenUrl(domainSettings.token_uri)
        .setPrivateKey(domainSettings.private_key)
        .setIssuer(domainSettings.client_email)
        .setSubject(user ? user : domainSettings.domainAdmin)
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope("https://www.googleapis.com/auth/admin.directory.user.readonly https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/calendar.events.readonly");
}
function getApiResponse(endPoint, user) {
    log(endPoint);
    var domainSettings = getDomainSettings();
    var service = getService(domainSettings, user ? user : domainSettings.domainAdmin);
    service.reset();
    if (service.hasAccess()) {
        var url = "https://www.googleapis.com/" + endPoint;
        var response = UrlFetchApp.fetch(url, {
            "headers": {
                "Authorization": 'Bearer ' + service.getAccessToken()
            },
            "muteHttpExceptions": true
        });
        if (response.getResponseCode() != 200) {
            var responseText = response.getContentText();
            //log(response.getResponseCode()+ ":" + responseText)
            var error = JSON.parse(responseText);
            throw error.error.message;
        }
        return response;
    }
    else {
        throw new Error(service.getLastError());
    }
}
/// <reference path="Service.ts" />
function getDomainUsers() {
    var users = getUsers(MAX_USERS);
    if (users.length > 0) {
        var rows = users.map(function (user) {
            return [user.primaryEmail, user.name.fullName];
        });
        var shu = SpreadsheetApp.getActive().getSheetByName(SHEET_DOMAIN_USERS);
        shu.getRange(DOMAIN_USERS_SHEET.FIRST_DATA_ROW, 1, shu.getMaxRows() - DOMAIN_USERS_SHEET.FIRST_DATA_ROW + 1, shu.getMaxColumns()).clearContent();
        shu.getRange(DOMAIN_USERS_SHEET.FIRST_DATA_ROW, 1, rows.length, rows[0].length).setValues(rows);
    }
}
var CACHE_EXPIRATION = 10;
var MAX_LOG_MESSAGE_LENGTH = 5000;
var SH_LOG = "Log";
var ActiveUsers = /** @class */ (function () {
    function ActiveUsers() {
        this.users = [];
        try {
            var usersProp = properties.activeUsers;
            if (usersProp)
                this.users = usersProp;
        }
        catch (e) {
            log(e);
        }
    }
    ActiveUsers.prototype.loadUsers = function (users) {
        this.users = users;
        this.storeUsers();
    };
    ActiveUsers.prototype.activeUser = function () {
        if (!this.active)
            throw "User list is empty or not initialized. Use moveNext()";
        return this.active;
    };
    ActiveUsers.prototype.moveToNext = function () {
        var u1 = this.active;
        if (this.active) {
            this.users.splice(0, 1);
            this.storeUsers();
        }
        this.active = this.users.length > 0 ? this.users[0] : undefined;
        log('Move to next user: ' + u1 + '=>' + this.active);
        return this.active;
    };
    ActiveUsers.prototype.storeUsers = function () {
        properties.activeUsers = this.users;
    };
    return ActiveUsers;
}());
/**
 * Convert date-only date string into datetime string at 0:00 in the script's time zone
 *
 * @param dateString date-only string
 */
function getDateInTimezone(dateString) {
    var timezone = Utilities.formatDate(new Date(dateString), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "Z");
    var dateTime = dateString + "T00:00:00" + timezone;
    return dateTime;
}
function checkDomainSettings() {
    var settings = getDomainSettings();
    if (!settings)
        throw "Settings does not exist";
    if (!settings.project_id)
        throw "Supervisor private key is missing";
    if (!settings.domainAdmin)
        throw "Domain admin is missing";
    if (!settings.domain)
        throw "Domain is missing";
    if (!settings.getMessages && !settings.getEvents)
        throw "Either Poll Gmail or Poll Calendar must be set";
}
function getDomainSettings() {
    var settings = properties.settings;
    if (!settings) {
        var admin = Session.getActiveUser().getEmail();
        var parse = /@(.*)/.exec(admin);
        var domain = parse && parse.length > 1 ? parse[1] : '';
        var before = new Date();
        before.setDate(before.getDate() - 1);
        var after = new Date();
        after.setDate(after.getDate() - 91);
        return {
            internalOnlyEvents: true,
            internalOnlyMessages: true,
            getEvents: true,
            getMessages: false,
            maxScriptRunTimeMin: DEFAULT_MAX_TIME_MINUTES.toString(),
            domainAdmin: admin,
            domain: domain,
            after: after.toLocaleDateString(),
            before: before.toLocaleDateString(),
            startTime: "09:00",
            endTime: "18:00"
        };
    }
    else {
        return settings;
    }
}
/**
 * Clears Progress column for all users
 */
function clearProgressColumn() {
    var _a;
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DOMAIN_USERS);
    (_a = sh) === null || _a === void 0 ? void 0 : _a.getRange(DOMAIN_USERS_SHEET.FIRST_DATA_ROW, DOMAIN_USERS_SHEET.COL_PROGRESS, sh.getLastRow() - DOMAIN_USERS_SHEET.FIRST_DATA_ROW + 1, 1).clearContent();
}
/**
 * Clears URL column for all users
 */
function clearUrlColumn() {
    var _a;
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DOMAIN_USERS);
    (_a = sh) === null || _a === void 0 ? void 0 : _a.getRange(DOMAIN_USERS_SHEET.FIRST_DATA_ROW, DOMAIN_USERS_SHEET.COL_SHEET_URL, sh.getLastRow() - DOMAIN_USERS_SHEET.FIRST_DATA_ROW + 1, 1).clearContent();
}
/**
 * Gets the row range for active users from 'Domain users' tab
 *
 * @param activeUser optional activeUser email. If ommitted, active users is read from activeUsers class
 * @returns the row range or null is there is no active user
 * @throws if active user email is not found in 'Domain users' tab
 *
 */
function getActiveUserRange(activeUser) {
    if (!activeUser) {
        activeUser = activeUsers.activeUser();
    }
    if (activeUser) {
        var shu = SpreadsheetApp.getActive().getSheetByName(SHEET_DOMAIN_USERS);
        var found = shu.getRange(DOMAIN_USERS_SHEET.FIRST_DATA_ROW, 1, shu.getDataRange().getLastRow() - DOMAIN_USERS_SHEET.FIRST_DATA_ROW + 1, 1)
            .createTextFinder(activeUser)
            .matchEntireCell(true)
            .findNext();
        if (found)
            return shu.getRange(found.getRow(), 1, 1, shu.getLastColumn());
        else
            throw "Cannot find user " + activeUser;
    }
    else {
        return null;
    }
}
function getActiveSheet(activeUser) {
    var activeUserRange = getActiveUserRange(activeUser);
    if (activeUserRange) {
        var sheetUrl = activeUserRange.getCell(1, DOMAIN_USERS_SHEET.COL_SHEET_URL).getValue();
        return SpreadsheetApp.openByUrl(sheetUrl);
    }
    else {
        return SpreadsheetApp.getActive();
    }
}
function log(message, error) {
    var _a;
    if (typeof message == 'object') {
        message = JSON.stringify(message);
    }
    message += error ? (error.message + "\n" + error.stack) : "";
    message = message.slice(0, MAX_LOG_MESSAGE_LENGTH);
    Logger.log(message);
    console.log(message);
    try {
        var shLog = (_a = SpreadsheetApp.getActive()) === null || _a === void 0 ? void 0 : _a.getSheetByName(SH_LOG);
        if (shLog) {
            shLog.insertRowBefore(1);
            var data = [new Date(), message];
            shLog.getRange(1, 1, 1, data.length).setValues([data]);
        }
        else {
            console.log(SH_LOG + " does not exist, the message not logged");
        }
    }
    catch (e) {
        Logger.log(e);
        console.log(e);
    }
}
function isValidDate(d) {
    return (d && typeof d.getTime == "function");
}
function getDecimalHours(d) {
    if (typeof d == "string") {
        var hhmm = /0?(\d*):0?(\d*)/.exec(d);
        if (!hhmm || hhmm.length < 3)
            throw 'Invalid time: ' + d;
        return parseInt(hhmm[1]) + parseInt(hhmm[2]) / 60;
    }
    else {
        var hours = d.getHours();
        hours += d.getMinutes() / 60;
        hours += d.getSeconds() / 60 * 60;
        return hours;
    }
}
var Properties = /** @class */ (function () {
    function Properties(propertiesStore) {
        if (propertiesStore === void 0) { propertiesStore = PropertiesService.getUserProperties(); }
        this.KEY_GMAIL = "GMAIL_FLAG";
        this.KEY_ACTIVE_USER = "active_user";
        this.KEY_APP_SETTING = "appSettings";
        this.propertiesStore = propertiesStore;
    }
    Properties.prototype.clearAll = function () {
        this.propertiesStore.deleteAllProperties();
    };
    Properties.prototype.parseJson = function (json) {
        try {
            return JSON.parse(json);
        }
        catch (e) {
            throw 'error parsing JSON (' + e + '): ' + +json;
        }
    };
    Properties.prototype.get = function (key) {
        var _a;
        return _a = this.propertiesStore.getProperty(key), (_a !== null && _a !== void 0 ? _a : '');
    };
    Properties.prototype.set = function (key, value) {
        if (value)
            this.propertiesStore.setProperty(key, value);
        else
            this.propertiesStore.deleteProperty(key);
    };
    Object.defineProperty(Properties.prototype, "settings", {
        get: function () {
            var val = this.get(this.KEY_APP_SETTING);
            return val ? this.parseJson(val) : undefined;
        },
        set: function (value) { this.set(this.KEY_APP_SETTING, JSON.stringify(value)); },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Properties.prototype, "flag", {
        get: function () {
            var _a;
            var val = (_a = this.get(this.KEY_GMAIL), (_a !== null && _a !== void 0 ? _a : ''));
            var flag = parseInt(val);
            return isNaN(flag) ? FLAG.FINISHED : flag;
        },
        set: function (value) { this.set(this.KEY_GMAIL, value.toString()); },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Properties.prototype, "activeUsers", {
        get: function () {
            var usersProp = this.get(this.KEY_ACTIVE_USER);
            return usersProp ? JSON.parse(usersProp) : undefined;
        },
        set: function (activeUsers) { this.set(this.KEY_ACTIVE_USER, JSON.stringify(activeUsers)); },
        enumerable: true,
        configurable: true
    });
    return Properties;
}());
/// <reference path="../Common/Timer.ts" />
/// <reference path="../Common/Users.ts" />
/// <reference path="../Common/Util.ts" />
/// <reference path="../Common/Properties.ts" />
// if set to true, some data will be mocked and execution delayed. Set to true only during development!
var TESTING = false;
// the header name expected in Volume tab to set the attendees list
var ATENDEES_LIST_HEADER = 'Attendees List';
var ATENDEES_LIST_HEADER_CELL = 'D1';
var MAX_THREADS = TESTING ? 2 : 500;
var MAX_MESSAGES = TESTING ? 2 : 500;
var MAX_EVENTS = TESTING ? 500 : 500;
var MAX_USERS = TESTING ? 1 : 500;
var DEFAULT_MAX_TIME_MINUTES = TESTING ? 1 : 4;
var properties = new Properties(PropertiesService.getUserProperties());
var DOMAIN_USERS_SHEET = {
    FIRST_DATA_ROW: 2,
    COL_EMAIL: 1,
    COL_FULL_NAME: 2,
    COL_SHEET_URL: 3,
    COL_PROGRESS: 4,
};
var INPROGRESS_SHEET = {
    // header row
    COL_QUERY: 1,
    COL_NEXT_PAGE_TOKEN: 2,
    COL_STARTING_THREAD_INDEX: 3,
    COL_TIME_ELAPSED: 4,
    COL_RESOURCE: 5,
    COL_STARTING_EVENT_INDEX: 6,
    // other rows
    COL_SENDER: 1,
    COL_RECEIVER: 2,
    COL_COUNT: 3,
};
var SHEET_DOMAIN_USERS = "People";
var SHEET_INPROGRESS = "In Progress";
var FLAG;
(function (FLAG) {
    FLAG[FLAG["SLEEPING"] = 1] = "SLEEPING";
    FLAG[FLAG["RUNNING"] = 2] = "RUNNING";
    FLAG[FLAG["FINISHED"] = 3] = "FINISHED";
})(FLAG || (FLAG = {}));
var activeUsers = new ActiveUsers();
function onOpen() {
    createMenu();
}
function createMissingUserSpreadsheets() {
    new UserSheet().createMissingUserSpreadsheets();
}
function checkUsersTab() {
    var shu = SpreadsheetApp.getActive().getSheetByName(SHEET_DOMAIN_USERS);
    if (!shu) {
        shu = SpreadsheetApp.getActive().insertSheet(SHEET_DOMAIN_USERS, 0);
        shu.getRange("A1:C1").setValues([["Email", "Full Name", "Output"]]);
    }
    var shp = SpreadsheetApp.getActive().getSheetByName(SHEET_INPROGRESS);
    if (!shp) {
        shp = SpreadsheetApp.getActive().insertSheet(SHEET_INPROGRESS);
        shp.hideSheet();
    }
}
function resetState() {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INPROGRESS);
    sh.clearContents();
}
/**
 * Loads all users with combined metrics file set
 */
function loadUsers() {
    var shu = SpreadsheetApp.getActive().getSheetByName(SHEET_DOMAIN_USERS);
    var users = shu.getDataRange().getValues().slice(1)
        .filter(function (row) { return row[2] != ""; })
        .map(function (row) { return row[0].toString(); });
    activeUsers.loadUsers(users);
    return users;
}
function startRunning() {
    try {
        checkDomainSettings();
    }
    catch (e) {
        toast(e);
        settings();
        return;
    }
    var users = loadUsers(); // load all users when starting from the menu
    if (users.length < 1) {
        toast("You need to load a list of domain users into " + SHEET_DOMAIN_USERS + " and setup their Output spreadsheet before running the script.");
        return;
    }
    var flag = getFlag();
    if (flag != FLAG.FINISHED) {
        toast("Script is already running and must be stopped first");
        return;
    }
    toast("Script is now running", 10);
    // make sure everything is cleaned up before starting
    resetState();
    clearProgressColumn();
    try {
        var settings_1 = getDomainSettings();
        new OutputSheet().updateSheetsVisibility(settings_1.getMessages, settings_1.getEvents);
    }
    catch (e) {
        log(e);
    }
    var flag = new Runner().run();
    if (flag == FLAG.SLEEPING) {
        toast("Script is now sleeping", 10);
    }
    else {
        toast("Script is now finished");
    }
}
function stopRunning() {
    var flag = getFlag();
    if (flag == FLAG.FINISHED) {
        toast("The script was not running");
        return;
    }
    if (flag == FLAG.RUNNING) {
        if (SpreadsheetApp.getUi().alert("Do you want to force-stop the app?", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL) == SpreadsheetApp.getUi().Button.OK) {
            setFlag(FLAG.FINISHED);
        }
    }
    while (flag == FLAG.RUNNING) {
        toast("Waiting for current run to finish...");
        Utilities.sleep(20);
        flag = getFlag();
    }
    deleteTrigger();
    setFlag(FLAG.FINISHED);
    clearProgressColumn();
    toast("The script has been stopped");
}
function displayStatus() {
    var msg = "The script is  ";
    var flag = getFlag();
    switch (flag) {
        case FLAG.SLEEPING:
            msg += "sleeping and will start running again shortly";
            break;
        case FLAG.RUNNING:
            msg += "currently running";
            break;
        case FLAG.FINISHED:
            msg += "not running";
    }
    toast(msg, 10);
}
function clearOutputs() {
    var _a, _b, _c;
    var response = SpreadsheetApp.getUi().alert("Heads Up!", "This will delete all files in the User Files folder.\nDo you want to proceed?", SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL);
    if (response != SpreadsheetApp.getUi().Button.YES) {
        return;
    }
    try {
        var appSettings = getDomainSettings();
        var userfiles = (_a = Drive.Files) === null || _a === void 0 ? void 0 : _a.list({
            q: "'" + appSettings.folderId + "' in parents"
        });
        (_c = (_b = userfiles) === null || _b === void 0 ? void 0 : _b.items) === null || _c === void 0 ? void 0 : _c.forEach(function (file) {
            var _a;
            try {
                (_a = Drive.Files) === null || _a === void 0 ? void 0 : _a.remove(file.id);
            }
            catch (e) {
                // if file is not found or other issues, just ignore the exception
                log(e);
            }
        });
    }
    catch (e) {
        // if folder is not found or other issues, just ignore the exception
        log(e);
    }
    clearProgressColumn();
    clearUrlColumn();
    toast("All files in the User Files folder have been deleted", 10);
}
function toast(msg, timeout) {
    if (timeout === void 0) { timeout = -1; }
    SpreadsheetApp.getActive().toast(msg, "Backtrack", timeout);
}
function extractHeader(headers, headerName) {
    var _a, _b;
    var header = (_b = (_a = headers) === null || _a === void 0 ? void 0 : _a.filter(function (h) { return h.name == headerName; }), (_b !== null && _b !== void 0 ? _b : []));
    if (header.length > 0) {
        var a = header[0].value.toLowerCase().match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
        return (a ? a : []);
    }
    else {
        return [];
    }
}
function setFlag(flag) {
    if (flag != FLAG.RUNNING && flag != FLAG.SLEEPING && flag != FLAG.FINISHED)
        throw "Invalid flag <" + flag + ">";
    Utilities.sleep(1500);
    var scriptProperties = PropertiesService.getUserProperties();
    properties.flag = flag;
}
function getFlag() {
    Utilities.sleep(1500);
    return properties.flag;
}
function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
}
function createMenu() {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createAddonMenu();
    try {
        menu
            .addItem("Get domain users", "getDomainUsers")
            .addItem("Create missing user spreadsheets", "createMissingUserSpreadsheets")
            .addSeparator()
            .addItem("Start", "startRunning")
            .addItem("Stop", "stopRunning")
            .addItem("Status", "displayStatus")
            .addSeparator()
            .addItem("Clear Outputs", "clearOutputs")
            .addSeparator();
    }
    catch (e) {
        log(e);
    }
    menu
        .addItem("Settings", "settings")
        .addToUi();
}
var OutputSheet = /** @class */ (function () {
    function OutputSheet(ss) {
        if (ss === void 0) { ss = getActiveSheet(); }
        this.ss = ss;
    }
    OutputSheet.appendCounts = function (values, counts) {
        for (var sender in counts) {
            for (var receiver in counts[sender]) {
                values.push([sender, receiver, counts[sender][receiver], "", "", ""]);
            }
        }
    };
    OutputSheet.prototype.saveGmail = function (state) {
        var sheet = this.ss.getSheetByName(OutputSheet.SHEET_GMAIL);
        if (!sheet) {
            return;
        }
        var values = [];
        OutputSheet.appendCounts(values, state.counts);
        if (values.length > 0) {
            sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
        }
        if (sheet.getLastRow() > values.length + 1)
            sheet.getRange(values.length + 2, 1, sheet.getLastRow() - values.length - 1, 3).clear();
        resetState();
    };
    OutputSheet.prototype.saveDaily = function (daily, start, end, startWorkday, endWorkday) {
        var sheet = this.ss.getSheetByName(OutputSheet.SHEET_DAILY);
        if (!sheet) {
            return;
        }
        var values = [];
        for (var i = start; i <= end; i += (24 * 60 * 60 * 1000)) {
            var date = Utilities.formatDate(new Date(i), Session.getScriptTimeZone(), "yyyyMMdd");
            if (daily[date]) {
                daily[date].meetings.sort(function (a, b) { return a.start > b.start ? 1 : -1; });
                var begin = startWorkday, max = 0, lastEnd = startWorkday;
                for (var j = 0; j < daily[date].meetings.length; j++) {
                    var meeting = daily[date].meetings[j];
                    if (meeting.start > endWorkday || meeting.end < startWorkday)
                        continue;
                    if (meeting.start - begin > max)
                        max = meeting.start - begin;
                    begin = meeting.end;
                    if (meeting.end > lastEnd)
                        lastEnd = meeting.end;
                }
                if (endWorkday - lastEnd > max)
                    max = endWorkday - lastEnd;
                values.push([date, daily[date].busy, max]);
            }
            else {
                values.push([date, 0, endWorkday - startWorkday]);
            }
        }
        if (values && values.length)
            sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
        if (sheet.getLastRow() > values.length + 1)
            sheet.getRange(values.length + 2, 1, sheet.getLastRow() - values.length - 1, 3).clearContent();
    };
    OutputSheet.prototype.savePeople = function (people, everyone, domain) {
        if (domain === void 0) { domain = getDomainSettings().domain; }
        var sheet = this.ss.getSheetByName(OutputSheet.SHEET_PEOPLE);
        if (!sheet) {
            return;
        }
        var values = [[""]];
        var all = Object.keys(everyone).filter(function (email) { return domain && email.indexOf(domain) >= 0; });
        sheet.clearContents();
        for (var i = 0; i < all.length; i++) {
            values[i + 1] = [all[i]];
            for (var j = 0; j < all.length; j++) {
                if (!i)
                    values[0].push(all[j]);
                values[i + 1][j + 1] = "";
                if (people[all[i]] && people[all[i]][all[j]]) {
                    values[i + 1][j + 1] = people[all[i]][all[j]].toString();
                }
            }
            if (values && values.length)
                sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
        }
        if (sheet.getLastRow() > values.length) {
            sheet.getRange(values.length + 1, 1, sheet.getLastRow() - values.length, sheet.getLastColumn()).clearContent();
        }
    };
    OutputSheet.prototype.saveCalendar = function (counts) {
        var sheet = this.ss.getSheetByName(OutputSheet.SHEET_CALENDAR);
        if (!sheet) {
            return;
        }
        var values = [];
        OutputSheet.appendCounts(values, counts);
        if (values && values.length)
            sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
        if (sheet.getLastRow() > values.length + 1)
            sheet.getRange(values.length + 2, 1, sheet.getLastRow() - values.length - 1, 3).clearContent();
    };
    OutputSheet.prototype.saveVolume = function (volumes) {
        var sheet = this.ss.getSheetByName(OutputSheet.SHEET_VOLUME);
        if (!sheet) {
            return;
        }
        var setAttendees = sheet.getRange(ATENDEES_LIST_HEADER_CELL).getValue() == ATENDEES_LIST_HEADER;
        var rows = volumes.map(function (vol) {
            return [
                vol.atendees.length,
                Utilities.formatDate(vol.date, Session.getScriptTimeZone(), "M/d/yyyy"),
                vol.duration,
                setAttendees ? anonymize(vol.atendees).join(',') : ''
            ];
        });
        if (rows.length)
            sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
        if (sheet.getLastRow() > rows.length + 1)
            sheet.getRange(rows.length + 2, 1, sheet.getLastRow() - rows.length - 1, 1).clearContent();
    };
    OutputSheet.prototype.saveRecent = function (recent) {
        var ss = getActiveSheet();
        var sheet = ss.getSheetByName(OutputSheet.SHEET_RECENT);
        if (!sheet) {
            return;
        }
        sheet.getRange("A2:C").clearContent();
        var rows = Object.keys(recent)
            .map(function (personB) {
            var email = personB;
            var count = recent[personB].length;
            var mostRecentMeeting = recent[personB].sort(function (d1, d2) { return d2.getTime() - d1.getTime(); })[0];
            return {
                email: email,
                count: count,
                mostRecentMeeting: mostRecentMeeting
            };
        })
            .sort(function (a, b) { return b.mostRecentMeeting.getTime() - a.mostRecentMeeting.getTime(); })
            .map(function (item) { return [item.email, item.count, item.mostRecentMeeting]; });
        if (rows.length > 0) {
            sheet.insertRows(2);
            sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
        }
    };
    /**
     * Hides or unhides the individual sheets related to gmail or calendar, for all user spreadsheets
     * Note: If showGmail and showCalendar are both set to true the method will throw an exception!
     *
     * @param showGmail show all gmail related sheets
     * @param showCalendar show all calendar related sheets
     *
     *
     */
    OutputSheet.prototype.updateSheetsVisibility = function (showGmail, showCalendar) {
        if (!showCalendar && !showGmail) {
            throw new Error('Cannot hide all sheets, either gmail or calendar items must be visible!');
        }
        var shu = SpreadsheetApp.getActive().getSheetByName(SHEET_DOMAIN_USERS);
        var spreadsheets = shu.getDataRange().getValues().slice(DOMAIN_USERS_SHEET.FIRST_DATA_ROW - 1)
            .filter(function (row) { return row[DOMAIN_USERS_SHEET.COL_SHEET_URL - 1]; })
            .map(function (row) { return row[DOMAIN_USERS_SHEET.COL_SHEET_URL - 1]; });
        var sheetRules = [
            { type: 'gmail', sheet: OutputSheet.SHEET_GMAIL },
            { type: 'calendar', sheet: OutputSheet.SHEET_CALENDAR },
            { type: 'calendar', sheet: OutputSheet.SHEET_DAILY },
            { type: 'calendar', sheet: OutputSheet.SHEET_PEOPLE },
            { type: 'calendar', sheet: OutputSheet.SHEET_VOLUME },
        ];
        spreadsheets.forEach(function (spreadsheet) {
            var ss = SpreadsheetApp.openByUrl(spreadsheet);
            sheetRules.forEach(function (rule) {
                var _a, _b;
                if (rule.type == 'gmail' && showGmail || rule.type == 'calendar' && showCalendar)
                    (_a = ss.getSheetByName(rule.sheet)) === null || _a === void 0 ? void 0 : _a.showSheet();
                else
                    (_b = ss.getSheetByName(rule.sheet)) === null || _b === void 0 ? void 0 : _b.hideSheet();
            });
        });
    };
    OutputSheet.SHEET_PEOPLE = "People";
    OutputSheet.SHEET_DAILY = "Daily";
    OutputSheet.SHEET_VOLUME = "Volume";
    OutputSheet.SHEET_GMAIL = "Gmail";
    OutputSheet.SHEET_RECENT = "Recent";
    OutputSheet.SHEET_CALENDAR = "Calendar";
    return OutputSheet;
}());
function settings() {
    checkUsersTab();
    var htmlTemplate = HtmlService.createTemplateFromFile('settings');
    var settings = getDomainSettings();
    console.log(JSON.stringify(settings));
    htmlTemplate.before = Utilities.formatDate(new Date(settings.before), 'GMT', 'yyyy-MM-dd');
    htmlTemplate.after = Utilities.formatDate(new Date(settings.after), 'GMT', 'yyyy-MM-dd');
    htmlTemplate.settings = JSON.stringify(settings);
    SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setWidth(500).setHeight(500), 'App Settings');
}
function clearSettings() {
    properties.activeUsers = undefined;
    createMenu();
}
function processSettings(form) {
    form.internalOnlyEvents = form.internalOnly == 'true';
    form.internalOnlyMessages = form.internalOnly == 'true';
    form.getEvents = (form.getEvents + '') == 'true';
    form.getMessages = (form.getMessages + '') == 'true';
    form.before = getDateInTimezone(form.before.toString());
    form.after = getDateInTimezone(form.after.toString());
    properties.settings = form;
    createMenu(); // show changes, if settings are ok
}
/// <reference path="../Common/Util.ts" />
var Runner = /** @class */ (function (_super) {
    __extends(Runner, _super);
    function Runner() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Runner.prototype.runGmail = function (userEmail, timer, state) {
        log('runGmail should not have ben caled for this app');
        return FLAG.FINISHED;
    };
    return Runner;
}(BaseRunner));
/// <reference path="../Common/BaseUserSheet.ts" />
var UserSheet = /** @class */ (function (_super) {
    __extends(UserSheet, _super);
    function UserSheet() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.SHEET_TEMPLATE = 'Recent-TEMPLATE';
        return _this;
    }
    /**
     *
    * @param userEmail Creates a new user spreadsheet
    * @returns spreadsheet URL
    */
    UserSheet.prototype.createFile = function (userEmail) {
        var _a;
        var title = 'Output for ' + userEmail;
        var newSS = SpreadsheetApp.create(title);
        (_a = SpreadsheetApp.getActive().getSheetByName(this.SHEET_TEMPLATE)) === null || _a === void 0 ? void 0 : _a.copyTo(newSS).setName(OutputSheet.SHEET_RECENT).showSheet();
        newSS.deleteSheet(newSS.getSheetByName('Sheet1'));
        return newSS.getUrl();
    };
    return UserSheet;
}(BaseUserSheet));
/// <reference path="../Common/Mock.ts" />
var GmailQuery = /** @class */ (function () {
    function GmailQuery(settings, activeUser) {
        this.settings = settings;
        this.activeUser = activeUser;
        this.before = new Date(settings.before);
        this.after = new Date(settings.after);
    }
    GmailQuery.prototype.toString = function () {
        var _a;
        var q = this.settings.internalOnlyMessages ? '((-to:' + this.activeUser + ' to:' + this.settings.domain + ') OR (-from:' + this.activeUser + ' from:' + this.settings.domain + ')) ' : '';
        return q + "before:" + Math.floor((_a = this.before, (_a !== null && _a !== void 0 ? _a : new Date())).getTime() / 1000) + (this.after ? " after:" + Math.floor(this.after.getTime() / 1000) : "");
    };
    return GmailQuery;
}());
function runGmail(userEmail, timer, state) {
    var _a, _b, _c, _d, _e, _f, _g;
    setFlag(FLAG.RUNNING);
    timer.setTask('Gmail');
    var allDone = false;
    if (!state.query)
        state.query = getGmailQuery();
    try {
        while (!timer.isTimerExpired() && !allDone) {
            timer.displayElapsedTime();
            log('fetch threads from Gmail API');
            var threads = TESTING ?
                getMockThreads(100, userEmail, MAX_THREADS, state.query, state.nextPageToken) :
                getThreads(userEmail, MAX_THREADS, state.query, state.nextPageToken);
            log('thread count: ' + ((_b = (_a = threads) === null || _a === void 0 ? void 0 : _a.threads) === null || _b === void 0 ? void 0 : _b.length));
            var i = 0;
            if ((_d = (_c = threads) === null || _c === void 0 ? void 0 : _c.threads) === null || _d === void 0 ? void 0 : _d.length) {
                for (i = (_e = state.startingThreadIndex, (_e !== null && _e !== void 0 ? _e : 0)); i < threads.threads.length && !timer.isTimerExpired(); i++) {
                    timer.displayElapsedTime();
                    var thread = threads.threads[i];
                    log('fetch messages for thread id ' + thread.id);
                    var messages = getMessages(userEmail, MAX_MESSAGES, thread.id);
                    log('message count: ' + ((_f = messages) === null || _f === void 0 ? void 0 : _f.length));
                    if (((_g = messages) === null || _g === void 0 ? void 0 : _g.length) > 0) {
                        for (var j = 0; j < messages.length; j++) {
                            var message = messages[j];
                            var senders = extractHeader(message.payload.headers, "From");
                            if (!senders || !senders.length)
                                continue;
                            var receivers = extractHeader(message.payload.headers, "To");
                            receivers = receivers.concat(extractHeader(message.payload.headers, "Cc"));
                            receivers = receivers.concat(extractHeader(message.payload.headers, "Bcc"));
                            if (!receivers || !receivers.length)
                                continue;
                            for (var k = 0; k < receivers.length; k++) {
                                incrementCount(state.counts, senders[0], receivers[k]);
                            }
                            timer.displayElapsedTime();
                        }
                    }
                }
                if (timer.isTimerExpired()) {
                    state.startingThreadIndex = i;
                }
                else {
                    state.nextPageToken = threads.nextPageToken;
                    state.startingThreadIndex = 0;
                    allDone = !threads.nextPageToken;
                }
            }
            else {
                allDone = true;
            }
        }
        log("all done: " + allDone);
        state.timeElapsed = timer.totalElapsedTime;
        var flag = void 0;
        if (allDone) {
            new OutputSheet().saveGmail(state);
            flag = FLAG.FINISHED;
        }
        else {
            saveState(state);
            createTrigger();
            flag = FLAG.SLEEPING;
        }
        log(state);
        return flag;
    }
    catch (e) {
        log(e);
        return FLAG.FINISHED;
    }
}
function incrementCount(counts, sender, receiver, count, domain) {
    if (count === void 0) { count = 1; }
    if (domain === void 0) { domain = getDomainSettings().domain; }
    // ignore this pair if either sender or receiver email is outside of the domain
    if (sender.indexOf((domain !== null && domain !== void 0 ? domain : '')) < 0 || receiver.indexOf((domain !== null && domain !== void 0 ? domain : '')) < 0) {
        return;
    }
    if (!counts[sender])
        counts[sender] = {};
    if (!counts[sender][receiver])
        counts[sender][receiver] = 0;
    counts[sender][receiver] += count;
}
function getGmailQuery() {
    var appSettings = getDomainSettings();
    if (!appSettings.before || !appSettings.after)
        throw 'Before/after date not set';
    var query = new GmailQuery(appSettings, activeUsers.activeUser());
    return query.toString();
}
