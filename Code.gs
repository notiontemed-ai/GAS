const SHEET_SCHEDULE_SOURCE = 'График';
const SHEET_DOCTOR_SCHEDULE = 'Обработанный график';
const SHEET_EMPLOYEES = 'Сотрудники';
const SHEET_REQUESTS = 'Заявки';
const SHEET_PROCESSED_REQUESTS = 'Обработанные заявки';
const SHEET_REPORT_LOG = 'Лог';
const SHEET_EXCLUSIONS = 'Исключения';
const MENU_NAME = 'TEMED';
const DEBUG_DOCTOR_NAME = 'Шепелев Александр Анатольевич';
const DEBUG_MONTH = '2025-11';

const TIME_TYPE_MAP = {
  'Рабочее время': 'Приемные часы',
  'ТОЛЬКО ДЛЯ ПП': 'Приемные часы',
  'Нет приема, курсовики !': 'Процедурные часы',
  'Обед': 'Обед',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('Обработать исходные данные', 'processSourceData')
    .addItem('Сформировать отчёт загрузки врачей', 'buildDoctorLoadReport')
    .addItem('Сформировать план загрузки 3/7 дней', 'buildForwardLoadReport')
    .addItem('Сформировать план загрузки 3/7 дней по всем клиникам', 'buildForwardLoadReportAllBranches')
    .addItem('Отправить рассылку по всем клиникам (после пакетного плана)', 'confirmAndSendForwardBatchEmails')
    .addItem('Собрать общий план', 'buildOverallForecastReport')
    .addItem('Архивировать план (ежедневно)', 'archivePlanReportsDaily')
    .addToUi();

  if (typeof maybePromptForwardBatchEmail_ === 'function') {
    maybePromptForwardBatchEmail_();
  }
}

function processSourceData() {
  buildDoctorSchedule();
  SpreadsheetApp.getActive().toast('Исходные данные (график + заявки) обработаны', MENU_NAME, 5);
}

function buildDoctorSchedule() {
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName(SHEET_SCHEDULE_SOURCE);
  if (!source) {
    throw new Error(`Не найден лист "${SHEET_SCHEDULE_SOURCE}"`);
  }

  const rows = getSheetObjects_(source);
  const requestsSheet = ss.getSheetByName(SHEET_REQUESTS);
  const requestRows = requestsSheet ? getSheetObjects_(requestsSheet) : [];
  const excludedEmployees = getExcludedEmployeesSet_(ss);
  const employeeRegistry = getEmployeeRegistry_(ss);
  const currentDay = detectCurrentDayFromRawRequests_(requestRows, employeeRegistry.allowed, excludedEmployees, employeeRegistry.unassignedFound);
  if (!currentDay) throw new Error('Не найдены выполненные заявки для определения актуальной даты');
  const rangeStart = addDays(currentDay, -42);
  const rangeEnd = addDays(currentDay, 7);
  const requestsRangeEnd = addDays(currentDay, 8);

  const grouped = new Map();
  let sourceGraphMaxCreatedAt = new Date(0);

  rows.forEach((row, idx) => {
    const employee = str_(row['Сотрудник']);
    const cabinet = str_(row['Кабинет']);
    const cabinetType = detectCabinetType_(cabinet);
    const isCabinetRow = Boolean(cabinetType);
    if (!isCabinetRow) {
      if (!employee || excludedEmployees.has(employee) || hasStandaloneTestWord_(employee)) return;
      if (!employeeRegistry.allowed.has(employee)) {
        employeeRegistry.unassignedFound.add(employee);
        return;
      }
    } else if (employee && (excludedEmployees.has(employee) || hasStandaloneTestWord_(employee))) {
      return;
    }

    const date = toDate_(row['Дата']);
    if (!isDateInRangeInclusive_(date, rangeStart, rangeEnd)) return;
    const dateLabel = formatDateDDMMYYYY_(date || row['Дата']);
    const branch = str_(row['Филиал']);
    const start = parseTimeToMinutes_(row['Время начала']);
    const end = parseTimeToMinutes_(row['Время окончания']);
    const timeKind = str_(row['Вид времени']);
    const sourceCreatedAt = toDate_(row['Дата создания']);
    if (sourceCreatedAt && sourceCreatedAt > sourceGraphMaxCreatedAt) sourceGraphMaxCreatedAt = sourceCreatedAt;

    const createdAtRaw = row['Дата создания'];
    const createdAt = toDate_(createdAtRaw) || new Date(0);
    const createdAtSort = sortableDateTime_(createdAtRaw) || createdAt.getTime();
    const createdAtLabel = formatDateTimeDDMMYYYY_(createdAtRaw) || formatDateTimeDDMMYYYY_(createdAt);

    if (!date || !dateLabel || !branch || start === null || end === null || !timeKind) return;
    if (end <= start) return;

    const key = [dateKey_(date), branch, employee, cabinet].join('|');
    if (!grouped.has(key)) grouped.set(key, []);

    grouped.get(key).push({
      date,
      dateLabel,
      branch,
      employee,
      cabinet,
      cabinetType,
      start,
      end,
      timeKind,
      timeType: mapTimeType_(timeKind),
      createdAt,
      createdAtSort,
      createdAtLabel,
      eventId: str_(row['Идентификатор события']),
      rowOrder: idx,
    });
  });

  const resolved = [];
  grouped.forEach((items) => {
    resolved.push(...resolveIntervals_(items));
  });

  resolved.sort((a, b) => {
    if (a.branch !== b.branch) return a.branch.localeCompare(b.branch, 'ru');
    const cmpDate = a.date.getTime() - b.date.getTime();
    if (cmpDate !== 0) return cmpDate;
    if (a.employee !== b.employee) return a.employee.localeCompare(b.employee, 'ru');
    return a.start - b.start;
  });

  const compacted = mergeAdjacentSlotsByType_(resolved);

  const target = getOrCreateSheet_(ss, SHEET_DOCTOR_SCHEDULE);
  target.clearContents();
  const header = [
    'Дата', 'Филиал', 'Сотрудник', 'Кабинет', 'Время начала', 'Время окончания', 'Вид времени', 'Тип времени', 'Дата создания', 'Идентификатор события',
  ];
  target.getRange(1, 1, 1, header.length).setValues([header]);

  const now = new Date();
  const minMarkerTime = sourceGraphMaxCreatedAt.getTime() > 0 ? sourceGraphMaxCreatedAt.getTime() + 1000 : 0;
  const markerDate = new Date(Math.max(now.getTime(), minMarkerTime));
  const markerRow = ['__SYSTEM_TIMESTAMP__', '', '', '', '', '', '', '', formatDateTimeDDMMYYYY_(markerDate), '__SYSTEM_TIMESTAMP__'];

  if (compacted.length) {
    const values = compacted.map((r) => [
      r.dateLabel, r.branch, r.employee, r.cabinet || '', minutesToTimeString_(r.start), minutesToTimeString_(r.end), r.timeKind, r.timeType, r.createdAtLabel, r.eventId,
    ]);
    values.push(markerRow);
    target.getRange(2, 1, values.length, header.length).setValues(values);
  } else {
    target.getRange(2, 1, 1, header.length).setValues([markerRow]);
  }

  buildProcessedRequestsSheet_(ss, requestsSheet, employeeRegistry.allowed, excludedEmployees, rangeStart, requestsRangeEnd, currentDay, employeeRegistry.unassignedFound);
  saveUnassignedEmployees_(ss, employeeRegistry.unassignedFound);
  SpreadsheetApp.getActive().toast('Обработанные данные сформированы', MENU_NAME, 5);
}

function buildProcessedRequestsSheet_(ss, requestsSheet, allowedEmployees, excludedEmployees, rangeStart, rangeEnd, currentDay, unassignedSet) {
  if (!requestsSheet) return;
  const values = requestsSheet.getDataRange().getValues();
  if (!values.length) return;

  const headers = values[0].map((h) => str_(h));
  const idxStatus = headers.indexOf('Состояние');
  const idxEmployee = headers.indexOf('Сотрудник');
  const idxClient = headers.indexOf('Клиент');
  const idxStart = headers.indexOf('ДатаНачала');
  const idxEnd = headers.indexOf('ДатаОкончания');
  const idxBranch = headers.indexOf('Филиал');

  const extraHeaders = ['НачалоМин', 'ОкончаниеМин', 'ДлительностьМин'];
  const outHeaders = [...headers];
  extraHeaders.forEach((h) => {
    if (outHeaders.indexOf(h) < 0) outHeaders.push(h);
  });

  const idxCabinet = headers.indexOf('Кабинет');

  const rows = values.slice(1).filter((row) => {
    const employee = str_(row[idxEmployee]);
    const client = str_(row[idxClient]);
    const status = str_(row[idxStatus]);
    const startDate = toDate_(row[idxStart]);
    const cabinetType = detectCabinetType_(idxCabinet >= 0 ? row[idxCabinet] : '');
    const isCabinetRequest = Boolean(cabinetType);

    if (!isCabinetRequest) {
      if (!employee || (excludedEmployees && excludedEmployees.has(employee)) || hasStandaloneTestWord_(employee) || hasStandaloneTestWord_(client)) return false;
      if (allowedEmployees && !allowedEmployees.has(employee)) {
        if (unassignedSet) unassignedSet.add(employee);
        return false;
      }
    } else {
      if (employee && ((excludedEmployees && excludedEmployees.has(employee)) || hasStandaloneTestWord_(employee))) return false;
      if (hasStandaloneTestWord_(client)) return false;
    }
    if (!isAllowedProcessedRequestStatus(status)) return false;
    if (isPlannedStatus(status) && startDate && currentDay && startDate < currentDay) return false;
    if (!isDateInRangeInclusive_(startDate, rangeStart, rangeEnd)) return false;
    return true;
  });

  rows.sort((a, b) => {
    const ba = str_(a[idxBranch]);
    const bb = str_(b[idxBranch]);
    if (ba !== bb) return ba.localeCompare(bb, 'ru');
    const da = toDate_(a[idxStart]);
    const db = toDate_(b[idxStart]);
    return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
  });

  const outRows = rows.map((row) => {
    const startDate = toDate_(row[idxStart]);
    const endDate = toDate_(row[idxEnd]);
    const timing = computeRequestTiming_(startDate, endDate);
    const item = row.slice();
    item[outHeaders.indexOf('НачалоМин')] = timing.startMin;
    item[outHeaders.indexOf('ОкончаниеМин')] = timing.endMin;
    item[outHeaders.indexOf('ДлительностьМин')] = timing.durationMin;
    return item;
  });

  const target = getOrCreateSheet_(ss, SHEET_PROCESSED_REQUESTS);
  target.clearContents();
  target.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);
  if (outRows.length) {
    target.getRange(2, 1, outRows.length, outHeaders.length).setValues(outRows);
  }
}


function buildDoctorLoadReport() {
  const ss = SpreadsheetApp.getActive();
  const runLog = [];
  const addLog = (level, message, details) => {
    runLog.push([new Date(), level, message, details || '']);
  };

  addLog('INFO', 'Старт формирования отчета загрузки врачей');

  const graphSheet = ss.getSheetByName(SHEET_SCHEDULE_SOURCE);
  const requestsSheet = ss.getSheetByName(SHEET_REQUESTS);
  const doctorScheduleSheet = ss.getSheetByName(SHEET_DOCTOR_SCHEDULE);

  if (!graphSheet || !requestsSheet || !doctorScheduleSheet) {
    addLog('ERROR', 'Не найдены обязательные листы', 'Требуются: График, Заявки, Обработанный график');
    writeReportBuildLog_(ss, runLog);
    throw new Error('Проверьте наличие листов: График, Заявки, Обработанный график');
  }

  const graphRows = getSheetObjects_(graphSheet);
  const requestRows = getSheetObjects_(requestsSheet);
  const scheduleRows = getSheetObjectsDisplay_(doctorScheduleSheet);
  addLog('INFO', 'Загружены строки источников', `График: ${graphRows.length}, Заявки: ${requestRows.length}, Обработанный график: ${scheduleRows.length}`);

  const excludedEmployees = getExcludedEmployeesSet_(ss);
  const employeeRegistry = getEmployeeRegistry_(ss);
  const latestByStatus = buildLatestRequests(requestRows, excludedEmployees, employeeRegistry.allowed, employeeRegistry.unassignedFound);
  const currentDay = detectCurrentDay(latestByStatus);
  if (!currentDay) throw new Error('Не найдены выполненные заявки для определения текущего дня');
  const rangeStart = addDays(currentDay, -42);
  const rangeEnd = addDays(currentDay, 7);
  logDebugDoctorNovember_(scheduleRows, requestRows, addLog, excludedEmployees);

  const graphMonths = new Set();
  graphRows.forEach((row) => {
    const date = toDate_(row['Дата']);
    const employee = str_(row['Сотрудник']);
    if (!date || !isDateInRangeInclusive_(date, rangeStart, rangeEnd)) return;
    if (!employee || !employeeRegistry.allowed.has(employee)) return;
    graphMonths.add(monthKey_(date));
  });

  const completedMonths = new Set();
  graphMonths.forEach((month) => {
    const [year, mon] = month.split('-').map(Number);
    const monthEnd = new Date(year, mon, 0, 0, 0, 0, 0);
    const hasCompleted = requestRows.some((row) => {
      const status = str_(row['Состояние']);
      const start = toDate_(row['ДатаНачала']);
      return isCompletedStatus(status) && start && start >= monthEnd;
    });
    if (hasCompleted) completedMonths.add(month);
  });

  addLog('INFO', 'Определены завершенные месяцы', completedMonths.size ? [...completedMonths].sort().join(', ') : 'Нет завершенных месяцев');

  const scheduleByMonth = buildScheduleIndex_(scheduleRows);

  [...completedMonths].sort().forEach((month) => {
    try {
      const result = buildMonthReport_(month, requestRows, scheduleByMonth.get(month) || [], excludedEmployees, employeeRegistry.allowed, rangeStart, rangeEnd, employeeRegistry.unassignedFound);
      renderMonthReport_(ss, month, result);
      addLog('INFO', `Сформирован лист отчета ${month}`, `Строк филиалы: ${result.byBranch.length}, врачи: ${result.byDoctor.length}, ошибок: ${result.logs.length}`);
    } catch (error) {
      addLog('ERROR', `Ошибка формирования листа ${month}`, error && error.stack ? error.stack : String(error));
    }
  });

  saveUnassignedEmployees_(ss, employeeRegistry.unassignedFound);
  writeReportBuildLog_(ss, runLog);

  SpreadsheetApp.getActive().toast('Отчеты по завершенным месяцам сформированы', MENU_NAME, 5);
}

function writeReportBuildLog_(ss, rows) {
  const sheet = getOrCreateSheet_(ss, SHEET_REPORT_LOG);
  sheet.clearContents();

  const header = ['Время', 'Уровень', 'Сообщение', 'Детали'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('dd.mm.yyyy hh:mm:ss');
  }

  sheet.autoResizeColumns(1, header.length);
}

function logDebugDoctorNovember_(scheduleRows, requestRows, addLog, excludedEmployees) {
  const scheduleDebugRows = scheduleRows.filter((row) => {
    const employee = str_(row['Сотрудник']);
    const date = toDate_(row['Дата']);
    return employee === DEBUG_DOCTOR_NAME && monthKey_(date) === DEBUG_MONTH;
  });

  const scheduleMinutes = scheduleDebugRows.reduce((acc, row) => {
    if (str_(row['Тип времени']) !== 'Приемные часы') return acc;
    const start = parseTimeToMinutes_(row['Время начала']);
    const end = parseTimeToMinutes_(row['Время окончания']);
    if (start === null || end === null || end <= start) return acc;
    return acc + (end - start);
  }, 0);

  addLog('INFO', 'DEBUG Шепелев ноябрь: строки графика врачей', String(scheduleDebugRows.length));
  addLog('INFO', 'DEBUG Шепелев ноябрь: приемные часы (ч)', String(round1_(scheduleMinutes / 60)));

  const details = scheduleDebugRows
    .sort((a, b) => {
      const d = toDate_(a['Дата']) - toDate_(b['Дата']);
      if (d !== 0) return d;
      return parseTimeToMinutes_(a['Время начала']) - parseTimeToMinutes_(b['Время начала']);
    })
    .slice(0, 40)
    .map((row) => {
      const date = formatDateDDMMYYYY_(row['Дата']);
      return `${date} | ${str_(row['Филиал'])} | ${formatTimeCell_(row['Время начала'])}-${formatTimeCell_(row['Время окончания'])} | ${str_(row['Тип времени'])}`;
    })
    .join('\n');

  addLog('INFO', 'DEBUG Шепелев ноябрь: интервалы графика врачей', details || 'Нет строк');

  const requestDebugRows = requestRows
    .map((r) => normalizeRequest_(r, excludedEmployees, allowedEmployees, rangeStart, rangeEnd, unassignedEmployees))
    .filter((r) => r && r.doctor === DEBUG_DOCTOR_NAME && r.month === DEBUG_MONTH);

  const requestHours = requestDebugRows.reduce((acc, r) => acc + r.durationMin, 0) / 60;
  addLog('INFO', 'DEBUG Шепелев ноябрь: выполненные заявки (шт)', String(requestDebugRows.length));
  addLog('INFO', 'DEBUG Шепелев ноябрь: выполненные заявки (ч)', String(round1_(requestHours)));
}

function buildMonthReport_(month, requestRows, monthScheduleRows, excludedEmployees, allowedEmployees, rangeStart, rangeEnd, unassignedEmployees) {
  const logs = [];
  const scheduleIndex = new Map();
  const metricsBranch = new Map();
  const metricsDoctor = new Map();

  monthScheduleRows.forEach((s) => {
    const dateKey = dateKey_(toDate_(s['Дата']));
    const branch = str_(s['Филиал']);
    const doctor = str_(s['Сотрудник']);
    if (excludedEmployees.has(doctor) || hasStandaloneTestWord_(doctor)) return;
    const start = parseTimeToMinutes_(s['Время начала']);
    const end = parseTimeToMinutes_(s['Время окончания']);
    const type = str_(s['Тип времени']);
    if (!dateKey || !branch || !doctor || start === null || end === null || end <= start) return;

    const key = [dateKey, branch, doctor].join('|');
    if (!scheduleIndex.has(key)) scheduleIndex.set(key, []);
    scheduleIndex.get(key).push({ start, end, type });

    accumulateTime_(metricsBranch, branch, type, end - start);
    accumulateTime_(metricsDoctor, `${branch}|${doctor}`, type, end - start, branch, doctor);
  });

  scheduleIndex.forEach((arr) => arr.sort((a, b) => a.start - b.start));

  const requests = requestRows
    .map((r) => normalizeRequest_(r, excludedEmployees, allowedEmployees, rangeStart, rangeEnd, unassignedEmployees))
    .filter((r) => r && r.month === month);

  const reqByDayDoctor = new Map();

  requests.forEach((req) => {
    const overlapKey = `${req.dateKey}|${req.branch}|${req.doctor}`;
    if (!reqByDayDoctor.has(overlapKey)) reqByDayDoctor.set(overlapKey, []);
    reqByDayDoctor.get(overlapKey).push(req);
  });

  reqByDayDoctor.forEach((arr) => {
    arr.sort((a, b) => a.startMin - b.startMin);
    let prev = null;
    arr.forEach((req) => {
      if (prev && req.startMin < prev.endMin) {
        logs.push(logRow_(req.month, 'Пересечение выполненных заявок', req.branch, req.doctor, req.date, req.id, 'Перекрытие с другой выполненной заявкой'));
        req.blockedByOverlap = true;
      }
      if (!prev || req.endMin > prev.endMin) prev = req;
    });
  });

  requests.forEach((req) => {
    if (req.blockedByOverlap) return;

    const key = [req.dateKey, req.branch, req.doctor].join('|');
    const dayIntervals = scheduleIndex.get(key) || [];
    const receptionIntervals = mergeIntervals_(dayIntervals.filter((i) => i.type === 'Приемные часы'));

    if (!receptionIntervals.length) {
      logs.push(logRow_(req.month, 'Нет приемных часов на дату', req.branch, req.doctor, req.date, req.id, 'Заявка учтена как факт выполненного приема, несмотря на отсутствие приемных часов в графике'));
    } else {
      const fits = receptionIntervals.some((it) => req.startMin >= it.start && req.endMin <= it.end);
      if (!fits) {
        logs.push(logRow_(req.month, 'Заявка вне приемных часов', req.branch, req.doctor, req.date, req.id, 'Заявка учтена как факт выполненного приема, несмотря на неполное попадание в приемные часы'));
      }
    }

    accumulateReception_(metricsBranch, req.branch, req.type, req.durationMin);
    accumulateReception_(metricsDoctor, `${req.branch}|${req.doctor}`, req.type, req.durationMin, req.branch, req.doctor);

    if (req.durationFallback) {
      logs.push(logRow_(req.month, 'Не распознана длительность', req.branch, req.doctor, req.date, req.id, 'Длительность не найдена в тексте, использовано 30 минут'));
    }
  });

  const procedureRequests = requestRows
    .map((r) => normalizeProcedureRequest_(r, excludedEmployees, allowedEmployees, rangeStart, rangeEnd, unassignedEmployees))
    .filter((r) => r && r.month === month);

  allocateProcedureMinutes_(procedureRequests, scheduleIndex, metricsBranch, metricsDoctor);

  return {
    byBranch: toMetricsRows_(metricsBranch),
    byDoctor: toMetricsRows_(metricsDoctor),
    logs,
  };
}

function renderMonthReport_(ss, month, result) {
  const sheet = getOrCreateSheet_(ss, month);
  sheet.clearContents();

  let row = 1;
  const title1 = 'Таблица 1 — Свод по филиалам';
  sheet.getRange(row++, 1).setValue(title1);

  const header = [
    'Филиал',
    'Врач',
    'Приемные часы (ч)',
    'Процедурные часы (ч)',
    'Обед (ч)',
    'Прочие часы (ч)',
    'Приёмы (ч)',
    'Первичные (ч)',
    'Повторные (ч)',
    'Прочие приёмы (ч)',
    'Процедуры (ч)',
    'Процедуры (шт)',
    'Утилизация приёмы (%)',
    'Утилизация процедурных часов (%)',
  ];

  sheet.getRange(row, 1, 1, header.length).setValues([header]);
  row++;

  row = renderMetricsTable_(sheet, row, result.byBranch, false);
  row += 1;

  sheet.getRange(row++, 1).setValue('Таблица 2 — Свод по врачам');
  sheet.getRange(row, 1, 1, header.length).setValues([header]);
  row++;
  row = renderMetricsTable_(sheet, row, result.byDoctor, true);
  row += 1;

  sheet.getRange(row++, 1).setValue('Лог ошибок');
  const logHeader = ['Месяц', 'Тип ошибки', 'Филиал', 'Врач', 'Дата', 'ID', 'Описание'];
  sheet.getRange(row, 1, 1, logHeader.length).setValues([logHeader]);
  row++;

  if (result.logs.length) {
    const logValues = result.logs.map((l) => [l.month, l.type, l.branch, l.doctor, l.date, l.id, l.description]);
    sheet.getRange(row, 1, logValues.length, logHeader.length).setValues(logValues);
    sheet.getRange(row, 5, logValues.length, 1).setNumberFormat('dd.mm.yyyy');
  }

  sheet.autoResizeColumns(1, 14);
}

function renderMetricsTable_(sheet, startRow, rows, withDoctor) {
  let row = startRow;
  const total = emptyMetric_();

  rows.forEach((r) => {
    addMetric_(total, r.metrics);
    const values = toOutputRow_(r.metrics, r.branch, withDoctor ? r.doctor : '');
    sheet.getRange(row++, 1, 1, values.length).setValues([values]);
  });

  const totalValues = toOutputRow_(total, 'ИТОГО', withDoctor ? 'ИТОГО' : '');
  sheet.getRange(row++, 1, 1, totalValues.length).setValues([totalValues]);
  return row;
}

function toOutputRow_(m, branch, doctor) {
  const utilizationReception = m.receptionHoursMin ? (m.receptionsMin / m.receptionHoursMin) * 100 : 0;
  const utilizationProcedure = m.procedureHoursMin ? (m.proceduresMin / m.procedureHoursMin) * 100 : 0;

  return [
    branch,
    doctor,
    round1_(m.receptionHoursMin / 60),
    round1_(m.procedureHoursMin / 60),
    round1_(m.lunchHoursMin / 60),
    round1_(m.otherHoursMin / 60),
    round1_(m.receptionsMin / 60),
    round1_(m.primaryMin / 60),
    round1_(m.repeatMin / 60),
    round1_(m.otherReceptionMin / 60),
    round1_(m.proceduresMin / 60),
    round1_(m.proceduresCount),
    round1_(utilizationReception),
    round1_(utilizationProcedure),
  ];
}

function toMetricsRows_(map) {
  return [...map.values()].sort((a, b) => {
    if (a.branch !== b.branch) return a.branch.localeCompare(b.branch, 'ru');
    return (a.doctor || '').localeCompare(b.doctor || '', 'ru');
  });
}

function normalizeRequest_(row, excludedEmployees, allowedEmployees, rangeStart, rangeEnd, unassignedSet) {
  const status = str_(row['Состояние']);
  if (!isCompletedStatus(status)) return null;

  const nomenclature = str_(row['Номенклатура']).toLowerCase();
  if (!nomenclature.includes('врача')) return null;
  if (nomenclature.includes('физ') || nomenclature.includes('лфк')) return null;

  const doctor = str_(row['Сотрудник']);
  const client = str_(row['Клиент']);
  if (!doctor || (excludedEmployees && excludedEmployees.has(doctor)) || hasStandaloneTestWord_(doctor) || hasStandaloneTestWord_(client)) return null;
  if (allowedEmployees && !allowedEmployees.has(doctor)) {
    if (unassignedSet) unassignedSet.add(doctor);
    return null;
  }

  const startDate = toDate_(row['ДатаНачала']);
  const endDate = toDate_(row['ДатаОкончания']);
  if (!startDate || !endDate) return null;
  if (rangeStart && rangeEnd && !isDateInRangeInclusive_(startDate, rangeStart, rangeEnd)) return null;

  const startMin = startDate.getHours() * 60 + startDate.getMinutes();
  const endMin = endDate.getHours() * 60 + endDate.getMinutes();
  let durationMin = endMin - startMin;
  if (durationMin <= 0) durationMin += 24 * 60;

  return {
    month: monthKey_(startDate),
    date: new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()),
    dateKey: dateKey_(startDate),
    branch: str_(row['Филиал']),
    doctor,
    id: str_(row['Номер заявки']) || str_(row['Номер']) || str_(row['ID']) || '',
    type: detectReceptionType_(nomenclature),
    durationMin,
    durationFallback: false,
    startMin,
    endMin,
  };
}


function normalizeProcedureRequest_(row, excludedEmployees, allowedEmployees, rangeStart, rangeEnd, unassignedSet) {
  const status = str_(row['Состояние']);
  if (!isCompletedStatus(status)) return null;

  const nomenclature = str_(row['Номенклатура']).toLowerCase();
  if (!nomenclature || nomenclature.includes('врача')) return null;

  const branch = str_(row['Филиал']);
  const cabinet = str_(row['Кабинет']);
  if (!branch || !cabinet) return null;

  const cabinetLower = cabinet.toLowerCase();
  if (!cabinetLower.includes('фтл') && !cabinetLower.includes('магнит')) return null;

  const doctor = str_(row['Сотрудник']);
  const client = str_(row['Клиент']);
  if (!doctor || (excludedEmployees && excludedEmployees.has(doctor)) || hasStandaloneTestWord_(doctor) || hasStandaloneTestWord_(client)) return null;
  if (allowedEmployees && !allowedEmployees.has(doctor)) {
    if (unassignedSet) unassignedSet.add(doctor);
    return null;
  }

  const startDate = toDate_(row['ДатаНачала']);
  const endDate = toDate_(row['ДатаОкончания']);
  if (!startDate || !endDate) return null;
  if (rangeStart && rangeEnd && !isDateInRangeInclusive_(startDate, rangeStart, rangeEnd)) return null;

  const startMin = startDate.getHours() * 60 + startDate.getMinutes();
  const endMin = endDate.getHours() * 60 + endDate.getMinutes();
  if (endMin <= startMin) return null;

  return {
    month: monthKey_(startDate),
    dateKey: dateKey_(startDate),
    branch,
    cabinet,
    startMin,
    endMin,
  };
}

function allocateProcedureMinutes_(procedureRequests, scheduleIndex, metricsBranch, metricsDoctor) {
  const byBranchDate = new Map();

  procedureRequests.forEach((req) => {
    const key = `${req.dateKey}|${req.branch}`;
    if (!byBranchDate.has(key)) byBranchDate.set(key, []);
    byBranchDate.get(key).push(req);
  });

  byBranchDate.forEach((reqs, key) => {
    const [dateKey, branch] = key.split('|');

    accumulateProcedureCount_(metricsBranch, branch, reqs.length);

    const doctorIntervals = [];
    scheduleIndex.forEach((intervals, doctorKey) => {
      const [dKey, br, doctor] = doctorKey.split('|');
      if (dKey !== dateKey || br !== branch) return;
      intervals
        .filter((i) => i.type === 'Процедурные часы')
        .forEach((i) => doctorIntervals.push({ doctor, start: i.start, end: i.end }));
    });

    const points = new Set();
    reqs.forEach((r) => {
      points.add(r.startMin);
      points.add(r.endMin);
    });
    doctorIntervals.forEach((d) => {
      points.add(d.start);
      points.add(d.end);
    });

    const sorted = [...points].sort((a, b) => a - b);
    for (let i = 0; i < sorted.length - 1; i += 1) {
      const segStart = sorted[i];
      const segEnd = sorted[i + 1];
      const segLen = segEnd - segStart;
      if (segLen <= 0) continue;

      const hasProcedure = reqs.some((r) => r.startMin < segEnd && r.endMin > segStart);
      if (!hasProcedure) continue;

      accumulateProcedure_(metricsBranch, branch, segLen);

      const activeDoctors = [...new Set(
        doctorIntervals
          .filter((d) => d.start < segEnd && d.end > segStart)
          .map((d) => d.doctor)
      )];

      if (!activeDoctors.length) continue;
      const share = segLen / activeDoctors.length;
      activeDoctors.forEach((doctor) => {
        accumulateProcedure_(metricsDoctor, `${branch}|${doctor}`, share, branch, doctor);
      });
    }

    reqs.forEach((req) => {
      allocateProcedureCountForRequest_(req, doctorIntervals, metricsDoctor, branch);
    });
  });
}

function allocateProcedureCountForRequest_(req, doctorIntervals, metricsDoctor, branch) {
  const relevantDoctorIntervals = doctorIntervals.filter((d) => d.start < req.endMin && d.end > req.startMin);
  if (!relevantDoctorIntervals.length) return;

  const points = new Set([req.startMin, req.endMin]);
  relevantDoctorIntervals.forEach((d) => {
    points.add(Math.max(req.startMin, d.start));
    points.add(Math.min(req.endMin, d.end));
  });

  const sorted = [...points].sort((a, b) => a - b);
  const reqLen = req.endMin - req.startMin;
  if (reqLen <= 0) return;

  for (let i = 0; i < sorted.length - 1; i += 1) {
    const segStart = sorted[i];
    const segEnd = sorted[i + 1];
    const segLen = segEnd - segStart;
    if (segLen <= 0) continue;

    const activeDoctors = [...new Set(
      relevantDoctorIntervals
        .filter((d) => d.start < segEnd && d.end > segStart)
        .map((d) => d.doctor)
    )];
    if (!activeDoctors.length) continue;

    const share = (segLen / reqLen) / activeDoctors.length;
    activeDoctors.forEach((doctor) => {
      accumulateProcedureCount_(metricsDoctor, `${branch}|${doctor}`, share, branch, doctor);
    });
  }
}


function detectReceptionType_(nomenclatureLower) {
  if (nomenclatureLower.includes('первичный')) return 'Первичный';
  if (nomenclatureLower.includes('повторный')) return 'Повторный';
  return 'Прочие';
}

function accumulateTime_(map, key, type, minutes, branch, doctor) {
  if (!map.has(key)) map.set(key, { branch: branch || key, doctor: doctor || '', metrics: emptyMetric_() });
  const entry = map.get(key).metrics;

  switch (type) {
    case 'Приемные часы':
      entry.receptionHoursMin += minutes;
      break;
    case 'Процедурные часы':
      entry.procedureHoursMin += minutes;
      break;
    case 'Обед':
      entry.lunchHoursMin += minutes;
      break;
    default:
      entry.otherHoursMin += minutes;
      break;
  }
}

function accumulateReception_(map, key, receptionType, minutes, branch, doctor) {
  if (!map.has(key)) map.set(key, { branch: branch || key, doctor: doctor || '', metrics: emptyMetric_() });
  const entry = map.get(key).metrics;

  entry.receptionsMin += minutes;
  if (receptionType === 'Первичный') entry.primaryMin += minutes;
  else if (receptionType === 'Повторный') entry.repeatMin += minutes;
  else entry.otherReceptionMin += minutes;
}


function accumulateProcedure_(map, key, minutes, branch, doctor) {
  if (!map.has(key)) map.set(key, { branch: branch || key, doctor: doctor || '', metrics: emptyMetric_() });
  map.get(key).metrics.proceduresMin += minutes;
}

function accumulateProcedureCount_(map, key, count, branch, doctor) {
  if (!map.has(key)) map.set(key, { branch: branch || key, doctor: doctor || '', metrics: emptyMetric_() });
  map.get(key).metrics.proceduresCount += count;
}

function emptyMetric_() {
  return {
    receptionHoursMin: 0,
    procedureHoursMin: 0,
    lunchHoursMin: 0,
    otherHoursMin: 0,
    receptionsMin: 0,
    primaryMin: 0,
    repeatMin: 0,
    otherReceptionMin: 0,
    proceduresMin: 0,
    proceduresCount: 0,
  };
}

function addMetric_(target, source) {
  Object.keys(target).forEach((k) => {
    target[k] += source[k] || 0;
  });
}

function logRow_(month, type, branch, doctor, date, id, description) {
  return { month, type, branch, doctor, date, id, description };
}

function buildScheduleIndex_(rows) {
  const byMonth = new Map();
  rows.forEach((row) => {
    const date = toDate_(row['Дата']);
    if (!date) return;
    const month = monthKey_(date);
    if (!byMonth.has(month)) byMonth.set(month, []);
    byMonth.get(month).push(row);
  });
  return byMonth;
}

function mergeIntervals_(intervals) {
  if (!intervals.length) return [];
  const sorted = [...intervals].sort((a, b) => a.start - b.start || a.end - b.end);
  const merged = [Object.assign({}, sorted[0])];

  for (let i = 1; i < sorted.length; i += 1) {
    const current = sorted[i];
    const last = merged[merged.length - 1];
    if (current.start <= last.end) {
      if (current.end > last.end) last.end = current.end;
    } else {
      merged.push(Object.assign({}, current));
    }
  }

  return merged;
}

function resolveIntervals_(records) {
  const points = new Set();
  records.forEach((r) => {
    points.add(r.start);
    points.add(r.end);
  });

  const sorted = [...points].sort((a, b) => a - b);
  const segments = [];

  for (let i = 0; i < sorted.length - 1; i += 1) {
    const segStart = sorted[i];
    const segEnd = sorted[i + 1];
    if (segEnd <= segStart) continue;

    const candidates = records.filter((r) => r.start < segEnd && r.end > segStart);
    if (!candidates.length) continue;

    candidates.sort((a, b) => {
      const cmp = a.createdAtSort - b.createdAtSort;
      if (cmp !== 0) return cmp;
      return a.rowOrder - b.rowOrder;
    });

    const winner = candidates[candidates.length - 1];
    const previous = segments[segments.length - 1];

    if (
      previous &&
      previous.end === segStart &&
      previous.timeKind === winner.timeKind &&
      previous.timeType === winner.timeType &&
      previous.createdAtSort === winner.createdAtSort &&
      previous.eventId === winner.eventId &&
      str_(previous.cabinet) === str_(winner.cabinet)
    ) {
      previous.end = segEnd;
    } else {
      segments.push({
        date: winner.date,
        dateLabel: winner.dateLabel,
        branch: winner.branch,
        employee: winner.employee,
        cabinet: winner.cabinet || '',
        cabinetType: winner.cabinetType || '',
        start: segStart,
        end: segEnd,
        timeKind: winner.timeKind,
        timeType: winner.timeType,
        createdAt: winner.createdAt,
        createdAtSort: winner.createdAtSort,
        createdAtLabel: winner.createdAtLabel,
        eventId: winner.eventId,
      });
    }
  }

  return segments;
}



function mergeAdjacentSlotsByType_(rows) {
  if (!rows.length) return [];
  const merged = [];

  rows.forEach((row) => {
    const prev = merged[merged.length - 1];
    const canMerge = prev
      && dateKey_(prev.date) === dateKey_(row.date)
      && prev.branch === row.branch
      && prev.employee === row.employee
      && str_(prev.cabinet) === str_(row.cabinet)
      && prev.timeType === row.timeType
      && prev.end === row.start;

    if (!canMerge) {
      merged.push(Object.assign({}, row));
      return;
    }

    prev.end = row.end;
    if (row.createdAtSort > prev.createdAtSort) {
      prev.createdAt = row.createdAt;
      prev.createdAtSort = row.createdAtSort;
      prev.createdAtLabel = row.createdAtLabel;
    }

    const eventIds = [str_(prev.eventId), str_(row.eventId)].filter((v) => v);
    prev.eventId = [...new Set(eventIds)].join(', ');
  });

  return merged;
}

function getEmployeeRegistry_(ss) {
  const unassignedFound = new Set();
  const allowed = new Set();

  const doctors = getDoctorsSetFromEmployeesSheet_(ss);
  const massage = getMassageSetFromEmployeesSheet_(ss);
  const ft = getFtSetFromEmployeesSheet_(ss);

  doctors.forEach((n) => allowed.add(n));
  massage.forEach((n) => allowed.add(n));
  ft.forEach((n) => allowed.add(n));

  return { allowed, unassignedFound };
}

function saveUnassignedEmployees_(ss, unassignedSet) {
  const sheet = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!sheet) return;

  const maxCols = sheet.getLastColumn() || 1;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 1, lastRow, maxCols).getDisplayValues();
  const header = values[0].map((v) => str_(v));

  let unassignedCol = header.findIndex((h) => h.toLowerCase() === 'нераспределенные');
  if (unassignedCol < 0) {
    unassignedCol = maxCols;
    sheet.getRange(1, unassignedCol + 1).setValue('Нераспределенные');
    // перечитаем диапазон после добавления колонки заголовка
    const newMaxCols = sheet.getLastColumn();
    const newLastRow = Math.max(sheet.getLastRow(), 1);
    values.length = 0;
    values.push(...sheet.getRange(1, 1, newLastRow, newMaxCols).getDisplayValues());
  }

  const rowCount = values.length - 1;
  const categorized = new Set();
  const currentUnassigned = [];

  for (let r = 1; r < values.length; r += 1) {
    const row = values[r];
    row.forEach((cell, c) => {
      const name = str_(cell);
      if (!name) return;
      if (c === unassignedCol) {
        currentUnassigned.push(name);
      } else {
        categorized.add(name);
      }
    });
  }

  // Чистим нераспределенных: удаляем тех, кто уже есть в любых других столбцах
  const keptUnassigned = currentUnassigned.filter((name, idx, arr) => !categorized.has(name) && arr.indexOf(name) === idx);

  // Добавляем новых только если их нет ни в одном столбце (включая нераспределенных)
  const knownAll = new Set([...categorized, ...keptUnassigned]);
  const incoming = unassignedSet ? [...unassignedSet].map((v) => str_(v)).filter((v) => v) : [];
  const toAppend = incoming.filter((name, idx, arr) => !knownAll.has(name) && arr.indexOf(name) === idx).sort((a, b) => a.localeCompare(b, 'ru'));
  const finalUnassigned = [...keptUnassigned, ...toAppend];

  const colIndex = unassignedCol + 1;
  const clearRows = Math.max(rowCount, finalUnassigned.length, 1);
  sheet.getRange(2, colIndex, clearRows, 1).clearContent();
  if (finalUnassigned.length) {
    sheet.getRange(2, colIndex, finalUnassigned.length, 1).setValues(finalUnassigned.map((v) => [v]));
  }
}

function getExcludedEmployeesSet_(ss) {
  const set = new Set();
  const sheet = ss.getSheetByName(SHEET_EXCLUSIONS);
  if (!sheet) return set;

  const rows = getSheetObjectsDisplay_(sheet);
  rows.forEach((row) => {
    const employee = str_(row['Сотрудник']);
    if (employee) set.add(employee);
  });
  return set;
}


function hasStandaloneTestWord_(value) {
  const tokens = str_(value).toLowerCase().split(/[^a-zа-яё0-9]+/i).filter((t) => t);
  return tokens.includes('тест');
}

function formatDateDDMMYYYY_(value) {
  const date = toDate_(value);
  if (!date) return '';
  return Utilities.formatDate(date, FL2_TIMEZONE, 'dd.MM.yyyy');
}

function formatDateTimeDDMMYYYY_(value) {
  const date = toDate_(value);
  if (!date) return '';
  return Utilities.formatDate(date, FL2_TIMEZONE, 'dd.MM.yyyy HH:mm:ss');
}

function sortableDateTime_(value) {
  const date = toDate_(value);
  if (!date) return null;
  return Number(Utilities.formatDate(date, FL2_TIMEZONE, 'yyyyMMddHHmmss'));
}

function mapTimeType_(timeKind) {
  return TIME_TYPE_MAP[timeKind] || 'Прочие часы';
}

function getSheetObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map((h) => str_(h));
  return values.slice(1).map((row) => {
    const item = {};
    headers.forEach((h, i) => {
      item[h] = row[i];
    });
    return item;
  });
}

function getSheetObjectsDisplay_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0].map((h) => str_(h));
  return values.slice(1).map((row) => {
    const item = {};
    headers.forEach((h, i) => {
      item[h] = row[i];
    });
    return item;
  });
}

function getSheetObjectsByBranch_(sheet, branchName, useDisplayValues) {
  const branch = str_(branchName);
  if (!branch) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const readFn = useDisplayValues ? 'getDisplayValues' : 'getValues';
  const headers = sheet.getRange(1, 1, 1, lastCol)[readFn]()[0].map((h) => str_(h));
  const branchIdx = headers.indexOf('Филиал');
  if (branchIdx < 0) return [];

  const branchColumn = sheet.getRange(2, branchIdx + 1, lastRow - 1, 1).getDisplayValues().map((r) => str_(r[0]));
  const ranges = [];
  let start = null;

  for (let i = 0; i < branchColumn.length; i += 1) {
    const isMatch = branchColumn[i] === branch;
    if (isMatch && start === null) start = i;
    if (!isMatch && start !== null) {
      ranges.push({ startIdx: start, endIdx: i - 1 });
      start = null;
    }
  }
  if (start !== null) ranges.push({ startIdx: start, endIdx: branchColumn.length - 1 });
  if (!ranges.length) return [];

  const out = [];
  ranges.forEach((r) => {
    const rowStart = r.startIdx + 2;
    const rowCount = r.endIdx - r.startIdx + 1;
    const values = sheet.getRange(rowStart, 1, rowCount, lastCol)[readFn]();
    out.push(...mapRowsToObjects_(headers, values));
  });
  return out;
}

function mapRowsToObjects_(headers, rows) {
  return rows.map((row) => {
    const item = {};
    headers.forEach((h, i) => {
      item[h] = row[i];
    });
    return item;
  });
}

function formatTimeCell_(value) {
  const minutes = parseTimeToMinutes_(value);
  if (minutes === null) return str_(value);
  return minutesToTimeString_(minutes);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function parseTimeToMinutes_(value) {
  if (value === '' || value === null || value === undefined) return null;
  if (typeof value === 'number') {
    const dayMinutes = Math.round((value % 1) * 24 * 60);
    return ((dayMinutes % 1440) + 1440) % 1440;
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return value.getHours() * 60 + value.getMinutes();
  }

  const str = String(value).trim();
  const match = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!match) return null;

  const h = Number(match[1]);
  const m = Number(match[2]);
  if (h < 0 || h > 23 || m < 0 || m > 59) return null;
  return h * 60 + m;
}

function minutesToTimeString_(minutes) {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function toDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return value;
  }
  if (value === '' || value === null || value === undefined) return null;

  if (typeof value === 'number') {
    const epochMs = Math.round((value - 25569) * 86400 * 1000);
    const parsedNum = new Date(epochMs);
    if (!Number.isNaN(parsedNum.getTime())) return parsedNum;
  }

  const normalized = str_(value);
  const isoDate = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoDate) {
    const year = Number(isoDate[1]);
    const month = Number(isoDate[2]);
    const day = Number(isoDate[3]);
    const parsedIsoDate = new Date(year, month - 1, day, 0, 0, 0, 0);
    if (!Number.isNaN(parsedIsoDate.getTime())) return parsedIsoDate;
  }

  const ruDateTime = normalized.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (ruDateTime) {
    const day = Number(ruDateTime[1]);
    const month = Number(ruDateTime[2]);
    const year = Number(ruDateTime[3]);
    const hour = Number(ruDateTime[4] !== undefined ? ruDateTime[4] : 12);
    const minute = Number(ruDateTime[5] || 0);
    const second = Number(ruDateTime[6] || 0);
    const parsedRu = new Date(year, month - 1, day, hour, minute, second, 0);
    if (!Number.isNaN(parsedRu.getTime())) return parsedRu;
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  return parsed;
}

function dateKey_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, FL2_TIMEZONE, 'yyyy-MM-dd');
}

function monthKey_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, FL2_TIMEZONE, 'yyyy-MM');
}

function str_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function round1_(value) {
  return Math.round(value * 10) / 10;
}
var FL2_SHEET_FORWARD_LOAD = 'План загрузки';
var FL2_SHEET_MAILING = 'Отправка';
var FL2_SHEET_OVERALL_FORECAST = 'Общий план';
var FL2_PROP_FORWARD_BATCH_STATE = 'FL2_PROP_FORWARD_BATCH_STATE';
var FL2_PROP_FORWARD_BATCH_PENDING_EMAIL = 'FL2_PROP_FORWARD_BATCH_PENDING_EMAIL';
var FL2_PROP_FORWARD_BATCH_RUN_ID = 'FL2_PROP_FORWARD_BATCH_RUN_ID';
var FL2_FORWARD_BATCH_HANDLER = 'runForwardLoadReportAllBranchesStep';
var FL2_SHEET_BATCH_SERVICE = 'Сервис план';
var FL2_BATCH_STEP_DELAY_MS = 5000; // 5 секунд
var FL2_FORWARD_BATCH_WATCHDOG_HANDLER = 'runForwardLoadBatchWatchdog';
var FL2_BATCH_WATCHDOG_INTERVAL_MIN = 1;
var FL2_BATCH_STALE_MS = 10 * 60 * 1000;
var FL2_PROP_FORWARD_BATCH_LOCK_OWNER = 'FL2_PROP_FORWARD_BATCH_LOCK_OWNER';
var FL2_BATCH_MAX_RUN_MS = 120000;
var FL2_BATCH_MAX_BRANCHES_PER_RUN = 1;
var FL2_ENABLE_LOG = false;
var FL2_PLAN_ARCHIVE_SPREADSHEET_ID = '1w9PLSlkxbJX6kigWtxYZElTyCLTM1ceDriA29yeFBgc';
var FL2_TIMEZONE = 'Europe/Moscow';

function fl2Log_(level, message, details) {
  if (!FL2_ENABLE_LOG) return;
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = getOrCreateSheet_(ss, SHEET_REPORT_LOG);
    const header = ['Время', 'Уровень', 'Сообщение', 'Детали'];
    const needHeader = sheet.getLastRow() < 1 || str_(sheet.getRange(1, 1).getDisplayValue()) !== 'Время';
    if (needHeader) {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
    }
    sheet.appendRow([new Date(), level || 'INFO', message || '', details || '']);
  } catch (e) {
    // намеренно игнорируем ошибки логирования, чтобы не ломать основной сценарий
  }
}

function fl2StateSummary_(state) {
  if (!state) return 'state=empty';
  const branches = Array.isArray(state.branches) ? state.branches.length : 0;
  return `index=${Number(state.index || 0)}, branches=${branches}`;
}

function fl2StageTimer_() {
  const startedAt = Date.now();
  return {
    ms: function ms() { return Date.now() - startedAt; },
    sec: function sec() { return (Date.now() - startedAt) / 1000; },
  };
}

function fl2LogStage_(level, message, details, timer) {
  const suffix = timer ? `, elapsedMs=${timer.ms()}` : '';
  fl2Log_(level, message, `${details || ''}${suffix}`);
}
function buildForwardLoadReport() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const precheckPrep = prepareForwardLoadInput_(ss, { persistUnassigned: false });
  if (!warnIfPlanDateDiffersFromActualData_(ui, precheckPrep.currentDay)) {
    SpreadsheetApp.getActive().toast('Построение плана отменено пользователем', MENU_NAME, 5);
    return;
  }

  const branches = getBranchesFromEmployeesSheet_(ss);
  const promptText = `Введите филиал (точно как в данных):
${branches.join('\n')}`;
  const response = ui.prompt('План загрузки 3/7 дней', promptText, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const selectedBranch = str_(response.getResponseText());
  if (!selectedBranch) throw new Error('Филиал не выбран');
  if (!branches.includes(selectedBranch)) throw new Error(`Филиал "${selectedBranch}" не найден в листе "Сотрудники"`);

  const prep = prepareForwardLoadInput_(ss, { selectedBranch, persistUnassigned: false, currentDay: precheckPrep.currentDay });
  const currentDay = buildForwardLoadReportForBranch_(ss, selectedBranch, prep);
  const sendAnswer = ui.alert('План загрузки', 'Сделать рассылку отчёта?', ui.ButtonSet.YES_NO);
  if (sendAnswer === ui.Button.YES) {
    sendForwardReportEmail(ss, selectedBranch, currentDay);
  }
  SpreadsheetApp.getActive().toast('План загрузки сформирован', MENU_NAME, 5);
}

function buildForwardLoadReportAllBranches() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(FL2_PROP_FORWARD_BATCH_STATE);
  if (existing) {
    const staleInfo = getForwardBatchStaleInfo_();
    fl2Log_('WARN', 'Пакетный план: найдено существующее состояние, принудительный перезапуск', staleInfo.reason);
  }

  // Переходим на однопроходный сценарий без триггеров/очереди
  clearForwardLoadBatchState_();
  props.deleteProperty(FL2_PROP_FORWARD_BATCH_PENDING_EMAIL);

  const doctorScheduleSheet = ss.getSheetByName(SHEET_DOCTOR_SCHEDULE);
  const processedRequestsSheet = ss.getSheetByName(SHEET_PROCESSED_REQUESTS);
  if (!doctorScheduleSheet || !processedRequestsSheet) {
    throw new Error('Проверьте наличие листов: Обработанный график, Обработанные заявки');
  }
  ensureDoctorScheduleFresh(ss);

  const builtAt = detectDoctorScheduleBuiltAtFallback(doctorScheduleSheet);
  const graphSheet = ss.getSheetByName(SHEET_SCHEDULE_SOURCE);
  const sourceMaxGraphCreated = graphSheet ? getSourceGraphMaxCreatedAt_(graphSheet) : new Date(0);
  if (!builtAt) {
    throw new Error('Сначала сформируйте лист "Обработанный график" (в последней строке нет системного timestamp)');
  }
  if (sourceMaxGraphCreated > builtAt) {
    throw new Error(`"Обработанный график" неактуален по timestamp (последняя сборка: ${formatDateTimeDDMMYYYY_(builtAt)}). Обновите его перед запуском.`);
  }
  fl2Log_('INFO', 'Пакетный план: проверка актуальности графика', `scheduleTimestamp=${formatDateTimeDDMMYYYY_(builtAt)}, sourceMaxCreated=${formatDateTimeDDMMYYYY_(sourceMaxGraphCreated)}`);

  const prep = prepareForwardLoadInput_(ss, { persistUnassigned: false });
  const ui = SpreadsheetApp.getUi();
  if (!warnIfPlanDateDiffersFromActualData_(ui, prep.currentDay)) {
    SpreadsheetApp.getActive().toast('Пакетный план отменен пользователем', MENU_NAME, 5);
    return;
  }
  const sendAnswer = ui.alert('Пакетный план по клиникам', 'После завершения очереди отправить письма автоматически?', ui.ButtonSet.YES_NO);
  const sendEmails = sendAnswer === ui.Button.YES;
  const runId = `${Date.now()}`;
  const branches = prep.branches || [];
  const errors = [];
  fl2Log_('INFO', 'Пакетный план: старт (однопроходный без триггеров)', `runId=${runId}, клиник: ${branches.length}, sendEmails=${sendEmails ? 'YES' : 'NO'}`);

  initForwardBatchServiceSheet_(ss, branches, runId);
  const serviceSheet = getOrCreateSheet_(ss, FL2_SHEET_BATCH_SERVICE);

  const sharedExcluded = prep.excludedEmployees;
  const sharedRegistry = { allowed: prep.allowedEmployees, unassignedFound: prep.unassignedEmployees };
  const sharedDoctorsSet = prep.doctorsSet;
  const sharedMassageSet = prep.massageSet;
  const sharedFtSet = prep.ftSet;

  branches.forEach((branchName, idx) => {
    const branch = str_(branchName);
    const rowNum = idx + 2;
    const stepStartedAt = new Date();
    serviceSheet.getRange(rowNum, 2, 1, 3).setValues([['В работе', formatDateTimeDDMMYYYY_(stepStartedAt), '']]);

    try {
      const branchPrep = prepareForwardLoadInput_(ss, {
        selectedBranch: branch,
        currentDay: prep.currentDay,
        persistUnassigned: false,
        excludedEmployees: sharedExcluded,
        employeeRegistry: sharedRegistry,
        doctorsSet: sharedDoctorsSet,
        massageSet: sharedMassageSet,
        ftSet: sharedFtSet,
      });
      buildForwardLoadReportForBranch_(ss, branch, branchPrep);

      if (sendEmails) {
        sendForwardReportEmail(ss, branch, prep.currentDay);
      }

      serviceSheet.getRange(rowNum, 2, 1, 3).setValues([['Готово', formatDateTimeDDMMYYYY_(stepStartedAt), formatDateTimeDDMMYYYY_(new Date())]]);
    } catch (e) {
      const err = `${branch || '(пусто)'}: ${e && e.message ? e.message : String(e)}`;
      errors.push(err);
      serviceSheet.getRange(rowNum, 2, 1, 3).setValues([[`Ошибка: ${err}`, formatDateTimeDDMMYYYY_(stepStartedAt), formatDateTimeDDMMYYYY_(new Date())]]);
    }
  });

  const doneMsg = errors.length
    ? `Пакетный план завершён с ошибками: ${errors.slice(0, 3).join(' | ')}`
    : 'Пакетный план по клиникам завершён';
  SpreadsheetApp.getActive().toast(doneMsg, MENU_NAME, 8);
}

function warnIfPlanDateDiffersFromActualData_(ui, actualDataDay) {
  const actual = toDate_(actualDataDay);
  const formedAt = new Date();
  if (!actual) return true;
  const actualDate = new Date(actual.getFullYear(), actual.getMonth(), actual.getDate());
  const formedDate = new Date(formedAt.getFullYear(), formedAt.getMonth(), formedAt.getDate());
  if (actualDate.getTime() === formedDate.getTime()) return true;

  const message = [
    'Внимание: дата формирования не совпадает с актуальной датой данных.',
    `Дата формирования: ${formatDateDDMMYYYY_(formedDate)}`,
    `Актуальная дата данных (последняя выполненная заявка): ${formatDateDDMMYYYY_(actualDate)}`,
    '',
    'Продолжить построение плана?',
  ].join('\n');

  let shouldContinue = true;
  if (ui) {
    const answer = ui.alert('План загрузки: предупреждение по дате', message, ui.ButtonSet.YES_NO);
    shouldContinue = answer === ui.Button.YES;
  }
  fl2Log_('WARN', 'План: дата формирования отличается от актуальной даты данных', message);
  return shouldContinue;
}


function runForwardLoadReportAllBranchesStep() {
  runForwardLoadReportAllBranchesStepInternal_({ scheduleOnLockFail: true, lockRetryMs: 5000 });
}

function runForwardLoadReportAllBranchesStepInternal_(options) {
  const opts = options || {};
  const lockRetryMs = Number(opts.lockRetryMs || 5000);
  const maxRunMs = Number(opts.maxRunMs || FL2_BATCH_MAX_RUN_MS);
  const lock = LockService.getScriptLock();
  const deadline = Date.now() + Math.max(500, lockRetryMs);
  let locked = false;

  while (Date.now() <= deadline) {
    if (lock.tryLock(500)) {
      locked = true;
      break;
    }
    Utilities.sleep(250);
  }

  if (!locked) {
    const props = PropertiesService.getScriptProperties();
    const owner = props.getProperty(FL2_PROP_FORWARD_BATCH_LOCK_OWNER);
    fl2Log_('WARN', 'Пакетный план: lock не получен', `Повторим шаг позже (delayMs=${FL2_BATCH_STEP_DELAY_MS}), owner=${owner || 'unknown'}`);
    if (opts.scheduleOnLockFail !== false) {
      scheduleNextForwardLoadBatchStep_();
    }
    return false;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(FL2_PROP_FORWARD_BATCH_STATE);
    if (!raw) {
      fl2Log_('WARN', 'Пакетный план: состояние не найдено', 'Шаг завершен без выполнения');
      return true;
    }

    const state = JSON.parse(raw);
    const branches = Array.isArray(state.branches) ? state.branches : [];
    let index = Number(state.index || 0);
    const runId = str_(state.runId) || str_(props.getProperty(FL2_PROP_FORWARD_BATCH_RUN_ID));
    props.setProperty(FL2_PROP_FORWARD_BATCH_LOCK_OWNER, JSON.stringify({ runId, index, startedAt: Date.now() }));
    fl2Log_('INFO', 'Пакетный план: шаг', `runId=${runId}, ${fl2StateSummary_(state)}`);

    if (!branches.length || index >= branches.length) {
      fl2Log_('INFO', 'Пакетный план: достигнут конец очереди', fl2StateSummary_(state));
      finishForwardLoadBatch_(state);
      return true;
    }

    const ss = SpreadsheetApp.getActive();
    const serviceSheet = getOrCreateSheet_(ss, FL2_SHEET_BATCH_SERVICE);
    const excludedEmployees = getExcludedEmployeesSet_(ss);
    const employeeRegistry = getEmployeeRegistry_(ss);
    const doctorsSet = getDoctorsSetFromEmployeesSheet_(ss);
    const startedAtMs = Date.now();
    let lastStepFinishedAt = Date.now();
    let processedThisRun = 0;

    while (index < branches.length && (Date.now() - startedAtMs) < maxRunMs && processedThisRun < FL2_BATCH_MAX_BRANCHES_PER_RUN) {
      const branch = str_(branches[index]);
      const rowNum = index + 2;
      const stepStartedAt = new Date();
      serviceSheet.getRange(rowNum, 2, 1, 3).setValues([['В работе', formatDateTimeDDMMYYYY_(stepStartedAt), '']]);

      try {
        state.updatedAt = Date.now();
        props.setProperty(FL2_PROP_FORWARD_BATCH_STATE, JSON.stringify(state));
        props.setProperty(FL2_PROP_FORWARD_BATCH_LOCK_OWNER, JSON.stringify({ runId, index, branch, startedAt: state.updatedAt }));
        if (!branch) {
          state.errors = state.errors || [];
          state.errors.push('Пустое название филиала в очереди');
          fl2Log_('ERROR', 'Пакетный план: пустой филиал в очереди', `шаг ${index + 1}/${branches.length}`);
        } else {
          const branchTimer = fl2StageTimer_();
          const stagePrefix = `runId=${runId}, branch=${branch}, step=${index + 1}/${branches.length}`;
          fl2LogStage_('INFO', 'Пакетный план: начало обработки филиала', stagePrefix, branchTimer);

          touchForwardBatchHeartbeat_(props);
          const prepTimer = fl2StageTimer_();
          fl2LogStage_('INFO', 'Пакетный план: этап prepareForwardLoadInput_ START', stagePrefix, prepTimer);
          const prep = prepareForwardLoadInput_(ss, {
            selectedBranch: branch,
            currentDay: state.currentDay,
            persistUnassigned: false,
            excludedEmployees,
            employeeRegistry,
            doctorsSet,
            logCtx: { runId, branch, step: `${index + 1}/${branches.length}` },
            onHeartbeat: () => touchForwardBatchHeartbeat_(props),
          });
          fl2LogStage_('INFO', 'Пакетный план: этап prepareForwardLoadInput_ END', `${stagePrefix}, scheduleRows=${prep.scheduleRows.length}, requestRows=${prep.requestRowsRaw.length}`, prepTimer);

          touchForwardBatchHeartbeat_(props);
          const calcTimer = fl2StageTimer_();
          fl2LogStage_('INFO', 'Пакетный план: этап buildForwardLoadReportForBranch_ START', stagePrefix, calcTimer);
          buildForwardLoadReportForBranch_(ss, branch, prep);
          fl2LogStage_('INFO', 'Пакетный план: этап buildForwardLoadReportForBranch_ END', stagePrefix, calcTimer);

          touchForwardBatchHeartbeat_(props);
          fl2LogStage_('INFO', 'Пакетный план: филиал обработан', stagePrefix, branchTimer);
        }
      } catch (e) {
        state.errors = state.errors || [];
        state.errors.push(`${branch || '(пусто)'}: ${e && e.message ? e.message : String(e)}`);
        fl2Log_('ERROR', 'Пакетный план: ошибка обработки филиала', `${branch || '(пусто)'}: ${e && e.message ? e.message : String(e)}`);
      }

      index += 1;
      processedThisRun += 1;
      lastStepFinishedAt = Date.now();
      state.index = index;
      state.updatedAt = lastStepFinishedAt;
      props.setProperty(FL2_PROP_FORWARD_BATCH_STATE, JSON.stringify(state));
      serviceSheet.getRange(rowNum, 2, 1, 3).setValues([['Готово', formatDateTimeDDMMYYYY_(stepStartedAt), formatDateTimeDDMMYYYY_(new Date(lastStepFinishedAt))]]);
    }

    if (index >= branches.length) {
      fl2Log_('INFO', 'Пакетный план: финал после последнего филиала', fl2StateSummary_(state));
      finishForwardLoadBatch_(state);
      return true;
    }

    scheduleNextForwardLoadBatchStep_(lastStepFinishedAt);
    fl2Log_('INFO', 'Пакетный план: шаг завершен, продолжаем по триггеру', `runId=${runId}, обработано за запуск=${processedThisRun}, до index=${index}, всего=${branches.length}`);
    return true;
  } catch (e) {
    fl2Log_('ERROR', 'Пакетный план: аварийная ошибка шага', e && e.stack ? e.stack : String(e));
    scheduleNextForwardLoadBatchStep_();
    throw e;
  } finally {
    try { PropertiesService.getScriptProperties().deleteProperty(FL2_PROP_FORWARD_BATCH_LOCK_OWNER); } catch (e) {}
    lock.releaseLock();
  }
}

function scheduleNextForwardLoadBatchStep_(stepFinishedAtMs) {
  clearForwardLoadBatchTriggers_();
  const finishedAt = Number(stepFinishedAtMs || Date.now());
  const targetAt = finishedAt + FL2_BATCH_STEP_DELAY_MS;
  const delayMs = Math.max(1000, targetAt - Date.now());

  ScriptApp.newTrigger(FL2_FORWARD_BATCH_HANDLER)
    .timeBased()
    .after(delayMs)
    .create();

  const count = getForwardLoadBatchTriggerCount_();
  fl2Log_('INFO', 'Пакетный план: триггер создан', `runForwardLoadReportAllBranchesStep через ${Math.round(delayMs / 1000)} сек от времени окончания шага (активных batch-триггеров: ${count})`);
}

function finishForwardLoadBatch_(state) {
  const ss = SpreadsheetApp.getActive();
  const currentDay = state && state.currentDay ? new Date(Number(state.currentDay)) : new Date();
  const props = PropertiesService.getScriptProperties();

  props.setProperty(FL2_PROP_FORWARD_BATCH_PENDING_EMAIL, JSON.stringify({
    runId: str_(state && state.runId),
    branches: Array.isArray(state && state.branches) ? state.branches : [],
    currentDay: currentDay.getTime(),
    errors: Array.isArray(state && state.errors) ? state.errors : [],
  }));

  clearForwardLoadBatchState_();
  fl2Log_('INFO', 'Пакетный план: очередь завершена', `runId=${str_(state && state.runId)}, подготовлена рассылка по ${Array.isArray(state && state.branches) ? state.branches.length : 0} клиникам`);

  const shouldSendEmails = Boolean(state && state.sendEmails);
  if (shouldSendEmails) {
    fl2Log_('INFO', 'Пакетный план: авторассылка после завершения', 'sendEmails=YES');
    confirmAndSendForwardBatchEmails();
    return;
  }

  fl2Log_('INFO', 'Пакетный план: авторассылка отключена', 'sendEmails=NO');
  const errors = state && Array.isArray(state.errors) ? state.errors.filter(Boolean) : [];
  const doneMsg = errors.length
    ? `Пакетный план завершён с ошибками: ${errors.slice(0, 3).join(' | ')}`
    : 'Пакетный план по клиникам завершён';
  SpreadsheetApp.getActive().toast(doneMsg, MENU_NAME, 8);
}

function confirmAndSendForwardBatchEmails() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(FL2_PROP_FORWARD_BATCH_PENDING_EMAIL);
  if (!raw) {
    fl2Log_('WARN', 'Пакетный план: нет payload для рассылки', 'FL2_PROP_FORWARD_BATCH_PENDING_EMAIL пуст');
    throw new Error('Нет подготовленного пакетного плана для рассылки');
  }

  const payload = JSON.parse(raw);
  const branches = Array.isArray(payload.branches) ? payload.branches : [];
  const currentDay = payload.currentDay ? new Date(Number(payload.currentDay)) : new Date();
  const errors = Array.isArray(payload.errors) ? payload.errors.slice() : [];
  const ss = SpreadsheetApp.getActive();

  branches.forEach((branch) => {
    try {
      if (str_(branch)) sendForwardReportEmail(ss, str_(branch), currentDay);
      fl2Log_('INFO', 'Пакетный план: письмо отправлено', str_(branch));
    } catch (e) {
      errors.push(`Рассылка ${branch || '(пусто)'}: ${e && e.message ? e.message : String(e)}`);
      fl2Log_('ERROR', 'Пакетный план: ошибка отправки письма', `${branch || '(пусто)'}: ${e && e.message ? e.message : String(e)}`);
    }
  });

  props.deleteProperty(FL2_PROP_FORWARD_BATCH_PENDING_EMAIL);
  const doneMsg = errors.length
    ? `Рассылка завершена с ошибками: ${errors.slice(0, 3).join(' | ')}`
    : 'Рассылка пакетного плана завершена';
  SpreadsheetApp.getActive().toast(doneMsg, MENU_NAME, 8);
}

function maybePromptForwardBatchEmail_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(FL2_PROP_FORWARD_BATCH_PENDING_EMAIL);
  if (!raw) return;

  const payload = JSON.parse(raw);
  const branches = Array.isArray(payload.branches) ? payload.branches.length : 0;
  const ui = tryGetUi_();
  if (!ui) return;

  const ask = ui.alert('Пакетный план по клиникам', `Пакетный план сформирован (клиник: ${branches}). Отправить рассылку сейчас?`, ui.ButtonSet.YES_NO);
  fl2Log_('INFO', 'Пакетный план: вопрос о рассылке при открытии', `Ответ: ${ask === ui.Button.YES ? 'YES' : 'NO'}`);
  if (ask === ui.Button.YES) {
    confirmAndSendForwardBatchEmails();
  }
}

function tryGetUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

function prepareForwardLoadInput_(ss, options) {
  const opts = options || {};
  const logCtx = opts.logCtx || null;
  const hb = typeof opts.onHeartbeat === 'function' ? opts.onHeartbeat : null;
  const stage = (msg, details, timer) => {
    if (!logCtx) return;
    const prefix = `runId=${str_(logCtx.runId)}, branch=${str_(logCtx.branch)}, step=${str_(logCtx.step)}`;
    fl2LogStage_('INFO', `Пакетный план: ${msg}`, `${prefix}${details ? `, ${details}` : ''}`, timer);
  };
  const tick = () => { if (hb) hb(); };

  const scheduleSheet = ss.getSheetByName(SHEET_DOCTOR_SCHEDULE);
  const requestsSheet = ss.getSheetByName(SHEET_PROCESSED_REQUESTS);
  if (!scheduleSheet || !requestsSheet) {
    throw new Error('Проверьте наличие листов: Обработанный график, Обработанные заявки');
  }
  ensureDoctorScheduleFresh(ss);

  const excludedEmployees = opts.excludedEmployees ? opts.excludedEmployees : getExcludedEmployeesSet_(ss);
  const employeeRegistry = opts.employeeRegistry ? opts.employeeRegistry : getEmployeeRegistry_(ss);
  const selectedBranch = str_(opts.selectedBranch);

  let currentDay = opts.currentDay ? new Date(Number(opts.currentDay)) : null;
  let requestRowsRawAll = null;
  if (!currentDay || Number.isNaN(currentDay.getTime())) {
    const detectDayTimer = fl2StageTimer_();
    requestRowsRawAll = getSheetObjectsDisplay_(requestsSheet);
    stage('prepare: loaded processed requests for currentDay detect', `rows=${requestRowsRawAll.length}`, detectDayTimer);
    tick();
    const latestRequestsAllBranches = buildLatestRequests(requestRowsRawAll, excludedEmployees, employeeRegistry.allowed, employeeRegistry.unassignedFound);
    currentDay = detectCurrentDay(latestRequestsAllBranches);
  }
  if (!currentDay) throw new Error('Не найдены выполненные заявки для определения текущего дня');

  const rangeStart = addDays(currentDay, -42);
  const rangeEnd = addDays(currentDay, 7);

  const readScheduleTimer = fl2StageTimer_();
  const scheduleSourceRows = selectedBranch
    ? getSheetObjectsByBranch_(scheduleSheet, selectedBranch, true)
    : getSheetObjectsDisplay_(scheduleSheet);

  stage('prepare: read schedule source', `rows=${scheduleSourceRows.length}`, readScheduleTimer);
  tick();

  const filterScheduleTimer = fl2StageTimer_();
  const scheduleRows = scheduleSourceRows
    .filter((r) => {
      const d = toDate_(r['Дата']);
      const emp = str_(r['Сотрудник']);
      const branch = str_(r['Филиал']);
      const cabinetType = detectCabinetType_(r['Кабинет']);
      const isCabinetRow = Boolean(cabinetType);
      if (selectedBranch && branch !== selectedBranch) return false;
      if (!isCabinetRow) {
        if (!emp || !employeeRegistry.allowed.has(emp)) {
          if (emp) employeeRegistry.unassignedFound.add(emp);
          return false;
        }
      }
      return isDateInRangeInclusive_(d, rangeStart, rangeEnd);
    })
    .sort((a, b) => {
      const ba = str_(a['Филиал']);
      const bb = str_(b['Филиал']);
      if (ba !== bb) return ba.localeCompare(bb, 'ru');
      const da = toDate_(a['Дата']);
      const db = toDate_(b['Дата']);
      return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
    });
  stage('prepare: filtered schedule rows', `rows=${scheduleRows.length}`, filterScheduleTimer);
  tick();

  const readRequestsTimer = fl2StageTimer_();
  const requestSourceRows = selectedBranch
    ? getSheetObjectsByBranch_(requestsSheet, selectedBranch, true)
    : (requestRowsRawAll || getSheetObjectsDisplay_(requestsSheet));

  stage('prepare: read requests source', `rows=${requestSourceRows.length}`, readRequestsTimer);
  tick();

  const filterRequestsTimer = fl2StageTimer_();
  const requestRowsRaw = requestSourceRows.filter((r) => {
    const d = toDate_(r['ДатаНачала']);
    const branch = str_(r['Филиал']);
    if (selectedBranch && branch !== selectedBranch) return false;
    return isDateInRangeInclusive_(d, rangeStart, rangeEnd);
  });

  stage('prepare: filtered request rows', `rows=${requestRowsRaw.length}`, filterRequestsTimer);
  tick();

  const branchesTimer = fl2StageTimer_();
  const branchesAll = getBranchesFromEmployeesSheet_(ss);
  const branches = selectedBranch ? branchesAll.filter((b) => b === selectedBranch) : branchesAll;
  if (!branches.length) throw new Error('Не удалось определить список филиалов');
  stage('prepare: branches resolved', `count=${branches.length}`, branchesTimer);
  tick();

  if (!opts || opts.persistUnassigned !== false) {
    saveUnassignedEmployees_(ss, employeeRegistry.unassignedFound);
  }
  return {
    scheduleRows,
    requestRowsRaw,
    branches,
    excludedEmployees,
    currentDay,
    rangeStart,
    rangeEnd,
    allowedEmployees: employeeRegistry.allowed,
    unassignedEmployees: employeeRegistry.unassignedFound,
    doctorsSet: opts && opts.doctorsSet ? opts.doctorsSet : getDoctorsSetFromEmployeesSheet_(ss),
    massageSet: opts && opts.massageSet ? opts.massageSet : getMassageSetFromEmployeesSheet_(ss),
    ftSet: opts && opts.ftSet ? opts.ftSet : getFtSetFromEmployeesSheet_(ss),
    logCtx,
    onHeartbeat: hb,
  };
}

function buildForwardLoadReportForBranch_(ss, selectedBranch, prep) {
  const branchName = str_(selectedBranch);
  if (!branchName) throw new Error('Филиал не задан для пакетного расчета');

  const logCtx = prep && prep.logCtx ? prep.logCtx : null;
  const hb = prep && typeof prep.onHeartbeat === 'function' ? prep.onHeartbeat : null;
  const stage = (msg, details, timer) => {
    if (!logCtx) return;
    const prefix = `runId=${str_(logCtx.runId)}, branch=${branchName}, step=${str_(logCtx.step)}`;
    fl2LogStage_('INFO', `Пакетный план: ${msg}`, `${prefix}${details ? `, ${details}` : ''}`, timer);
  };
  const tick = () => { if (hb) hb(); };

  const filterTimer = fl2StageTimer_();
  const filteredScheduleRows = prep.scheduleRows.filter((r) => str_(r['Филиал']) === branchName);
  const filteredRequestRows = prep.requestRowsRaw.filter((r) => str_(r['Филиал']) === branchName);
  stage('branch: filtered branch rows', `scheduleRows=${filteredScheduleRows.length}, requestRows=${filteredRequestRows.length}`, filterTimer);
  tick();

  const scheduleIndexTimer = fl2StageTimer_();
  const scheduleByDoctorDay = buildReceptionScheduleIndex(filteredScheduleRows, prep.excludedEmployees, prep.allowedEmployees);
  stage('branch: buildReceptionScheduleIndex done', `keys=${scheduleByDoctorDay.size}`, scheduleIndexTimer);
  logDoctorForwardDebug_(ss, branchName, 'Долинский Павел Максимович', prep.currentDay, scheduleByDoctorDay);
  tick();

  const doctorsSet = prep.doctorsSet || getDoctorsSetFromEmployeesSheet_(ss);
  const massageSet = prep.massageSet || getMassageSetFromEmployeesSheet_(ss);
  const ftSet = prep.ftSet || getFtSetFromEmployeesSheet_(ss);

  const doctorScheduleByDoctorDay = new Map();
  scheduleByDoctorDay.forEach((intervals, key) => {
    const doctor = str_(key.split('|')[2]);
    if (doctorsSet.has(doctor)) doctorScheduleByDoctorDay.set(key, intervals);
  });
  stage('branch: doctor schedule subset built', `keys=${doctorScheduleByDoctorDay.size}`, fl2StageTimer_());
  tick();

  const latestReqTimer = fl2StageTimer_();
  const latestRequests = buildLatestRequests(filteredRequestRows, prep.excludedEmployees, prep.allowedEmployees, prep.unassignedEmployees, prep.rangeStart, prep.rangeEnd);
  stage('branch: buildLatestRequests done', `rows=${latestRequests.length}`, latestReqTimer);
  tick();

  const h3Timer = fl2StageTimer_();
  const horizon3 = computeWindowMetrics(prep.currentDay, 3, doctorScheduleByDoctorDay, latestRequests);
  stage('branch: computeWindowMetrics(3) done', `byBranch=${horizon3.byBranch.size}, byDoctor=${horizon3.byDoctor.size}`, h3Timer);
  tick();

  const h7Timer = fl2StageTimer_();
  const horizon7 = computeWindowMetrics(prep.currentDay, 7, doctorScheduleByDoctorDay, latestRequests);
  stage('branch: computeWindowMetrics(7) done', `byBranch=${horizon7.byBranch.size}, byDoctor=${horizon7.byDoctor.size}`, h7Timer);
  tick();

  const b3Timer = fl2StageTimer_();
  const baseline3 = computeBaselineMetrics(prep.currentDay, 3, doctorScheduleByDoctorDay, latestRequests);
  stage('branch: computeBaselineMetrics(3) done', `byBranch=${baseline3.byBranch.size}, byDoctor=${baseline3.byDoctor.size}`, b3Timer);
  tick();

  const b7Timer = fl2StageTimer_();
  const baseline7 = computeBaselineMetrics(prep.currentDay, 7, doctorScheduleByDoctorDay, latestRequests);
  stage('branch: computeBaselineMetrics(7) done', `byBranch=${baseline7.byBranch.size}, byDoctor=${baseline7.byDoctor.size}`, b7Timer);
  tick();

  const h3AllTimer = fl2StageTimer_();
  const horizon3All = computeWindowMetrics(prep.currentDay, 3, scheduleByDoctorDay, latestRequests, { includeAllNomenclature: true });
  stage('branch: computeWindowMetrics(3, all nomenclature) done', `byBranch=${horizon3All.byBranch.size}, byDoctor=${horizon3All.byDoctor.size}`, h3AllTimer);
  tick();

  const h7AllTimer = fl2StageTimer_();
  const horizon7All = computeWindowMetrics(prep.currentDay, 7, scheduleByDoctorDay, latestRequests, { includeAllNomenclature: true });
  stage('branch: computeWindowMetrics(7, all nomenclature) done', `byBranch=${horizon7All.byBranch.size}, byDoctor=${horizon7All.byDoctor.size}`, h7AllTimer);
  tick();

  const b3AllTimer = fl2StageTimer_();
  const baseline3All = computeBaselineMetrics(prep.currentDay, 3, scheduleByDoctorDay, latestRequests, { includeAllNomenclature: true });
  stage('branch: computeBaselineMetrics(3, all nomenclature) done', `byBranch=${baseline3All.byBranch.size}, byDoctor=${baseline3All.byDoctor.size}`, b3AllTimer);
  tick();

  const b7AllTimer = fl2StageTimer_();
  const baseline7All = computeBaselineMetrics(prep.currentDay, 7, scheduleByDoctorDay, latestRequests, { includeAllNomenclature: true });
  stage('branch: computeBaselineMetrics(7, all nomenclature) done', `byBranch=${baseline7All.byBranch.size}, byDoctor=${baseline7All.byDoctor.size}`, b7AllTimer);
  tick();

  const mergeTimer = fl2StageTimer_();
  const rowsBranch = mergeForwardRows(horizon3.byBranch, horizon7.byBranch, baseline3.byBranch, baseline7.byBranch);
  const rowsDoctorAll = mergeForwardRows(horizon3.byDoctor, horizon7.byDoctor, baseline3.byDoctor, baseline7.byDoctor);
  const rowsAllNomenclature = mergeForwardRows(horizon3All.byDoctor, horizon7All.byDoctor, baseline3All.byDoctor, baseline7All.byDoctor);
  stage('branch: mergeForwardRows done', `rowsBranch=${rowsBranch.length}, rowsDoctorAll=${rowsDoctorAll.length}`, mergeTimer);
  tick();

  const doctorFilterTimer = fl2StageTimer_();
  const rowsDoctor = rowsDoctorAll.filter((r) => doctorsSet.has(str_(r.doctor)));
  stage('branch: filtered doctors rows', `rowsDoctor=${rowsDoctor.length}`, doctorFilterTimer);
  tick();

  const massageSummary = aggregateForwardCategorySummary_(rowsAllNomenclature, massageSet);
  const ftSummary = aggregateForwardCategorySummary_(rowsAllNomenclature, ftSet);
  const cabinetSummary = buildCabinetTypeSummary_(prep.currentDay, filteredScheduleRows, latestRequests);
  const massageRows = rowsAllNomenclature.filter((r) => massageSet.has(str_(r.doctor)));
  const ftRows = rowsAllNomenclature.filter((r) => ftSet.has(str_(r.doctor)));
  const cabinetRows = buildCabinetRowsByName_(prep.currentDay, filteredScheduleRows, latestRequests);

  const renderTimer = fl2StageTimer_();
  renderForwardLoadSheet(ss, prep.currentDay, branchName, rowsBranch, rowsDoctor, massageSummary, ftSummary, cabinetSummary, massageRows, ftRows, cabinetRows);
  stage('branch: renderForwardLoadSheet done', `rowsBranch=${rowsBranch.length}, rowsDoctor=${rowsDoctor.length}`, renderTimer);
  tick();

  return prep.currentDay;
}

function logDoctorForwardDebug_(ss, branchName, doctorName, anchorDate, scheduleByDoctorDay) {
  try {
    const startDate = addDays(anchorDate, 1);
    const endDate = addDays(anchorDate, 3);
    const parts = [];
    let totalReceptionMin = 0;

    for (let d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()); d <= endDate; d = addDays(d, 1)) {
      const key = `${dateKey_(d)}|${branchName}|${doctorName}`;
      const intervals = scheduleByDoctorDay.get(key) || [];
      const reception = mergeIntervals_(intervals.filter((i) => i.type === 'Приемные часы'));
      const receptionMin = sumIntervalsMin(reception);
      totalReceptionMin += receptionMin;
      parts.push(`${formatDateDDMMYYYY_(d)}=${round1_(receptionMin / 60)}ч`);
    }

    const details = `Окно3д(${formatDateDDMMYYYY_(startDate)}..${formatDateDDMMYYYY_(endDate)}), ${doctorName}, ${parts.join('; ')}, Итого=${round1_(totalReceptionMin / 60)}ч`;
    const sheet = getOrCreateSheet_(ss, SHEET_REPORT_LOG);
    const header = ['Время', 'Уровень', 'Сообщение', 'Детали'];
    const needHeader = sheet.getLastRow() < 1 || str_(sheet.getRange(1, 1).getDisplayValue()) !== 'Время';
    if (needHeader) {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
    }
    sheet.appendRow([new Date(), 'DEBUG', 'Forward debug doctor', details]);
  } catch (e) {
    // не ломаем основной расчет из-за диагностического лога
  }
}
function clearForwardLoadBatchTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction && t.getHandlerFunction() === FL2_FORWARD_BATCH_HANDLER)
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

function getForwardLoadBatchTriggerCount_() {
  return ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction && t.getHandlerFunction() === FL2_FORWARD_BATCH_HANDLER).length;
}

function getForwardBatchStaleInfo_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(FL2_PROP_FORWARD_BATCH_STATE);
  if (!raw) return { isStale: false, reason: 'Состояние отсутствует' };

  let state = null;
  try {
    state = JSON.parse(raw);
  } catch (e) {
    return { isStale: true, reason: 'Состояние повреждено (не JSON)' };
  }

  const triggerCount = getForwardLoadBatchTriggerCount_();
  const updatedAt = Number(state && state.updatedAt ? state.updatedAt : state && state.startedAt ? state.startedAt : 0);
  const ageMs = updatedAt > 0 ? Date.now() - updatedAt : null;
  const staleByNoTrigger = triggerCount === 0;
  const staleByAge = ageMs !== null && ageMs > 15 * 60 * 1000;

  if (staleByNoTrigger) {
    return { isStale: true, reason: `Нет активных batch-триггеров (index=${Number(state && state.index || 0)})` };
  }
  if (staleByAge) {
    return { isStale: true, reason: `Состояние не обновлялось более 15 минут (ageMs=${ageMs})` };
  }

  return {
    isStale: false,
    reason: `Очередь активна: triggerCount=${triggerCount}, index=${Number(state && state.index || 0)}`,
  };
}


function clearForwardLoadBatchWatchdogTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction && t.getHandlerFunction() === FL2_FORWARD_BATCH_WATCHDOG_HANDLER)
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

function ensureForwardLoadBatchWatchdog_() {
  const exists = ScriptApp.getProjectTriggers().some((t) => t.getHandlerFunction && t.getHandlerFunction() === FL2_FORWARD_BATCH_WATCHDOG_HANDLER);
  if (exists) return;
  ScriptApp.newTrigger(FL2_FORWARD_BATCH_WATCHDOG_HANDLER)
    .timeBased()
    .everyMinutes(FL2_BATCH_WATCHDOG_INTERVAL_MIN)
    .create();
  fl2Log_('INFO', 'Пакетный план: watchdog включен', `interval=${FL2_BATCH_WATCHDOG_INTERVAL_MIN} мин`);
}

function touchForwardBatchHeartbeat_(props) {
  try {
    const raw = props.getProperty(FL2_PROP_FORWARD_BATCH_STATE);
    if (!raw) return;
    const state = JSON.parse(raw);
    state.updatedAt = Date.now();
    props.setProperty(FL2_PROP_FORWARD_BATCH_STATE, JSON.stringify(state));
  } catch (e) {
    // heartbeat best effort
  }
}

function runForwardLoadBatchWatchdog() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(FL2_PROP_FORWARD_BATCH_STATE);
  if (!raw) {
    clearForwardLoadBatchWatchdogTriggers_();
    return;
  }

  let state = null;
  try {
    state = JSON.parse(raw);
  } catch (e) {
    fl2Log_('ERROR', 'Пакетный план: watchdog обнаружил поврежденное состояние', String(e));
    clearForwardLoadBatchState_();
    return;
  }

  const updatedAt = Number(state && state.updatedAt ? state.updatedAt : state && state.startedAt ? state.startedAt : 0);
  const ageMs = updatedAt > 0 ? Date.now() - updatedAt : Number.MAX_SAFE_INTEGER;
  if (ageMs < FL2_BATCH_STALE_MS) return;

  const owner = props.getProperty(FL2_PROP_FORWARD_BATCH_LOCK_OWNER);
  fl2Log_('WARN', 'Пакетный план: watchdog обнаружил зависание', `ageMs=${ageMs}, owner=${owner || 'unknown'}`);

  props.deleteProperty(FL2_PROP_FORWARD_BATCH_LOCK_OWNER);
  clearForwardLoadBatchTriggers_();
  scheduleNextForwardLoadBatchStep_();
  fl2Log_('INFO', 'Пакетный план: watchdog', 'Шаговый триггер принудительно пересоздан');
}

function clearForwardLoadBatchState_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(FL2_PROP_FORWARD_BATCH_STATE);
  props.deleteProperty(FL2_PROP_FORWARD_BATCH_RUN_ID);
  props.deleteProperty(FL2_PROP_FORWARD_BATCH_LOCK_OWNER);
  clearForwardLoadBatchTriggers_();
  clearForwardLoadBatchWatchdogTriggers_();
}

function initForwardBatchServiceSheet_(ss, branches, runId) {
  const sheet = getOrCreateSheet_(ss, FL2_SHEET_BATCH_SERVICE);
  resetSheetForRender_(sheet);
  const header = ['Филиал', 'Статус', 'Время старта шага', 'Время окончания шага'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (branches && branches.length) {
    const rows = branches.map((b) => [str_(b), 'Ожидание', '', '']);
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  sheet.getRange(1, 6).setValue('runId');
  sheet.getRange(1, 7).setValue(runId || '');
}
function getEmployeesSetByHeader_(ss, headerNames, fallbackColIndex) {
  const sheet = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!sheet) throw new Error('Не найден лист "Сотрудники"');
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return new Set();

  const header = values[0].map((h) => str_(h).toLowerCase());
  const names = (headerNames || []).map((h) => str_(h).toLowerCase());
  let col = header.findIndex((h) => names.includes(h));
  if (col < 0 && Number.isInteger(fallbackColIndex) && fallbackColIndex >= 0) col = fallbackColIndex;
  if (col < 0) return new Set();

  return new Set(
    values
      .slice(1)
      .map((row) => str_(row[col]))
      .filter((name) => name)
  );
}

function getDoctorsSetFromEmployeesSheet_(ss) {
  return getEmployeesSetByHeader_(ss, ['врачи', 'врач'], 0);
}

function getMassageSetFromEmployeesSheet_(ss) {
  return getEmployeesSetByHeader_(ss, ['массажисты', 'массажист'], 1);
}

function getFtSetFromEmployeesSheet_(ss) {
  return getEmployeesSetByHeader_(ss, ['фт'], 2);
}

function getBranchesFromEmployeesSheet_(ss) {
  const sheet = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!sheet) throw new Error('Не найден лист "Сотрудники"');
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) throw new Error('Лист "Сотрудники" пуст');

  const header = values[0].map((h) => str_(h).toLowerCase());
  let branchCol = header.findIndex((h) => h === 'филиал');
  if (branchCol < 0) {
    throw new Error('В листе "Сотрудники" не найдена колонка "Филиал"');
  }

  const set = new Set();
  values.slice(1).forEach((row) => {
    const b = str_(row[branchCol]);
    if (b) set.add(b);
  });
  const branches = [...set].sort((a, b) => a.localeCompare(b, 'ru'));
  if (!branches.length) throw new Error('В колонке "Филиал" листа "Сотрудники" нет данных');
  return branches;
}

function collectBranches(scheduleRows, requestRows) {
  const set = new Set();
  scheduleRows.forEach((r) => {
    const b = str_(r['Филиал']);
    if (b) set.add(b);
  });
  requestRows.forEach((r) => {
    const b = str_(r['Филиал']);
    if (b) set.add(b);
  });
  return [...set].sort((a, b) => a.localeCompare(b, 'ru'));
}
function detectCurrentDayFromRawRequests_(rows, allowedEmployees, excludedEmployees, unassignedSet) {
  let maxDate = null;
  rows.forEach((row) => {
    if (!isCompletedStatus(str_(row['Состояние']))) return;
    const doctor = str_(row['Сотрудник']);
    if (!doctor || (excludedEmployees && excludedEmployees.has(doctor)) || hasStandaloneTestWord_(doctor)) return;
    if (allowedEmployees && !allowedEmployees.has(doctor)) {
      if (unassignedSet) unassignedSet.add(doctor);
      return;
    }
    const d = toDate_(row['ДатаНачала']);
    if (!d) return;
    if (!maxDate || d > maxDate) maxDate = d;
  });
  if (!maxDate) return null;
  return new Date(maxDate.getFullYear(), maxDate.getMonth(), maxDate.getDate());
}

function detectCurrentDay(requests) {
  const completed = requests.filter((r) => isCompletedStatus(r.status) && r.startDate);
  if (!completed.length) return null;
  completed.sort((a, b) => a.startDate - b.startDate);
  const max = completed[completed.length - 1].startDate;
  return new Date(max.getFullYear(), max.getMonth(), max.getDate());
}
function buildReceptionScheduleIndex(scheduleRows, excludedEmployees, allowedEmployees) {
  const map = new Map();
  scheduleRows.forEach((row) => {
    const date = toDate_(row['Дата']);
    const doctor = str_(row['Сотрудник']);
    if (!doctor || (excludedEmployees && excludedEmployees.has(doctor)) || hasStandaloneTestWord_(doctor)) return;
    if (allowedEmployees && !allowedEmployees.has(doctor)) return;
    const branch = str_(row['Филиал']);
    const type = str_(row['Тип времени']);
    const start = parseTimeToMinutes_(row['Время начала']);
    const end = parseTimeToMinutes_(row['Время окончания']);
    if (!date || !doctor || !branch || !type || start === null || end === null || end <= start) return;
    const key = `${dateKey_(date)}|${branch}|${doctor}`;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push({ type, start, end });
  });
  map.forEach((arr) => arr.sort((a, b) => a.start - b.start));
  return map;
}
function buildLatestRequests(rows, excludedEmployees, allowedEmployees, unassignedSet, rangeStart, rangeEnd) {
  const latest = new Map();
  rows.forEach((row, idx) => {
    const status = str_(row['Состояние']);
    if (!isAllowedProcessedRequestStatus(status)) return;
    const id = str_(row['Номер заявки']) || str_(row['Номер']) || str_(row['ID']) || `ROW-${idx}`;
    const statusDate = toDate_(row['Дата']) || toDate_(row['ДатаНачала']) || new Date(0);
    const statusSort = Number(Utilities.formatDate(statusDate, FL2_TIMEZONE, 'yyyyMMddHHmmss'));
    const prev = latest.get(id);
    if (!prev || statusSort > prev.statusSort || (statusSort === prev.statusSort && idx > prev.idx)) {
      const startDate = toDate_(row['ДатаНачала']);
      const endDate = toDate_(row['ДатаОкончания']);
      const timing = computeRequestTiming_(startDate, endDate);
      const startMinStored = toIntOrNull_(row['НачалоМин']);
      const endMinStored = toIntOrNull_(row['ОкончаниеМин']);
      const durationMinStored = toIntOrNull_(row['ДлительностьМин']);
      const doctor = str_(row['Сотрудник']);
      const client = str_(row['Клиент']);
      const cabinet = str_(row['Кабинет']);
      const cabinetType = detectCabinetType_(cabinet);
      const isCabinetRequest = Boolean(cabinetType);
      const blockedByCategory = doctor && allowedEmployees && !allowedEmployees.has(doctor);
      if (blockedByCategory && unassignedSet && doctor) unassignedSet.add(doctor);
      const skipByName = isCabinetRequest
        ? ((excludedEmployees && excludedEmployees.has(doctor)) || hasStandaloneTestWord_(doctor) || hasStandaloneTestWord_(client))
        : (!doctor || blockedByCategory || (excludedEmployees && excludedEmployees.has(doctor)) || hasStandaloneTestWord_(doctor) || hasStandaloneTestWord_(client));
      const nomenclature = str_(row['Номенклатура']).toLowerCase();
      const isDoctorVisit = nomenclature.includes('врача');
      const isPrimary = nomenclature.includes('первичный');
      latest.set(id, {
        idx,
        id,
        statusSort,
        status,
        doctor,
        branch: str_(row['Филиал']),
        startDate,
        endDate,
        startMin: startMinStored !== null ? startMinStored : timing.startMin,
        endMin: endMinStored !== null ? endMinStored : timing.endMin,
        durationMin: durationMinStored !== null ? durationMinStored : timing.durationMin,
        isDoctorVisit,
        isPrimary,
        cabinet,
        cabinetType,
        skipByName,
      });
    }
  });
  return [...latest.values()].filter((r) => !r.skipByName && (!rangeStart || !rangeEnd || isDateInRangeInclusive_(r.startDate, rangeStart, rangeEnd)));
}
function computeWindowMetrics(anchorDate, horizonDays, scheduleByDoctorDay, requests, options) {
  const startDate = addDays(anchorDate, 1);
  const endDate = addDays(anchorDate, horizonDays);
  const byBranch = new Map();
  const byDoctor = new Map();

  const includeAllNomenclature = Boolean(options && options.includeAllNomenclature);

  const requestIndex = new Map();
  requests.forEach((r) => {
    if (!isActivePlannedStatus(r.status) || !r.startDate) return;
    if (!includeAllNomenclature && !r.isDoctorVisit) return;
    const key = `${dateKey_(r.startDate)}|${r.branch}|${r.doctor}`;
    if (!requestIndex.has(key)) requestIndex.set(key, { intervals: [], reqCount: 0, primaryCount: 0 });
    const entry = requestIndex.get(key);
    entry.intervals.push(requestInterval(r));
    entry.reqCount += 1;
    if (r.isPrimary) entry.primaryCount += 1;
  });

  scheduleByDoctorDay.forEach((intervals, key) => {
    const [dKey, branch, doctor] = key.split('|');
    const day = toDate_(dKey);
    if (!day || day < startDate || day > endDate) return;

    const req = requestIndex.get(key) || { intervals: [], reqCount: 0, primaryCount: 0 };
    const reception = mergeIntervals_(intervals.filter((i) => i.type === 'Приемные часы'));
    const nonReception = mergeIntervals_(intervals.filter((i) => i.type !== 'Приемные часы'));
    const occupied = mergeIntervals_(req.intervals);
    const occupiedWithinReception = intersectIntervals(occupied, reception);
    const occupiedWithinNonReception = intersectIntervals(occupied, nonReception);
    const nonReceptionOccupiedMin = sumIntervalsMin(occupiedWithinNonReception);

    if (!reception.length && nonReceptionOccupiedMin <= 0) return;

    const free = subtractIntervals(reception, occupiedWithinReception);

    const receptionMin = sumIntervalsMin(reception) + nonReceptionOccupiedMin;
    const occupiedMin = sumIntervalsMin(occupiedWithinReception) + nonReceptionOccupiedMin;
    const freeMin = sumIntervalsMin(free);
    const freePP = free.reduce((acc, i) => acc + Math.floor((i.end - i.start) / 60), 0);

    accumulateForward(byBranch, branch, { receptionMin, occupiedMin, freeMin, freePP, reqCount: req.reqCount, primaryCount: req.primaryCount }, branch, '');
    accumulateForward(byDoctor, `${branch}|${doctor}`, { receptionMin, occupiedMin, freeMin, freePP, reqCount: req.reqCount, primaryCount: req.primaryCount }, branch, doctor);
  });
  return { byBranch, byDoctor };
}

function computeBaselineMetrics(currentDay, horizonDays, scheduleByDoctorDay, requests, options) {
  const byBranchSeries = new Map();
  const byDoctorSeries = new Map();
  const baselineRequests = requests.map((r) => ({
    ...r,
    status: isCompletedStatus(r.status) ? 'Запланирована' : r.status,
  }));

  for (let i = 1; i <= 42; i += 1) {
    const ref = addDays(currentDay, -i);
    if (ref.getDay() !== currentDay.getDay()) continue;
    const res = computeWindowMetrics(ref, horizonDays, scheduleByDoctorDay, baselineRequests, options);
    res.byBranch.forEach((v, k) => pushSeries(byBranchSeries, k, v, v.branch, ''));
    res.byDoctor.forEach((v, k) => pushSeries(byDoctorSeries, k, v, v.branch, v.doctor));
  }
  return {
    byBranch: finalizeSeriesAverage(byBranchSeries),
    byDoctor: finalizeSeriesAverage(byDoctorSeries),
  };
}
function pushSeries(seriesMap, key, value, branch, doctor) {
  if (!seriesMap.has(key)) seriesMap.set(key, { branch, doctor, values: [] });
  seriesMap.get(key).values.push(value);
}
function finalizeSeriesAverage(seriesMap) {
  const out = new Map();
  seriesMap.forEach((entry, key) => {
    const avg = emptyForwardMetric();
    if (!entry.values.length) {
      out.set(key, { branch: entry.branch, doctor: entry.doctor, metrics: avg });
      return;
    }
    entry.values.forEach((v) => {
      avg.receptionMin += v.metrics.receptionMin;
      avg.occupiedMin += v.metrics.occupiedMin;
      avg.freeMin += v.metrics.freeMin;
      avg.freePP += v.metrics.freePP;
      avg.reqCount += v.metrics.reqCount;
      avg.primaryCount += v.metrics.primaryCount;
    });
    const n = entry.values.length;
    Object.keys(avg).forEach((k) => { avg[k] /= n; });
    out.set(key, { branch: entry.branch, doctor: entry.doctor, metrics: avg });
  });
  return out;
}
function mergeForwardRows(current3, current7, base3, base7) {
  const keys = new Set([...current3.keys(), ...current7.keys(), ...base3.keys(), ...base7.keys()]);
  const rows = [];
  keys.forEach((key) => {
    const c3 = current3.get(key) || emptyForwardEntry(key);
    const c7 = current7.get(key) || emptyForwardEntry(key);
    const b3 = base3.get(key) || emptyForwardEntry(key);
    const b7 = base7.get(key) || emptyForwardEntry(key);
    const branch = c3.branch || c7.branch || b3.branch || b7.branch || key;
    const doctor = c3.doctor || c7.doctor || b3.doctor || b7.doctor || '';
    rows.push({ key, branch, doctor, c3: c3.metrics, c7: c7.metrics, b3: b3.metrics, b7: b7.metrics });
  });
  rows.sort((a, b) => {
    if (a.branch !== b.branch) return a.branch.localeCompare(b.branch, 'ru');
    return (a.doctor || '').localeCompare(b.doctor || '', 'ru');
  });
  return rows;
}
function renderForwardLoadSheet(ss, currentDay, selectedBranch, branchRows, doctorRows, massageSummary, ftSummary, cabinetSummary, massageRows, ftRows, cabinetRows) {
  const sheet = getOrCreateSheet_(ss, selectedBranch);
  resetSheetForRender_(sheet);
  let row = 1;
  sheet.getRange(row, 1, 1, 8)
    .merge()
    .setValue(`План загрузки на дату ${formatDateDDMMYYYY_(currentDay)} (без текущего дня)`)
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#d9ead3');
  row += 1;
  sheet.getRange(row, 1, 1, 8)
    .merge()
    .setValue(`Филиал: ${selectedBranch}`)
    .setFontWeight('bold')
    .setBackground('#eaf3ff');
  row += 1;
  sheet.getRange(row, 1, 1, 8)
    .merge()
    .setValue(`Сформировано: ${formatDateTimeDDMMYYYY_(new Date())}`)
    .setFontSize(10)
    .setBackground('#f3f8ff');
  row += 2;

  const branchRow = branchRows.length ? branchRows[0] : { branch: selectedBranch, c3: emptyForwardMetric(), c7: emptyForwardMetric(), b3: emptyForwardMetric(), b7: emptyForwardMetric() };
  const physioSummary = mergeCabinetTypeSummary_(cabinetSummary);

  row = renderPlanBlockByHorizon_(sheet, row, '3 дня', {
    doctorsSummary: { current: branchRow.c3, baseline: branchRow.b3 },
    massageSummary: { current: massageSummary && massageSummary.c3 ? massageSummary.c3 : emptyForwardMetric(), baseline: massageSummary && massageSummary.b3 ? massageSummary.b3 : emptyForwardMetric() },
    ftSummary: { current: ftSummary && ftSummary.c3 ? ftSummary.c3 : emptyForwardMetric(), baseline: ftSummary && ftSummary.b3 ? ftSummary.b3 : emptyForwardMetric() },
    physioSummary: { current: physioSummary.c3, baseline: physioSummary.b3 },
    doctorRows,
    massageRows,
    ftRows,
    cabinetRows,
    horizon: '3d',
    tableBase: 1,
  });

  row += 1;
  row = renderPlanBlockByHorizon_(sheet, row, '7 дней', {
    doctorsSummary: { current: branchRow.c7, baseline: branchRow.b7 },
    massageSummary: { current: massageSummary && massageSummary.c7 ? massageSummary.c7 : emptyForwardMetric(), baseline: massageSummary && massageSummary.b7 ? massageSummary.b7 : emptyForwardMetric() },
    ftSummary: { current: ftSummary && ftSummary.c7 ? ftSummary.c7 : emptyForwardMetric(), baseline: ftSummary && ftSummary.b7 ? ftSummary.b7 : emptyForwardMetric() },
    physioSummary: { current: physioSummary.c7, baseline: physioSummary.b7 },
    doctorRows,
    massageRows,
    ftRows,
    cabinetRows,
    horizon: '7d',
    tableBase: 5,
  });

  sheet.autoResizeColumns(1, 8);
}

function renderPlanBlockByHorizon_(sheet, row, horizonLabel, ctx) {
  const h = ctx.horizon;
  const summaryHeader = ['Категория', 'Приемные (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)', 'Заявки на ПП (шт)', 'Свободные слоты под ПП'];
  const titleRange = sheet.getRange(row, 1, 1, summaryHeader.length);
  titleRange.breakApart();
  titleRange
    .merge()
    .setValue(`Сводный план на ${horizonLabel}`)
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setBackground('#6fa8dc');
  row += 1;
  sheet.getRange(row, 1, 1, summaryHeader.length).setValues([summaryHeader]);
  styleForecastHeaderRow(sheet, row, summaryHeader.length);
  row += 1;

  row = renderSummaryCategoryRow_(sheet, row, 'Врачи', ctx.doctorsSummary.current, ctx.doctorsSummary.baseline, h, summaryHeader.length, true);
  row = renderSummaryCategoryRow_(sheet, row, 'Массажисты', ctx.massageSummary.current, ctx.massageSummary.baseline, h, summaryHeader.length, false);
  row = renderSummaryCategoryRow_(sheet, row, 'ФТ', ctx.ftSummary.current, ctx.ftSummary.baseline, h, summaryHeader.length, false);
  row = renderSummaryCategoryRow_(sheet, row, 'Физиотерапия', ctx.physioSummary.current, ctx.physioSummary.baseline, h, summaryHeader.length, false);

  row += 1;
  row = renderForwardDetailsTable_(sheet, row, `Таблица ${ctx.tableBase} — Врачи (${horizonLabel})`, 'Врач', ctx.doctorRows, h, true);
  row += 1;
  row = renderForwardDetailsTable_(sheet, row, `Таблица ${ctx.tableBase + 1} — Массажисты (${horizonLabel})`, 'Сотрудник', ctx.massageRows, h, false);
  row += 1;
  row = renderForwardDetailsTable_(sheet, row, `Таблица ${ctx.tableBase + 2} — ФТ (${horizonLabel})`, 'Сотрудник', ctx.ftRows, h, false);
  row += 1;
  row = renderCabinetDetailsTable_(sheet, row, `Таблица ${ctx.tableBase + 3} — Кабинеты физиотерапии (${horizonLabel})`, ctx.cabinetRows, h);

  return row;
}

function renderSummaryCategoryRow_(sheet, row, name, current, baseline, horizon, columns, includePrimaryCols) {
  const loadCur = current.receptionMin ? (current.occupiedMin / current.receptionMin) * 100 : 0;
  const loadBase = baseline.receptionMin ? (baseline.occupiedMin / baseline.receptionMin) * 100 : 0;
  const values = [
    name,
    round1_(current.receptionMin / 60),
    `${round1_(current.occupiedMin / 60)} (${round1_(baseline.occupiedMin / 60)})`,
    `${round0(loadCur)} (${round0(loadBase)})`,
    round1_(current.freeMin / 60),
    round1_(current.reqCount),
    includePrimaryCols ? round1_(current.primaryCount) : '',
    includePrimaryCols ? round1_(current.freePP) : '',
  ];
  sheet.getRange(row, 1, 1, columns).setValues([values]);
  styleForecastDataRow(sheet, row, columns);
  styleForecastRow(sheet, row, current, baseline, horizon, 4);
  return row + 1;
}

function renderForwardDetailsTable_(sheet, row, title, firstColTitle, rows, horizon, includePrimaryCols) {
  const includePrimary = includePrimaryCols !== false;
  const header = [firstColTitle, 'Приемные (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)'];
  if (includePrimary) {
    header.push('Заявки на ПП (шт)', 'Свободные слоты под ПП');
  }
  const columns = header.length;
  sheet.getRange(row, 1, 1, columns).merge().setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  row += 1;
  sheet.getRange(row, 1, 1, columns).setValues([header]);
  styleForecastHeaderRow(sheet, row, columns);
  row += 1;

  (rows || []).forEach((r) => {
    const current = horizon === '3d' ? r.c3 : r.c7;
    if (!current || !current.receptionMin) return;
    const baseline = horizon === '3d' ? r.b3 : r.b7;
    const loadCur = current.receptionMin ? (current.occupiedMin / current.receptionMin) * 100 : 0;
    const loadBase = baseline.receptionMin ? (baseline.occupiedMin / baseline.receptionMin) * 100 : 0;
    const values = [
      str_(r.doctor),
      round1_(current.receptionMin / 60),
      `${round1_(current.occupiedMin / 60)} (${round1_(baseline.occupiedMin / 60)})`,
      `${round0(loadCur)} (${round0(loadBase)})`,
      round1_(current.freeMin / 60),
      round1_(current.reqCount),
    ];
    if (includePrimary) {
      values.push(round1_(current.primaryCount), round1_(current.freePP));
    }
    sheet.getRange(row, 1, 1, columns).setValues([values]);
    styleForecastDataRow(sheet, row, columns);
    styleForecastRow(sheet, row, current, baseline, horizon, 4);
    row += 1;
  });

  return row;
}

function renderCabinetDetailsTable_(sheet, row, title, rows, horizon) {
  const columns = 6;
  const header = ['Кабинет', 'Приемные (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)'];
  sheet.getRange(row, 1, 1, columns).merge().setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  row += 1;
  sheet.getRange(row, 1, 1, columns).setValues([header]);
  styleForecastHeaderRow(sheet, row, columns);
  row += 1;

  (rows || []).forEach((r) => {
    const current = horizon === '3d' ? r.c3 : r.c7;
    if (!current || !current.receptionMin) return;
    const baseline = horizon === '3d' ? r.b3 : r.b7;
    const loadCur = current.receptionMin ? (current.occupiedMin / current.receptionMin) * 100 : 0;
    const loadBase = baseline.receptionMin ? (baseline.occupiedMin / baseline.receptionMin) * 100 : 0;
    const values = [
      str_(r.cabinet),
      round1_(current.receptionMin / 60),
      `${round1_(current.occupiedMin / 60)} (${round1_(baseline.occupiedMin / 60)})`,
      `${round0(loadCur)} (${round0(loadBase)})`,
      round1_(current.freeMin / 60),
      round1_(current.reqCount),
    ];
    sheet.getRange(row, 1, 1, columns).setValues([values]);
    styleForecastDataRow(sheet, row, columns);
    styleForecastRow(sheet, row, current, baseline, horizon, 4);
    row += 1;
  });

  return row;
}

function renderForwardRowsSingle(sheet, startRow, rows, horizon) {
  let row = startRow;
  rows.forEach((r) => {
    const current = horizon === '3d' ? r.c3 : r.c7;
    if (!current.receptionMin) return;
    const baseline = horizon === '3d' ? r.b3 : r.b7;
    const horizonLabel = horizon === '3d' ? '3 дня' : '7 дней';
    const values = toForwardOutputRowSingle(r.branch, r.doctor, horizonLabel, current, baseline);
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    styleForecastDataRow(sheet, row, values.length);
    styleForecastRow(sheet, row, current, baseline, horizon);
    row += 1;
  });
  return row;
}
function toForwardOutputRowSingle(branch, doctor, horizonLabel, current, baseline) {
  const currentLoad = current.receptionMin ? (current.occupiedMin / current.receptionMin) * 100 : 0;
  const usualLoad = baseline.receptionMin ? (baseline.occupiedMin / baseline.receptionMin) * 100 : 0;
  return [
    doctor || horizonLabel,
    round1_(current.receptionMin / 60),
    `${round1_(current.occupiedMin / 60)} (${round1_(baseline.occupiedMin / 60)})`,
    `${round1_(current.primaryCount)} (${round1_(baseline.primaryCount)})`,
    `${round0(currentLoad)} (${round0(usualLoad)})`,
    round1_(current.freeMin / 60),
    round1_(current.freePP),
    round1_(current.reqCount),
  ];
}
function styleForecastHeaderRow(sheet, row, columns) {
  sheet.getRange(row, 1, 1, columns)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setBorder(true, true, true, true, true, true);
}
function styleForecastDataRow(sheet, row, columns) {
  const range = sheet.getRange(row, 1, 1, columns);
  range.setBorder(true, true, true, true, true, true);
  sheet.getRange(row, 1).setBackground('#fafafa');
}
function styleForecastRow(sheet, row, current, baseline, horizon, loadColumn) {
  const loadCur = current.receptionMin ? (current.occupiedMin / current.receptionMin) * 100 : 0;
  const targetCol = loadColumn || 5;
  sheet.getRange(row, targetCol).setBackground(forecastLoadColorByHorizon(loadCur, horizon));
}
function forecastLoadColorByHorizon(loadPercent, horizon) {
  if (horizon === '7d') {
    if (loadPercent < 20) return '#ff0000';
    if (loadPercent < 35) return '#fff2cc';
    if (loadPercent > 60) return '#f4cccc';
    if (loadPercent > 50) return '#d9ead3';
    return '#ffffff';
  }
  if (loadPercent < 25) return '#ff0000';
  if (loadPercent < 40) return '#fff2cc';
  if (loadPercent > 75) return '#f4cccc';
  if (loadPercent > 60) return '#d9ead3';
  return '#ffffff';
}
function sendForwardReportEmail(ss, selectedBranch, currentDay) {
  const sheet = ss.getSheetByName(selectedBranch);
  const mailingSheet = ss.getSheetByName(FL2_SHEET_MAILING);
  if (!sheet || !mailingSheet) {
    throw new Error('Не найден лист отчета филиала или лист Отправка');
  }
  const mailingRows = getSheetObjects_(mailingSheet);
  const target = mailingRows.find((r) => str_(r['Филиал']) === selectedBranch);
  if (!target) throw new Error(`Для филиала "${selectedBranch}" нет строки в листе Отправка`);
  const to = parseEmailList(target['Кому']);
  const cc = parseEmailList(target['Копия']);
  if (!to.length) throw new Error(`Для филиала "${selectedBranch}" не заполнено поле Кому`);
  const subject = `План загрузки на дату ${formatDateDDMMYYYY_(currentDay)}`;
  const body = [
    sheetToHtmlTable(sheet),
    '<p><b>Цветовая интерпретация загрузки:</b> для горизонта 3 дня: <span style="background:#ff0000;padding:2px 6px;">красный</span> &lt; 25%, <span style="background:#fff2cc;padding:2px 6px;">желтый</span> 25–40%, белый 40–60%, <span style="background:#d9ead3;padding:2px 6px;">зеленый</span> 60–75%, <span style="background:#f4cccc;padding:2px 6px;">бледно-розовый</span> > 75%. Для горизонта 7 дней: <span style="background:#ff0000;padding:2px 6px;">красный</span> &lt; 20%, <span style="background:#fff2cc;padding:2px 6px;">желтый</span> 20–35%, белый 35–50%, <span style="background:#d9ead3;padding:2px 6px;">зеленый</span> 50–60%, <span style="background:#f4cccc;padding:2px 6px;">бледно-розовый</span> > 60%.</p>',
    '<p><b>Пояснение:</b> значения в скобках — это обычный уровень (среднее за последние 6 недель по тому же дню недели).</p>',
    '<p><b>Пояснение:</b> Приемные (ч) — часы из графика врача; Занятые (ч) — часы занятых заявками; Загрузка (%) = Занятые/Приемные; Свободные (ч) = Приемные − Занятые; Свободные ПП — количество свободных часовых слотов под первичный прием.</p>',
    buildMailingServiceInfoHtml_(ss, selectedBranch, to, cc, sheet),
  ].join('');
  MailApp.sendEmail({
    to: to.join(','),
    cc: cc.join(','),
    subject,
    htmlBody: body,
  });
}
function buildOverallForecastReport() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  ensureDoctorScheduleFresh(ss);
  const sheets = ss.getSheets();
  const clinicRows = [];
  let reportDate = '';
  sheets.forEach((sheet) => {
    const data = sheet.getDataRange().getDisplayValues();
    if (!data.length) return;
    if (!isBranchPlanSheetData_(data)) return;
    const parsed = parseClinicForecastSheet(data);
    if (!parsed) return;
    if (!reportDate && parsed.reportDate) reportDate = parsed.reportDate;
    clinicRows.push(...parsed.rows);
  });
  if (!clinicRows.length) {
    throw new Error('Не найдены листы филиалов с планом для сборки общего отчета');
  }
  renderOverallForecastSheet(ss, reportDate || formatDateDDMMYYYY_(new Date()), clinicRows);
  const sendAnswer = ui.alert('Общий план', 'Сделать рассылку общего плана?', ui.ButtonSet.YES_NO);
  if (sendAnswer === ui.Button.YES) {
    sendOverallForecastEmail(ss, reportDate || formatDateDDMMYYYY_(new Date()));
  }
  SpreadsheetApp.getActive().toast('Общий план собран', MENU_NAME, 5);
}

function isBranchPlanSheetData_(data) {
  if (!data || !data.length) return false;
  const firstCol = data.map((r) => str_(r && r[0]));
  const hasPlanHeader = firstCol.some((v) => v.startsWith('План загрузки на дату'));
  const hasBranchHeader = firstCol.some((v) => v.startsWith('Филиал:'));
  return hasPlanHeader && hasBranchHeader;
}

function parseClinicForecastSheet(data) {
  const firstCol = (data || []).map((r) => str_(r && r[0]));
  const planHeader = firstCol.find((v) => v.startsWith('План загрузки на дату')) || '';
  const dateMatch = planHeader.match(/(\d{2}\.\d{2}\.\d{4})/);
  const reportDate = dateMatch ? dateMatch[1] : '';
  const branchLine = firstCol.find((v) => v.startsWith('Филиал:')) || '';
  const branch = branchLine.replace(/^Филиал:\s*/i, '').trim();
  if (!branch) return null;

  let currentCategory = 'doctor';
  let currentHorizon = '';
  let cabinetLegacyMode = false;
  const rows = [];

  const pushRow = (category, horizon, reception, occupiedVal, loadVal, free, requests, cabinetType, primaryCount, freePP) => {
    const occupied = parseCombinedMetric(occupiedVal);
    const load = parseCombinedMetric(loadVal);
    rows.push({
      category,
      branch,
      cabinetType: cabinetType || '',
      horizon,
      reception: parseLocaleNumber(reception),
      occupiedCur: occupied.current,
      occupiedUsual: occupied.usual,
      loadCur: load.current,
      loadUsual: load.usual,
      free: parseLocaleNumber(free),
      freePP: parseLocaleNumber(freePP),
      requests: parseLocaleNumber(requests),
      primaryCount: parseLocaleNumber(primaryCount),
    });
  };

  data.forEach((row) => {
    const first = str_(row[0]);
    if (!first) return;

    if (first.includes('Блок 3 дня') || first.includes('Сводный план на 3 дня')) currentHorizon = '3 дня';
    if (first.includes('Блок 7 дней') || first.includes('Сводный план на 7 дней')) currentHorizon = '7 дней';

    if (first.includes('Таблица 1 — Свод по филиалу')) currentCategory = 'doctor';
    if (first.includes('Свод по массажистам')) currentCategory = 'massage';
    if (first.includes('Свод по ФТ')) currentCategory = 'ft';
    if (first.includes('Кабинеты ФТЛ:')) currentCategory = 'cabinet_ftl';
    if (first.includes('Кабинеты Магнит:')) currentCategory = 'cabinet_magnet';
    if (first.includes('Загрузка кабинетов ФТЛ/Магнит')) {
      currentCategory = 'cabinet_legacy';
      cabinetLegacyMode = true;
    }

    // Новый формат: свод по категориям внутри блока горизонта
    if (currentHorizon && (first === 'Врачи' || first === 'Массажисты' || first === 'ФТ' || first === 'Физиотерапия')) {
      const category = first === 'Врачи'
        ? 'doctor'
        : (first === 'Массажисты' ? 'massage' : (first === 'ФТ' ? 'ft' : 'physio'));
      pushRow(category, currentHorizon, row[1], row[2], row[3], row[4], row[5], '', row[6], row[7]);
      return;
    }

    // Новый формат: общий лист non-doctor, категория встроена в колонку филиала
    if ((first.includes('— Массажисты и ФТ') || first.includes('— Физиотерапия')) && (currentHorizon === '3 дня' || currentHorizon === '7 дней')) {
      const category = first.includes('Физиотерапия') ? 'physio' : 'massage';
      pushRow(category, currentHorizon, row[1], row[2], row[3], row[4], row[5], '', row[6], row[7]);
      return;
    }

    // Legacy cabinet separate tables
    if (currentCategory === 'cabinet_ftl' || currentCategory === 'cabinet_magnet') {
      const horizon = str_(row[0]);
      if (horizon !== '3 дня' && horizon !== '7 дней') return;
      pushRow('cabinet', horizon, row[1], row[2], row[3], row[4], row[5], currentCategory === 'cabinet_ftl' ? 'ФТЛ' : 'Магнит', 0, 0);
      return;
    }

    if (currentCategory === 'cabinet_legacy' && cabinetLegacyMode) {
      const horizon = str_(row[1]);
      if (horizon !== '3 дня' && horizon !== '7 дней') return;
      const type = str_(row[0]);
      if (type !== 'ФТЛ' && type !== 'Магнит') return;
      pushRow('cabinet', horizon, 0, row[2], row[3], row[4], row[5], type, 0, 0);
      return;
    }

    // Старый формат (horizon в первом столбце)
    const horizon = first;
    if (horizon !== '3 дня' && horizon !== '7 дней') return;

    if (currentCategory === 'doctor') {
      const occupied = parseCombinedMetric(row[2]);
      const load = parseCombinedMetric(row[4]);
      rows.push({
        category: currentCategory,
        branch,
        horizon,
        reception: parseLocaleNumber(row[1]),
        occupiedCur: occupied.current,
        occupiedUsual: occupied.usual,
        loadCur: load.current,
        loadUsual: load.usual,
        free: parseLocaleNumber(row[5]),
        freePP: parseLocaleNumber(row[6]),
        requests: parseLocaleNumber(row[7]),
        primaryCount: parseLocaleNumber(row[8]),
      });
      return;
    }

    pushRow(currentCategory, horizon, row[1], row[2], row[3], row[4], row[5], '', row[6], row[7]);
  });

  return rows.length ? { reportDate, rows } : null;
}

function parseCombinedMetric(value) {
  const text = str_(value);
  const m = text.match(/(-?[\d.,]+)\s*\((-?[\d.,]+)\)/);
  if (!m) return { current: parseLocaleNumber(text), usual: 0 };
  return {
    current: parseLocaleNumber(m[1]),
    usual: parseLocaleNumber(m[2]),
  };
}
function parseLocaleNumber(value) {
  const normalized = str_(value).replace(/\s/g, '').replace(',', '.');
  const num = Number(normalized);
  return Number.isFinite(num) ? num : 0;
}
function resetSheetForRender_(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (maxRows > 0 && maxCols > 0) {
    const allRange = sheet.getRange(1, 1, maxRows, maxCols);
    allRange.breakApart();
    allRange.clearContent();
    allRange.clearFormat();
  }
}

function renderOverallForecastSheet(ss, reportDate, clinicRows) {
  const sheet = getOrCreateSheet_(ss, FL2_SHEET_OVERALL_FORECAST);
  resetSheetForRender_(sheet);
  let row = 1;
  const columns = 7;
  sheet.getRange(row, 1, 1, columns).merge().setValue(`Общий план по клиникам на ${reportDate}`).setFontWeight('bold').setBackground('#d9ead3');
  row += 1;
  sheet.getRange(row, 1, 1, columns).merge().setValue(`Сформировано: ${formatDateTimeDDMMYYYY_(new Date())}`).setFontSize(10).setBackground('#f3f8ff');
  row += 2;

  sheet.getRange(row, 1, 1, columns).merge().setValue('Блок 1 — Горизонт 3 дня').setFontWeight('bold').setBackground('#eaf3ff');
  row += 1;
  row = renderOverallHorizonTable(sheet, row, clinicRows.filter((r) => str_(r.category || 'doctor') === 'doctor'), '3 дня', 'Таблица 1 — Врачи: общий план на 3 дня');
  row += 1;
  row = renderOverallUnifiedAuxTable_(sheet, row, clinicRows.filter((r) => str_(r.category) === 'massage' || str_(r.category) === 'ft'), '3 дня', 'Таблица 2 — Массажисты и ФТ: общий план на 3 дня');
  row += 1;
  row = renderOverallUnifiedAuxTable_(sheet, row, clinicRows.filter((r) => str_(r.category) === 'cabinet' || str_(r.category) === 'physio'), '3 дня', 'Таблица 3 — Физиотерапия: общий план на 3 дня');

  row += 1;
  sheet.getRange(row, 1, 1, columns).merge().setValue('Блок 2 — Горизонт 7 дней').setFontWeight('bold').setBackground('#eaf3ff');
  row += 1;
  row = renderOverallHorizonTable(sheet, row, clinicRows.filter((r) => str_(r.category || 'doctor') === 'doctor'), '7 дней', 'Таблица 4 — Врачи: общий план на 7 дней');
  row += 1;
  row = renderOverallUnifiedAuxTable_(sheet, row, clinicRows.filter((r) => str_(r.category) === 'massage' || str_(r.category) === 'ft'), '7 дней', 'Таблица 5 — Массажисты и ФТ: общий план на 7 дней');
  row += 1;
  row = renderOverallUnifiedAuxTable_(sheet, row, clinicRows.filter((r) => str_(r.category) === 'cabinet' || str_(r.category) === 'physio'), '7 дней', 'Таблица 6 — Физиотерапия: общий план на 7 дней');

  sheet.autoResizeColumns(1, columns);
}

function renderOverallUnifiedAuxTable_(sheet, row, clinicRows, horizonLabel, title) {
  const columns = 6;
  const header = ['Клиника', 'Приемные (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)'];

  sheet.getRange(row, 1, 1, columns).merge().setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  row += 1;
  sheet.getRange(row, 1, 1, columns).setValues([header]);
  styleForecastHeaderRow(sheet, row, columns);
  row += 1;

  const aggregated = new Map();
  (clinicRows || [])
    .filter((r) => r.horizon === horizonLabel)
    .forEach((r) => {
      const key = str_(r.branch);
      if (!aggregated.has(key)) {
        aggregated.set(key, {
          branch: str_(r.branch),
          reception: 0,
          occupiedCur: 0,
          occupiedUsual: 0,
          free: 0,
          requests: 0,
        });
      }
      const item = aggregated.get(key);
      item.reception += Number(r.reception) || 0;
      item.occupiedCur += Number(r.occupiedCur) || 0;
      item.occupiedUsual += Number(r.occupiedUsual) || 0;
      item.free += Number(r.free) || 0;
      item.requests += Number(r.requests) || 0;
    });

  const subset = [...aggregated.values()].sort((a, b) => {
    return a.branch.localeCompare(b.branch, 'ru');
  });

  subset.forEach((r) => {
    const loadCur = r.reception ? (r.occupiedCur / r.reception) * 100 : 0;
    const loadUsual = r.reception ? (r.occupiedUsual / r.reception) * 100 : 0;
    const values = [
      r.branch,
      round1_(r.reception),
      `${round1_(r.occupiedCur)} (${round1_(r.occupiedUsual)})`,
      `${round0(loadCur)} (${round0(loadUsual)})`,
      round1_(r.free),
      round1_(r.requests),
    ];
    sheet.getRange(row, 1, 1, columns).setValues([values]);
    styleForecastDataRow(sheet, row, columns);
    styleForecastRow(sheet, row, { occupiedMin: r.occupiedCur * 60, receptionMin: Math.max(1, r.reception * 60) }, { occupiedMin: r.occupiedUsual * 60, receptionMin: Math.max(1, r.reception * 60) }, horizonLabel === '7 дней' ? '7d' : '3d', 4);
    row += 1;
  });

  const reception = subset.reduce((a, r) => a + r.reception, 0);
  const occupiedCur = subset.reduce((a, r) => a + r.occupiedCur, 0);
  const occupiedUsual = subset.reduce((a, r) => a + r.occupiedUsual, 0);
  const weightedCur = reception ? (occupiedCur / reception) * 100 : 0;
  const weightedUsual = reception ? (occupiedUsual / reception) * 100 : 0;
  const free = subset.reduce((a, r) => a + r.free, 0);
  const requests = subset.reduce((a, r) => a + r.requests, 0);

  const totalValues = [
    'ИТОГО',
    round1_(reception),
    `${round1_(occupiedCur)} (${round1_(occupiedUsual)})`,
    `${round0(weightedCur)} (${round0(weightedUsual)})`,
    round1_(free),
    round1_(requests),
  ];
  sheet.getRange(row, 1, 1, columns).setValues([totalValues]);
  styleForecastDataRow(sheet, row, columns);
  styleForecastRow(sheet, row, { occupiedMin: occupiedCur * 60, receptionMin: Math.max(1, reception * 60) }, { occupiedMin: occupiedUsual * 60, receptionMin: Math.max(1, reception * 60) }, horizonLabel === '7 дней' ? '7d' : '3d', 4);
  sheet.getRange(row, 1, 1, columns).setFontWeight('bold').setBackground('#f3f3f3');
  return row + 1;
}

function renderOverallHorizonTable(sheet, row, clinicRows, horizonLabel, title) {
  const columns = 8;
  const header = ['Клиника', 'Приемные (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)', 'Заявки на ПП (шт)', 'Свободные слоты под ПП'];
  sheet.getRange(row, 1, 1, columns).merge().setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  row += 1;
  sheet.getRange(row, 1, 1, columns).setValues([header]);
  styleForecastHeaderRow(sheet, row, columns);
  row += 1;
  const subset = clinicRows
    .filter((r) => r.horizon === horizonLabel)
    .sort((a, b) => a.branch.localeCompare(b.branch, 'ru'));
  subset.forEach((r) => {
    const values = [
      r.branch,
      round1_(r.reception),
      `${round1_(r.occupiedCur)} (${round1_(r.occupiedUsual)})`,
      `${round0(r.loadCur)} (${round0(r.loadUsual)})`,
      round1_(r.free),
      round1_(r.requests),
      round1_(r.primaryCount),
      round1_(r.freePP),
    ];
    sheet.getRange(row, 1, 1, columns).setValues([values]);
    styleForecastDataRow(sheet, row, columns);
    styleForecastRow(sheet, row, { occupiedMin: r.occupiedCur * 60, receptionMin: Math.max(1, r.reception * 60) }, { occupiedMin: r.occupiedUsual * 60, receptionMin: Math.max(1, r.reception * 60) }, horizonLabel === '7 дней' ? '7d' : '3d', 4);
    row += 1;
  });
  const reception = subset.reduce((a, r) => a + r.reception, 0);
  const occupiedCur = subset.reduce((a, r) => a + r.occupiedCur, 0);
  const occupiedUsual = subset.reduce((a, r) => a + r.occupiedUsual, 0);
  const weightedCur = reception ? subset.reduce((a, r) => a + (r.loadCur * r.reception), 0) / reception : 0;
  const weightedUsual = reception ? subset.reduce((a, r) => a + (r.loadUsual * r.reception), 0) / reception : 0;
  const free = subset.reduce((a, r) => a + r.free, 0);
  const requests = subset.reduce((a, r) => a + r.requests, 0);
  const primaryCount = subset.reduce((a, r) => a + (r.primaryCount || 0), 0);
  const freePP = subset.reduce((a, r) => a + r.freePP, 0);
  const totalValues = [
    'ИТОГО',
    round1_(reception),
    `${round1_(occupiedCur)} (${round1_(occupiedUsual)})`,
    `${round0(weightedCur)} (${round0(weightedUsual)})`,
    round1_(free),
    round1_(requests),
    round1_(primaryCount),
    round1_(freePP),
  ];
  sheet.getRange(row, 1, 1, columns).setValues([totalValues]);
  styleForecastDataRow(sheet, row, columns);
  styleForecastRow(sheet, row, { occupiedMin: occupiedCur * 60, receptionMin: Math.max(1, reception * 60) }, { occupiedMin: occupiedUsual * 60, receptionMin: Math.max(1, reception * 60) }, horizonLabel === '7 дней' ? '7d' : '3d', 4);
  sheet.getRange(row, 1, 1, columns).setFontWeight('bold').setBackground('#f3f3f3');
  return row + 1;
}
function aggregateForwardCategorySummary_(rowsDoctorAll, employeeSet) {
  const set = employeeSet || new Set();
  const out = {
    c3: emptyForwardMetric(),
    c7: emptyForwardMetric(),
    b3: emptyForwardMetric(),
    b7: emptyForwardMetric(),
  };
  (rowsDoctorAll || []).forEach((r) => {
    if (!set.has(str_(r.doctor))) return;
    addForwardMetric(out.c3, r.c3 || emptyForwardMetric());
    addForwardMetric(out.c7, r.c7 || emptyForwardMetric());
    addForwardMetric(out.b3, r.b3 || emptyForwardMetric());
    addForwardMetric(out.b7, r.b7 || emptyForwardMetric());
  });
  return out;
}

function toForwardOutputRowLite_(horizonLabel, current, baseline) {
  const currentLoad = current.receptionMin ? (current.occupiedMin / current.receptionMin) * 100 : 0;
  const usualLoad = baseline.receptionMin ? (baseline.occupiedMin / baseline.receptionMin) * 100 : 0;
  return [
    horizonLabel,
    round1_(current.receptionMin / 60),
    `${round1_(current.occupiedMin / 60)} (${round1_(baseline.occupiedMin / 60)})`,
    `${round0(currentLoad)} (${round0(usualLoad)})`,
    round1_(current.freeMin / 60),
    round1_(current.reqCount),
  ];
}

function renderForwardCategorySummaryTable_(sheet, row, title, summary) {
  const header = ['Горизонт', 'Приемные (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)'];
  sheet.getRange(row, 1, 1, header.length).merge().setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  row += 1;
  sheet.getRange(row, 1, 1, header.length).setValues([header]);
  styleForecastHeaderRow(sheet, row, header.length);
  row += 1;

  const c3 = summary && summary.c3 ? summary.c3 : emptyForwardMetric();
  const c7 = summary && summary.c7 ? summary.c7 : emptyForwardMetric();
  const b3 = summary && summary.b3 ? summary.b3 : emptyForwardMetric();
  const b7 = summary && summary.b7 ? summary.b7 : emptyForwardMetric();

  const row3 = toForwardOutputRowLite_('3 дня', c3, b3);
  sheet.getRange(row, 1, 1, header.length).setValues([row3]);
  styleForecastDataRow(sheet, row, header.length);
  styleForecastRow(sheet, row, c3, b3, '3d', 4);
  row += 1;

  const row7 = toForwardOutputRowLite_('7 дней', c7, b7);
  sheet.getRange(row, 1, 1, header.length).setValues([row7]);
  styleForecastDataRow(sheet, row, header.length);
  styleForecastRow(sheet, row, c7, b7, '7d', 4);
  row += 1;

  return row;
}

function renderOverallCategoryCombinedTable_(sheet, row, clinicRows, title) {
  const columns = 7;
  const header = ['Клиника', 'Горизонт', 'Приемные (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)'];
  sheet.getRange(row, 1, 1, columns).merge().setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  row += 1;
  sheet.getRange(row, 1, 1, columns).setValues([header]);
  styleForecastHeaderRow(sheet, row, columns);
  row += 1;

  const horizonOrder = { '3 дня': 0, '7 дней': 1 };
  const subset = (clinicRows || [])
    .filter((r) => r.horizon === '3 дня' || r.horizon === '7 дней')
    .sort((a, b) => {
      if (a.branch !== b.branch) return a.branch.localeCompare(b.branch, 'ru');
      return (horizonOrder[a.horizon] || 99) - (horizonOrder[b.horizon] || 99);
    });

  subset.forEach((r) => {
    const values = [
      r.branch,
      r.horizon,
      round1_(r.reception),
      `${round1_(r.occupiedCur)} (${round1_(r.occupiedUsual)})`,
      `${round0(r.loadCur)} (${round0(r.loadUsual)})`,
      round1_(r.free),
      round1_(r.requests),
    ];
    sheet.getRange(row, 1, 1, columns).setValues([values]);
    styleForecastDataRow(sheet, row, columns);
    styleForecastRow(sheet, row, { occupiedMin: r.occupiedCur * 60, receptionMin: Math.max(1, r.reception * 60) }, { occupiedMin: r.occupiedUsual * 60, receptionMin: Math.max(1, r.reception * 60) }, r.horizon === '7 дней' ? '7d' : '3d', 5);
    row += 1;
  });

  return row;
}


function mergeCabinetTypeSummary_(summary) {
  const ftl = summary && summary.ftl ? summary.ftl : { c3: emptyForwardMetric(), c7: emptyForwardMetric(), b3: emptyForwardMetric(), b7: emptyForwardMetric() };
  const magnet = summary && summary.magnet ? summary.magnet : { c3: emptyForwardMetric(), c7: emptyForwardMetric(), b3: emptyForwardMetric(), b7: emptyForwardMetric() };
  const out = { c3: emptyForwardMetric(), c7: emptyForwardMetric(), b3: emptyForwardMetric(), b7: emptyForwardMetric() };
  addForwardMetric(out.c3, ftl.c3 || emptyForwardMetric());
  addForwardMetric(out.c3, magnet.c3 || emptyForwardMetric());
  addForwardMetric(out.c7, ftl.c7 || emptyForwardMetric());
  addForwardMetric(out.c7, magnet.c7 || emptyForwardMetric());
  addForwardMetric(out.b3, ftl.b3 || emptyForwardMetric());
  addForwardMetric(out.b3, magnet.b3 || emptyForwardMetric());
  addForwardMetric(out.b7, ftl.b7 || emptyForwardMetric());
  addForwardMetric(out.b7, magnet.b7 || emptyForwardMetric());
  return out;
}

function buildCabinetRowsByName_(currentDay, scheduleRows, requests) {
  const c3 = computeCabinetByNameWindowMetrics_(currentDay, 3, scheduleRows, requests);
  const c7 = computeCabinetByNameWindowMetrics_(currentDay, 7, scheduleRows, requests);
  const b3 = computeCabinetByNameBaselineMetrics_(currentDay, 3, scheduleRows, requests);
  const b7 = computeCabinetByNameBaselineMetrics_(currentDay, 7, scheduleRows, requests);
  const keys = new Set([...c3.keys(), ...c7.keys(), ...b3.keys(), ...b7.keys()]);
  return [...keys].map((cabinet) => ({
    cabinet,
    c3: c3.get(cabinet) || emptyForwardMetric(),
    c7: c7.get(cabinet) || emptyForwardMetric(),
    b3: b3.get(cabinet) || emptyForwardMetric(),
    b7: b7.get(cabinet) || emptyForwardMetric(),
  })).sort((a, b) => str_(a.cabinet).localeCompare(str_(b.cabinet), 'ru'));
}

function buildCabinetScheduleByNameIndex_(scheduleRows) {
  const map = new Map();
  (scheduleRows || []).forEach((row) => {
    const d = toDate_(row['Дата']);
    const cabinet = str_(row['Кабинет']);
    const type = detectCabinetType_(cabinet);
    const timeType = str_(row['Тип времени']);
    const start = parseTimeToMinutes_(row['Время начала']);
    const end = parseTimeToMinutes_(row['Время окончания']);
    if (!d || !type || !cabinet || timeType !== 'Приемные часы' || start === null || end === null || end <= start) return;
    const key = `${dateKey_(d)}|${cabinet}`;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push({ start, end });
  });
  map.forEach((arr, key) => map.set(key, mergeIntervals_(arr)));
  return map;
}

function buildCabinetRequestByNameIndex_(requests, startDate, endDate) {
  const map = new Map();
  (requests || []).forEach((r) => {
    if (!isActivePlannedStatus(r.status) || !r.startDate) return;
    if (startDate && endDate && (r.startDate < startDate || r.startDate > endDate)) return;
    const cabinet = str_(r.cabinet || r['Кабинет']);
    const type = detectCabinetType_(r.cabinetType || cabinet);
    if (!type || !cabinet) return;
    const key = `${dateKey_(r.startDate)}|${cabinet}`;
    if (!map.has(key)) map.set(key, { intervals: [], reqCount: 0 });
    const e = map.get(key);
    e.intervals.push(requestInterval(r));
    e.reqCount += 1;
  });
  return map;
}

function computeCabinetByNameWindowMetrics_(anchorDate, horizonDays, scheduleRows, requests) {
  const startDate = addDays(anchorDate, 1);
  const endDate = addDays(anchorDate, horizonDays);
  const scheduleIndex = buildCabinetScheduleByNameIndex_(scheduleRows);
  const requestIndex = buildCabinetRequestByNameIndex_(requests, startDate, endDate);
  const out = new Map();
  const touched = new Set();

  scheduleIndex.forEach((workIntervals, key) => {
    const [dKey, cabinet] = key.split('|');
    const day = toDate_(dKey);
    if (!day || day < startDate || day > endDate) return;
    const req = requestIndex.get(key) || { intervals: [], reqCount: 0 };
    touched.add(key);
    const occupied = mergeIntervals_(req.intervals);
    const occupiedWithinWork = intersectIntervals(occupied, workIntervals);
    const free = subtractIntervals(workIntervals, occupiedWithinWork);
    if (!out.has(cabinet)) out.set(cabinet, emptyForwardMetric());
    addForwardMetric(out.get(cabinet), {
      receptionMin: sumIntervalsMin(workIntervals),
      occupiedMin: sumIntervalsMin(occupiedWithinWork),
      freeMin: sumIntervalsMin(free),
      freePP: 0,
      reqCount: req.reqCount,
      primaryCount: 0,
    });
  });

  requestIndex.forEach((req, key) => {
    if (touched.has(key)) return;
    const [, cabinet] = key.split('|');
    if (!out.has(cabinet)) out.set(cabinet, emptyForwardMetric());
    addForwardMetric(out.get(cabinet), {
      receptionMin: 0,
      occupiedMin: sumIntervalsMin(mergeIntervals_(req.intervals)),
      freeMin: 0,
      freePP: 0,
      reqCount: req.reqCount,
      primaryCount: 0,
    });
  });

  return out;
}

function computeCabinetByNameBaselineMetrics_(currentDay, horizonDays, scheduleRows, requests) {
  const scheduleIndex = buildCabinetScheduleByNameIndex_(scheduleRows);
  const baselineRequests = (requests || []).map((r) => ({ ...r, status: isCompletedStatus(r.status) ? 'Запланирована' : r.status }));
  const series = new Map();

  for (let i = 1; i <= 42; i += 1) {
    const ref = addDays(currentDay, -i);
    if (ref.getDay() !== currentDay.getDay()) continue;
    const startDate = addDays(ref, 1);
    const endDate = addDays(ref, horizonDays);
    const reqIndex = buildCabinetRequestByNameIndex_(baselineRequests, startDate, endDate);
    const win = new Map();
    const touched = new Set();

    scheduleIndex.forEach((workIntervals, key) => {
      const [dKey, cabinet] = key.split('|');
      const day = toDate_(dKey);
      if (!day || day < startDate || day > endDate) return;
      const req = reqIndex.get(key) || { intervals: [], reqCount: 0 };
      touched.add(key);
      const occupied = mergeIntervals_(req.intervals);
      const occupiedWithinWork = intersectIntervals(occupied, workIntervals);
      const free = subtractIntervals(workIntervals, occupiedWithinWork);
      if (!win.has(cabinet)) win.set(cabinet, emptyForwardMetric());
      addForwardMetric(win.get(cabinet), {
        receptionMin: sumIntervalsMin(workIntervals),
        occupiedMin: sumIntervalsMin(occupiedWithinWork),
        freeMin: sumIntervalsMin(free),
        freePP: 0,
        reqCount: req.reqCount,
        primaryCount: 0,
      });
    });

    reqIndex.forEach((req, key) => {
      if (touched.has(key)) return;
      const [, cabinet] = key.split('|');
      if (!win.has(cabinet)) win.set(cabinet, emptyForwardMetric());
      addForwardMetric(win.get(cabinet), {
        receptionMin: 0,
        occupiedMin: sumIntervalsMin(mergeIntervals_(req.intervals)),
        freeMin: 0,
        freePP: 0,
        reqCount: req.reqCount,
        primaryCount: 0,
      });
    });

    win.forEach((m, cabinet) => {
      if (!series.has(cabinet)) series.set(cabinet, []);
      series.get(cabinet).push(m);
    });
  }

  const out = new Map();
  series.forEach((arr, cabinet) => {
    const avg = emptyForwardMetric();
    arr.forEach((m) => addForwardMetric(avg, m));
    const n = arr.length || 1;
    Object.keys(avg).forEach((k) => { avg[k] /= n; });
    out.set(cabinet, avg);
  });
  return out;
}

function detectCabinetType_(cabinetValue) {
  const c = str_(cabinetValue).toLowerCase();
  if (!c) return '';
  if (c.includes('фтл')) return 'ФТЛ';
  if (c.includes('магнит')) return 'Магнит';
  return '';
}

function buildCabinetScheduleIndex_(scheduleRows) {
  const map = new Map();
  (scheduleRows || []).forEach((row) => {
    const d = toDate_(row['Дата']);
    const cabinet = str_(row['Кабинет']);
    const type = detectCabinetType_(cabinet);
    const timeType = str_(row['Тип времени']);
    const start = parseTimeToMinutes_(row['Время начала']);
    const end = parseTimeToMinutes_(row['Время окончания']);
    if (!d || !type || !cabinet || timeType !== 'Приемные часы' || start === null || end === null || end <= start) return;
    const key = `${dateKey_(d)}|${type}|${cabinet}`;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push({ start, end });
  });
  map.forEach((arr, key) => map.set(key, mergeIntervals_(arr)));
  return map;
}

function buildCabinetRequestIndex_(requests, startDate, endDate) {
  const map = new Map();
  (requests || []).forEach((r) => {
    if (!isActivePlannedStatus(r.status) || !r.startDate) return;
    if (startDate && endDate && (r.startDate < startDate || r.startDate > endDate)) return;
    const cabinet = str_(r.cabinet || r['Кабинет']);
    const type = detectCabinetType_(r.cabinetType || cabinet);
    if (!type || !cabinet) return;
    const key = `${dateKey_(r.startDate)}|${type}|${cabinet}`;
    if (!map.has(key)) map.set(key, { intervals: [], reqCount: 0 });
    const e = map.get(key);
    e.intervals.push(requestInterval(r));
    e.reqCount += 1;
  });
  return map;
}

function computeCabinetTypeWindowMetrics_(anchorDate, horizonDays, scheduleRows, requests) {
  const startDate = addDays(anchorDate, 1);
  const endDate = addDays(anchorDate, horizonDays);
  const scheduleIndex = buildCabinetScheduleIndex_(scheduleRows);
  const requestIndex = buildCabinetRequestIndex_(requests, startDate, endDate);

  const out = new Map();
  const touchedReq = new Set();
  scheduleIndex.forEach((workIntervals, key) => {
    const [dKey, type] = key.split('|');
    const day = toDate_(dKey);
    if (!day || day < startDate || day > endDate) return;

    const req = requestIndex.get(key) || { intervals: [], reqCount: 0 };
    touchedReq.add(key);
    const occupied = mergeIntervals_(req.intervals);
    const occupiedWithinWork = intersectIntervals(occupied, workIntervals);
    const free = subtractIntervals(workIntervals, occupiedWithinWork);

    const metric = {
      receptionMin: sumIntervalsMin(workIntervals),
      occupiedMin: sumIntervalsMin(occupiedWithinWork),
      freeMin: sumIntervalsMin(free),
      freePP: 0,
      reqCount: req.reqCount,
      primaryCount: 0,
    };
    accumulateForward(out, type, metric, type, '');
  });

  requestIndex.forEach((req, key) => {
    if (touchedReq.has(key)) return;
    const [, type] = key.split('|');
    const occupied = mergeIntervals_(req.intervals);
    const metric = {
      receptionMin: 0,
      occupiedMin: sumIntervalsMin(occupied),
      freeMin: 0,
      freePP: 0,
      reqCount: req.reqCount,
      primaryCount: 0,
    };
    accumulateForward(out, type, metric, type, '');
  });

  return out;
}

function computeCabinetTypeBaselineMetrics_(currentDay, horizonDays, scheduleRows, requests) {
  const scheduleIndex = buildCabinetScheduleIndex_(scheduleRows);
  const baselineRequests = (requests || []).map((r) => ({
    ...r,
    status: isCompletedStatus(r.status) ? 'Запланирована' : r.status,
  }));

  const series = new Map();
  for (let i = 1; i <= 42; i += 1) {
    const ref = addDays(currentDay, -i);
    if (ref.getDay() !== currentDay.getDay()) continue;

    const startDate = addDays(ref, 1);
    const endDate = addDays(ref, horizonDays);
    const reqIndex = buildCabinetRequestIndex_(baselineRequests, startDate, endDate);

    const winByType = new Map();
    const touchedReq = new Set();
    scheduleIndex.forEach((workIntervals, key) => {
      const [dKey, type] = key.split('|');
      const day = toDate_(dKey);
      if (!day || day < startDate || day > endDate) return;

      const req = reqIndex.get(key) || { intervals: [], reqCount: 0 };
      touchedReq.add(key);
      const occupied = mergeIntervals_(req.intervals);
      const occupiedWithinWork = intersectIntervals(occupied, workIntervals);
      const free = subtractIntervals(workIntervals, occupiedWithinWork);
      if (!winByType.has(type)) winByType.set(type, emptyForwardMetric());
      addForwardMetric(winByType.get(type), {
        receptionMin: sumIntervalsMin(workIntervals),
        occupiedMin: sumIntervalsMin(occupiedWithinWork),
        freeMin: sumIntervalsMin(free),
        freePP: 0,
        reqCount: req.reqCount,
        primaryCount: 0,
      });
    });

    reqIndex.forEach((req, key) => {
      if (touchedReq.has(key)) return;
      const [, type] = key.split('|');
      if (!winByType.has(type)) winByType.set(type, emptyForwardMetric());
      addForwardMetric(winByType.get(type), {
        receptionMin: 0,
        occupiedMin: sumIntervalsMin(mergeIntervals_(req.intervals)),
        freeMin: 0,
        freePP: 0,
        reqCount: req.reqCount,
        primaryCount: 0,
      });
    });

    winByType.forEach((m, type) => {
      if (!series.has(type)) series.set(type, []);
      series.get(type).push(m);
    });
  }

  const out = new Map();
  series.forEach((arr, type) => {
    const n = arr.length || 1;
    const avg = emptyForwardMetric();
    arr.forEach((m) => addForwardMetric(avg, m));
    Object.keys(avg).forEach((k) => { avg[k] /= n; });
    out.set(type, avg);
  });
  return out;
}

function buildCabinetTypeSummary_(currentDay, scheduleRows, requests) {
  const c3 = computeCabinetTypeWindowMetrics_(currentDay, 3, scheduleRows, requests);
  const c7 = computeCabinetTypeWindowMetrics_(currentDay, 7, scheduleRows, requests);
  const b3 = computeCabinetTypeBaselineMetrics_(currentDay, 3, scheduleRows, requests);
  const b7 = computeCabinetTypeBaselineMetrics_(currentDay, 7, scheduleRows, requests);

  const getSummary = (type) => ({
    c3: c3.get(type) ? c3.get(type).metrics : emptyForwardMetric(),
    c7: c7.get(type) ? c7.get(type).metrics : emptyForwardMetric(),
    b3: b3.get(type) || emptyForwardMetric(),
    b7: b7.get(type) || emptyForwardMetric(),
  });

  return {
    ftl: getSummary('ФТЛ'),
    magnet: getSummary('Магнит'),
  };
}

function toForwardOutputRowCabinet_(horizonLabel, current, baseline) {
  const currentLoad = current.receptionMin ? (current.occupiedMin / current.receptionMin) * 100 : 0;
  const usualLoad = baseline.receptionMin ? (baseline.occupiedMin / baseline.receptionMin) * 100 : 0;
  return [
    horizonLabel,
    round1_(current.receptionMin / 60),
    `${round1_(current.occupiedMin / 60)} (${round1_(baseline.occupiedMin / 60)})`,
    `${round0(currentLoad)} (${round0(usualLoad)})`,
    round1_(current.freeMin / 60),
    round1_(current.reqCount),
  ];
}

function renderCabinetTypeSummaryTable_(sheet, row, title, summary) {
  const header = ['Горизонт', 'Рабочие (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)'];
  sheet.getRange(row, 1, 1, header.length).merge().setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  row += 1;
  sheet.getRange(row, 1, 1, header.length).setValues([header]);
  styleForecastHeaderRow(sheet, row, header.length);
  row += 1;

  const c3 = summary && summary.c3 ? summary.c3 : emptyForwardMetric();
  const c7 = summary && summary.c7 ? summary.c7 : emptyForwardMetric();
  const b3 = summary && summary.b3 ? summary.b3 : emptyForwardMetric();
  const b7 = summary && summary.b7 ? summary.b7 : emptyForwardMetric();

  const row3 = toForwardOutputRowCabinet_('3 дня', c3, b3);
  sheet.getRange(row, 1, 1, header.length).setValues([row3]);
  styleForecastDataRow(sheet, row, header.length);
  styleForecastRow(sheet, row, c3, b3, '3d', 4);
  row += 1;

  const row7 = toForwardOutputRowCabinet_('7 дней', c7, b7);
  sheet.getRange(row, 1, 1, header.length).setValues([row7]);
  styleForecastDataRow(sheet, row, header.length);
  styleForecastRow(sheet, row, c7, b7, '7d', 4);
  row += 1;

  return row;
}

function renderOverallCabinetTypeTable_(sheet, row, clinicRows, title) {
  const columns = 7;
  const header = ['Клиника', 'Горизонт', 'Рабочие (ч)', 'Занятые (ч)', 'Загрузка (%)', 'Свободные (ч)', 'Заявки (шт)'];
  sheet.getRange(row, 1, 1, columns).merge().setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  row += 1;
  sheet.getRange(row, 1, 1, columns).setValues([header]);
  styleForecastHeaderRow(sheet, row, columns);
  row += 1;

  const horizonOrder = { '3 дня': 0, '7 дней': 1 };
  const subset = (clinicRows || []).slice().sort((a, b) => {
    if (a.branch !== b.branch) return a.branch.localeCompare(b.branch, 'ru');
    return (horizonOrder[a.horizon] || 99) - (horizonOrder[b.horizon] || 99);
  });

  subset.forEach((r) => {
    const values = [
      r.branch,
      r.horizon,
      round1_(r.reception),
      `${round1_(r.occupiedCur)} (${round1_(r.occupiedUsual)})`,
      `${round0(r.loadCur)} (${round0(r.loadUsual)})`,
      round1_(r.free),
      round1_(r.requests),
    ];
    sheet.getRange(row, 1, 1, columns).setValues([values]);
    styleForecastDataRow(sheet, row, columns);
    styleForecastRow(sheet, row, { occupiedMin: r.occupiedCur * 60, receptionMin: Math.max(1, r.reception * 60) }, { occupiedMin: r.occupiedUsual * 60, receptionMin: Math.max(1, r.reception * 60) }, r.horizon === '7 дней' ? '7d' : '3d', 5);
    row += 1;
  });

  return row;
}

function getSourceGraphMaxCreatedAt_(graphSheet) {
  if (!graphSheet) return new Date(0);
  const graphRows = getSheetObjects_(graphSheet);
  let sourceMaxGraphCreated = new Date(0);
  graphRows.forEach((row) => {
    const d = toDate_(row['Дата создания']);
    if (d && d > sourceMaxGraphCreated) sourceMaxGraphCreated = d;
  });
  return sourceMaxGraphCreated;
}

function ensureDoctorScheduleFresh(ss) {
  const graphSheet = ss.getSheetByName(SHEET_SCHEDULE_SOURCE);
  const doctorScheduleSheet = ss.getSheetByName(SHEET_DOCTOR_SCHEDULE);
  if (!graphSheet || !doctorScheduleSheet) {
    throw new Error('Проверьте наличие листов: График, Обработанный график');
  }
  const builtAt = detectDoctorScheduleBuiltAtFallback(doctorScheduleSheet);
  if (!builtAt) {
    buildDoctorSchedule();
    const doctorScheduleSheet2 = ss.getSheetByName(SHEET_DOCTOR_SCHEDULE);
    const rebuiltAt = detectDoctorScheduleBuiltAtFallback(doctorScheduleSheet2);
    if (!rebuiltAt) {
      throw new Error('Сначала сформируйте лист "Обработанный график" (в последней строке нет системного timestamp)');
    }
    return ensureDoctorScheduleFresh(ss);
  }
  const sourceMaxGraphCreated = getSourceGraphMaxCreatedAt_(graphSheet);
  if (sourceMaxGraphCreated > builtAt) {
    throw new Error(`"Обработанный график" неактуален (последняя сборка: ${formatDateTimeDDMMYYYY_(builtAt)}). Обновите его перед запуском.`);
  }
}

function detectDoctorScheduleBuiltAtFallback(doctorScheduleSheet) {
  const lastRow = doctorScheduleSheet.getLastRow();
  if (lastRow < 2) return null;
  const marker = doctorScheduleSheet.getRange(lastRow, 1, 1, 10).getDisplayValues()[0];
  const markerKey = str_(marker[0]);
  const markerDate = toDate_(marker[8]);
  if (markerKey !== '__SYSTEM_TIMESTAMP__' || !markerDate) return null;
  return markerDate;
}
function sendOverallForecastEmail(ss, reportDate) {
  const sheet = ss.getSheetByName(FL2_SHEET_OVERALL_FORECAST);
  const mailingSheet = ss.getSheetByName(FL2_SHEET_MAILING);
  if (!sheet || !mailingSheet) {
    throw new Error('Не найден лист "Общий план" или "Отправка"');
  }
  const mailingRows = getSheetObjects_(mailingSheet);
  const target = mailingRows.find((r) => str_(r['Филиал']) === FL2_SHEET_OVERALL_FORECAST);
  if (!target) throw new Error('Для "Общий план" нет строки в листе Отправка');
  const to = parseEmailList(target['Кому']);
  const cc = parseEmailList(target['Копия']);
  if (!to.length) throw new Error('Для "Общий план" не заполнено поле Кому');
  const subject = `План загрузки на дату ${reportDate}`;
  const body = [
    sheetToHtmlTable(sheet, { clinicLinkMap: buildClinicSheetLinkMap_(ss) }),
    '<p><b>Цветовая интерпретация загрузки:</b> для горизонта 3 дня: <span style="background:#ff0000;padding:2px 6px;">красный</span> &lt; 25%, <span style="background:#fff2cc;padding:2px 6px;">желтый</span> 25–40%, белый 40–60%, <span style="background:#d9ead3;padding:2px 6px;">зеленый</span> 60–75%, <span style="background:#f4cccc;padding:2px 6px;">бледно-розовый</span> > 75%. Для горизонта 7 дней: <span style="background:#ff0000;padding:2px 6px;">красный</span> &lt; 20%, <span style="background:#fff2cc;padding:2px 6px;">желтый</span> 20–35%, белый 35–50%, <span style="background:#d9ead3;padding:2px 6px;">зеленый</span> 50–60%, <span style="background:#f4cccc;padding:2px 6px;">бледно-розовый</span> > 60%.</p>',
  ].join('');
  MailApp.sendEmail({
    to: to.join(','),
    cc: cc.join(','),
    subject,
    htmlBody: body,
  });
}

function buildClinicSheetLinkMap_(ss) {
  const map = {};
  if (!ss) return map;
  const spreadsheetUrl = ss.getUrl();
  ss.getSheets().forEach((sheet) => {
    const name = str_(sheet.getName());
    if (!name) return;
    if (name === FL2_SHEET_OVERALL_FORECAST) return;
    if (name === FL2_SHEET_MAILING) return;
    if (name === FL2_SHEET_BATCH_SERVICE) return;
    const a1 = str_(sheet.getRange(1, 1).getDisplayValue());
    if (!a1.startsWith('План загрузки на дату')) return;
    map[name] = `${spreadsheetUrl}#gid=${sheet.getSheetId()}`;
  });
  return map;
}

function collectMailingParticipantsByBranch_(ss, branchName) {
  const sheet = ss.getSheetByName(SHEET_DOCTOR_SCHEDULE);
  if (!sheet) {
    return { doctors: [], massage: [], ft: [] };
  }

  const doctorsSet = getDoctorsSetFromEmployeesSheet_(ss);
  const massageSet = getMassageSetFromEmployeesSheet_(ss);
  const ftSet = getFtSetFromEmployeesSheet_(ss);

  const rows = getSheetObjectsByBranch_(sheet, branchName, true);
  const doctors = new Set();
  const massage = new Set();
  const ft = new Set();

  rows.forEach((row) => {
    const employee = str_(row['Сотрудник']);
    if (!employee || employee === '__SYSTEM_TIMESTAMP__') return;
    if (doctorsSet.has(employee)) doctors.add(employee);
    if (massageSet.has(employee)) massage.add(employee);
    if (ftSet.has(employee)) ft.add(employee);
  });

  const toSorted = (set) => [...set].sort((a, b) => a.localeCompare(b, 'ru'));
  return {
    doctors: toSorted(doctors),
    massage: toSorted(massage),
    ft: toSorted(ft),
  };
}

function buildMailingServiceInfoHtml_(ss, branchName, toList, ccList, reportSheet) {
  const participants = collectMailingParticipantsByBranch_(ss, branchName);
  const reportBuiltAt = detectPlanBuiltAtFromReportSheet_(reportSheet);

  const renderList = (arr) => arr && arr.length ? arr.join(', ') : '—';
  const toText = renderList((toList || []).map((v) => str_(v)).filter((v) => v));
  const ccText = renderList((ccList || []).map((v) => str_(v)).filter((v) => v));

  return [
    '<hr style="border:none;border-top:1px solid #d9d9d9;margin:12px 0;">',
    '<p style="margin:0 0 6px 0;"><b>Служебная информация</b></p>',
    `<p style="margin:0 0 4px 0;"><b>Филиал:</b> ${str_(branchName) || '—'}</p>`,
    `<p style="margin:0 0 4px 0;"><b>Кому:</b> ${toText}</p>`,
    `<p style="margin:0 0 8px 0;"><b>Копия:</b> ${ccText}</p>`,
    `<p style="margin:0 0 4px 0;"><b>Врачи (по "Обработанный график"):</b> ${renderList(participants.doctors)}</p>`,
    `<p style="margin:0 0 4px 0;"><b>Массажисты (по "Обработанный график"):</b> ${renderList(participants.massage)}</p>`,
    `<p style="margin:0;"><b>ФТ (по "Обработанный график"):</b> ${renderList(participants.ft)}</p>`,
  ].join('');
}

function detectPlanBuiltAtFromReportSheet_(sheet) {
  if (!sheet) return null;
  const maxRows = Math.min(sheet.getLastRow(), 6);
  if (maxRows < 1) return null;
  const values = sheet.getRange(1, 1, maxRows, Math.max(1, sheet.getLastColumn())).getDisplayValues();
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i] || [];
    const first = str_(row[0]);
    if (!first.startsWith('Сформировано:')) continue;
    const m = first.match(/Сформировано:\s*(.+)$/i);
    if (!m) continue;
    const d = toDate_(m[1]);
    if (d) return d;
  }
  return null;
}

function isPlanReportSheet_(sheet) {
  if (!sheet) return false;
  const a1 = str_(sheet.getRange(1, 1).getDisplayValue());
  return a1.startsWith('План загрузки на дату') || a1.startsWith('Общий план по клиникам на');
}

function extractPlanReportDate_(sheet) {
  const a1 = str_(sheet.getRange(1, 1).getDisplayValue());
  const m = a1.match(/(\d{2}\.\d{2}\.\d{4})/);
  return m ? m[1] : '';
}

function computeSheetDisplayHash_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const serialized = JSON.stringify(values);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, serialized, Utilities.Charset.UTF_8);
  return digest.map((b) => (b + 256).toString(16).slice(-2)).join('');
}

function archivePlanReportsDaily() {
  const src = SpreadsheetApp.getActive();
  const archive = SpreadsheetApp.openById(FL2_PLAN_ARCHIVE_SPREADSHEET_ID);
  const indexSheet = getOrCreateSheet_(archive, 'Индекс плана');

  if (indexSheet.getLastRow() < 1) {
    indexSheet.getRange(1, 1, 1, 7).setValues([['Ключ', 'Лист', 'ОтчетДата', 'МеткаПлана', 'Хэш', 'Обновлено', 'ИсточникКнига']]);
  }

  const categorySheets = {
    doctor: getOrCreateSheet_(archive, 'Архив Врачи'),
    massage: getOrCreateSheet_(archive, 'Архив Массажисты'),
    ft: getOrCreateSheet_(archive, 'Архив ФТ'),
    physio: getOrCreateSheet_(archive, 'Архив Физиотерапия'),
  };

  const dataHeader = ['Ключ', 'Филиал', 'Лист', 'ОтчетДата', 'МеткаПлана', 'Горизонт', 'ТипСтроки', 'Объект', 'Приемные (ч)', 'Занятые (ч) текущее', 'Занятые (ч) обычное', 'Загрузка (%) текущее', 'Загрузка (%) обычное', 'Свободные (ч)', 'Заявки (шт)', 'Заявки на ПП (шт)', 'Свободные слоты под ПП', 'Хэш', 'Обновлено', 'ИсточникКнига'];
  Object.keys(categorySheets).forEach((k) => {
    ensureArchiveDataSheetHeader_(categorySheets[k], dataHeader);
  });

  const indexVals = indexSheet.getLastRow() > 1 ? indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 7).getDisplayValues() : [];
  const indexMap = new Map();
  const reportFreshnessMap = new Map();
  indexVals.forEach((r, i) => {
    const key = str_(r[0]);
    const ts = toDate_(r[3]);
    if (key) indexMap.set(key, { row: i + 2, ts });

    const sheetName = str_(r[1]);
    const reportDate = str_(r[2]);
    if (sheetName && reportDate && ts) {
      const reportKey = `${sheetName}|${reportDate}`;
      const prev = reportFreshnessMap.get(reportKey);
      if (!prev || ts.getTime() > prev.getTime()) {
        reportFreshnessMap.set(reportKey, ts);
      }
    }
  });

  const upsertsByCategory = {
    doctor: [],
    massage: [],
    ft: [],
    physio: [],
  };

  const planSheets = src.getSheets();
  planSheets.forEach((sheet) => {
    const values = sheet.getDataRange().getDisplayValues();
    if (!isBranchPlanSheetData_(values)) return;

    const reportDate = extractPlanReportDate_(sheet);
    const builtAt = detectPlanBuiltAtFromReportSheet_(sheet);
    if (!reportDate || !builtAt) return;

    const reportKey = `${sheet.getName()}|${reportDate}`;
    const archivedBuiltAt = reportFreshnessMap.get(reportKey);
    if (archivedBuiltAt && archivedBuiltAt.getTime() >= builtAt.getTime()) return;

    const hash = computeSheetDisplayHash_(sheet);
    const branch = extractBranchFromPlanSheetData_(values) || sheet.getName();
    const detailRows = parsePlanDetailRowsForArchive_(values);
    detailRows.forEach((r) => {
      const category = normalizeArchiveCategory_(r.category);
      if (!category || !r.horizon) return;
      const objectName = r.isTotal ? 'Итого' : str_(r.objectName);
      const key = `${sheet.getName()}|${reportDate}|${category}|${r.horizon}|${objectName}`;
      const idxCurrent = indexMap.get(key);
      if (idxCurrent && idxCurrent.ts && idxCurrent.ts.getTime() >= builtAt.getTime()) return;

      const primaryCount = category === 'doctor' ? round1_(Number(r.primaryCount) || 0) : '';
      const freePP = category === 'doctor' ? round1_(Number(r.freePP) || 0) : '';

      upsertsByCategory[category].push({
        key,
        row: [
          key,
          branch,
          sheet.getName(),
          reportDate,
          formatDateTimeDDMMYYYY_(builtAt),
          r.horizon,
          r.isTotal ? 'Итого' : 'Детализация',
          objectName,
          round1_(Number(r.reception) || 0),
          round1_(Number(r.occupiedCur) || 0),
          round1_(Number(r.occupiedUsual) || 0),
          round1_(Number(r.loadCur) || 0),
          round1_(Number(r.loadUsual) || 0),
          round1_(Number(r.free) || 0),
          round1_(Number(r.requests) || 0),
          primaryCount,
          freePP,
          hash,
          formatDateTimeDDMMYYYY_(new Date()),
          src.getId(),
        ],
        idxRowVals: [key, sheet.getName(), reportDate, formatDateTimeDDMMYYYY_(builtAt), hash, formatDateTimeDDMMYYYY_(new Date()), src.getId()],
      });
    });
  });

  const totalUpserts = Object.keys(upsertsByCategory).reduce((acc, k) => acc + upsertsByCategory[k].length, 0);
  if (!totalUpserts) {
    SpreadsheetApp.getActive().toast('Архив плана: новых изменений нет', MENU_NAME, 5);
    return;
  }

  Object.keys(upsertsByCategory).forEach((category) => {
    upsertArchiveDataRows_(categorySheets[category], upsertsByCategory[category], dataHeader.length);
  });

  Object.keys(upsertsByCategory).forEach((category) => {
    upsertsByCategory[category].forEach((u) => {
      const idxInfo = indexMap.get(u.key);
      if (idxInfo) {
        indexSheet.getRange(idxInfo.row, 1, 1, 7).setValues([u.idxRowVals]);
      } else {
        indexSheet.getRange(indexSheet.getLastRow() + 1, 1, 1, 7).setValues([u.idxRowVals]);
      }
    });
  });

  SpreadsheetApp.getActive().toast(`Архив плана обновлен: ${totalUpserts} строк`, MENU_NAME, 5);
}

function ensureArchiveDataSheetHeader_(sheet, header) {
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    return;
  }
  const width = Math.max(1, sheet.getLastColumn());
  const existing = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map((v) => str_(v));
  if (existing.length < header.length || existing[0] !== 'Ключ' || existing[1] !== 'Филиал') {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
}

function normalizeArchiveCategory_(category) {
  const c = str_(category);
  if (c === 'doctor') return 'doctor';
  if (c === 'massage') return 'massage';
  if (c === 'ft') return 'ft';
  if (c === 'physio' || c === 'cabinet') return 'physio';
  return '';
}

function extractBranchFromPlanSheetData_(data) {
  const firstCol = (data || []).map((r) => str_(r && r[0]));
  const branchLine = firstCol.find((v) => v.startsWith('Филиал:')) || '';
  return branchLine.replace(/^Филиал:\s*/i, '').trim();
}

function parsePlanDetailRowsForArchive_(data) {
  const rows = [];
  let horizon = '';
  let category = '';
  let expectedFirstCol = '';

  (data || []).forEach((row) => {
    const first = str_(row && row[0]);
    if (!first) return;

    if (first.includes('Сводный план на 3 дня') || first.includes('Блок 3 дня')) {
      horizon = '3 дня';
      category = '';
      expectedFirstCol = '';
      return;
    }
    if (first.includes('Сводный план на 7 дней') || first.includes('Блок 7 дней')) {
      horizon = '7 дней';
      category = '';
      expectedFirstCol = '';
      return;
    }

    if (first.includes('— Врачи (')) {
      category = 'doctor';
      expectedFirstCol = 'Врач';
      return;
    }
    if (first.includes('— Массажисты (')) {
      category = 'massage';
      expectedFirstCol = 'Сотрудник';
      return;
    }
    if (first.includes('— ФТ (')) {
      category = 'ft';
      expectedFirstCol = 'Сотрудник';
      return;
    }
    if (first.includes('— Кабинеты физиотерапии')) {
      category = 'physio';
      expectedFirstCol = 'Кабинет';
      return;
    }

    if (!category || !horizon) return;
    if (first === expectedFirstCol) return;
    if (first.startsWith('Таблица ') || first.startsWith('Сводный план')) return;

    const reception = parseLocaleNumber(row[1]);
    if (!reception && first !== 'ИТОГО') return;
    const occupied = parseCombinedMetric(row[2]);
    const load = parseCombinedMetric(row[3]);
    const free = parseLocaleNumber(row[4]);
    const requests = parseLocaleNumber(row[5]);
    const primaryCount = parseLocaleNumber(row[6]);
    const freePP = parseLocaleNumber(row[7]);
    rows.push({
      category,
      horizon,
      objectName: first,
      isTotal: first === 'ИТОГО',
      reception,
      occupiedCur: occupied.current,
      occupiedUsual: occupied.usual,
      loadCur: load.current,
      loadUsual: load.usual,
      free,
      requests,
      primaryCount,
      freePP,
    });
  });

  return rows;
}

function upsertArchiveDataRows_(sheet, upserts, width) {
  if (!upserts || !upserts.length) return;
  const lastRow = sheet.getLastRow();
  const vals = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues() : [];
  const rowMap = new Map();
  vals.forEach((r, i) => {
    const key = str_(r[0]);
    if (key) rowMap.set(key, i + 2);
  });

  upserts.forEach((u) => {
    const rowNum = rowMap.get(u.key);
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, width).setValues([u.row]);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, width).setValues([u.row]);
    }
  });
}

function parseEmailList(value) {
  return str_(value)
    .split(',')
    .map((v) => v.trim())
    .filter((v) => v);
}
function sheetToHtmlTable(sheet, options) {
  const opts = options || {};
  const clinicLinkMap = opts.clinicLinkMap || null;
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();
  const fontWeights = range.getFontWeights();
  let currentHorizon = '3d';
  let loadColumnIdx = 4;
  const htmlRows = values
    .map((row, rowIndex) => ({ row, rowIndex }))
    .filter((item) => item.row.some((cell) => str_(cell) !== ''))
    .map((item) => {
    const row = item.row;
    const sourceRowIndex = item.rowIndex;
    const first = str_(row[0]);
    if (first.includes('Таблица 2')) currentHorizon = '3d';
    if (first.includes('Таблица 3')) currentHorizon = '7d';
    if (first === '3 дня') currentHorizon = '3d';
    if (first === '7 дней') currentHorizon = '7d';

    const headerLoadIdx = row.findIndex((cell) => /^Загрузка\s*\(%\)/i.test(str_(cell)) || /^Загрузка/i.test(str_(cell)));
    if (headerLoadIdx >= 0) {
      loadColumnIdx = headerLoadIdx;
    }

    const cells = row.map((cell, idx) => {
      const cellText = str_(cell);
      const sourceBg = backgrounds[sourceRowIndex][idx] || '#ffffff';
      const sourceWeight = fontWeights[sourceRowIndex][idx] || 'normal';
      let background = sourceBg;
      if (idx === loadColumnIdx && isLoadMetricCellForEmail_(first, cellText)) {
        const parsed = parseCombinedMetric(cellText);
        background = forecastLoadColorByHorizon(parsed.current, currentHorizon);
      }
      const tag = sourceWeight === 'bold' ? 'th' : 'td';
      const align = idx === 0 ? 'left' : 'center';
      const style = [
        'border:1px solid #cfd8dc',
        'padding:6px 8px',
        `background:${background}`,
        `font-weight:${sourceWeight}`,
        `text-align:${align}`,
        'vertical-align:middle',
      ].join(';');
      let textHtml = cellText || '';
      if (clinicLinkMap && idx === 0 && shouldLinkClinicCellForEmail_(first, cellText, clinicLinkMap)) {
        const href = clinicLinkMap[cellText];
        textHtml = `<a href="${href}" target="_blank" style="color:inherit;text-decoration:none;">${cellText}</a>`;
      }
      return `<${tag} style="${style}">${textHtml}</${tag}>`;
    }).join('');
    return `<tr>${cells}</tr>`;
  }).join('');
  return `<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;line-height:1.35;">${htmlRows}</table>`;
}

function shouldLinkClinicCellForEmail_(firstCellText, cellText, clinicLinkMap) {
  const first = str_(firstCellText);
  const cell = str_(cellText);
  if (!cell || !clinicLinkMap || !clinicLinkMap[cell]) return false;
  if (first === 'Клиника' || first === 'ИТОГО') return false;
  if (first.startsWith('Таблица') || first.startsWith('Блок') || first.startsWith('Общий план') || first.startsWith('Сформировано:')) return false;
  return true;
}

function isLoadMetricCellForEmail_(firstCellText, cellText) {
  const first = str_(firstCellText);
  const value = str_(cellText);
  if (!value) return false;
  if (/^Загрузка/i.test(value)) return false;
  const hasPair = /^-?[\d.,]+\s*\(-?[\d.,]+\)$/.test(value);
  const hasSingleNumber = /^-?[\d.,]+$/.test(value);
  if (!hasPair && !hasSingleNumber) return false;
  if (first === '3 дня' || first === '7 дней' || first === 'ИТОГО') return true;
  if (first.startsWith('Темед')) return true;
  if (first.startsWith('Таблица') || first.startsWith('План') || first.startsWith('Филиал:')) return false;
  return true;
}
function isCompletedStatus(status) {
  return status === 'Проведен прием'
    || status === 'Выполнена'
    || status === 'Пришел';
}

function isPlannedStatus(status) {
  return status === 'Запланирована'
    || status === 'Подтвердил запись'
    || status === 'Недозвон. Отправить смс';
}

function isAllowedProcessedRequestStatus(status) {
  return isCompletedStatus(status) || isPlannedStatus(status);
}

function isActivePlannedStatus(status) {
  return isPlannedStatus(status) || isCompletedStatus(status);
}
function toIntOrNull_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const num = Number(value);
  if (!Number.isFinite(num)) return null;
  return Math.trunc(num);
}

function computeRequestTiming_(startDate, endDate) {
  if (!startDate || !endDate) {
    return { startMin: null, endMin: null, durationMin: null };
  }
  const startMin = startDate.getHours() * 60 + startDate.getMinutes();
  let endMin = endDate.getHours() * 60 + endDate.getMinutes();
  let durationMin = endMin - startMin;
  if (durationMin <= 0) {
    durationMin += 24 * 60;
    endMin = startMin + durationMin;
  }
  return { startMin, endMin, durationMin };
}

function requestInterval(req) {
  if (req && Number.isFinite(req.startMin) && Number.isFinite(req.durationMin) && req.durationMin > 0) {
    return { start: req.startMin, end: req.startMin + req.durationMin };
  }
  if (req && Number.isFinite(req.startMin) && Number.isFinite(req.endMin) && req.endMin > req.startMin) {
    return { start: req.startMin, end: req.endMin };
  }
  if (!req || !req.startDate || !req.endDate) {
    return { start: 0, end: 1 };
  }
  const timing = computeRequestTiming_(req.startDate, req.endDate);
  if (!Number.isFinite(timing.startMin) || !Number.isFinite(timing.durationMin) || timing.durationMin <= 0) {
    return { start: 0, end: 1 };
  }
  return { start: timing.startMin, end: timing.startMin + timing.durationMin };
}
function sumIntervalsMin(intervals) {
  return intervals.reduce((acc, i) => acc + (i.end - i.start), 0);
}
function intersectIntervals(a, b) {
  const out = [];
  a.forEach((x) => {
    b.forEach((y) => {
      const start = Math.max(x.start, y.start);
      const end = Math.min(x.end, y.end);
      if (end > start) out.push({ start, end });
    });
  });
  return mergeIntervals_(out);
}
function subtractIntervals(base, sub) {
  let result = [...base];
  sub.forEach((s) => {
    const next = [];
    result.forEach((b) => {
      if (s.end <= b.start || s.start >= b.end) {
        next.push(b);
        return;
      }
      if (s.start > b.start) next.push({ start: b.start, end: s.start });
      if (s.end < b.end) next.push({ start: s.end, end: b.end });
    });
    result = next;
  });
  return mergeIntervals_(result);
}
function addDays(date, days) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
}
function isDateInRangeInclusive_(date, start, end) {
  if (!date || !start || !end) return false;
  return date >= start && date <= end;
}
function emptyForwardMetric() {
  return { receptionMin: 0, occupiedMin: 0, freeMin: 0, freePP: 0, reqCount: 0, primaryCount: 0 };
}
function addForwardMetric(target, src) {
  Object.keys(target).forEach((k) => {
    target[k] += src[k] || 0;
  });
}
function round0(value) {
  return Math.round(value);
}
function emptyForwardEntry(key) {
  return { branch: key, doctor: '', metrics: emptyForwardMetric() };
}
function accumulateForward(map, key, metric, branch, doctor) {
  if (!map.has(key)) map.set(key, { branch: branch || key, doctor: doctor || '', metrics: emptyForwardMetric() });
  addForwardMetric(map.get(key).metrics, metric);
}
