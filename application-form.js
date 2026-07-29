let database = window.scheduleDatabase;
const calendarMonths = [
  { year: 2026, month: 8 },
];
const weekdayTimes = ["14:00", "15:00", "16:00", "17:00", "18:00", "19:00"];
const saturdayTimes = ["11:00", "12:00", "13:00", "14:00", "15:00"];
const closedDateKeys = [];

const state = {
  campusId: "",
  subjectId: "",
  date: "",
  time: "",
  monthIndex: getInitialMonthIndex(),
};

function getInitialMonthIndex() {
  const today = new Date();
  const currentKey = today.getFullYear() * 100 + today.getMonth() + 1;
  const exactIndex = calendarMonths.findIndex(({ year, month }) => year * 100 + month === currentKey);

  if (exactIndex >= 0) return exactIndex;
  if (currentKey < calendarMonths[0].year * 100 + calendarMonths[0].month) return 0;
  return calendarMonths.length - 1;
}

const campusSelect = document.querySelector("#campusSelect");
const subjectSelect = document.querySelector("#subjectSelect");
const calendarDays = document.querySelector("#calendarDays");
const timeList = document.querySelector("#timeList");
const selectedDateText = document.querySelector("#selectedDateText");
const selectedTimeText = document.querySelector("#selectedTimeText");
const monthTitle = document.querySelector("#monthTitle");
const prevMonthButton = document.querySelector("#prevMonthButton");
const nextMonthButton = document.querySelector("#nextMonthButton");
const applicationForm = document.querySelector("#applicationForm");
const gradeSelect = applicationForm.querySelector('select[name="grade"]');
const submitStatus = document.querySelector("#submitStatus");
const phoneInput = applicationForm.querySelector('input[name="phone"]');
const privacyAgree = document.querySelector("#privacyAgree");
const submitButton = document.querySelector("#submitButton");
const completeModal = document.querySelector("#completeModal");
const completeCloseButton = document.querySelector("#completeCloseButton");
const completeCaption = completeModal?.querySelector(".complete-caption");
const submitButtonDefaultHtml = submitButton.innerHTML;
let isSubmitting = false;
let completionRedirectUrl = "";
let completionSelection = null;
let scheduleRefreshPromise = null;
let isScheduleReady = false;

function getTrackingParams() {
  const params = new URLSearchParams(window.location.search);
  const utmSource = params.get("utm_source") || "";
  const utmMedium = params.get("utm_medium") || "";
  const utmCampaign = params.get("utm_campaign") || "";
  const utmContent = params.get("utm_content") || "";
  const utmTerm = params.get("utm_term") || "";

  return {
    utmSource,
    utmMedium,
    utmCampaign,
    utmContent,
    utmTerm,
    landingUrl: window.location.href,
    referrer: document.referrer || "",
  };
}

function buildReferralFromTracking(tracking) {
  return {
    referralSource: "랜딩",
    referralDetail: tracking.utmSource || "기타",
  };
}

function getSelectedCampus() {
  return database.campuses.find((campus) => campus.id === state.campusId);
}

function getSelectedSubject() {
  return database.subjects.find((subject) => subject.id === state.subjectId);
}

function getSubjectsByCampus(campusId) {
  return database.subjects.filter((subject) => {
    const isCombinedSubject = subject.id === "math_english" || subject.name === "수학+영어";
    return subject.campusIds.includes(campusId) && !isCombinedSubject;
  });
}

function scheduleValues(value) {
  if (Array.isArray(value)) return value;
  return String(value || "")
    .split(/[,，]/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function normalizeScheduleValue(value) {
  return String(value || "")
    .replace(/캠퍼스/g, "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

function normalizeScheduleDate(value) {
  const text = String(value || "").trim();
  const match = text.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!match) return text;
  return `${match[1]}-${String(match[2]).padStart(2, "0")}-${String(match[3]).padStart(2, "0")}`;
}

function isClosedScheduleStatus(value) {
  const status = normalizeScheduleValue(value);
  return ["휴무", "미운영", "신청불가", "closed", "n", "false", "0"].includes(status);
}

function isAllScheduleValue(value) {
  const normalized = normalizeScheduleValue(value);
  return !normalized || normalized === "전체" || normalized === "all";
}

function scheduleDimensionMatches(values, candidates) {
  const items = scheduleValues(values);
  if (items.length === 0 || items.some(isAllScheduleValue)) return true;
  if (!candidates.length || candidates.every((candidate) => !candidate)) return true;

  const normalizedCandidates = candidates.map(normalizeScheduleValue);
  return items.some((item) => normalizedCandidates.includes(normalizeScheduleValue(item)));
}

function getSelectedGradeGroup() {
  const grade = gradeSelect?.value || "";
  if (grade.startsWith("초")) return "초등";
  if (grade.startsWith("중")) return "중등";
  if (grade.startsWith("고")) return "고등";
  return "";
}

function getScheduleCandidates() {
  const campus = getSelectedCampus();
  const subject = getSelectedSubject();
  return {
    campuses: [state.campusId, campus?.name, campus?.olympiadName],
    subjects: [state.subjectId, subject?.name, subject?.olympiadName],
    gradeGroups: [getSelectedGradeGroup()],
  };
}

function ruleMatchesSelection(rule) {
  const candidates = getScheduleCandidates();
  return (
    scheduleDimensionMatches(rule.campuses, candidates.campuses) &&
    scheduleDimensionMatches(rule.subjects, candidates.subjects) &&
    scheduleDimensionMatches(rule.gradeGroups, candidates.gradeGroups)
  );
}

function uniqueSortedTimes(rules) {
  return [...new Set(rules.flatMap((rule) => scheduleValues(rule.times))
    .map((time) => {
      const match = String(time || "").match(/(?:^|\s)(\d{1,2}:\d{2})(?::\d{2})?/);
      return match ? match[1] : "";
    }))]
    .filter(Boolean)
    .sort();
}

function getSpecialTimesForDate(date) {
  const rules = database.scheduleRules?.special || [];
  const normalizedDate = normalizeScheduleDate(date);
  const dateRules = rules.filter((rule) => normalizeScheduleDate(rule.date) === normalizedDate);
  if (dateRules.length === 0) return null;

  const globalClosedRule = dateRules.find(
    (rule) =>
      isClosedScheduleStatus(rule.status) &&
      scheduleValues(rule.campuses).some(isAllScheduleValue) &&
      scheduleValues(rule.subjects).some(isAllScheduleValue) &&
      scheduleValues(rule.gradeGroups).some(isAllScheduleValue)
  );
  if (globalClosedRule) return [];

  const matchingRules = dateRules.filter(ruleMatchesSelection);
  if (matchingRules.length === 0) return [];
  if (matchingRules.some((rule) => isClosedScheduleStatus(rule.status))) return [];
  return uniqueSortedTimes(matchingRules);
}

function weekdayRuleMatches(rule, dayOfWeek) {
  const weekdays = scheduleValues(rule.weekdays);
  const dayNames = ["일", "월", "화", "수", "목", "금", "토"];
  return weekdays.some((weekday) => {
    const normalized = normalizeScheduleValue(weekday);
    if (isAllScheduleValue(weekday)) return true;
    if (["월~금", "월-금", "평일"].includes(normalized)) return dayOfWeek >= 1 && dayOfWeek <= 5;
    return normalized === dayNames[dayOfWeek];
  });
}

function scheduleRuleSpecificity(rule) {
  let score = 0;
  const campusValues = scheduleValues(rule.campuses);
  const subjectValues = scheduleValues(rule.subjects);
  const gradeGroupValues = scheduleValues(rule.gradeGroups);
  if (campusValues.length && !campusValues.some(isAllScheduleValue)) score += 8;
  if (subjectValues.length && !subjectValues.some(isAllScheduleValue)) score += 4;
  if (gradeGroupValues.length && !gradeGroupValues.some(isAllScheduleValue)) score += 2;
  return score;
}

function getBasicTimesForDate(date) {
  const dayOfWeek = parseDateKey(date).getDay();
  const rules = (database.scheduleRules?.basic || []).filter(
    (rule) => weekdayRuleMatches(rule, dayOfWeek) && ruleMatchesSelection(rule)
  );

  if (rules.length === 0) {
    if (dayOfWeek === 0) return [];
    if (isSaturdayBlocked(date)) return [];
    return dayOfWeek === 6 ? saturdayTimes : weekdayTimes;
  }

  const highestSpecificity = Math.max(...rules.map(scheduleRuleSpecificity));
  const selectedRules = rules.filter((rule) => scheduleRuleSpecificity(rule) === highestSpecificity);
  if (selectedRules.some((rule) => isClosedScheduleStatus(rule.status))) return [];
  return uniqueSortedTimes(selectedRules);
}

function updateSubmitButtonState() {
  submitButton.disabled = isSubmitting || !privacyAgree.checked;
}

function setSubmitting(value) {
  isSubmitting = value;
  submitButton.innerHTML = value ? "신청 중..." : submitButtonDefaultHtml;
  updateSubmitButtonState();
}

function trackApplicationComplete(payload) {
  if (typeof window.gtag !== "function") {
    return;
  }

  window.gtag("event", "generate_lead", {
    send_to: "G-GQLE2L3HRB",
    form_name: "2026 가을학기 학력진단평가 신청",
    campus_id: payload.campusId,
    campus_name: payload.campusName,
    subject_id: payload.subjectId,
    subject_name: payload.subjectName,
  });
}

function loadJsonp(url, params = {}) {
  return new Promise((resolve, reject) => {
    const callbackName = `campusInfoCallback_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
    const script = document.createElement("script");
    const separator = url.includes("?") ? "&" : "?";
    const query = new URLSearchParams({
      ...params,
      callback: callbackName,
      _: Date.now().toString(),
    });
    const timeoutId = window.setTimeout(() => {
      delete window[callbackName];
      script.remove();
      reject(new Error("캠퍼스정보 응답 시간이 초과되었습니다."));
    }, 5000);

    window[callbackName] = (data) => {
      window.clearTimeout(timeoutId);
      delete window[callbackName];
      script.remove();
      resolve(data);
    };

    script.onerror = () => {
      window.clearTimeout(timeoutId);
      delete window[callbackName];
      script.remove();
      reject(new Error("캠퍼스정보를 불러오지 못했습니다."));
    };

    script.src = `${url}${separator}${query.toString()}`;
    document.head.append(script);
  });
}

async function loadRemoteCampusOptions() {
  const gasWebAppUrl = window.appConfig?.gasWebAppUrl;
  if (!gasWebAppUrl) return;

  const allowedCampusIds = database.campuses.map((campus) => campus.id);
  const result = await loadJsonp(gasWebAppUrl, { action: "campus-options" });
  if (!result?.ok || !result.data?.campuses?.length || !result.data?.subjects?.length) {
    return;
  }

  const campuses = result.data.campuses.filter((campus) => allowedCampusIds.includes(campus.id));
  const subjects = result.data.subjects
    .map((subject) => ({
      ...subject,
      campusIds: subject.campusIds.filter((id) => allowedCampusIds.includes(id)),
    }))
    .filter((subject) => subject.campusIds.length > 0);

  if (campuses.length && subjects.length) {
    const remoteScheduleRules = result.data.scheduleRules;
    const hasRemoteScheduleRules =
      remoteScheduleRules?.basic?.length || remoteScheduleRules?.special?.length;
    database = {
      ...database,
      campuses,
      subjects,
      scheduleRules: hasRemoteScheduleRules ? remoteScheduleRules : database.scheduleRules,
    };
  }
}

async function refreshRemoteScheduleRules() {
  const gasWebAppUrl = window.appConfig?.gasWebAppUrl;
  if (!gasWebAppUrl) return;
  if (scheduleRefreshPromise) return scheduleRefreshPromise;

  scheduleRefreshPromise = loadJsonp(gasWebAppUrl, { action: "campus-options" })
    .then((result) => {
      const remoteScheduleRules = result?.data?.scheduleRules;
      if (!result?.ok || (!remoteScheduleRules?.basic?.length && !remoteScheduleRules?.special?.length)) {
        return;
      }

      database = { ...database, scheduleRules: remoteScheduleRules };
      if (state.date && getTimesForDate(state.date).length === 0) {
        state.date = "";
        state.time = "";
      } else if (state.time && !getTimesForDate(state.date).includes(state.time)) {
        state.time = "";
      }
      renderCalendar();
      renderTimes();
      updateSummary();
    })
    .catch(() => {})
    .finally(() => {
      scheduleRefreshPromise = null;
    });

  return scheduleRefreshPromise;
}

function initCampusOptions() {
  campusSelect.innerHTML = '<option value="">캠퍼스 선택</option>';
  database.campuses.forEach((campus) => {
    campusSelect.append(new Option(campus.name, campus.id));
  });
}

function updateSubjectOptions() {
  subjectSelect.innerHTML = "";

  if (!state.campusId) {
    subjectSelect.disabled = true;
    subjectSelect.append(new Option("캠퍼스를 먼저 선택", ""));
    return;
  }

  const subjects = getSubjectsByCampus(state.campusId);

  if (subjects.length === 1) {
    subjectSelect.append(new Option(subjects[0].name, subjects[0].id));
    subjectSelect.disabled = true;
    state.subjectId = subjects[0].id;
    return;
  }

  subjectSelect.disabled = false;
  subjectSelect.append(new Option("전형과목 선택", ""));
  subjects.forEach((subject) => {
    subjectSelect.append(new Option(subject.name, subject.id));
  });
}


function getCurrentMonth() {
  return calendarMonths[state.monthIndex];
}

function dateKey(year, month, day) {
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function parseDateKey(date) {
  const [year, month, day] = date.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function getTodayKey() {
  const today = new Date();
  return dateKey(today.getFullYear(), today.getMonth() + 1, today.getDate());
}

function isClosedDate(date) {
  return date <= getTodayKey() || closedDateKeys.includes(date);
}

function isSaturdayBlocked(date) {
  const subject = getSelectedSubject();
  const blockedCampusIds = subject?.unavailableSaturdayCampusIds || [];
  return parseDateKey(date).getDay() === 6 && blockedCampusIds.includes(state.campusId);
}

function getTimesForDate(date) {
  if (!isScheduleReady) {
    return [];
  }

  if (isClosedDate(date)) {
    return [];
  }

  if (!state.campusId || !state.subjectId) {
    return [];
  }

  const specialTimes = getSpecialTimesForDate(date);
  if (specialTimes) {
    return specialTimes;
  }

  return getBasicTimesForDate(date);
}

function updateMonthControls() {
  const { year, month } = getCurrentMonth();
  monthTitle.textContent = `${year}.${month}`;
  prevMonthButton.disabled = state.monthIndex === 0;
  nextMonthButton.disabled = state.monthIndex === calendarMonths.length - 1;
}

function renderCalendar() {
  const { year, month } = getCurrentMonth();
  const firstWeekday = new Date(year, month - 1, 1).getDay();
  const lastDay = new Date(year, month, 0).getDate();

  updateMonthControls();
  calendarDays.innerHTML = "";

  for (let blank = 0; blank < firstWeekday; blank += 1) {
    calendarDays.append(document.createElement("span"));
  }

  for (let day = 1; day <= lastDay; day += 1) {
    const key = dateKey(year, month, day);
    const isClosed = isClosedDate(key);
    const times = getTimesForDate(key);
    const button = document.createElement("button");
    button.type = "button";
    button.className = "cal-day mx-auto flex flex-col items-center justify-center rounded-lg text-sm text-slate-800";
    button.textContent = day;

    if (isClosed) {
      button.disabled = true;
      button.className = "cal-day mx-auto flex flex-col items-center justify-center rounded-lg text-sm text-slate-300";
      button.setAttribute("aria-label", `${key} 선택 불가`);
    } else if (times.length > 0) {
      button.className = "cal-day mx-auto flex flex-col items-center justify-center rounded-lg bg-orange-50 text-sm font-black text-orange-700";
      button.innerHTML = `${day}<small class="text-[9px] leading-none">응시</small>`;
      button.addEventListener("click", () => selectDate(key));
    } else {
      button.disabled = true;
      button.className = "cal-day mx-auto flex flex-col items-center justify-center rounded-lg text-sm text-slate-300";
      button.setAttribute("aria-label", `${key} 선택 불가`);
    }

    if (!isClosed && state.date === key) {
      button.className = "cal-day mx-auto flex flex-col items-center justify-center rounded-lg bg-orange-600 text-sm font-black text-white";
      button.innerHTML = `${day}<small class="text-[9px] leading-none">응시</small>`;
    }

    calendarDays.append(button);
  }
}

function selectDate(date) {
  if (isClosedDate(date)) {
    return;
  }

  state.date = date;
  state.time = "";
  renderCalendar();
  renderTimes();
  updateSummary();
}

function renderTimes() {
  timeList.innerHTML = "";

  if (!state.campusId || !state.subjectId || !state.date) {
    const empty = document.createElement("p");
    empty.className = "text-sm text-slate-500";
    empty.textContent = "캠퍼스명, 전형과목, 희망날짜를 선택하면 희망시가 표시됩니다.";
    timeList.append(empty);
    return;
  }

  const times = getTimesForDate(state.date);

  if (times.length === 0) {
    const empty = document.createElement("p");
    empty.className = "text-sm text-slate-500";
    empty.textContent = "선택 가능한 희망시가 없습니다.";
    timeList.append(empty);
    return;
  }

  times.forEach((time) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "h-11 rounded-md border border-slate-200 bg-white px-5 text-sm font-black text-slate-800";
    button.textContent = time;

    if (state.time === time) {
      button.className = "h-11 rounded-md border border-orange-600 bg-white px-5 text-sm font-black text-orange-700";
    }

    button.addEventListener("click", () => {
      state.time = time;
      renderTimes();
      updateSummary();
    });

    timeList.append(button);
  });
}

function updateSummary() {
  selectedDateText.textContent = state.date || "날짜 선택 전";
  selectedTimeText.textContent = state.time || "시간 선택 전";
}

function getCompletionRedirectUrl() {
  const campus = getSelectedCampus();
  const redirectSite = campus?.redirectSites?.[state.subjectId] || "";
  const inputSite = campus?.inputSites?.[state.subjectId] || "";
  return normalizeRedirectUrl(redirectSite) || getSiteOriginUrl(inputSite);
}

function normalizeRedirectUrl(value) {
  if (!value) return "";
  try {
    const url = new URL(value, window.location.href);
    return url.protocol === "http:" || url.protocol === "https:" ? url.href : "";
  } catch (error) {
    return "";
  }
}

function getSiteOriginUrl(value) {
  const normalizedUrl = normalizeRedirectUrl(value);
  if (!normalizedUrl) return "";

  try {
    return new URL(normalizedUrl).origin;
  } catch (error) {
    return "";
  }
}

async function resolveCompletionRedirectUrl() {
  const cachedRedirectUrl = completionRedirectUrl;
  const selectedApplication = completionSelection;
  if (cachedRedirectUrl) return cachedRedirectUrl;
  if (!selectedApplication) return "";

  const gasWebAppUrl = window.appConfig?.gasWebAppUrl;
  if (!gasWebAppUrl) return "";

  try {
    const result = await loadJsonp(gasWebAppUrl, {
      action: "redirect-site",
      campusId: selectedApplication.campusId,
      subjectId: selectedApplication.subjectId,
    });
    return normalizeRedirectUrl(result?.url || "");
  } catch (error) {
    return "";
  }
}

function showCompletionModal(redirectUrl = "", selection = null) {
  if (!completeModal) return;
  completionRedirectUrl = redirectUrl;
  completionSelection = selection;
  if (completeCaption) {
    completeCaption.textContent = "※ 캠퍼스에서 확인을 위해 연락을 드릴 수 있습니다.";
  }
  completeModal.hidden = false;
  document.body.style.overflow = "hidden";
  completeCloseButton?.focus();
}

function closeCompletionModal() {
  if (!completeModal) return;
  completeModal.hidden = true;
  document.body.style.overflow = "";
}

async function confirmCompletionModal() {
  const redirectPromise = resolveCompletionRedirectUrl();
  completionRedirectUrl = "";
  completionSelection = null;
  closeCompletionModal();

  const redirectUrl = await redirectPromise;
  if (redirectUrl) {
    window.location.assign(redirectUrl);
  }
}


function buildPayload(form) {
  const formData = new FormData(form);
  const campus = getSelectedCampus();
  const subject = getSelectedSubject();
  const tracking = getTrackingParams();
  const referral = buildReferralFromTracking(tracking);

  return {
    studentName: formData.get("studentName"),
    region: formData.get("region"),
    campusId: state.campusId,
    campusName: campus?.olympiadName || campus?.name || "",
    campusLocation: campus?.location || "",
    school: formData.get("school"),
    grade: formData.get("grade"),
    phone: formData.get("phone"),
    preferredDate: state.date,
    preferredTime: state.time,
    subjectId: state.subjectId,
    subjectName: subject?.olympiadName || subject?.name || "",
    referralSource: referral.referralSource,
    referralDetail: referral.referralDetail,
    ...tracking,
  };
}

function formatPhoneNumber(value) {
  const digits = value.replace(/\D/g, "").slice(0, 11);

  if (digits.length <= 3) {
    return digits;
  }

  if (digits.length <= 7) {
    return `${digits.slice(0, 3)}-${digits.slice(3)}`;
  }

  return `${digits.slice(0, 3)}-${digits.slice(3, 7)}-${digits.slice(7)}`;
}

function submitApplication(payload) {
  const gasWebAppUrl = window.appConfig?.gasWebAppUrl;

  if (!gasWebAppUrl) {
    throw new Error("config.js에 GAS 웹앱 URL을 입력해 주세요.");
  }

  fetch(gasWebAppUrl, {
    method: "POST",
    mode: "no-cors",
    keepalive: true,
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  }).catch(() => {});

  return "신청 정보가 전송되었습니다.";
}

campusSelect.addEventListener("change", (event) => {
  state.campusId = event.target.value;
  state.subjectId = "";
  state.date = "";
  state.time = "";
  updateSubjectOptions();
  renderCalendar();
  renderTimes();
  updateSummary();
});

subjectSelect.addEventListener("change", (event) => {
  state.subjectId = event.target.value;
  state.date = "";
  state.time = "";
  renderCalendar();
  renderTimes();
  updateSummary();
});

gradeSelect?.addEventListener("change", () => {
  state.date = "";
  state.time = "";
  renderCalendar();
  renderTimes();
  updateSummary();
});

prevMonthButton.addEventListener("click", () => {
  state.monthIndex = Math.max(0, state.monthIndex - 1);
  renderCalendar();
});

nextMonthButton.addEventListener("click", () => {
  state.monthIndex = Math.min(calendarMonths.length - 1, state.monthIndex + 1);
  renderCalendar();
});

phoneInput.addEventListener("input", (event) => {
  event.target.value = formatPhoneNumber(event.target.value);
});

privacyAgree.addEventListener("change", updateSubmitButtonState);

window.addEventListener("focus", refreshRemoteScheduleRules);
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") {
    refreshRemoteScheduleRules();
  }
});

applicationForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (isSubmitting) {
    return;
  }

  if (!privacyAgree.checked) {
    submitStatus.textContent = "개인정보 수집 및 이용에 동의해 주세요.";
    return;
  }

  if (!state.date || !state.time) {
    submitStatus.textContent = "희망날짜와 희망시를 선택해 주세요.";
    return;
  }

  submitStatus.textContent = "신청 정보를 전송하고 있습니다.";
  setSubmitting(true);

  try {
    const payload = buildPayload(applicationForm);
    submitStatus.textContent = submitApplication(payload);
    trackApplicationComplete(payload);
    showCompletionModal(getCompletionRedirectUrl(), {
      campusId: state.campusId,
      subjectId: state.subjectId,
    });
    applicationForm.reset();
    state.campusId = "";
    state.subjectId = "";
    state.date = "";
    state.time = "";
    updateSubjectOptions();
    updateSubmitButtonState();
    renderCalendar();
    renderTimes();
    updateSummary();
  } catch (error) {
    submitStatus.textContent = error.message;
  } finally {
    setSubmitting(false);
  }
});

completeCloseButton?.addEventListener("click", confirmCompletionModal);
completeModal?.addEventListener("click", (event) => {
  if (event.target === completeModal) {
    completionRedirectUrl = "";
    completionSelection = null;
    closeCompletionModal();
  }
});


async function boot() {
  initCampusOptions();
  updateSubjectOptions();
  updateSubmitButtonState();
  renderCalendar();
  renderTimes();
  updateSummary();

  try {
    await loadRemoteCampusOptions();
  } catch (error) {
    submitStatus.textContent = "";
  } finally {
    isScheduleReady = true;
  }

  initCampusOptions();
  updateSubjectOptions();
  renderCalendar();
  renderTimes();
  updateSummary();
}

boot();
