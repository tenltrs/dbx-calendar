<script>
  import { onMount } from "svelte";
  // Imported statically on purpose. As a dynamic import Vite emits a relative
  // specifier, and the DBX host inlines the entry script, so it resolves against
  // the document root instead of /assets/ and the chunk 404s. Bundling it costs
  // size up front but cannot fail at runtime.
  //
  // The full build is used rather than the mini one because only it can write
  // and read the legacy .xls format; mini strips the CFB codec entirely.
  import XLSX from "xlsx/dist/xlsx.full.min.js";
  // Both icons are under Vite's inline threshold, so they become data URIs and
  // never turn into a runtime fetch.
  import xlsIcon from "../assets/xls.svg";
  import txtIcon from "../assets/txt.svg";
  import previousIcon from "../assets/previous.svg";
  import nextIcon from "../assets/next.svg";
  import productIcon from "../assets/icon.svg";
  // One icon family (single-colour, 1024 grid) for the menu and the settings
  // rail, so the two lists read as the same hand.
  import settingsIcon from "../assets/settings.svg";
  import exportIcon from "../assets/export.svg";
  import importIcon from "../assets/import.svg";
  import calendarIcon from "../assets/calendar.svg";
  import infoIcon from "../assets/info.svg";
  // The About pane names a version, and the manifest is the only place it is
  // actually declared; reading it here keeps the two from drifting apart.
  import manifest from "../manifest.json";
  import { dayInfo } from "./lunar.js";
  import { highlightMatches, navigateSearchResults, searchNotes } from "./search.js";
  import { invoke, whenReady, hostLocale, bridgeAvailable, readSettingsStorage, writeSettingsStorage } from "./dbx.js";

  /** Vite inlines small SVGs as data URIs whose payload contains single quotes,
   *  and an unquoted CSS url() rejects quote characters outright. Quoting here
   *  keeps the value legal wherever it ends up. */
  const maskVar = (url) => `--icon: url("${url}")`;

  const copy = {
    en: {
      prev: "Previous month",
      next: "Next month",
      today: "Today",
      chooseMonth: "Choose month",
      previousYear: "Previous year",
      nextYear: "Next year",
      yearLabel: "Year",
      weekdays: ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
      editorTitle: "Edit entry",
      placeholder: "Write something for this day…",
      save: "Save",
      cancel: "Cancel",
      clear: "Clear",
      hint: "Ctrl + Enter saves, Esc closes",
      hasNote: "has an entry",
      noNote: "no entry",
      dayOff: "Off",
      workday: "Work",
      edit: "double-click to edit",
      search: "Search",
      searchTitle: "Search notes",
      searchPlaceholder: "Search note content…",
      searchPrompt: "Enter a keyword to search notes",
      searchEmpty: "No matching notes",
      searchCount: (count) => `${count} ${count === 1 ? "match" : "matches"}`,
      selectedDateLabel: "selected date",
      closeSearch: "Close search",
      searchHelpNavigate: "navigate",
      searchHelpClose: "close",
      menu: "Menu",
      settings: "Settings",
      settingsClose: "Close settings",
      tabGrid: "Calendar",
      tabAbout: "About",
      groupStyle: "Calendar style",
      groupCellStyle: "Cell style",
      groupDisplay: "Display",
      groupFestival: "Festivals",
      groupLunar: "Lunar calendar",
      styleFull: "Full month",
      styleFullNote: "Always starts on the 1st.",
      styleFloat: "Floating month",
      styleFloatNote: "Keeps the anchor day on one row.",
      pastWeeks: "Weeks above",
      pastWeeksHint: "Week rows kept above the anchor day; 0 puts it in the first row.",
      pastWeeksValue: (count) => `${count} ${count === 1 ? "row" : "rows"}`,
      bulletNone: "No bullets",
      bulletNumber: "Numbered lines",
      bulletDot: "Bulleted lines",
      bulletHint: "Each line of a note is prefixed for display only; the prefix is never saved.",
      showWeekend: "Weekend tint",
      showWeekendHint: "Wash Saturday and Sunday in a softer tone.",
      showToday: "Today marker",
      showTodayHint: "Mark today's number in the accent colour.",
      showFestivals: "Festival names",
      showFestivalsHint: "Lunar festivals and statutory holidays.",
      showHolidayBadges: "Day off / work marks",
      showHolidayBadgesHint: "Show the off and workday badges when a schedule exists.",
      showLunar: "Lunar date",
      showLunarHint: "The lunar day under the Gregorian number.",
      showTerms: "Solar terms",
      showTermsHint: "On a term day, show the term instead of the lunar date.",
      aboutName: "Calendar Notes",
      aboutTagline: "A month view for keeping notes, day by day.",
      aboutIntro:
        "Runs as a DBX workbench panel. Each cell carries its lunar date, solar term or festival, the statutory holiday arrangement when one is published, and the note you wrote on that day.",
      aboutFeatures: [
        "Month grid with lunar dates, solar terms and holiday markers",
        "Search every note, keyboard-first",
        "Export to .xls or .txt, import back from .xls or .xlsx",
        "Notes kept by the local sidecar, never in the page",
      ],
      aboutMeta: "Details",
      metaPublisher: "Publisher",
      metaLicense: "License",
      metaRequires: "Requires",
      metaBuilt: "Built with",
      metaHomepage: "Homepage",
      metaSource: "Source",
      aboutRights: (year) => `© ${year} tenltrs. Released under the Apache License 2.0.`,
      colDate: "Date",
      colContent: "Note",
      sheetName: "Notes",
      exported: (count) => `Exported ${count} ${count === 1 ? "entry" : "entries"}`,
      imported: (count, total, skipped) =>
        [`Imported ${count} ${count === 1 ? "entry" : "entries"}`, `${total} total`, skipped ? `${skipped} rows skipped` : null]
          .filter(Boolean)
          .join(", "),
      exportFailed: "Export failed",
      savedTo: (path) => `Saved to ${path}`,
      outputLabel: "Save to",
      choosePath: "Choose…",
      copyPath: "Copy path",
      openFolder: "Open folder",
      pathCopied: "Path copied",
      copyFailed: "Could not copy the path",
      openFolderFailed: "Could not open the folder",
      importFailed: "Could not read that file",
      saveFailed: "Could not save",
      fileExcel: "Excel sheet",
      fileText: "Text file",
      rangeLabel: "Date range",
      ranges: {
        "1m": "Past month",
        "3m": "Past 3 months",
        "6m": "Past 6 months",
        "1y": "Past year",
        all: "All dates",
        custom: "Custom",
      },
      entries: (count) => `${count} ${count === 1 ? "entry" : "entries"}`,
      nothingInRange: "Nothing in this range",
      export: "Export",
      chooseFile: "Choose a file",
      unsupportedFile: "Only .xls and .xlsx files can be imported",
      noStorage: "Notes cannot be saved here, and will be lost when the page closes. Export to keep them.",
      conflicts: (count) => `${count} conflicting ${count === 1 ? "date" : "dates"}`,
      modeOverwrite: "Overwrite",
      modeMerge: "Merge",
      modeSkip: "Skip",
      import: "Import",
    },
    zh: {
      prev: "上个月",
      next: "下个月",
      today: "今天",
      chooseMonth: "选择年月",
      previousYear: "上一年",
      nextYear: "下一年",
      yearLabel: "年份",
      weekdays: ["周一", "周二", "周三", "周四", "周五", "周六", "周日"],
      editorTitle: "编辑内容",
      placeholder: "写点什么…",
      save: "保存",
      cancel: "取消",
      clear: "清除",
      hint: "Ctrl + Enter 保存，Esc 关闭",
      hasNote: "已有内容",
      noNote: "暂无内容",
      dayOff: "休",
      workday: "班",
      edit: "双击编辑",
      search: "搜索",
      searchTitle: "搜索笔记",
      searchPlaceholder: "输入关键词搜索笔记…",
      searchPrompt: "输入关键词搜索笔记",
      searchEmpty: "没有找到匹配的笔记",
      searchCount: (count) => `找到 ${count} 条`,
      selectedDateLabel: "选中日期",
      closeSearch: "关闭搜索",
      searchHelpNavigate: "切换结果",
      searchHelpClose: "关闭",
      menu: "菜单",
      settings: "设置",
      settingsClose: "关闭设置",
      tabGrid: "日历",
      tabAbout: "关于我们",
      groupStyle: "日历样式",
      groupCellStyle: "单元格样式",
      groupDisplay: "显示",
      groupFestival: "节日",
      groupLunar: "农历节气",
      styleFull: "完整月份",
      styleFullNote: "始终从 1 号开始排列。",
      styleFloat: "浮动月份",
      styleFloatNote: "当期日期固定在设定的行。",
      pastWeeks: "已过星期数",
      pastWeeksHint: "当期日期上方保留的星期行数，0 即落在第一行。",
      pastWeeksValue: (count) => `${count} 行`,
      bulletNone: "无项目符号",
      bulletNumber: "项目数字编号",
      bulletDot: "项目圆点",
      bulletHint: "每行内容在显示时加上前缀，前缀不会写入文件。",
      showWeekend: "周末底色",
      showWeekendHint: "周六、周日用更淡的底色区分。",
      showToday: "今天标记",
      showTodayHint: "用强调色标出今天的日期。",
      showFestivals: "节日名称",
      showFestivalsHint: "农历节日与法定节日名称。",
      showHolidayBadges: "休 / 班标记",
      showHolidayBadgesHint: "有放假安排时显示「休」和「班」。",
      showLunar: "农历日期",
      showLunarHint: "公历日下方显示农历。",
      showTerms: "二十四节气",
      showTermsHint: "节气当天用它替代农历日期。",
      aboutName: "日历笔记",
      aboutTagline: "在月视图上，按天记录笔记。",
      aboutIntro:
        "以工作台面板运行在 DBX 中。每个日期格写着公历日、农历、节气或节日，遇到法定节假日还会标出放假与调休；那天写过的笔记也直接显示在格里。",
      aboutFeatures: [
        "月视图，附农历、节气与节假日标记",
        "全文搜索笔记，全程可用键盘操作",
        "导出 .xls 或 .txt，支持 .xls / .xlsx 导回",
        "笔记由本地 sidecar 保存，不落在页面里",
      ],
      aboutMeta: "详细信息",
      metaPublisher: "开发者",
      metaLicense: "许可证",
      metaRequires: "运行环境",
      metaBuilt: "技术栈",
      metaHomepage: "主页",
      metaSource: "源码",
      aboutRights: (year) => `© ${year} tenltrs，基于 Apache License 2.0 发布。`,
      colDate: "日期",
      colContent: "内容",
      sheetName: "笔记",
      exported: (count) => `已导出 ${count} 条`,
      imported: (count, total, skipped) =>
        [`已导入 ${count} 条`, `共 ${total} 条`, skipped ? `跳过 ${skipped} 行` : null].filter(Boolean).join("，"),
      exportFailed: "导出失败",
      savedTo: (path) => `已保存到 ${path}`,
      outputLabel: "输出地址",
      choosePath: "选择…",
      copyPath: "复制路径",
      openFolder: "打开文件夹",
      pathCopied: "已复制路径",
      copyFailed: "复制路径失败",
      openFolderFailed: "打开文件夹失败",
      importFailed: "无法读取该文件",
      saveFailed: "保存失败",
      fileExcel: "Excel 表格",
      fileText: "文本文件",
      rangeLabel: "日期范围",
      ranges: {
        "1m": "近1个月",
        "3m": "近3个月",
        "6m": "近半年",
        "1y": "近一年",
        all: "全部范围",
        custom: "自定义",
      },
      entries: (count) => `共 ${count} 条`,
      nothingInRange: "该范围内没有内容",
      export: "导出",
      chooseFile: "选择文件",
      unsupportedFile: "只支持 .xls 和 .xlsx 文件",
      noStorage: "此处无法保存笔记，页面关闭后会丢失，请先导出保存。",
      conflicts: (count) => `冲突 ${count} 条`,
      modeOverwrite: "覆盖",
      modeMerge: "合并",
      modeSkip: "跳过",
      import: "导入",
    },
  };

  // Chinese is the default: the plugin ships for a Chinese-first host, and a
  // page opened directly in a browser has no locale to read, so anything else
  // would show English to the people most likely to open it that way.
  //
  // A locale the bridge already carries is read synchronously here rather than
  // waited for, so an English host never paints a frame of Chinese first.
  // init() then confirms it against whatever the ready handshake reports.
  const bootstrapLocale = hostLocale();
  let lang = $state(
    bootstrapLocale && !bootstrapLocale.toLowerCase().startsWith("zh") ? "en" : "zh",
  );
  let text = $derived(copy[lang]);

  const now = new Date();
  let viewYear = $state(now.getFullYear());
  let viewMonth = $state(now.getMonth());

  let monthPickerOpen = $state(false);
  let monthPickerYear = $state(String(now.getFullYear()));
  let monthPickerEl = $state(null);
  let monthPickerPanel = $state(null);
  let monthPickerTrigger = $state(null);

  /** date key ("YYYY-MM-DD") -> note text */
  let notes = $state({});
  let editingKey = $state(null);
  let draft = $state("");
  /** The editor's field, focused whenever the editor opens. */
  let editorInput = $state(null);

  let menuOpen = $state(false);
  let menuEl = $state(null);
  let menuButton = $state(null);
  let settingsOpen = $state(false);
  /** Which settings pane the rail is showing. The dialog always reopens on
   *  the first pane so its shape is predictable. */
  let settingsTab = $state("grid");
  let settingsPanel = $state(null);

  /** How the grid is framed. "full" is one calendar month from its first day;
   *  "float" slides the window so an anchor day always lands on the same week
   *  row, which is how a planner reads dates — by distance from today, not by
   *  which month page they fall on. */
  let calendarStyle = $state("full");
  /** Week rows kept above the anchor day in the floating frame. */
  let pastWeeks = $state(1);
  /** How each line of a note is marked in the cell and the editor: no prefix,
   *  a running number, or a dot. Presentation only — the stored note is plain. */
  let cellBullet = $state("none");
  let showWeekend = $state(true);
  let showToday = $state(true);
  let showFestivals = $state(true);
  let showHolidayBadges = $state(true);
  let showLunar = $state(true);
  let showTerms = $state(true);
  /** Set once the saved preferences have been read, so the first save does not
   *  overwrite them before they are applied. */
  let settingsLoaded = $state(false);

  const SETTINGS_KEY = "dbx-calendar/settings";

  /** Fields the user can pick in the settings panel, with the narrow shape
   *  each one accepts. Centralised so the sidecar reader, the localStorage
   *  reader and the writer all validate the same way. */
  const SETTINGS_FIELDS = {
    style: (value) => (value === "full" || value === "float" ? value : null),
    pastWeeks: (value) => (Number.isInteger(value) ? Math.min(3, Math.max(0, value)) : null),
    bullet: (value) => (value === "none" || value === "number" || value === "dot" ? value : null),
    weekend: (value) => (typeof value === "boolean" ? value : null),
    today: (value) => (typeof value === "boolean" ? value : null),
    festivals: (value) => (typeof value === "boolean" ? value : null),
    holidayBadges: (value) => (typeof value === "boolean" ? value : null),
    lunar: (value) => (typeof value === "boolean" ? value : null),
    terms: (value) => (typeof value === "boolean" ? value : null),
  };

  /** Applies a saved record to the state variables. Unknown or malformed keys
   *  keep the current value, so an old cache can never break the panel. */
  function applySettings(saved) {
    if (!saved || typeof saved !== "object") return;
    if (saved.style !== undefined) calendarStyle = SETTINGS_FIELDS.style(saved.style) ?? calendarStyle;
    if (saved.pastWeeks !== undefined) pastWeeks = SETTINGS_FIELDS.pastWeeks(saved.pastWeeks) ?? pastWeeks;
    if (saved.bullet !== undefined) cellBullet = SETTINGS_FIELDS.bullet(saved.bullet) ?? cellBullet;
    if (saved.weekend !== undefined) showWeekend = SETTINGS_FIELDS.weekend(saved.weekend) ?? showWeekend;
    if (saved.today !== undefined) showToday = SETTINGS_FIELDS.today(saved.today) ?? showToday;
    if (saved.festivals !== undefined) showFestivals = SETTINGS_FIELDS.festivals(saved.festivals) ?? showFestivals;
    if (saved.holidayBadges !== undefined) showHolidayBadges = SETTINGS_FIELDS.holidayBadges(saved.holidayBadges) ?? showHolidayBadges;
    if (saved.lunar !== undefined) showLunar = SETTINGS_FIELDS.lunar(saved.lunar) ?? showLunar;
    if (saved.terms !== undefined) showTerms = SETTINGS_FIELDS.terms(saved.terms) ?? showTerms;
  }

  /** Preferences are a convenience, not data: when the host webview hides
   *  localStorage the calendar still works, it just forgets the choices. */
  function readLocalSettings() {
    try {
      return JSON.parse(localStorage.getItem(SETTINGS_KEY) ?? "null");
    } catch {
      // No storage, or a malformed record: keep the defaults.
      return null;
    }
  }

  function currentSettings() {
    return {
      style: calendarStyle,
      pastWeeks,
      bullet: cellBullet,
      weekend: showWeekend,
      today: showToday,
      festivals: showFestivals,
      holidayBadges: showHolidayBadges,
      lunar: showLunar,
      terms: showTerms,
    };
  }

  function writeLocalSettings(settings) {
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
    } catch {
      // Storage unavailable: the session keeps its choices, nothing else to do.
    }
  }

  const savedLocal = readLocalSettings();
  if (savedLocal) applySettings(savedLocal);

  settingsLoaded = true;
  /** Set once the sidecar has been asked for its copy. Until then the effect
   *  below must not write the sidecar: the in-page cache may hold stale values
   *  from an older release, and letting them win would defeat the upgrade
   *  path the persisted file exists for. */
  let sidecarChecked = $state(false);

  /** Pulls the persisted preferences from the sidecar once the bridge is up.
   *  The sidecar file survives reinstalls and upgrades, so it wins over the
   *  in-page cache whenever it has anything to say. */
  async function loadPersistedSettings() {
    const reply = await readSettingsStorage();
    if (reply.ok && reply.persisted && Object.keys(reply.settings).length > 0) {
      applySettings(reply.settings);
      writeLocalSettings(reply.settings);
    }
    sidecarChecked = true;
  }

  $effect(() => {
    if (!settingsLoaded) return;
    const settings = currentSettings();
    writeLocalSettings(settings);
    if (sidecarChecked) writeSettingsStorage(settings);
  });
  let searchOpen = $state(false);
  let searchQuery = $state("");
  let searchInput = $state(null);
  let searchButton = $state(null);
  /** True while the user is moving with the keyboard, so focus frames are wanted. */
  let keyboardFocus = $state(false);
  let activeSearchIndex = $state(0);
  let selectedDate = $state(null);
  let fileInput = $state(null);
  let status = $state("");
  let statusTimer;
  /** Set when the sidecar refuses to store notes, so the UI stops implying it did. */
  let notesError = $state("");

  let exportOpen = $state(false);
  let exportType = $state("xls");
  let exportRange = $state("all");
  let exportFrom = $state("");
  let exportTo = $state("");
  /** The absolute file the export will write, shown in the dialog and used as
   *  given. Picking one in the save dialog replaces it wholesale. */
  let exportPath = $state("");
  /** Set once the user named the file themselves, which is what tells the
   *  sidecar an existing file may be replaced. */
  let exportChosen = $state(false);

  let importOpen = $state(false);
  let importFileName = $state("");
  let importRows = $state([]);
  let importSkipped = $state(0);
  let importError = $state("");
  let importMode = $state("overwrite");

  const todayKey = keyOf(now.getFullYear(), now.getMonth(), now.getDate());

  const monthTitle = $derived(
    new Intl.DateTimeFormat(lang === "zh" ? "zh-CN" : "en-US", {
      year: "numeric",
      month: "long",
    }).format(new Date(viewYear, viewMonth, 1)),
  );

  /** Short month names for the picker grid, localized with the rest of the UI. */
  const monthNames = $derived(
    Array.from({ length: 12 }, (_, month) =>
      new Intl.DateTimeFormat(lang === "zh" ? "zh-CN" : "en-US", { month: "short" }).format(
        new Date(2024, month, 1),
      ),
    ),
  );

  const weeks = $derived.by(() => buildWeeks(viewYear, viewMonth));
  const searchResults = $derived(searchNotes(notes, searchQuery));

  /** Lunar date, solar term and holiday marker for the visible grid, resolved once per month. */
  const dayInfos = $derived.by(() => {
    const map = new Map();
    const layers = { festivals: showFestivals, terms: showTerms, lunar: showLunar };
    for (const week of weeks) for (const cell of week) map.set(cell.key, dayInfo(cell.key, lang, layers));
    return map;
  });

  function keyOf(year, month, day) {
    return `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  /** Rows in the floating frame: a fixed window keeps the calendar height
   *  steady while the dates slide from day to day. */
  const FLOAT_ROWS = 5;

  /** Six-column week rows starting on Monday, padded with the neighbouring months' days. */
  function buildWeeks(year, month) {
    return calendarStyle === "float" ? buildFloatingWeeks(year, month) : buildMonthWeeks(year, month);
  }

  function buildMonthWeeks(year, month) {
    const first = new Date(year, month, 1);
    const lead = (first.getDay() + 6) % 7;
    const rows = Math.ceil((lead + new Date(year, month + 1, 0).getDate()) / 7);

    return Array.from({ length: rows }, (_, row) =>
      Array.from({ length: 7 }, (_, col) => {
        const date = new Date(year, month, 1 - lead + row * 7 + col);
        return {
          key: keyOf(date.getFullYear(), date.getMonth(), date.getDate()),
          day: date.getDate(),
          inMonth: date.getMonth() === month,
          weekend: col >= 5,
        };
      }),
    );
  }

  /** Week rows in the floating frame: the anchor day keeps the same row every
   *  month, so the weeks above and below it are the dates a reader actually
   *  plans with. The anchor is today's day of month, clamped for short months,
   *  which keeps the current month identical to "today is on this row".
   *
   *  Every cell is in range here: the window deliberately runs past the month
   *  boundaries, so dimming the neighbours would grey out most of the grid. */
  function buildFloatingWeeks(year, month) {
    const lastDay = new Date(year, month + 1, 0).getDate();
    const anchorDay = Math.min(now.getDate(), lastDay);
    const anchor = new Date(year, month, anchorDay);
    const lead = (anchor.getDay() + 6) % 7;
    const start = new Date(year, month, anchorDay - lead - pastWeeks * 7);

    return Array.from({ length: FLOAT_ROWS }, (_, row) =>
      Array.from({ length: 7 }, (_, col) => {
        const date = new Date(start.getFullYear(), start.getMonth(), start.getDate() + row * 7 + col);
        return {
          key: keyOf(date.getFullYear(), date.getMonth(), date.getDate()),
          day: date.getDate(),
          inMonth: true,
          weekend: col >= 5,
          /** The 1st of a month carries its month name instead of the bare
           *  number: in a window that runs across months, "1" alone would not
           *  say which month it opens. */
          monthLabel: date.getDate() === 1
            ? new Intl.DateTimeFormat(lang === "zh" ? "zh-CN" : "en-US", { month: "long" }).format(date)
            : null,
        };
      }),
    );
  }

  function step(delta) {
    const next = new Date(viewYear, viewMonth + delta, 1);
    viewYear = next.getFullYear();
    viewMonth = next.getMonth();
  }

  function goToday() {
    viewYear = now.getFullYear();
    viewMonth = now.getMonth();
  }

  function openMonthPicker() {
    menuOpen = false;
    // The year field starts from the grid so the steppers continue from where
    // the user is; the highlighted month is always the grid's own year+month,
    // so stepping the year to another one leaves nothing highlighted.
    monthPickerYear = String(viewYear);
    monthPickerOpen = true;
  }

  function closeMonthPicker({ restoreFocus = true } = {}) {
    if (!monthPickerOpen) return;
    monthPickerOpen = false;
    if (restoreFocus) monthPickerTrigger?.focus();
  }

  function toggleMonthPicker() {
    if (monthPickerOpen) closeMonthPicker();
    else openMonthPicker();
  }

  /** The pending year, or null while the field is incomplete/out of range. */
  function pickerYearValue() {
    const year = Number(monthPickerYear);
    return monthPickerYear.length === 4 && Number.isInteger(year) && year >= 1900 && year <= 2200
      ? year
      : null;
  }

  /** Steps the pending year without moving the calendar until a month is chosen. */
  function stepPickerYear(delta) {
    monthPickerYear = String(Math.min(2200, Math.max(1900, (pickerYearValue() ?? viewYear) + delta)));
  }

  function normalizePickerYear() {
    // A partial year (e.g. "20" on the way to 2020) falls back rather than
    // silently sending the calendar to 1900 when focus leaves the field.
    monthPickerYear = String(pickerYearValue() ?? viewYear);
  }

  /** Keeps the year field to four digits while it is being typed. */
  function inputPickerYear(event) {
    const digits = event.currentTarget.value.replace(/\D/g, "").slice(0, 4);
    event.currentTarget.value = digits;
    monthPickerYear = digits;
  }

  function chooseMonth(month) {
    normalizePickerYear();
    const year = Number(monthPickerYear);
    if (year < 1900 || year > 2200) return;
    viewYear = year;
    viewMonth = month;
    closeMonthPicker();
  }

  /** True for the one month the grid is actually showing, and only while the
   *  year field still names that year. Stepping to another year clears it:
   *  the pill marks a year+month pair, not a month index that rides along. */
  function isChosenMonth(index) {
    return index === viewMonth && pickerYearValue() === viewYear;
  }

  // Opening the panel puts focus on the pending month so it can be confirmed
  // with Enter and walked with Tab, rather than leaving focus on the trigger.
  $effect(() => {
    if (!monthPickerOpen || !monthPickerPanel) return;
    monthPickerPanel.querySelector(".picker-month.active")?.focus();
  });

  function openSearch() {
    menuOpen = false;
    monthPickerOpen = false;
    searchQuery = "";
    activeSearchIndex = 0;
    searchOpen = true;
  }

  /** The rail is built from the copy so its order and labels stay in one place.
   *  Icons travel with the label for the same reason: one definition feeds both
   *  the rail and anything else that has to name a pane. */
  const settingsTabs = $derived([
    { id: "grid", label: text.tabGrid, icon: calendarIcon },
    { id: "about", label: text.tabAbout, icon: infoIcon },
  ]);

  /** Toggle rows are described as data so the three groups render the same
   *  label-hint-switch shape without three copies of it in the markup. */
  const displaySettings = $derived([
    { key: "weekend", label: text.showWeekend, hint: text.showWeekendHint, value: showWeekend, set: (on) => (showWeekend = on) },
    { key: "today", label: text.showToday, hint: text.showTodayHint, value: showToday, set: (on) => (showToday = on) },
  ]);

  const festivalSettings = $derived([
    { key: "festivals", label: text.showFestivals, hint: text.showFestivalsHint, value: showFestivals, set: (on) => (showFestivals = on) },
    { key: "badges", label: text.showHolidayBadges, hint: text.showHolidayBadgesHint, value: showHolidayBadges, set: (on) => (showHolidayBadges = on) },
  ]);

  const lunarSettings = $derived([
    { key: "lunar", label: text.showLunar, hint: text.showLunarHint, value: showLunar, set: (on) => (showLunar = on) },
    { key: "terms", label: text.showTerms, hint: text.showTermsHint, value: showTerms, set: (on) => (showTerms = on) },
  ]);

  function openSettings() {
    menuOpen = false;
    monthPickerOpen = false;
    settingsTab = "grid";
    settingsOpen = true;
  }

  function closeSettings() {
    settingsOpen = false;
    // Focus goes back to the menu it was opened from, so a keyboard user is
    // where they left off rather than at the top of the document.
    menuButton?.focus();
  }

  // Focus lands on the pane already showing: a keyboard user starts inside the
  // sheet rather than at the top of the document, and Tab moves on from there.
  $effect(() => {
    if (!settingsOpen || !settingsPanel) return;
    settingsPanel.querySelector(".settings-tab.active")?.focus();
  });

  // Closing returns focus to the trigger so typing continues from where it was,
  // but that button never paints a focus ring of its own.
  function closeSearch() {
    searchOpen = false;
    searchButton?.focus();
  }

  /** The one selection state shared by search results and dates opened from the grid. */
  function selectDate(date) {
    selectedDate = date;
  }

  function selectSearchResult(date) {
    const [year, month] = date.split("-").map(Number);
    viewYear = year;
    viewMonth = month - 1;
    selectDate(date);
    closeSearch();
  }

  /** The prefix a fresh note starts with, so a new editor already reads as a
   *  list instead of waiting for the first Enter. */
  function firstBulletPrefix() {
    return cellBullet === "number" ? "1. " : "• ";
  }

  function openEditor(cell) {
    // Opening the editor is also an explicit choice of that day, so it takes
    // the same amber selection as a search result.
    selectDate(cell.key);
    editingKey = cell.key;
    // Prefixes are part of the editing surface, not the note: showing them here
    // keeps the textarea and the cell reading the same, and stripping them on
    // save keeps the stored text plain.
    draft = decorateNote(notes[cell.key] ?? "");
    if (!draft && cellBullet !== "none") draft = firstBulletPrefix();
  }

  /**
   * A line prefix is derived from the note text each time it renders, so the
   * number always follows the line order and nothing extra reaches storage.
   * Blank lines stay blank: they are spacing, not items.
   */
  function decorateNote(value) {
    if (cellBullet === "none" || !value) return value;
    let item = 0;
    return value
      .split("\n")
      .map((line) => {
        if (!line.trim()) return line;
        item += 1;
        return cellBullet === "number" ? `${item}. ${line}` : `• ${line}`;
      })
      .join("\n");
  }

  /** Removes exactly the prefixes this style introduces, and nothing a person
   *  may have typed themselves in a different shape. */
  function stripNotePrefixes(value) {
    if (cellBullet === "none") return value;
    const pattern = cellBullet === "number" ? /^\s*\d+[.、)]\s?/ : /^\s*[•·]\s?/;
    return value
      .split("\n")
      .map((line) => line.replace(pattern, ""))
      .join("\n");
  }

  /** Enter in the editor starts the next item, carrying the same prefix the
   *  cell would show. Shift+Enter keeps its native newline so a note can still
   *  hold a soft break. */
  function handleEditorKeydown(event) {
    if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) {
      event.preventDefault();
      saveNote();
      return;
    }
    if (event.key !== "Enter" || event.shiftKey || event.altKey || cellBullet === "none") {
      return;
    }
    event.preventDefault();
    const field = event.currentTarget;
    const start = field.selectionStart;
    const before = field.value.slice(0, start);
    const item = before.split("\n").filter((line) => line.trim()).length + 1;
    const prefix = cellBullet === "number" ? `${item}. ` : "• ";
    const inserted = `\n${prefix}`;
    field.value = field.value.slice(0, start) + inserted + field.value.slice(field.selectionEnd);
    const caret = start + inserted.length;
    field.setSelectionRange(caret, caret);
    draft = field.value;
  }

  function closeEditor() {
    editingKey = null;
    draft = "";
  }

  // Typing has to be possible the moment the editor appears, so the field is
  // focused here rather than left to the `autofocus` attribute: that one is
  // unreliable for an element added after the page loaded, and this UI lives in
  // an iframe the host may not have focused. The caret goes to the end, which is
  // where a note usually continues from.
  $effect(() => {
    if (!editingKey || !editorInput) return;
    editorInput.focus();
    const end = editorInput.value.length;
    editorInput.setSelectionRange(end, end);
  });

  $effect(() => {
    if (!searchOpen || !searchInput) return;
    searchInput.focus();
  });

  async function saveNote() {
    // A draft that only holds prefixes was never really written: saving it
    // would store an empty entry over a note the user may have meant to keep.
    const value = stripNotePrefixes(draft).trim();
    if (!value) {
      closeEditor();
      return;
    }
    await writeNote(editingKey, value);
    closeEditor();
  }

  async function removeNote() {
    await writeNote(editingKey, "");
    closeEditor();
  }

  function deleteNote(key) {
    const { [key]: _removed, ...rest } = notes;
    notes = rest;
  }

  /**
   * The sidecar owns the note set: it stores the day and answers with the whole
   * map, so the grid never diverges from what is actually persisted. An empty
   * value clears the day.
   */
  async function writeNote(key, value) {
    const result = await invoke("dbx-calendar/notes/set", { date: key, value });
    if (!result.ok) {
      // A failure leaves the grid untouched rather than showing a note that was
      // never stored.
      notesError = result.error;
      flash(`${text.saveFailed}: ${result.error}`);
      return;
    }
    notesError = "";
    notes = result.value?.notes ?? {};
  }

  /**
   * Replaces the whole note set at once, which is what an import does. Posting
   * every date separately would be one round trip per row and could leave a
   * half-applied import behind if one of them failed.
   */
  async function replaceNotes(next) {
    const result = await invoke("dbx-calendar/notes/replace", { notes: next });
    if (!result.ok) {
      notesError = result.error;
      flash(`${text.saveFailed}: ${result.error}`);
      return false;
    }
    notesError = "";
    notes = result.value?.notes ?? {};
    return true;
  }

  async function loadNotes() {
    const result = await invoke("dbx-calendar/notes/list");
    if (!result.ok) {
      notesError = result.error;
      notes = {};
      return;
    }
    notesError = "";
    notes = result.value?.notes ?? {};
  }

  function flash(message) {
    status = message;
    clearTimeout(statusTimer);
    statusTimer = setTimeout(() => (status = ""), 4000);
  }

  function toggleMenu() {
    menuOpen = !menuOpen;
  }

  function handlePointerDown() {
    // A pointer click is not a request for a keyboard focus frame; pressing
    // Escape afterwards should leave the toolbar looking untouched.
    keyboardFocus = false;
  }

  function handleWindowClick(event) {
    if (menuOpen && menuEl && !menuEl.contains(event.target)) menuOpen = false;
    if (
      monthPickerOpen &&
      monthPickerEl &&
      !monthPickerEl.contains(event.target) &&
      !monthPickerPanel?.contains(event.target)
    ) {
      // The click that dismissed the panel already moved focus somewhere useful.
      closeMonthPicker({ restoreFocus: false });
    }
  }

  function handleKeydown(event) {
    // Only focus navigation asks for a visible frame. Typing, arrow keys inside
    // a dialog and Escape must not turn a pointer click into a focus ring.
    if (event.key === "Tab") keyboardFocus = true;

    // The settings sheet is modal, so Tab cycles inside it instead of leaking
    // out to the toolbar behind the scrim.
    if (event.key === "Tab" && settingsOpen && settingsPanel) {
      const focusable = [...settingsPanel.querySelectorAll("button, a[href], [tabindex]:not([tabindex='-1'])")]
        .filter((el) => el.offsetParent !== null);
      if (focusable.length) {
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        const active = document.activeElement;
        if (!settingsPanel.contains(active)) {
          event.preventDefault();
          first.focus();
          return;
        }
        if (event.shiftKey && active === first) {
          event.preventDefault();
          last.focus();
          return;
        }
        if (!event.shiftKey && active === last) {
          event.preventDefault();
          first.focus();
          return;
        }
      }
    }

    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "f") {
      keyboardFocus = true;
      // Do not open search on top of another modal; there is nothing to search
      // from until that one is dealt with.
      if (editingKey || exportOpen || importOpen || settingsOpen) return;
      // Suppress the host's own Find UI only when this shortcut is actually ours.
      event.preventDefault();
      if (searchOpen) {
        searchInput?.focus();
        searchInput?.select();
      } else {
        openSearch();
      }
      return;
    }

    if (event.key === "Escape") {
      // The editor and the export dialog are modal, so they take Escape first.
      if (editingKey) closeEditor();
      else if (exportOpen) exportOpen = false;
      else if (importOpen) closeImport();
      else if (settingsOpen) closeSettings();
      else if (monthPickerOpen) closeMonthPicker();
      else if (searchOpen) closeSearch();
      else if (menuOpen) menuOpen = false;
      return;
    }

    if (!searchOpen || event.target !== searchInput) return;

    const navigation = navigateSearchResults(activeSearchIndex, searchResults.length, event.key);
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      activeSearchIndex = navigation.index;
      if (activeSearchIndex >= 0) {
        document.getElementById(`search-result-${activeSearchIndex}`)?.scrollIntoView({ block: "nearest" });
      }
    } else if (navigation.activate) {
      event.preventDefault();
      selectSearchResult(searchResults[navigation.index].date);
    }
  }

  /**
   * Reads the date column as it can realistically come back: the "YYYY-MM-DD"
   * text we write, a real date cell (Excel converts a typed date into one), or
   * the serial number behind such a cell. Returns null for anything else, which
   * is also what makes the header row and blank rows fall away on import.
   */
  function dateKeyOf(value) {
    if (value instanceof Date && !Number.isNaN(+value)) {
      return keyOf(value.getFullYear(), value.getMonth(), value.getDate());
    }

    if (typeof value === "number") {
      // Excel day 1 is 1900-01-01, offset by its imaginary 1900-02-29.
      const serial = new Date(Date.UTC(1899, 11, 30) + value * 86_400_000);
      return keyOf(serial.getUTCFullYear(), serial.getUTCMonth(), serial.getUTCDate());
    }

    const match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(value ?? "").trim());
    if (!match) return null;

    const [year, month, day] = match.slice(1).map(Number);
    const date = new Date(year, month - 1, day);
    // Date rolls impossible values over (2026-02-30 becomes 2026-03-02), so
    // compare the parts back to reject them.
    return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day
      ? keyOf(year, month - 1, day)
      : null;
  }

  /** Subtracts whole months, clamping to the target month's length so that
   *  31 March minus one month lands on 28/29 February instead of rolling into March. */
  function shiftMonths(date, months) {
    const shifted = new Date(date.getFullYear(), date.getMonth() - months, 1);
    const lastDay = new Date(shifted.getFullYear(), shifted.getMonth() + 1, 0).getDate();
    shifted.setDate(Math.min(date.getDate(), lastDay));
    return shifted;
  }

  function boundsFor(range) {
    if (range === "all") {
      const keys = Object.keys(notes).sort();
      return keys.length ? [keys[0], keys[keys.length - 1]] : [todayKey, todayKey];
    }
    const start = shiftMonths(now, { "1m": 1, "3m": 3, "6m": 6, "1y": 12 }[range]);
    return [keyOf(start.getFullYear(), start.getMonth(), start.getDate()), todayKey];
  }

  async function openExport() {
    menuOpen = false;
    monthPickerOpen = false;
    exportType = "xls";
    exportRange = "all";
    const [from, to] = boundsFor("all");
    exportFrom = from;
    exportTo = to;
    exportChosen = false;
    exportOpen = true;
    // The dialog opens on a path rather than filling one in later, so the user
    // sees where the file is going before deciding to write it.
    await loadDestination();
  }

  /** Switching format keeps the folder and the stem, so a renamed file stays
   *  renamed and the destination does not jump back to the default. */
  function setExportType(type) {
    exportType = type;
    exportPath = replaceExtension(exportPath, type);
  }

  function pickRange(range) {
    exportRange = range;
    // Custom leaves the pickers alone so they keep whatever the user last saw.
    if (range === "custom") return;
    const [from, to] = boundsFor(range);
    exportFrom = from;
    exportTo = to;
  }

  /** Notes inside the closed range, oldest first. Date keys sort chronologically
   *  as plain strings, so the comparison needs no parsing. */
  const exportRows = $derived.by(() => {
    if (!exportOpen) return [];
    return Object.keys(notes)
      .filter((key) => key >= (exportFrom || "0000-00-00") && key <= (exportTo || "9999-99-99"))
      .sort()
      .map((key) => [key, notes[key]]);
  });

  /** The file name an export of this format would like to use. */
  function exportFileName(type) {
    return `calendar-notes-${todayKey}.${type}`;
  }

  /** The UI cannot join paths, and the host hands back a native one, so this
   *  only has to supply whichever separator the folder is missing. */
  function joinPath(directory, name) {
    if (!directory) return name;
    const separator = directory.includes("\\") ? "\\" : "/";
    return directory.endsWith(separator) ? `${directory}${name}` : `${directory}${separator}${name}`;
  }

  /** Swaps the extension, keeping the folder and the stem the user may have
   *  renamed: switching format should not move the file or rename it back. */
  function replaceExtension(path, extension) {
    // No destination to rewrite yet; an invented ".txt" would be worse than an
    // empty field, which at least shows that the host never answered.
    if (!path) return "";
    const cut = path.lastIndexOf(".");
    const slash = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
    return `${cut > slash ? path.slice(0, cut) : path}.${extension}`;
  }

  /**
   * Asks the host where an export should go: the folder the user last chose, or
   * their downloads folder the first time.
   */
  async function loadDestination() {
    const result = await invoke("dbx-calendar/export/destination");
    const directory = result.ok ? (result.value?.directory ?? "") : "";
    exportPath = joinPath(directory, exportFileName(exportType));
  }

  /** Opens the OS save dialog. Whatever it returns becomes the destination, so
   *  the user can rename the file as well as move it. */
  async function chooseExportPath() {
    const result = await invoke("dbx-calendar/export/pick", {
      initial: exportPath,
      title: text.export,
      label: exportType === "txt" ? text.fileText : text.fileExcel,
      extension: exportType,
    });
    if (!result.ok) {
      flash(`${text.exportFailed}: ${result.error}`);
      return;
    }
    if (result.value?.cancelled) return;
    const chosen = result.value?.path;
    if (typeof chosen === "string" && chosen) {
      exportPath = chosen;
      // Named by hand in a dialog that already asked about replacing it, so the
      // export writes exactly here rather than beside an existing file.
      exportChosen = true;
    }
  }

  async function copyExportPath() {
    const result = await invoke("dbx-calendar/export/copy", { text: exportPath });
    flash(result.ok ? text.pathCopied : `${text.copyFailed}: ${result.error}`);
  }

  async function revealExportPath() {
    const result = await invoke("dbx-calendar/export/reveal", { path: exportPath });
    if (!result.ok) flash(`${text.openFolderFailed}: ${result.error}`);
  }

  /** The browser's own download, for a page opened outside the host: there is
   *  no sidecar to write for us there, and no sandbox to drop the click. */
  function anchorDownload(blob, filename) {
    try {
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      link.click();
      // Revoking straight away can cancel the download before it starts.
      setTimeout(() => URL.revokeObjectURL(url), 10_000);
    } catch (error) {
      flash(`${text.exportFailed}: ${error.message}`);
    }
  }

  /**
   * Writes the file through sidecar and reports where it landed. The reply is
   * the only evidence either side of this feature can trust, which is why it is
   * attempted before any browser route.
   */
  async function runExport() {
    const rows = exportRows;
    if (!rows.length) return;

    try {
      let blob;
      if (exportType === "txt") {
        // Named `body` rather than `text`: `text` is the copy table.
        // A note may hold line breaks of its own, which would otherwise split
        // one record across several lines and leave everything after the first
        // without a date. They travel as the two characters `\n` instead, so
        // every record stays exactly one line: date, tab, note. CR and CRLF are
        // folded in too — both are "a line break here", whatever the note was
        // pasted from.
        const oneLine = (value) => value.replace(/\r\n?|\n/g, "\\n");
        const body = rows.map(([key, content]) => `${key}\t${oneLine(content)}`).join("\n");
        blob = new Blob([`${body}\n`], { type: "text/plain;charset=utf-8" });
      } else {
        const sheet = XLSX.utils.aoa_to_sheet([[text.colDate, text.colContent], ...rows]);
        sheet["!cols"] = [{ wch: 12 }, { wch: 60 }];
        const book = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(book, sheet, text.sheetName);
        blob = new Blob([XLSX.write(book, { bookType: "xls", type: "array" })], {
          type: "application/vnd.ms-excel",
        });
      }

      // Outside the host there is no sidecar to write for us, and no sandbox to
      // drop the click either — a browser tab still downloads the file.
      if (!bridgeAvailable) {
        anchorDownload(blob, exportFileName(exportType));
        exportOpen = false;
        flash(text.exported(rows.length));
        return;
      }

      const bytes = new Uint8Array(await blob.arrayBuffer());
      // Chunked so a large workbook cannot blow the argument limit of from().
      let binary = "";
      for (let index = 0; index < bytes.length; index += 0x8000) {
        binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
      }

      const result = await invoke("dbx-calendar/export/write", {
        path: exportPath,
        data: btoa(binary),
        overwrite: exportChosen,
      });
      if (!result.ok) {
        flash(`${text.exportFailed}: ${result.error}`);
        return;
      }

      exportOpen = false;
      // The reply names the file that was actually written, which is not always
      // the one asked for: a taken name gains " (2)" rather than replacing it.
      const written = result.value?.path || exportPath;
      flash([text.exported(rows.length), text.savedTo(written)].filter(Boolean).join(" · "));
    } catch (error) {
      flash(`${text.exportFailed}: ${error.message}`);
    }
  }

  function openImport() {
    menuOpen = false;
    monthPickerOpen = false;
    importFileName = "";
    importRows = [];
    importSkipped = 0;
    importError = "";
    importMode = "overwrite";
    importOpen = true;
  }

  function closeImport() {
    importOpen = false;
    importFileName = "";
    importRows = [];
    importSkipped = 0;
    importError = "";
  }

  /** Reads the first sheet into [dateKey, content] pairs, dropping rows whose
   *  first cell is not a date — which is how the header and blanks fall away. */
  async function readSheet(file) {
    const book = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    const sheet = book.Sheets[book.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });

    const parsed = [];
    let skipped = 0;
    for (const [rawDate, rawContent] of rows) {
      const key = dateKeyOf(rawDate);
      const content = String(rawContent ?? "").trim();
      if (!key) {
        if (String(rawDate ?? "").trim() || content) skipped += 1;
        continue;
      }
      if (content) parsed.push([key, content]);
    }
    return { parsed, skipped };
  }

  /** The accept attribute only filters the picker; a user can still switch to
   *  "All files", so the extension is checked again here. */
  function isImportable(file) {
    const extension = file.name.split(".").pop()?.toLowerCase();
    return extension === "xls" || extension === "xlsx";
  }

  async function pickImportFile(event) {
    const file = event.target.files?.[0];
    // Clear the input so picking the same file twice still fires a change.
    event.target.value = "";
    if (!file) return;

    importFileName = file.name;
    importRows = [];
    importSkipped = 0;
    importMode = "overwrite";
    importError = "";

    if (!isImportable(file)) {
      importError = text.unsupportedFile;
      return;
    }

    try {
      const { parsed, skipped } = await readSheet(file);
      importRows = parsed;
      importSkipped = skipped;
    } catch (error) {
      importError = `${text.importFailed}: ${error.message}`;
    }
  }

  /** Dates the file and the calendar both have an entry for. */
  const importConflicts = $derived.by(
    () => importRows.filter(([key]) => notes[key] !== undefined).length,
  );

  async function runImport() {
    if (!importRows.length) return;

    const next = { ...notes };
    for (const [key, content] of importRows) {
      if (next[key] === undefined) {
        // Nothing to conflict with, so every mode adds it.
        next[key] = content;
        continue;
      }
      if (importMode === "skip") continue;
      next[key] = importMode === "merge" ? `${next[key]}\n${content}` : content;
    }

    const count = importRows.length;
    // Report only what the sidecar actually accepted.
    if (!(await replaceNotes(next))) return;
    closeImport();
    flash(text.imported(count, Object.keys(notes).length, importSkipped));
  }

  /**
   * Notes live in the Go sidecar, not in the page: the plugin frame runs on an
   * opaque origin, so localStorage throws on every access. The bridge is absent
   * when the page is opened directly in a browser, which leaves the calendar
   * functional but unable to store anything.
   */
  async function init() {
    // An explicit host locale always wins, in both directions: the mount
    // default is only what applies when nobody tells us anything.
    const context = await whenReady();
    const locale = context?.locale ?? hostLocale();
    if (locale) lang = locale.toLowerCase().startsWith("zh") ? "zh" : "en";

    // The sidecar file survives reinstalls and upgrades, so it is pulled after
    // the bridge is up and wins over the in-page cache.
    await loadPersistedSettings();
    await loadNotes();
  }

  onMount(() => {
    void init();
  });

  // The document keeps the same language as the copy, so assistive tech and the
  // browser's own spellcheck/line-breaking follow a switch to English too.
  $effect(() => {
    document.documentElement.lang = lang === "zh" ? "zh-CN" : "en";
  });
</script>

<svelte:head><title>DBX Calendar</title></svelte:head>
<svelte:window
  onkeydown={handleKeydown}
  onpointerdown={handlePointerDown}
  onclick={handleWindowClick}
/>

<main>
  <header class="toolbar" class:keyboard-focus={keyboardFocus}>
    <div class="month-menu" bind:this={monthPickerEl}>
      <h1>
        <button
          type="button"
          class="month-trigger"
          bind:this={monthPickerTrigger}
          aria-haspopup="dialog"
          aria-expanded={monthPickerOpen}
          aria-controls="month-picker"
          aria-label={`${monthTitle}, ${text.chooseMonth}`}
          onclick={toggleMonthPicker}
        >
          <span>{monthTitle}</span>
          <svg class="chevron" viewBox="0 0 16 16" aria-hidden="true" focusable="false">
            <path
              d="m4 6 4 4 4-4"
              fill="none"
              stroke="currentColor"
              stroke-width="1.5"
              stroke-linecap="round"
              stroke-linejoin="round"
            />
          </svg>
        </button>
      </h1>

    </div>
    <div class="actions">
      <button
        type="button"
        class="menu-button search-button"
        bind:this={searchButton}
        aria-label={text.search}
        aria-expanded={searchOpen}
        onclick={openSearch}
      >
        <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
          <circle cx="6.8" cy="6.8" r="4.6" fill="none" stroke="currentColor" stroke-width="1.5" />
          <path d="m10.2 10.2 3.1 3.1" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" />
        </svg>
      </button>
      <div class="menu" bind:this={menuEl}>
        <button
          type="button"
          class="menu-button"
          bind:this={menuButton}
          aria-label={text.menu}
          aria-expanded={menuOpen}
          onclick={toggleMenu}
        >
          <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
            <path
              d="M2 4h12M2 8h12M2 12h12"
              fill="none"
              stroke="currentColor"
              stroke-width="1.6"
              stroke-linecap="round"
            />
          </svg>
        </button>
        {#if menuOpen}
          <div class="dropdown">
            <button type="button" onclick={openSettings}>
              <span class="dropdown-icon" style={maskVar(settingsIcon)} aria-hidden="true"></span>
              {text.settings}
            </button>
            <button type="button" onclick={openExport}>
              <span class="dropdown-icon" style={maskVar(exportIcon)} aria-hidden="true"></span>
              {text.export}
            </button>
            <button type="button" onclick={openImport}>
              <span class="dropdown-icon" style={maskVar(importIcon)} aria-hidden="true"></span>
              {text.import}
            </button>
          </div>
        {/if}
      </div>

      <div class="nav">
        <button
          type="button"
          class="arrow"
          style={maskVar(previousIcon)}
          onclick={() => step(-1)}
          aria-label={text.prev}
        ></button>
        <button type="button" class="today" onclick={goToday}>{text.today}</button>
        <button
          type="button"
          class="arrow"
          style={maskVar(nextIcon)}
          onclick={() => step(1)}
          aria-label={text.next}
        ></button>
      </div>

      <input
        type="file"
        accept=".xls,.xlsx"
        bind:this={fileInput}
        onchange={pickImportFile}
        hidden
      />
    </div>

    {#if monthPickerOpen}
      <div
        class="month-picker"
        id="month-picker"
        role="dialog"
        aria-label={text.chooseMonth}
        bind:this={monthPickerPanel}
      >
        <div class="picker-year">
          <button
            type="button"
            class="picker-step"
            aria-label={text.previousYear}
            onclick={() => stepPickerYear(-1)}
          >
            <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
              <path
                d="M9.5 3.5 5.5 8l4 4.5"
                fill="none"
                stroke="currentColor"
                stroke-width="1.6"
                stroke-linecap="round"
                stroke-linejoin="round"
              />
            </svg>
          </button>

          <div class="picker-year-field">
            <input
              class="picker-year-input"
              type="text"
              inputmode="numeric"
              maxlength="4"
              value={monthPickerYear}
              aria-label={text.yearLabel}
              onfocus={(event) => event.currentTarget.select()}
              oninput={inputPickerYear}
              onblur={normalizePickerYear}
              onkeydown={(event) => {
                if (event.key !== "Enter") return;
                event.preventDefault();
                normalizePickerYear();
              }}
            />
            {#if lang === "zh"}<span class="picker-year-suffix" aria-hidden="true">年</span>{/if}
          </div>

          <button
            type="button"
            class="picker-step"
            aria-label={text.nextYear}
            onclick={() => stepPickerYear(1)}
          >
            <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
              <path
                d="m6.5 3.5 4 4.5-4 4.5"
                fill="none"
                stroke="currentColor"
                stroke-width="1.6"
                stroke-linecap="round"
                stroke-linejoin="round"
              />
            </svg>
          </button>
        </div>

        <div class="picker-months" role="group" aria-label={text.chooseMonth}>
          {#each monthNames as name, index}
            <button
              type="button"
              class="picker-month"
              class:active={isChosenMonth(index)}
              aria-pressed={isChosenMonth(index)}
              onclick={() => chooseMonth(index)}
            >
              {name}
            </button>
          {/each}
        </div>
      </div>
    {/if}
  </header>

  {#if notesError}
    <p class="notice">{text.noStorage}: {notesError}</p>
  {/if}

  <div
    class="grid"
    style="grid-template-rows: auto repeat({weeks.length}, minmax(0, 1fr));"
  >
    {#each text.weekdays as label, index (label)}
      <div class="weekday" class:weekend={showWeekend && index >= 5}>{label}</div>
    {/each}

    {#each weeks as week, weekIndex (weekIndex)}
      {#each week as cell (cell.key)}
        {@const info = dayInfos.get(cell.key)}
        <button
          type="button"
          class="cell"
          class:outside={!cell.inMonth}
          class:weekend={showWeekend && cell.weekend}
          class:today={showToday && cell.key === todayKey}
          class:selected={cell.key === selectedDate}
          aria-label="{cell.key}, {info.text}{info.status
            ? `, ${info.status === 'off' ? text.dayOff : text.workday}`
            : ''}, {notes[cell.key] ? text.hasNote : text.noNote}, {text.edit}{cell.key === selectedDate
            ? `, ${text.selectedDateLabel}`
            : ''}"
          ondblclick={() => openEditor(cell)}
          onclick={(event) => {
            // One click marks the day (the amber frame), which is all a pointer
            // user needs to point at a date. The editor still needs the double
            // click. Keyboard activation reports detail 0 and never fires
            // dblclick, so Enter/Space opens the editor directly.
            selectDate(cell.key);
            if (event.detail === 0) openEditor(cell);
          }}
        >
          <span class="cell-head">
            {#if cell.monthLabel}
              <span class="date month-day">{cell.monthLabel}</span>
            {:else}
              <span class="date">{cell.day}</span>
            {/if}
            {#if info.text}
              <span class="lunar">{info.text}</span>
            {/if}
            {#if info.status && showHolidayBadges}
              <span class="badge" class:off={info.status === "off"}>
                {info.status === "off" ? text.dayOff : text.workday}
              </span>
            {/if}
          </span>
          {#if notes[cell.key]}
            <span class="note" title={decorateNote(notes[cell.key])}>{decorateNote(notes[cell.key])}</span>
          {/if}
        </button>
      {/each}
    {/each}
  </div>
</main>

{#if searchOpen}
  <div class="backdrop search-backdrop">
    <div class="search-dialog" role="dialog" aria-modal="true" aria-labelledby="search-title" tabindex="-1">
      <header class="search-head">
        <h2 id="search-title">{text.searchTitle}</h2>
        <button
          type="button"
          class="search-close"
          aria-label={text.closeSearch}
          onclick={() => closeSearch()}
        >
          <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
            <path d="m4 4 8 8M12 4l-8 8" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" />
          </svg>
        </button>
      </header>

      <label class="search-field">
        <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
          <circle cx="6.8" cy="6.8" r="4.6" fill="none" stroke="currentColor" stroke-width="1.5" />
          <path d="m10.2 10.2 3.1 3.1" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" />
        </svg>
        <input
          type="search"
          bind:this={searchInput}
          bind:value={searchQuery}
          placeholder={text.searchPlaceholder}
          aria-label={text.searchPlaceholder}
          role="combobox"
          aria-controls="search-results"
          aria-haspopup="listbox"
          aria-expanded={Boolean(searchQuery.trim() && searchResults.length)}
          aria-activedescendant={searchResults.length ? `search-result-${activeSearchIndex}` : undefined}
          oninput={() => (activeSearchIndex = 0)}
        />
      </label>

      <p class="search-count" aria-live="polite">
        {searchQuery.trim() ? text.searchCount(searchResults.length) : ""}
      </p>

      <div
        id="search-results"
        class="search-results"
        role={searchQuery.trim() && searchResults.length ? "listbox" : undefined}
        aria-label={searchQuery.trim() && searchResults.length ? text.searchTitle : undefined}
      >
        {#if searchQuery.trim() && searchResults.length}
          {#each searchResults as result, index (result.date)}
            <button
              type="button"
              id="search-result-{index}"
              class="search-result"
              class:active={activeSearchIndex === index}
              role="option"
              aria-selected={activeSearchIndex === index}
              aria-label="{result.date}: {result.content}"
              onclick={() => selectSearchResult(result.date)}
            >
              <time class="search-date" datetime={result.date}>{result.date}</time>
              <span class="search-content">
                {#each highlightMatches(result.content, searchQuery) as segment}
                  {#if segment.highlighted}
                    <mark>{segment.text}</mark>
                  {:else}
                    {segment.text}
                  {/if}
                {/each}
              </span>
            </button>
          {/each}
        {:else if searchQuery.trim()}
          <p class="search-empty">{text.searchEmpty}</p>
        {:else}
          <p class="search-empty">{text.searchPrompt}</p>
        {/if}
      </div>

      <p class="search-help">
        <span><kbd>↑</kbd><kbd>↓</kbd> {text.searchHelpNavigate}</span>
        <span><kbd>Esc</kbd> {text.searchHelpClose}</span>
      </p>
    </div>
  </div>
{/if}

{#if status}
  <div class="status" role="status">{status}</div>
{/if}

{#if editingKey}
  <!-- 蒙版只做视觉遮罩，点击不关闭；关闭走「取消」按钮或 Esc，避免误触丢失正在编辑的内容 -->
  <div class="backdrop">
    <div
      class="editor"
      role="dialog"
      aria-modal="true"
      aria-label={text.editorTitle}
      tabindex="-1"
    >
      <div class="editor-head">
        <span class="editor-date">{editingKey}</span>
        <span class="editor-hint">{text.hint}</span>
      </div>
      <textarea
        bind:this={editorInput}
        bind:value={draft}
        placeholder={text.placeholder}
        onkeydown={handleEditorKeydown}
      ></textarea>
      <div class="editor-actions">
        {#if notes[editingKey]}
          <button type="button" class="danger" onclick={removeNote}
            >{text.clear}</button
          >
        {/if}
        <span class="spacer"></span>
        <button type="button" class="ghost" onclick={closeEditor}
          >{text.cancel}</button
        >
        <button type="button" class="primary" onclick={saveNote}
          >{text.save}</button
        >
      </div>
    </div>
  </div>
{/if}

{#if exportOpen}
  <div class="backdrop">
    <div class="dialog" role="dialog" aria-modal="true" aria-label={text.export} tabindex="-1">
      <div class="types" role="radiogroup" aria-label={text.export}>
        <button
          type="button"
          class="type"
          class:active={exportType === "xls"}
          role="radio"
          aria-checked={exportType === "xls"}
          onclick={() => setExportType("xls")}
        >
          <img src={xlsIcon} alt="" />
          <span>{text.fileExcel}</span>
        </button>
        <button
          type="button"
          class="type"
          class:active={exportType === "txt"}
          role="radio"
          aria-checked={exportType === "txt"}
          onclick={() => setExportType("txt")}
        >
          <img src={txtIcon} alt="" />
          <span>{text.fileText}</span>
        </button>
      </div>

      <div class="destination">
        <span class="destination-label">{text.outputLabel}</span>
        <p class="destination-path" title={exportPath}>{exportPath}</p>
        <div class="destination-actions">
          <button type="button" class="ghost" onclick={chooseExportPath}>{text.choosePath}</button>
          <button type="button" class="ghost" onclick={copyExportPath}>{text.copyPath}</button>
          <button type="button" class="ghost" onclick={revealExportPath}>{text.openFolder}</button>
        </div>
      </div>

      <div class="range-inputs">
        <input
          type="date"
          aria-label={text.rangeLabel}
          bind:value={exportFrom}
          disabled={exportRange !== "custom"}
        />
        <span class="dash">–</span>
        <input
          type="date"
          aria-label={text.rangeLabel}
          bind:value={exportTo}
          disabled={exportRange !== "custom"}
        />
      </div>

      <div class="ranges" role="radiogroup" aria-label={text.rangeLabel}>
        {#each Object.entries(text.ranges) as [key, label] (key)}
          <label class:active={exportRange === key}>
            <input
              type="radio"
              name="export-range"
              value={key}
              checked={exportRange === key}
              onchange={() => pickRange(key)}
            />
            {label}
          </label>
        {/each}
      </div>

      <div class="dialog-actions">
        <span class="count">
          {#if exportRows.length}
            {text.entries(exportRows.length)}
          {:else if notesError}
            {text.nothingInRange} — {text.saveFailed}: {notesError}
          {:else}
            {text.nothingInRange}
          {/if}
        </span>
        <button type="button" class="ghost" onclick={() => (exportOpen = false)}>{text.cancel}</button>
        <button type="button" class="primary" disabled={!exportRows.length} onclick={runExport}>
          {text.export}
        </button>
      </div>
    </div>
  </div>
{/if}

{#if importOpen}
  <div class="backdrop">
    <div class="dialog" role="dialog" aria-modal="true" aria-label={text.import} tabindex="-1">
      <div class="import-icon">
        <img src={xlsIcon} alt="" />
      </div>

      <button type="button" class="file-button" onclick={() => fileInput?.click()}>
        {importFileName || text.chooseFile}
      </button>

      {#if importError}
        <p class="error">{importError}</p>
      {:else if importFileName}
        <p class="summary">
          {text.entries(importRows.length)}{importConflicts
            ? ` · ${text.conflicts(importConflicts)}`
            : ""}
        </p>

        {#if importConflicts}
          <div class="modes" role="radiogroup" aria-label={text.conflicts(importConflicts)}>
            {#each [["overwrite", text.modeOverwrite], ["merge", text.modeMerge], ["skip", text.modeSkip]] as [key, label] (key)}
              <label class:active={importMode === key}>
                <input
                  type="radio"
                  name="import-mode"
                  value={key}
                  checked={importMode === key}
                  onchange={() => (importMode = key)}
                />
                {label}
              </label>
            {/each}
          </div>
        {/if}
      {/if}

      <div class="dialog-actions">
        <button type="button" class="ghost" onclick={closeImport}>{text.cancel}</button>
        <button type="button" class="primary" disabled={!importRows.length} onclick={runImport}>
          {text.import}
        </button>
      </div>
    </div>
  </div>
{/if}

{#if settingsOpen}
  <!-- A settings sheet reads as one surface: an index rail on the left, one
       pane of content on the right. The rail is not a nav menu that navigates
       away, it is a table of contents for the same panel, so the two share the
       dialog's single frame instead of splitting into floating cards. -->
  <div class="backdrop">
    <div
      class="settings"
      role="dialog"
      aria-modal="true"
      aria-label={text.settings}
      tabindex="-1"
      bind:this={settingsPanel}
    >
      <div class="settings-head">
        <h2>{text.settings}</h2>
        <button
          type="button"
          class="settings-close"
          aria-label={text.settingsClose}
          onclick={closeSettings}
        >
          <svg viewBox="0 0 16 16" aria-hidden="true" focusable="false">
            <path
              d="m4 4 8 8m0-8-8 8"
              fill="none"
              stroke="currentColor"
              stroke-width="1.5"
              stroke-linecap="round"
            />
          </svg>
        </button>
      </div>

      <div class="settings-body">
        <nav class="settings-rail" aria-label={text.settings}>
          {#each settingsTabs as tab (tab.id)}
            <button
              type="button"
              class="settings-tab"
              class:active={settingsTab === tab.id}
              aria-current={settingsTab === tab.id ? "page" : undefined}
              onclick={() => (settingsTab = tab.id)}
            >
              <span class="settings-tab-icon" style={maskVar(tab.icon)} aria-hidden="true"></span>
              {tab.label}
            </button>
          {/each}
        </nav>

        <section class="settings-pane" aria-live="polite">
          {#if settingsTab === "grid"}
            <div class="settings-groups">
              <section class="settings-group">
                <h3 class="group-head">{text.groupStyle}</h3>
                <div class="group-body">
                  <!-- A pair of small grids rather than a dropdown: the choice is
                       spatial, so it reads faster shown than named. -->
                  <div class="style-choices">
                    <label class="style-choice" class:active={calendarStyle === "full"}>
                      <input
                        type="radio"
                        name="calendar-style"
                        value="full"
                        class="style-radio"
                        checked={calendarStyle === "full"}
                        onchange={() => (calendarStyle = "full")}
                      />
                      <span class="style-preview" aria-hidden="true">
                        {#each Array(28) as _, index (index)}
                          <span class:blank={index < 3}></span>
                        {/each}
                      </span>
                      <span class="style-name">{text.styleFull}</span>
                      <span class="style-note">{text.styleFullNote}</span>
                    </label>
                    <label class="style-choice" class:active={calendarStyle === "float"}>
                      <input
                        type="radio"
                        name="calendar-style"
                        value="float"
                        class="style-radio"
                        checked={calendarStyle === "float"}
                        onchange={() => (calendarStyle = "float")}
                      />
                      <span class="style-preview" aria-hidden="true">
                        {#each Array(28) as _, index (index)}
                          <span class:anchor={index === 9}></span>
                        {/each}
                      </span>
                      <span class="style-name">{text.styleFloat}</span>
                      <span class="style-note">{text.styleFloatNote}</span>
                    </label>
                  </div>

                  {#if calendarStyle === "float"}
                    <div class="setting-row">
                      <span class="setting-copy">
                        <span class="setting-label">{text.pastWeeks}</span>
                        <span class="setting-hint">{text.pastWeeksHint}</span>
                      </span>
                      <select
                        class="setting-select"
                        aria-label={text.pastWeeks}
                        value={String(pastWeeks)}
                        onchange={(event) => (pastWeeks = Number(event.currentTarget.value))}
                      >
                        {#each [0, 1, 2, 3] as count (count)}
                          <option value={String(count)}>{text.pastWeeksValue(count)}</option>
                        {/each}
                      </select>
                    </div>
                  {/if}
                </div>
              </section>

              <section class="settings-group">
                <h3 class="group-head">{text.groupCellStyle}</h3>
                <div class="group-body">
                  <!-- The three marks are compared, not chosen from a list, so
                       they sit side by side like the calendar styles above. -->
                  <div class="bullet-choices">
                    {#each [
                      { id: "none", label: text.bulletNone },
                      { id: "number", label: text.bulletNumber },
                      { id: "dot", label: text.bulletDot },
                    ] as option (option.id)}
                      <label class="style-choice" class:active={cellBullet === option.id}>
                        <input
                          type="radio"
                          name="cell-bullet"
                          value={option.id}
                          class="style-radio"
                          checked={cellBullet === option.id}
                          onchange={() => (cellBullet = option.id)}
                        />
                        <span class="bullet-preview" aria-hidden="true">
                          {#each [1, 2, 3] as item (item)}
                            <span class="bullet-line">
                              <span class="bullet-mark"
                                >{option.id === "number"
                                  ? item
                                  : option.id === "dot"
                                    ? "•"
                                    : ""}</span
                              >
                              <span class="bullet-bar"></span>
                            </span>
                          {/each}
                        </span>
                        <span class="style-name">{option.label}</span>
                      </label>
                    {/each}
                  </div>
                  <p class="group-note">{text.bulletHint}</p>
                </div>
              </section>

              {#each [[text.groupDisplay, displaySettings], [text.groupFestival, festivalSettings], [text.groupLunar, lunarSettings]] as [title, rows] (title)}
                <section class="settings-group">
                  <h3 class="group-head">{title}</h3>
                  <div class="group-body">
                    {#each rows as row (row.key)}
                      <label class="setting-row">
                        <span class="setting-copy">
                          <span class="setting-label">{row.label}</span>
                          <span class="setting-hint">{row.hint}</span>
                        </span>
                        <input
                          type="checkbox"
                          class="switch"
                          checked={row.value}
                          onchange={(event) => row.set(event.currentTarget.checked)}
                        />
                      </label>
                    {/each}
                  </div>
                </section>
              {/each}
            </div>
          {:else}
            <div class="about">
              <header class="about-identity">
                <img class="about-mark" src={productIcon} alt="" width="56" height="56" />
                <div class="about-title">
                  <h3>{text.aboutName}</h3>
                  <p class="about-tagline">{text.aboutTagline}</p>
                </div>
                <span class="about-version">v{manifest.version}</span>
              </header>

              <p class="about-intro">{text.aboutIntro}</p>

              <ul class="about-features">
                {#each text.aboutFeatures as feature (feature)}
                  <li>{feature}</li>
                {/each}
              </ul>

              <h4 class="about-meta-title">{text.aboutMeta}</h4>
              <dl class="about-meta">
                <div>
                  <dt>{text.metaPublisher}</dt>
                  <dd>{manifest.publisher}</dd>
                </div>
                <div>
                  <dt>{text.metaLicense}</dt>
                  <dd>Apache-2.0</dd>
                </div>
                <div>
                  <dt>{text.metaRequires}</dt>
                  <dd>DBX ≥ 0.6.0 · Host API 1</dd>
                </div>
                <div>
                  <dt>{text.metaBuilt}</dt>
                  <dd>Svelte 5 · Vite 7 · Go 1.22</dd>
                </div>
              </dl>

              <div class="about-links">
                <a href={manifest.homepage} target="_blank" rel="noreferrer noopener">
                  {text.metaHomepage}
                </a>
                <a href={manifest.source} target="_blank" rel="noreferrer noopener">
                  {text.metaSource}
                </a>
              </div>

              <p class="about-rights">{text.aboutRights(new Date().getFullYear())}</p>
            </div>
          {/if}
        </section>
      </div>
    </div>
  </div>
{/if}

<style>
  :global(*) {
    box-sizing: border-box;
  }
  :global(html),
  :global(body),
  :global(#app) {
    height: 100%;
  }
  :global(body) {
    --accent: #6d5dfc;
    --select: #e0a12e;
    --rule: color-mix(in srgb, CanvasText 14%, transparent);
    --muted: color-mix(in srgb, CanvasText 45%, transparent);
    --tint: color-mix(in srgb, CanvasText 4%, transparent);
    margin: 0;
    color: CanvasText;
    background: Canvas;
    font-family: ui-sans-serif, system-ui, "Segoe UI", "PingFang SC",
      "Microsoft YaHei", sans-serif;
    font-variant-numeric: tabular-nums;
    -webkit-font-smoothing: antialiased;
  }

  main {
    display: flex;
    flex-direction: column;
    gap: 10px;
    height: 100%;
    padding: clamp(12px, 2vw, 20px);
  }

  /* Toolbar stays quiet: the grid is the interface. */
  .toolbar {
    position: relative;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 8px 16px;
    padding: 0 2px;
  }
  h1 {
    margin: 0;
    font-size: 19px;
    font-weight: 600;
    line-height: 1.2;
    letter-spacing: -0.01em;
  }
  /* The title is the month picker's handle: a quiet label that reacts to the
     pointer, not a button-shaped control competing with the grid. */
  .month-menu {
    position: relative;
    min-width: 0;
  }
  .month-trigger {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    margin: -4px 0 -4px -7px;
    padding: 4px 7px;
    border: 0;
    border-radius: 7px;
    color: inherit;
    background: transparent;
    font: inherit;
    letter-spacing: inherit;
    cursor: pointer;
  }
  .month-trigger > span {
    overflow: hidden;
    white-space: nowrap;
    text-overflow: ellipsis;
  }
  .month-trigger:hover {
    background: var(--tint);
  }
  .month-trigger:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: 2px;
  }
  .month-trigger .chevron {
    flex: none;
    width: 12px;
    height: 12px;
    color: var(--muted);
    transition: transform 140ms ease;
  }
  .month-trigger[aria-expanded="true"] .chevron {
    transform: rotate(180deg);
  }

  /* A compact almanac index: year stepper on top, twelve months below. */
  .month-picker {
    position: absolute;
    top: calc(100% + 8px);
    left: 2px;
    z-index: 12;
    width: min(292px, calc(100vw - 32px));
    padding: 12px;
    border: 1px solid var(--rule);
    border-radius: 12px;
    background: Canvas;
    box-shadow: 0 16px 36px rgba(0, 0, 0, 0.2);
    animation: picker-in 120ms ease-out;
  }
  @keyframes picker-in {
    from {
      opacity: 0;
      transform: translateY(-4px);
    }
  }
  .picker-year {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    padding-bottom: 10px;
    border-bottom: 1px solid var(--rule);
  }
  .picker-step {
    display: grid;
    place-items: center;
    width: 26px;
    height: 26px;
    padding: 0;
    border: 0;
    border-radius: 6px;
    color: var(--muted);
    background: transparent;
    cursor: pointer;
  }
  .picker-step:hover {
    color: inherit;
    background: var(--tint);
  }
  .picker-step svg {
    width: 12px;
    height: 12px;
  }
  .picker-year-field {
    display: flex;
    align-items: baseline;
    gap: 3px;
  }
  .picker-year-input {
    width: 4.6ch;
    padding: 2px 2px 1px;
    border: 0;
    border-bottom: 1px dashed color-mix(in srgb, CanvasText 28%, transparent);
    border-radius: 4px 4px 0 0;
    color: inherit;
    background: transparent;
    font: inherit;
    font-size: 14px;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
    text-align: center;
  }
  .picker-year-input:hover {
    background: var(--tint);
  }
  .picker-year-input:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: 1px;
  }
  .picker-year-suffix {
    color: var(--muted);
    font-size: 12px;
  }
  .picker-months {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 4px;
    padding-top: 10px;
  }
  .picker-month {
    height: 34px;
    padding: 0;
    border: 0;
    border-radius: 7px;
    color: inherit;
    background: transparent;
    font: inherit;
    font-size: 12.5px;
    cursor: pointer;
  }
  .picker-month:hover {
    background: var(--tint);
  }
  .picker-month.active {
    color: #fff;
    background: var(--accent);
    font-weight: 600;
  }
  .picker-month.active:focus-visible {
    outline-color: CanvasText;
    outline-offset: 2px;
  }
  .picker-month:focus-visible,
  .picker-step:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: 1px;
  }
  .nav {
    display: flex;
    flex: none;
    gap: 4px;
  }
  .nav button {
    min-width: 32px;
    height: 30px;
    padding: 0 10px;
    border: 1px solid var(--rule);
    border-radius: 7px;
    color: inherit;
    background: transparent;
    font: inherit;
    font-size: 14px;
    line-height: 1;
    white-space: nowrap;
    cursor: pointer;
  }
  .nav button:hover {
    background: var(--tint);
  }
  /* The arrow SVGs are filled with a fixed dark navy, which would disappear on
     a dark theme as a plain <img>. Masking them instead paints the shape with
     currentColor, so they follow the host theme. Drawn on a pseudo-element so
     the button's own hover background stays independent of the icon. */
  .nav .arrow {
    display: grid;
    place-items: center;
    padding: 0;
  }
  .nav .arrow::before {
    content: "";
    width: 13px;
    height: 13px;
    background-color: currentColor;
    -webkit-mask: var(--icon) center / contain no-repeat;
    mask: var(--icon) center / contain no-repeat;
  }
  .nav .today {
    font-size: 12.5px;
  }

  .actions {
    display: flex;
    flex: none;
    align-items: center;
    gap: 8px;
    margin-left: auto;
  }

  .notice {
    margin: 0;
    padding: 7px 10px;
    border: 1px solid color-mix(in srgb, #e5484d 40%, transparent);
    border-radius: 8px;
    color: color-mix(in srgb, #e5484d 80%, CanvasText);
    font-size: 12px;
    line-height: 1.5;
  }

  .menu {
    position: relative;
  }
  .menu-button {
    display: grid;
    place-items: center;
    width: 30px;
    height: 30px;
    padding: 0;
    border: 1px solid var(--rule);
    border-radius: 7px;
    color: inherit;
    background: transparent;
    cursor: pointer;
  }
  .menu-button:hover {
    background: var(--tint);
  }
  .menu-button svg {
    width: 15px;
    height: 15px;
  }

  /* Anchored to the button's right edge so it opens leftward, away from the
     panel edge. */
  .dropdown {
    position: absolute;
    top: calc(100% + 6px);
    right: 0;
    z-index: 10;
    display: flex;
    flex-direction: column;
    min-width: 150px;
    padding: 4px;
    border: 1px solid var(--rule);
    border-radius: 10px;
    background: Canvas;
    box-shadow: 0 12px 30px rgba(0, 0, 0, 0.18);
  }
  .dropdown button {
    display: grid;
    grid-template-columns: 15px minmax(0, 1fr);
    align-items: center;
    gap: 9px;
    padding: 7px 10px;
    border: 0;
    border-radius: 6px;
    color: inherit;
    background: transparent;
    font: inherit;
    font-size: 13px;
    text-align: left;
    white-space: nowrap;
    cursor: pointer;
  }
  /* Masked so the glyph inherits the row's colour and its hover/active states
     come for free; nothing has to keep a second palette in sync. */
  .dropdown-icon {
    width: 15px;
    height: 15px;
    background-color: color-mix(in srgb, CanvasText 62%, transparent);
    -webkit-mask: var(--icon) center / contain no-repeat;
    mask: var(--icon) center / contain no-repeat;
  }
  .dropdown button:hover {
    background: var(--tint);
  }
  .dropdown button:hover .dropdown-icon {
    background-color: CanvasText;
  }

  /* Floating so that reporting a result never shifts the grid. */
  .status {
    position: fixed;
    bottom: 18px;
    left: 50%;
    z-index: 20;
    padding: 8px 14px;
    transform: translateX(-50%);
    border-radius: 8px;
    color: Canvas;
    background: color-mix(in srgb, CanvasText 85%, Canvas);
    font-size: 12.5px;
  }

  .grid {
    display: grid;
    flex: 1;
    min-height: 0;
    grid-template-columns: repeat(7, minmax(0, 1fr));
    gap: 1px;
    overflow: hidden;
    border: 1px solid var(--rule);
    border-radius: 10px;
    background: var(--rule);
  }

  .weekday {
    padding: 7px 10px;
    background: Canvas;
    color: var(--muted);
    font-size: 11.5px;
    font-weight: 500;
  }

  .cell {
    display: flex;
    flex-direction: column;
    gap: 3px;
    min-height: 0;
    padding: 7px 9px;
    overflow: hidden;
    border: 0;
    border-radius: 0;
    background: Canvas;
    color: inherit;
    font: inherit;
    text-align: left;
    cursor: pointer;
    /* Double-click is the edit gesture, so suppress the text selection it would otherwise make. */
    user-select: none;
  }
  .cell:hover {
    background: color-mix(in srgb, var(--accent) 6%, Canvas);
  }
  .cell:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: -2px;
  }
  /* Days from the neighbouring months recede so the month reads as one block. */
  .cell.outside {
    color: color-mix(in srgb, CanvasText 35%, transparent);
  }
  .cell.weekend {
    background: color-mix(in srgb, var(--accent) 5%, Canvas);
  }
  .cell.weekend:hover {
    background: color-mix(in srgb, var(--accent) 10%, Canvas);
  }
  /* Today is the one cell allowed to break the hairline grid, so it carries a
     solid ring plus a tint. Drawn inset because the 1px gaps leave no room for
     a real border without shifting the row, and left square-cornered because
     rounding would expose the rule colour behind the cell's own background. */
  .cell.today {
    background: color-mix(in srgb, var(--accent) 11%, Canvas);
    box-shadow: inset 0 0 0 2px var(--accent);
  }
  .cell.today:hover {
    background: color-mix(in srgb, var(--accent) 16%, Canvas);
  }
  /* A search result is framed in amber, the way a date gets marked in a paper
     diary. The frame does the work, so the date and note keep their normal ink;
     today keeps its purple date pill but yields the frame to the selection. */
  .cell.selected {
    background: color-mix(in srgb, var(--select) 7%, Canvas);
    box-shadow: inset 0 0 0 2px var(--select);
  }
  .cell.selected:hover {
    background: color-mix(in srgb, var(--select) 11%, Canvas);
  }
  .cell.today.selected {
    background: color-mix(in srgb, var(--select) 8%, Canvas);
    box-shadow: inset 0 0 0 2px var(--select);
  }

  .cell-head {
    display: flex;
    align-items: center;
    gap: 4px;
    min-width: 0;
  }

  .date {
    display: grid;
    place-items: center;
    min-width: 20px;
    height: 20px;
    padding: 0 4px;
    border-radius: 5px;
    font-size: 12px;
    font-weight: 500;
  }

  /* "十一月" instead of a bare 1: it is wider than a two-digit day, so it
     loses the pill's fixed width and drops a size to sit beside the rest. */
  .date.month-day {
    min-width: 0;
    height: 18px;
    padding: 0 5px;
    font-size: 11px;
    white-space: nowrap;
  }

  /* Lunar date, solar term or festival name, sitting right of the day number. */
  .lunar {
    flex: 1;
    min-width: 0;
    overflow: hidden;
    color: var(--muted);
    font-size: 11px;
    white-space: nowrap;
    text-overflow: ellipsis;
  }
  .cell.outside .lunar {
    color: color-mix(in srgb, CanvasText 25%, transparent);
  }

  /* 休 / 班 marker. Rendered only for years with a published holiday schedule. */
  .badge {
    flex: none;
    display: grid;
    place-items: center;
    min-width: 15px;
    height: 15px;
    padding: 0 3px;
    border-radius: 4px;
    color: CanvasText;
    background: color-mix(in srgb, CanvasText 16%, Canvas);
    font-size: 10px;
    font-weight: 600;
  }
  .badge.off {
    color: #fff;
    background: #e5484d;
  }
  .cell.today .date {
    background: var(--accent);
    color: #fff;
  }

  .note {
    overflow: hidden;
    font-size: 12px;
    line-height: 1.45;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
    opacity: 0.78;
    display: -webkit-box;
    -webkit-line-clamp: 4;
    line-clamp: 4;
    -webkit-box-orient: vertical;
  }

  .backdrop {
    position: fixed;
    inset: 0;
    display: grid;
    place-items: center;
    padding: 20px;
    background: color-mix(in srgb, #000 45%, transparent);
  }
  .search-backdrop {
    z-index: 30;
  }
  .editor {
    display: flex;
    flex-direction: column;
    gap: 12px;
    width: min(460px, 100%);
    padding: 16px;
    border: 1px solid var(--rule);
    border-radius: 14px;
    background: Canvas;
    box-shadow: 0 18px 48px rgba(0, 0, 0, 0.3);
  }
  .editor-head {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 12px;
  }
  .editor-date {
    font-size: 14px;
    font-weight: 600;
  }
  .editor-hint {
    color: var(--muted);
    font-size: 11.5px;
  }
  textarea {
    min-height: 150px;
    padding: 10px 12px;
    resize: vertical;
    border: 1px solid var(--rule);
    border-radius: 9px;
    color: inherit;
    background: color-mix(in srgb, CanvasText 3%, Canvas);
    font: inherit;
    font-size: 13.5px;
    line-height: 1.55;
  }
  textarea:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: -1px;
  }
  .editor-actions,
  .dialog-actions {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  /* Keeps cancel/confirm on the right in both dialogs. The export dialog has a
     leading count, but the import one does not, so it needs this explicitly. */
  .dialog-actions {
    justify-content: flex-end;
  }
  .spacer {
    flex: 1;
  }
  .editor-actions button,
  .dialog-actions button {
    height: 32px;
    padding: 0 14px;
    border: 1px solid transparent;
    border-radius: 8px;
    font: inherit;
    font-size: 13px;
    cursor: pointer;
  }
  .editor-actions .primary,
  .dialog-actions .primary {
    color: #fff;
    background: var(--accent);
  }
  .dialog-actions .primary:disabled {
    opacity: 0.45;
    cursor: not-allowed;
  }
  .editor-actions .ghost,
  .editor-actions .danger,
  .dialog-actions .ghost {
    border-color: var(--rule);
    color: inherit;
    background: transparent;
  }
  .editor-actions .ghost:hover,
  .editor-actions .danger:hover,
  .dialog-actions .ghost:hover {
    background: var(--tint);
  }
  .editor-actions .danger {
    color: color-mix(in srgb, #e5484d 80%, CanvasText);
  }

  /* Search dialog */
  .search-dialog {
    display: flex;
    flex-direction: column;
    gap: 12px;
    width: min(560px, 100%);
    height: min(560px, 84vh);
    padding: 16px;
    overflow: hidden;
    border: 1px solid var(--rule);
    border-radius: 14px;
    background: Canvas;
    box-shadow: 0 18px 48px rgba(0, 0, 0, 0.3);
  }
  .search-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
  }
  .search-head h2 {
    margin: 0;
    font-size: 15px;
    font-weight: 600;
  }
  .search-close {
    display: grid;
    flex: none;
    place-items: center;
    width: 30px;
    height: 30px;
    padding: 0;
    border: 1px solid var(--rule);
    border-radius: 7px;
    color: inherit;
    background: transparent;
    cursor: pointer;
  }
  .search-close:hover {
    background: var(--tint);
  }
  .search-close svg {
    width: 14px;
    height: 14px;
  }
  .search-field {
    display: flex;
    align-items: center;
    gap: 9px;
    min-height: 38px;
    padding: 0 10px;
    border: 1px solid var(--rule);
    border-radius: 9px;
    background: color-mix(in srgb, CanvasText 3%, Canvas);
  }
  .search-field:focus-within {
    outline: 2px solid var(--accent);
    outline-offset: -1px;
  }
  .search-field svg {
    flex: none;
    width: 15px;
    height: 15px;
    color: var(--muted);
  }
  .search-field input {
    width: 100%;
    min-width: 0;
    height: 36px;
    padding: 0;
    border: 0;
    outline: 0;
    color: inherit;
    background: transparent;
    font: inherit;
    font-size: 13px;
  }
  .search-field input::placeholder {
    color: var(--muted);
  }
  .search-count {
    min-height: 16px;
    margin: 0;
    color: var(--muted);
    font-size: 11.5px;
  }
  .search-results {
    display: flex;
    flex: 1;
    flex-direction: column;
    gap: 3px;
    min-height: 0;
    overflow: auto;
    overscroll-behavior: contain;
  }
  .search-result {
    display: flex;
    align-items: flex-start;
    gap: 12px;
    width: 100%;
    padding: 10px;
    border: 1px solid transparent;
    border-radius: 8px;
    color: inherit;
    background: transparent;
    font: inherit;
    text-align: left;
    cursor: pointer;
  }
  .search-result:hover {
    border-color: var(--rule);
    background: var(--tint);
  }
  .search-result.active {
    border-color: color-mix(in srgb, var(--accent) 48%, transparent);
    background: color-mix(in srgb, var(--accent) 9%, Canvas);
  }
  .search-result:focus-visible,
  .search-close:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: 2px;
  }
  /* The search trigger is an icon-only shortcut target that focus returns to
     after Escape, so it stays frameless in every state; its hover tint is the
     only affordance it needs. */
  .search-button:focus,
  .search-button:focus-visible {
    outline: none;
  }
  /* Pointer clicks must not leave a focus frame behind when Escape switches
     the browser back to keyboard mode; Tab and other keyboard navigation set
     the class that keeps the normal focus treatment. */
  .toolbar:not(.keyboard-focus) button:focus {
    outline: none;
  }
  .search-date {
    flex: none;
    min-width: 92px;
    padding-top: 1px;
    color: var(--muted);
    font-size: 11.5px;
    font-variant-numeric: tabular-nums;
  }
  .search-content {
    min-width: 0;
    font-size: 12.5px;
    line-height: 1.5;
    overflow-wrap: anywhere;
    white-space: pre-wrap;
  }
  .search-content mark {
    border-radius: 2px;
    color: inherit;
    background: color-mix(in srgb, var(--accent) 28%, transparent);
    font-weight: 650;
  }
  .search-empty {
    display: grid;
    flex: 1;
    place-items: center;
    margin: 0;
    padding: 18px;
    color: var(--muted);
    font-size: 13px;
    text-align: center;
  }
  .search-help {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 6px 16px;
    margin: 0;
    padding-top: 10px;
    border-top: 1px solid var(--rule);
    color: var(--muted);
    font-size: 11.5px;
  }
  .search-help span {
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }
  .search-help kbd {
    display: inline-grid;
    place-items: center;
    min-width: 20px;
    height: 20px;
    padding: 0 5px;
    border: 1px solid var(--rule);
    border-radius: 5px;
    background: color-mix(in srgb, CanvasText 4%, Canvas);
    color: inherit;
    font: inherit;
    font-size: 10.5px;
    line-height: 1;
  }

  /* Export dialog */
  .dialog {
    display: flex;
    flex-direction: column;
    gap: 14px;
    width: min(420px, 100%);
    padding: 16px;
    border: 1px solid var(--rule);
    border-radius: 14px;
    background: Canvas;
    box-shadow: 0 18px 48px rgba(0, 0, 0, 0.3);
  }

  .types {
    display: flex;
    gap: 10px;
  }
  .type {
    display: flex;
    flex: 1;
    flex-direction: column;
    align-items: center;
    gap: 6px;
    padding: 12px 8px;
    border: 1px solid var(--rule);
    border-radius: 10px;
    color: inherit;
    background: transparent;
    font: inherit;
    font-size: 12.5px;
    cursor: pointer;
  }
  .type img {
    width: 32px;
    height: 32px;
  }
  .type:hover {
    background: var(--tint);
  }
  .type.active {
    border-color: var(--accent);
    background: color-mix(in srgb, var(--accent) 8%, Canvas);
  }

  /* The destination is the one piece of the dialog the user may need to read
     character by character, so the path wraps rather than truncating. */
  .destination {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .destination-label {
    color: var(--muted);
    font-size: 11.5px;
  }
  .destination-path {
    margin: 0;
    padding: 7px 9px;
    border: 1px solid var(--rule);
    border-radius: 8px;
    background: color-mix(in srgb, CanvasText 3%, Canvas);
    font-size: 11.5px;
    line-height: 1.5;
    overflow-wrap: anywhere;
  }
  .destination-actions {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
  }
  .destination-actions button {
    height: 28px;
    padding: 0 10px;
    border: 1px solid var(--rule);
    border-radius: 7px;
    color: inherit;
    background: transparent;
    font: inherit;
    font-size: 12px;
    cursor: pointer;
  }
  .destination-actions button:hover {
    background: var(--tint);
  }

  .range-inputs {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .range-inputs input {
    flex: 1;
    min-width: 0;
    height: 32px;
    padding: 0 8px;
    border: 1px solid var(--rule);
    border-radius: 8px;
    color: inherit;
    background: color-mix(in srgb, CanvasText 3%, Canvas);
    font: inherit;
    font-size: 12.5px;
  }
  .range-inputs input:disabled {
    opacity: 0.55;
  }
  .dash {
    color: var(--muted);
  }

  .ranges,
  .modes {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
  }
  .ranges label,
  .modes label {
    display: flex;
    align-items: center;
    gap: 5px;
    padding: 5px 11px;
    border: 1px solid var(--rule);
    border-radius: 20px;
    font-size: 12.5px;
    cursor: pointer;
  }
  .ranges label:hover,
  .modes label:hover {
    background: var(--tint);
  }
  .ranges label.active,
  .modes label.active {
    border-color: var(--accent);
    background: color-mix(in srgb, var(--accent) 8%, Canvas);
  }
  .ranges input,
  .modes input {
    margin: 0;
    accent-color: var(--accent);
  }

  /* Import dialog */
  .import-icon {
    display: grid;
    place-items: center;
  }
  .import-icon img {
    width: 44px;
    height: 44px;
  }
  .file-button {
    height: 34px;
    padding: 0 12px;
    overflow: hidden;
    border: 1px solid var(--rule);
    border-radius: 8px;
    color: inherit;
    background: transparent;
    font: inherit;
    font-size: 12.5px;
    white-space: nowrap;
    text-overflow: ellipsis;
    cursor: pointer;
  }
  .file-button:hover {
    background: var(--tint);
  }
  .summary,
  .error {
    margin: 0;
    font-size: 12.5px;
  }
  .summary {
    color: var(--muted);
  }
  .error {
    color: color-mix(in srgb, #e5484d 80%, CanvasText);
  }

  .count {
    margin-right: auto;
    color: var(--muted);
    font-size: 12px;
  }

  /* Settings sheet. One panel with an index rail rather than a stack of cards:
     the hairline between rail and pane is the same rule the calendar grid uses,
     so the dialog speaks the product's language instead of a generic one. */
  .settings {
    display: flex;
    flex-direction: column;
    width: min(720px, 100%);
    height: min(560px, 86vh);
    overflow: hidden;
    border: 1px solid var(--rule);
    border-radius: 14px;
    background: Canvas;
    box-shadow: 0 18px 48px rgba(0, 0, 0, 0.3);
  }
  .settings:focus {
    outline: none;
  }
  .settings-head {
    display: flex;
    flex: none;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    padding: 13px 16px;
    border-bottom: 1px solid var(--rule);
  }
  .settings-head h2 {
    margin: 0;
    font-size: 15px;
    font-weight: 600;
  }
  .settings-close {
    display: grid;
    flex: none;
    place-items: center;
    width: 30px;
    height: 30px;
    padding: 0;
    border: 1px solid var(--rule);
    border-radius: 7px;
    color: inherit;
    background: transparent;
    cursor: pointer;
  }
  .settings-close:hover {
    background: var(--tint);
  }
  .settings-close svg {
    width: 13px;
    height: 13px;
  }
  .settings-close:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: 2px;
  }

  .settings-body {
    display: grid;
    flex: 1;
    min-height: 0;
    grid-template-columns: 176px minmax(0, 1fr);
  }

  /* The rail is tinted a shade so the two columns read as index and page,
     the way a ledger separates tabs from the sheet they open. */
  .settings-rail {
    display: flex;
    flex-direction: column;
    gap: 2px;
    padding: 10px 8px;
    overflow-y: auto;
    border-right: 1px solid var(--rule);
    background: color-mix(in srgb, CanvasText 3.5%, Canvas);
  }
  .settings-tab {
    position: relative;
    display: grid;
    grid-template-columns: 14px minmax(0, 1fr);
    align-items: center;
    gap: 8px;
    padding: 8px 10px 8px 13px;
    border: 0;
    border-radius: 7px;
    color: color-mix(in srgb, CanvasText 70%, transparent);
    background: transparent;
    font: inherit;
    font-size: 12.5px;
    text-align: left;
    cursor: pointer;
  }
  .settings-tab-icon {
    width: 14px;
    height: 14px;
    background-color: color-mix(in srgb, CanvasText 45%, transparent);
    -webkit-mask: var(--icon) center / contain no-repeat;
    mask: var(--icon) center / contain no-repeat;
  }
  .settings-tab:hover {
    color: CanvasText;
    background: color-mix(in srgb, CanvasText 5%, transparent);
  }
  .settings-tab:hover .settings-tab-icon {
    background-color: color-mix(in srgb, CanvasText 75%, transparent);
  }
  /* The active marker is a short amber tick, the same accent a chosen date
     carries in the grid. */
  .settings-tab.active {
    color: CanvasText;
    background: Canvas;
    font-weight: 600;
    box-shadow: inset 0 0 0 1px var(--rule);
  }
  .settings-tab.active .settings-tab-icon {
    background-color: CanvasText;
  }
  .settings-tab.active::before {
    content: "";
    position: absolute;
    top: 50%;
    left: 5px;
    width: 2px;
    height: 14px;
    transform: translateY(-50%);
    border-radius: 2px;
    background: var(--select);
  }
  .settings-tab:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: -2px;
  }

  .settings-pane {
    min-width: 0;
    padding: 18px 20px 16px;
    overflow-y: auto;
    overscroll-behavior: contain;
  }

  /* Settings are grouped by the part of the calendar they change, and every
     group is a hairline ledger: label and explanation on the left, the control
     on the right. The rule between rows is the only divider, so the pane reads
     as one sheet rather than a stack of cards. */
  .settings-groups {
    display: flex;
    flex-direction: column;
    gap: 22px;
    max-width: 60ch;
  }
  /* A short amber tick before the title — the same accent a chosen date and
     the active rail entry carry — so each group reads as a marked heading
     rather than a floating label. */
  .group-head {
    position: relative;
    display: flex;
    align-items: center;
    gap: 10px;
    margin: 0 0 8px;
    padding-left: 10px;
    color: CanvasText;
    font-size: 13px;
    font-weight: 600;
    letter-spacing: 0.02em;
  }
  .group-head::before {
    content: "";
    position: absolute;
    top: 50%;
    left: 0;
    width: 2.5px;
    height: 13px;
    transform: translateY(-50%);
    border-radius: 2px;
    background: var(--select);
  }
  .group-head::after {
    content: "";
    flex: 1;
    height: 1px;
    background: var(--rule);
  }
  .group-body {
    display: flex;
    flex-direction: column;
  }
  .setting-row {
    display: flex;
    align-items: center;
    gap: 18px;
    padding: 9px 0;
    border-bottom: 1px solid var(--rule);
  }
  .group-body > .setting-row:last-child {
    padding-bottom: 0;
    border-bottom: 0;
  }
  .setting-copy {
    display: flex;
    flex: 1;
    flex-direction: column;
    gap: 1px;
    min-width: 0;
  }
  .setting-label {
    font-size: 12.5px;
  }
  .setting-hint {
    color: var(--muted);
    font-size: 11.5px;
    line-height: 1.45;
  }
  .setting-select {
    flex: none;
    min-width: 96px;
    padding: 5px 8px;
    border: 1px solid var(--rule);
    border-radius: 6px;
    color: inherit;
    background: Canvas;
    font: inherit;
    font-size: 12px;
    cursor: pointer;
  }
  .setting-select:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: 1px;
  }

  /* A pair of small grids rather than a dropdown: the choice is spatial, so it
     reads faster shown than named. */
  .style-choices {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 10px;
    margin-bottom: 10px;
  }
  .style-choice {
    position: relative;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 7px;
    padding: 12px 10px 10px;
    border: 1px solid var(--rule);
    border-radius: 9px;
    cursor: pointer;
  }
  .style-choice:hover {
    background: var(--tint);
  }
  .style-choice.active {
    border-color: var(--select);
    background: color-mix(in srgb, var(--select) 7%, transparent);
  }
  .style-choice:has(.style-radio:focus) {
    outline: 2px solid var(--accent);
    outline-offset: 1px;
  }
  .style-radio {
    position: absolute;
    width: 1px;
    height: 1px;
    margin: 0;
    opacity: 0;
    pointer-events: none;
  }
  .style-preview {
    display: grid;
    grid-template-columns: repeat(7, 1fr);
    gap: 2px;
    width: 84px;
  }
  .style-preview span {
    aspect-ratio: 1;
    border-radius: 1.5px;
    background: color-mix(in srgb, CanvasText 17%, transparent);
  }
  .style-preview span.blank {
    background: transparent;
  }
  .style-preview span.anchor {
    background: var(--select);
  }
  .style-name {
    font-size: 12.5px;
    font-weight: 600;
    line-height: 1.2;
  }
  .style-note {
    color: var(--muted);
    font-size: 11px;
    line-height: 1.4;
    text-align: center;
  }

  /* Three marks, same card as the calendar styles so the two choices read as
     siblings rather than two different widgets. Cards align their content to
     one edge: the three names are different lengths, and centring each would
     leave the row ragged. */
  .bullet-choices {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 8px;
  }
  .bullet-choices .style-choice {
    align-items: flex-start;
    text-align: left;
  }
  .bullet-choices .style-name {
    align-self: stretch;
    white-space: nowrap;
  }
  .bullet-preview {
    display: flex;
    flex-direction: column;
    gap: 5px;
    width: 100%;
    max-width: 64px;
    padding: 6px 0;
  }
  .bullet-line {
    display: flex;
    align-items: center;
    gap: 4px;
    /* A fixed row keeps the three previews the same height: the "none" mark
       is an empty string and would otherwise collapse its line. */
    height: 11px;
  }
  .bullet-mark {
    flex: none;
    width: 9px;
    color: var(--muted);
    font-size: 9px;
    line-height: 11px;
    text-align: right;
  }
  .bullet-bar {
    flex: 1;
    height: 4px;
    border-radius: 1px;
    background: color-mix(in srgb, CanvasText 17%, transparent);
  }
  .group-note {
    margin: 10px 0 0;
    color: var(--muted);
    font-size: 11.5px;
    line-height: 1.5;
  }

  .switch {
    position: relative;
    flex: none;
    width: 34px;
    height: 20px;
    margin: 0;
    padding: 0;
    border: 0;
    border-radius: 999px;
    background: color-mix(in srgb, CanvasText 20%, transparent);
    appearance: none;
    -webkit-appearance: none;
    cursor: pointer;
    transition: background 120ms ease;
  }
  .switch::after {
    content: "";
    position: absolute;
    top: 2px;
    left: 2px;
    width: 16px;
    height: 16px;
    border-radius: 50%;
    background: Canvas;
    box-shadow: 0 1px 2px rgb(0 0 0 / 0.25);
    transition: transform 120ms ease;
  }
  .switch:checked {
    background: var(--accent);
  }
  .switch:checked::after {
    transform: translateX(14px);
  }
  .switch:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: 2px;
  }

  .about {
    display: flex;
    flex-direction: column;
    gap: 12px;
    max-width: 60ch;
  }
  .about-identity {
    display: flex;
    align-items: center;
    gap: 12px;
  }
  .about-mark {
    flex: none;
    width: 44px;
    height: 44px;
  }
  .about-title {
    min-width: 0;
  }
  .about-title h3 {
    margin: 0;
    font-size: 16px;
    font-weight: 600;
    letter-spacing: -0.01em;
  }
  .about-tagline {
    margin: 2px 0 0;
    color: var(--muted);
    font-size: 12px;
    line-height: 1.5;
  }
  .about-version {
    flex: none;
    margin-left: auto;
    align-self: flex-start;
    padding: 2px 7px;
    border: 1px solid var(--rule);
    border-radius: 5px;
    color: var(--muted);
    font-size: 11px;
    font-variant-numeric: tabular-nums;
  }
  .about-intro {
    margin: 0;
    font-size: 12.5px;
    line-height: 1.65;
    color: color-mix(in srgb, CanvasText 82%, transparent);
  }

  /* Ruled rows, not floating pills: the list is a table of contents for what
     the plugin does, and hairlines are how this interface separates things. */
  .about-features {
    margin: 0;
    padding: 0;
    list-style: none;
    border-top: 1px solid var(--rule);
  }
  .about-features li {
    padding: 6px 0 6px 15px;
    border-bottom: 1px solid var(--rule);
    font-size: 12.5px;
    line-height: 1.5;
  }
  .about-features li::before {
    content: "·";
    float: left;
    width: 15px;
    margin-left: -15px;
    color: var(--select);
    font-weight: 700;
  }

  .about-meta-title {
    margin: 0;
    color: var(--muted);
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.04em;
    text-transform: uppercase;
  }
  .about-meta {
    display: grid;
    gap: 0;
    margin: -10px 0 0;
    font-size: 12.5px;
  }
  .about-meta > div {
    display: flex;
    gap: 12px;
    padding: 5px 0;
    border-bottom: 1px solid var(--rule);
  }
  .about-meta dt {
    flex: none;
    width: 72px;
    color: var(--muted);
  }
  .about-meta dd {
    margin: 0;
    min-width: 0;
    overflow-wrap: anywhere;
  }

  .about-links {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
  }
  .about-links a {
    padding: 5px 11px;
    border: 1px solid var(--rule);
    border-radius: 7px;
    color: inherit;
    font-size: 12.5px;
    text-decoration: none;
  }
  .about-links a:hover {
    background: var(--tint);
  }
  .about-links a:focus-visible {
    outline: 2px solid var(--accent);
    outline-offset: 2px;
  }

  .about-rights {
    margin: 0;
    padding-top: 2px;
    color: var(--muted);
    font-size: 11px;
    line-height: 1.5;
  }

  @media (max-width: 560px) {
    .settings {
      height: min(600px, 90vh);
    }
    /* On a narrow panel the rail turns into a horizontal tab strip so the
       content keeps the full width. */
    .settings-body {
      grid-template-columns: minmax(0, 1fr);
      grid-template-rows: auto minmax(0, 1fr);
    }
    .settings-rail {
      flex-direction: row;
      gap: 4px;
      padding: 8px;
      overflow-x: auto;
      overflow-y: hidden;
      border-right: 0;
      border-bottom: 1px solid var(--rule);
    }
    .settings-tab {
      flex: none;
      padding: 6px 10px;
      white-space: nowrap;
    }
    .settings-tab.active::before {
      display: none;
    }
    .settings-pane {
      padding: 14px;
    }
  }

  @media (prefers-reduced-motion: reduce) {
    * {
      transition: none !important;
      animation: none !important;
    }
  }

  @media (max-width: 480px) {
    .search-dialog {
      height: min(560px, 90vh);
      padding: 12px;
    }
    .search-result {
      flex-direction: column;
      gap: 4px;
      padding: 9px;
    }
  }
</style>
