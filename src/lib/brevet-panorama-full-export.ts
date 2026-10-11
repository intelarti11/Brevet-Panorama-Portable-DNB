import autoTable from "jspdf-autotable";
import * as XLSX from "./spreadsheet";

import { createBrevetPanoramaPdfDoc } from "@/lib/brevet-panorama-export";
import type { BrevetPanoramaReportData, BrevetPanoramaSubscore, BrevetPanoramaSubsubject } from "@/lib/brevet-panorama-report";

type CellValue = string | number | null | undefined;

/**
 * The report builder grows over time.  Keeping the export's view of the
 * extended rows local makes the exporter tolerant while the report data and
 * its tests evolve together.
 */
type ExportStudentRow = BrevetPanoramaReportData["studentRows"][number] & {
  sex?: "Fille" | "Garçon";
  scholarship?: boolean;
  rank?: number;
};

type ExportClassRow = BrevetPanoramaReportData["classRows"][number] & {
  girls?: number;
  boys?: number;
  genderKnown?: number;
  scholarshipKnown?: number;
  scholarshipYes?: number;
};

type ExportGroupStats = {
  totalStudents: number;
  averageCount: number;
  decidedCount: number;
  admis: number;
  refuse: number;
  successRate?: number;
  averageDnb?: number;
  averageBb1?: number;
  averageBb2?: number;
  averageDeltaBb2Dnb?: number;
};

type ExportComparisonRow = {
  subject: string;
  groupACount: number;
  groupAAverage?: number;
  groupBCount: number;
  groupBAverage?: number;
  gap?: number;
};

type ExportBreakdown = {
  groupA: ExportGroupStats;
  groupB: ExportGroupStats;
  specifiedCount: number;
  unspecifiedCount: number;
  averageGap?: number;
  successRateGap?: number;
  subjectRows: ExportComparisonRow[];
};

type ExportClassDetail = {
  className: string;
  summary: ExportClassRow;
  alphabeticalRows: ExportStudentRow[];
  rankedRows: ExportStudentRow[];
};

type ExtendedReportData = BrevetPanoramaReportData & {
  genderBreakdown?: ExportBreakdown;
  scholarshipBreakdown?: ExportBreakdown;
  alphabeticalRows?: ExportStudentRow[];
  rankedRows?: ExportStudentRow[];
  classDetails?: ExportClassDetail[];
};

const BLUE = "2563EB";
const LIGHT_BLUE = "EFF6FF";
const BORDER = "CBD5E1";
const UNKNOWN = "Inconnu";

const titleStyle = {
  font: { bold: true, size: 15, color: { rgb: BLUE } },
  alignment: { horizontal: "left", vertical: "center" },
};

const headerStyle = {
  font: { bold: true, color: { rgb: "FFFFFF" } },
  fill: { patternType: "solid", fgColor: { rgb: BLUE } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  border: {
    top: { style: "thin", color: { rgb: BORDER } },
    bottom: { style: "thin", color: { rgb: BORDER } },
    left: { style: "thin", color: { rgb: BORDER } },
    right: { style: "thin", color: { rgb: BORDER } },
  },
};

const bodyStyle = {
  alignment: { vertical: "center", wrapText: true },
  border: {
    top: { style: "thin", color: { rgb: "E2E8F0" } },
    bottom: { style: "thin", color: { rgb: "E2E8F0" } },
    left: { style: "thin", color: { rgb: "E2E8F0" } },
    right: { style: "thin", color: { rgb: "E2E8F0" } },
  },
};

const twoDecimalStyle = { ...bodyStyle, numFmt: "0.00" };

const summaryLabelStyle = {
  ...bodyStyle,
  font: { bold: true, color: { rgb: BLUE } },
  fill: { patternType: "solid", fgColor: { rgb: LIGHT_BLUE } },
};

const formatNumber = (value: number | undefined, digits = 2): string =>
  value === undefined ? "" : value.toLocaleString("fr-FR", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });

const safeFileName = (value: string): string =>
  value.replace(/[\\/:*?"<>|]/g, "-");

const formatUnknownNumber = (value: number | undefined, digits = 2): string =>
  value === undefined ? UNKNOWN : formatNumber(value, digits);

const formatPercent = (value: number | undefined): string =>
  value === undefined ? UNKNOWN : `${formatNumber(value, 1)} %`;

const textOrUnknown = (value: string | undefined): string =>
  value?.trim() ? value : UNKNOWN;

const sexLabel = (row: ExportStudentRow): string => row.sex ?? UNKNOWN;

const subsubjectHeader = (subject: BrevetPanoramaSubsubject): string =>
  `${subject.label} /${subject.maxScore}`;

const subsubjectPdfValue = (value: BrevetPanoramaSubscore | undefined): string =>
  typeof value === "string" ? value : pdfUnknownNumber(value);

const asExtendedReport = (report: BrevetPanoramaReportData): ExtendedReportData =>
  report as ExtendedReportData;

function compareNames(left: ExportStudentRow, right: ExportStudentRow): number {
  return left.lastName.localeCompare(right.lastName, "fr", {sensitivity: "base"}) ||
    left.firstName.localeCompare(right.firstName, "fr", {sensitivity: "base"}) ||
    (left.formerClass ?? "ZZZ").localeCompare(right.formerClass ?? "ZZZ", "fr", {numeric: true});
}

function compareRankedStudents(left: ExportStudentRow, right: ExportStudentRow): number {
  const averageComparison = (right.averageDnb ?? -Infinity) - (left.averageDnb ?? -Infinity);
  return averageComparison || compareNames(left, right);
}

function getAlphabeticalRows(report: ExtendedReportData): ExportStudentRow[] {
  return report.alphabeticalRows ?? [...report.studentRows as ExportStudentRow[]].sort(compareNames);
}

function getRankedRows(report: ExtendedReportData): ExportStudentRow[] {
  return report.rankedRows ?? [...report.studentRows as ExportStudentRow[]]
    .sort(compareRankedStudents);
}

function getClassDetails(report: ExtendedReportData): ExportClassDetail[] {
  if (report.classDetails) {
    return report.classDetails;
  }

  const rowsByClass = new Map<string, ExportStudentRow[]>();
  (report.studentRows as ExportStudentRow[]).forEach((row) => {
    if (!row.formerClass) return;
    const rows = rowsByClass.get(row.formerClass) ?? [];
    rows.push(row);
    rowsByClass.set(row.formerClass, rows);
  });

  return (report.classRows as ExportClassRow[]).map((summary) => {
    const rows = rowsByClass.get(summary.className) ?? [];
    return {
      className: summary.className,
      summary,
      alphabeticalRows: [...rows].sort(compareNames),
      rankedRows: [...rows].sort(compareRankedStudents),
    };
  });
}

function hasKnownBreakdown(breakdown: ExportBreakdown | undefined): boolean {
  return Boolean(breakdown && breakdown.specifiedCount > 0);
}

function displayGroupValue(
  breakdown: ExportBreakdown,
  value: number | undefined,
  digits = 2,
): string {
  if (!hasKnownBreakdown(breakdown)) return UNKNOWN;
  return value === undefined ? UNKNOWN : formatNumber(value, digits);
}

function displayGroupPercent(
  breakdown: ExportBreakdown,
  value: number | undefined,
): string {
  if (!hasKnownBreakdown(breakdown)) return UNKNOWN;
  return formatPercent(value);
}

function sanitizeSheetName(value: string, fallback = "Feuille"): string {
  const sanitized = value.replace(/[\\/*?:\[\]]/g, "-").trim() || fallback;
  return sanitized.slice(0, 31);
}

function uniqueSheetName(workbook: XLSX.WorkBook, value: string): string {
  const baseName = sanitizeSheetName(value);
  if (!workbook.SheetNames.includes(baseName)) return baseName;

  for (let suffix = 2; suffix < 1000; suffix += 1) {
    const suffixText = `-${suffix}`;
    const candidate = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`;
    if (!workbook.SheetNames.includes(candidate)) return candidate;
  }

  return `${baseName.slice(0, 27)}-999`;
}

function applyStyle(
  sheet: XLSX.WorkSheet,
  address: string,
  style: Record<string, unknown>,
) {
  const cell = sheet[address] as (XLSX.CellObject & { s?: Record<string, unknown> }) | undefined;
  if (cell) cell.s = style;
}

function createTableSheet(
  title: string,
  subtitle: string,
  headers: string[],
  rows: CellValue[][],
  widths: number[],
  twoDecimalColumns: number[] = [],
): XLSX.WorkSheet {
  const data: CellValue[][] = [
    [title],
    [subtitle],
    [],
    headers,
    ...rows,
  ];
  const sheet = XLSX.utils.aoa_to_sheet(data);
  const lastColumn = XLSX.utils.encode_col(Math.max(headers.length - 1, 0));
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(headers.length - 1, 0) } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(headers.length - 1, 0) } },
  ];
  sheet["!cols"] = widths.map((width) => ({ wch: width }));
  sheet["!rows"] = [{ hpt: 24 }, { hpt: 19 }, {}, { hpt: 34 }];
  sheet["!autofilter"] = { ref: `A4:${lastColumn}${Math.max(rows.length + 4, 4)}` };
  sheet["!freeze"] = { xSplit: 0, ySplit: 4 };

  applyStyle(sheet, "A1", titleStyle);
  applyStyle(sheet, "A2", {
    font: { italic: true, color: { rgb: "64748B" } },
  });

  headers.forEach((_, columnIndex) => {
    applyStyle(sheet, XLSX.utils.encode_cell({ r: 3, c: columnIndex }), headerStyle);
  });
  rows.forEach((_, rowIndex) => {
    headers.forEach((__, columnIndex) => {
      const value = rows[rowIndex][columnIndex];
      applyStyle(
        sheet,
        XLSX.utils.encode_cell({ r: rowIndex + 4, c: columnIndex }),
        columnIndex === 0 && headers.length === 2 ? summaryLabelStyle :
          twoDecimalColumns.includes(columnIndex) && typeof value === "number" ? twoDecimalStyle : bodyStyle,
      );
    });
  });

  return sheet;
}

interface SheetSection {
  title: string;
  headers: string[];
  rows: CellValue[][];
}

function createSectionedSheet(
  title: string,
  subtitle: string,
  sections: SheetSection[],
  widths: number[],
): XLSX.WorkSheet {
  const data: CellValue[][] = [[title], [subtitle], []];
  const sectionRows: Array<{headerRow: number; dataStart: number; dataEnd: number; headers: string[]}> = [];

  sections.forEach((section, sectionIndex) => {
    if (sectionIndex > 0) data.push([]);
    data.push([section.title]);
    const headerRow = data.length;
    data.push(section.headers);
    const dataStart = data.length;
    data.push(...section.rows);
    const dataEnd = data.length - 1;
    sectionRows.push({headerRow, dataStart, dataEnd, headers: section.headers});
  });

  const sheet = XLSX.utils.aoa_to_sheet(data);
  const lastColumn = XLSX.utils.encode_col(Math.max(widths.length - 1, 0));
  sheet["!cols"] = widths.map((width) => ({wch: width}));
  sheet["!rows"] = data.map((_row, rowIndex) => ({
    hpt: rowIndex === 0 ? 24 : rowIndex === 1 ? 19 : 20,
  }));
  sheet["!freeze"] = {xSplit: 0, ySplit: 4};
  sheet["!autofilter"] = {ref: `A1:${lastColumn}${Math.max(data.length, 1)}`};

  applyStyle(sheet, "A1", titleStyle);
  applyStyle(sheet, "A2", {
    font: {italic: true, color: {rgb: "64748B"}},
  });

  sectionRows.forEach(({headerRow, dataStart, dataEnd, headers}) => {
    const sectionTitleAddress = `A${headerRow}`;
    if (sheet[sectionTitleAddress]) {
      sheet[sectionTitleAddress].s = summaryLabelStyle;
    }

    headers.forEach((_, columnIndex) => {
      const address = XLSX.utils.encode_cell({r: headerRow, c: columnIndex});
      if (sheet[address]) sheet[address].s = headerStyle;
    });

    for (let rowIndex = dataStart; rowIndex <= dataEnd; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < widths.length; columnIndex += 1) {
        const address = XLSX.utils.encode_cell({r: rowIndex, c: columnIndex});
        if (sheet[address]) sheet[address].s = bodyStyle;
      }
    }
  });

  return sheet;
}

function buildSummaryRows(report: BrevetPanoramaReportData): CellValue[][] {
  const { stats, coverage } = report;
  const extended = asExtendedReport(report);
  const rows: CellValue[][] = [
    ["Session", report.yearLabel],
    ["Élèves dans la sélection", stats.totalStudents],
    ["Admis", stats.admis],
    ["Refusés", stats.refuse],
    ["Taux de réussite", `${formatNumber(stats.successRate, 1)} %`],
    ["Moyenne des admis", formatNumber(stats.averageOverallScoreAdmitted)],
    ["Mention Très Bien", stats.mentions.tresBien],
    ["Mention Bien", stats.mentions.bien],
    ["Mention Assez Bien", stats.mentions.assezBien],
    ["Admis sans mention", stats.mentions.sansMention],
    ["Élèves rattachés à une ancienne classe", coverage.studentsWithClass],
    ["Couverture des anciennes classes", `${formatNumber(coverage.classCoverageRate, 1)} %`],
    ["Élèves rattachés par division officielle", coverage.officialClassCount],
    ["Élèves rattachés par la liste brevet blanc (INE)", coverage.brevetBlancClassCount],
    ["Élèves sans ancienne classe", coverage.studentsWithoutClass],
    ["Élèves avec au moins une moyenne de brevet blanc", coverage.studentsWithBrevetBlancScores],
  ];

  const appendBreakdownCoverage = (label: string, breakdown: ExportBreakdown | undefined) => {
    if (!breakdown) return;
    const total = breakdown.specifiedCount + breakdown.unspecifiedCount;
    rows.push(
      [`${label} connu`, `${breakdown.specifiedCount}/${total}`],
      [`${label} inconnu`, breakdown.unspecifiedCount],
    );
  };

  appendBreakdownCoverage("Sexe", extended.genderBreakdown);
  appendBreakdownCoverage("Statut boursier", extended.scholarshipBreakdown);
  return rows;
}

function buildGroupSummaryRows(
  breakdown: ExportBreakdown,
  groupALabel: string,
  groupBLabel: string,
): CellValue[][] {
  const metrics: Array<{
    label: string;
    value: (group: ExportGroupStats) => number | undefined;
    percent?: boolean;
  }> = [
    {label: "Effectif", value: (group) => group.totalStudents},
    {label: "Moyennes DNB connues", value: (group) => group.averageCount},
    {label: "Décisions connues", value: (group) => group.decidedCount},
    {label: "Admis", value: (group) => group.admis},
    {label: "Refusés", value: (group) => group.refuse},
    {label: "Réussite", value: (group) => group.successRate, percent: true},
    {label: "Moyenne DNB", value: (group) => group.averageDnb},
    {label: "Moyenne BB1", value: (group) => group.averageBb1},
    {label: "Moyenne BB2", value: (group) => group.averageBb2},
    {label: "Écart BB2 → DNB", value: (group) => group.averageDeltaBb2Dnb},
  ];
  const breakdownRows = metrics.map((metric) => [
    metric.label,
    metric.percent ? displayGroupPercent(breakdown, metric.value(breakdown.groupA)) : displayGroupValue(breakdown, metric.value(breakdown.groupA)),
    metric.percent ? displayGroupPercent(breakdown, metric.value(breakdown.groupB)) : displayGroupValue(breakdown, metric.value(breakdown.groupB)),
    metric.percent ? formatPercent(metric.value(breakdown.groupA) !== undefined && metric.value(breakdown.groupB) !== undefined ?
      (metric.value(breakdown.groupA) as number) - (metric.value(breakdown.groupB) as number) : undefined) :
      formatUnknownNumber(metric.value(breakdown.groupA) !== undefined && metric.value(breakdown.groupB) !== undefined ?
        (metric.value(breakdown.groupA) as number) - (metric.value(breakdown.groupB) as number) : undefined),
  ]);

  return [
    ["Métrique globale", groupALabel, groupBLabel, "Écart A-B"],
    ...breakdownRows,
  ];
}

function buildComparisonSubjectRows(
  breakdown: ExportBreakdown | undefined,
): CellValue[][] {
  if (!breakdown || !hasKnownBreakdown(breakdown)) {
    return [[UNKNOWN, UNKNOWN, UNKNOWN, UNKNOWN, UNKNOWN, UNKNOWN]];
  }

  return breakdown.subjectRows.map((row) => [
    row.subject,
    row.groupACount,
    formatUnknownNumber(row.groupAAverage),
    row.groupBCount,
    formatUnknownNumber(row.groupBAverage),
    formatUnknownNumber(row.gap),
  ]);
}

function appendBreakdownSheet(
  workbook: XLSX.WorkBook,
  report: ExtendedReportData,
  sheetName: string,
  title: string,
  groupALabel: string,
  groupBLabel: string,
  breakdown: ExportBreakdown | undefined,
) {
  const known = hasKnownBreakdown(breakdown);
  const total = breakdown ? breakdown.specifiedCount + breakdown.unspecifiedCount : 0;
  const subtitle = breakdown ?
    `${report.yearLabel} — métadonnées connues : ${breakdown.specifiedCount}/${total}; les absences restent « ${UNKNOWN} »` :
    `${report.yearLabel} — données indisponibles`;
  const safeBreakdown = breakdown ?? {
    groupA: {totalStudents: 0, averageCount: 0, decidedCount: 0, admis: 0, refuse: 0},
    groupB: {totalStudents: 0, averageCount: 0, decidedCount: 0, admis: 0, refuse: 0},
    specifiedCount: 0,
    unspecifiedCount: 0,
    subjectRows: [],
  } satisfies ExportBreakdown;

  const summaryRows = known ?
    buildGroupSummaryRows(safeBreakdown, groupALabel, groupBLabel).slice(1) :
    [["Aucune métadonnée connue", UNKNOWN, UNKNOWN, UNKNOWN]];
  const subjectRows = buildComparisonSubjectRows(breakdown);

  XLSX.utils.book_append_sheet(
    workbook,
    createSectionedSheet(
      title,
      subtitle,
      [
        {
          title: "Synthèse globale",
          headers: ["Métrique", groupALabel, groupBLabel, "Écart A-B"],
          rows: summaryRows,
        },
        {
          title: "Comparaison par matière",
          headers: ["Matière / composante", `Effectif ${groupALabel}`, `Moy. ${groupALabel}`, `Effectif ${groupBLabel}`, `Moy. ${groupBLabel}`, "Écart"],
          rows: subjectRows,
        },
      ],
      [32, 19, 19, 19, 19, 14],
    ),
    uniqueSheetName(workbook, sheetName),
  );
}

function appendClassMatrixSheet(workbook: XLSX.WorkBook, report: BrevetPanoramaReportData) {
  const french = report.subsubjects.filter((subject) => subject.group === "francais");
  const sciences = report.subsubjects.filter((subject) => subject.group === "sciences");
  const rows = (report.classRows as ExportClassRow[]).map((row) => [
    row.className,
    row.totalStudents,
    formatNumber(row.averageDnb),
    formatNumber(row.averageControleContinu),
    formatNumber(row.averageEpreuvesTerminales),
    formatNumber(row.averageFrancais),
    ...french.map((subject) => formatNumber(row[subject.classKey])),
    formatNumber(row.averageMaths),
    formatNumber(row.averageHistoireGeoEmc),
    formatNumber(row.averageSciences),
    ...sciences.map((subject) => formatNumber(row[subject.classKey])),
    formatNumber(row.averageBb1),
    formatNumber(row.averageBb2),
    formatNumber(row.averageDeltaBb2Dnb),
  ]);

  XLSX.utils.book_append_sheet(
    workbook,
    createTableSheet(
      "Matrice des moyennes par classe",
      `Session ${report.yearLabel} — sous-épreuves aux barèmes indiqués`,
      ["Classe", "Effectif", "DNB", "Contrôle continu", "Épreuves terminales", "Français", ...french.map(subsubjectHeader), "Maths", "HG-EMC", "Sciences", ...sciences.map(subsubjectHeader), "BB1", "BB2", "BB2 → DNB"],
      rows,
      [14, 10, 11, 17, 19, 12, ...french.map(() => 20), 12, 12, 12, ...sciences.map(() => 20), 11, 11, 14],
    ),
    "Matrice classes",
  );
}

function getDnbSheetDescription(sheetName: string): string {
  if (sheetName === "Synthèse") return "Indicateurs globaux, mentions et couverture";
  if (sheetName === "Par classe") return "Résultats et moyennes par ancienne classe";
  if (sheetName === "Par matière") return "Moyennes, minima et maxima par composante";
  if (sheetName === "Top 10") return "Dix meilleures moyennes DNB";
  if (sheetName === "Élèves") return "Détail complet des résultats individuels";
  if (sheetName === "03 - Filles-Garçons") return "Comparatif filles / garçons, global et par matière";
  if (sheetName === "04 - Boursiers") return "Comparatif boursiers / non-boursiers, global et par matière";
  if (sheetName === "Matrice classes") return "Matrice des moyennes DNB, matières et brevets blancs";
  if (sheetName === "Tout alpha") return "Tous les élèves par ordre alphabétique";
  if (sheetName === "Classement") return "Tous les élèves par moyenne DNB décroissante";
  if (sheetName.endsWith(" - Alpha")) return "Élèves de la classe par ordre alphabétique";
  return "Élèves de la classe par moyenne décroissante";
}

function appendWorkbookMenuSheet(workbook: XLSX.WorkBook, report: BrevetPanoramaReportData) {
  const sheetNames = workbook.SheetNames.filter((sheetName) => sheetName !== "Menu");
  const menuRows: CellValue[][] = [
    ["Menu du bilan complet DNB"],
    [`Session ${report.yearLabel} — cliquer sur un onglet pour y accéder`],
    [],
    ["Onglet", "Description"],
    ...sheetNames.map((sheetName) => [sheetName, getDnbSheetDescription(sheetName)]),
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(menuRows);
  worksheet["!cols"] = [{wch: 30}, {wch: 68}];
  worksheet["!rows"] = menuRows.map((_, index) => ({hpt: index === 0 ? 24 : index === 3 ? 30 : 20}));
  worksheet["!merges"] = [
    {s: {r: 0, c: 0}, e: {r: 0, c: 1}},
    {s: {r: 1, c: 0}, e: {r: 1, c: 1}},
  ];
  worksheet["!freeze"] = {xSplit: 0, ySplit: 4};

  applyStyle(worksheet, "A1", titleStyle);
  applyStyle(worksheet, "A2", {font: {italic: true, color: {rgb: "64748B"}}});
  applyStyle(worksheet, "A4", headerStyle);
  applyStyle(worksheet, "B4", headerStyle);
  sheetNames.forEach((sheetName, index) => {
    const rowIndex = index + 4;
    const sheetCell = `A${rowIndex + 1}`;
    const descriptionCell = `B${rowIndex + 1}`;
    if (worksheet[sheetCell]) {
      worksheet[sheetCell].s = {...bodyStyle, font: {bold: true, color: {rgb: BLUE}}};
      worksheet[sheetCell].l = {
        Target: `#'${sheetName.replace(/'/g, "''")}'!A1`,
        Tooltip: `Aller à ${sheetName}`,
      };
    }
    if (worksheet[descriptionCell]) worksheet[descriptionCell].s = bodyStyle;
  });

  XLSX.utils.book_append_sheet(workbook, worksheet, "Menu");
  workbook.SheetNames = ["Menu", ...workbook.SheetNames.filter((sheetName) => sheetName !== "Menu")];
}

export function createBrevetPanoramaWorkbook(report: BrevetPanoramaReportData): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new();
  const subtitle = `Bilan complet du DNB — session ${report.yearLabel}`;
  const frenchSubsubjects = report.subsubjects.filter((subject) => subject.group === "francais");
  const scienceSubsubjects = report.subsubjects.filter((subject) => subject.group === "sciences");
  const subsubjectCount = report.subsubjects.length;

  XLSX.utils.book_append_sheet(
    workbook,
    createTableSheet(
      "Synthèse du DNB",
      subtitle,
      ["Indicateur", "Valeur"],
      buildSummaryRows(report),
      [45, 24],
    ),
    "Synthèse",
  );

  XLSX.utils.book_append_sheet(
    workbook,
    createTableSheet(
      "Comparaison par ancienne classe",
      `${subtitle} — rattachement par division officielle ou INE exact`,
      [
        "Classe",
        "Effectif",
        "Filles",
        "Garçons",
        "Sexe connu",
        "Boursiers",
        "Bourse connue",
        "Admis",
        "Refusés",
        "Réussite (%)",
        "Moy. DNB",
        "Moy. contrôle continu",
        "Moy. épreuves terminales",
        "Français",
        ...frenchSubsubjects.map(subsubjectHeader),
        "Maths",
        "HG-EMC",
        "Sciences",
        ...scienceSubsubjects.map(subsubjectHeader),
        "Moy. BB1",
        "Moy. BB2",
        "Écart BB2 → DNB",
        "Élèves appariés BB",
      ],
      report.classRows.map((row) => [
        row.className,
        row.totalStudents,
        row.genderKnown > 0 ? row.girls ?? UNKNOWN : UNKNOWN,
        row.genderKnown > 0 ? row.boys ?? UNKNOWN : UNKNOWN,
        row.genderKnown > 0 ? row.genderKnown : UNKNOWN,
        row.scholarshipKnown > 0 ? row.scholarshipYes ?? UNKNOWN : UNKNOWN,
        row.scholarshipKnown > 0 ? row.scholarshipKnown : UNKNOWN,
        row.admis,
        row.refuse,
        row.successRate,
        row.averageDnb,
        row.averageControleContinu,
        row.averageEpreuvesTerminales,
        row.averageFrancais,
        ...frenchSubsubjects.map((subject) => row[subject.classKey]),
        row.averageMaths,
        row.averageHistoireGeoEmc,
        row.averageSciences,
        ...scienceSubsubjects.map((subject) => row[subject.classKey]),
        row.averageBb1,
        row.averageBb2,
        row.averageDeltaBb2Dnb,
        row.matchedBbCount,
      ]),
      [14, 10, 9, 9, 11, 11, 12, 9, 9, 13, 11, 17, 18, 11, ...frenchSubsubjects.map(() => 21), 11, 11, 11, 11, ...scienceSubsubjects.map(() => 21), 11, 11, 15, 16],
      Array.from({ length: 10 + subsubjectCount }, (_, index) => index + 10),
    ),
    "Par classe",
  );

  XLSX.utils.book_append_sheet(
    workbook,
    createTableSheet(
      "Analyse par matière et composante",
      `${subtitle} — sous-épreuves aux barèmes indiqués`,
      ["Matière ou composante", "Coefficient / poids", "Notes", "Moyenne", "Minimum", "Maximum"],
      report.subjectRows.map((row) => [
        row.subject,
        row.coefficient,
        row.count,
        row.average,
        row.minimum,
        row.maximum,
      ]),
      [35, 18, 10, 12, 12, 12],
      [3, 4, 5],
    ),
    "Par matière",
  );

  XLSX.utils.book_append_sheet(
    workbook,
    createTableSheet(
      "Classement des dix meilleures moyennes",
      subtitle,
      ["Rang", "Nom", "Prénom", "Ancienne classe", "Moyenne DNB", "Résultat"],
      report.top10.map((row, index) => [
        index + 1,
        row.lastName,
        row.firstName,
        row.formerClass,
        row.averageDnb,
        row.result,
      ]),
      [8, 22, 20, 15, 14, 24],
      [4],
    ),
    "Top 10",
  );

  XLSX.utils.book_append_sheet(
    workbook,
    createTableSheet(
      "Détail des élèves",
      `${subtitle} — les anciennes classes issues du brevet blanc sont appariées uniquement par INE exact`,
      [
        "Rang DNB",
        "Nom",
        "Prénom",
        "Ancienne classe",
        "Série",
        "Sexe",
        "Résultat",
        "Moy. DNB",
        "Contrôle continu",
        "Épreuves terminales",
        "Français",
        ...frenchSubsubjects.map(subsubjectHeader),
        "Maths",
        "Histoire-Géo",
        "EMC",
        "HG-EMC pondérés",
        "Sciences",
        ...scienceSubsubjects.map(subsubjectHeader),
        "Oral",
        "Moy. BB1",
        "Moy. BB2",
        "Écart BB2 → DNB",
      ],
      getRankedRows(asExtendedReport(report)).map((row) => [
        row.rank ?? "",
        row.lastName,
        row.firstName,
        row.formerClass,
        row.serie,
        sexLabel(row),
        row.result,
        row.averageDnb,
        row.controleContinu,
        row.epreuvesTerminales,
        row.francais,
        ...frenchSubsubjects.map((subject) => row[subject.studentKey]),
        row.maths,
        row.histoireGeo,
        row.emc,
        row.histoireGeoEmc,
        row.sciences,
        ...scienceSubsubjects.map((subject) => row[subject.studentKey]),
        row.oral,
        row.averageBb1,
        row.averageBb2,
        row.deltaBb2Dnb,
      ]),
      [10, 22, 20, 15, 16, 11, 25, 12, 16, 17, 11, ...frenchSubsubjects.map(() => 21), 11, 14, 10, 16, 11, ...scienceSubsubjects.map(() => 21), 11, 11, 11, 15],
      Array.from({ length: 13 + subsubjectCount }, (_, index) => index + 7),
    ),
    "Élèves",
  );

  const extended = asExtendedReport(report);
  appendBreakdownSheet(
    workbook,
    extended,
    "03 - Filles-Garçons",
    "Comparaison filles / garçons",
    "Filles",
    "Garçons",
    extended.genderBreakdown,
  );
  appendBreakdownSheet(
    workbook,
    extended,
    "04 - Boursiers",
    "Comparaison boursiers / non-boursiers",
    "Boursiers",
    "Non-boursiers",
    extended.scholarshipBreakdown,
  );
  appendClassMatrixSheet(workbook, report);

  const rankedHeaders = [
    "Rang DNB",
    "Nom",
    "Prénom",
    "Ancienne classe",
    "Sexe",
    "Résultat",
    "Moy. DNB",
    "Contrôle continu",
    "Épreuves terminales",
    "Français",
    ...frenchSubsubjects.map(subsubjectHeader),
    "Maths",
    "Histoire-Géo",
    "EMC",
    "HG-EMC pondérés",
    "Sciences",
    ...scienceSubsubjects.map(subsubjectHeader),
    "Oral",
    "Moy. BB1",
    "Moy. BB2",
    "Écart BB2 → DNB",
  ];
  const alphabeticalHeaders = ["N°", ...rankedHeaders.slice(1)];
  const studentWidths = [10, 20, 18, 15, 11, 20, 12, 16, 17, 11, ...frenchSubsubjects.map(() => 21), 11, 14, 10, 16, 11, ...scienceSubsubjects.map(() => 21), 11, 11, 11, 15];
  const toStudentValues = (row: ExportStudentRow, index: number, includeRank: boolean): CellValue[] => [
    includeRank ? row.rank : index + 1,
    row.lastName,
    row.firstName,
    row.formerClass,
    sexLabel(row),
    row.result,
    row.averageDnb,
    row.controleContinu,
    row.epreuvesTerminales,
    row.francais,
    ...frenchSubsubjects.map((subject) => row[subject.studentKey]),
    row.maths,
    row.histoireGeo,
    row.emc,
    row.histoireGeoEmc,
    row.sciences,
    ...scienceSubsubjects.map((subject) => row[subject.studentKey]),
    row.oral,
    row.averageBb1,
    row.averageBb2,
    row.deltaBb2Dnb,
  ];

  XLSX.utils.book_append_sheet(
    workbook,
    createTableSheet(
      "Tous les élèves — ordre alphabétique",
      `${subtitle} — sans fiche brevet blanc appariée, sexe et bourse restent inconnus`,
      alphabeticalHeaders,
      getAlphabeticalRows(extended).map((row, index) => toStudentValues(row, index, false)),
      studentWidths,
      Array.from({ length: 13 + subsubjectCount }, (_, index) => index + 6),
    ),
    "Tout alpha",
  );

  XLSX.utils.book_append_sheet(
    workbook,
    createTableSheet(
      "Tous les élèves — classement par moyenne",
      `${subtitle} — rang DNB, moyenne décroissante; les élèves sans moyenne restent en fin de tableau`,
      rankedHeaders,
      getRankedRows(extended).map((row, index) => toStudentValues(row, index, true)),
      studentWidths,
      Array.from({ length: 13 + subsubjectCount }, (_, index) => index + 6),
    ),
    "Classement",
  );

  getClassDetails(extended).forEach((classDetail) => {
    const classSubtitle = `${subtitle} — ${classDetail.className}; inconnues conservées comme « ${UNKNOWN} »`;
    XLSX.utils.book_append_sheet(
      workbook,
      createTableSheet(
        `${classDetail.className} — ordre alphabétique`,
        classSubtitle,
        alphabeticalHeaders.filter((header) => header !== "Ancienne classe"),
        classDetail.alphabeticalRows.map((row, index) => toStudentValues(row, index, false).filter((_, columnIndex) => columnIndex !== 3)),
        studentWidths.filter((_, columnIndex) => columnIndex !== 3),
        Array.from({ length: 13 + subsubjectCount }, (_, index) => index + 5),
      ),
      uniqueSheetName(workbook, `${classDetail.className} - Alpha`),
    );

    XLSX.utils.book_append_sheet(
      workbook,
      createTableSheet(
        `${classDetail.className} — classement par moyenne`,
        classSubtitle,
        rankedHeaders.filter((header) => header !== "Ancienne classe"),
        classDetail.rankedRows.map((row, index) => toStudentValues(row, index, true).filter((_, columnIndex) => columnIndex !== 3)),
        studentWidths.filter((_, columnIndex) => columnIndex !== 3),
        Array.from({ length: 13 + subsubjectCount }, (_, index) => index + 5),
      ),
      uniqueSheetName(workbook, classDetail.className),
    );
  });

  appendWorkbookMenuSheet(workbook, report);

  return workbook;
}

export function exportBrevetPanoramaXlsx(report: BrevetPanoramaReportData) {
  const workbook = createBrevetPanoramaWorkbook(report);
  XLSX.writeFile(workbook, `Bilan_complet_DNB_${safeFileName(report.yearLabel)}.xlsx`);
}

function getPdfFont(doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>): string {
  const fontList = doc.getFontList();
  return Object.prototype.hasOwnProperty.call(fontList, "NotoSansBrevetPdf")
    ? "NotoSansBrevetPdf"
    : "helvetica";
}

function drawPdfHeader(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  title: string,
  subtitle: string,
) {
  const font = getPdfFont(doc);
  doc.setFont(font, "bold");
  doc.setFontSize(17);
  doc.setTextColor(37, 99, 235);
  doc.text(title, 12, 14);
  doc.setFont(font, "normal");
  doc.setFontSize(9);
  doc.setTextColor(100, 116, 139);
  doc.text(subtitle, 12, 21);
}

function pdfTableStyles(doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>) {
  return {
    font: getPdfFont(doc),
    fontSize: 7.2,
    cellPadding: 1.8,
    textColor: [15, 23, 42] as [number, number, number],
    lineColor: [226, 232, 240] as [number, number, number],
    lineWidth: 0.15,
  };
}

function pdfUnknownNumber(value: number | undefined, digits = 2): string {
  return value === undefined ? UNKNOWN : formatNumber(value, digits);
}

function pdfGroupNumber(
  breakdown: ExportBreakdown,
  value: number | undefined,
  digits = 2,
): string {
  return hasKnownBreakdown(breakdown) ? pdfUnknownNumber(value, digits) : UNKNOWN;
}

function pdfGroupPercent(breakdown: ExportBreakdown, value: number | undefined): string {
  return hasKnownBreakdown(breakdown) ? formatPercent(value) : UNKNOWN;
}

function addGroupComparisonPage(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  report: ExtendedReportData,
  breakdown: ExportBreakdown | undefined,
  title: string,
  groupALabel: string,
  groupBLabel: string,
) {
  doc.addPage("a4", "landscape");
  const total = breakdown ? breakdown.specifiedCount + breakdown.unspecifiedCount : 0;
  drawPdfHeader(
    doc,
    title,
    breakdown ?
      `Session ${report.yearLabel} — métadonnées connues : ${breakdown.specifiedCount}/${total}; ${breakdown.unspecifiedCount} « ${UNKNOWN} »` :
      `Session ${report.yearLabel} — données indisponibles`,
  );

  if (!breakdown || !hasKnownBreakdown(breakdown)) {
    doc.setFont(getPdfFont(doc), "normal");
    doc.setFontSize(11);
    doc.setTextColor(100, 116, 139);
    doc.text(`Aucune métadonnée ${groupALabel.toLowerCase()} / ${groupBLabel.toLowerCase()} connue.`, 12, 40);
    return;
  }

  autoTable(doc, {
    startY: 29,
    margin: {left: 10, right: 10},
    theme: "grid",
    head: [["Groupe", "Effectif", "Moy. DNB", "Moy. BB1", "Moy. BB2", "Écart BB2-DNB", "Admis", "Refusés", "Réussite"]],
    body: [
      [
        groupALabel,
        breakdown.groupA.totalStudents,
        pdfGroupNumber(breakdown, breakdown.groupA.averageDnb),
        pdfGroupNumber(breakdown, breakdown.groupA.averageBb1),
        pdfGroupNumber(breakdown, breakdown.groupA.averageBb2),
        pdfGroupNumber(breakdown, breakdown.groupA.averageDeltaBb2Dnb),
        breakdown.groupA.admis,
        breakdown.groupA.refuse,
        pdfGroupPercent(breakdown, breakdown.groupA.successRate),
      ],
      [
        groupBLabel,
        breakdown.groupB.totalStudents,
        pdfGroupNumber(breakdown, breakdown.groupB.averageDnb),
        pdfGroupNumber(breakdown, breakdown.groupB.averageBb1),
        pdfGroupNumber(breakdown, breakdown.groupB.averageBb2),
        pdfGroupNumber(breakdown, breakdown.groupB.averageDeltaBb2Dnb),
        breakdown.groupB.admis,
        breakdown.groupB.refuse,
        pdfGroupPercent(breakdown, breakdown.groupB.successRate),
      ],
    ],
    styles: {...pdfTableStyles(doc), fontSize: 8.2, cellPadding: 2},
    headStyles: {fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold"},
    alternateRowStyles: {fillColor: [248, 250, 252]},
  });

  if (breakdown.subjectRows.length === 0) {
    doc.setFont(getPdfFont(doc), "normal");
    doc.setFontSize(10);
    doc.setTextColor(100, 116, 139);
    doc.text(`Aucune moyenne par matière connue pour les groupes ${groupALabel} / ${groupBLabel}.`, 12, 65);
    return;
  }

  autoTable(doc, {
    startY: 65,
    margin: {left: 10, right: 10},
    theme: "grid",
    head: [["Matière / composante", `Effectif ${groupALabel}`, `Moy. ${groupALabel}`, `Effectif ${groupBLabel}`, `Moy. ${groupBLabel}`, "Écart A-B"]],
    body: breakdown.subjectRows.map((row) => [
      row.subject,
      row.groupACount,
      pdfUnknownNumber(row.groupAAverage),
      row.groupBCount,
      pdfUnknownNumber(row.groupBAverage),
      pdfUnknownNumber(row.gap),
    ]),
    styles: {...pdfTableStyles(doc), fontSize: 7.4, cellPadding: 1.7},
    headStyles: {fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold"},
    alternateRowStyles: {fillColor: [248, 250, 252]},
  });
}

function studentPdfHeaders(includeRank: boolean): string[] {
  return [
    ...(includeRank ? ["Rang"] : []),
    "Nom",
    "Prénom",
    "Classe",
    "Sexe",
    "Résultat",
    "Moy. DNB",
    "CC",
    "Terminales",
    "Français",
    "Maths",
    "HG",
    "EMC",
    "HG-EMC",
    "Sciences",
    "Oral",
    "BB1",
    "BB2",
    "Écart",
  ];
}

function studentPdfRow(row: ExportStudentRow, includeRank: boolean): Array<string | number> {
  return [
    ...(includeRank ? [row.rank ?? ""] : []),
    textOrUnknown(row.lastName),
    textOrUnknown(row.firstName),
    textOrUnknown(row.formerClass),
    sexLabel(row),
    textOrUnknown(row.result),
    pdfUnknownNumber(row.averageDnb),
    pdfUnknownNumber(row.controleContinu),
    pdfUnknownNumber(row.epreuvesTerminales),
    pdfUnknownNumber(row.francais),
    pdfUnknownNumber(row.maths),
    pdfUnknownNumber(row.histoireGeo),
    pdfUnknownNumber(row.emc),
    pdfUnknownNumber(row.histoireGeoEmc),
    pdfUnknownNumber(row.sciences),
    pdfUnknownNumber(row.oral),
    pdfUnknownNumber(row.averageBb1),
    pdfUnknownNumber(row.averageBb2),
    pdfUnknownNumber(row.deltaBb2Dnb),
  ];
}

function addStudentListPage(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  title: string,
  subtitle: string,
  rows: ExportStudentRow[],
  includeRank: boolean,
) {
  doc.addPage("a4", "landscape");
  if (rows.length === 0) {
    drawPdfHeader(doc, title, subtitle);
    doc.setFont(getPdfFont(doc), "normal");
    doc.setFontSize(11);
    doc.setTextColor(100, 116, 139);
    doc.text("Aucun élève à afficher.", 12, 40);
    return;
  }

  autoTable(doc, {
    startY: 29,
    margin: {top: 28, left: 5, right: 5, bottom: 12},
    theme: "grid",
    head: [studentPdfHeaders(includeRank)],
    body: rows.map((row) => studentPdfRow(row, includeRank)),
    styles: {...pdfTableStyles(doc), fontSize: 5.35, cellPadding: 0.85, overflow: "linebreak", valign: "middle"},
    headStyles: {fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold", fontSize: 5.35},
    alternateRowStyles: {fillColor: [248, 250, 252]},
    columnStyles: {
      0: {cellWidth: includeRank ? 8 : 22},
      1: {cellWidth: includeRank ? 22 : 20},
      2: {cellWidth: includeRank ? 20 : 15},
      3: {cellWidth: includeRank ? 15 : 10},
      4: {cellWidth: includeRank ? 10 : 20},
      5: {cellWidth: includeRank ? 20 : 12},
    },
    willDrawPage: () => drawPdfHeader(doc, title, subtitle),
  });
}

function addClassComparisonPage(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  report: BrevetPanoramaReportData,
) {
  doc.addPage("a4", "landscape");
  drawPdfHeader(
    doc,
    "Comparaison par ancienne classe",
    `Session ${report.yearLabel} — ${report.coverage.studentsWithClass}/${report.coverage.totalStudents} élèves rattachés (${formatNumber(report.coverage.classCoverageRate, 1)} %)`,
  );

  if (report.classRows.length === 0) {
    doc.setFont(getPdfFont(doc), "normal");
    doc.setFontSize(11);
    doc.setTextColor(100, 116, 139);
    doc.text(
      "Aucune ancienne classe n'a pu être rattachée. Le rapport global reste disponible.",
      12,
      38,
    );
    return;
  }

  autoTable(doc, {
    startY: 29,
    margin: { top: 28, left: 10, right: 10, bottom: 12 },
    theme: "grid",
    head: [[
      "Classe",
      "Eff.",
      "Admis",
      "Réussite",
      "Moy. DNB",
      "Contrôle continu",
      "Épreuves terminales",
      "Français",
      "Maths",
      "HG-EMC",
      "Sciences",
      "BB1",
      "BB2",
      "BB2-DNB",
    ]],
    body: report.classRows.map((row) => [
      row.className,
      row.totalStudents,
      row.admis,
      row.successRate === undefined ? "" : `${formatNumber(row.successRate, 1)} %`,
      formatNumber(row.averageDnb),
      formatNumber(row.averageControleContinu),
      formatNumber(row.averageEpreuvesTerminales),
      formatNumber(row.averageFrancais),
      formatNumber(row.averageMaths),
      formatNumber(row.averageHistoireGeoEmc),
      formatNumber(row.averageSciences),
      formatNumber(row.averageBb1),
      formatNumber(row.averageBb2),
      formatNumber(row.averageDeltaBb2Dnb),
    ]),
    styles: pdfTableStyles(doc),
    headStyles: { fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold" },
    alternateRowStyles: { fillColor: [248, 250, 252] },
    willDrawPage: () => drawPdfHeader(
      doc,
      "Comparaison par ancienne classe",
      `Session ${report.yearLabel} — ${report.coverage.studentsWithClass}/${report.coverage.totalStudents} élèves rattachés (${formatNumber(report.coverage.classCoverageRate, 1)} %)`,
    ),
  });

}

function addClassSubsubjectPage(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  report: BrevetPanoramaReportData,
) {
  if (report.subsubjects.length === 0 || report.classRows.length === 0) return;

  doc.addPage("a4", "landscape");
  const title = "Sous-épreuves par ancienne classe";
  const subtitle = `Session ${report.yearLabel} — effectif total; moyennes sur les notes disponibles, dans le barème indiqué`;
  autoTable(doc, {
    startY: 29,
    margin: { top: 28, left: 10, right: 10, bottom: 12 },
    theme: "grid",
    head: [["Classe", "Effectif", ...report.subsubjects.map(subsubjectHeader)]],
    body: report.classRows.map((row) => [
      row.className,
      row.totalStudents,
      ...report.subsubjects.map((subject) => formatNumber(row[subject.classKey])),
    ]),
    styles: pdfTableStyles(doc),
    headStyles: { fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold" },
    alternateRowStyles: { fillColor: [248, 250, 252] },
    willDrawPage: () => drawPdfHeader(doc, title, subtitle),
  });
}

function addSubjectAndRankingPage(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  report: BrevetPanoramaReportData,
) {
  doc.addPage("a4", "landscape");
  drawPdfHeader(doc, "Composantes et classement", `Session ${report.yearLabel} — sous-épreuves aux barèmes indiqués`);

  autoTable(doc, {
    startY: 29,
    margin: { left: 10, right: 154 },
    theme: "grid",
    head: [["Matière / composante", "Coeff. / poids", "Notes", "Moy.", "Min.", "Max."]],
    body: report.subjectRows.map((row) => [
      row.subject,
      row.coefficient ?? "",
      row.count,
      formatNumber(row.average),
      formatNumber(row.minimum),
      formatNumber(row.maximum),
    ]),
    styles: pdfTableStyles(doc),
    headStyles: { fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold" },
    alternateRowStyles: { fillColor: [248, 250, 252] },
  });

  autoTable(doc, {
    startY: 29,
    margin: { left: 151, right: 10 },
    theme: "grid",
    head: [["Rang", "Élève", "Classe", "Moy. DNB", "Résultat"]],
    body: report.top10.map((row, index) => [
      index + 1,
      `${row.lastName} ${row.firstName}`,
      row.formerClass ?? "",
      formatNumber(row.averageDnb),
      row.result ?? "",
    ]),
    styles: pdfTableStyles(doc),
    headStyles: { fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold" },
    alternateRowStyles: { fillColor: [248, 250, 252] },
  });

  // Le second tableau modifie l'état graphique de jsPDF ; redessiner l'en-tête
  // garantit qu'il reste intact dans tous les moteurs de rendu PDF.
  drawPdfHeader(doc, "Composantes et classement", `Session ${report.yearLabel} — sous-épreuves aux barèmes indiqués`);
}

function addStudentAppendix(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  report: BrevetPanoramaReportData,
) {
  doc.addPage("a4", "landscape");
  autoTable(doc, {
    startY: 28,
    margin: { top: 28, left: 7, right: 7, bottom: 12 },
    theme: "grid",
    head: [[
      "Nom",
      "Prénom",
      "Classe",
      "Résultat",
      "Moy. DNB",
      "CC",
      "Terminales",
      "Fr.",
      "Maths",
      "HG",
      "EMC",
      "Sciences",
      "Oral",
      "BB1",
      "BB2",
      "Écart",
    ]],
    body: report.studentRows.map((row) => [
      row.lastName,
      row.firstName,
      row.formerClass ?? "",
      row.result ?? "",
      formatNumber(row.averageDnb),
      formatNumber(row.controleContinu),
      formatNumber(row.epreuvesTerminales),
      formatNumber(row.francais),
      formatNumber(row.maths),
      formatNumber(row.histoireGeo),
      formatNumber(row.emc),
      formatNumber(row.sciences),
      formatNumber(row.oral),
      formatNumber(row.averageBb1),
      formatNumber(row.averageBb2),
      formatNumber(row.deltaBb2Dnb),
    ]),
    styles: { ...pdfTableStyles(doc), fontSize: 6.4, cellPadding: 1.25 },
    headStyles: { fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold" },
    alternateRowStyles: { fillColor: [248, 250, 252] },
    willDrawPage: () => {
      drawPdfHeader(
        doc,
        "Annexe — détail des élèves",
        `Session ${report.yearLabel} — notes sur 20`,
      );
    },
  });
}

function addSubsubjectStudentAppendix(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  report: BrevetPanoramaReportData,
) {
  if (report.subsubjects.length === 0) return;

  const french = report.subsubjects.filter((subject) => subject.group === "francais");
  const sciences = report.subsubjects.filter((subject) => subject.group === "sciences");
  doc.addPage("a4", "landscape");
  const title = "Annexe — détail français et sciences";
  const subtitle = `Session ${report.yearLabel} — notes brutes et barèmes officiels; « Absent » conservé`;
  autoTable(doc, {
    startY: 29,
    margin: { top: 28, left: 7, right: 7, bottom: 12 },
    theme: "grid",
    head: [[
      "Nom", "Prénom", "Classe", "Français /20",
      ...french.map(subsubjectHeader),
      "Sciences /20",
      ...sciences.map(subsubjectHeader),
    ]],
    body: report.studentRows.map((row) => [
      row.lastName,
      row.firstName,
      row.formerClass ?? "",
      pdfUnknownNumber(row.francais),
      ...french.map((subject) => subsubjectPdfValue(row[subject.studentKey])),
      pdfUnknownNumber(row.sciences),
      ...sciences.map((subject) => subsubjectPdfValue(row[subject.studentKey])),
    ]),
    styles: { ...pdfTableStyles(doc), fontSize: 7, cellPadding: 1.2 },
    headStyles: { fillColor: [37, 99, 235], textColor: 255, fontStyle: "bold" },
    alternateRowStyles: { fillColor: [248, 250, 252] },
    columnStyles: { 0: { cellWidth: 25 }, 1: { cellWidth: 25 }, 2: { cellWidth: 15 } },
    willDrawPage: () => drawPdfHeader(doc, title, subtitle),
  });
}

function addPdfFooters(
  doc: Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>,
  yearLabel: string,
  firstPage: number,
) {
  const pageCount = doc.getNumberOfPages();
  const font = getPdfFont(doc);
  for (let pageNumber = firstPage; pageNumber <= pageCount; pageNumber += 1) {
    doc.setPage(pageNumber);
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    doc.setDrawColor(226, 232, 240);
    doc.line(10, pageHeight - 8, pageWidth - 10, pageHeight - 8);
    doc.setFont(font, "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(100, 116, 139);
    doc.text(`Bilan complet DNB — ${yearLabel}`, 10, pageHeight - 4);
    doc.text(`Page ${pageNumber}/${pageCount}`, pageWidth - 10, pageHeight - 4, { align: "right" });
  }
}

export async function createBrevetPanoramaFullPdfDoc(
  report: BrevetPanoramaReportData,
): Promise<Awaited<ReturnType<typeof createBrevetPanoramaPdfDoc>>> {
  const doc = await createBrevetPanoramaPdfDoc(report.yearLabel, report.stats);
  const extended = asExtendedReport(report);
  const firstAdditionalPage = doc.getNumberOfPages() + 1;
  addClassComparisonPage(doc, report);
  addClassSubsubjectPage(doc, report);
  addSubjectAndRankingPage(doc, report);
  addGroupComparisonPage(
    doc,
    extended,
    extended.genderBreakdown,
    "Comparaison filles / garçons",
    "Filles",
    "Garçons",
  );
  addGroupComparisonPage(
    doc,
    extended,
    extended.scholarshipBreakdown,
    "Comparaison boursiers / non-boursiers",
    "Boursiers",
    "Non-boursiers",
  );
  addStudentListPage(
    doc,
    "Tous les élèves — ordre alphabétique",
    `Session ${report.yearLabel} — ${extended.alphabeticalRows?.length ?? report.studentRows.length} élèves; sans fiche appariée, sexe et bourse restent « ${UNKNOWN} »`,
    getAlphabeticalRows(extended),
    false,
  );
  addStudentListPage(
    doc,
    "Tous les élèves — classement par moyenne DNB",
    `Session ${report.yearLabel} — moyenne DNB décroissante; les absences de données restent « ${UNKNOWN} »`,
    getRankedRows(extended),
    true,
  );
  getClassDetails(extended).forEach((classDetail) => {
    const classSubtitle = `Session ${report.yearLabel} — ${classDetail.className}; informations manquantes : « ${UNKNOWN} »`;
    addStudentListPage(
      doc,
      `${classDetail.className} — ordre alphabétique`,
      classSubtitle,
      classDetail.alphabeticalRows,
      false,
    );
    addStudentListPage(
      doc,
      `${classDetail.className} — classement par moyenne DNB`,
      classSubtitle,
      classDetail.rankedRows,
      true,
    );
  });
  addStudentAppendix(doc, report);
  addSubsubjectStudentAppendix(doc, report);
  addPdfFooters(doc, report.yearLabel, firstAdditionalPage);
  return doc;
}

export async function exportBrevetPanoramaFullPdf(report: BrevetPanoramaReportData) {
  const doc = await createBrevetPanoramaFullPdfDoc(report);
  doc.save(`Bilan_complet_DNB_${safeFileName(report.yearLabel)}.pdf`);
}
