import assert from "node:assert/strict";
import test from "node:test";
import * as XLSX from "./spreadsheet";

import type { ProcessedStudentData } from "@/contexts/FilterContext";
import type { BrevetPanoramaPdfStats } from "@/lib/brevet-panorama-export";
import { createBrevetPanoramaWorkbook } from "@/lib/brevet-panorama-full-export";
import { buildBrevetPanoramaReportData } from "@/lib/brevet-panorama-report";

const distribution = { gte15: 0, gte10lt15: 0, gte8lt10: 0, lt8: 0, count: 0 };
const stats: BrevetPanoramaPdfStats = {
  totalStudents: 3,
  admis: 2,
  refuse: 0,
  successRate: 100,
  mentions: { tresBien: 0, bien: 0, assezBien: 0, sansMention: 2 },
  mentionPercentages: { tresBien: 0, bien: 0, assezBien: 0, sansMention: 100 },
  scoreDistribution: {
    francais: { ...distribution },
    maths: { ...distribution },
    histoireGeo: { ...distribution },
    sciences: { ...distribution },
  },
};

const students: ProcessedStudentData[] = [
  { id: "1", ine: "1", nom: "Zulu", prenom: "Alice", etablissement: "", formerClass: "3e 2", sexe: "F", isBoursier: true, moyenne: 15, resultat: "ADMIS", scoreFrancais: 14, scoreFrancaisGrammaireComprehension: 29, scoreFrancaisDictee: 1.5, scoreFrancaisRedaction: 28, scoreSciences: 14, scoreSciencesSvt: 9, scoreSciencesPhysiqueChimie: 5 },
  { id: "2", ine: "2", nom: "Alpha", prenom: "Benoît", etablissement: "", formerClass: "3e 1", sexe: "G", isBoursier: false, moyenne: 12, resultat: "ADMIS", brevetBlancBb1Average: 10, brevetBlancBb2Average: 11, scoreFrancais: 10, scoreFrancaisGrammaireComprehension: "Absent", scoreFrancaisDictee: "Absent", scoreFrancaisRedaction: "Absent", scoreSciencesSvt: "Absent", scoreSciencesPhysiqueChimie: "Absent" },
  { id: "3", ine: "3", nom: "Bravo", prenom: "Chloé", etablissement: "", formerClass: "3e 1" },
];

test("le classeur contient les vues globales et par classe sans faux rang ni faux statut boursier", () => {
  const report = buildBrevetPanoramaReportData(students, "2026", stats);
  const workbook = createBrevetPanoramaWorkbook(report);
  assert.equal(workbook.SheetNames[0], "Menu");
  for (const name of ["Synthèse", "Par classe", "Par matière", "03 - Filles-Garçons", "04 - Boursiers", "Matrice classes", "Tout alpha", "Classement", "3e 1 - Alpha", "3e 1", "3e 2 - Alpha", "3e 2"]) {
    assert.ok(workbook.SheetNames.includes(name), `${name} absent`);
  }

  const alpha = XLSX.utils.sheet_to_json<Array<string | number>>(workbook.Sheets["Tout alpha"], { header: 1, defval: "" });
  const ranked = XLSX.utils.sheet_to_json<Array<string | number>>(workbook.Sheets["Classement"], { header: 1, defval: "" });
  const top10 = XLSX.utils.sheet_to_json<Array<string | number>>(workbook.Sheets["Top 10"], { header: 1, defval: "" });
  assert.equal(alpha[3][0], "N°");
  assert.equal(ranked[3][0], "Rang DNB");
  assert.deepEqual(alpha.slice(4).map((row) => row[1]), ["Alpha", "Bravo", "Zulu"]);
  assert.deepEqual(ranked.slice(4).map((row) => row[1]), ["Zulu", "Alpha", "Bravo"]);
  assert.deepEqual(ranked.slice(4).map((row) => row[0]), [1, 2, ""]);
  assert.equal(alpha[5][4], "Inconnu");
  assert.ok(!top10[3].includes("INE"));
  assert.ok(!alpha[3].includes("INE"));
  assert.ok(!ranked[3].includes("INE"));

  const classStatistics = XLSX.utils.sheet_to_json<Array<string | number>>(workbook.Sheets["Par classe"], { header: 1, defval: "" });
  assert.ok(classStatistics[3].includes("Boursiers"));
  assert.ok(classStatistics[3].includes("Bourse connue"));

  const details = XLSX.utils.sheet_to_json<Array<string | number>>(workbook.Sheets["Élèves"], { header: 1, defval: "" });
  assert.equal(details[3][3], "Ancienne classe");
  assert.equal(details[3][4], "Série");
  assert.ok(!details[3].includes("INE"));
  assert.ok(!details[3].includes("Boursier"));
  assert.ok(!details[3].includes("Source classe"));
  assert.equal(details[3].length, details[4].length);
  assert.ok(details[3].includes("Grammaire et compréhension /50"));
  assert.ok(details[3].includes("SVT /10"));
  assert.ok(details[3].includes("Physique-chimie /10"));
  assert.ok(!details[3].includes("Technologie /10"));
  const grammarColumn = details[3].indexOf("Grammaire et compréhension /50");
  assert.equal(details[4][grammarColumn], 29);
  assert.equal(details[5][grammarColumn], "Absent");
  assert.equal(workbook.Sheets["Élèves"].H5.s?.numFmt, "0.00");
  assert.equal(workbook.Sheets["Par classe"].K5.s?.numFmt, "0.00");
  assert.equal(workbook.Sheets["Top 10"].E5.s?.numFmt, "0.00");
  assert.equal(workbook.Sheets["Tout alpha"].G5.s?.numFmt, "0.00");
  const grammarAddress = XLSX.utils.encode_cell({ r: 4, c: grammarColumn });
  assert.equal(workbook.Sheets["Élèves"][grammarAddress].s?.numFmt, "0.00");
  const savedWorkbook = XLSX.read(XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }), {
    type: "buffer", cellNF: true, cellText: true,
  });
  assert.equal(savedWorkbook.Sheets["Élèves"].H5.z, "0.00");
  assert.equal(savedWorkbook.Sheets["Élèves"].H5.w, "15.00");

  const classAlpha = XLSX.utils.sheet_to_json<Array<string | number>>(workbook.Sheets["3e 1 - Alpha"], { header: 1, defval: "" });
  assert.deepEqual(classAlpha.slice(4).map((row) => row[1]), ["Alpha", "Bravo"]);
  assert.ok(!classAlpha[3].includes("INE"));
  assert.ok(!classAlpha[3].includes("Boursier"));
  assert.equal(workbook.Sheets["3e 1 - Alpha"].F5.s?.numFmt, "0.00");
  const classAlphaLastColumn = XLSX.utils.encode_cell({ r: 4, c: classAlpha[3].length - 1 });
  assert.equal(workbook.Sheets["3e 1 - Alpha"][classAlphaLastColumn].s?.numFmt, "0.00");
});

test("la colonne technologie apparaît seulement lorsqu'une note de technologie existe", () => {
  const report = buildBrevetPanoramaReportData([
    { id: "t", ine: "t", nom: "Test", prenom: "Tech", etablissement: "", scoreSciencesSvt: 6, scoreSciencesTechnologie: 8 },
  ], "2027", stats);
  const workbook = createBrevetPanoramaWorkbook(report);
  const details = XLSX.utils.sheet_to_json<Array<string | number>>(workbook.Sheets["Élèves"], { header: 1, defval: "" });
  assert.ok(details[3].includes("SVT /10"));
  assert.ok(details[3].includes("Technologie /10"));
  assert.ok(!details[3].includes("Physique-chimie /10"));
});
