import { program } from 'commander';
import ExcelJS from 'exceljs';
const { Workbook } = ExcelJS;

program.version('1.0.0');

program
	.requiredOption('-f, --fichier-evaluations <fichier>', 'Fichier Excel contenant les grilles d\'évaluation des équipes.')
	.requiredOption('-c, --cellule-note <cellule>', 'Identifiant de la cellule contenant la note de l\'équipe, ex. A1.');

program.parse(process.argv);
const options = program.opts();

(async () => {
	const { fichierEvaluations, celluleNote } = options;

	const evaluations = new Workbook();
	await evaluations.xlsx.readFile(fichierEvaluations);
	evaluations.worksheets.forEach(grilleEquipe => {
		const celluleResultat = grilleEquipe.getCell(celluleNote);
		if (celluleResultat.value) {
			const noteEquipe = grilleEquipe.getCell(celluleNote).value.result ?? 0;
			console.log(`${grilleEquipe.name} : ${noteEquipe}`);
		}
	});
})()
