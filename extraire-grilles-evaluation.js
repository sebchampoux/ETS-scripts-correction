import { program } from 'commander';
import slugify from 'slugify';
import ExcelJs from 'exceljs';

const { Workbook } = ExcelJs;

program.version('1.0.0');

program
	.requiredOption('-f, --fichier-evaluations <file>', 'Fichier Excel dans lequel seront extraites.')
	.option('-t, --template-name <name>', 'Nom de la worksheet template, qui sera ignorée.', 'Sheet1')
	.option('-d, --destination <folder_name>', 'Dossier dans lequel seront enregistrées les grilles d\'évaluation extraites. Ce dossier doit déjà exister.', './grilles/');

program.parse(process.argv);

const options = program.opts();

(async () => {
	const {
		fichierEvaluations,
		templateName,
		destination,
	} = options;

	const evaluations = new Workbook();
	await evaluations.xlsx.readFile(fichierEvaluations);

	evaluations.worksheets
		.filter(s => s.name !== templateName)
		.forEach(async (sheet) => {
			const nomEtudiant = sheet.name;
			const nouveauWorkbook = new Workbook();
			const grilleCopiee = nouveauWorkbook.addWorksheet(nomEtudiant);
			grilleCopiee.model = sheet.model;
			grilleCopiee.name = sheet.name;
			await nouveauWorkbook.xlsx.writeFile(`${destination}${slugify(nomEtudiant)}.xlsx`);
		});
	console.log('Grilles extraites avec succès!');
})()
