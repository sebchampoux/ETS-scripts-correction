import { program } from 'commander';
import ExcelJs from 'exceljs';
import { getListeEtudiants } from './get-liste-etudiants.js';
import { stringify } from 'csv';
import { writeFile } from 'fs';

const { Workbook } = ExcelJs;

program.version('1.0.0');

program
	.requiredOption('-f, --fichier-evaluations <file>', 'Fichier Excel dans lequel sont les grilles d\'évaluation.')
	.requiredOption('-le, --liste-etudiants <file>', 'Fichier de liste d\'étudiants au format CSV, dans lequel seront enregistrées les notes.')
	.requiredOption('-cn, --cellule-note <cellule>', 'Cellule de la grille d\'évaluation dans laquelle se trouve la note de l\'étudiant.')
	.requiredOption('-ir, --index-resultat <index>', 'Index (à partir de 0) de la colonne du CSV dans laquelle doit être enregistrée la note de l\'étudiant pour ce travail.')
	.option('-ip, --index-prenom <index>', 'Index (à partir de 0) de la colonne du CSV dans laquelle est enregistré les prénoms.', '0')
	.option('-in, --index-nom <index>', 'Index (à partir de 0) de la colonne du CSV dans laquelle est enregistré les noms de famille.', '1');

program.parse(process.argv);

const options = program.opts();

(async () => {
	const {
		fichierEvaluations,
		listeEtudiants: fichierListeEtudiants,
		celluleNote,
		indexPrenom,
		indexNom,
		indexResultat,
	} = options;

	const listeEtudiants = await getListeEtudiants(fichierListeEtudiants);

	const evaluations = new Workbook();
	await evaluations.xlsx.readFile(fichierEvaluations);

	// Ignore rangée 0 car en-tête du CSV
	listeEtudiants.slice(1).forEach(etudiant => {
		const nomEtudiant = `${etudiant[indexPrenom]} ${etudiant[indexNom]}`;
		const grilleEtudiant = evaluations.worksheets.find(s => s.name === nomEtudiant);
		if (!grilleEtudiant) {
			throw new Error(`Pas trouvé la grille pour ${nomEtudiant}`);
			return;
		}
		const noteEtudiant = grilleEtudiant.getCell(celluleNote).value.result ?? 0;
		etudiant[indexResultat] = noteEtudiant;
	});
	const saveDataToFile = (error, output) => {
		writeFile(fichierListeEtudiants, output, () => {});
	};
	stringify(listeEtudiants, saveDataToFile);
	console.log(`Extrait les notes dans ${fichierListeEtudiants}`);
})()
