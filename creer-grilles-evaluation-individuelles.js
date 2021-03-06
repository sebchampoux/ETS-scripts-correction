import { program } from 'commander';
import ExcelJs from 'exceljs';
import { getListeEtudiants } from './utils/get-liste-etudiants.js';

const { Workbook } = ExcelJs;

program.version('1.0.0');

program
	.requiredOption('-f, --fichier-evaluations <file>', 'Fichier Excel dans lequel seront créées les grilles d\'évaluation.')
	.requiredOption('-le, --liste-etudiants <file>', 'Fichier de liste d\'étudiants au format CSV, tel que généré par Moodle.')
	.option('-t, --template-name <name>', 'Nom de la worksheet qui sera utilisée comme modèle pour les grilles d\'évaluation.', 'Sheet1')
	.option('-g, --numero-groupe <groupe>', 'Numéro du groupe', null)
	.option('-c, --cellule-groupe <cellule>', 'Identifiant de la cellule dans lequel indiquer le numéro du groupe. ex. "A3"', null)
	.option('-ce, --cellule-etudiant <cellule>', 'Identifiant de la cellule dans lequel indiquer le nom de l\'étudiant(e). ex. "A3"', null)
	.option('-ip, --index-prenom <index>', 'Index (à partir de 0) de la colonne du CSV dans laquelle est enregistré les prénoms.', '0')
	.option('-in, --index-nom <index>', 'Index (à partir de 0) de la colonne du CSV dans laquelle est enregistré les noms de famille.', '1');

program.parse(process.argv);

const options = program.opts();

(async () => {
	const {
		fichierEvaluations,
		listeEtudiants: fichierListeEtudiants,
		templateName,
		numeroGroupe,
		celluleGroupe,
		celluleEtudiant,
		indexPrenom,
		indexNom,
	} = options;

	const evaluations = new Workbook();
	await evaluations.xlsx.readFile(fichierEvaluations);
	const template = evaluations.getWorksheet(templateName);

	const listeEtudiants = await getListeEtudiants(fichierListeEtudiants);

	// Ignore rangée 0 parce que c'est le header du CSV
	listeEtudiants.slice(1).forEach(etudiant => {
		const nomEtudiant = `${etudiant[indexPrenom]} ${etudiant[indexNom]}`;

		const worksheetEquipe = evaluations.addWorksheet(nomEtudiant);
		worksheetEquipe.model = template.model;
		worksheetEquipe.name = nomEtudiant;

		if (celluleGroupe && numeroGroupe) {
			worksheetEquipe.getCell(celluleGroupe).value = `Groupe ${numeroGroupe}`;
		}
		if (celluleEtudiant) {
			worksheetEquipe.getCell(celluleEtudiant).value = nomEtudiant;
		}
	});

	await evaluations.xlsx.writeFile(fichierEvaluations);
	console.log(`Créé les grilles d'évaluation pour ${listeEtudiants.length - 1} étudiants avec succès.`);
})()
