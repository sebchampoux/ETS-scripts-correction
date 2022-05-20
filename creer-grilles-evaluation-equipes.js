import { program } from 'commander';
import ExcelJs from 'exceljs';

const { Workbook } = ExcelJs;

program.version('1.0.0');

program
	.requiredOption('-f, --fichier-evaluations <file>', 'Fichier Excel dans lequel seront créées les grilles d\'évaluation.')
	.requiredOption('-n, --nombre-equipes <nbr>', 'Nombre d\'équipes. Des grilles d\'évaluation seront créées pour les équipes 1 à n.')
	.option('-d, --debut <nbr>', 'Numéro de la première équipe si la première équipe n\'est pas la #1.', 1)
	.option('-e, --exclure-equipes <nbrs>', 'Numéros d\'équipes inutilisés. Écrire comme une string, séparer les numéros par une virgule, ex. "1, 5, 7".', '')
	.option('-t, --template-name <name>', 'Nom de la worksheet qui sera utilisée comme modèle pour les grilles d\'évaluation.', '__TEMPLATE__')
	.option('-g, --numero-groupe <groupe>', 'Numéro du groupe', null)
	.option('-c, --cellule-groupe <cellule>', 'Identifiant de la cellule dans lequel indiquer le numéro du groupe. ex. "A3"', null);

program.parse(process.argv);

const options = program.opts();

(async () => {
	const {
		celluleGroupe,
		numeroGroupe,
		fichierEvaluations,
		templateName,
		exclureEquipes,
		debut,
		nombreEquipes: nbrEquipesString
	} = options;

	const evaluations = new Workbook();
	await evaluations.xlsx.readFile(fichierEvaluations);
	const template = evaluations.getWorksheet(templateName);
	const equipesExclues = exclureEquipes.split(',').map(n => parseInt(n, 10));
	const nombreEquipes = parseInt(nbrEquipesString, 10);

	for (let i = debut; i <= (debut + nombreEquipes); i++) {
		if (equipesExclues.includes(i)) {
			continue;
		}
		const nomEquipe = `Équipe ${i}`;
		const worksheetEquipe = evaluations.addWorksheet(nomEquipe);
		worksheetEquipe.model = template.model;
		worksheetEquipe.name = nomEquipe;

		if (celluleGroupe && numeroGroupe) {
			worksheetEquipe.getCell(celluleGroupe).value = `Groupe ${numeroGroupe}`;
		}
	}
	await evaluations.xlsx.writeFile(fichierEvaluations);
	console.log(`Créé les grilles d'évaluation pour ${nombreEquipes} équipes avec succès.`);
})()