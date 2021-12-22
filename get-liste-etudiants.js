import { createReadStream } from 'fs';
import { parse } from 'csv';

export const getListeEtudiants = async (fichierListeEtudiants) => {
	const result = [];
	const parser = createReadStream(fichierListeEtudiants).pipe(parse());
	for await (const etudiant of parser) {
		result.push(etudiant);
	}
	return result;
};
