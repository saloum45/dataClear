const XLSX = require("xlsx");
const fs = require("fs");

// Fonction pour nettoyer les prix
const nettoyerPrix = (prix) => {
  // Supprimer tout sauf les chiffres
  const chiffres = prix.replace(/[^\d]/g, "");
  return chiffres ? parseInt(chiffres, 10) : 0;
};

// Fonction pour traiter le fichier Excel
const traiterFichierExcel = (filePath, outputJsonPath) => {
  try {
    // Charger le fichier Excel
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Utiliser la première feuille
    const sheet = workbook.Sheets[sheetName];

    // Convertir la feuille en JSON brut
    const rawData = XLSX.utils.sheet_to_json(sheet);

    // Créer un tableau d'objets avec les colonnes formatées
    // const tableauObjets = rawData.map((row) => ({
    //   libelle: row["libelle"],
    //   prix_achat: nettoyerPrix(row["prix_achat"]),
    //   prix_vente: nettoyerPrix(row["prix_vente"]),
    // }));
    const tableauObjets = rawData
      .filter(
        (row) =>
          row["libelle"] && // Ignorer les lignes où "libelle" est vide
          row["prix_achat"] && // Ignorer les lignes où "prix_achat" est vide
          row["prix_vente"] // Ignorer les lignes où "prix_vente" est vide
      )
      .map((row) => ({
        libelle: row["libelle"],
        prix_achat: nettoyerPrix(row["prix_achat"]),
        prix_vente: nettoyerPrix(row["prix_vente"]),
      }));


    // Écrire le tableau d'objets dans un fichier JSON
    fs.writeFileSync(outputJsonPath, JSON.stringify(tableauObjets, null, 2), "utf-8");
    console.log(`Fichier JSON généré avec succès : ${outputJsonPath}`);
  } catch (error) {
    console.error("Une erreur s'est produite :", error.message);
  }
};

// Chemin du fichier Excel et du fichier JSON
const filePath = "./data/laborex_donnee.xlsx"; // Remplacez par le chemin de votre fichier Excel
const outputJsonPath = "laborex.json"; // Nom du fichier JSON de sortie

// Appeler la fonction pour traiter le fichier
traiterFichierExcel(filePath, outputJsonPath);
