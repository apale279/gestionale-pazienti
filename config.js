// Configurazione dell'app. Non contiene dati dei pazienti.
export const CONFIG = {
  // "Application (client) ID" ottenuto registrando l'app su Microsoft (vedi GUIDA.md).
  // Se lasci vuoto, l'app te lo chiede al primo avvio.
  clientId: "607cc1c1-3d7a-4165-9e4b-40fba2019238",

  // Percorso della cartella pazienti DENTRO OneDrive (senza "OneDrive/" davanti).
  rootPath: "Lavoro_medico/Pazienti",

  // Cartella (dentro rootPath) dove l'app salva le proprie impostazioni.
  appFolder: "_Gestionale",

  // Cartelle da non considerare pazienti (confronto senza maiuscole/minuscole).
  // Le cartelle che iniziano con "_" o "." sono sempre ignorate, tranne ".Pazienti".
  ignoredFolders: ["GestionePazienti", "Referti assistenze", "node_modules"],

  // Cartelle speciali che contengono altre cartelle-paziente.
  extraPatientContainers: [".Pazienti"],

  docTypes: ["Visita", "Referto", "Certificato", "Anamnesi", "Relazione", "Ricetta", "Consenso", "Lettera", "Nota"],
};
