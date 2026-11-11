# MMC Batch

Application de bureau (Tauri + Vue) qui convertit des classeurs Excel en fichiers d'import :

- **Standard** : classeur déjà au format `MSISDN,<numéro>,<Paramètre>,…` → un CSV par feuille avec
  l'entête `HDR,UpdateIndividualRatingParameter,…` (style **Init** ou **Set**).
- **CRPLMT** : une colonne de numéros (avec un titre) → fichier `AddAlias` `.txt` :

  ```
  HDR,"AddAlias","Add Alias Title","ExtId12340","<nb>","Agent  advance pilote","1"
  MSISDN,"237653282055","CRPLMT_237653282055@100000"
  ```

  Montant, titre et description sont modifiables dans l'interface.
- **Migration** : une colonne de numéros (avec un titre) → fichier `SetProfile` `.csv` :

  ```
  HDR,"SetProfile","SetProfile Title","ExtId20240724","<nb>","SetProfile rollback","1"
  MSISDN,237671598153,<Profil>,<Raison>
  ```

  Profil et raison peuvent rester vides (`MSISDN,237671598153,,`) ; ils ne doivent contenir ni
  virgule ni guillemet. Un numéro à 9 chiffres reçoit l'indicatif `237`. Titre et description
  de l'en-tête sont modifiables.

L'interface affiche un aperçu des premières lignes du fichier produit avant le dépôt.

Le résultat est un zip, gardé dans l'historique (retéléchargeable depuis le panneau de gauche).

## Organisation

| Dossier | Rôle |
|---|---|
| `BackendRust/` | Bibliothèque Rust `mmc_batch_core` : lecture Excel, conversion, zip, historique SQLite. Contient aussi un serveur HTTP optionnel. |
| `Frontend/` | Interface Vue 3 + Tailwind. |
| `Frontend/src-tauri/` | Appli Tauri : appelle directement `mmc_batch_core` (commandes `convertir`, `lister_historique`, `lire_zip_historique`, `supprimer_historique`). |

La conversion tourne **dans l'appli elle-même** : il n'y a plus de backend séparé à compiler,
copier ou lancer, ni de port réseau.

Les données (historique et zips) sont dans `%APPDATA%\com.mmc.batch`.

## Lancer l'appli

```powershell
cd Frontend
npm install
npm run tauri dev      # développement
npm run tauri build    # installateur (.msi / .exe) dans C:\cargo-build\mmc-batch-tauri\release\bundle
```

## Travailler sur l'interface dans un navigateur

```powershell
cd BackendRust
cargo run              # serveur HTTP sur http://127.0.0.1:8000
```

puis, dans un autre terminal :

```powershell
cd Frontend
npm run dev
```

## Tests

```powershell
cd BackendRust
cargo test
```
