# rapports-pedagogiques-clp
Un programme pour générer un PDF permettant de présenter de manière visuelle les rapports pédagogiques.
Pour l'utiliser ajouter un fichier `rapports.xlsx` à la racine.

Après avoir installer nodejs, et copier ce fichier, il suffit d'éxécuter: 
```
npm install
npm start
```

## Pour faire fonctionner sous MacOS

Installer libreoffice et suivre [ce tutoriel](https://gist.github.com/pankaj28843/3ad78df6290b5ba931c1).

Il faut en fait avoir `soffice` installé pour convertir les fichiers Word en PDF.

Si vous n'y arriver pas vous pouvez toujours modifier le fichier `index.js` (voir commentaires) pour générer des `.docx`.