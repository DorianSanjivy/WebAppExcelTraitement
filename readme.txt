Excel'AI : Solution de webapp pour l'analyse de sentiment sur des commentaires en utilisant un fichier Excel
[IMPORTANT : Le fichier Excel doit être au format .xlsx et contenir une colonne "Review" avec les commentaires à analyser]

main.py : Application flask qui gère le backend du site sur le port 5000.
Gère :
-les requêtes d'api
-la génération de graphs
-le remplissage du fichier Excel

upload.html : page web frontend,
Gère :
-L'affichage htlm du site
-Le Css avec l'aide de Bootstrap
-Le javascript permettant de télécharger le fichier Excel traité et afficher les graphs


requirements.txt : liste des modules python à installer
```