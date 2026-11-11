# Atelier Excel ONG

Application web pilotée par des pages JSP pour importer et archiver des classeurs, parcourir leurs feuilles, sélectionner des colonnes et exporter les résultats en Excel ou PDF.

## Fonctions

- Import `.xls`, `.xlsx`, `.xlsm`, `.xltx` et `.xltm` avec validation et limite de 50 Mo.
- Aperçu des colonnes, feuilles et données par pages de 100 lignes.
- Export `.xlsx` ou PDF paysage avec les colonnes choisies.
- Archive persistante des imports et exports, avec liste et téléchargement.

Les fichiers binaires `.xlsb` ne sont pas pris en charge par Apache POI. Les macros ne sont pas exécutées; un export personnalisé est produit en `.xlsx`.

## Prérequis et lancement

- JDK 21, Maven 3.9 ou plus récent et Apache Tomcat 10.1 (Jakarta Servlet 6). Tomcat 9 n'est pas compatible avec Jakarta Servlet 6.
- Depuis le dossier `ONGEze`, lancer `mvn clean package`, puis copier `target/excel-atelier.war` dans le dossier `webapps` de Tomcat. URL après démarrage: `http://localhost:8080/excel-atelier/`.

Les fichiers sont conservés dans `ong-excel-archive` du répertoire personnel du compte système Tomcat, indépendamment du WAR. Pour déplacer l’archive, ajouter dans `WEB-INF/web.xml` un paramètre de contexte nommé `storageRoot` et lui donner un chemin accessible en lecture/écriture au compte Tomcat. Sauvegarder ce répertoire régulièrement.

Cette version ne comporte pas de comptes ni de rôles. La déployer sur un réseau de confiance; une exposition publique nécessite notamment authentification, autorisations, protection CSRF et sauvegardes adaptées.