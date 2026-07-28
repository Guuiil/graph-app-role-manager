# Graph App Role Manager

[English version](README.md)

**Version multiplateforme actuelle : v1.1**

Graph App Role Manager est une interface graphique permettant de consulter, d'attribuer et
de supprimer les permissions d'application Microsoft Graph d'une identité managée ou d'un
Service Principal.

Le dépôt contient deux versions :

- une interface Windows native en PowerShell et Windows Forms ;
- une interface Python/Tkinter pour Windows, macOS et les environnements de bureau Linux.

Les permissions disponibles sont lues directement depuis le Service Principal Microsoft
Graph du tenant. Elles ne sont pas enregistrées dans une liste statique dans le code.

## Fonctionnalités

- Authentification interactive avec Microsoft Entra ID.
- Recherche des identités managées et Service Principals par nom d'affichage.
- Consultation et filtrage des permissions d'application Microsoft Graph actives.
- Affichage des App Roles Graph déjà attribués à l'identité sélectionnée.
- Attribution de plusieurs permissions en évitant les doublons.
- Suppression des permissions sélectionnées après confirmation explicite.
- Filtrage local de la liste des identités du tenant ou recherche par nom côté serveur.
- Conservation des permissions sélectionnées pendant le filtrage.
- Dialogue dédié au Device Code sous Windows pour éviter les erreurs de Window Handle WAM.
- Interface multiplateforme modernisée avec des contrôles et tableaux plus lisibles.
- Journal d'activité et récapitulatif de chaque opération.
- Aucun mot de passe, secret, certificat ou jeton d'accès enregistré sur le disque.

## Choisir une version

| Version | Usage conseillé | Authentification | Prérequis |
| --- | --- | --- | --- |
| [Interface PowerShell Windows](windows/Graph-App-Role-Manager.ps1) | Poste d'administration Windows | `Connect-MgGraph` interactif | Windows, PowerShell 7+, modules Microsoft Graph |
| [Interface Python multiplateforme v1.1](cross-platform/graph_app_role_manager.py) | Windows, macOS ou Linux | Device Code sous Windows ; navigateur interactif sous macOS/Linux | Python 3.10+, Tkinter, PowerShell 7, `Microsoft.Graph.Authentication` |

La version multiplateforme ne nécessite ni App Registration dédiée, ni Client ID, ni
secret, ni certificat, ni paquet Python supplémentaire.

## Permissions déléguées requises

La session de l'administrateur demande :

- `Application.Read.All`
- `AppRoleAssignment.ReadWrite.All`

Ces permissions sont sensibles. Elles nécessitent un consentement administrateur et un rôle
Microsoft Entra adapté. Consultez
[Permissions et sécurité](docs/permissions-and-safety.md) avant toute utilisation.

## Démarrage rapide

### PowerShell sous Windows

```powershell
pwsh -File .\windows\Graph-App-Role-Manager.ps1
```

L'interface peut proposer l'installation des modules Graph absents pour l'utilisateur
courant, mais demande une confirmation préalable.

### Version multiplateforme

```bash
pwsh -NoProfile -Command "Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
python cross-platform/graph_app_role_manager.py
```

Conservez `graph_app_role_manager.py`, `graph_backend.ps1` et
`graph-app-role-manager-icon.png` dans le même dossier. Le champ du tenant est facultatif.
Les prérequis par système et le dépannage se trouvent dans le
[guide multiplateforme](docs/cross-platform.md).

## Utilisation prudente

1. Connectez-vous avec un compte administrateur autorisé.
2. Recherchez l'identité cible.
3. Vérifiez son nom, son Object ID, son Application ID et son type.
4. Consultez les permissions existantes.
5. Sélectionnez uniquement les permissions nécessaires au scénario.
6. Relisez la confirmation avant toute attribution ou suppression.
7. Contrôlez le résultat dans l'outil et dans le centre d'administration Microsoft Entra.

## Périmètre

L'outil gère uniquement les **permissions d'application Microsoft Graph** représentées par
des App Role Assignments. Il ne gère pas le consentement délégué des utilisateurs, les rôles
d'annuaire Entra, Azure RBAC ou les autorisations de site SharePoint `Sites.Selected`.

## Documentation

- [Architecture](docs/architecture.md)
- [Installation Windows](docs/windows.md)
- [Installation multiplateforme](docs/cross-platform.md)
- [Permissions et sécurité](docs/permissions-and-safety.md)

Les captures d'écran pourront être ajoutées ultérieurement sans modifier la structure ou le
fonctionnement du projet.

## Licence

Projet distribué sous [licence MIT](LICENSE).
