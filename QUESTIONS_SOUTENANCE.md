# Questions de Soutenance — AGSV (PFE)

> Préparation exhaustive couvrant : contexte, architecture, choix techniques,
> sécurité, PWA/offline, base de données, modélisation, limites et perspectives.

---

## 1. Contexte et justification du projet

- Pourquoi avoir choisi une architecture SaaS multi-tenant plutôt qu'une application dédiée à COOPRANORD ?
- COOPRANORD a deux sites (Korhogo et Bamako). Comment votre système gère-t-il la séparation des données entre dépôts tout en offrant une vue consolidée à la direction ?
- Quelles alternatives avez-vous envisagées avant de partir sur une solution full-custom ? (ERP existant, Odoo, etc.) Pourquoi les avez-vous écartées ?
- Vous dites que la gestion se faisait "sur cahier et Excel". Quels risques concrets cela posait-il, et comment votre solution y répond-elle précisément ?
- Comment avez-vous recueilli les besoins auprès de COOPRANORD ? Des entretiens, des ateliers ?

---

## 2. Architecture générale et choix techniques

- Pourquoi Spring Boot côté backend et React/Vite côté frontend ? Quels critères ont guidé ces choix ?
- Pourquoi MySQL plutôt que PostgreSQL ou une base NoSQL ?
- Vous avez séparé backend et frontend en deux services distincts. Quels avantages et inconvénients cela implique-t-il ?
- Comment le frontend communique-t-il avec le backend en développement ? Et en production ?
- Pourquoi utiliser Docker Compose ? Qu'est-ce que cela apporte concrètement pour ce projet ?
- Qu'est-ce que Vite apporte par rapport à Create React App ? Pourquoi ce choix ?

---

## 3. Modélisation et base de données

- Expliquez le choix de l'héritage **JOINED** pour l'entité `User` (AdminPrincipal, Administrator, Seller). Quelles sont les alternatives (SINGLE_TABLE, TABLE_PER_CLASS) et pourquoi avez-vous écarté ces options ?
- Votre héritage JOINED génère plusieurs tables liées. Comment une requête de connexion charge-t-elle le bon sous-type ? Qu'est-ce que `AccountService.getConcreteUser()` résout exactement ?
- Pourquoi les UUID sont-ils stockés en `CHAR(36)` et non en `BINARY(16)` (plus compact) ?
- La règle `ddl-auto=update` est utilisée. Est-ce une bonne pratique en production ? Quels risques cela comporte-t-il ?
- Comment le stock est-il géré par dépôt ? Y a-t-il une table intermédiaire `ProductDepot` ou une autre structure ?
- Comment modélisez-vous une vente avec plusieurs produits ? Décrivez le modèle entité-relation.
- La lettre de change a trois parties (tireur, bénéficiaire, tiré). Comment sont-elles modélisées en base ?
- Comment gérez-vous les colonnes d'enum avec Hibernate 6 et MySQL ? Pourquoi `VARCHAR(30)` et non l'enum MySQL natif ?

---

## 4. Sécurité

- Expliquez votre mécanisme d'authentification JWT de bout en bout (login → génération du token → validation sur chaque requête).
- Quelles informations sont stockées dans le JWT (`claims`) ? N'est-ce pas un risque de sécurité de mettre `companyId` et `depotId` dans le token ?
- Comment fonctionne votre protection anti-brute-force (`BruteForceProtectionService`) ?
- Les mots de passe sont-ils hachés ? Avec quel algorithme ? Pourquoi BCrypt plutôt que SHA-256 ?
- Vous avez un mot de passe par défaut `BienvenueSurAGSV`. Comment forcez-vous le changement à la première connexion ?
- Comment protégez-vous les endpoints selon les rôles ? Spring Security, annotations, ou contrôles manuels dans les contrôleurs ?
- Y a-t-il une gestion de l'expiration du token ? Que se passe-t-il si le token expire pendant une session active ?
- Comment évitez-vous les injections SQL ou les attaques CSRF dans votre application ?

---

## 5. Multi-tenancy

- Comment l'isolation des données entre entreprises est-elle garantie techniquement ? Est-ce au niveau de la base, du service, ou du contrôleur ?
- Un administrateur gérant plusieurs entreprises reçoit un token de sélection (`requiresCompanySelection`). Comment ce mécanisme fonctionne-t-il ?
- Qu'est-ce que la table `AdminCompany` ? À quoi sert-elle dans votre modèle ?
- Que se passe-t-il si un administrateur tente d'accéder aux données d'une entreprise à laquelle il n'est pas rattaché ?

---

## 6. Mode hors-ligne et PWA

- Expliquez votre architecture offline de bout en bout : qu'est-ce qui se passe quand un vendeur crée une vente sans connexion ?
- Pourquoi avoir utilisé IndexedDB plutôt que `localStorage` pour la file d'attente ?
- Votre synchronisation (`syncAll`) gère la résolution d'ID temporaires pour les clients créés hors-ligne. Expliquez ce mécanisme en détail. Que se passe-t-il si le sync échoue à mi-chemin ?
- Vous avez deux couches de cache : IndexedDB (manuel) ET le Service Worker Workbox. Ne font-ils pas doublon ? Pourquoi cette redondance ?
- Quelle est la différence entre `NetworkFirst` et `CacheFirst` dans Workbox ? Pourquoi avoir choisi `NetworkFirst` pour les produits et clients ?
- Que se passe-t-il si un produit change de prix côté serveur pendant que le vendeur est hors-ligne ? Comment gérez-vous les conflits à la synchronisation ?
- L'offline n'est implémenté que dans `VentePage`. Pourquoi cette limitation ? Que faudrait-il pour l'étendre aux inventaires et pertes ?
- Votre `vite-plugin-pwa` est configuré avec `registerType: 'autoUpdate'`. Qu'est-ce que cela signifie concrètement pour l'utilisateur ?

---

## 7. Modules métier — questions de fond

**Ventes :**
- Comment gérez-vous une vente avec paiement partiel ? Qu'est-ce qui est créé automatiquement en base ?
- Un vendeur peut-il faire une vente à crédit à un client occasionnel ? Pourquoi ou pourquoi pas ?
- Comment fonctionne la conversion d'une commande en vente ?

**Crédits / Créances :**
- Comment le statut "En retard" est-il déclenché ? Est-ce un job planifié (scheduler) ou un calcul à la volée ?
- Comment les alertes email de rappel de crédit sont-elles envoyées ? Quel service mail utilisez-vous ?

**Pertes et inventaires :**
- Pourquoi soumettre la perte d'un vendeur à validation avant de déduire le stock ? Quels risques métier cela couvre-t-il ?
- Comment évitez-vous qu'un vendeur vende un produit dont le stock est "en attente de déduction" suite à une perte non encore validée ?

**Approvisionnements :**
- Comment le stock est-il mis à jour après un approvisionnement ? Y a-t-il un mouvement de stock enregistré ?

**Rapports :**
- Vos rapports sont générés à la volée (requêtes en temps réel) ou précalculés ? Quelles sont les implications sur les performances ?
- Comment générez-vous le PDF d'une facture ?

---

## 8. Journal d'audit

- Pourquoi avoir mis le log d'audit dans un try-catch ? Qu'est-ce que cela implique si le log échoue ?
- Quelles sont les 40+ actions auditées ? Donnez des exemples représentatifs.
- Le journal est en lecture seule pour tous les utilisateurs, même l'AdminPrincipal. Comment cela est-il garanti techniquement ?

---

## 9. Performance et scalabilité

- Vous visez un chargement du tableau de bord en moins de 2 secondes. Comment le mesurez-vous ? Qu'avez-vous fait pour l'optimiser ?
- Avec un chargement paresseux (lazy loading) des associations JPA, comment évitez-vous les problèmes N+1 ?
- Pourquoi ajouter `@Transactional` sur les contrôleurs du package `products` ? N'est-ce pas une mauvaise pratique de mettre `@Transactional` dans la couche contrôleur ?
- Si COOPRANORD s'étend à 10 entreprises avec 50 dépôts, votre architecture tient-elle ? Quels seraient les premiers goulots d'étranglement ?

---

## 10. Tests et qualité

- Avez-vous écrit des tests unitaires ou d'intégration ? Si oui, qu'est-ce qui est couvert ?
- Comment avez-vous testé le mode hors-ligne ? Et la synchronisation ?
- Avez-vous effectué des tests utilisateurs avec COOPRANORD ? Quels retours avez-vous eu ?

---

## 11. Déploiement et DevOps

- Votre `docker-compose.yml` démarre MySQL, MailHog et les deux services. Est-ce la configuration de prod ou seulement de dev ?
- Comment géreriez-vous les secrets (mots de passe, clé JWT) en production ? Vous utilisez un fichier `.env` — comment le sécuriseriez-vous ?
- Avez-vous pensé à la haute disponibilité ? Que se passe-t-il si le serveur backend tombe ?
- Comment mettriez-vous à jour l'application en production sans interruption de service ?

---

## 12. Limites et perspectives

- Quelles sont les trois principales limitations techniques de votre solution actuelle ?
- Vous n'avez pas de route guards côté frontend — n'importe quel utilisateur authentifié peut accéder à n'importe quelle URL. Pourquoi ce choix ? Quels risques cela comporte-t-il ?
- Les fichiers de rapports exportés ne sont pas stockés (seul l'historique est enregistré, pas les bytes). Pourquoi ? Comment rééditeriez-vous une facture ancienne si le serveur redémarre ?
- Si vous aviez deux semaines supplémentaires, quelle fonctionnalité ajouteriez-vous en priorité ?
- Comment passeriez-vous votre application de SaaS à une version on-premise pour un client qui ne veut pas être hébergé dans le cloud ?

---

## 13. Questions de recul général

- Quel a été le module le plus difficile à développer et pourquoi ?
- Avec le recul, quel choix technique ferait-il autrement ?
- Quelle est la différence entre votre AGSV et un ERP standard comme Odoo ? Pourquoi ne pas avoir utilisé Odoo directement ?
- Comment garantissez-vous que les données de COOPRANORD sont conformes au RGPD ou aux réglementations locales de Côte d'Ivoire ?
- Si une autre coopérative agricole voulait utiliser AGSV, que faudrait-il paramétrer ou adapter ?