# Politique de Confidentialité — Octano

**Dernière mise à jour : 5 juillet 2026**

Octano est une application iOS gratuite, maintenue par un développeur indépendant, qui aide les conducteurs à choisir où faire le plein à partir de données publiques officielles. Cette politique explique, conformément à la Loi Générale de Protection des Données du Brésil (Loi n° 13.709/2018 — LGPD), quelles données sont traitées et comment.

**Responsable du traitement :** Luis Fontinelles — contact : **[fontinelles.com/contact](https://fontinelles.com/contact)**

## Le résumé honnête

Octano **n'a pas de compte, pas de connexion, pas de publicité et ne vend pas de données**. Le contenu que vous créez (prix, votes) reste sur votre appareil. Deux éléments quittent l'appareil : votre **localisation approximative**, envoyée à Google pour trouver des stations proches, et des **données d'usage anonymes**, envoyées à Firebase (Google) pour comprendre comment l'app est utilisée et l'améliorer. Les achats de soutien facultatifs sont traités par **Apple** — nous ne voyons jamais les données de votre carte.

## 1. Données traitées par l'app

### 1.1 Localisation (avec votre permission)
- **Pourquoi :** trouver des stations proches, trier la liste par distance et tracer des itinéraires.
- **Comment :** traitée sur l'appareil. Elle est envoyée **uniquement** à l'API Google Places (recherche de stations) et à Apple (MKDirections, calcul d'itinéraire), qui agissent comme sous-traitants de ces fonctions.
- **Base légale :** consentement (vous l'autorisez dans iOS et pouvez le révoquer à tout moment dans Réglages → Confidentialité → Localisation).
- L'app **ne** stocke **pas** d'historique de localisation sur un serveur du développeur.

### 1.2 Données d'usage et de diagnostic (analytics)
- **Quoi :** événements de navigation (écrans visités, temps passé sur chaque écran, appuis sur des boutons comme rechercher, démarrer un itinéraire, ouvrir une station), modèle de l'appareil et version d'iOS, région approximative (à partir de l'IP), et un **identifiant pseudonyme de l'app** généré par Firebase. Nous collectons aussi des rapports de **plantage (crash)** pour corriger les erreurs.
- **Pourquoi :** comprendre comment l'app est utilisée, prioriser les améliorations et corriger les problèmes. Nous **ne** l'utilisons **pas** à des fins publicitaires et **ne** collectons **pas** l'identifiant publicitaire (IDFA).
- **Qui traite :** **Firebase (Google LLC)** — Analytics, Crashlytics, Performance. Les données sont traitées selon la politique de Google.
- **Base légale :** intérêt légitime à améliorer le service ; les données sont pseudonymisées et agrégées.
- **Notifications :** si vous acceptez les notifications, un jeton Firebase Messaging est utilisé uniquement pour les livrer.

### 1.3 Achats intégrés (contributions facultatives)
- L'app propose des **contributions uniques** (« pourboires ») pour soutenir le projet. Elles ne déverrouillent aucune fonctionnalité.
- Le **paiement est traité par Apple** (App Store). Le développeur reçoit uniquement la confirmation de la transaction — **jamais** votre numéro de carte ni de données financières.

### 1.4 Données que l'app NE collecte PAS
Nom, e-mail, téléphone, contacts, photos, données de carte et identifiant publicitaire (IDFA). Il n'y a ni connexion ni compte.

## 2. Services tiers

| Service | Ce qu'il reçoit | Politique |
|---|---|---|
| Google Maps / Places (Google LLC) | Localisation approximative lors des recherches ; télémétrie propre au SDK cartographique | [policies.google.com/privacy](https://policies.google.com/privacy) |
| Firebase (Google LLC) | Événements d'usage, diagnostics/plantages et un identifiant pseudonyme de l'app | [firebase.google.com/support/privacy](https://firebase.google.com/support/privacy) |
| Apple (App Store / StoreKit) | Traitement des achats de soutien | [apple.com/privacy](https://www.apple.com/privacy/) |
| Apple (MapKit/MKDirections) | Coordonnées d'origine/destination pour le calcul d'itinéraires | [apple.com/privacy](https://www.apple.com/privacy/) |

## 3. Données publiques affichées par l'app

L'app reproduit des bases **publiques et officielles** de l'ANP (inspection, registre des revendeurs et prix), des **Procons régionaux** et du système national **SINDEC/Senacon** (entreprises verbalisées et réclamations fondées), de l'**IPEM-SP** (pompes certifiées) et du contenu de Google (avis, notes), toujours avec la source et la date.

- Les données d'**entreprises** (CNPJ, raison sociale) sont affichées telles qu'elles figurent dans les bases publiques.
- Les **CPF** (identifiant fiscal des personnes physiques) des revendeurs individuels, lorsqu'ils sont présents dans les bases, sont **masqués** par l'app et le pipeline de données — le document complet n'est pas rediffusé.
- Les bases de protection du consommateur contiennent des entreprises (personnes morales), collectées depuis les portails de transparence et les données ouvertes (Loi d'Accès à l'Information).
- Les avis affichés sont rédigés par des utilisateurs de Google et restent sous la responsabilité de leurs auteurs et de la plateforme d'origine.

## 4. Vos droits (art. 18 de la LGPD)

Vous pouvez demander la confirmation du traitement, l'accès, la correction ou la suppression des données via notre canal de contact : **[fontinelles.com/contact](https://fontinelles.com/contact)**. En pratique : **désinstaller l'app supprime les données que vous avez créées sur l'appareil** ; la permission de localisation est révocable dans les Réglages iOS ; et vous pouvez désactiver la personnalisation publicitaire/le suivi iOS à tout moment.

Si vous êtes une **personne concernée par des données présentes dans les bases publiques** affichées (par ex. un revendeur personne physique), vous pouvez nous contacter pour réviser l'affichage — et, pour les bases d'origine, exercer vos droits auprès de l'ANP, du Procon ou de l'IPEM.

## 5. Enfants et adolescents

L'app est destinée aux conducteurs titulaires du permis et n'est pas destinée aux mineurs.

## 6. Modifications

Cette politique peut être mise à jour ; la version en vigueur sera toujours à cette adresse, avec la date en haut.
