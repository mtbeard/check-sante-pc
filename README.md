# 🖥️ Chek-Sante-PC

Script PowerShell qui génère un **rapport HTML complet de santé** de votre machine Windows — disques, RAM, batterie, mises à jour et programmes au démarrage — en un seul clic.

---

## ✨ Fonctionnalités

- **Espace disque** : utilisation par partition avec jauge colorée (vert / orange / rouge)
- **Mémoire RAM** : taux d'utilisation global + détail de chaque barrette (capacité, vitesse, fabricant)
- **Batterie** : santé en %, capacité actuelle vs capacité de conception, nombre de cycles de charge
- **Mises à jour Windows** : liste des mises à jour en attente avec sévérité et taille
- **Programmes au démarrage** : inventaire complet depuis le registre et WMI
- **Rapport HTML moderne** : interface responsive avec dégradé violet, jauges animées et code couleur intuitif

---

## 📋 Prérequis

- Windows 10 ou Windows 11
- PowerShell 5.1+ (inclus nativement)
- Droits administrateur recommandés (requis pour `powercfg /batteryreport` et Windows Update)

---

## 🚀 Utilisation

### Méthode 1 — Clic droit (recommandée)

1. Faites un **clic droit** sur `check-sante-pc.ps1`
2. Sélectionnez **"Exécuter avec PowerShell"**

### Méthode 2 — Terminal PowerShell

```powershell
# Depuis le dossier du script
.\check-sante-pc.ps1
```

> Si l'exécution est bloquée par la politique d'exécution, lancez d'abord :
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
> ```

---

## 📁 Sortie

Le rapport HTML est automatiquement enregistré dans :

```
%USERPROFILE%\Documents\SantePC\rapport-sante_YYYY-MM-DD_HH-mm.html
```

Le fichier s'ouvre automatiquement dans votre navigateur par défaut à la fin de l'exécution.

---

## 🎨 Aperçu du rapport

| Section | Indicateurs |
|---|---|
| Espace disque | Barre de progression, Go utilisé / total, % libre |
| RAM | Utilisation globale, détail des slots mémoire |
| Batterie | Santé %, cycles, capacité mWh, état de charge |
| Mises à jour | Nombre, sévérité (Critique / Important / Standard), taille |
| Démarrage | Nom, commande, source (Registre HKLM / HKCU / WMI) |

**Code couleur :**
- 🟢 Vert — situation normale (< 75 % d'utilisation / batterie > 85 %)
- 🟡 Orange — à surveiller (75–90 % / batterie 70–85 %)
- 🔴 Rouge — critique (> 90 % / batterie < 70 %)

---

## 🛡️ Confidentialité

Le script tourne **entièrement en local**. Aucune donnée n'est envoyée sur Internet. Les rapports HTML restent sur votre machine.

---

## 📄 Licence

MIT — libre d'utilisation, de modification et de distribution.
