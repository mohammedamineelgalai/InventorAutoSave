---
agent: speckit.analyze
---

# RALPH LOOP PROTOCOL v3.0 - ANALYSE CRITIQUE OBLIGATOIRE

**REFERENCE COMPLETE**: Voir `RALPH_LOOP_PROTOCOL_V3.md` a la racine du projet.

## OBJECTIF DE L'ANALYSE

L'analyse est le **COEUR** du Ralph Loop Protocol. Elle compare ce qui a ete 
**REELLEMENT produit** avec la **SPECIFICATION** pour calculer la conformite.

## PROCEDURE D'ANALYSE

### 1. Relire la specification INTEGRALEMENT
- Identifier TOUS les requis (controles, bindings, noms, valeurs)
- Compter le nombre total de requis

### 2. Verifier CHAQUE requis
Pour chaque requis de la specification:
- [ ] Est-il implemente?
- [ ] L'implementation est-elle CORRECTE?
- [ ] Les noms sont-ils EXACTS?
- [ ] Les valeurs par defaut sont-elles correctes?

### 3. Detecter les problemes
- Exigences manquantes
- Implementations partielles
- Approximations
- Interpretations incorrectes
- Noms differents de la specification

### 4. Calculer le score
```
SCORE = (REQUIS_CONFORMES / REQUIS_TOTAL) * 100%
```

### 5. Generer le rapport de conformite

```
╔═══════════════════════════════════════════════════════════════════════════════╗
║  RAPPORT DE CONFORMITE - STEP_XX                                              ║
╠═══════════════════════════════════════════════════════════════════════════════╣
║  ITERATION: X                                                                 ║
║  BUILD STATUS: SUCCESS/FAILED                                                 ║
║                                                                               ║
║  CHECKLIST DES REQUIS:                                                        ║
║  [OK/FAIL] Requis 1                                                          ║
║  [OK/FAIL] Requis 2                                                          ║
║  ...                                                                          ║
║                                                                               ║
║  SCORE: XX/YY = ZZ%                                                          ║
║                                                                               ║
║  ECARTS DETECTES:                                                            ║
║  - Ecart 1: Description + Action corrective                                  ║
║  - Ecart 2: Description + Action corrective                                  ║
║                                                                               ║
║  DECISION:                                                                    ║
║  [ ] CONTINUER (score < 100%) -> /speckit.fix                                ║
║  [ ] TERMINER (score = 100%) -> Rapport final                                ║
║  [ ] ESCALADER (blocage technique) -> Declarer le blocage                    ║
╚═══════════════════════════════════════════════════════════════════════════════╝
```

## REGLES ABSOLUES

1. **ANALYSER ce qui a ete REELLEMENT produit** (pas ce qu'on pense avoir produit)
2. **COMPARER point par point** avec la specification
3. **NE JAMAIS** declarer termine si score < 100%
4. **Si blocage technique**: declarer explicitement pour escalade
