---
agent: speckit.implement
---

# RALPH LOOP PROTOCOL v3.0 - INTEGRATION OBLIGATOIRE

**REFERENCE COMPLETE**: Voir `RALPH_LOOP_PROTOCOL_V3.md` a la racine du projet.

## PRINCIPE FONDAMENTAL

Ce prompt DOIT etre execute avec le Ralph Loop Protocol v3.0:

```
WHILE conformite < 100%
{
    1. LIRE la specification INTEGRALEMENT
    2. IMPLEMENTER selon les requis
    3. BUILD (si echec: corriger et rebuild)
    4. ANALYSER: comparer OUTPUT vs SPECIFICATION
    5. Si ecarts detectes: CORRIGER et recommencer
}
```

## REGLES ABSOLUES

1. **NE JAMAIS** dire "termine" si conformite < 100%
2. **BUILD SUCCESS** n'est PAS suffisant - CONFORMITE TOTALE est le but
3. **TOUJOURS** comparer implementation vs specification (pas vs intuition)
4. **Si blocage technique**: DECLARER explicitement pour escalade

## BUILD COMMAND (XNRGY)

```powershell
cd XnrgyEngineeringAutomationTools
.\build-and-run.ps1 -BuildOnly
```

NEVER use `dotnet build` (breaks WPF .g.cs files)
