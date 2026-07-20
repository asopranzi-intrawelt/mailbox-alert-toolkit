# Work-log

> Append-only, ordine cronologico inverso.

## 2026-06-18 - Popolamento schede context e ancoraggio a HEAD

Eseguita sync-context su repo appena inizializzato. Tutte e sei le schede di `.claude/context/` erano template vuoti: popolate da zero leggendo il codice (`mailbox-alert.ps1`, `launcher.ps1`, `generate_mailbox_report.py`, `generate_trends.py`, `config.example.json`, `README.md`). Schede ancorate: STACK.md, design-and-security.md, deployment.md, dev-testing.md, current-work.md, roadmap.md. `memory/index.md` compilato con stato e punto di ripresa. Commit di riferimento: a8eab4c.

## 2026-06-18 - Adozione standard e identita git

Adottato lo standard portabile (PROJECT-SYSTEM, rules, engine skills + onboard, bundle/PACKAGES, schede context/memory scaffold) in .claude, ora TRACCIATO (.gitignore esclude solo settings.local.json e memory/*.local.md). Identita locale asopranzi/asopranzi@intrawelt + github-corp + OpenSSH; remoto switchato da HTTPS a SSH github-corp. Igiene: credentials/config.json/cert mai committati. Creato CLAUDE.md. Schede da ancorare con sync-context (HEAD 285943e) e popolare.

