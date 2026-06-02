# Skill backup — multifamily-financial-review

`multifamily-financial-review.SKILL.md` is a backup of the **edited** skill that
includes the owner "House Rules" (see `../REVIEW_METHODOLOGY.md` for the same rules
in standalone form).

## Live (installed) location

The skill the app actually loads lives in the workspace skills directory:

```
C:\Users\daisy\AppData\Roaming\Claude\local-agent-mode-sessions\skills-plugin\c0ce1413-1c06-4211-8997-60a80c105fd1\6754e018-63ae-48be-a04d-b0e1593c0d89\skills\multifamily-financial-review\SKILL.md
```

This is the real installed copy (it sits alongside docx/pdf/pptx and the other
anthropic-skills), so edits here are picked up by future sessions in this workspace.

## Restore (if a marketplace re-sync ever overwrites the live skill)

Re-copy the backup over the live file (PowerShell):

```powershell
Copy-Item `
  "C:\Users\daisy\Projects\wilson-capital-financial\skill-backup\multifamily-financial-review.SKILL.md" `
  "C:\Users\daisy\AppData\Roaming\Claude\local-agent-mode-sessions\skills-plugin\c0ce1413-1c06-4211-8997-60a80c105fd1\6754e018-63ae-48be-a04d-b0e1593c0d89\skills\multifamily-financial-review\SKILL.md" `
  -Force
```

> The `c0ce1413…` GUID is this workspace's ID. If it changes, find the live path with:
> `Get-ChildItem "$env:APPDATA\Claude\local-agent-mode-sessions\skills-plugin" -Recurse -Filter SKILL.md | ? FullName -like *multifamily*`
