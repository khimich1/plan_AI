# Cursor sessionStart: РЅР°РїРѕРјРёРЅР°РЅРёРµ РїСЂРѕ agent-skills
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
if (-not $root) { $root = (Get-Location).Path }
$meta = Join-Path $root ".cursor\skills\using-agent-skills\SKILL.md"
$project = Join-Path $root ".cursor\skills\project-shishov\SKILL.md"

$lines = @(
    "agent-skills Р·Р°РіСЂСѓР¶РµРЅС‹ РґР»СЏ РїСЂРѕРµРєС‚Р° РЁРёС€РѕРІ.",
    "РЎРЅР°С‡Р°Р»Р° СѓС‡С‚Рё project-shishov, Р·Р°С‚РµРј РІС‹Р±РµСЂРё СЃРєРёР»Р» РїРѕ Р±Р»РѕРє-СЃС…РµРјРµ using-agent-skills.",
    "Р СѓРєРѕРІРѕРґСЃС‚РІРѕ РЅР° СЂСѓСЃСЃРєРѕРј: .cursor/РЎРљРР›Р›Р«-Р РЈРљРћР’РћР”РЎРўР’Рћ.md"
)
if (Test-Path $meta) { $lines += "(meta-skill: using-agent-skills)" }
if (Test-Path $project) { $lines += "(РєРѕРЅС‚РµРєСЃС‚ РїСЂРѕРµРєС‚Р°: project-shishov)" }

$msg = $lines -join "`n"
$obj = @{ additional_context = $msg }
$obj | ConvertTo-Json -Compress
