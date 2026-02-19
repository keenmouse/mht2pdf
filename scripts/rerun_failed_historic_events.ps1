$ErrorActionPreference = "Stop"

$repoRoot = "<project-root>"
$pyScript = Join-Path $repoRoot "scripts\convert_mht_to_pdf.py"
$logPath = Join-Path $repoRoot "logs\historic-events-rerun.log"

$files = @(
  "<source-root>\Elizabeth Smart\The Salt Lake Tribune -- It has been a year since Elizabeth Smart returned home, her kidnapping an event that changed her family and the community_ Now the focus is on _ _ _ 'Going Forward'.mht",
  "<source-root>\Terrorism & War, beginning 2001-09-11\al-Qa'ida (The Base) - Maktab al-Khidamat (MAK - Services Office) - International Islamic Front for Jihad Against the Jews and Crusaders - Osama bin Laden - Ussama Ibn Ladin.mht",
  "<source-root>\Terrorism & War, beginning 2001-09-11\FindLaw's Writ - Lazarus Warrantless Wiretapping Why It Seriously Imperils the Separation of Powers, And Continues the Executive's Sapping of Power From Congress and the Courts.mht",
  "<source-root>\Terrorism & War, beginning 2001-09-11\Unclassified Report to Congress on the Acquisition of Technology Relating to Weapons of Mass Destruction and Advanced Conventional Munitions, 1 January Through 30 June 2001.mht",
  "<source-root>\Terrorism & War, beginning 2001-09-11\Unclassified Report to Congress on the Acquisition of Technology Relating to Weapons of Mass Destruction and Advanced Conventional Munitions, 1 July Through 31 December 2001.mht"
)

$args = @(
  $pyScript,
  "--log-path", $logPath
)

foreach ($f in $files) {
  $args += @("--source-file", $f)
}

Write-Host "Running mht2pdf rerun for $($files.Count) failed files..."
Write-Host "Log: $logPath"
python @args
