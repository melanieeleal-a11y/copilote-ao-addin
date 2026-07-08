# Installe le complément Word « Copilote AO » (Énergies Locales).
# À exécuter UNE SEULE FOIS, dans PowerShell en tant qu'administrateur :
#   clic droit sur le menu Démarrer > « Terminal (administrateur) », puis :
#   powershell -ExecutionPolicy Bypass -File .\installer-complement.ps1
#
# Ce que fait ce script :
#  1. Crée le dossier C:\CopiloteAO et y télécharge le manifeste du complément.
#  2. Partage ce dossier en lecture (requis par Word pour les catalogues de compléments).
#  3. Déclare ce partage comme « catalogue approuvé » dans Word (registre, utilisateur courant).
# Ensuite, dans Word : Fichier > Options > Centre de gestion de la confidentialité
# n'est PAS nécessaire — allez directement dans Insertion > Compléments >
# Mes compléments > DOSSIER PARTAGÉ > Copilote AO.

$ErrorActionPreference = 'Stop'

$dossier = 'C:\CopiloteAO'
$nomPartage = 'CopiloteAO'
$urlManifeste = 'https://melanieeleal-a11y.github.io/copilote-ao-addin/manifest.xml'
$idCatalogue = '0f7bfa9b-3f6a-4f6e-9f2a-1c2d3e4f5a6b'

Write-Host '1/3 - Création du dossier et téléchargement du manifeste...'
New-Item -ItemType Directory -Path $dossier -Force | Out-Null
Invoke-WebRequest -Uri $urlManifeste -OutFile (Join-Path $dossier 'manifest.xml') -UseBasicParsing

Write-Host '2/3 - Partage du dossier (lecture seule, pour Word)...'
if (-not (Get-SmbShare -Name $nomPartage -ErrorAction SilentlyContinue)) {
  New-SmbShare -Name $nomPartage -Path $dossier -ReadAccess "$env:USERDOMAIN\$env:USERNAME" | Out-Null
}
$chemin = "\\$env:COMPUTERNAME\$nomPartage"

Write-Host '3/3 - Déclaration du catalogue approuvé dans Word...'
$cleBase = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
$cle = Join-Path $cleBase $idCatalogue
New-Item -Path $cle -Force | Out-Null
Set-ItemProperty -Path $cle -Name 'Id' -Value $idCatalogue
Set-ItemProperty -Path $cle -Name 'Url' -Value $chemin
Set-ItemProperty -Path $cle -Name 'Flags' -Value 1 -Type DWord

Write-Host ''
Write-Host '================================================================'
Write-Host ' Installation terminée !'
Write-Host ''
Write-Host ' 1. Fermez complètement Word puis rouvrez-le.'
Write-Host " 2. Onglet Insertion > Compléments (ou « Mes compléments »)."
Write-Host " 3. Ouvrez l'onglet « DOSSIER PARTAGÉ » et choisissez « Copilote AO »."
Write-Host " 4. Le bouton « Copilote AO » apparaît aussi dans l'onglet Accueil."
Write-Host '================================================================'
